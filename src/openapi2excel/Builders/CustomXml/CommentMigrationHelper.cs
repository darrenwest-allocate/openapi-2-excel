using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2019.Excel.ThreadedComments;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using openapi2excel.core.Common;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace openapi2excel.core.Builders.CustomXml;

/// <summary>
/// Reasons why a threaded comment migration might fail.
/// </summary>
public enum CommentMigrationFailureReason
{
    /// <summary>
    /// The comment has no associated OpenAPI anchor, making it impossible to map to the new workbook.
    /// </summary>
    NoOpenApiAnchorFound,

    /// <summary>
    /// The OpenAPI anchor exists but cannot be found in the new workbook's cell mappings.
    /// </summary>
    OpenApiAnchorNotFoundInNewWorkbook,

    /// <summary>
    /// The target worksheet referenced in the mapping does not exist in the new workbook.
    /// </summary>
    TargetWorksheetNotFound,

    /// <summary>
    /// An unexpected error occurred during the migration process (e.g., file I/O, XML parsing).
    /// </summary>
    UnexpectedErrorDuringMigration,
	Unknown
}

/// <summary>
/// Helper class for migrating Excel threaded comments from an existing workbook to a new workbook
/// based on OpenAPI anchor mappings.
/// </summary>
public static class CommentMigrationHelper
{
    /// <summary>
    /// Migrates unresolved comments from an existing Excel workbook to a new workbook
    /// using OpenAPI anchor mappings to determine the correct cell placement.
    /// Uses a two-phase approach: Phase 1 creates legacy comments with ClosedXML,
    /// Phase 2 adds threaded comment parts using OpenXML.
    /// </summary>
    public static List<(ThreadedCommentWithContext, CommentMigrationFailureReason)> MigrateComments(
         string existingWorkbookPath,
         string newWorkbookPath,
         List<WorksheetOpenApiMapping> newWorkbookMappings)
    {        
        // Step 1: Extract unresolved comments from existing workbook with annotations
        var commentsToMigrate = ExcelOpenXmlHelper.ExtractAndAnnotateUnresolvedComments(existingWorkbookPath);
        
        if (!commentsToMigrate.Any()) return new List<(ThreadedCommentWithContext, CommentMigrationFailureReason)>();

        // Tracks old ID -> new ID mapping to preserve threaded comment chains
        var idMapping = new Dictionary<string, string>();
        var sortedComments = SortCommentsForMigration(commentsToMigrate);
        var nonMigratableComments = new List<(ThreadedCommentWithContext, CommentMigrationFailureReason)>();
        var processedCells = new HashSet<string>(); // Track cells that already have legacy comments

        // PHASE 1: Create legacy comments using ClosedXML (only root comments)
        using (var newWorkbook = new XLWorkbook(newWorkbookPath))
        {
            foreach (var comment in sortedComments)
            {
                // Only process root comments for legacy comment creation
                if (!comment.IsRootComment) continue;
                
                var migrationResult = TryMigrateThreadedComment(
                    comment,
                    newWorkbook,
                    newWorkbookMappings,
                    idMapping,
                    existingWorkbookPath,
                    processedCells);

                if (!migrationResult.Success)
                {
                    nonMigratableComments.Add((comment, migrationResult.FailureReason ?? CommentMigrationFailureReason.Unknown));
                }
            }
            newWorkbook.Save();
        }

        // PHASE 2: Add threaded comment parts using OpenXML
        AddThreadedCommentParts(existingWorkbookPath, newWorkbookPath, sortedComments, newWorkbookMappings, idMapping);

        return nonMigratableComments;
    }

    /// <summary>
    /// Attempts to migrate a single threaded comment to the new workbook based on OpenAPI anchor mapping.
    /// </summary>
    private static (bool Success, CommentMigrationFailureReason? FailureReason, string? ErrorDetails) TryMigrateThreadedComment(
        ThreadedCommentWithContext comment,
        IXLWorkbook workbook,
        List<WorksheetOpenApiMapping> newWorkbookMappings,
        Dictionary<string, string> idMapping,
        string existingWorkbookPath,
        HashSet<string> processedCells)
    {
        try
        {
            if (string.IsNullOrEmpty(comment.OpenApiAnchor))
            {
                return (false, CommentMigrationFailureReason.NoOpenApiAnchorFound, "Comment has no OpenAPI anchor.");
            }

            var (targetMapping, worksheetName) = FindMatchingMapping(comment.OpenApiAnchor, newWorkbookMappings);
            if (targetMapping == null)
            {
                return (false, CommentMigrationFailureReason.OpenApiAnchorNotFoundInNewWorkbook, $"Anchor '{comment.OpenApiAnchor}' not found in new workbook mappings.");
            }

            if (!workbook.Worksheets.TryGetWorksheet(worksheetName, out var worksheet))
            {
                return (false, CommentMigrationFailureReason.TargetWorksheetNotFound, $"Worksheet '{worksheetName}' not found in the new workbook.");
            }

            if (!TryGetTargetCell(comment, targetMapping, out string targetCellReference))
            {
                return (false, CommentMigrationFailureReason.TargetWorksheetNotFound, "Could not determine target cell for migration.");
            }

            // Create a unique key for this cell to avoid duplicate legacy comments
            var cellKey = $"{worksheetName}:{targetCellReference}";
            if (!processedCells.Contains(cellKey))
            {
                ReplicateSourceCommentOnNewWorksheet(worksheet, targetCellReference, comment);
                processedCells.Add(cellKey);
            }
            
            return (true, null, null);
        }
        catch (Exception ex)
        {
            return (false, CommentMigrationFailureReason.UnexpectedErrorDuringMigration, ex.Message);
        }
    }

    private static bool TryGetTargetCell(ThreadedCommentWithContext comment, CellOpenApiMapping targetMapping, out string targetCellReference)
    {
        if (!string.IsNullOrEmpty(targetMapping.Cell))
        {
            // Exact cell match - use the mapped cell
            targetCellReference = targetMapping.Cell;
        }
        else if (targetMapping.Row > 0)
        {
            // Row match - preserve original column, use mapped row
            var originalColumn = ExtractColumnFromCellReference(comment.CellReference);
            targetCellReference = $"{originalColumn}{targetMapping.Row}";
        }
        else
        {
            targetCellReference = string.Empty;
            return false;
        }
        return true;
    }

    /// <summary>
    /// Finds a matching cell mapping based on OpenAPI anchor.
    /// </summary>
    private static (CellOpenApiMapping? Mapping, string WorksheetName) FindMatchingMapping(
        string openApiAnchor,
        List<WorksheetOpenApiMapping> mappings)
    {
        foreach (var wsMapping in mappings)
        {
            var cellMapping = wsMapping.Mappings.FirstOrDefault(cm =>
                cm.OpenApiRef.Equals(openApiAnchor, StringComparison.OrdinalIgnoreCase));

            if (cellMapping != null)
            {
                return (cellMapping, wsMapping.WorksheetName);
            }
        }

        return (null, string.Empty);
    }

    /// <summary>
    /// Extracts the column part from a cell reference (e.g., "A5" -> "A").
    /// </summary>
    private static string ExtractColumnFromCellReference(string cellReference)
    {
        return new string([.. cellReference.TakeWhile(c => !char.IsDigit(c))]);
    }

    /// <summary>
    /// Adds a legacy comment for Excel backward compatibility and visibility.
    /// Legacy comments are required for Excel to display threaded comments properly.
    /// </summary>
    private static void ReplicateSourceCommentOnNewWorksheet(
        IXLWorksheet newWorksheet,
        string cellReference,
        ThreadedCommentWithContext sourceComment)
    {
        var cell = newWorksheet.Cell(cellReference);
        var comment = cell.GetComment();
        if (comment == null)
        {
            comment = cell.CreateComment();
        }
        comment.AddText(sourceComment.CommentText);
        comment.Author = sourceComment.Author;
        return;
    }

    /// <summary>
    /// Sorts comments to ensure parent comments are migrated before their replies.
    /// This prevents broken parent-child ID relationships in the migrated workbook.
    /// </summary>
    private static List<ThreadedCommentWithContext> SortCommentsForMigration(List<ThreadedCommentWithContext> comments)
    {
        var sortedComments = new List<ThreadedCommentWithContext>();
        var processed = new HashSet<string>();
        
        // First pass: add all root comments (comments with no parent)
        var rootComments = comments.Where(c => c.IsRootComment).ToList();
        foreach (var rootComment in rootComments)
        {
            sortedComments.Add(rootComment);
            processed.Add(rootComment.CommentId);
        }
        
        // Second pass: add reply comments in order of their parent dependencies
        var remainingComments = comments.Where(c => !processed.Contains(c.CommentId)).ToList();
        var maxIterations = remainingComments.Count + 1; // Prevent infinite loops
        var iteration = 0;
        
        while (remainingComments.Any() && iteration < maxIterations)
        {
            var addedThisIteration = new List<ThreadedCommentWithContext>();
            
            foreach (var comment in remainingComments)
            {
                var parentId = comment.Comment.ParentId?.Value;
                if (!string.IsNullOrEmpty(parentId) && processed.Contains(parentId))
                {
                    // Parent has been processed, safe to add this reply
                    sortedComments.Add(comment);
                    processed.Add(comment.CommentId);
                    addedThisIteration.Add(comment);
                }
            }
            
            foreach (var comment in addedThisIteration)
            {
                remainingComments.Remove(comment);
            }
            
            iteration++;
        }
        
        sortedComments.AddRange(remainingComments);
        
        return sortedComments;
    }

    /// <summary>
    /// Adds threaded comment parts to the workbook using OpenXML.
    /// This method handles the creation of WorksheetThreadedCommentsPart and ThreadedComment objects
    /// for root comments and their replies.
    /// </summary>
    private static void AddThreadedCommentParts(
        string existingWorkbookPath,
        string newWorkbookPath,
        List<ThreadedCommentWithContext> sortedComments,
        List<WorksheetOpenApiMapping> newWorkbookMappings,
        Dictionary<string, string> idMapping)
    {
        using var document = SpreadsheetDocument.Open(newWorkbookPath, true);
        var workbookPart = document.WorkbookPart;
        if (workbookPart == null) return;

        // Group comments by their target worksheet
        var commentsByWorksheet = new Dictionary<string, List<ThreadedCommentWithContext>>();
        
        foreach (var comment in sortedComments)
        {
            if (string.IsNullOrEmpty(comment.OpenApiAnchor)) continue;
            
            var (targetMapping, worksheetName) = FindMatchingMapping(comment.OpenApiAnchor, newWorkbookMappings);
            if (targetMapping == null) continue;
            
            if (!commentsByWorksheet.ContainsKey(worksheetName))
            {
                commentsByWorksheet[worksheetName] = new List<ThreadedCommentWithContext>();
            }
            commentsByWorksheet[worksheetName].Add(comment);
        }

        // Process each worksheet
        foreach (var (worksheetName, worksheetComments) in commentsByWorksheet)
        {
            var worksheetPart = FindWorksheetPart(workbookPart, worksheetName);
            if (worksheetPart == null) continue;

            CreateThreadedCommentsForWorksheet(worksheetPart, worksheetComments, newWorkbookMappings, existingWorkbookPath, idMapping);
        }

        document.Save();
    }

    /// <summary>
    /// Creates threaded comments for a specific worksheet.
    /// Uses the 2018 schema format to match Excel's expected structure.
    /// </summary>
    private static void CreateThreadedCommentsForWorksheet(
        WorksheetPart worksheetPart,
        List<ThreadedCommentWithContext> comments,
        List<WorksheetOpenApiMapping> newWorkbookMappings,
        string existingWorkbookPath,
        Dictionary<string, string> idMapping)
    {        
        // Get or create the threaded comments part
        var threadedCommentsPart = worksheetPart.GetPartsOfType<WorksheetThreadedCommentsPart>().FirstOrDefault();
        if (threadedCommentsPart != null)
        {
            worksheetPart.DeletePart(threadedCommentsPart);
        }
        
        threadedCommentsPart = worksheetPart.AddNewPart<WorksheetThreadedCommentsPart>();
        
        // Generate XML manually to ensure correct namespace structure (no prefixes)
        var xml = CreateThreadedCommentsXml(comments, newWorkbookMappings, idMapping);
        
        // Write the XML directly to the part
        using (var stream = threadedCommentsPart.GetStream(FileMode.Create))
        using (var writer = new StreamWriter(stream))
        {
            writer.Write(xml);
        }

        // Ensure the relationship is properly set up in the worksheet
        // This is crucial for Excel to recognize the threaded comments
        var worksheet = worksheetPart.Worksheet;
        if (worksheet != null)
        {
            var relationshipId = worksheetPart.GetIdOfPart(threadedCommentsPart);
        }

        // Ensure persons part exists for all comment authors
        var document = worksheetPart.OpenXmlPackage as SpreadsheetDocument;
        foreach (var comment in comments)
        {
            if (comment.Comment.PersonId?.Value != null)
            {
                EnsurePersonsPartExists(document, comment.Comment.PersonId.Value, existingWorkbookPath);
            }
        }
    }

    /// <summary>
    /// Creates the threaded comments XML content manually to ensure correct 2018 schema format.
    /// </summary>
    private static string CreateThreadedCommentsXml(
        List<ThreadedCommentWithContext> comments,
        List<WorksheetOpenApiMapping> newWorkbookMappings,
        Dictionary<string, string> idMapping)
    {
        var xmlBuilder = new System.Text.StringBuilder();
        
        // XML header and root element with correct 2018 namespaces
        xmlBuilder.AppendLine("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        xmlBuilder.AppendLine("<ThreadedComments xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments\" xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
        
        // Process root comments and their replies
        foreach (var rootComment in comments.Where(c => c.IsRootComment))
        {
            var (targetMapping, _) = FindMatchingMapping(rootComment.OpenApiAnchor, newWorkbookMappings);
            if (targetMapping == null) continue;

            if (!TryGetTargetCell(rootComment, targetMapping, out string targetCellReference)) continue;

            // Create the root threaded comment
            var rootId = Guid.NewGuid().ToString("B").ToUpper(); // Format: {GUID}
            var rootDateTime = rootComment.Comment.DT?.Value.ToString("yyyy-MM-ddTHH:mm:ss.ff") ?? DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.ff");
            var personId = rootComment.Comment.PersonId?.Value ?? "";
            
            var doneAttr = rootComment.Comment.Done?.Value == true ? " done=\"1\"" : "";
            
            xmlBuilder.AppendLine($"<threadedComment ref=\"{targetCellReference}\" dT=\"{rootDateTime}\" personId=\"{personId}\" id=\"{rootId}\"{doneAttr}>");
            xmlBuilder.AppendLine($"<text>{System.Security.SecurityElement.Escape(rootComment.CommentText)}</text>");
            xmlBuilder.AppendLine("</threadedComment>");
            
            // Track the ID mapping
            if (!string.IsNullOrEmpty(rootComment.CommentId))
            {
                idMapping[rootComment.CommentId] = rootId;
            }

            // Add replies to this root comment
            var replies = rootComment.GetReplies(comments).ToList();
            foreach (var reply in replies)
            {
                var replyId = Guid.NewGuid().ToString("B").ToUpper();
                var replyDateTime = reply.Comment.DT?.Value.ToString("yyyy-MM-ddTHH:mm:ss.ff") ?? DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.ff");
                var replyPersonId = reply.Comment.PersonId?.Value ?? "";
                
                xmlBuilder.AppendLine($"<threadedComment ref=\"{targetCellReference}\" dT=\"{replyDateTime}\" personId=\"{replyPersonId}\" id=\"{replyId}\" parentId=\"{rootId}\">");
                xmlBuilder.AppendLine($"<text>{System.Security.SecurityElement.Escape(reply.CommentText)}</text>");
                xmlBuilder.AppendLine("</threadedComment>");
                
                // Track the reply ID mapping
                if (!string.IsNullOrEmpty(reply.CommentId))
                {
                    idMapping[reply.CommentId] = replyId;
                }
            }
        }
        
        xmlBuilder.AppendLine("</ThreadedComments>");
        return xmlBuilder.ToString();
    }


    /// <summary>
    /// Finds the WorksheetPart for a given worksheet name.
    /// </summary>
    private static WorksheetPart? FindWorksheetPart(WorkbookPart workbookPart, string worksheetName)
    {
        var sheet = workbookPart.Workbook.Descendants<Sheet>()
            .FirstOrDefault(s => s.Name?.Value?.Equals(worksheetName, StringComparison.OrdinalIgnoreCase) == true);
        
        if (sheet?.Id?.Value != null)
        {
            return workbookPart.GetPartById(sheet.Id.Value) as WorksheetPart;
        }
        
        return null;
    }

    /// <summary>
    /// Ensures that the persons part exists in the workbook with the specified person ID.
    /// This is required for Excel to properly validate personId references in threaded comments.
    /// Uses a hybrid approach: extract persons from source workbook using OpenXML objects,
    /// then create proper WorkbookPersonPart in target workbook.
    /// </summary>
    private static void EnsurePersonsPartExists(SpreadsheetDocument? document, string? personId, string existingWorkbookPath)
    {
        if (document?.WorkbookPart == null || string.IsNullOrEmpty(personId))
        {
            return;
        }

        var workbookPart = document.WorkbookPart;

        var existingPersonPart = workbookPart.GetPartsOfType<WorkbookPersonPart>().FirstOrDefault();
        if (existingPersonPart != null)
        {
            // Check if the specific personId already exists
            var existingPerson = existingPersonPart.PersonList?.Elements<Person>()
                .FirstOrDefault(p => p.Id?.Value == personId);
            
            if (existingPerson != null)
            {
                return;
            }
        }

        // Extract the specific person from source workbook
        var sourcePerson = ExtractPersonFromSourceWorkbook(existingWorkbookPath, personId);

        var newPerson = new Person
        {
            Id = sourcePerson.Id?.Value,
            DisplayName = sourcePerson.DisplayName?.Value,
            ProviderId = sourcePerson.ProviderId?.Value,
            UserId = sourcePerson.UserId?.Value
        };
        // Add the person to existing or new persons part
        if (existingPersonPart != null)
        {
            existingPersonPart.PersonList?.AppendChild(newPerson);
        }
        else
        {
            // Create new persons part
            var personPart = workbookPart.AddNewPart<WorkbookPersonPart>();
            var personList = new PersonList();
            personList.AppendChild(newPerson);
            personPart.PersonList = personList;
        }
    }

    /// <summary>
    /// Extracts a specific Person by ID from the source workbook using OpenXML objects.
    /// Returns null if no persons part exists or the person is not found.
    /// </summary>
    private static Person ExtractPersonFromSourceWorkbook(string existingWorkbookPath, string personId)
    {

        var defaultPerson =  new Person
            {
                Id = personId,
                DisplayName = "Comment Author",
                ProviderId = "Excel"
            };

        try
        {
            using var sourceDocument = SpreadsheetDocument.Open(existingWorkbookPath, false);
            var sourceWorkbookPart = sourceDocument.WorkbookPart;
            if (sourceWorkbookPart == null)
            {
                return defaultPerson;
            }

            var sourcePersonPart = sourceWorkbookPart.GetPartsOfType<WorkbookPersonPart>().FirstOrDefault();
            if (sourcePersonPart?.PersonList == null)
            {
                return defaultPerson;
            }

            // Find the specific person by ID
            var person = sourcePersonPart.PersonList.Elements<Person>()
                .FirstOrDefault(p => p.Id?.Value == personId);

            if (person != null)
            {
                return (Person)person.Clone();
            }
            else
            {
                return defaultPerson;
            }
        }
        catch (Exception)
        {
            return defaultPerson;
        }
    }
}