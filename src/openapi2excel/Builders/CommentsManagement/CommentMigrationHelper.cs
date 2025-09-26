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

namespace openapi2excel.core.Builders.CommentsManagement;

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
    /// Uses the OFFICIAL OpenXML SDK pattern from ThreadedCommentExample.
    /// This is the proven working approach that generates Excel-compatible comments.
    /// </summary>
    private static void CreateThreadedCommentsForWorksheet(
        WorksheetPart worksheetPart,
        List<ThreadedCommentWithContext> comments,
        List<WorksheetOpenApiMapping> newWorkbookMappings,
        string existingWorkbookPath,
        Dictionary<string, string> idMapping)
    {
        if (!comments.Any()) return;

        // **STEP 1: Create PersonPart (required for ThreadedComments)**
        var workbookPart = worksheetPart.OpenXmlPackage.GetPartsOfType<WorkbookPart>().First();
        EnsurePersonsPartExistsForComments(workbookPart, comments, existingWorkbookPath);

        // **STEP 2: Create WorksheetCommentsPart (legacy comments for visibility)**
        var legacyCommentsPart = worksheetPart.GetPartsOfType<WorksheetCommentsPart>().FirstOrDefault();
        if (legacyCommentsPart == null)
        {
            legacyCommentsPart = worksheetPart.AddNewPart<WorksheetCommentsPart>();
        }
        CreateLegacyCommentsUsingOfficialPattern(legacyCommentsPart, comments, newWorkbookMappings);

        // **STEP 3: Create WorksheetThreadedCommentsPart**
        var threadedCommentsPart = worksheetPart.GetPartsOfType<WorksheetThreadedCommentsPart>().FirstOrDefault();
        if (threadedCommentsPart == null)
        {
            threadedCommentsPart = worksheetPart.AddNewPart<WorksheetThreadedCommentsPart>();
        }
        CreateThreadedCommentsUsingOfficialPattern(threadedCommentsPart, comments, newWorkbookMappings, existingWorkbookPath, idMapping);

        // **STEP 4: Create VmlDrawingPart using official pattern**
        CreateVmlDrawingPartUsingOfficialPattern(worksheetPart, comments, newWorkbookMappings);

        // **STEP 5: Add LegacyDrawing reference**
        EnsureLegacyDrawingReference(worksheetPart);
    }

    /// <summary>
    /// Ensures the worksheet has a LegacyDrawing reference (required for comment display).
    /// Uses the official Microsoft SDK pattern.
    /// </summary>
    private static void EnsureLegacyDrawingReference(WorksheetPart worksheetPart)
    {
        var worksheet = worksheetPart.Worksheet;
        var legacyDrawing = worksheet.GetFirstChild<LegacyDrawing>();
        
        if (legacyDrawing == null)
        {
            worksheet.AppendChild(new LegacyDrawing() { Id = "rId1" });
        }
    }

    /// <summary>
    /// Creates WorksheetThreadedCommentsPart with proper GUID matching to legacy comments.
    /// Uses manual XML generation to ensure correct 2018 schema format as expected by tests.
    /// </summary>
    private static void CreateThreadedCommentsXmlContent(
        WorksheetThreadedCommentsPart threadedCommentsPart,
        List<ThreadedCommentWithContext> comments,
        List<WorksheetOpenApiMapping> newWorkbookMappings,
        string existingWorkbookPath,
        Dictionary<string, string> idMapping)
    {
        var xml = CreateThreadedCommentsXml(comments, newWorkbookMappings, idMapping);
        
        // Write the XML content to the part
        using (var stream = threadedCommentsPart.GetStream(FileMode.Create))
        using (var writer = new StreamWriter(stream))
        {
            writer.Write(xml);
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
    /// Extracts a specific Person by ID from the source workbook using OpenXML objects.
    /// Returns null if no persons part exists or the person is not found.
    /// </summary>
    private static Person ExtractPersonFromSourceWorkbook(string existingWorkbookPath, string personId)
    {

        var defaultPerson = new Person
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

    /// <summary>
    /// Extracts the row number from a cell reference like "A1" -> 1, "B23" -> 23
    /// </summary>
    private static int ExtractRowFromCellReference(string cellReference)
    {
        var digitStart = cellReference.IndexOf(cellReference.First(char.IsDigit));
        var rowString = cellReference.Substring(digitStart);
        return int.Parse(rowString);
    }

    /// <summary>
    /// Extracts the column index (0-based) from a cell reference like "A1" -> 0, "B23" -> 1
    /// </summary>
    private static int ExtractColumnIndexFromCellReference(string cellReference)
    {
        var columnString = cellReference.Substring(0, cellReference.IndexOf(cellReference.First(char.IsDigit)));
        int columnIndex = 0;
        for (int i = 0; i < columnString.Length; i++)
        {
            columnIndex = columnIndex * 26 + (columnString[i] - 'A' + 1);
        }
        return columnIndex - 1; // Convert to 0-based index
    }

    /// <summary>
    /// Ensures PersonPart exists for all comment authors using official SDK pattern.
    /// </summary>
    private static void EnsurePersonsPartExistsForComments(
        WorkbookPart workbookPart, 
        List<ThreadedCommentWithContext> comments, 
        string existingWorkbookPath)
    {
        var personPart = workbookPart.GetPartsOfType<WorkbookPersonPart>().FirstOrDefault();
        if (personPart == null)
        {
            personPart = workbookPart.AddNewPart<WorkbookPersonPart>();
            personPart.PersonList = new PersonList();
        }

        // Add persons for each unique author
        foreach (var comment in comments)
        {
            if (comment.Comment.PersonId?.Value != null)
            {
                var personId = comment.Comment.PersonId.Value;
                
                // Check if person already exists
                var existingPerson = personPart.PersonList.Elements<Person>()
                    .FirstOrDefault(p => p.Id?.Value == personId);
                
                if (existingPerson == null)
                {
                    // Extract person from source or create default
                    var person = ExtractPersonFromSourceWorkbook(existingWorkbookPath, personId);
                    personPart.PersonList.AppendChild(new Person
                    {
                        Id = personId,
                        DisplayName = person.DisplayName ?? "OpenAPI2Excel User",
                        ProviderId = person.ProviderId ?? "Excel",
                        UserId = person.UserId ?? "user@example.com"
                    });
                }
            }
        }
    }

    /// <summary>
    /// Creates legacy comments using the official SDK pattern.
    /// Legacy comments are required for Excel to show comment indicators.
    /// </summary>
    private static void CreateLegacyCommentsUsingOfficialPattern(
        WorksheetCommentsPart legacyCommentsPart,
        List<ThreadedCommentWithContext> comments,
        List<WorksheetOpenApiMapping> newWorkbookMappings)
    {
        var authors = new Authors();
        var commentList = new CommentList();

        // Create author entries using the EXACT official pattern
        var processedComments = new List<(ThreadedCommentWithContext comment, string cellRef, string tcId)>();
        
        foreach (var comment in comments.Where(c => c.IsRootComment))
        {
            var (targetMapping, _) = FindMatchingMapping(comment.OpenApiAnchor, newWorkbookMappings);
            if (targetMapping == null) continue;

            if (!TryGetTargetCell(comment, targetMapping, out string targetCellReference)) continue;

            // Generate unique tcId for this comment (like official example)
            string tcId = comment.CommentId;
            
            // Add author using EXACT official pattern
            authors.AppendChild(new Author("tc=" + tcId));
            
            processedComments.Add((comment, targetCellReference, tcId));
        }

        // Create legacy comments using EXACT official pattern
        for (int i = 0; i < processedComments.Count; i++)
        {
            var (comment, cellRef, tcId) = processedComments[i];
            
            // Create legacy comment following EXACT official pattern
            var legacyComment = new Comment(
                new CommentText(new Text($"Comment: {comment.CommentText}")))
            {
                Reference = cellRef,
                AuthorId = (uint)i,  // Sequential author ID
                ShapeId = 0,         // Official example uses 0
                Guid = tcId          // CRITICAL: Must match ThreadedComment.Id
            };
            
            commentList.AppendChild(legacyComment);
        }

        legacyCommentsPart.Comments = new Comments(authors, commentList);
    }

    /// <summary>
    /// Creates threaded comments using the official SDK pattern.
    /// </summary>
    private static void CreateThreadedCommentsUsingOfficialPattern(
        WorksheetThreadedCommentsPart threadedCommentsPart,
        List<ThreadedCommentWithContext> comments,
        List<WorksheetOpenApiMapping> newWorkbookMappings,
        string existingWorkbookPath,
        Dictionary<string, string> idMapping)
    {
        var threadedCommentsList = new List<ThreadedComment>();

        // Create threaded comments using EXACT official pattern
        foreach (var comment in comments)
        {
            var (targetMapping, _) = FindMatchingMapping(comment.OpenApiAnchor, newWorkbookMappings);
            if (targetMapping == null) continue;

            if (!TryGetTargetCell(comment, targetMapping, out string targetCellReference)) continue;

            // Create threaded comment following EXACT official pattern
            var threadedComment = new ThreadedComment(
                new ThreadedCommentText(comment.CommentText))
            {
                Ref = targetCellReference,
                PersonId = comment.Comment.PersonId?.Value ?? Guid.NewGuid().ToString(),
                Id = comment.CommentId, // CRITICAL: Must match legacy Comment.Guid
                DT = comment.CreatedDate ?? DateTime.Now
            };

            // Add parent reference for replies (official pattern)
            if (!comment.IsRootComment && !string.IsNullOrEmpty(comment.Comment.ParentId?.Value))
            {
                threadedComment.ParentId = comment.Comment.ParentId.Value;
            }

            threadedCommentsList.Add(threadedComment);
        }

        threadedCommentsPart.ThreadedComments = new ThreadedComments(threadedCommentsList);
    }

    /// <summary>
    /// Creates VML Drawing Part using the exact official SDK pattern.
    /// This is the proven working VML that Excel accepts.
    /// </summary>
    private static void CreateVmlDrawingPartUsingOfficialPattern(
        WorksheetPart worksheetPart,
        List<ThreadedCommentWithContext> comments,
        List<WorksheetOpenApiMapping> newWorkbookMappings)
    {
        // Check if ClosedXML already created a VML part
        var existingVmlPart = worksheetPart.GetPartsOfType<VmlDrawingPart>().FirstOrDefault();
        VmlDrawingPart vmlDrawingPart;
        
        if (existingVmlPart != null)
        {
            // Use the existing VML part but replace its content
            vmlDrawingPart = existingVmlPart;
        }
        else
        {
            // Create new VML part with specific relationship ID like official example
            vmlDrawingPart = worksheetPart.AddNewPart<VmlDrawingPart>("rId1");
        }

        using var writer = new System.Xml.XmlTextWriter(vmlDrawingPart.GetStream(FileMode.Create), System.Text.Encoding.UTF8);
        // Use the EXACT VML from the official SDK example that works
        string vmlContent = @"<xml xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:o=""urn:schemas-microsoft-com:office:office"" xmlns:x=""urn:schemas-microsoft-com:office:excel"">
                <o:shapelayout v:ext=""edit"">
                    <o:idmap v:ext=""edit"" data=""1""/>
                </o:shapelayout>
                <v:shapetype id=""_x0000_t202"" coordsize=""21600,21600"" o:spt=""202"" path=""m,l,21600r21600,l21600,xe"">
                    <v:stroke joinstyle=""miter""/>
                    <v:path gradientshapeok=""t"" o:connecttype=""rect""/>
                </v:shapetype>";

        int shapeId = 1025; // Use official example's starting shape ID

        // Create VML shape for each root comment
        foreach (var comment in comments.Where(c => c.IsRootComment))
        {
            var (targetMapping, _) = FindMatchingMapping(comment.OpenApiAnchor, newWorkbookMappings);
            if (targetMapping == null) continue;

            if (!TryGetTargetCell(comment, targetMapping, out string targetCellReference)) continue;

            // Extract row and column for VML anchor (0-based for VML)
            var row = ExtractRowFromCellReference(targetCellReference) - 1;
            var col = ExtractColumnIndexFromCellReference(targetCellReference);

            // Use EXACT VML shape pattern from official example - CRITICAL: no space after semicolon
            vmlContent += $@"
                <v:shape id=""_x0000_s{shapeId}"" type=""#_x0000_t202"" style=""position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;z-index:1;visibility:hidden"" fillcolor=""#ffffe1"" o:insetmode=""auto"">
                    <v:fill color2=""#ffffe1""/>
                    <v:shadow on=""t"" color=""black"" obscured=""t""/>
                    <v:path o:connecttype=""none""/>
                    <v:textbox style=""mso-direction-alt:auto"">
                        <div style=""text-align:left""></div>
                    </v:textbox>
                    <x:ClientData ObjectType=""Note"">
                        <x:MoveWithCells/>
                        <x:SizeWithCells/>
                        <x:Anchor>1, 15, {row}, 2, 3, 15, {row + 3}, 16</x:Anchor>
                        <x:AutoFill>False</x:AutoFill>
                        <x:Row>{row}</x:Row>
                        <x:Column>{col}</x:Column>
                    </x:ClientData>
                </v:shape>";

            shapeId++;
        }

        vmlContent += "</xml>";
        writer.WriteRaw(vmlContent);
        writer.Flush();
    }
}