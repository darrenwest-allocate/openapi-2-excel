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
    
    /// <summary>
    /// Comment successfully migrated with no OpenAPI anchor - placed near title row on existing worksheet.
    /// </summary>
    SuccessfullyMigratedAsNoAnchorComment,
    
    /// <summary>
    /// Comment successfully migrated from missing worksheet - placed on Info sheet.
    /// </summary>
    SuccessfullyMigratedAsNoWorksheetComment,
    
    Unknown
}

/// <summary>
/// Context object that holds the migration state during comment processing.
/// </summary>
public class MigrationContext
{
    public Dictionary<string, string> IdMapping { get; } = new Dictionary<string, string>();
    public List<ThreadedCommentWithContext> SortedComments { get; set; } = new List<ThreadedCommentWithContext>();
    public HashSet<string> ProcessedCells { get; } = new HashSet<string>();
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
        var commentsToMigrate = ExtractCommentsToMigrate(existingWorkbookPath);
        if (!commentsToMigrate.Any()) return new List<(ThreadedCommentWithContext, CommentMigrationFailureReason)>();

        var migrationContext = InitializeMigrationContext(commentsToMigrate);
        var nonMigratableComments = CreateLegacyComments(newWorkbookPath, newWorkbookMappings, migrationContext);
        AddThreadedCommentParts(existingWorkbookPath, newWorkbookPath, migrationContext.SortedComments, newWorkbookMappings, migrationContext.IdMapping);

        return nonMigratableComments;
    }

    /// <summary>
    /// Extracts unresolved comments from the existing workbook.
    /// </summary>
    private static List<ThreadedCommentWithContext> ExtractCommentsToMigrate(string existingWorkbookPath)
    {
        return ExcelOpenXmlHelper.ExtractAndAnnotateUnresolvedComments(existingWorkbookPath);
    }

    /// <summary>
    /// Initializes the migration context with sorted comments and tracking structures.
    /// </summary>
    private static MigrationContext InitializeMigrationContext(List<ThreadedCommentWithContext> commentsToMigrate)
    {
        return new MigrationContext
        {
            SortedComments = SortCommentsForMigration(commentsToMigrate)
        };
    }

    /// <summary>
    /// Creates legacy comments in the new workbook using ClosedXML (Phase 1).
    /// Processes only root comments and tracks migration failures.
    /// </summary>
    private static List<(ThreadedCommentWithContext, CommentMigrationFailureReason)> CreateLegacyComments(
        string newWorkbookPath,
        List<WorksheetOpenApiMapping> newWorkbookMappings,
        MigrationContext migrationContext)
    {
        var nonMigratableComments = new List<(ThreadedCommentWithContext, CommentMigrationFailureReason)>();

        using var newWorkbook = new XLWorkbook(newWorkbookPath);
        
        foreach (var comment in migrationContext.SortedComments)
        {
            // Only process root comments for legacy comment creation
            if (!comment.IsRootComment) continue;
            
            var migrationResult = TryMigrateThreadedComment(
                comment,
                newWorkbook,
                newWorkbookMappings,
                migrationContext.IdMapping,
                migrationContext.ProcessedCells,
                migrationContext.SortedComments);

            if (!migrationResult.Success)
            {
                nonMigratableComments.Add((comment, migrationResult.FailureReason ?? CommentMigrationFailureReason.Unknown));
            }
        }
        
        newWorkbook.Save();
        return nonMigratableComments;
    }

    /// <summary>
    /// Attempts to migrate a single threaded comment using available strategies.
    /// </summary>
    private static (bool Success, CommentMigrationFailureReason? FailureReason, string? ErrorDetails) TryMigrateThreadedComment(
        ThreadedCommentWithContext comment,
        IXLWorkbook workbook,
        List<WorksheetOpenApiMapping> newWorkbookMappings,
        Dictionary<string, string> idMapping,
        HashSet<string> processedCells,
        List<ThreadedCommentWithContext> sortedComments)
    {
        return ProcessCommentWithStrategies(comment, workbook, newWorkbookMappings, processedCells, sortedComments);
    }

    /// <summary>
    /// Processes a comment using the available migration strategies in order of preference.
    /// </summary>
    private static (bool Success, CommentMigrationFailureReason? FailureReason, string? ErrorDetails) ProcessCommentWithStrategies(
        ThreadedCommentWithContext comment,
        IXLWorkbook workbook,
        List<WorksheetOpenApiMapping> newWorkbookMappings,
        HashSet<string> processedCells,
        List<ThreadedCommentWithContext> sortedComments)
    {
        var strategies = CreateMigrationStrategies();

        // Try each strategy until one succeeds
        foreach (var strategy in strategies)
        {
            if (strategy.CanHandle(comment, workbook, newWorkbookMappings))
            {
                return strategy.TryMigrate(comment, workbook, processedCells, sortedComments, newWorkbookMappings);
            }
        }

        // If no strategy can handle the comment
        return (false, CommentMigrationFailureReason.NoOpenApiAnchorFound, "No migration strategy could handle this comment.");
    }

    /// <summary>
    /// Creates the list of migration strategies in order of preference.
    /// </summary>
    private static List<ICommentMigrationStrategy> CreateMigrationStrategies()
    {
        var collisionResolver = new CellCollisionResolver();
        var targetResolver = new CommentTargetResolver();
        
        return new List<ICommentMigrationStrategy>
        {
            new MappedCommentMigrationStrategy(targetResolver),
            new NoAnchorCommentMigrationStrategy(collisionResolver, targetResolver),
            new NoWorksheetCommentMigrationStrategy(collisionResolver, targetResolver)
        };
    }



    /// <summary>
    /// Gets the target cell reference for a comment using the CommentTargetResolver.
    /// </summary>
    private static bool TryGetTargetCellForThreadedComment(
        ThreadedCommentWithContext comment, 
        List<WorksheetOpenApiMapping> newWorkbookMappings, 
        out string targetCellReference)
    {
        var targetResolver = new CommentTargetResolver();
        return targetResolver.TryGetTargetCellForThreadedComment(comment, newWorkbookMappings, out targetCellReference);
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
            string worksheetName;
            
            // Handle Type A comments with override target cells
            if (!string.IsNullOrEmpty(comment.OverrideTargetCell) && !string.IsNullOrEmpty(comment.OverrideWorksheetName))
            {
                worksheetName = comment.OverrideWorksheetName;
            }
            // Handle regular comments with OpenAPI anchors
            else if (!string.IsNullOrEmpty(comment.OpenApiAnchor))
            {
                var targetResolver = new CommentTargetResolver();
                var (targetMapping, mappedWorksheetName) = targetResolver.FindMatchingMapping(comment.OpenApiAnchor, newWorkbookMappings);
                if (targetMapping == null) continue;
                worksheetName = mappedWorksheetName;
            }
            else
            {
                continue; // Skip comments that cannot be mapped
            }
            
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
    /// Creates threaded comments for a specific worksheet using factory classes.
    /// </summary>
    private static void CreateThreadedCommentsForWorksheet(
        WorksheetPart worksheetPart,
        List<ThreadedCommentWithContext> comments,
        List<WorksheetOpenApiMapping> newWorkbookMappings,
        string existingWorkbookPath,
        Dictionary<string, string> idMapping)
    {
        if (!comments.Any()) return;

        var workbookPart = worksheetPart.OpenXmlPackage.GetPartsOfType<WorkbookPart>().First();
        
        // Initialize factories
        var targetResolver = new CommentTargetResolver();
        var personPartManager = new PersonPartManager();
        var threadedCommentXmlFactory = new ThreadedCommentXmlFactory(targetResolver);
        var vmlDrawingFactory = new VmlDrawingFactory(targetResolver);

        // **STEP 1: Create PersonPart (required for ThreadedComments)**
        personPartManager.EnsurePersonsPartExistsForComments(workbookPart, comments, existingWorkbookPath);

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

        // **STEP 4: Create VmlDrawingPart using factory**
        vmlDrawingFactory.CreateVmlDrawingPartUsingOfficialPattern(worksheetPart, comments, newWorkbookMappings);

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
            var targetResolver = new CommentTargetResolver();
            var (targetMapping, _) = targetResolver.FindMatchingMapping(rootComment.OpenApiAnchor, newWorkbookMappings);
            if (targetMapping == null) continue;

            if (!targetResolver.TryGetTargetCell(rootComment, targetMapping, out string targetCellReference)) continue;

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
            if (!TryGetTargetCellForThreadedComment(comment, newWorkbookMappings, out string targetCellReference))
                continue;

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
            if (!TryGetTargetCellForThreadedComment(comment, newWorkbookMappings, out string targetCellReference))
                continue;

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
}