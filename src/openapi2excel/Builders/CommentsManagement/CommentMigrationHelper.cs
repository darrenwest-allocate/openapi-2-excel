using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2019.Excel.ThreadedComments;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using openapi2excel.core.Builders.CommentsManagement.MigrationStrategy;
using openapi2excel.core.Common;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace openapi2excel.core.Builders.CommentsManagement;

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

            var (Success, FailureReason, ErrorDetails) = TryMigrateThreadedComment(
                comment,
                newWorkbook,
                newWorkbookMappings,
                migrationContext.IdMapping,
                migrationContext.ProcessedCells,
                migrationContext.SortedComments);

            if (!Success)
            {
                nonMigratableComments.Add((comment, FailureReason ?? CommentMigrationFailureReason.Unknown));
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
        foreach (var strategy in strategies)
        {
            if (strategy.CanHandle(comment, workbook, newWorkbookMappings))
            {
                return strategy.TryMigrate(comment, workbook, processedCells, sortedComments, newWorkbookMappings);
            }
        }

        return (false, CommentMigrationFailureReason.NoOpenApiAnchorFound, "No migration strategy could handle this comment.");
    }

    /// <summary>
    /// Creates the list of migration strategies in order of preference.
    /// </summary>
    private static List<ICommentMigrationStrategy> CreateMigrationStrategies()
    {
        return
        [
            new MappedCommentMigrationStrategy(),
            new NoAnchorCommentMigrationStrategy(),
            new NoWorksheetCommentMigrationStrategy()
        ];
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
        return CommentTargetResolver.TryGetTargetCellForThreadedComment(comment, newWorkbookMappings, out targetCellReference);
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
            if (!string.IsNullOrEmpty(comment.OverrideTargetCell) && !string.IsNullOrEmpty(comment.OverrideWorksheetName))
            {
                worksheetName = comment.OverrideWorksheetName;
            }
            else if (!string.IsNullOrEmpty(comment.OpenApiAnchor))
            {
                var targetResolver = new CommentTargetResolver();
                var (targetMapping, mappedWorksheetName) = CommentTargetResolver.FindMatchingMapping(comment.OpenApiAnchor, newWorkbookMappings);
                if (targetMapping == null) continue;
                worksheetName = mappedWorksheetName;
            }
            else
            {
                continue; // Skip comments that cannot be mapped
            }

            if (!commentsByWorksheet.TryGetValue(worksheetName, out List<ThreadedCommentWithContext>? value))
            {
                value = new List<ThreadedCommentWithContext>();
                commentsByWorksheet[worksheetName] = value;
            }

            value.Add(comment);
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

		// **STEP 1: Create PersonPart (required for ThreadedComments)**
		PersonPartManager.EnsurePersonsPartExistsForComments(workbookPart, comments, existingWorkbookPath);

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
		VmlDrawingFactory.CreateVmlDrawingPartUsingOfficialPattern(worksheetPart, comments, newWorkbookMappings);

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

        foreach (var comment in comments)
        {
            if (comment.Comment.PersonId?.Value != null)
            {
                var personId = comment.Comment.PersonId.Value;
                var existingPerson = personPart.PersonList.Elements<Person>()
                    .FirstOrDefault(p => p.Id?.Value == personId);

                if (existingPerson == null)
                {
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

        var processedComments = new List<(ThreadedCommentWithContext comment, string cellRef, string tcId)>();
        foreach (var comment in comments.Where(c => c.IsRootComment))
        {
            if (!TryGetTargetCellForThreadedComment(comment, newWorkbookMappings, out string targetCellReference))
                continue;
            // Generate unique tcId for this comment (like official example)
            string tcId = comment.CommentId;
            authors.AppendChild(new Author("tc=" + tcId));

            processedComments.Add((comment, targetCellReference, tcId));
        }

        // Create legacy comments using official pattern
        for (int i = 0; i < processedComments.Count; i++)
        {
            var (comment, cellRef, tcId) = processedComments[i];
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

        // Create threaded comments using official pattern
        foreach (var comment in comments)
        {
            if (!TryGetTargetCellForThreadedComment(comment, newWorkbookMappings, out string targetCellReference))
                continue;

            // Create threaded comment following official pattern
            var threadedComment = new ThreadedComment(
                new ThreadedCommentText(comment.CommentText))
            {
                Ref = targetCellReference,
                PersonId = comment.Comment.PersonId?.Value ?? Guid.NewGuid().ToString(),
                Id = comment.CommentId, // CRITICAL: Must match legacy Comment.Guid
                DT = comment.CreatedDate ?? DateTime.Now
            };

            // Add parent reference for replies
            if (!comment.IsRootComment && !string.IsNullOrEmpty(comment.Comment.ParentId?.Value))
            {
                threadedComment.ParentId = comment.Comment.ParentId.Value;
            }
            threadedCommentsList.Add(threadedComment);
        }

        threadedCommentsPart.ThreadedComments = new ThreadedComments(threadedCommentsList);
    }
}
