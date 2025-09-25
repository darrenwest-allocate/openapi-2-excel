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
    /// </summary>
    public static List<(ThreadedCommentWithContext, CommentMigrationFailureReason)> MigrateComments(
         string existingWorkbookPath,
         IXLWorkbook newWorkbook,
         List<WorksheetOpenApiMapping> newWorkbookMappings)
    {
        // Step 1: Extract unresolved comments from existing workbook with annotations
        var commentsToMigrate = ExcelOpenXmlHelper.ExtractAndAnnotateUnresolvedComments(existingWorkbookPath);
        if (!commentsToMigrate.Any()) return new List<(ThreadedCommentWithContext, CommentMigrationFailureReason)>();

        // Tracks old ID -> new ID mapping to preserve threaded comment chains
        var idMapping = new Dictionary<string, string>();
        // Prevents broken parent-child ID relationships with a ort of comments to ensure parent comments are migrated before their replies
        var sortedComments = SortCommentsForMigration(commentsToMigrate);

        // Migrate comments directly to the saved workbook using OpenXML
        var migratedCommentCount = 0;
        var nonMigratableComments = new List<(ThreadedCommentWithContext, CommentMigrationFailureReason)>();

        foreach (var comment in sortedComments.Where(c => c.IsRootComment))
        {
            var migrationResult = TryMigrateThreadedComment(
                comment,
                newWorkbook,
                newWorkbookMappings,
                idMapping,
                existingWorkbookPath);

            if (migrationResult.Success)
            {
                migratedCommentCount++;
                // Recursively migrate replies
                // MigrateReplies(comment, workbookPart, newWorkbookMappings, idMapping, existingWorkbookPath, sortedComments);
            }
            else
            {
                nonMigratableComments.Add((comment, migrationResult.FailureReason ?? CommentMigrationFailureReason.Unknown));
                // Also add replies to non-migratable list
                // nonMigratableComments.AddRange(comment.GetReplies(sortedComments).Select(r => (r, CommentMigrationFailureReason.ParentFailedToMigrate)));
            }
        }


        // TODO: Handle nonMigratableComments in a future iteration (create "Lost Commentary" worksheet)
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
        string existingWorkbookPath)
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

            ReplicateSourceCommentOnNewWorksheet(worksheet, targetCellReference, comment);
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
    private static string ReplicateSourceCommentOnNewWorksheet(
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
        return string.Empty;
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
        
        Console.WriteLine($"[DEBUG] Sorted {sortedComments.Count} comments for migration (Root: {rootComments.Count}, Replies: {comments.Count - rootComments.Count})");
        
        return sortedComments;
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
            Console.WriteLine($"[DEBUG] Persons part already exists, checking for personId: {personId}");
            
            // Check if the specific personId already exists
            var existingPerson = existingPersonPart.PersonList?.Elements<Person>()
                .FirstOrDefault(p => p.Id?.Value == personId);
            
            if (existingPerson != null)
            {
                Console.WriteLine($"[DEBUG] Person {personId} already exists in persons part");
                return;
            }
            
            Console.WriteLine($"[DEBUG] Person {personId} not found, need to add from source workbook");
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
            Console.WriteLine($"[DEBUG] Added person to existing persons part: {newPerson.Id} - {newPerson.DisplayName}");
        }
        else
        {
            // Create new persons part
            var personPart = workbookPart.AddNewPart<WorkbookPersonPart>();
            var personList = new PersonList();
            personList.AppendChild(newPerson);
            personPart.PersonList = personList;
            Console.WriteLine($"[DEBUG] Created new WorkbookPersonPart with person: {newPerson.Id} - {newPerson.DisplayName}");
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
                Console.WriteLine("[DEBUG] No workbook part in source document");
                return defaultPerson;
            }

            var sourcePersonPart = sourceWorkbookPart.GetPartsOfType<WorkbookPersonPart>().FirstOrDefault();
            if (sourcePersonPart?.PersonList == null)
            {
                Console.WriteLine("[DEBUG] No persons part found in source workbook");
                return defaultPerson;
            }

            // Find the specific person by ID
            var person = sourcePersonPart.PersonList.Elements<Person>()
                .FirstOrDefault(p => p.Id?.Value == personId);

            if (person != null)
            {
                Console.WriteLine($"[DEBUG] Found person {personId} in source workbook: {person.DisplayName?.Value}");
                return (Person)person.Clone();
            }
            else
            {
                Console.WriteLine($"[DEBUG] Person {personId} not found in source workbook");
                return defaultPerson;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[DEBUG] Error extracting person {personId} from source: {ex.Message}");
            return defaultPerson;
        }
    }
}