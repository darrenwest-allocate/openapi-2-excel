using ClosedXML.Excel;
using openapi2excel.core.Common;
using System;
using System.Collections.Generic;
using System.Linq;

namespace openapi2excel.core.Builders.CommentsManagement;

/// <summary>
/// Strategy for migrating comments with no OpenAPI anchor on existing worksheets.
/// Places the comment near the nearest title row above its original position, or at row 1 if no title found.
/// </summary>
public class NoAnchorCommentMigrationStrategy : ICommentMigrationStrategy
{
    private readonly CellCollisionResolver _collisionResolver;
    private readonly CommentTargetResolver _targetResolver;

    public string StrategyName => "Type A (NoAnchor)";

    public NoAnchorCommentMigrationStrategy(CellCollisionResolver collisionResolver, CommentTargetResolver targetResolver)
    {
        _collisionResolver = collisionResolver ?? throw new ArgumentNullException(nameof(collisionResolver));
        _targetResolver = targetResolver ?? throw new ArgumentNullException(nameof(targetResolver));
    }

    public bool CanHandle(ThreadedCommentWithContext comment, IXLWorkbook workbook, List<WorksheetOpenApiMapping> newWorkbookMappings)
    {
        // Type A: Comment has no OpenAPI anchor AND the source worksheet exists in the new workbook
        return string.IsNullOrEmpty(comment.OpenApiAnchor) 
               && workbook.Worksheets.TryGetWorksheet(comment.WorksheetName, out _);
    }

    public (bool Success, CommentMigrationFailureReason? FailureReason, string? ErrorDetails) TryMigrate(
        ThreadedCommentWithContext comment,
        IXLWorkbook workbook,
        HashSet<string> processedCells,
        List<ThreadedCommentWithContext> allComments,
        List<WorksheetOpenApiMapping> newWorkbookMappings)
    {
        try
        {
            if (!workbook.Worksheets.TryGetWorksheet(comment.WorksheetName, out var worksheet))
            {
                return (false, CommentMigrationFailureReason.TargetWorksheetNotFound, 
                    $"Worksheet '{comment.WorksheetName}' not found in new workbook for Type A migration.");
            }

            // Preserve the original column
            var originalColumn = _targetResolver.ExtractColumnFromCellReference(comment.CellReference);
            
            // Find the target row near a title row using OpenAPI mappings
            var targetRow = FindTargetRowForTypeAComment(worksheet, comment, newWorkbookMappings);
            
            var targetCellReference = $"{originalColumn}{targetRow}";
            
            // Check for collision and adjust if necessary
            var finalCellReference = _collisionResolver.HandleTypeACollision(worksheet, targetCellReference, processedCells);
            
            // Create the legacy comment (this creates the visible comment indicator)
            CommentMigrationSharedHelper.ReplicateSourceCommentOnNewWorksheet(worksheet, finalCellReference, comment);
            
            // Mark cell as processed
            var cellKey = $"{comment.WorksheetName}:{finalCellReference}";
            processedCells.Add(cellKey);
            
            // Store the target cell reference for ThreadedComment processing
            CommentMigrationSharedHelper.SetOverrideTargetCellForCommentAndReplies(
                comment, finalCellReference, comment.WorksheetName, allComments);
            
            return (true, CommentMigrationFailureReason.SuccessfullyMigratedAsNoAnchorComment, "Successfully migrated comment with no OpenAPI anchor");
        }
        catch (Exception ex)
        {
            return (false, CommentMigrationFailureReason.UnexpectedErrorDuringMigration, 
                $"Error during Type A migration: {ex.Message}");
        }
    }

    private int FindTargetRowForTypeAComment(IXLWorksheet worksheet, ThreadedCommentWithContext comment, List<WorksheetOpenApiMapping> newWorkbookMappings)
    {
        // Start from the original row and search upward for title rows
        var originalRow = _targetResolver.ExtractRowFromCellReference(comment.CellReference);
        
        // Find the worksheet mapping for this comment's worksheet
        var worksheetMapping = newWorkbookMappings.FirstOrDefault(w => w.WorksheetName == comment.WorksheetName);
        if (worksheetMapping == null)
        {
            return 1; // Fallback to row 1 if no mapping found
        }
        
        // Find all title row mappings
        var titleRowMappings = worksheetMapping.Mappings
            .Where(m => m.OpenApiRef.EndsWith("/TitleRow", StringComparison.OrdinalIgnoreCase))
            .Where(m => m.Row > 0) // Only row mappings, not cell mappings
            .OrderByDescending(m => m.Row) // Order by row descending to find closest title above
            .ToList();
            
        // Find the closest title row above the original comment
        var bestTitleRow = titleRowMappings
            .Where(m => m.Row < originalRow) // Only consider title rows above the comment
            .OrderByDescending(m => m.Row) // Get the closest one (highest row number below originalRow)
            .FirstOrDefault();
            
        if (bestTitleRow != null)
        {
            return bestTitleRow.Row;
        }
        
        // No title found, default to row 1
        return 1;
    }
}