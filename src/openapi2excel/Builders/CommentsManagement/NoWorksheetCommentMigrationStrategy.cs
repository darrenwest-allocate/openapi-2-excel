using ClosedXML.Excel;
using openapi2excel.core.Common;
using System;
using System.Collections.Generic;

namespace openapi2excel.core.Builders.CommentsManagement;

/// <summary>
/// Strategy for migrating comments from worksheets that don't exist in the new workbook.
/// Places the comment on the Info sheet in column V (column 22).
/// </summary>
public class NoWorksheetCommentMigrationStrategy : ICommentMigrationStrategy
{
    private readonly CellCollisionResolver _collisionResolver;
    private readonly CommentTargetResolver _targetResolver;

    public string StrategyName => "Type B (NoWorksheet)";

    public NoWorksheetCommentMigrationStrategy(CellCollisionResolver collisionResolver, CommentTargetResolver targetResolver)
    {
        _collisionResolver = collisionResolver ?? throw new ArgumentNullException(nameof(collisionResolver));
        _targetResolver = targetResolver ?? throw new ArgumentNullException(nameof(targetResolver));
    }

    public bool CanHandle(ThreadedCommentWithContext comment, IXLWorkbook workbook, List<WorksheetOpenApiMapping> newWorkbookMappings)
    {
        // Type B handles comments that cannot be placed on their original worksheet:
        // 1. Comments with no OpenAPI anchor and their source worksheet doesn't exist in new workbook
        // 2. Comments with an OpenAPI anchor that maps to a non-existent worksheet
        // 3. Comments with unmappable OpenAPI anchors
        
        if (string.IsNullOrEmpty(comment.OpenApiAnchor))
        {
            // No anchor - check if source worksheet exists
            return !workbook.Worksheets.TryGetWorksheet(comment.WorksheetName, out _);
        }
        
        // Has anchor - check if it maps to an existing worksheet
        var (targetMapping, worksheetName) = _targetResolver.FindMatchingMapping(comment.OpenApiAnchor, newWorkbookMappings);
        if (targetMapping == null)
        {
            return true; // Unmappable anchor - Type B candidate
        }
        
        return !workbook.Worksheets.TryGetWorksheet(worksheetName, out _);
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
            // Get the Info sheet - this is where Type B comments go
            const string infoSheetName = "Info"; // Using the standard Info sheet name
            if (!workbook.Worksheets.TryGetWorksheet(infoSheetName, out var infoSheet))
            {
                return (false, CommentMigrationFailureReason.TargetWorksheetNotFound, 
                    $"Info sheet not found in new workbook for Type B migration.");
            }

            // Type B comments go in column V (column 22)
            const int targetColumn = 22; // Column V
            
            // Find the next available row in column V
            var targetRow = _collisionResolver.FindNextAvailableRowInColumn(infoSheet, targetColumn, processedCells);
            
            var targetCellReference = $"V{targetRow}";
            
            // Create the legacy comment on the Info sheet
            CommentMigrationSharedHelper.ReplicateSourceCommentOnNewWorksheet(infoSheet, targetCellReference, comment);
            
            // Mark cell as processed
            var cellKey = $"{infoSheetName}:{targetCellReference}";
            processedCells.Add(cellKey);
            
            // Store the target cell reference for ThreadedComment processing
            CommentMigrationSharedHelper.SetOverrideTargetCellForCommentAndReplies(
                comment, targetCellReference, infoSheetName, allComments);
            
            return (true, CommentMigrationFailureReason.SuccessfullyMigratedAsNoWorksheetComment, "Successfully migrated comment from missing worksheet");
        }
        catch (Exception ex)
        {
            return (false, CommentMigrationFailureReason.UnexpectedErrorDuringMigration, 
                $"Error during Type B migration: {ex.Message}");
        }
    }
}