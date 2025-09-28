using ClosedXML.Excel;
using System;
using System.Collections.Generic;

namespace openapi2excel.core.Builders.CommentsManagement.MigrationStrategy;

/// <summary>
/// Strategy for migrating comments from worksheets that don't exist in the new workbook.
/// Places the comment on the Info sheet in column V (column 22).
/// </summary>
public class NoWorksheetCommentMigrationStrategy : ICommentMigrationStrategy
{
    public string StrategyName => "No Worksheet";

    public NoWorksheetCommentMigrationStrategy() { }

    public bool CanHandle(ThreadedCommentWithContext comment, IXLWorkbook workbook, List<WorksheetOpenApiMapping> newWorkbookMappings)
    {
        // Type B handles comments that cannot be placed on their original worksheet:
        // 1. Comments with no OpenAPI anchor and their source worksheet doesn't exist in new workbook
        // 2. Comments with an OpenAPI anchor that maps to a non-existent worksheet
        // 3. Comments with non-mappable OpenAPI anchors
        
        if (string.IsNullOrEmpty(comment.OpenApiAnchor))
        {
            return !workbook.Worksheets.TryGetWorksheet(comment.WorksheetName, out _);
        }
        
        var (targetMapping, worksheetName) = CommentTargetResolver.FindMatchingMapping(comment.OpenApiAnchor, newWorkbookMappings);
        if (targetMapping == null)
        {
            return true; // non-mappable anchor - NoWorksheet comment candidate
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
            // Get the Info sheet - this is where NoWorksheet comments go
            const string infoSheetName = "Info"; // Using the standard Info sheet name
            if (!workbook.Worksheets.TryGetWorksheet(infoSheetName, out var infoSheet))
            {
                return (false, CommentMigrationFailureReason.TargetWorksheetNotFound, 
                    $"Info sheet not found in new workbook for  NoWorksheet comment migration.");
            }

            const int targetColumn = 22; // Column V
            var targetCellReference = $"V{CellCollisionResolver.FindNextAvailableRowInColumn(infoSheet, targetColumn, processedCells)}";
            StrategyHelper.ReplicateSourceCommentOnNewWorksheet(infoSheet, targetCellReference, comment);
            
            var cellKey = $"{infoSheetName}:{targetCellReference}";
            processedCells.Add(cellKey);
            
            // Store the target cell reference for ThreadedComment processing
            StrategyHelper.SetOverrideTargetCellForCommentAndReplies(
                comment, targetCellReference, infoSheetName, allComments);
            
            return (true, CommentMigrationFailureReason.SuccessfullyMigratedAsNoWorksheetComment, "Successfully migrated comment from missing worksheet");
        }
        catch (Exception ex)
        {
            return (false, CommentMigrationFailureReason.UnexpectedErrorDuringMigration, 
                $"Error during NoWorksheet comment migration: {ex.Message}");
        }
    }
}