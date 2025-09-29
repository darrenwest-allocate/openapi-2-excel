using ClosedXML.Excel;
using System;
using System.Collections.Generic;

namespace openapi2excel.core.Builders.CommentsManagement.MigrationStrategy;

/// <summary>
/// Strategy for migrating comments that have OpenAPI anchors and can be mapped to their target locations.
/// </summary>
public class MappedCommentMigrationStrategy : ICommentMigrationStrategy
{
    public string StrategyName => "Anchored comment";

    public MappedCommentMigrationStrategy() { }

    public bool CanHandle(ThreadedCommentWithContext comment, IXLWorkbook workbook, List<WorksheetOpenApiMapping> newWorkbookMappings)
    {
        // Can handle comments that have OpenAPI anchors and can be mapped to existing worksheets
        if (string.IsNullOrEmpty(comment.OpenApiAnchor))
        {
            return false;
        }

        var (targetMapping, worksheetName) = CommentTargetResolver.FindMatchingMapping(comment.OpenApiAnchor, newWorkbookMappings);
        if (targetMapping == null)
        {
            return false; // Cannot map
        }

        return workbook.Worksheets.TryGetWorksheet(worksheetName, out _);
    }

    public (bool Success, CommentMigrationState? MigrationState) TryMigrate(
        ThreadedCommentWithContext comment,
        IXLWorkbook workbook,
        HashSet<string> processedCells,
        List<ThreadedCommentWithContext> allComments,
        List<WorksheetOpenApiMapping> newWorkbookMappings)
    {
        try
        {
            var (targetMapping, worksheetName) = CommentTargetResolver.FindMatchingMapping(comment.OpenApiAnchor, newWorkbookMappings);
            if (targetMapping == null)
            {
                return (false, CommentMigrationState.OpenApiAnchorNotFoundInNewWorkbook);
            }

            if (!workbook.Worksheets.TryGetWorksheet(worksheetName, out var worksheet))
            {
                return (false, CommentMigrationState.TargetWorksheetNotFound);
            }

            if (!CommentTargetResolver.TryGetTargetCell(comment, targetMapping, out string targetCellReference))
            {
                return (false, CommentMigrationState.TargetWorksheetNotFound);
            }

            // Create a unique key for this cell to avoid duplicate legacy comments
            var cellKey = $"{worksheetName}:{targetCellReference}";
            if (!processedCells.Contains(cellKey))
            {
                StrategyHelper.ReplicateSourceCommentOnNewWorksheet(worksheet, targetCellReference, comment);
                processedCells.Add(cellKey);
            }
            
            return (true, CommentMigrationState.Successful);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error migrating comment from '{comment.WorksheetName}' at '{comment.CellReference}' to Info sheet.\n{ex}");
            return (false, CommentMigrationState.UnexpectedErrorDuringMigration);
        }
    }
}