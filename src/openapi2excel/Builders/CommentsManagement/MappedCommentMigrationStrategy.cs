using ClosedXML.Excel;
using openapi2excel.core.Common;
using System;
using System.Collections.Generic;

namespace openapi2excel.core.Builders.CommentsManagement;

/// <summary>
/// Strategy for migrating comments that have OpenAPI anchors and can be mapped to their target locations.
/// </summary>
public class MappedCommentMigrationStrategy : ICommentMigrationStrategy
{
    private readonly CommentTargetResolver _targetResolver;

    public string StrategyName => "Mapped Comment";

    public MappedCommentMigrationStrategy(CommentTargetResolver targetResolver)
    {
        _targetResolver = targetResolver ?? throw new ArgumentNullException(nameof(targetResolver));
    }

    public bool CanHandle(ThreadedCommentWithContext comment, IXLWorkbook workbook, List<WorksheetOpenApiMapping> newWorkbookMappings)
    {
        // Can handle comments that have OpenAPI anchors and can be mapped to existing worksheets
        if (string.IsNullOrEmpty(comment.OpenApiAnchor))
        {
            return false;
        }

        var (targetMapping, worksheetName) = _targetResolver.FindMatchingMapping(comment.OpenApiAnchor, newWorkbookMappings);
        if (targetMapping == null)
        {
            return false; // Cannot map
        }

        return workbook.Worksheets.TryGetWorksheet(worksheetName, out _);
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
            var (targetMapping, worksheetName) = _targetResolver.FindMatchingMapping(comment.OpenApiAnchor, newWorkbookMappings);
            if (targetMapping == null)
            {
                return (false, CommentMigrationFailureReason.OpenApiAnchorNotFoundInNewWorkbook, 
                    $"Anchor '{comment.OpenApiAnchor}' not found in new workbook mappings.");
            }

            if (!workbook.Worksheets.TryGetWorksheet(worksheetName, out var worksheet))
            {
                return (false, CommentMigrationFailureReason.TargetWorksheetNotFound, 
                    $"Worksheet '{worksheetName}' not found in the new workbook.");
            }

            if (!_targetResolver.TryGetTargetCell(comment, targetMapping, out string targetCellReference))
            {
                return (false, CommentMigrationFailureReason.TargetWorksheetNotFound, 
                    "Could not determine target cell for migration.");
            }

            // Create a unique key for this cell to avoid duplicate legacy comments
            var cellKey = $"{worksheetName}:{targetCellReference}";
            if (!processedCells.Contains(cellKey))
            {
                CommentMigrationSharedHelper.ReplicateSourceCommentOnNewWorksheet(worksheet, targetCellReference, comment);
                processedCells.Add(cellKey);
            }
            
            return (true, null, null);
        }
        catch (Exception ex)
        {
            return (false, CommentMigrationFailureReason.UnexpectedErrorDuringMigration, ex.Message);
        }
    }
}