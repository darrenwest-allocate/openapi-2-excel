using ClosedXML.Excel;
using System.Collections.Generic;

namespace openapi2excel.core.Builders.CommentsManagement.MigrationStrategy;

/// <summary>
/// Strategy interface for different types of comment migration approaches.
/// </summary>
public interface ICommentMigrationStrategy
{
    /// <summary>
    /// Gets the strategy name for logging and debugging purposes.
    /// </summary>
    string StrategyName { get; }
    
    /// <summary>
    /// Determines if this strategy can handle the given comment based on its characteristics.
    /// </summary>
    /// <param name="comment">The comment to evaluate</param>
    /// <param name="workbook">The target workbook</param>
    /// <param name="newWorkbookMappings">OpenAPI mappings for the new workbook</param>
    /// <returns>True if this strategy can handle the comment, false otherwise</returns>
    bool CanHandle(ThreadedCommentWithContext comment, IXLWorkbook workbook, List<WorksheetOpenApiMapping> newWorkbookMappings);
    
    /// <summary>
    /// Attempts to migrate the comment using this strategy.
    /// </summary>
    /// <param name="comment">The comment to migrate</param>
    /// <param name="workbook">The target workbook</param>
    /// <param name="processedCells">Set of cells that already have legacy comments</param>
    /// <param name="allComments">All comments in the migration batch (for reply handling)</param>
    /// <param name="newWorkbookMappings">OpenAPI mappings for the new workbook</param>
    /// <returns>Result indicating success/failure and details</returns>
    (bool Success, CommentMigrationFailureReason? FailureReason, string? ErrorDetails) TryMigrate(
        ThreadedCommentWithContext comment,
        IXLWorkbook workbook,
        HashSet<string> processedCells,
        List<ThreadedCommentWithContext> allComments,
        List<WorksheetOpenApiMapping> newWorkbookMappings);
}