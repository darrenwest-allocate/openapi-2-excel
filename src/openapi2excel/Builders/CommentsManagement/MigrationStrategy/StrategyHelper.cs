using ClosedXML.Excel;
using System.Collections.Generic;

namespace openapi2excel.core.Builders.CommentsManagement.MigrationStrategy;

/// <summary>
/// Shared helper methods used by multiple migration strategies and the main migration helper.
/// </summary>
public static class StrategyHelper
{
    /// <summary>
    /// Creates a legacy comment for Excel backward compatibility and visibility.
    /// Legacy comments are required for Excel to display threaded comments properly.
    /// </summary>
    public static void ReplicateSourceCommentOnNewWorksheet(
        IXLWorksheet newWorksheet,
        string cellReference,
        ThreadedCommentWithContext sourceComment)
    {
        var cell = newWorksheet.Cell(cellReference);
        var comment = cell.GetComment();
        comment ??= cell.CreateComment();
        comment.AddText(sourceComment.CommentText);
        comment.Author = sourceComment.Author;
    }

    /// <summary>
    /// Sets override target cells for a comment and all its replies to ensure threaded comments 
    /// migrate to the same location.
    /// </summary>
    public static void SetOverrideTargetCellForCommentAndReplies(
        ThreadedCommentWithContext comment,
        string targetCellReference,
        string targetWorksheetName,
        List<ThreadedCommentWithContext> allComments)
    {
        // Set override for the main comment
        comment.SetOverrideTargetCell(targetCellReference, targetWorksheetName);
        
        // Set override for all replies so they migrate to the same location
        var replies = comment.GetReplies(allComments);
        foreach (var reply in replies)
        {
            reply.SetOverrideTargetCell(targetCellReference, targetWorksheetName);
        }
    }
}