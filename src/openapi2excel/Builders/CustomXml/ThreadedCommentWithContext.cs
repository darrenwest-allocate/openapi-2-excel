using System;
using System.Linq;
using DocumentFormat.OpenXml.Office2019.Excel.ThreadedComments;

namespace openapi2excel.core.Builders.CustomXml;

/// <summary>
/// Data structure to capture threaded comment with its worksheet context and OpenAPI mapping information.
/// </summary>
public class ThreadedCommentWithContext
{

    public ThreadedCommentWithContext() { }

    public ThreadedCommentWithContext(ThreadedComment comment, string worksheetName)
    {
        Comment = comment;
        WorksheetName = worksheetName;
    }

    public ThreadedComment Comment { get; set; } = null!;
    public string WorksheetName { get; set; } = string.Empty;
    public string OpenApiAnchor { get; set; } = string.Empty; // Added for mapping
    public string CellReference => Comment?.Ref?.Value ?? string.Empty;
    public string CommentText
    {
        get
        {
            // Extract text from the ThreadedComment XML structure
            var textElement = Comment?.Elements().FirstOrDefault(e => e.LocalName == "text");
            return textElement?.InnerText ?? string.Empty;
        }
    }
    public string CommentId => Comment?.Id?.Value ?? string.Empty;
    public DateTime? CreatedDate
    {
        get
        {
            var dtValue = Comment?.DT?.Value;
            if (dtValue != null && DateTime.TryParse(dtValue.ToString(), out var date))
                return date;
            return null;
        }
    }

    /// <summary>
    /// Gets the reply texts for this comment by finding other comments in the collection that reference this comment's ID as their parentId.
    /// This method recursively traverses the entire reply chain to get all nested replies.
    /// </summary>
    /// <param name="allComments">The complete collection of comments to search for replies</param>
    /// <returns>The text of all replies to this comment, including nested replies</returns>
    public IEnumerable<string> GetReplyTexts(IEnumerable<ThreadedCommentWithContext> allComments)
    {
        return GetReplies(allComments).Select(c => c.CommentText);
    }

    public IEnumerable<ThreadedCommentWithContext> GetReplies(IEnumerable<ThreadedCommentWithContext> allComments)
    {
        return GetReplyTextsRecursive(allComments, []);
    }

    /// <summary>
    /// Internal recursive helper that prevents infinite loops by tracking visited comment IDs.
    /// </summary>
    private IEnumerable<ThreadedCommentWithContext> GetReplyTextsRecursive(IEnumerable<ThreadedCommentWithContext> allComments, HashSet<string> visitedIds)
    {
        if (visitedIds.Contains(this.CommentId)) yield break; // Prevent infinite recursion by tracking visited comment IDs
        visitedIds.Add(CommentId);
        foreach (var reply in allComments.Where(c => c.Comment.ParentId?.Value == CommentId))
        {
            yield return reply;
            foreach (var nestedReply in reply.GetReplyTextsRecursive(allComments, visitedIds))
            {
                yield return nestedReply;
            }
        }
    }

    /// <summary>
    /// Determines if this comment has any replies by checking if any other comments reference this comment's ID as their parentId.
    /// </summary>
    /// <param name="allComments">The complete collection of comments to search for replies</param>
    /// <returns>True if this comment has at least one reply</returns>
    public bool HasReplies(IEnumerable<ThreadedCommentWithContext> allComments)
    {
        return allComments.Any(c => c.Comment.ParentId?.Value == this.CommentId);
    }
}
