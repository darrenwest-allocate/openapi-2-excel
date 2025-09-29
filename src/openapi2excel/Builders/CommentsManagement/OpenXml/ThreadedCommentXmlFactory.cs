using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace openapi2excel.core.Builders.CommentsManagement;

/// <summary>
/// Factory for creating threaded comment XML content using the official Microsoft SDK patterns.
/// </summary>
public class ThreadedCommentXmlFactory
{

    public ThreadedCommentXmlFactory() { }

    /// <summary>
    /// Creates WorksheetThreadedCommentsPart with proper GUID matching to legacy comments.
    /// Uses manual XML generation to ensure correct 2018 schema format as expected by tests.
    /// </summary>
    public void CreateThreadedCommentsXmlContent(
        WorksheetThreadedCommentsPart threadedCommentsPart,
        List<ThreadedCommentWithContext> comments,
        List<WorksheetOpenApiMapping> newWorkbookMappings,
        Dictionary<string, string> idMapping)
    {
        var xml = CreateThreadedCommentsXml(comments, newWorkbookMappings, idMapping);
        
        // Write the XML content to the part
        using (var stream = threadedCommentsPart.GetStream(FileMode.Create))
        using (var writer = new StreamWriter(stream))
        {
            writer.Write(xml);
        }
    }

    /// <summary>
    /// Creates the threaded comments XML content manually to ensure correct 2018 schema format.
    /// </summary>
    private string CreateThreadedCommentsXml(
        List<ThreadedCommentWithContext> comments,
        List<WorksheetOpenApiMapping> newWorkbookMappings,
        Dictionary<string, string> idMapping)
    {
        var xmlBuilder = new StringBuilder();
        
        // XML header and root element with correct 2018 namespaces
        xmlBuilder.AppendLine("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        xmlBuilder.AppendLine("<ThreadedComments xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments\" xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
        
        // Process root comments and their replies
        foreach (var rootComment in comments.Where(c => c.IsRootComment))
        {
            var (targetMapping, _) = CommentTargetResolver.FindMatchingMapping(rootComment.OpenApiAnchor, newWorkbookMappings);
            if (targetMapping == null) continue;

            if (!CommentTargetResolver.TryGetTargetCell(rootComment, targetMapping, out string targetCellReference)) continue;

            // Create the root threaded comment
            var rootId = Guid.NewGuid().ToString("B").ToUpper(); // Format: {GUID}
            var rootDateTime = rootComment.Comment.DT?.Value.ToString("yyyy-MM-ddTHH:mm:ss.ff") ?? DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.ff");
            var personId = rootComment.Comment.PersonId?.Value ?? "";
            
            var doneAttr = rootComment.Comment.Done?.Value == true ? " done=\"1\"" : "";
            
            xmlBuilder.AppendLine($"<threadedComment ref=\"{targetCellReference}\" dT=\"{rootDateTime}\" personId=\"{personId}\" id=\"{rootId}\"{doneAttr}>");
            xmlBuilder.AppendLine($"<text>{System.Security.SecurityElement.Escape(rootComment.CommentText)}</text>");
            xmlBuilder.AppendLine("</threadedComment>");
            
            // Track the ID mapping
            if (!string.IsNullOrEmpty(rootComment.CommentId))
            {
                idMapping[rootComment.CommentId] = rootId;
            }

            // Add replies to this root comment
            var replies = rootComment.GetReplies(comments).ToList();
            foreach (var reply in replies)
            {
                var replyId = Guid.NewGuid().ToString("B").ToUpper();
                var replyDateTime = reply.Comment.DT?.Value.ToString("yyyy-MM-ddTHH:mm:ss.ff") ?? DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.ff");
                var replyPersonId = reply.Comment.PersonId?.Value ?? "";
                
                xmlBuilder.AppendLine($"<threadedComment ref=\"{targetCellReference}\" dT=\"{replyDateTime}\" personId=\"{replyPersonId}\" id=\"{replyId}\" parentId=\"{rootId}\">");
                xmlBuilder.AppendLine($"<text>{System.Security.SecurityElement.Escape(reply.CommentText)}</text>");
                xmlBuilder.AppendLine("</threadedComment>");
                
                // Track the reply ID mapping
                if (!string.IsNullOrEmpty(reply.CommentId))
                {
                    idMapping[reply.CommentId] = replyId;
                }
            }
        }
        
        xmlBuilder.AppendLine("</ThreadedComments>");
        return xmlBuilder.ToString();
    }
}