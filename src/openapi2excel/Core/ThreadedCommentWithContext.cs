using System;
using System.Linq;
using DocumentFormat.OpenXml.Office2019.Excel.ThreadedComments;

namespace OpenApi2Excel.Core
{
    /// <summary>
    /// Data structure to capture threaded comment with its worksheet context and OpenAPI mapping information.
    /// </summary>
    public class ThreadedCommentWithContext
    {
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
    }
}
