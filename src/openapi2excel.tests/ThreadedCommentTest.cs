
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using openapi2excel.core.Common;

namespace OpenApi2Excel.Tests;

public class ThreadedCommentTest
{
	[Fact]
	public void Can_List_Replies()
	{
		const string existingWorkbook = "Sample/sample-api-gw-with-mappings.xlsx";
		const string discussionStart = "A comment about a description";
		const string discussionEnd = "This is the end of the discussion";
		var allComments = ExcelOpenXmlHelper.ExtractAndAnnotateAllComments(existingWorkbook);
		var originalComment = allComments.FirstOrDefault(c => c.CommentText.Contains(discussionStart));

		Assert.NotNull(originalComment);
		Assert.True(originalComment.HasReplies(allComments));

		var replyTexts = originalComment.GetReplyTexts(allComments).ToList();
		Assert.True(replyTexts.Count > 2, "Should have found reply texts");
		Assert.Contains(discussionEnd, replyTexts);
	}

}