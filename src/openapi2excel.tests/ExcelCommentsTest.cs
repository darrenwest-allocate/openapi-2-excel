using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Office2019.Excel.ThreadedComments;
using DocumentFormat.OpenXml.Spreadsheet;
using openapi2excel.core;
using openapi2excel.core.Common;
using openapi2excel.core.Builders.CommentsManagement;

namespace OpenApi2Excel.Tests;

public class ExcelCommentsTest
{
    [Fact]
    public void ExtractUnresolvedComments_FindsOnlyUnresolvedComments_FromOldWorkbook()
    {
        var samplePath = "Sample/sample-api-gw.xlsx"; // This file should contain both types

        // Act: Extract threaded comments with worksheet context using the new helper
        var unresolvedComments = ExcelOpenXmlHelper.ExtractUnresolvedThreadedCommentsFromWorkbook(samplePath);

        Assert.NotEmpty(unresolvedComments);

        Assert.All(unresolvedComments, commentWithContext =>
        {
            Assert.NotEqual("1", commentWithContext.Comment.Done);
            Assert.NotNull(commentWithContext.WorksheetName);
            Assert.NotEmpty(commentWithContext.WorksheetName);
        });

        const string resolvedComment = "A comment that we will marked resolved";
        const string unresolvedComment = "A comment in a field that was not populated by the open api json";

        Assert.Contains(unresolvedComments, c => c.CommentText.Contains(unresolvedComment));
        Assert.DoesNotContain(unresolvedComments, c => c.CommentText.Contains(resolvedComment));
    }

    [Fact]
    public void ExtractCustomXmlMappings_LinksCommentsToOpenApiAnchors_FromOldWorkbook()
    {
        var samplePath = "Sample/sample-api-gw-with-mappings.xlsx";

        var annotatedComments = ExcelOpenXmlHelper.ExtractAndAnnotateUnresolvedComments(samplePath);

        Assert.NotNull(annotatedComments);
        Assert.True(annotatedComments.Count > 10);

        const string unMappableComment = "This is a comment that is unlikely to be mapped because it is on a blank row";
        foreach (var comment in annotatedComments)
        {
            Assert.NotEmpty(comment.WorksheetName);
            Assert.NotEmpty(comment.CellReference);
            if (comment.CommentText.Contains(unMappableComment))
            {
                Assert.Empty(comment.OpenApiAnchor);
                continue;
            }
            Assert.NotEmpty(comment.OpenApiAnchor); // May be empty if no mapping exists
        }
    }

    [Fact]
    public async Task MigratedComments_CreatesLegacyComments()
    {
        var tempNewWorkbookPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.xlsx");
        try
        {
            await PrepareWorkbookWithMigratedComments(tempNewWorkbookPath);

            // Assert: Check that legacy comments were created
            using (var workbook = new XLWorkbook(tempNewWorkbookPath))
            {
                var legacyCommentCount = workbook.Worksheets.Sum(ws => ws.Cells().Count(c => c.HasComment));

                // There are 10 root comments, one of which is unmappable. So 9 comments should be migrated.
                Assert.Equal(9, legacyCommentCount);
            }
        }
        finally
        {
            // Cleanup
            if (File.Exists(tempNewWorkbookPath)) File.Delete(tempNewWorkbookPath);
        }
    }

    private static async Task PrepareWorkbookWithMigratedComments(string tempNewWorkbookPath)
    {
        var existingWorkbookPath = "Sample/sample-api-gw-with-mappings.xlsx";

        // Create a sample new workbook using the OpenApiDocumentationGenerator
        var sampleOpenApiPath = "Sample/sample-api-gw.json";
        await OpenApiDocumentationGenerator.GenerateDocumentation(
            sampleOpenApiPath,
            tempNewWorkbookPath,
            new OpenApiDocumentationOptions { IncludeMappings = true }
        );

        // Extract mappings from the new workbook for migration
        var newWorkbookMappings = ExcelOpenXmlHelper.ExtractCustomXmlMappingsFromWorkbook(tempNewWorkbookPath);

        // Act: Migrate comments to the new workbook
        var nonMigratableComments = CommentMigrationHelper.MigrateComments(existingWorkbookPath, tempNewWorkbookPath, newWorkbookMappings);
    }

    [Fact]
    public async Task MigratedComments_IncludesThreadedComments()
    {
        var tempNewWorkbookPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.xlsx");
        try
        {
            await PrepareWorkbookWithMigratedComments(tempNewWorkbookPath);

            using var spreadsheetDocument = SpreadsheetDocument.Open(tempNewWorkbookPath, false);
            var workbookPart = spreadsheetDocument.WorkbookPart;
            Assert.NotNull(workbookPart);

            var allThreadedComments = new List<ThreadedComment>();
            foreach (var worksheetPart in workbookPart.WorksheetParts)
            {
                var threadedCommentsPart = worksheetPart.GetPartsOfType<WorksheetThreadedCommentsPart>().FirstOrDefault();
                if (threadedCommentsPart?.ThreadedComments != null)
                {
                    allThreadedComments.AddRange(threadedCommentsPart.ThreadedComments.Elements<ThreadedComment>());
                }
            }

            // A "discussion" is a root comment with at least one reply.
            // We convert to ThreadedCommentWithContext to use its helper methods.
            var allCommentsWithContext = allThreadedComments
                .Select(c => new ThreadedCommentWithContext { Comment = c })
                .ToList();

            var rootComments = allCommentsWithContext.Where(c => c.IsRootComment).ToList();
            var discussionCount = rootComments.Count(rc => rc.HasReplies(allCommentsWithContext));

            // The sample file has at least two distinct comment threads.
            Assert.True(discussionCount >= 2, $"Expected at least 2 discussions, but found {discussionCount}.");
        }
        finally
        {
            // Cleanup
            if (File.Exists(tempNewWorkbookPath)) File.Delete(tempNewWorkbookPath);
        }
    }

    [Fact]
    public void ListReplies_ReturnsAllReplies_ForDiscussion()
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
