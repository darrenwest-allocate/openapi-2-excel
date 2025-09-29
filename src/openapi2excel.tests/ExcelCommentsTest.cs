using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Office2019.Excel.ThreadedComments;
using openapi2excel.core;
using openapi2excel.core.Common;
using openapi2excel.core.Builders.CommentsManagement;
using openapi2excel.core.Builders.WorksheetPartsBuilders.Common;

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

        const string commentWithoutAnchorFlag = "[NoAnchor]";
        foreach (var comment in annotatedComments)
        {
            Assert.NotEmpty(comment.WorksheetName);
            Assert.NotEmpty(comment.CellReference);
            if (comment.CommentText.Contains(commentWithoutAnchorFlag))
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
			using var workbook = new XLWorkbook(tempNewWorkbookPath);
			var legacyCommentCount = workbook.Worksheets.Sum(ws => ws.Cells().Count(c => c.HasComment));
			// - Original exact matches: 9
			// - (NoAnchor) comment migrations: 5 
			// - (NoWorksheet) comment migrations: 3
			// Total expected: 17
			Assert.Equal(17, legacyCommentCount);
		}
        finally
        {
            if (File.Exists(tempNewWorkbookPath)) File.Delete(tempNewWorkbookPath);
        }
    }

    private static async Task PrepareWorkbookWithMigratedComments(string tempNewWorkbookPath)
    {
        var existingWorkbookPath = "Sample/sample-api-gw-with-mappings.xlsx";
        var sampleOpenApiPath = "Sample/sample-api-gw.json";
        await OpenApiDocumentationGenerator.GenerateDocumentation(
            sampleOpenApiPath,
            tempNewWorkbookPath,
            new OpenApiDocumentationOptions { IncludeMappings = true }
        );

        var newWorkbookMappings = ExcelOpenXmlHelper.ExtractCustomXmlMappingsFromWorkbook(tempNewWorkbookPath);
        CommentMigrationHelper.MigrateComments(existingWorkbookPath, tempNewWorkbookPath, newWorkbookMappings);
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

    [Fact]
    public async Task MigrateComments_HandlesNoAnchorComments_ByPlacingNearTitleRows()
    {
        var tempNewWorkbookPath = Path.Combine(Path.GetTempPath(), $"test_noanchor_{Guid.NewGuid():N}.xlsx");
        try
        {
            var sampleOpenApiPath = "Sample/sample-api-gw.json";
            await OpenApiDocumentationGenerator.GenerateDocumentation(
                sampleOpenApiPath,
                tempNewWorkbookPath,
                new OpenApiDocumentationOptions { IncludeMappings = true }
            );

            var headingTitles = new string[] {
                WorksheetLanguage.Schema.Title,
                WorksheetLanguage.Operations.Title,
                WorksheetLanguage.Request.Title,
                WorksheetLanguage.Response.Title,
                WorksheetLanguage.Parameters.Title
            };

            var existingWorkbookPath = "Sample/sample-api-gw-with-mappings.xlsx";
            var newWorkbookMappings = ExcelOpenXmlHelper.ExtractCustomXmlMappingsFromWorkbook(tempNewWorkbookPath);
            // Act: Migrate comments including NoAnchor types
            CommentMigrationHelper.MigrateComments(existingWorkbookPath, tempNewWorkbookPath, newWorkbookMappings);

            // Assert: Verify NoAnchor comments were successfully migrated
            const string commentWithoutAnchorFlag = "[NoAnchor]";
            var allComments = ExcelOpenXmlHelper.ExtractAndAnnotateAllComments(tempNewWorkbookPath);
            var commentsWithoutAnchor = allComments.Where(c => c.CommentText.Contains(commentWithoutAnchorFlag)).ToList();
            Assert.True(commentsWithoutAnchor.Count > 7, $"Expected more than 7 migrated NoAnchor comments, got {commentsWithoutAnchor.Count}");

            // Verify comments were placed at expected NoAnchor (alternative) location
            using var workbook = new XLWorkbook(tempNewWorkbookPath);
            foreach (var noAnchorComment in commentsWithoutAnchor)
            {
                var commentRow = CommentTargetResolver.ExtractRowFromCellReference(noAnchorComment.CellReference);
                var commentColumn = CommentTargetResolver.ExtractColumnFromCellReference(noAnchorComment.CellReference);
                if (commentColumn == "V") // "No Worksheet" comments
                {
                    continue;
                }
                var worksheet = workbook.Worksheets.First(ws => ws.Name.Equals(noAnchorComment.WorksheetName));
                var textOfFirstCellForRow = worksheet.Row( commentRow ).CellsUsed().FirstOrDefault()?.GetText() ?? "";
                if (headingTitles.Contains(textOfFirstCellForRow))
                {
                    continue;
                }
                if (commentRow == 1) // fallback row
                {
                    continue;
                }
                Assert.Fail($"{commentWithoutAnchorFlag} comment at {noAnchorComment.WorksheetName}!{noAnchorComment.CellReference} (row {commentRow}) is not near a title row. Found text: '{textOfFirstCellForRow}'");
            }
        }
        finally
        {
            if (File.Exists(tempNewWorkbookPath)) File.Delete(tempNewWorkbookPath);
        }
    }

    [Fact]
    public async Task MigrateComments_HandlesNoWorksheetComments_ByPlacingOnInfoSheet()
    {
        const string commentWithoutAnchorFlag = "[NoAnchor]";
        var tempNewWorkbookPath = Path.Combine(Path.GetTempPath(), $"test_noworksheet_{Guid.NewGuid():N}.xlsx");
        try
        {
            // Arrange: Create new workbook with mappings
            var sampleOpenApiPath = "Sample/sample-api-gw.json";
            await OpenApiDocumentationGenerator.GenerateDocumentation(
                sampleOpenApiPath,
                tempNewWorkbookPath,
                new OpenApiDocumentationOptions { IncludeMappings = true }
            );

            var existingWorkbookPath = "Sample/sample-api-gw-with-mappings.xlsx";
            var newWorkbookMappings = ExcelOpenXmlHelper.ExtractCustomXmlMappingsFromWorkbook(tempNewWorkbookPath);

            // Act: Migrate comments including NoWorksheet types
            var nonMigratableComments = CommentMigrationHelper.MigrateComments(existingWorkbookPath, tempNewWorkbookPath, newWorkbookMappings);

            // Assert: Verify NoWorksheet comments were successfully migrated to Info sheet
            using (var workbook = new XLWorkbook(tempNewWorkbookPath))
            {
                var infoSheetName = OpenApiDocumentationLanguageConst.Info;
                var infoSheet = workbook.Worksheets.FirstOrDefault(ws => ws.Name.Equals(infoSheetName, StringComparison.OrdinalIgnoreCase));
                Assert.NotNull(infoSheet);
                
                // Look for comments in column V using proper OpenXML extraction
                // (ClosedXML HasComment cannot detect threaded comments properly)
                var allComments = ExcelOpenXmlHelper.ExtractAndAnnotateAllComments(tempNewWorkbookPath);
                
                // Filter for comments on Info sheet in column V
                var infoColumnVComments = allComments.Where(c => c.WorksheetName.Equals(infoSheetName) && c.CellReference.StartsWith("V")).ToList();
                Assert.True(infoColumnVComments.Count > 2, "Should have found NoWorksheet comments migrated to column V of Info sheet");
                
                // Verify one of the comments is from the ROGUE* worksheet content
                var foundRogueComment = infoColumnVComments.All(c => c.CommentText.Contains(commentWithoutAnchorFlag));
                
                Assert.True(foundRogueComment, "All comments from the ROGUE* worksheet migrated to Info sheet");
            }
        }
        finally
        {
            if (File.Exists(tempNewWorkbookPath)) File.Delete(tempNewWorkbookPath);
        }
    }
}
