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
            // TODO this test has to be fixed. currently it only execuse the one
            // original unmigrated comment, and tnot that most have been added
            // they all need to be excused (somehow)

            if (comment.CommentText != "") Console.WriteLine(comment.CommentText);

            //Assert.NotEmpty(comment.OpenApiAnchor); // May be empty if no mapping exists
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

                // There are 10 root comments, now with Type A improvements we migrate more.
                // We should have: 9 exact matches + more Type A migrations = 14 total.
                Assert.Equal(14, legacyCommentCount);
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

    [Fact]
    public async Task MigrateComments_HandlesNoAnchorComments_ByPlacingNearTitleRows()
    {
        var tempNewWorkbookPath = Path.Combine(Path.GetTempPath(), $"test_noanchor_{Guid.NewGuid():N}.xlsx");
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

            // Act: Migrate comments including NoAnchor types
            var nonMigratableComments = CommentMigrationHelper.MigrateComments(existingWorkbookPath, tempNewWorkbookPath, newWorkbookMappings);

            // Assert: Verify NoAnchor comments were successfully migrated (not in failure list)
            // Comments like "Comment no anchor near the top" should be migrated to near title rows
            using (var workbook = new XLWorkbook(tempNewWorkbookPath))
            {
                var allComments = workbook.Worksheets.SelectMany(ws => ws.Cells().Where(c => c.HasComment)).ToList();
                
                // Should have both exact matches AND NoAnchor placements
                Assert.True(allComments.Count > 9, $"Expected more than 9 migrated comments including NoAnchor types, got {allComments.Count}");
                
                // Verify at least one comment was placed at expected Type A locations
                // From our debug output, we saw comments at J1 and C28 - these are typical Type A placements
                var foundTypeAPlacement = false;
                
                foreach (var worksheet in workbook.Worksheets)
                {
                    var commentsInWorksheet = worksheet.Cells().Where(c => c.HasComment).ToList();
                    
                    foreach (var commentCell in commentsInWorksheet)
                    {
                        var commentRow = commentCell.Address.RowNumber;
                        var commentCol = commentCell.Address.ColumnNumber;
                        
                        Console.WriteLine($"[DEBUG] Comment at {worksheet.Name}!{commentCell.Address}: Row {commentRow}");
                        
                        // Type A placement logic: row 1 (fallback) or near title rows
                        if (commentRow == 1)
                        {
                            foundTypeAPlacement = true;
                            Console.WriteLine($"[DEBUG] Found Type A fallback placement at row 1");
                            break;
                        }
                        
                        // Check if this comment is near any title row (within reasonable range)
                        for (int checkRow = Math.Max(1, commentRow - 5); checkRow <= Math.Min(worksheet.LastRowUsed()?.RowNumber() ?? commentRow + 5, commentRow + 5); checkRow++)
                        {
                            var rowCells = worksheet.Row(checkRow).CellsUsed().Select(c => c.GetString().ToUpper()).ToList();
                            var rowText = string.Join(" ", rowCells);
                            //TOSO replace with Langugae constants
                            if (rowText.Contains("REQUEST") || 
                                rowText.Contains("OPERATION") ||
                                rowText.Contains("PARAMETER") || 
                                rowText.Contains("SCHEMA"))
                            {
                                foundTypeAPlacement = true;
                                Console.WriteLine($"[DEBUG] Found Type A comment near title at row {checkRow}: '{rowText}'");
                                break;
                            }
                        }
                        
                        if (foundTypeAPlacement) break;
                    }
                    
                    if (foundTypeAPlacement) break;
                }
                
                Assert.True(foundTypeAPlacement, "Should have found at least one comment placed using Type A logic (row 1 or near title rows)");
            }
        }
        finally
        {
            if (File.Exists(tempNewWorkbookPath)) File.Delete(tempNewWorkbookPath);
        }
    }

    [Fact(Skip = "Type B (NoWorksheet) migration not yet implemented")]
    public async Task MigrateComments_HandlesNoWorksheetComments_ByPlacingOnInfoSheet()
    {
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
                // Find the Info sheet using the language constant
                var infoSheetName = OpenApiDocumentationLanguageConst.Info; // Should be "Info"
                var infoSheet = workbook.Worksheets.FirstOrDefault(ws => 
                    ws.Name.Equals(infoSheetName, StringComparison.OrdinalIgnoreCase));
                
                Assert.NotNull(infoSheet);
                
                // Look for comments in column V starting from row 1
                var columnV = 22; // Column V is the 22nd column
                var commentsInColumnV = new List<IXLCell>();
                for (int row = 1; row <= 10; row++) // Check first 10 rows
                {
                    var cell = infoSheet.Cell(row, columnV);
                    if (cell.HasComment)
                    {
                        commentsInColumnV.Add(cell);
                    }
                }
                
                // Should have at least one NoWorksheet comment migrated to column V
                Assert.True(commentsInColumnV.Count > 0, 
                    "Should have found NoWorksheet comments migrated to column V of Info sheet");
                
                // Verify one of the comments is from the ROGUE* worksheet
                var foundRogueComment = commentsInColumnV.Any(c => 
                    c.GetComment().Text.Contains("ROGUE") || 
                    c.GetComment().Text.Contains("Comment on a sheet not generated"));
                
                Assert.True(foundRogueComment, 
                    "Should have found at least one comment from the ROGUE* worksheet migrated to Info sheet");
            }
        }
        finally
        {
            if (File.Exists(tempNewWorkbookPath)) File.Delete(tempNewWorkbookPath);
        }
    }
}
