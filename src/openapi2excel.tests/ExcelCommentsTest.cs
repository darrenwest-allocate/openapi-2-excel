using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Office2019.Excel.ThreadedComments;
using DocumentFormat.OpenXml.Spreadsheet;
using openapi2excel.core;
using openapi2excel.core.Common;
using openapi2excel.core.Builders.CustomXml;

namespace OpenApi2Excel.Tests;

public class ExcelCommentsTest
{
    [Fact]
    public void ExtractUnresolvedComments_FromOldWorkbook_FindsOnlyUnresolvedComments()
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
    public void ExtractCustomXmlMappings_FromOldWorkbook_LinksCommentsToOpenApiAnchors()
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
    public async Task MigratedComments_ShouldCreateLegacyCommentsForExcelVisibility()
    {
        // Arrange: Use the existing sample files for migration
        var existingWorkbookPath = "Sample/sample-api-gw-with-mappings.xlsx";
        var tempNewWorkbookPath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid():N}.xlsx");
        
        try
        {
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
            CommentMigrationHelper.MigrateComments(existingWorkbookPath, tempNewWorkbookPath, newWorkbookMappings);
            
            // Check if any migration occurred by looking for threaded comments
            bool hasThreadedComments = false;
            using (var document = SpreadsheetDocument.Open(tempNewWorkbookPath, false))
            {
                hasThreadedComments = document.WorkbookPart!.WorksheetParts
                    .Any(ws => ws.GetPartsOfType<WorksheetThreadedCommentsPart>().Any());
            }
            
            Assert.True(hasThreadedComments, "No threaded comments found after migration - migration may have failed");
            
            // Assert: Check that legacy comments part was created
            using (var document = SpreadsheetDocument.Open(tempNewWorkbookPath, false))
            {
                var workbookPart = document.WorkbookPart;
                Assert.NotNull(workbookPart);
                
                // Find worksheets that should have migrated comments
                var worksheetsWithComments = workbookPart.WorksheetParts
                    .Where(ws => ws.GetPartsOfType<WorksheetThreadedCommentsPart>().Any());
                
                Assert.True(worksheetsWithComments.Any(), "Should have worksheets with threaded comments");
                
                // Verify each worksheet with threaded comments also has legacy comments
                foreach (var worksheetPart in worksheetsWithComments)
                {
                    var threadedCommentsPart = worksheetPart.GetPartsOfType<WorksheetThreadedCommentsPart>().First();
                    var threadedCommentsCount = threadedCommentsPart.ThreadedComments.Elements<ThreadedComment>().Count();
                    
                    // CRITICAL: Each worksheet with threaded comments must have corresponding legacy comments
                    var legacyCommentsPart = worksheetPart.GetPartsOfType<WorksheetCommentsPart>().FirstOrDefault();
                    Assert.NotNull(legacyCommentsPart);
                    Assert.NotNull(legacyCommentsPart.Comments);
                    Assert.NotNull(legacyCommentsPart.Comments.CommentList);
                    Assert.NotNull(legacyCommentsPart.Comments.Authors);
                    
                    var legacyCommentsCount = legacyCommentsPart.Comments.CommentList.Elements<Comment>().Count();
                    
                    // Each threaded comment should have a corresponding legacy comment
                    Assert.Equal(threadedCommentsCount, legacyCommentsCount);
                    
                    // Legacy comments should have the "[Threaded comment]" format
                    var legacyComments = legacyCommentsPart.Comments.CommentList.Elements<Comment>();
                    foreach (var legacyComment in legacyComments)
                    {
                        Assert.NotNull(legacyComment.Reference?.Value);
                        Assert.NotNull(legacyComment.AuthorId?.Value);
                        Assert.NotNull(legacyComment.CommentText);
                        
                        // Text should be "[Threaded comment]" for threaded comment compatibility
                        var commentText = legacyComment.CommentText.InnerText;
                        Assert.Equal("[Threaded comment]", commentText);
                    }
                    
                    // Authors should be populated
                    var authors = legacyCommentsPart.Comments.Authors.Elements<Author>();
                    Assert.True(authors.Any(), "Should have at least one author");
                }
            }
        }
        finally
        {
            // Cleanup
            if (File.Exists(tempNewWorkbookPath))
            {
                File.Delete(tempNewWorkbookPath);
            }
        }
    }

}
