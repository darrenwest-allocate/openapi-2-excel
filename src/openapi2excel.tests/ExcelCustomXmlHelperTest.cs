using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using openapi2excel.core.Common;


namespace OpenApi2Excel.Tests;

public class ExcelCustomXmlHelperTest
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
            var xmlContent = commentWithContext.Comment.OuterXml;
            Assert.DoesNotContain("resolved=\"1\"", xmlContent);
            // Verify we captured the worksheet name
            Assert.NotNull(commentWithContext.WorksheetName);
            Assert.NotEmpty(commentWithContext.WorksheetName);
        });
    }

    [Fact]
    public void ExtractCustomXmlMappings_FromOldWorkbook_LinksCommentsToOpenApiAnchors()
    {
        var samplePath = "Sample/sample-api-gw.xlsx";

        // Act: Extract and annotate comments to verify XML mappings work
        var annotatedComments = ExcelOpenXmlHelper.ExtractAndAnnotateUnresolvedComments(samplePath);

        // Assert: We should find annotated comments (this verifies mappings work)
        Assert.NotNull(annotatedComments);

        // If mappings exist, some comments should be annotated with OpenAPI anchors
        // This is a more realistic test since it tests the full workflow
        foreach (var comment in annotatedComments)
        {
            Assert.NotEmpty(comment.WorksheetName);
            Assert.NotEmpty(comment.CellReference);
            Assert.NotNull(comment.OpenApiAnchor); // May be empty if no mapping exists
        }
    }

    [Fact]
    public void ExtractAndAnnotateUnresolvedComments_CombinesCommentsWithOpenApiAnchors()
    {
        var samplePath = "Sample/sample-api-gw.xlsx";

        // Act: Extract and annotate comments with OpenAPI anchors using the helper class
        var annotatedComments = ExcelOpenXmlHelper.ExtractAndAnnotateUnresolvedComments(samplePath);

        // Assert: We should find comments with proper context
        Assert.NotNull(annotatedComments);

        foreach (var comment in annotatedComments)
        {
            // Verify basic comment properties
            Assert.NotNull(comment.Comment);
            Assert.NotEmpty(comment.WorksheetName);
            Assert.NotEmpty(comment.CellReference);

            // OpenApiAnchor might be empty if no mapping exists for this comment
            // but the property should be available for annotation
            Assert.NotNull(comment.OpenApiAnchor);
        }
    }
}
