
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using openapi2excel.core;
using openapi2excel.core.Builders;
using openapi2excel.core.Builders.CustomXml;
using Microsoft.OpenApi.Readers;
using ClosedXML.Excel;
using openapi2excel.core.Common;

namespace OpenApi2Excel.Tests;

public class OpenApiDocumentationGeneratorTest
{
   [Fact]
   public async Task GenerateDocumentation_create_excel_file_for_correct_openapi_document()
   {
      const string inputFIle = "Sample/Sample1.yaml";
      const string outputFile = "output.xlsx";
      await using var file = File.OpenRead(inputFIle);

      await OpenApiDocumentationGenerator.GenerateDocumentation(file, outputFile, new OpenApiDocumentationOptions());
      Assert.True(File.Exists(outputFile));
   }

   [Fact]
   public void CustomXmlMappingParts_AreWrittenAndReadPerWorksheet()
   {
      // Arrange
      var tempFile = Path.GetTempFileName().Replace(".tmp", ".xlsx");
      var worksheetName = "TestSheet";

      // Use a builder class to create the worksheet mapping
      var mappings = new[] {
           ("A5", "paths./pets.get.responses.200"),
           ("B12", "components.parameters.PetId")
        }.Select(i => new CellOpenApiMapping() { Cell = i.Item1, OpenApiRef = i.Item2 }).ToList();

      // Use the serializer class
      var customXmlContent = WorksheetOpenApiMapping.Serialize([new WorksheetOpenApiMapping(worksheetName) { Mappings = mappings }]);

      // Act: Write and read back
      ExcelCustomXmlHelper.WriteCustomXmlMapping(tempFile, worksheetName, customXmlContent);
      var actualDoc = ExcelCustomXmlHelper.ReadAllCustomXmlMappings(tempFile)[worksheetName];
      var expectedDoc = XDocument.Parse(customXmlContent);
      Assert.True(XNode.DeepEquals(expectedDoc, actualDoc), "XML content does not match after round-trip.");
   }

   [Fact]
   public async Task OperationWorksheetBuilder_CreatesExpectedMappings()
   {
      const string openApiFile = "Sample/sample-api-gw.json";
      var readResult = await new OpenApiStreamReader().ReadAsync(File.OpenRead(openApiFile));
      using var workbook = new XLWorkbook();
      var worksheetBuilder = new OperationWorksheetBuilder(workbook, new OpenApiDocumentationOptions());
      var path = readResult.OpenApiDocument.Paths.First();
      var operation = path.Value.Operations.First();
      WorksheetOpenApiMapping.AllWorksheetMappings.Clear();
      var worksheet = worksheetBuilder.Build(path.Key, path.Value, operation.Key, operation.Value);
      var mappings = worksheetBuilder.CurrentWorksheetMapping;
      var customXmlContent = WorksheetOpenApiMapping.Serialize(mappings);

      // Verify that the worksheet was created and mappings were added
      Assert.NotNull(worksheet);
      Assert.Equal(worksheet.Name, mappings.WorksheetName);
      Assert.NotEmpty(mappings.Mappings);

      const string sampleWorkbookFirstSheetMappings = "Sample/sample-api-gw-workbook-first-mappings.xml";
      var expectedDoc = XDocument.Load(sampleWorkbookFirstSheetMappings);
      var actualDoc = XDocument.Parse(customXmlContent);

      // iterate the nodes of both expected and actual until a mismatch is found

      var lastMismatch = expectedDoc.DescendantNodes().Zip(actualDoc.DescendantNodes(), (e, a) => (Expected: e, Actual: a))
         .Reverse()
         .Where(path => path.Expected is XElement el && el.Elements().Any())
         .FirstOrDefault(pair => !XNode.DeepEquals(pair.Expected, pair.Actual));

      if (lastMismatch != default)
      {
         Console.WriteLine($"Expected: {lastMismatch.Expected}");
         Console.WriteLine($"Actual: {lastMismatch.Actual}");
         Console.WriteLine($"Actual XML:\n{actualDoc}");
         Assert.Fail();
      }
   }

   [Fact]
   public async Task Generated_excel_file_records_mappings_in_custom_xml_for_openapi()
   {
      const string openApiFile = "Sample/sample-api-gw.json";
      const string outputFile = "output-with-mappings.xlsx";
      await using var file = File.OpenRead(openApiFile);

      await OpenApiDocumentationGenerator.GenerateDocumentation(file, outputFile, new OpenApiDocumentationOptions() { IncludeMappings = true });
      Assert.True(File.Exists(outputFile));

      var mappings = ExcelOpenXmlHelper.ExtractCustomXmlMappingsFromWorkbook(outputFile);
      var worksheetNames = ExcelOpenXmlHelper.GetAllWorksheetNames(outputFile);

      Assert.Equal(worksheetNames.Count, mappings.Count);
      foreach (var mapping in mappings)
      {
         Assert.Contains(mapping.WorksheetName, worksheetNames);
         Assert.True(mapping.Mappings.Any(), $"Worksheet '{mapping.WorksheetName}' should contain OpenAPI mappings but none were found");
      }
   }

   [Fact]
   public async Task MigrateComments_ToNewWorkbook_PreservesUnresolvedCommentsInCorrectCells()
   {
      // Arrange
      const string openApiFile = "Sample/sample-api-gw.json";
      const string existingWorkbook = "Sample/sample-api-gw-with-mappings.xlsx";
      const string outputFile = "output-with-migrated-comments.xlsx";
      
      // Extract original comments for comparison
      var originalUnresolvedComments = ExcelOpenXmlHelper.ExtractAndAnnotateUnresolvedComments(existingWorkbook);
      var originalResolvedComments = ExcelOpenXmlHelper.ExtractAndAnnotateResolvedComments(existingWorkbook);
      var originalAllComments = ExcelOpenXmlHelper.ExtractAndAnnotateAllComments(existingWorkbook);
      
      // Act
      await using var file = File.OpenRead(openApiFile);
      await OpenApiDocumentationGenerator.GenerateDocumentation(
         file, 
         outputFile, 
         new OpenApiDocumentationOptions() 
         { 
            IncludeMappings = true,
            FilepathToPreserveComments = existingWorkbook 
         }
      );
      
      // Assert
      Assert.True(File.Exists(outputFile));
      
      var migratedComments = ExcelOpenXmlHelper.ExtractAndAnnotateAllComments(outputFile);
      
      // 1. Should have fewer comments than original (resolved ones abandoned)
      Assert.True(migratedComments.Count <= originalAllComments.Count, 
         $"Expected migrated comments ({migratedComments.Count}) to be <= original total comments ({originalAllComments.Count})");
      Assert.True(migratedComments.Count > 0, "Should have some migrated comments");
      
      // 2. Should have migrated approximately the same number as original unresolved comments
      Assert.Equal(originalUnresolvedComments.Count, migratedComments.Count);
      
      // 3. Verify resolved comments were NOT migrated
      Assert.True(originalResolvedComments.Count > 0, "Test requires some resolved comments in the source workbook");
      
      // 4. Each migrated comment should match an original unresolved comment
      foreach (var migratedComment in migratedComments)
      {
         var matchingOriginal = originalUnresolvedComments.FirstOrDefault(orig => 
            orig.CommentText == migratedComment.CommentText &&
            orig.OpenApiAnchor == migratedComment.OpenApiAnchor);
            
         Assert.NotNull(matchingOriginal); // Should find a match
         
         // 5. Metadata should be preserved
         Assert.Equal(matchingOriginal.CreatedDate, migratedComment.CreatedDate);
         Assert.Equal(matchingOriginal.Comment.PersonId?.Value, migratedComment.Comment.PersonId?.Value);
         
         // 6. Should be in correct cell based on mapping
         Assert.NotEmpty(migratedComment.CellReference);
         Assert.NotEmpty(migratedComment.OpenApiAnchor);
      }
      
      // 7. Verify only unresolved comments were migrated
      // All migrated comments should have been unresolved in the original workbook
      Assert.All(migratedComments, migratedComment =>
      {
         var matchingOriginal = originalUnresolvedComments.First(orig => 
            orig.CommentText == migratedComment.CommentText &&
            orig.OpenApiAnchor == migratedComment.OpenApiAnchor);
         Assert.NotEqual("1", matchingOriginal.Comment.Done); // Should be unresolved
      });
   }

}


