
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
        }.Select(i => new CellOpenApiMapping(){ Cell = i.Item1, OpenApiRef = i.Item2 }).ToList();

     // Use the serializer class
     var customXmlContent = WorksheetOpenApiMapping.Serialize([new WorksheetOpenApiMapping(worksheetName) { Mappings = mappings }]);

     // Act: Write and read back
     ExcelCustomXmlHelper.WriteCustomXmlMapping(tempFile, worksheetName, customXmlContent);
     var readXml = ExcelCustomXmlHelper.ReadCustomXmlMapping(tempFile, worksheetName);
     var expectedDoc = XDocument.Parse(customXmlContent);
     var actualDoc = XDocument.Parse(readXml);
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

}


