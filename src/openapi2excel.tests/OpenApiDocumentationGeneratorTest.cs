
using openapi2excel.core;
using Microsoft.OpenApi.Readers;
using ClosedXML.Excel;
using openapi2excel.core.Builders;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Xunit;
using System.Threading.Tasks;
using OpenApi2Excel.Core.CustomXml;

namespace OpenApi2Excel.Tests
{
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
         var mappings = WorksheetOpenApiMapping.CreateMappings(
            worksheetName,
            new[]
            {
               ("A5", "paths./pets.get.responses.200"),
               ("B12", "components.parameters.PetId")
            }
         );

         // Use the serializer class
         var customXmlContent = WorksheetOpenApiMapping.Serialize(mappings.ToList());

         // Act: Write and read back
         OpenApi2Excel.Common.ExcelCustomXmlHelper.WriteCustomXmlMapping(tempFile, worksheetName, customXmlContent);
         var readXml = OpenApi2Excel.Common.ExcelCustomXmlHelper.ReadCustomXmlMapping(tempFile, worksheetName);
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
         var worksheet = worksheetBuilder.Build(path.Key, path.Value, operation.Key, operation.Value);
         var mappings = worksheetBuilder.WorksheetMapping;
         var customXmlContent = WorksheetOpenApiMapping.Serialize(new List<WorksheetOpenApiMapping> { mappings });

         // Verify that the worksheet was created and mappings were added
         Assert.NotNull(worksheet);
         Assert.Equal(worksheet.Name, mappings.Worksheet);
         Assert.NotEmpty(mappings.Mappings);

         const string sampleWorkbookFirstSheetMappings = "Sample/sample-api-gw-workbook-first-mappings.xml";
         var expectedDoc = XDocument.Load(sampleWorkbookFirstSheetMappings);
         var actualDoc = XDocument.Parse(customXmlContent);
         Assert.True(XNode.DeepEquals(expectedDoc, actualDoc), "XML content does not match expected mappings.");
      }

      [Fact]
      public void Generated_excel_file_records_mappings_in_custom_xml_for_openapi()
      { 


      }

   }
}


