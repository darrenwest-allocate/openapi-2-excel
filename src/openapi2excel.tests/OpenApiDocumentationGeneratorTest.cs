using openapi2excel.core;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Xunit;
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
         var customXmlContent = WorksheetOpenApiMapping.Serialize(mappings);

         // Act: Write and read back
         OpenApi2Excel.Common.ExcelCustomXmlHelper.WriteCustomXmlMapping(tempFile, worksheetName, customXmlContent);
         var readXml = OpenApi2Excel.Common.ExcelCustomXmlHelper.ReadCustomXmlMapping(tempFile, worksheetName);
         var expectedDoc = XDocument.Parse(customXmlContent);
         var actualDoc = XDocument.Parse(readXml);
         Assert.True(XNode.DeepEquals(expectedDoc, actualDoc), "XML content does not match after round-trip.");
      }
   }
}


