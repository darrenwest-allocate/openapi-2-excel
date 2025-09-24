using System.Text;
using ClosedXML.Excel;
using Microsoft.OpenApi.Readers;
using openapi2excel.core.Builders;
using openapi2excel.core.Builders.CustomXml;
using openapi2excel.core.Common;

namespace openapi2excel.core;

public static class OpenApiDocumentationGenerator
{
   public static async Task GenerateDocumentation(string openApiFile, string outputFile,
      OpenApiDocumentationOptions options)
   {
      if (!File.Exists(openApiFile))
         throw new FileNotFoundException($"Invalid input file path: {openApiFile}.");

      if (string.IsNullOrEmpty(outputFile))
         throw new ArgumentNullException(outputFile, "Invalid output file path.");

      await using var fileStream = File.OpenRead(openApiFile);
      await GenerateDocumentationImpl(fileStream, outputFile, options);
   }

   public static async Task GenerateDocumentation(Stream openApiFileStream, string outputFile,
      OpenApiDocumentationOptions options)
   {
      if (string.IsNullOrEmpty(outputFile))
         throw new ArgumentNullException(outputFile, "Invalid output file path.");

      await GenerateDocumentationImpl(openApiFileStream, outputFile, options);
   }

   private static async Task GenerateDocumentationImpl(Stream openApiFileStream, string outputFile,
      OpenApiDocumentationOptions options)
   {
      var readResult = await new OpenApiStreamReader().ReadAsync(openApiFileStream);
      AssertReadResult(readResult);

      WorksheetOpenApiMapping.AllWorksheetMappings.Clear();

      using var workbook = new XLWorkbook();
      var infoWorksheetsBuilder = new InfoWorksheetBuilder(workbook, options);
      infoWorksheetsBuilder.Build(readResult.OpenApiDocument);

      var worksheetBuilder = new OperationWorksheetBuilder(workbook, options);
      readResult.OpenApiDocument.Paths.ForEach(path
         => path.Value.Operations.ForEach(operation
               =>
               {
                  var worksheet = worksheetBuilder.Build(path.Key, path.Value, operation.Key, operation.Value);
                  infoWorksheetsBuilder.AddLink(operation.Key, path.Key, worksheet);

                  var mappings = worksheetBuilder.CurrentWorksheetMapping;
               }
         ));

      var filePath = new FileInfo(outputFile).FullName;
      workbook.SaveAs(filePath);

      if (options.IncludeMappings)
      {
         WorksheetOpenApiMapping.AllWorksheetMappings
            .ForEach(worksheetMapping => ExcelCustomXmlHelper.WriteCustomXmlMapping(filePath, worksheetMapping));
      }

      // Migrate comments from existing workbook if specified
      if (!string.IsNullOrEmpty(options.FilepathToPreserveComments))
      {
         CommentMigrationHelper.MigrateComments(
            options.FilepathToPreserveComments, 
            filePath, 
            WorksheetOpenApiMapping.AllWorksheetMappings);
      }
   }

   private static void AssertReadResult(ReadResult readResult)
   {
      if (!readResult.OpenApiDiagnostic.Errors.Any())
         return;

      var errorMessageBuilder = new StringBuilder();
      errorMessageBuilder.AppendLine("Some errors occurred while processing input file.");
      readResult.OpenApiDiagnostic.Errors.ToList().ForEach(e => errorMessageBuilder.AppendLine($"{e.Message} ({e.Pointer})"));
      throw new InvalidOperationException(errorMessageBuilder.ToString());
   }
}