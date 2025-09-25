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

      if (options.IncludeMappings || !string.IsNullOrEmpty(options.FilepathToPreserveComments))
      {
         // Save to a temporary file to add custom XML parts
         var tempFile = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
         workbook.SaveAs(tempFile);

         if (options.IncludeMappings)
         {
            WorksheetOpenApiMapping.AllWorksheetMappings
               .ForEach(worksheetMapping => ExcelCustomXmlHelper.WriteCustomXmlMapping(tempFile, worksheetMapping));
         }

         // Migrate comments from existing workbook if specified
         if (!string.IsNullOrEmpty(options.FilepathToPreserveComments))
         {
            var newWorkbookMappings = ExcelOpenXmlHelper.ExtractCustomXmlMappingsFromWorkbook(tempFile);
            var nonMigratableComments = CommentMigrationHelper.MigrateComments(
               options.FilepathToPreserveComments,
               tempFile,
               newWorkbookMappings);
         }

         // Move the modified temp file to the final output path
         File.Copy(tempFile, filePath, true);
         File.Delete(tempFile);
      }
      else
      {
         workbook.SaveAs(filePath);
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