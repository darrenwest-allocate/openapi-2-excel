using openapi2excel.core;
using openapi2excel.core.Common;
using Spectre.Console;
using Spectre.Console.Cli;
using System.ComponentModel;
using Path = System.IO.Path;

namespace OpenApi2Excel.cli;

[Description("Generate Rest API specification in a MS Excel format")]
public class GenerateExcelCommand : Command<GenerateExcelCommand.GenerateExcelSettings>
{
   public class GenerateExcelSettings : CommandSettings
   {
      [Description("The path or URL to a YAML or JSON file with Rest API specification.")]
      [CommandArgument(0, "<INPUT_FILE>")]
      public string InputFile { get; init; } = null!;

      [Description("The path for output excel file.")]
      [CommandArgument(1, "<OUTPUT_FILE>")]
      public string OutputFile { get; init; } = null!;

      [Description("Run tool without logo.")]
      [CommandOption("-n|--no-logo")]
      public bool NoLogo { get; init; }

      [Description("Maximum depth level for documenting object hiearchies (defaults to 10).")]
      [CommandOption("-d|--depth")]
      public int Depth { get; init; } = 10;

      [Description("Run tool with debug mode.")]
      [CommandOption("-g|--debug")]
      public bool Debug { get; init; }

      [Description("Path to existing Excel workbook from which to preserve comments.")]
      [CommandOption("-e|--existing-workbook")]
      public string? ExistingWorkbook { get; init; }

      internal FileInfo InputFileParsed { get; set; } = null!;
      internal FileInfo OutputFileParsed { get; set; } = null!;
      internal bool IsOutputDirectory { get; set; }

      /// <summary>
      /// Determines if the given path represents a directory rather than a file.
      /// </summary>
      private static bool IsDirectoryPath(string path)
      {
         // Check if path ends with directory separator
         if (path.EndsWith(Path.DirectorySeparatorChar.ToString()) || 
             path.EndsWith(Path.AltDirectorySeparatorChar.ToString()))
         {
            return true;
         }

         // Check if path exists and is a directory
         if (Directory.Exists(path))
         {
            return true;
         }

         // Check if path doesn't have an extension and doesn't contain a file-like name
         var fileName = Path.GetFileName(path);
         if (string.IsNullOrEmpty(fileName) || !Path.HasExtension(path))
         {
            // If the parent directory exists, this is likely a directory path
            var parentDir = Path.GetDirectoryName(path);
            if (!string.IsNullOrEmpty(parentDir) && Directory.Exists(parentDir))
            {
               return true;
            }
         }

         return false;
      }

      public override ValidationResult Validate()
      {
         var inputFilePath = InputFile.Trim();
         if (File.Exists(inputFilePath))
         {
            InputFileParsed = new FileInfo(inputFilePath);
         }
         else if (Uri.TryCreate(inputFilePath, UriKind.RelativeOrAbsolute, out var uri))
         {
            var inputFileTempPath = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".xlsx");
            if (TryDownloadFileTaskAsync(uri, inputFileTempPath).GetAwaiter().GetResult())
            {
               InputFileParsed = new FileInfo(inputFileTempPath);
            }
            else
            {
               return ValidationResult.Error("Invalid input file path.");
            }
         }
         else
         {
            return ValidationResult.Error("Invalid input file path.");
         }

         var outputFilePath = OutputFile.Trim();

         // Check if the output path is a directory
         if (IsDirectoryPath(outputFilePath))
         {
            IsOutputDirectory = true;
            
            // Validate that the directory exists or can be created
            try
            {
               if (!Directory.Exists(outputFilePath))
               {
                  // Try to create the directory to validate the path
                  var testPath = Path.GetFullPath(outputFilePath);
                  Directory.CreateDirectory(testPath);
               }
            }
            catch (Exception ex) when (ex is UnauthorizedAccessException || ex is DirectoryNotFoundException || 
                                      ex is ArgumentException || ex is NotSupportedException || ex is IOException)
            {
               return ValidationResult.Error($"Invalid output directory path: {outputFilePath}. {ex.Message}");
            }
            
            // Generate filename from OpenAPI document
            try
            {
               var (title, version) = OpenApiInfoExtractor.ExtractInfoAsync(InputFileParsed.FullName).GetAwaiter().GetResult();
               var generatedFilename = OpenApiInfoExtractor.GenerateFilename(title, version);
               OutputFileParsed = new FileInfo(Path.Combine(outputFilePath, generatedFilename));
            }
            catch (Exception)
            {
               // Fallback to default filename if OpenAPI parsing fails
               OutputFileParsed = new FileInfo(Path.Combine(outputFilePath, "api_documentation.xlsx"));
            }
         }
         else
         {
            IsOutputDirectory = false;
            // Existing file path logic
            if (!outputFilePath.EndsWith(".xlsx", StringComparison.CurrentCultureIgnoreCase))
            {
               outputFilePath += ".xlsx";
            }
            OutputFileParsed = new FileInfo(outputFilePath);
         }

         return ValidationResult.Success();
      }

      private static async Task<bool> TryDownloadFileTaskAsync(Uri uri, string fileName)
      {
         try
         {
            var client = new HttpClient();
            await using var s = await client.GetStreamAsync(uri);
            await using var fs = new FileStream(fileName, FileMode.CreateNew);
            await s.CopyToAsync(fs);
            return true;
         }
         catch
         {
            return false;
         }
      }
   }

   public override int Execute(CommandContext context, GenerateExcelSettings settings)
   {
      if (!settings.NoLogo)
      {
         foreach (var renderable in CustomHelpProvider.GetHeaderText())
         {
            AnsiConsole.Write(renderable);
         }
      }

      try
      {
         var options = new OpenApiDocumentationOptions 
         { 
            MaxDepth = settings.Depth,
            FilepathToPreserveComments = settings.ExistingWorkbook ?? string.Empty
         };

         OpenApiDocumentationGenerator
            .GenerateDocumentation(settings.InputFileParsed.FullName, settings.OutputFileParsed.FullName, options)
            .ConfigureAwait(false).GetAwaiter().GetResult();

         AnsiConsole.MarkupLine($"Excel file saved to [green]{settings.OutputFileParsed.FullName.EscapeMarkup()}[/]");
      }
      catch (IOException exc)
      {
         AnsiConsole.MarkupLine(settings.Debug ? $"[red]{exc.ToString().EscapeMarkup()}[/]" : $"[red]{exc.Message.EscapeMarkup()}[/]");

         return 1;
      }
      catch (Exception exc)
      {
         AnsiConsole.MarkupLine(settings.Debug
            ? $"[red]An unexpected error occurred: {exc.ToString().EscapeMarkup()}[/]"
            : $"[red]An unexpected error occurred: {exc.Message.EscapeMarkup()}[/]");

         return 1;
      }

      return 0;
   }
}