using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing;
using Microsoft.OpenApi.Models;
using openapi2excel.core.Common;

namespace openapi2excel.core.Builders.WorksheetPartsBuilders.Common;

internal class OpenApiSchemaDescriptor(IXLWorksheet worksheet, OpenApiDocumentationOptions options)
{
   public int AddNameHeader(RowPointer actualRow, int startColumn)
      => worksheet.Cell(actualRow, startColumn).SetTextBold("Name").GetColumnNumber();

   public int AddNameValue(string name, int actualRow, int startColumn, CustomXML.Anchor mappingAnchor)
      => worksheet.Cell(actualRow, startColumn).SetText(name)
         .MapTableCell(mappingAnchor, WorksheetLanguage.Generic.Name)
         .MapRow(mappingAnchor)
         .GetColumnNumber();

   public int AddSchemaDescriptionHeader(RowPointer actualRow, int startColumn, CustomXML.Anchor mappingAnchor)
   {
      var cell = worksheet.Cell(actualRow, startColumn).SetTextBold("Type")
         .CellRight().SetTextBold("Object type")
         .CellRight().SetTextBold("Format")
         .CellRight().SetTextBold("Length")
         .CellRight().SetTextBold("Required")
         .CellRight().SetTextBold("Nullable")
         .CellRight().SetTextBold("Range")
         .CellRight().SetTextBold("Pattern")
         .CellRight().SetTextBold("Enum")
         .CellRight().SetTextBold("Deprecated")
         .CellRight().SetTextBold("Default")
         .CellRight().SetTextBold("Example")
         .CellRight().SetTextBold("Description");

      return cell.GetColumnNumber();
   }

   public int AddSchemaDescriptionValues(OpenApiSchema schema, bool required, RowPointer actualRow, int startColumn, CustomXML.Anchor mappingAnchor, string? description = null, bool includeArrayItemType = false)
   {
      if (schema.Items != null && includeArrayItemType)
      {
         var cell = worksheet.Cell(actualRow, startColumn).SetText("array").MapTableCell(mappingAnchor, WorksheetLanguage.Parameters.Type)
            .CellRight().SetText(schema.GetObjectDescription())
            .CellRight().SetText(schema.Items.Type).MapTableCell(mappingAnchor, WorksheetLanguage.Parameters.ObjectType)
            .CellRight().SetText(schema.Items.Format).MapTableCell(mappingAnchor, WorksheetLanguage.Parameters.Format)
            .CellRight().SetText(schema.GetPropertyLengthDescription()).SetHorizontalAlignment(XLAlignmentHorizontalValues.Center).MapTableCell(mappingAnchor, WorksheetLanguage.Parameters.Length)
            .CellRight().SetText(options.Language.Get(required)).SetHorizontalAlignment(XLAlignmentHorizontalValues.Center).MapTableCell(mappingAnchor, WorksheetLanguage.Parameters.Required)
            .CellRight().SetText(options.Language.Get(schema.Nullable)).SetHorizontalAlignment(XLAlignmentHorizontalValues.Center).MapTableCell(mappingAnchor, WorksheetLanguage.Parameters.Nullable)
            .CellRight().SetText(schema.GetPropertyRangeDescription()).SetHorizontalAlignment(XLAlignmentHorizontalValues.Center).MapTableCell(mappingAnchor, WorksheetLanguage.Parameters.Range)
            .CellRight().SetText(schema.Items.Pattern).MapTableCell(mappingAnchor, WorksheetLanguage.Parameters.Pattern)
            .CellRight().SetText(schema.Items.GetEnumDescription()).MapTableCell(mappingAnchor, WorksheetLanguage.Parameters.Enum)
            .CellRight().SetText(options.Language.Get(schema.Deprecated)).SetHorizontalAlignment(XLAlignmentHorizontalValues.Center).MapTableCell(mappingAnchor, WorksheetLanguage.Parameters.Deprecated)
            .CellRight().SetText(schema.GetExampleDescription()).MapTableCell(mappingAnchor, WorksheetLanguage.Parameters.Example)
            .CellRight().SetText((string.IsNullOrEmpty(schema.Description) ? description : schema.Description).StripHtmlTags()).MapTableCell(mappingAnchor, WorksheetLanguage.Parameters.Description);

         return cell.GetColumnNumber();
      }
      else
      {
         var cell = worksheet.Cell(actualRow, startColumn).SetText(schema.GetTypeDescription()).MapTableCell(mappingAnchor, WorksheetLanguage.Parameters.Type)
            .CellRight().SetText(schema.GetObjectDescription()).MapTableCell(mappingAnchor, WorksheetLanguage.Parameters.ObjectType)
            .CellRight().SetText(schema.Format).MapTableCell(mappingAnchor, WorksheetLanguage.Parameters.Format)
            .CellRight().SetText(schema.GetPropertyLengthDescription()).SetHorizontalAlignment(XLAlignmentHorizontalValues.Center).MapTableCell(mappingAnchor, WorksheetLanguage.Parameters.Length)
            .CellRight().SetText(options.Language.Get(required)).SetHorizontalAlignment(XLAlignmentHorizontalValues.Center).MapTableCell(mappingAnchor, WorksheetLanguage.Parameters.Required)
            .CellRight().SetText(options.Language.Get(schema.Nullable)).SetHorizontalAlignment(XLAlignmentHorizontalValues.Center).MapTableCell(mappingAnchor, WorksheetLanguage.Parameters.Nullable)
            .CellRight().SetText(schema.GetPropertyRangeDescription()).SetHorizontalAlignment(XLAlignmentHorizontalValues.Center).MapTableCell(mappingAnchor, WorksheetLanguage.Parameters.Range)
            .CellRight().SetText(schema.Pattern).MapTableCell(mappingAnchor, WorksheetLanguage.Parameters.Pattern)
            .CellRight().SetText(schema.GetEnumDescription()).MapTableCell(mappingAnchor, WorksheetLanguage.Parameters.Enum)
            .CellRight().SetText(options.Language.Get(schema.Deprecated)).SetHorizontalAlignment(XLAlignmentHorizontalValues.Center).MapTableCell(mappingAnchor, WorksheetLanguage.Parameters.Deprecated)
            .CellRight().SetText(schema.GetDefaultDescription()).MapTableCell(mappingAnchor, WorksheetLanguage.Parameters.Default)
            .CellRight().SetText(schema.GetExampleDescription()).MapTableCell(mappingAnchor, WorksheetLanguage.Parameters.Example)
            .CellRight().SetText((string.IsNullOrEmpty(schema.Description) ? description : schema.Description).StripHtmlTags()).MapTableCell(mappingAnchor, WorksheetLanguage.Parameters.Description);

         return cell.GetColumnNumber();
      }
   }
}