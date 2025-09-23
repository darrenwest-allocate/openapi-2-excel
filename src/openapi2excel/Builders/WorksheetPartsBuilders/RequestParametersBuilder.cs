using ClosedXML.Excel;
using Microsoft.OpenApi.Models;
using openapi2excel.core.Builders.WorksheetPartsBuilders.Common;
using openapi2excel.core.Common;
using openapi2excel.core.CustomXML;

namespace openapi2excel.core.Builders.WorksheetPartsBuilders;

internal class RequestParametersBuilder(
   RowPointer actualRow,
   int attributesColumnIndex,
   IXLWorksheet worksheet,
   OpenApiDocumentationOptions options)
   : WorksheetPartBuilder(actualRow, worksheet, options)
{
   private readonly OpenApiSchemaDescriptor _schemaDescriptor = new(worksheet, options);

   public void AddRequestParametersPart(OpenApiOperation operation, Anchor anchor)
   {
      attributesColumnIndex = attributesColumnIndex > 1 ? attributesColumnIndex : 2;
      if (!operation.Parameters.Any())
         return;

      Cell(1).SetTextBold(WorksheetLanguage.Parameters.Title).MapRow(AnchorGenerator.GenerateParameterAnchor("/Title"));
      ActualRow.MoveNext();
      using (var _ = new Section(Worksheet, ActualRow))
      {
         var nextCell = Cell(1).SetTextBold(WorksheetLanguage.Parameters.Name).MapRow(AnchorGenerator.GenerateParameterAnchor("/ParameterHeadings"))
            .CellRight(attributesColumnIndex - 1).SetTextBold(WorksheetLanguage.Parameters.Location)
            .CellRight().SetTextBold(WorksheetLanguage.Parameters.Serialization)
            .CellRight();

         var lastUsedColumn = _schemaDescriptor.AddSchemaDescriptionHeader(ActualRow, nextCell.Address.ColumnNumber, anchor);

         Cell(1).SetBackground(lastUsedColumn, HeaderBackgroundColor).SetBottomBorder(lastUsedColumn);

         ActualRow.MoveNext();
         foreach (var operationParameter in operation.Parameters)
         {
            AddPropertyRow(operationParameter, AnchorGenerator.GenerateParameterAnchor(operationParameter.Name));
         }
         ActualRow.MovePrev();
      }

      ActualRow.MoveNext(2);
   }

   private void AddPropertyRow(OpenApiParameter parameter, Anchor mappingAnchor)
   {
		var nextCell = Cell(1).SetText(parameter.Name)
         .MapRow(mappingAnchor)
         .MapTableCell(mappingAnchor, WorksheetLanguage.Parameters.Name)
         .CellRight(attributesColumnIndex - 1).SetText(parameter.In.ToString()?.ToUpper())
         .MapTableCell(mappingAnchor, WorksheetLanguage.Parameters.Location)
         .CellRight().SetText(parameter.Style?.ToString())
         .MapTableCell(mappingAnchor, WorksheetLanguage.Parameters.Serialization)
         .CellRight();

      _schemaDescriptor.AddSchemaDescriptionValues(parameter.Schema, parameter.Required, ActualRow, nextCell.Address.ColumnNumber, mappingAnchor, parameter.Description, true );
      ActualRow.MoveNext();
   }
}