using ClosedXML.Excel;
using Microsoft.OpenApi.Models;
using openapi2excel.core.Builders.CustomXml;
using openapi2excel.core.Builders.WorksheetPartsBuilders.Common;
using openapi2excel.core.Common;

namespace openapi2excel.core.Builders.WorksheetPartsBuilders;

internal class OperationInfoBuilder(
   RowPointer actualRow,
   int attributesColumnIndex,
   IXLWorksheet worksheet,
   OpenApiDocumentationOptions options)
   : WorksheetPartBuilder(actualRow, worksheet, options)
{
   public void AddOperationInfoSection(string path, OpenApiPathItem pathItem, OperationType operationType,
      OpenApiOperation operation, List<CellOpenApiMapping> mappings)
   {
   Cell(1).SetTextBold(WorksheetLanguage.Operations.Heading);
      ActualRow.MoveNext();

      Anchor mappingAnchor = AnchorGenerator.GenerateOperationInfoAnchor(path, pathItem, operationType, operation); 

      using (var _ = new Section(Worksheet, ActualRow))
      {
         var cell = Cell(1).SetTextBold(WorksheetLanguage.Operations.OperationType).CellRight(attributesColumnIndex).SetText(operationType.ToString().ToUpper()).MapRow(mappingAnchor)
            .IfNotEmpty(operation.OperationId, c => c.NextRow().SetTextBold(WorksheetLanguage.Operations.Id).MapRowWithDetail(mappingAnchor).CellRight(attributesColumnIndex).SetText(operation.OperationId))
            .NextRow().SetTextBold(WorksheetLanguage.Operations.Path).MapRowWithDetail(mappingAnchor).CellRight(attributesColumnIndex).SetText(path)
            .IfNotEmpty(pathItem.Description, c => c.NextRow().SetTextBold(WorksheetLanguage.Operations.PathDescription).MapRowWithDetail(mappingAnchor).CellRight(attributesColumnIndex).SetText(pathItem.Description))
            .IfNotEmpty(pathItem.Summary, c => c.NextRow().SetTextBold(WorksheetLanguage.Operations.PathSummary).MapRowWithDetail(mappingAnchor).CellRight(attributesColumnIndex).SetText(pathItem.Summary))
            .IfNotEmpty(operation.Description, c => c.NextRow().SetTextBold(WorksheetLanguage.Operations.OperationDescription).MapRowWithDetail(mappingAnchor).CellRight(attributesColumnIndex).SetText(operation.Description))
            .IfNotEmpty(operation.Summary, c => c.NextRow().SetTextBold(WorksheetLanguage.Operations.OperationSummary).MapRowWithDetail(mappingAnchor).CellRight(attributesColumnIndex).SetText(operation.Summary))
            .NextRow().SetTextBold(WorksheetLanguage.Operations.Deprecated).MapRowWithDetail(mappingAnchor).CellRight(attributesColumnIndex).SetText(Options.Language.Get(operation.Deprecated));

         ActualRow.GoTo(cell.Address.RowNumber);
      }

      ActualRow.MoveNext(2);
   }
}