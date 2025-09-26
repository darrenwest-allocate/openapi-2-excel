using ClosedXML.Excel;
using Microsoft.OpenApi.Models;
using openapi2excel.core.Builders.CommentsManagement;
using openapi2excel.core.Builders.WorksheetPartsBuilders.Common;
using openapi2excel.core.Common;

namespace openapi2excel.core.Builders.WorksheetPartsBuilders;

internal class RequestBodyBuilder(
   RowPointer actualRow,
   int attributesColumnIndex,
   IXLWorksheet worksheet,
   OpenApiDocumentationOptions options) : WorksheetPartBuilder(actualRow, worksheet, options)
{
   public void AddRequestBodyPart(OpenApiOperation operation, Anchor anchor)
   {
      if (operation.RequestBody is null)
         return;

      Cell(1).SetTextBold(WorksheetLanguage.Request.Title)
         .MapRowWithDetail(AnchorGenerator.GenerateParameterAnchor("/Title"));

      ActualRow.MoveNext();

      var mappingAnchor = AnchorGenerator.GenerateRequestBodyAnchor(anchor);

      using (var _ = new Section(Worksheet, ActualRow))
      {
         var builder = new PropertiesTreeBuilder(attributesColumnIndex, Worksheet, Options);
         builder.AddPropertiesTreeForMediaTypes(ActualRow, operation.RequestBody.Content, Options, mappingAnchor);
         ActualRow.MovePrev();
      }

      ActualRow.MoveNext(2);
   }
}