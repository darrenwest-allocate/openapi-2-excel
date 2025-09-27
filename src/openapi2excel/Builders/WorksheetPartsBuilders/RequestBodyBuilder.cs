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

      var mappingAnchor = AnchorGenerator.GenerateRequestBodyAnchor(anchor);

      Cell(1).SetTextBold(WorksheetLanguage.Request.Title)
         .MapRow(mappingAnchor.With(WorksheetLanguage.Generic.TitleRow));

      ActualRow.MoveNext();

      using (var _ = new Section(Worksheet, ActualRow))
      {
         var builder = new PropertiesTreeBuilder(attributesColumnIndex, Worksheet, Options);
         builder.AddPropertiesTreeForMediaTypes(ActualRow, operation.RequestBody.Content, Options, mappingAnchor);
         ActualRow.MovePrev();
      }

      ActualRow.MoveNext(2);
   }
}