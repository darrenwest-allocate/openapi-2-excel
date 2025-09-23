using ClosedXML.Excel;
using Microsoft.OpenApi.Models;
using openapi2excel.core.Builders.CustomXml;
using openapi2excel.core.Builders.WorksheetPartsBuilders.Common;
using openapi2excel.core.Common;

namespace openapi2excel.core.Builders.WorksheetPartsBuilders;

internal class ResponseBodyBuilder(
   RowPointer actualRow,
   int attributesColumnIndex,
   IXLWorksheet worksheet,
   OpenApiDocumentationOptions options) : WorksheetPartBuilder(actualRow, worksheet, options)
{
   public void AddResponseBodyPart(OpenApiOperation operation, Anchor anchor)
   {
      if (!operation.Responses.Any())
         return;
      var responseAnchor = AnchorGenerator.GenerateResponseBodyAnchor(anchor);
      Cell(1).SetTextBold("RESPONSE").MapRow(responseAnchor);
      ActualRow.MoveNext();
      using (var _ = new Section(Worksheet, ActualRow))
      {
         var builder = new PropertiesTreeBuilder(attributesColumnIndex, Worksheet, Options);
         foreach (var response in operation.Responses)
         {
            AddResponseHttpCode(response.Key, response.Value.Description, responseAnchor);
            AddResponseHeaders(response.Value.Headers, responseAnchor);
            var mappingAnchor = AnchorGenerator.GenerateResponseAnchor(anchor, response.Key);
            builder.AddPropertiesTreeForMediaTypes(ActualRow, response.Value.Content, Options, mappingAnchor);
         }
      }
      ActualRow.MoveNext();
   }

   private void AddResponseHeaders(IDictionary<string, OpenApiHeader> valueHeaders, Anchor anchor)
   {
      if (!valueHeaders.Any())
         return;

      ActualRow.MoveNext();
      var headersAnchor = AnchorGenerator.GenerateHeadersAnchor(anchor);

      var responseHeaderRowPointer = ActualRow.Copy();
      Cell(1).SetTextBold("Response headers")
         .MapRowWithDetail(headersAnchor);
      ActualRow.MoveNext();

      using (var _ = new Section(Worksheet, ActualRow))
      {
         var schemaDescriptor = new OpenApiSchemaDescriptor(Worksheet, Options);

         InsertHeader(schemaDescriptor, headersAnchor);
         ActualRow.MoveNext();

         foreach (var openApiHeader in valueHeaders)
         {
            InsertProperty(openApiHeader, schemaDescriptor, headersAnchor);
            ActualRow.MoveNext();
         }
      }
      ActualRow.MoveNext();

      void InsertHeader(OpenApiSchemaDescriptor schemaDescriptor, Anchor anchor)
      {
         var nextCell = Cell(1).SetTextBold("Name")
            .MapRowWithDetail(anchor)
            .CellRight(attributesColumnIndex + 1).GetColumnNumber();

         var lastUsedColumn = schemaDescriptor.AddSchemaDescriptionHeader(ActualRow, nextCell, anchor);

         Worksheet.Cell(ActualRow, 1)
            .SetBackground(lastUsedColumn, HeaderBackgroundColor)
            .SetBottomBorder(lastUsedColumn);

         Worksheet.Cell(responseHeaderRowPointer, 1)
            .SetBackground(lastUsedColumn, HeaderBackgroundColor);
      }

      void InsertProperty(KeyValuePair<string, OpenApiHeader> openApiHeader, OpenApiSchemaDescriptor schemaDescriptor, Anchor mappingAnchor)
      {
         var nextCellNumber = Cell(1).SetText(openApiHeader.Key)
            .MapRowWithDetail(mappingAnchor)
            .CellRight(attributesColumnIndex + 1).GetColumnNumber();

         nextCellNumber = schemaDescriptor.AddSchemaDescriptionValues(openApiHeader.Value.Schema, openApiHeader.Value.Required, ActualRow, nextCellNumber, mappingAnchor);

         Cell(nextCellNumber).SetText(openApiHeader.Value.Description);
      }
   }

   private void AddResponseHttpCode(string httpCode, string? description, Anchor anchor)
   {
      var responseCode = httpCode.Equals("default") ? "Default response" : $"Response HttpCode: {httpCode}";
      if (!string.IsNullOrEmpty(description) && !description.Equals("default response"))
      {
         responseCode += $": {description}";
      }

      Cell(1).SetTextBold(responseCode).MapRow(anchor.With(httpCode));

      ActualRow.MoveNext();
   }
}