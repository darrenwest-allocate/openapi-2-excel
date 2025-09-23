using ClosedXML.Excel;
using Microsoft.OpenApi.Models;
using openapi2excel.core.Builders.CustomXml;
using openapi2excel.core.Builders.WorksheetPartsBuilders;
using openapi2excel.core.Builders.WorksheetPartsBuilders.Common;
using openapi2excel.core.Common;
using System.Text.RegularExpressions;

namespace openapi2excel.core.Builders;

public class OperationWorksheetBuilder : WorksheetBuilder
{
   private readonly RowPointer _actualRowPointer = new(1);
   private readonly IXLWorkbook _workbook;
   private IXLWorksheet _worksheet = null!;
   private int _attributesColumnsStartIndex;
   private WorksheetOpenApiMapping? _worksheetMapping;

   public OperationWorksheetBuilder(IXLWorkbook workbook, OpenApiDocumentationOptions options)
      : base(options)
   {
      _workbook = workbook;
   }

   public static List<IXLWorksheet> OperationWorksheets { get; } = [];

   public string CurrentWorksheetName => _worksheet.Name;

   public WorksheetOpenApiMapping CurrentWorksheetMapping => _worksheetMapping!;

   public IXLWorksheet Build(string path, OpenApiPathItem pathItem, OperationType operationType,
      OpenApiOperation operation)
   {
      var worksheetName = GetWorksheetName(path, operation, operationType);
      OperationWorksheets.Add(CreateNewWorksheet(worksheetName));

      _actualRowPointer.GoTo(1);

      _attributesColumnsStartIndex = MaxPropertiesTreeLevel.Calculate(operation, Options.MaxDepth);
      AdjustColumnsWidthToRequestTreeLevel();
      
      WorksheetOpenApiMapping.AllWorksheetMappings.Add(CreateWorksheetMapping(worksheetName));
      var anchor = AnchorGenerator.GenerateOperationAnchor(path, operationType);

      AddHomePageLink();
      AddOperationInfos(path, pathItem, operationType, operation);
      AddRequestParameters(operation, anchor);
      AddRequestBody(operation, anchor);
      AddResponseBody(operation, anchor);
      AdjustLastNamesColumnToContents();
      AdjustDescriptionColumnToContents();

      return _worksheet;
   }

   private WorksheetOpenApiMapping CreateWorksheetMapping(string worksheetName)
   {
      _worksheetMapping = new WorksheetOpenApiMapping(worksheetName);
      return _worksheetMapping;
   }


   private string GetWorksheetName(string path, OpenApiOperation operation, OperationType operationType)
   {
      var maxLength = 28;
      var name = "";
      if (!string.IsNullOrEmpty(operation.OperationId))
      {
         // take worksheet name from OperationId
         name = operation.OperationId;
      }
      else
      {
         // generate worksheet name based on operationType and path
         var pathName = path.Replace("/", "-");
         name = operationType.ToString().ToUpper() + "_" + pathName[1..];
      }

      name = Regex.Replace(name, "&", "and");
      name = Regex.Replace(name, "\\+", "plus");
      name = Regex.Replace(name, "[{}'\"<>]", string.Empty);

      // check if the name is not too long
      if (name.Length > maxLength)
      {
         name = name[..maxLength];
      }

      // check if the name is unique
      var nr = 2;
      var tmpName = name;
      while (_workbook.Worksheets.Any(s => s.Name.Equals(tmpName, StringComparison.CurrentCultureIgnoreCase)))
      {
         tmpName = name[..maxLength] + "_" + nr++;
      }
      return tmpName;
   }

   private IXLWorksheet CreateNewWorksheet(string operation)
   {
      _worksheet = _workbook.Worksheets.Add(operation);
      _worksheet.Style.Font.FontSize = 10;
      _worksheet.Style.Font.FontName = "Arial";
      _worksheet.Outline.SummaryHLocation = XLOutlineSummaryHLocation.Left;
      _worksheet.Outline.SummaryVLocation = XLOutlineSummaryVLocation.Top;
      return _worksheet;
   }

   private void AdjustColumnsWidthToRequestTreeLevel()
   {
      for (var columnIndex = 1; columnIndex < _attributesColumnsStartIndex - 1; columnIndex++)
      {
         _worksheet.Column(columnIndex).Width = 1.8;
      }
   }

   private void AdjustLastNamesColumnToContents()
   {
      if (_attributesColumnsStartIndex > 1)
      {
         _worksheet.Column(_attributesColumnsStartIndex - 1).AdjustToContents();
      }
   }

   private void AdjustDescriptionColumnToContents()
   {
      _worksheet.LastColumnUsed()?.AdjustToContents();
   }

   private void AddOperationInfos(string path, OpenApiPathItem pathItem, OperationType operationType,
      OpenApiOperation operation) =>
   new OperationInfoBuilder(_actualRowPointer, _attributesColumnsStartIndex, _worksheet, Options)
      .AddOperationInfoSection(path, pathItem, operationType, operation, _worksheetMapping!.Mappings);

   private void AddRequestParameters(OpenApiOperation operation, Anchor anchor) =>
      new RequestParametersBuilder(_actualRowPointer, _attributesColumnsStartIndex, _worksheet, Options)
         .AddRequestParametersPart(operation, anchor);

   private void AddRequestBody(OpenApiOperation operation, Anchor anchor) =>
      new RequestBodyBuilder(_actualRowPointer, _attributesColumnsStartIndex, _worksheet, Options)
         .AddRequestBodyPart(operation, anchor);

   private void AddResponseBody(OpenApiOperation operation, Anchor anchor) =>
      new ResponseBodyBuilder(_actualRowPointer, _attributesColumnsStartIndex, _worksheet, Options)
         .AddResponseBodyPart(operation, anchor);

   private void AddHomePageLink() => new HomePageLinkBuilder(_actualRowPointer, _worksheet, Options)
      .AddHomePageLinkSection();
}
