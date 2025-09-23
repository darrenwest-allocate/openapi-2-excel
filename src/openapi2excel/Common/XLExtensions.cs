using ClosedXML.Excel;
using openapi2excel.core.Builders;
using openapi2excel.core.Builders.CustomXml;

namespace openapi2excel.core.Common;

// ReSharper disable once InconsistentNaming
internal static class XLExtensions
{
   public static IXLCell SetBoldStyle(this IXLCell cell)
   {
      cell.Style.Font.SetBold(true);
      return cell;
   }

   public static IXLCell SetText(this IXLCell cell, string? value)
      => cell.SetValue(value?.Trim());

   public static IXLCell SetTextBold(this IXLCell cell, string? value)
      => cell.SetBoldStyle().SetValue(value?.Trim());

   public static IXLCell NextRow(this IXLCell cell)
      => cell.Worksheet.Cell(cell.Address.RowNumber + 1, 1);

   public static int GetColumnNumber(this IXLCell cell)
      => cell.Address.ColumnNumber;

   public static IXLCell If(this IXLCell cell, bool condition, Func<IXLCell, IXLCell> func)
      => condition ? func(cell) : cell;

   public static IXLCell IfNotEmpty(this IXLCell cell, string text, Func<IXLCell, IXLCell> func)
      => string.IsNullOrEmpty(text) ? cell : func(cell);

   public static IXLCell SetHorizontalAlignment(this IXLCell cell, XLAlignmentHorizontalValues alignment)
   {
      cell.Style.Alignment.SetHorizontal(alignment);
      return cell;
   }

   public static IXLCell SetBackground(this IXLCell cell, XLColor color)
   {
      cell.Style.Fill.SetBackgroundColor(color);
      return cell;
   }

   public static IXLCell SetBackground(this IXLCell cell, int toColumn, XLColor color)
   {
      var tmpCell = cell;
      while (tmpCell.Address.ColumnNumber <= toColumn)
         tmpCell = tmpCell.SetBackground(color).CellRight();

      return cell;
   }

   public static IXLCell SetBottomBorder(this IXLCell cell)
   {
      cell.Style.Border.SetBottomBorder(XLBorderStyleValues.Medium);
      return cell;
   }

   public static IXLCell SetBottomBorder(this IXLCell cell, int toColumn)
   {
      var tmpCell = cell;
      while (tmpCell.Address.ColumnNumber <= toColumn)
         tmpCell = tmpCell.SetBottomBorder().CellRight();

      return cell;
   }

   public static IXLCell MapRow(this IXLCell cell, Anchor mappingAnchor)
   {
      var mappings = OperationWorksheetBuilder.OperationWorksheets
         .FirstOrDefault(w => w.Name == cell.Worksheet.Name)?.WorksheetMapping.Mappings;

      mappings?.Add(new CellOpenApiMapping() { Row = cell.WorksheetRow().RowNumber(), Cell = cell.Address.ToString() ?? string.Empty, OpenApiRef = mappingAnchor.ToString() });
      return cell;
   }

   public static IXLCell MapRowWithDetail(this IXLCell cell, Anchor mappingAnchor)
   {
      return MapRow(cell, AnchorGenerator.AppendDetailToAnchor(mappingAnchor, cell.GetText() ?? string.Empty));
   }

   public static IXLCell MapTableCell(this IXLCell cell, Anchor mappingAnchor, string columnName)
   {
      var mappings = OperationWorksheetBuilder.OperationWorksheets
         .FirstOrDefault(w => w.Name == cell.Worksheet.Name)?.WorksheetMapping.Mappings;

      mappings?.Add(new CellOpenApiMapping() { Cell = cell.Address.ToString() ?? string.Empty, OpenApiRef = $"{mappingAnchor}/@{columnName.Replace(" ", string.Empty).ToLowerInvariant()}" });
      return cell;
   }
}