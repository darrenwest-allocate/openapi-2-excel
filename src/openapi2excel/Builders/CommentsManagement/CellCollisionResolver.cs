using ClosedXML.Excel;
using System.Collections.Generic;
using System.Linq;

namespace openapi2excel.core.Builders.CommentsManagement;

/// <summary>
/// Handles collision detection and resolution when placing comments in cells.
/// </summary>
public class CellCollisionResolver
{
    /// <summary>
    /// Handles collision detection for Type A comments and finds an alternative cell if needed.
    /// </summary>
    public string HandleTypeACollision(IXLWorksheet worksheet, string targetCellReference, HashSet<string> processedCells)
    {
        var targetCell = worksheet.Cell(targetCellReference);
        var cellKey = $"{worksheet.Name}:{targetCellReference}";
        
        // If the target cell is empty and not already processed, use it
        if ((targetCell.IsEmpty() || !targetCell.HasComment) && !processedCells.Contains(cellKey))
        {
            return targetCellReference;
        }
        
        // Find the next available row below
        var originalColumn = ExtractColumnFromCellReference(targetCellReference);
        var startRow = ExtractRowFromCellReference(targetCellReference);
        
        for (int row = startRow + 1; row <= startRow + 5; row++) // Check up to 5 rows below
        {
            var candidateCell = $"{originalColumn}{row}";
            var candidateKey = $"{worksheet.Name}:{candidateCell}";
            var cell = worksheet.Cell(candidateCell);
            
            if ((cell.IsEmpty() || !cell.HasComment) && !processedCells.Contains(candidateKey))
            {
                return candidateCell;
            }
        }
        
        // If still no space, use the original target (will merge with existing comment)
        return targetCellReference;
    }

    /// <summary>
    /// Finds the next available row in a specific column for Type B comment placement.
    /// </summary>
    public int FindNextAvailableRowInColumn(IXLWorksheet worksheet, int column, HashSet<string> processedCells)
    {
        // Start from row 1 and find the first available row
        for (int row = 1; row <= 1000; row++) // Reasonable limit
        {
            var cellReference = worksheet.Cell(row, column).Address.ToString();
            var cellKey = $"{worksheet.Name}:{cellReference}";
            
            var cell = worksheet.Cell(row, column);
            
            // Check if this cell is available (empty, no comment, not processed)
            if ((cell.IsEmpty() || !cell.HasComment) && !processedCells.Contains(cellKey))
            {
                return row;
            }
        }
        
        // Fallback to row 1 if no space found
        return 1;
    }

    private static string ExtractColumnFromCellReference(string cellReference)
    {
        return new string([.. cellReference.TakeWhile(c => !char.IsDigit(c))]);
    }

    private static int ExtractRowFromCellReference(string cellReference)
    {
        var digitStart = cellReference.IndexOf(cellReference.First(char.IsDigit));
        var rowString = cellReference.Substring(digitStart);
        return int.Parse(rowString);
    }
}