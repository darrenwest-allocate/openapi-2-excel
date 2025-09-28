using System;
using System.Collections.Generic;
using System.Linq;

namespace openapi2excel.core.Builders.CommentsManagement;

/// <summary>
/// Resolves target cell locations for comment migration based on OpenAPI mappings and comment characteristics.
/// </summary>
public class CommentTargetResolver
{
    /// <summary>
    /// Gets the target cell reference for a comment, handling both regular mapped comments and override scenarios.
    /// </summary>
    public static bool TryGetTargetCellForThreadedComment(
        ThreadedCommentWithContext comment, 
        List<WorksheetOpenApiMapping> newWorkbookMappings, 
        out string targetCellReference)
    {
        targetCellReference = string.Empty;
        
        // Handle comments with override target cells (NoAnchor and NoWorksheet comment migrations)
        if (!string.IsNullOrEmpty(comment.OverrideTargetCell))
        {
            targetCellReference = comment.OverrideTargetCell;
            return true;
        }
        
        // Handle regular comments with OpenAPI anchors
        if (!string.IsNullOrEmpty(comment.OpenApiAnchor))
        {
            var (targetMapping, _) = FindMatchingMapping(comment.OpenApiAnchor, newWorkbookMappings);
            if (targetMapping != null)
            {
                return TryGetTargetCell(comment, targetMapping, out targetCellReference);
            }
        }
        
        return false;
    }

    /// <summary>
    /// Gets target cell reference from a specific mapping.
    /// </summary>
    public static bool TryGetTargetCell(ThreadedCommentWithContext comment, CellOpenApiMapping targetMapping, out string targetCellReference)
    {
        if (!string.IsNullOrEmpty(targetMapping.Cell))
        {
            targetCellReference = targetMapping.Cell;
        }
        else if (targetMapping.Row > 0)
        {
            // Row match - preserve original column, use mapped row
            var originalColumn = ExtractColumnFromCellReference(comment.CellReference);
            targetCellReference = $"{originalColumn}{targetMapping.Row}";
        }
        else
        {
            targetCellReference = string.Empty;
            return false;
        }
        return true;
    }

    /// <summary>
    /// Finds a matching cell mapping based on OpenAPI anchor.
    /// </summary>
    public static (CellOpenApiMapping? Mapping, string WorksheetName) FindMatchingMapping(
        string openApiAnchor,
        List<WorksheetOpenApiMapping> mappings)
    {
        foreach (var wsMapping in mappings)
        {
            var cellMapping = wsMapping.Mappings.FirstOrDefault(cm =>
                cm.OpenApiRef.Equals(openApiAnchor, StringComparison.OrdinalIgnoreCase));

            if (cellMapping != null)
            {
                return (cellMapping, wsMapping.WorksheetName);
            }
        }

        return (null, string.Empty);
    }

    /// <summary>
    /// Extracts the column part from a cell reference (e.g., "A5" -> "A").
    /// </summary>
    public static string ExtractColumnFromCellReference(string cellReference)
    {
        return new string([.. cellReference.TakeWhile(c => !char.IsDigit(c))]);
    }

    /// <summary>
    /// Extracts the row number from a cell reference like "A1" -> 1, "B23" -> 23
    /// </summary>
    public static int ExtractRowFromCellReference(string cellReference)
    {
        var digitStart = cellReference.IndexOf(cellReference.First(char.IsDigit));
        var rowString = cellReference.Substring(digitStart);
        return int.Parse(rowString);
    }

    /// <summary>
    /// Extracts the column index (0-based) from a cell reference like "A1" -> 0, "B23" -> 1
    /// </summary>
    public static int ExtractColumnIndexFromCellReference(string cellReference)
    {
        var columnString = cellReference.Substring(0, cellReference.IndexOf(cellReference.First(char.IsDigit)));
        int columnIndex = 0;
        for (int i = 0; i < columnString.Length; i++)
        {
            columnIndex = columnIndex * 26 + (columnString[i] - 'A' + 1);
        }
        return columnIndex - 1; // Convert to 0-based index
    }
}