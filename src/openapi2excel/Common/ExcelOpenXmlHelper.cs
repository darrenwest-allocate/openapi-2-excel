using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Office2019.Excel.ThreadedComments;
using openapi2excel.core.Builders.CommentsManagement;

namespace openapi2excel.core.Common;

/// <summary>
/// Helper class for working with Excel files using OpenXML SDK, particularly for threaded comments and worksheet operations.
/// </summary>
public static class ExcelOpenXmlHelper
{

    /// <summary>
    /// Gets the name of a worksheet from its WorksheetPart.
    /// </summary>
    public static string GetWorksheetName(WorkbookPart workbookPart, WorksheetPart worksheetPart)
    {
        var relationshipId = workbookPart.GetIdOfPart(worksheetPart);
        var sheet = workbookPart.Workbook.Sheets?.Elements<Sheet>()
            .FirstOrDefault(s => s.Id?.Value == relationshipId);
        return sheet?.Name?.Value ?? "Unknown";
    }

    /// <summary>
    /// Gets all worksheet names from a workbook.
    /// </summary>
    public static List<string> GetAllWorksheetNames(WorkbookPart workbookPart)
    {
        return [.. workbookPart.WorksheetParts.Select(worksheetPart => GetWorksheetName(workbookPart, worksheetPart))];
    }

    /// <summary>
    /// Gets all worksheet names from an Excel file.
    /// </summary>
    public static List<string> GetAllWorksheetNames(string filePath)
    {
        using var document = SpreadsheetDocument.Open(filePath, false);
        var workbookPart = document.WorkbookPart;
        if (workbookPart == null) return new List<string>();

        return GetAllWorksheetNames(workbookPart);
    }


    /// <summary>
    /// Extracts unresolved threaded comments from all worksheets in an Excel workbook.
    /// </summary>
    public static List<ThreadedCommentWithContext> ExtractUnresolvedThreadedCommentsFromWorkbook(string filePath, bool annotateWithOpenApiAnchors = false)
    {
        return ExtractThreadedCommentsFromWorkbook(filePath, includeResolved: false, includeUnresolved: true, annotateWithOpenApiAnchors);
    }

    /// <summary>
    /// Extracts unresolved threaded comments and annotates them with OpenAPI anchors from custom XML mappings.
    /// </summary>
    public static List<ThreadedCommentWithContext> ExtractAndAnnotateUnresolvedComments(string filePath)
    {
        return ExtractUnresolvedThreadedCommentsFromWorkbook(filePath, annotateWithOpenApiAnchors: true);
    }

    /// <summary>
    /// Extracts all threaded comments (both resolved and unresolved) and annotates them with OpenAPI anchors from custom XML mappings.
    /// </summary>
    public static List<ThreadedCommentWithContext> ExtractAndAnnotateAllComments(string filePath)
    {
        return ExtractThreadedCommentsFromWorkbook(filePath, includeResolved: true, includeUnresolved: true, annotateWithOpenApiAnchors: true);
    }

    /// <summary>
    /// Extracts only resolved threaded comments and annotates them with OpenAPI anchors from custom XML mappings.
    /// </summary>
    public static List<ThreadedCommentWithContext> ExtractAndAnnotateResolvedComments(string filePath)
    {
        return ExtractThreadedCommentsFromWorkbook(filePath, includeResolved: true, includeUnresolved: false, annotateWithOpenApiAnchors: true);
    }


    /// <summary>
    /// Extracts all threaded comments from all worksheets in an Excel workbook, with optional filtering by resolution status.
    /// </summary>
    /// <param name="filePath">Path to the Excel workbook</param>
    /// <param name="includeResolved">If true, includes resolved comments; if false, only unresolved comments</param>
    /// <param name="includeUnresolved">If true, includes unresolved comments; if false, only resolved comments</param>
    /// <param name="annotateWithOpenApiAnchors">If true, annotates comments with OpenAPI anchors from custom XML mappings</param>
    /// <returns>List of threaded comments with their worksheet context, optionally annotated with OpenAPI anchors</returns>
    private static List<ThreadedCommentWithContext> ExtractThreadedCommentsFromWorkbook(string filePath, bool includeResolved = true, bool includeUnresolved = true, bool annotateWithOpenApiAnchors = false)
    {
        var comments = new List<ThreadedCommentWithContext>();
		using var document = SpreadsheetDocument.Open(filePath, false);
		var workbookPart = document.WorkbookPart;
		if (workbookPart == null) return comments;

		foreach (var worksheet in workbookPart.WorksheetParts)
		{
			var worksheetName = GetWorksheetName(workbookPart, worksheet);
			var threadedCommentsPart = worksheet.GetPartsOfType<WorksheetThreadedCommentsPart>().FirstOrDefault();
			if (threadedCommentsPart == null || threadedCommentsPart.ThreadedComments == null) continue;

			foreach (var comment in threadedCommentsPart.ThreadedComments.Elements<ThreadedComment>())
			{
				var isResolved = comment.Done == "1";
				if ((isResolved && includeResolved) || (!isResolved && includeUnresolved))
				{
					comments.Add(new ThreadedCommentWithContext(comment, worksheetName, filePath));
				}
			}
		}
		if (annotateWithOpenApiAnchors)
		{
			var mappings = ExtractCustomXmlMappingsFromWorkbook(workbookPart);
			foreach (var comment in comments)
			{
				var mapping = MapToCell(mappings, comment) ?? MapToRow(mappings, comment);
				if (mapping != null)
				{
					comment.OpenApiAnchor = mapping.OpenApiRef;
				}
			}
		}
		return comments;
	}

    private static CellOpenApiMapping? MapToRow(List<WorksheetOpenApiMapping> mappings, ThreadedCommentWithContext comment)
    {
        return mappings.FirstOrDefault(m =>
            m.WorksheetName.Equals(comment.WorksheetName, StringComparison.OrdinalIgnoreCase))
            ?.Mappings.FirstOrDefault(cellMapping =>
                cellMapping.Row.Equals(RowForCellReference(comment.CellReference)));
    }

    private static CellOpenApiMapping? MapToCell(List<WorksheetOpenApiMapping> mappings, ThreadedCommentWithContext comment)
    {
        return mappings.FirstOrDefault(m =>
            m.WorksheetName.Equals(comment.WorksheetName, StringComparison.OrdinalIgnoreCase))
            ?.Mappings.FirstOrDefault(cellMapping =>
                cellMapping.Cell.Equals(comment.CellReference, StringComparison.OrdinalIgnoreCase));
    }

    private static int RowForCellReference(string cellReference)
    {
        var rowPart = new string(cellReference.SkipWhile(c => !char.IsDigit(c)).ToArray());
        if (int.TryParse(rowPart, out var rowNumber))
            return rowNumber;
        else
            return -1;
    }


    /// <summary>
    /// Extracts custom XML mappings from workbook using the existing ExcelCustomXmlHelper infrastructure
    /// </summary>
    public static List<WorksheetOpenApiMapping> ExtractCustomXmlMappingsFromWorkbook(string filePath)
    {
        using var document = SpreadsheetDocument.Open(filePath, false);
        return ExtractCustomXmlMappingsFromWorkbook(document.WorkbookPart!);
    }

    /// <summary>
    /// Extracts custom XML mappings from workbook using the existing ExcelCustomXmlHelper infrastructure
    /// </summary>
    public static List<WorksheetOpenApiMapping> ExtractCustomXmlMappingsFromWorkbook(WorkbookPart workbookPart)
    {
        var worksheetMappings = new List<WorksheetOpenApiMapping>();
        if (workbookPart == null) return worksheetMappings;
        var mappingsDictionary = ExcelCustomXmlHelper.ReadAllCustomXmlMappings(workbookPart);
        foreach (var worksheetName in GetAllWorksheetNames(workbookPart))
        {
            if (!mappingsDictionary.TryGetValue(worksheetName, out var xmlDocument))
                continue;
            var cellMappings = xmlDocument.Root?
                .Elements("MapOpenApiRef")
                .Select(element => new CellOpenApiMapping(element)).ToList() ?? [];
            if (cellMappings.Any())
            {
                worksheetMappings.Add(new WorksheetOpenApiMapping(worksheetName) { Mappings = cellMappings });
            }
        }
        return worksheetMappings;
    }
}
