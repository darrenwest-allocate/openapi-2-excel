using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Office2019.Excel.ThreadedComments;
using openapi2excel.core.Builders.CustomXml;

namespace openapi2excel.core.Common;

/// <summary>
/// Helper class for working with Excel files using OpenXML SDK, particularly for threaded comments and worksheet operations.
/// </summary>
public static class ExcelOpenXmlHelper
{

    /// <summary>
    /// Gets the name of a worksheet from its WorksheetPart.
    /// </summary>
    /// <param name="workbookPart">The workbook part containing the worksheet</param>
    /// <param name="worksheetPart">The worksheet part to get the name for</param>
    /// <returns>The worksheet name or "Unknown" if not found</returns>
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
    /// <param name="workbookPart">The workbook part to get worksheet names from</param>
    /// <returns>List of all worksheet names in the workbook</returns>
    public static List<string> GetAllWorksheetNames(WorkbookPart workbookPart)
    {
        var worksheetNames = new List<string>();

        foreach (var worksheetPart in workbookPart.WorksheetParts)
        {
            var worksheetName = GetWorksheetName(workbookPart, worksheetPart);
            worksheetNames.Add(worksheetName);
        }

        return worksheetNames;
    }

    /// <summary>
    /// Gets all worksheet names from an Excel file.
    /// </summary>
    /// <param name="filePath">Path to the Excel file</param>
    /// <returns>List of all worksheet names in the workbook</returns>
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
    /// <param name="filePath">Path to the Excel workbook</param>
    /// <param name="annotateWithOpenApiAnchors">If true, annotates comments with OpenAPI anchors from custom XML mappings</param>
    /// <returns>List of unresolved threaded comments with their worksheet context, optionally annotated with OpenAPI anchors</returns>
    public static List<ThreadedCommentWithContext> ExtractUnresolvedThreadedCommentsFromWorkbook(string filePath, bool annotateWithOpenApiAnchors = false)
    {
        var unresolvedComments = new List<ThreadedCommentWithContext>();

        using (var document = SpreadsheetDocument.Open(filePath, false))
        {
            var workbookPart = document.WorkbookPart;
            if (workbookPart == null) return unresolvedComments;

            // Iterate through all worksheets
            foreach (var worksheetPart in workbookPart.WorksheetParts)
            {
                // Get worksheet name from the workbook
                var worksheetName = GetWorksheetName(workbookPart, worksheetPart);

                // Check if this worksheet has threaded comments
                var threadedCommentsPart = worksheetPart.GetPartsOfType<WorksheetThreadedCommentsPart>().FirstOrDefault();
                if (threadedCommentsPart == null) continue;

                var threadedComments = threadedCommentsPart.ThreadedComments;
                if (threadedComments == null) continue;

                // Extract unresolved threaded comments with worksheet context
                foreach (var comment in threadedComments.Elements<ThreadedComment>())
                {
                    var xmlContent = comment.OuterXml;
                    bool isResolved = xmlContent.Contains("resolved=\"1\"");

                    if (!isResolved)
                    {
                        unresolvedComments.Add(new ThreadedCommentWithContext
                        {
                            Comment = comment,
                            WorksheetName = worksheetName
                        });
                    }
                }
            }
        }

        // Optionally annotate with OpenAPI anchors
        if (annotateWithOpenApiAnchors)
        {
            var mappings = ExtractCustomXmlMappingsFromWorkbook(filePath);

            foreach (var comment in unresolvedComments)
            {
                var mapping = mappings.FirstOrDefault(m =>
                    m.Worksheet.Equals(comment.WorksheetName, StringComparison.OrdinalIgnoreCase))
                    ?.Mappings.FirstOrDefault(cellMapping =>
                        cellMapping.Cell.Equals(comment.CellReference, StringComparison.OrdinalIgnoreCase));

                if (mapping != null)
                {
                    comment.OpenApiAnchor = mapping.OpenApiRef;
                }
                else
                {
                    // try to map on the row
                    mapping = mappings.FirstOrDefault(m =>
                        m.Worksheet.Equals(comment.WorksheetName, StringComparison.OrdinalIgnoreCase))
                        ?.Mappings.FirstOrDefault(cellMapping =>
                            cellMapping.Row.Equals(RowForCellReference(comment.CellReference)));
                    if (mapping != null)
                    {
                        comment.OpenApiAnchor = mapping.OpenApiRef;
                    }
                }


            }
        }

        return unresolvedComments;
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
    /// Extracts unresolved threaded comments and annotates them with OpenAPI anchors from custom XML mappings.
    /// </summary>
    /// <param name="filePath">Path to the Excel workbook</param>
    /// <returns>List of unresolved threaded comments annotated with OpenAPI anchors where available</returns>
    public static List<ThreadedCommentWithContext> ExtractAndAnnotateUnresolvedComments(string filePath)
    {
        return ExtractUnresolvedThreadedCommentsFromWorkbook(filePath, annotateWithOpenApiAnchors: true);
    }

    /// <summary>
/// Extracts custom XML mappings from workbook using the existing ExcelCustomXmlHelper infrastructure
/// </summary>
/// <param name="filePath">Path to the Excel workbook</param>
/// <returns>List of worksheet mappings with OpenAPI anchors</returns>
    private static List<WorksheetOpenApiMapping> ExtractCustomXmlMappingsFromWorkbook(string filePath)
    {
        var worksheetMappings = new List<WorksheetOpenApiMapping>();

        using (var document = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(filePath, false))
        {
            var workbookPart = document.WorkbookPart;
            if (workbookPart == null) return worksheetMappings;

            var worksheetNames = GetAllWorksheetNames(workbookPart);

            foreach (var worksheetName in worksheetNames)
            {
                try
                {
                    // Use the existing ExcelCustomXmlHelper to read mappings for each worksheet
                    var xmlContent = ExcelCustomXmlHelper.ReadCustomXmlMapping(filePath, worksheetName);

                    // Parse the XML content using the proper Core models
                    var cellMappings = ParseCustomXmlContent(xmlContent);

                    if (cellMappings.Any())
                    {
                        worksheetMappings.Add(new WorksheetOpenApiMapping
                        {
                            Worksheet = worksheetName,
                            Mappings = cellMappings
                        });
                    }
                }
                catch (System.IO.FileNotFoundException)
                {
                    // Skip worksheets that don't have custom XML mappings
                    continue;
                }
                catch (System.InvalidOperationException)
                {
                    // Skip worksheets that can't be processed due to invalid operation
                    continue;
                }
                catch (System.Xml.XmlException)
                {
                    // Skip worksheets with malformed custom XML
                    continue;
                }
            }
        }

        return worksheetMappings;
    }

    /// <summary>
    /// Parse custom XML content to extract mappings
    /// </summary>
    /// <param name="xmlContent">XML content from custom XML part</param>
    /// <returns>List of cell mappings for the worksheet</returns>
    private static List<CellOpenApiMapping> ParseCustomXmlContent(string xmlContent)
    {
        var mappings = new List<CellOpenApiMapping>();

        try
        {
            var doc = XDocument.Parse(xmlContent);

            // Look for mapping elements according to the existing structure
            var mappingElements = doc.Descendants("Mapping");

            foreach (var mapping in mappingElements)
            {
                var worksheet = mapping.Element("Worksheet")?.Value;
                var cell = mapping.Element("Cell")?.Value;
                var openApiRef = mapping.Element("OpenApiRef")?.Value;

                if (!string.IsNullOrEmpty(cell) && !string.IsNullOrEmpty(openApiRef))
                {
                    mappings.Add(new CellOpenApiMapping
                    {
                        Cell = cell,
                        OpenApiRef = openApiRef
                    });
                }
            }
        }
        catch (System.Xml.XmlException)
        {
            // Return empty list if XML parsing fails
        }
        catch (ArgumentException)
        {
            // Return empty list if XML parsing fails
        }

        return mappings;
    }
}
