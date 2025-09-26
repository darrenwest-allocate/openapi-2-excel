
using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using System.Xml.Linq;

namespace openapi2excel.core.Builders.CommentsManagement;

public static class ExcelCustomXmlHelper
{
	public static void WriteCustomXmlMapping(string filePath, WorksheetOpenApiMapping worksheetMapping)
		=> WriteCustomXmlMapping(filePath, worksheetMapping.WorksheetName, WorksheetOpenApiMapping.Serialize(worksheetMapping));

	public static void WriteCustomXmlMapping(string filePath, string worksheetName, string xmlContent)
	{
		// Ensure file exists (create empty workbook if needed)
		if (!File.Exists(filePath))
		{
			using var mem = new MemoryStream();
			using (var doc = SpreadsheetDocument.Create(mem, SpreadsheetDocumentType.Workbook))
			{
				var wbPart = doc.AddWorkbookPart();
				wbPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();
				wbPart.Workbook.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Sheets());
			}
			File.WriteAllBytes(filePath, mem.ToArray());
		}

		using (var doc = SpreadsheetDocument.Open(filePath, true))
		{
			var wbPart = doc.WorkbookPart ?? doc.AddWorkbookPart();
			if (wbPart.Workbook == null)
				wbPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();

			// Use a valid NCName for the relationship ID
			var relId = GetRelId(worksheetName);
			// Remove any existing custom XML part for this worksheet by relId
			var existing = wbPart.Parts.FirstOrDefault(p => p.RelationshipId == relId && p.OpenXmlPart is CustomXmlPart);
			if (existing.OpenXmlPart != null)
				wbPart.DeletePart(existing.OpenXmlPart);

			var customXmlPart = wbPart.AddCustomXmlPart(CustomXmlPartType.CustomXml, relId);
			using var stream = customXmlPart.GetStream(FileMode.Create, FileAccess.Write);
			using var writer = new StreamWriter(stream);
			writer.Write(xmlContent);
		}
	}

	private static string GetRelId(string worksheetName)
	{
		return $"mapping_{worksheetName.Replace(" ", "_")}";
	}

	/// <summary>
	/// Reads the custom XML mapping for a specific worksheet from an Excel file.
	/// </summary>
	public static Dictionary<string, XDocument> ReadAllCustomXmlMappings(WorkbookPart workbookPart)
	{ 
		var dictionary = new Dictionary<string, XDocument>(StringComparer.OrdinalIgnoreCase);
		foreach (var partInfo in workbookPart.Parts.Where(p => p.OpenXmlPart is CustomXmlPart))
		{
			using var stream = partInfo.OpenXmlPart.GetStream();
			using var reader = new StreamReader(stream);
			var document = XDocument.Load(reader);
			var worksheetName = document.Root?.Element("Worksheet")?.Value;
			if (string.IsNullOrWhiteSpace(worksheetName))
				continue;
			dictionary[worksheetName] = document;
		}
		return dictionary;
	}

	/// <summary>
	/// Reads the custom XML mapping for a specific worksheet from an Excel file.
	/// </summary>
	public static Dictionary<string, XDocument> ReadAllCustomXmlMappings(string filePath)
	{
		using var doc = SpreadsheetDocument.Open(filePath, false);
		var wbPart = doc.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart not found in the document.");
		return ReadAllCustomXmlMappings(wbPart);
	}
}
