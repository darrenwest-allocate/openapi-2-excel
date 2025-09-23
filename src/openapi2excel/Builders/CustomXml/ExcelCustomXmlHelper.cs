
using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;

namespace openapi2excel.core.Builders.CustomXml;

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
			var relId = $"mapping_{worksheetName.Replace(" ", "_")}";
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

	public static string ReadCustomXmlMapping(string filePath, string worksheetName)
	{
		using var doc = SpreadsheetDocument.Open(filePath, false);
		var wbPart = doc.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart not found in the document.");
		var relId = $"mapping_{worksheetName.Replace(" ", "_")}";
		if (wbPart.Parts.FirstOrDefault(p => p.RelationshipId == relId && p.OpenXmlPart is CustomXmlPart).OpenXmlPart is not CustomXmlPart part)
			throw new InvalidOperationException($"Custom XML part for worksheet '{worksheetName}' not found.");
		using var stream = part.GetStream();
		using var reader = new StreamReader(stream);
		return reader.ReadToEnd();
	}
}
