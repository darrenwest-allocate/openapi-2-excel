using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Bibliography;

namespace openapi2excel.core.Builders.CustomXml;

/// <summary>
/// Represents a mapping between a single cell and an OpenAPI reference.
/// </summary>
public class CellOpenApiMapping
{
	public CellOpenApiMapping()	{ }

	public CellOpenApiMapping(XElement element)
	{
		if (element == null) return;
		Cell = element.Attribute("Cell")?.Value ?? string.Empty;
		OpenApiRef = element.Value ?? string.Empty;
		if (int.TryParse(element.Attribute("Row")?.Value, out var rowNumber)) Row = rowNumber;
	}

	public string Cell { get; set; } = string.Empty;
	/// <summary>
	/// The OpenAPI JSON reference this cell maps to, e.g. "paths./pets.get.responses.200"
	/// defined with the AnchorGenerator class.
	/// </summary>
	public string OpenApiRef { get; set; } = string.Empty;
	/// <summary>
	/// The row number in the worksheet (1-based). This is set when the mapping is created for a whole row.
	/// </summary>
	public int Row { get; internal set; }
}

/// <summary>
/// Represents a mapping between worksheet cells and OpenAPI references for a single worksheet.
/// </summary>
public class WorksheetOpenApiMapping
{
	public static List<WorksheetOpenApiMapping> AllWorksheetMappings { get; } = new(); /* I regret doing this */

	/// <summary>
	/// Clears the static AllWorksheetMappings collection. Used for test isolation.
	/// </summary>
	public static void ClearAllMappings()
	{
		AllWorksheetMappings.Clear();
	}

	public string WorksheetName { get; set; }
	public List<CellOpenApiMapping> Mappings { get; set; } = new();

	public WorksheetOpenApiMapping(string worksheetName)
	{
		WorksheetName = worksheetName;
	}

	public static string Serialize(WorksheetOpenApiMapping mapping)
	{ 
		return Serialize([mapping]);
	}
	public static string Serialize(IEnumerable<WorksheetOpenApiMapping> mappings)
	{
		var doc = new XDocument(
			new XElement("OpenApiMappings",
				mappings.SelectMany(mapping =>
					new[] { new XElement("Worksheet", mapping.WorksheetName) }
					.Union(
						mapping.Mappings.Select(m =>
						{
							return new XElement("MapOpenApiRef", 
								m.Row > 0
									? new XAttribute("Row", m.Row)
									: new XAttribute("Cell", m.Cell),
								new XText(m.OpenApiRef)
							);
						}))
					)
			)
		);
		return doc.ToString();
	}
}
