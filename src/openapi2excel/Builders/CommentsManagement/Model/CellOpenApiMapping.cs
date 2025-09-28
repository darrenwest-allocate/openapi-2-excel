using System.Xml.Linq;
namespace openapi2excel.core.Builders.CommentsManagement;

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
