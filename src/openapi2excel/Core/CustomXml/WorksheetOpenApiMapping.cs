using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace OpenApi2Excel.Core.CustomXml
{

	/// <summary>
	/// Represents a mapping between a single cell and an OpenAPI reference.
	/// </summary>
	public class CellOpenApiMapping
	{
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
		public string Worksheet { get; set; } = string.Empty;
		public List<CellOpenApiMapping> Mappings { get; set; } = new();

		public static string Serialize(IEnumerable<WorksheetOpenApiMapping> mappings)
		{
			var doc = new XDocument(
				new XElement("OpenApiMappings",
					mappings.SelectMany(mapping =>
						new[] { new XElement("Worksheet", mapping.Worksheet) }
						.Union(
							mapping.Mappings.Select(m =>
							{
								return new XElement("Mapping", 
									m.Row > 0
										? new XElement("Row", m.Row)
										: new XElement("Cell", m.Cell),
									new XElement("OpenApiRef", m.OpenApiRef)
								);
							}))
						)
				)
			);
			return doc.ToString();
		}

		public static IEnumerable<WorksheetOpenApiMapping> CreateMappings(string worksheet, IEnumerable<(string cell, string openApiRef)> items)
		{
			return [ new WorksheetOpenApiMapping
			{
				Worksheet = worksheet,
				Mappings = [.. items.Select(i => new CellOpenApiMapping
				{
					Cell = i.cell,
					OpenApiRef = i.openApiRef
				})]
			}];
		}
	}
}
