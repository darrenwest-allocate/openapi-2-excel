using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace OpenApi2Excel.Core.CustomXml
{
    /// <summary>
    /// Represents the meta custom XML part for the workbook, listing all mapping parts and global metadata.
    /// </summary>
    public class OpenApiExcelMeta
    {
        public List<MappingPartRef> Mappings { get; set; } = new();
        public string Version { get; set; } = "1.0";
        public DateTime Generated { get; set; } = DateTime.UtcNow;
    }

    public class MappingPartRef
    {
        public string Worksheet { get; set; } = string.Empty;
        public string PartUri { get; set; } = string.Empty;
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
						mapping.Mappings.Select(m =>
							new XElement("Mapping",
								new XElement("Worksheet", mapping.Worksheet),
								new XElement("Cell", m.Cell),
								new XElement("OpenApiRef", m.OpenApiRef)
							)
						)
					)
				)
			);
			return doc.ToString();
		}

		public static IEnumerable<WorksheetOpenApiMapping> CreateMappings(string worksheet, IEnumerable<(string cell, string openApiRef)> items)
		{
			return new[] {new WorksheetOpenApiMapping
			{
				Worksheet = worksheet,
				Mappings = items.Select(i => new CellOpenApiMapping
					{
						Cell = i.cell,
						OpenApiRef = i.openApiRef
					}).ToList()
				}};
		}
    }

    public class CellOpenApiMapping
    {
        public string Cell { get; set; } = string.Empty;
        public string OpenApiRef { get; set; } = string.Empty;
    }
}
