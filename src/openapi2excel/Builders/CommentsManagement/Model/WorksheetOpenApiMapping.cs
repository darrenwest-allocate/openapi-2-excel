using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
namespace openapi2excel.core.Builders.CommentsManagement;

/// <summary>
/// Represents a mapping between worksheet cells and OpenAPI references for a single worksheet.
/// </summary>
public class WorksheetOpenApiMapping
{
    /* TODO: correct this use of a global static variable because if forces all tests to clear this list to not clash with each other */
    public static List<WorksheetOpenApiMapping> AllWorksheetMappings { get; } = new();

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
