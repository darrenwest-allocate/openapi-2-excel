
using System;
using System.Collections.Generic;

namespace openapi2excel.core.Builders.CustomXml;

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
