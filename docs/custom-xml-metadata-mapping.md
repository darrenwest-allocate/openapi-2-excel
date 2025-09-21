# Using Custom XML Parts in Excel Workbooks for Metadata Mapping

## Overview
Excel workbooks (.xlsx) are ZIP archives containing multiple XML files. The OpenXML standard allows you to add custom XML parts to a workbook, which can store arbitrary metadata invisible to end users.

## Use Case: Mapping Excel Cells to OpenAPI Entities
Custom XML parts can be used to store mappings between worksheet cells and their source OpenAPI entities. This is useful for features like comment migration, traceability, and advanced automation.

## Example Custom XML Structure
```xml
<OpenApiMappings>
  <Mapping>
    <Worksheet>Operations</Worksheet>
    <Cell>A5</Cell>
    <OpenApiRef>paths./pets.get.responses.200</OpenApiRef>
  </Mapping>
  <Mapping>
    <Worksheet>Parameters</Worksheet>
    <Cell>B12</Cell>
    <OpenApiRef>components.parameters.PetId</OpenApiRef>
  </Mapping>
  <!-- More mappings... -->
</OpenApiMappings>
```
- `<Worksheet>`: Name of the worksheet.
- `<Cell>`: Cell address (e.g., A5, B12).
- `<OpenApiRef>`: OpenAPI anchor (see below for syntax).

## Defining Referenceable OpenAPI JSON Anchors

To reliably map Excel cells to their OpenAPI source, each mapping should use a referenceable OpenAPI anchor. The anchor syntax should be deterministic and unambiguous, allowing programmatic lookup in the OpenAPI document.

### Recommended Anchor Syntax

- **Paths and Operations:**
  - Format: `paths.{path}.{method}`
  - Example: `paths./pets.get`
- **Responses:**
  - Format: `paths.{path}.{method}.responses.{status}`
  - Example: `paths./pets.get.responses.200`
- **Parameters:**
  - Format: `components.parameters.{parameterName}`
  - Example: `components.parameters.PetId`
- **Schemas:**
  - Format: `components.schemas.{schemaName}`
  - Example: `components.schemas.Pet`
- **Properties:**
  - Format: `components.schemas.{schemaName}.properties.{propertyName}`
  - Example: `components.schemas.Pet.properties.name`

### General Rules
- Use the OpenAPI JSON path, with keys separated by dots.
- For path segments, encode slashes as-is (e.g., `/pets`).
- For arrays or indexed items, use the array index (e.g., `parameters.0`).
- This syntax should match the way you traverse the OpenAPI document in code.

### Example
```xml
<OpenApiRef>paths./pets.get.responses.200</OpenApiRef>
<OpenApiRef>components.schemas.Pet.properties.name</OpenApiRef>
```

This approach ensures that every cell mapped in the custom XML can be traced back to a unique location in the OpenAPI JSON.

## How to Add Custom XML with ClosedXML/OpenXML SDK
- ClosedXML does not directly expose custom XML APIs, but you can access the underlying OpenXML package.
- Example pseudocode:

```csharp
using (var workbook = new XLWorkbook("file.xlsx"))
{
    var doc = new XDocument(
        new XElement("OpenApiMappings",
            new XElement("Mapping",
                new XElement("Worksheet", "Operations"),
                new XElement("Cell", "A5"),
                new XElement("OpenApiRef", "paths./pets.get.responses.200")
            )
            // ... more mappings
        )
    );

    var customXml = workbook.WorkbookPart.AddNewPart<CustomXmlPart>();
    using (var stream = customXml.GetStream())
    {
        doc.Save(stream);
    }
    // Save workbook as usual
}
```

## Advantages
- **Invisible to users**
- **Programmatically accessible**
- **Flexible and extensible**



## Recommendation: Use a Meta Custom XML Part (Per Worksheet)

Given the likelihood of large workbooks and future feature expansion, it is recommended to adopt a Meta Custom XML part approach, with mapping parts separated by worksheet:

- **Meta Part:** Acts as a manifest, listing all mapping XML parts (one per worksheet) and storing global metadata (e.g., version, generation date, feature flags).
- **Mapping Parts:** Each worksheet has its own Custom XML part for modularity and scalability.

### Rationale
- Each worksheet in the workbook describes an individual endpoint from the OpenAPI spec.
- There is also a summary worksheet used for navigation between endpoint worksheets.
- Separating mapping parts by worksheet (not by OpenAPI content type) aligns with the workbook structure and makes it easier to manage mappings as endpoints are added, removed, or updated.

### Benefits
- Scales well for large and complex workbooks.
- Supports modular development and future extensibility (e.g., new features, per-worksheet metadata).
- Facilitates discovery and management of all mapping and metadata parts.

### Example Meta Part Structure
```xml
<OpenApiExcelMeta>
  <Mappings>
    <MappingPart worksheet="GET /pets">customXml/mapping-get-pets.xml</MappingPart>
    <MappingPart worksheet="POST /pets">customXml/mapping-post-pets.xml</MappingPart>
    <MappingPart worksheet="Summary">customXml/mapping-summary.xml</MappingPart>
    <!-- More mapping parts, one per worksheet... -->
  </Mappings>
  <Version>1.0</Version>
  <Generated>2025-09-21</Generated>
</OpenApiExcelMeta>
```

This approach is robust and future-proof, making it easier to add new features and manage large workbooks as the project evolves.

---
This technique is recommended for robust, future-proof mapping of Excel content to its OpenAPI source, especially for advanced features like comment migration and traceability.
