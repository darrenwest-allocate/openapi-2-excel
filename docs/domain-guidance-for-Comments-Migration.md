# Domain Guidance: Excel Comments Migration System

## Overview

The Comments Migration system enables the transfer of unresolved Excel comments from an existing workbook to a newly generated workbook, maintaining traceability to OpenAPI entities. This feature preserves stakeholder feedback across OpenAPI specification updates.

## System Architecture

### Core Components

**1. Custom XML Metadata Mapping**
- **Purpose**: Creates invisible cell-to-OpenAPI entity mappings within Excel workbooks
- **Structure**: Meta Custom XML part + individual mapping parts per worksheet
- **Location**: Stored as custom XML parts in the OpenXML package structure
- **Key Classes**: Integration with `ExcelOpenXmlHelper`, worksheet builders

**2. Comment Migration Engine**
- **Purpose**: Extracts, maps, and transfers threaded comments between workbooks
- **Strategy**: Three migration types (Exact Match, Type A NoAnchor, Type B NoWorksheet)
- **Technology**: Hybrid Legacy + ThreadedComment approach for maximum Excel compatibility

**3. OpenAPI Anchor System**
- **Purpose**: Provides deterministic references to OpenAPI entities for stable mapping
- **Format**: Dot-notation paths (e.g., `paths./pets.get.responses.200`)
- **Coverage**: All OpenAPI content types (paths, components, schemas, parameters)

### Migration Types

**Exact Match Migration**
- Comments migrate to cells with identical OpenAPI anchors
- Preserves original cell relationships and context
- Highest fidelity migration type

**Type A (NoAnchor) Migration**
- Comments from cells without OpenAPI anchors
- Migrates to nearest `/TitleRow` heading row in same worksheet
- Preserves original column, handles collisions intelligently

**Type B (NoWorksheet) Migration**
- Comments from worksheets not present in new workbook
- Migrates to Info sheet, column V, stacking downward
- Includes original worksheet context in comment metadata

## Technical Implementation Details

### Custom XML Structure

**Meta Part Structure:**
```xml
<OpenApiExcelMeta>
  <Mappings>
    <MappingPart worksheet="GET /pets">customXml/mapping-get-pets.xml</MappingPart>
    <!-- One per worksheet... -->
  </Mappings>
  <Version>1.0</Version>
  <Generated>2025-09-21</Generated>
</OpenApiExcelMeta>
```

### OpenAPI Anchor Syntax

**Paths and Operations:**
- Format: `paths.{path}.{method}`
- Example: `paths./pets.get`

**Responses:**
- Format: `paths.{path}.{method}.responses.{status}`
- Example: `paths./pets.get.responses.200`

**Components:**
- Parameters: `components.parameters.{parameterName}`
- Schemas: `components.schemas.{schemaName}`
- Properties: `components.schemas.{schemaName}.properties.{propertyName}`

### Comment Technology Stack

**Hybrid Approach Benefits:**
- **Legacy Comments**: Ensure visibility in older Excel versions (using `ClosedXML.Excel`)
- **Threaded Comments**: Support modern conversation features and threading (using `DocumentFormat.OpenXml.Office2019.Excel`)
- **Persons Integration**: Maintain author attribution and timestamps

**Key Implementation Pattern:**
```csharp
// Simplified example - actual implementation is more complex
var legacyComment = cell.CreateComment();
legacyComment.SetText(commentText);

var threadedComment = cell.CreateThreadedComment();
threadedComment.SetText(commentText);
threadedComment.SetAuthor(author);
```

## Integration Points

### CLI Integration
- `--existing-workbook` parameter accepts path to old workbook
- Migration occurs automatically during new workbook generation
- No separate migration step required

### Workbook Generation Pipeline
- Custom XML mapping generation integrated into worksheet builders
- Comment migration occurs before final workbook save
- Preserves all existing workbook generation features

### Testing Infrastructure
- Behavioral tests validate end-to-end migration scenarios
- Sample test data: `sample-api-gw.json`, `sample-api-gw-with-mappings.xlsx`
- Hybrid comment detection via `ExcelOpenXmlHelper.ExtractAndAnnotateAllComments`

## Development Guidelines

### When Modifying This System

**1. Understand the Mapping First**
- Custom XML mappings are the foundation - changes here affect all migration logic
- Test with `ExcelOpenXmlHelper` methods to verify mapping integrity
- Remember: mappings are per-worksheet, not per-OpenAPI-entity

**2. Preserve Thread Integrity**
- Comment threads (parent-child relationships) must be maintained
- Use the established Legacy + ThreadedComment pattern
- Test in actual Excel to verify comment visibility and threading

**3. Test Behavioral Outcomes**
- Tests validate observable behavior, not implementation details
- Focus on: "Can I see the comment in Excel?" rather than "Does the code path execute?"
- Use manual Excel verification for complex threading scenarios

### Adding New Migration Types

**Pattern to Follow:**
1. Define new `CommentMigrationFailureReason` enum value (or success tracking)
2. Add detection logic in comment extraction phase
3. Implement placement strategy (worksheet + cell selection)
4. Extend ID mapping dictionary for tracking
5. Add behavioral test with real Excel validation

### Extending Custom XML Capabilities

**Current Limitations:**
- One mapping per cell (no multi-entity cells)
- Simple dot-notation anchors (no complex query support)
- Worksheet-based partitioning only

**Extension Points:**
- Additional metadata in mapping parts (versioning, feature flags)
- Cross-worksheet reference support
- Complex anchor expressions for computed cells

## Troubleshooting Common Issues

### Comments Not Visible in Excel
- **Likely Cause**: Missing legacy comment component
- **Solution**: Ensure both legacy and threaded comments are created
- **Test**: Open in Excel to look for opening errors or simply missing comments

### Migration Mismatches
- **Likely Cause**: OpenAPI anchor changes between versions
- **Solution**: Review anchor generation logic in worksheet builders
- **Test**: Compare custom XML between old and new workbooks

### Performance Issues with Large Workbooks
- **Likely Cause**: Custom XML processing overhead
- **Solution**: Consider caching mapping lookups or partial processing
- **Test**: Profile with large OpenAPI specifications (100+ endpoints)

### Excel Repair Dialogs
- **Likely Cause**: Malformed threaded comment XML or missing persons parts
- **Solution**: Validate persons part creation and comment structure. Stick to OpenXML documented examples.
- **Test**: Open generated workbook in Excel and check for repair prompts
