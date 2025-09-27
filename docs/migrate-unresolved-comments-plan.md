# See also: [docs/unit-testing-policy-migrate-comments.md](unit-testing-policy-migrate-comments.md) for the unit testing policy and behavioral test outline for this feature.

# Implementation Plan: Migrating Unresolved Excel Comments
Reference: Parent Issue #5

**Note:** This implementation will use the ClosedXML library for all Excel operations. Only modern threaded comments are supported; legacy (non-threaded) comments will not be migrated.


## 1. Identifying and Extracting Unresolved Comments
- Use the ClosedXML .NET library to read both old and new Excel workbooks.
- Iterate through all worksheets, cells, and threaded comments in the old workbook.
- For each threaded comment, check if it is unresolved.
- Extract:
  - Worksheet name
  - Cell address
  - Comment text
  - Author
  - Thread/conversation metadata (if available)

**See also:** [docs/custom-xml-metadata-mapping.md](custom-xml-metadata-mapping.md) for details on how to map cells to OpenAPI entities using Custom XML parts.

## 2. Matching Comments to Relevant Sections in the New Workbook
- Determine the logical "anchor" for each comment (e.g., operationId, parameter name, schema property).
- If possible, encode a hidden mapping or metadata in the workbook linking each cell to its OpenAPI entity.
- In the new workbook, use the same logic to map cells to OpenAPI entities.
- Match old comments to new cells by comparing these anchors, not just cell addresses.
- If no direct match, flag for manual review.

**See also:** [docs/custom-xml-metadata-mapping.md](custom-xml-metadata-mapping.md) for recommended anchor syntax and metadata structure.

## 3. Handling Conflicts, Missing References, or Structure Changes
- **Missing References**: If a comment's anchor no longer exists, collect for reporting.
- **Conflicts**: If a cell in the new workbook already has a comment:
  - Merge comments (append old to new, with attribution)
  - Keep both as separate threads (if supported)
  - Flag for manual resolution
- **Structure Changes**: If layout changes significantly, fallback to fuzzy matching (e.g., by operation summary, parameter name, or property path).

## 3A. Handling Non-Migratable Comments (Types A & B)

### Type A: "NoAnchor" Comments
Comments found on worksheets that exist in the new workbook but are located on rows without OpenAPI anchors:
- **Migration Strategy**: Migrate to the same worksheet, preserving the original column
- **Row Placement**: Find the nearest row above that contains an anchor ending with `/TitleRow` (heading row)
- **Collision Handling**: If target cell has content/comments, place in the row below the heading row
- **Full Thread Migration**: Migrate complete comment threads (root + replies) using the established Legacy + ThreadedComment approach

### Type B: "NoWorksheet" Comments  
Comments found on worksheets that will not exist in the new workbook:
- **Migration Strategy**: Migrate to the Info sheet (identified by `OpenApiDocumentationLanguageConst.Info`)
- **Placement**: Column V, starting at row 1, stacking downward for subsequent comments
- **Full Thread Migration**: Preserve complete comment threads as threaded comments
- **Metadata Preservation**: Include reference to original worksheet name in comment context

## 4. Integrating the Old Workbook as an Input
- Update CLI/API to accept an optional "previous workbook" input parameter.
- Processing flow:
  1. Load old workbook and extract unresolved comments.
  2. Generate new workbook from updated OpenAPI spec.
  3. Before saving new workbook, inject migrated comments into appropriate locations.

## 5. Tools, Libraries, and APIs
- **ClosedXML**: High-level .NET library for Excel, supports threaded comments (required for this feature).
- **Diff/Merge Libraries**: For fuzzy matching or diffing OpenAPI entities.
- **Custom XML Metadata Mapping**: See [docs/custom-xml-metadata-mapping.md](custom-xml-metadata-mapping.md) for implementation details and rationale for using Custom XML parts for cell-to-OpenAPI mapping.

## 6. Risks, Edge Cases, and Testability
- **Risks**:
  - Major spec changes may make matching unreliable.
  - Comments on deleted/renamed entities may be orphaned.
  - Threaded comments or advanced Excel features may not be fully supported.
- **Edge Cases**:
  - Multiple comments per cell.
  - Comments on merged cells or non-standard layouts.
  - Comments on summary/info sheets.
- **Testability**:
  - See [docs/unit-testing-policy-migrate-comments.md](unit-testing-policy-migrate-comments.md) for the behavioral test outline and test artefact strategy.
  - Create test cases with various workbook versions and comment scenarios.
  - Validate all unresolved comments are migrated or reported as unmigratable.
  - Ensure no data loss or corruption in the new workbook.

## 7. Potential Blockers / Open Questions
1. How are OpenAPI entities currently mapped to Excel cells? Is there a stable anchor for matching?
2. Are there any remaining legacy (non-threaded) comments in use? 
  - **Answer:** No. Legacy (non-threaded) comments will not be supported or migrated. Only modern threaded comments are in scope.
3. Is there a requirement to preserve comment authorship and timestamps?
  - **Answer:** Yes. Preserving the stakeholder (author) and timestamp is important context when reading a comment and will be supported in the migration process.
4. Should the tool support partial/fuzzy matches, or only exact matches?
  - **Answer:** Only exact matches will be supported in the first iteration to keep the feature simple. A future enhancement request may add fuzzy/partial matching. As a backup, if an exact match is not found, the tool may attempt a case-insensitive and special-character-insensitive comparison. If no match is found, the comment will be considered unmigratable.

5. How should the user be notified of unmigratable comments?
  - **Answer:** Unmigratable comments will be added to a new worksheet named "Lost Commentary" in the generated workbook. This worksheet will list all comments that could not be mapped to an exact location.
  - **Note:** Comments of Type A (NoAnchor) and Type B (NoWorksheet) are considered successfully migrated when placed in their designated fallback locations, not "lost commentary."

### Requirements for the "Lost Commentary" Worksheet

Each row in the "Lost Commentary" worksheet should include:

- **Original Worksheet**: The name of the worksheet where the comment was originally found.
- **Original Cell Address**: The cell address (e.g., B12) in the old workbook.
- **OpenAPI Reference**: The OpenAPI anchor (if available) that the cell/comment was associated with.
- **Comment Thread**: The full text of the unresolved comment thread (multi-line if needed).
- **Comment Author**: The author of the comment (or main participants, if threaded).
- **Original Timestamp**: The timestamp of the comment or the last message in the thread.
- **Date Lost**: The date the comment was determined to be unmigratable (i.e., when the new workbook was generated).
- **Reason Lost**: A short code or message indicating why the comment could not be mapped (e.g., "No matching OpenAPI entity", "Endpoint removed", "Structure changed").

This structure ensures that all lost commentary is preserved for review, audit, or manual migration.
