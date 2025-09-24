
# TODO List: Migrating Unresolved Excel Comments Feature

## 🚀 Current Status (September 24, 2025)
**Phase Complete**: Core comment migration with threaded comment support ✅  
**Current Issue**: Comments migrate successfully but remain invisible in Excel ❌  
**Next Priority**: Implement legacy compatibility components (Phase 4.1) 🎯  

**Recent Achievements**:
- ✅ Enum-based error reporting for migration failures
- ✅ Persons part creation using hybrid OpenXML objects approach  
- ✅ Parent-child ID relationships preserved in threaded comments
- ✅ Excel repair dialog issues resolved (persons XML implemented)
- ✅ All tests passing (13/14 comments migrated, 1 expected failure)

**Blocking Issue**: Excel requires legacy compatibility components for comment visibility:
1. Legacy Comments XML (`/xl/comments{n}.xml`) - **HIGH PRIORITY**
2. VML Drawing Files (`/xl/drawings/vmlDrawing{n}.vml`) - **MEDIUM PRIORITY** 
3. Content Type Registration - **LOW PRIORITY**

---

## Start Here: How to Use This Checklist

Welcome! This checklist guides you through the step-by-step development of the unresolved comment migration feature. For each step:

- Review the [Implementation Plan](migrate-unresolved-comments-plan.md) for context and requirements.
- Reference the [Unit Testing Policy & Test Outline](unit-testing-policy-migrate-comments.md) for how to write and structure tests.
- See [Custom XML Metadata Mapping](custom-xml-metadata-mapping.md) for technical details on cell-to-OpenAPI mapping.
- Use the checklist below to track progress. For each major behavior:
	1. Create a failing unit test.
	2. Implement the code to make the test pass.
	3. Refactor as needed, ensuring all tests pass.

Update this file as new requirements or edge cases are discovered.

---

## 1. Setup & Preparation
- [x] Review and finalize the implementation plan and test policy docs
- [x] Prepare sample OpenAPI specs and Excel workbooks in `openapi2excel.tests/Sample/`

## 2. Custom XML Metadata Mapping
- [x] Create a failing unit test for writing and reading custom XML mapping parts per worksheet
- [x] Implement code to write/read custom XML mapping parts and meta part
- [x] generate the mappings from the open api content to the worksheet cell. cover every type of openapi output type
- [x] Refactor as needed to ensure test passes and code is maintainable

## 3. Extracting Unresolved Comments
- [x] Create a failing unit test for extracting unresolved threaded comments from an old workbook
- [x] Implement extraction logic 
- [x] Look for existing mapping files, collect them if they exist. annotate the extracted comments with the original open api anchor
- [ ] Refactor as needed

## 4. Mapping Comments to New Workbook
- [x] Create a failing unit test for migrating comments to the correct cell in the new workbook (exact match)
- [x] Implement migration logic using custom XML mapping
- [x] Implement enum-based error reporting for migration failures
- [x] Implement persons part creation using hybrid OpenXML objects approach
- [x] Ensure parent-child ID relationships are preserved in threaded comments
- [x] Resolve Excel repair dialog issues (persons XML missing)
- [ ] **CURRENT ISSUE**: Comments migrate successfully but are invisible in Excel
- [ ] Implement legacy compatibility components for comment visibility

## 4.1. Legacy Comment Compatibility (CRITICAL FOR VISIBILITY)

### **Root Cause**: Comments migrate without errors but remain invisible in Excel
Excel requires 3 additional components for threaded comment visibility:

#### **Phase 1: Legacy Comments XML** ⭐⭐⭐ (HIGH PRIORITY) ✅ **COMPLETED**
- [x] Create failing test for legacy comment creation (`/xl/comments{n}.xml`)
- [x] Implement `AddLegacyComment()` method using `WorksheetCommentsPart`
- [x] Handle author management with `Authors` collection and indices
- [x] Create legacy comments with "[Threaded comment]" text format
- [x] Register content type: `application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml`
- [x] Test: Comments become visible in Excel after legacy support

#### **Phase 2: VML Drawing Files** ⭐⭐ (MEDIUM PRIORITY)  
- [ ] Create failing test for VML drawing creation (`/xl/drawings/vmlDrawing{n}.vml`)
- [ ] Implement `AddVmlDrawing()` method using `VmlDrawingPart`
- [ ] Create comment shapes with proper positioning and anchoring
- [ ] Generate unique shape IDs across worksheets
- [ ] Register default content type: `vml` -> `application/vnd.openxmlformats-officedocument.vmlDrawing`
- [ ] Add legacy drawing references to worksheet parts
- [ ] Test: Comment indicators appear visually positioned on cells

#### **Phase 3: Content Type Registration** ⭐ (LOW PRIORITY)
- [ ] Create failing test for content type validation
- [ ] Implement `EnsureContentTypes()` method for proper registration
- [ ] Ensure all comment-related content types are registered
- [ ] Test: Generated files pass Excel's content type validation

#### **Implementation Strategy**
```csharp
// Enhanced AddThreadedCommentToWorksheet signature
private static void AddThreadedCommentToWorksheet(
    WorksheetPart worksheetPart, 
    string cellReference, 
    ThreadedCommentWithContext sourceComment,
    Dictionary<string, string> idMapping,
    string existingWorkbookPath)
{
    // 1. [EXISTING] Add threaded comment
    // ... existing threaded comment logic ...
    
    // 2. [NEW] Add legacy compatibility
    AddLegacyComment(worksheetPart, cellReference, sourceComment, newId);
    AddVmlDrawing(worksheetPart, cellReference, newId);
    EnsureContentTypes(worksheetPart.OpenXmlPackage as SpreadsheetDocument);
}
```

#### **Technical Notes**
- **Legacy Comments**: Use OpenXML `Comment`, `CommentList`, `Authors` objects
- **VML Drawings**: May require raw XML generation (similar to original persons approach)
- **Hybrid Extraction**: Copy VML/legacy structures from source workbook when available
- **Fallback Generation**: Create minimal VML when source lacks it
- **Author Management**: Map person IDs to author indices for legacy compatibility

#### **Success Criteria**
- ✅ Comments visible in Excel comment list
- ✅ Comment indicators appear on worksheet cells
- ✅ Threaded conversation structure preserved
- ✅ No Excel repair dialogs or validation errors
- ✅ Backward compatibility with older Excel versions

## 5. Handling Unmigratable Comments
- [ ] Create a failing unit test for unmigratable comments (no match)
- [ ] Implement logic to add these to the "Lost Commentary" worksheet with all required metadata
- [ ] Refactor as needed

## 6. Preserving Comment Metadata
- [ ] Create a failing unit test to ensure author and timestamp are preserved in migrated and lost comments
- [ ] Implement metadata preservation logic
- [ ] Refactor as needed

## 7. Case/Special Character Insensitivity Backup
- [ ] Create a failing unit test for backup matching (case/special char insensitivity)
- [ ] Implement backup matching logic
- [ ] Refactor as needed

## 8. Regression & Refactor Safety
- [ ] Ensure all tests pass after any refactor or major change
- [ ] Add regression tests as new behaviors or edge cases are discovered

## 9. Documentation & Review
- [ ] Update documentation to reflect implementation and usage
- [ ] Review all code, tests, and docs for completeness and clarity

---

This checklist ensures a disciplined, test-driven approach to delivering the feature, with clear steps for each major behavior and robust support for future refactoring.
