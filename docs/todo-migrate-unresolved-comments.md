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
- [x] Implement legacy compatibility components for comment visibility
- [x] Implement threaded comments  ✅ **BREAKTHROUGH ACHIEVED!**


#### **Success Criteria**
- ✅ Comments visible in Excel comment list
- ✅ Comment indicators appear on worksheet cells
- ✅ Threaded conversation structure preserved
- ✅ No Excel repair dialogs or validation errors. (Currently only detectable by human execution)
- ✅ Backward compatibility with older Excel versions

## 5. Handling non-migratable Comments

### 5A. Type A - "NoAnchor" Comments (on existing worksheets, no OpenAPI anchor)
- [x] Create a failing unit test for "NoAnchor" comments that should migrate to heading rows
- [x] Implement logic to find nearest anchor ending with `/TitleRow` above the comment location  
- [x] Implement column preservation from original location
- [x] Implement collision handling (place in row below heading if target cell occupied)
- [x] Extend ID mapping dictionary to track Type A comment migrations
- [x] Ensure full comment thread migration (Legacy + ThreadedComment approach)
- [x] Refactor as needed  ✅ **TYPE A MIGRATION COMPLETE!**

### 5B. Type B - "NoWorksheet" Comments (on worksheets not in new workbook)  
- [x] Create a failing unit test for "NoWorksheet" comments that should migrate to Info sheet
- [x] Implement logic to identify Info sheet using `OpenApiDocumentationLanguageConst.Info`
- [x] Implement placement in column V starting at row 1, stacking downward
- [x] Extend ID mapping dictionary to track Type B comment migrations  
- [x] Ensure full comment thread migration with original worksheet context
- [x] Refactor as needed  ✅ **TYPE B MIGRATION COMPLETE!**

### 5C. Integration with Existing Migration Flow
- [x] Extend `CommentMigrationFailureReason` enum with new success categories (or remove from "failure" tracking)
- [x] Update main `MigrateComments` method to handle Type A and Type B cases  
- [x] Ensure Type A/B processing doesn't interfere with existing exact-match migrations
- [x] Update tests to account for additional successful migrations (adjust expected counts)
- [x] Refactor as needed  ✅ **INTEGRATION COMPLETE!**


## 6 NEW. Github action that will allow the upload of the old workbook so that its comments can be migrated across
- [x] decide options
- [ ] implement

## 9. Documentation & Review
- [ ] Update documentation to reflect implementation and usage
- [ ] Review all code, tests, and docs for completeness and clarity

---

This checklist ensures a disciplined, test-driven approach to delivering the feature, with clear steps for each major behavior and robust support for future refactoring.
