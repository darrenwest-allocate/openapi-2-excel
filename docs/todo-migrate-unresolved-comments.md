
# TODO List: Migrating Unresolved Excel Comments Feature

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
- [ ] generate the mappings from the open api content to the worksheet cell
- [x] Refactor as needed to ensure test passes and code is maintainable

## 3. Extracting Unresolved Comments
- [x] Create a failing unit test for extracting unresolved threaded comments from an old workbook
- [x] Implement extraction logic 
- [ ] Look for existing mapping files, collect them if they exist. annotate the extracted comments with the original open api anchor
- [ ] Refactor as needed

## 4. Mapping Comments to New Workbook
- [ ] Create a failing unit test for migrating comments to the correct cell in the new workbook (exact match)
- [ ] Implement migration logic using custom XML mapping
- [ ] Refactor as needed

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
