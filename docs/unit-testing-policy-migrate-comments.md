# Unit Testing Policy and Behavioural Test Outline: Migrating Unresolved Comments

## Unit Testing Policy

- **Behavior-Driven Focus:**
  - Write tests that validate observable behaviors and outcomes, not just internal implementation details.
  - Tests should describe “what” the system does, not “how” it does it.

- **Refactor-Friendly:**
  - Tests should be resilient to internal refactoring. As long as the behavior/output is unchanged, tests should continue to pass.

- **Representative Coverage:**
  - Prioritize meaningful scenarios and edge cases over exhaustive line coverage.
  - Aim for high confidence in correctness, not 100% code coverage.

- **Test Data as Artefacts:**
  - Use small but realistic sample OpenAPI JSON and .xlsx files as test artefacts.
  - Store these in a dedicated test data directory (e.g., `openapi2excel.tests/Sample/`).

- **Regression Safety:**
  - All new features and bug fixes must be accompanied by or covered by at least one test.
  - Tests should fail if a breaking change is introduced.

---

## Test Outline for Migrating Unresolved Comments

### Test Data Location
- Place sample OpenAPI specs and Excel files in `openapi2excel.tests/Sample/`
  - e.g., `Sample1.yaml`, `OldWorkbook1.xlsx`, `ExpectedNewWorkbook1.xlsx`

### Test Scenarios

1. **Extract Unresolved Comments from Old Workbook**
   - Given: An old workbook with threaded comments (some resolved, some unresolved)
   - Expect: Only unresolved comments are extracted, with correct metadata (author, timestamp, etc.)

2. **Map Comments to New Workbook (Exact Match)**
   - Given: An old workbook and a new workbook with matching OpenAPI anchors
   - Expect: Comments are migrated to the correct cells in the new workbook

3. **Handle Unmigratable Comments**
   - Given: Comments in the old workbook with no matching anchor in the new workbook
   - Expect: These comments appear in the “Lost Commentary” worksheet with all required metadata

4. **Preserve Comment Metadata**
   - Given: Comments with author and timestamp
   - Expect: These fields are preserved in the new workbook or “Lost Commentary” sheet

5. **Case/Special Character Insensitivity Backup Check**
   - Given: Minor differences in anchor case or special characters between old and new workbooks
   - Expect: Comments are still matched if possible

6. **Meta Custom XML Mapping**
   - Given: Workbooks with custom XML mapping parts
   - Expect: The mapping is correctly read and written, and used for migration

7. **Regression: Refactor Safety**
   - Given: Refactoring of migration logic
   - Expect: All above tests continue to pass, ensuring no behavioral regressions

---


## Test OpenAPI Specification

For unit and integration tests related to the migration of unresolved Excel comments, by default use the minimal OpenAPI specification at:

```
src/openapi2excel.tests/Sample/sample-api-gw.json
```

This file contains a reduced set of paths and operations (GET, PUT, POST) for both `/{tenant}/duties/` and `/{tenant}/duty-assignments` endpoints, and should be referenced in all new test cases for this feature.


## Example Test Artefacts

- `Sample/sample-api-gw.json`: Minimal OpenAPI spec for focused testing
- `Sample/sample-api-gw.xlsx`: Vanilla Excel workbook generated from the minimal spec (use as the basis for comment and custom XML migration tests; make copies as needed)
- `Sample/Sample1.yaml`: The OpenAPI spec used to generate the workbooks

---

## Summary Table

| Test Scenario                        | Test Data Files                  | Expected Outcome                        |
|--------------------------------------|----------------------------------|-----------------------------------------|
| Extract unresolved comments          | OldWorkbook1.xlsx                | Only unresolved comments extracted      |
| Migrate comments (exact match)       | OldWorkbook1.xlsx, Sample1.yaml  | Comments appear in correct cells        |
| Handle unmigratable comments         | OldWorkbook1.xlsx, Sample1.yaml  | “Lost Commentary” sheet populated       |
| Preserve comment metadata            | OldWorkbook1.xlsx                | Author/timestamp preserved              |
| Case/special char insensitivity      | OldWorkbook1.xlsx, Sample1.yaml  | Comments matched if possible            |
| Meta custom XML mapping              | OldWorkbook1.xlsx, NewWorkbook1.xlsx | Mapping read/written and used        |
| Regression/refactor safety           | All above                        | Tests pass after refactor               |

---

This policy and outline will guide the development and maintenance of robust, behavior-driven tests for the unresolved comment migration feature.
