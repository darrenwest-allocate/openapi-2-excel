# OpenAPI-2-Excel Copilot Instructions

## Repository Overview

**OpenAPI-2-Excel** is a .NET 9.0 tool that converts OpenAPI/Swagger specifications (YAML or JSON) into human-readable Microsoft Excel documentation. The tool generates Excel files with an information worksheet and detailed operation worksheets for REST API documentation that's accessible to Business Analysts and developers.

### Project Structure
- **Language**: C# (.NET 9.0)
- **Project Type**: Multi-project solution with CLI tool, core library, and unit tests
- **Dependencies**: ClosedXML (Excel generation), Microsoft.OpenApi (OpenAPI parsing), Spectre.Console (CLI), xUnit (testing)
- **Size**: ~50 source files across 3 projects

## Build & Development Instructions

### Prerequisites
- **.NET 9.0 SDK** (verified working with 9.0.305)
- No additional runtime or build tools required

### Build Commands
**ALWAYS run commands from the repository root directory.**

#### 1. Restore Dependencies
```powershell
dotnet restore
```
**Duration**: ~1 second. **Always run this first** before any build operations.

#### 2. Build Solution
```powershell
# Debug build (default)
dotnet build

# Release build (recommended for validation)
dotnet build --configuration Release
```
**Duration**: ~2-4 seconds. Release build is required for accurate performance testing.

#### 3. Run Tests
```powershell
# Debug tests
dotnet test

# Release tests (recommended)
dotnet test --configuration Release
```
**Duration**: ~3-4 seconds. 
All tests must pass for CI validation.

#### 4. Run CLI Tool
```powershell
# Show help
dotnet run --project src/openapi2excel.cli -- --help

# Convert OpenAPI to Excel (example)
dotnet run --project src/openapi2excel.cli -- input.yaml output.xlsx --no-logo

# With configuration options
dotnet run --project src/openapi2excel.cli --configuration Release -- input.yaml output.xlsx --depth 5 --no-logo
```

### Validation Pipeline
The CI/CD runs on GitHub Actions (`.github/workflows/test.yml`):
1. **Platforms**: Ubuntu, Windows, macOS
2. **Framework**: net9.0
3. **Steps**: 
   - Smoke test: Converts sample YAML to Excel
   - Unit tests: All tests must pass
4. **Timeout**: 30 minutes

**To replicate CI locally:**
```powershell
# Smoke test (replicates CI)
dotnet run --configuration Release --framework net9.0 --project src/openapi2excel.cli src/openapi2excel.tests/Sample/Sample1.yaml Sample.xlsx

# Full test suite
dotnet test --configuration Release --framework net9.0

# create a workbook with migrated comments
dotnet run --project src/openapi2excel.cli --configuration Release -- "src/openapi2excel.tests/Sample/sample-api-gw.json" "test-output/debug-workbook.xlsx" --existing-workbook "src/openapi2excel.tests/Sample/sample-api-gw-with-mappings.xlsx" --no-logo
```

## Project Architecture

### Solution Structure
```
openapi2excel.sln                    # Main solution file
├── src/openapi2excel/               # Core library (openapi2excel.core.csproj)
│   ├── OpenApiDocumentationGenerator.cs  # Main entry point
│   ├── OpenApiDocumentationOptions.cs    # Configuration options
│   ├── Builders/                         # Excel worksheet builders
│   │   ├── InfoWorksheetsBuilder.cs     # Info worksheet generation
│   │   ├── OperationWorksheetBuilder.cs # Operation worksheet generation
│   │   ├── WorksheetBuilder.cs          # Base worksheet builder
│   │   ├── CommentsManagement/          # Excel comments migration system
│   │   │   ├── AnchorGenerator.cs       # Comment anchor generation
│   │   │   ├── CommentMigrationHelper.cs # Main migration logic
│   │   │   ├── CommentTargetResolver.cs # Target cell resolution
│   │   │   ├── MigrationStrategy/       # Migration strategies
│   │   │   ├── Model/                   # Comment data models
│   │   │   └── OpenXml/                 # OpenXML integration
│   │   └── WorksheetPartsBuilders/      # Worksheet component builders
│   │       ├── Common/                  # Shared builder utilities
│   │       ├── HomePageLinkBuilder.cs   # Home page navigation
│   │       ├── OperationInfoBuilder.cs  # Operation information
│   │       ├── PropertiesTreeBuilder.cs # Schema properties
│   │       ├── RequestBodyBuilder.cs    # Request body documentation
│   │       ├── RequestParametersBuilder.cs # Request parameters
│   │       ├── ResponseBodyBuilder.cs   # Response body documentation
│   │       └── WorksheetPartBuilder.cs  # Base part builder
│   ├── Common/                          # Utility extensions
│   │   ├── EnumerableExtensions.cs     # Collection utilities
│   │   ├── ExcelOpenXmlHelper.cs       # Excel/OpenXML helpers
│   │   ├── OpenApiInfoExtractor.cs     # OpenAPI metadata extraction
│   │   ├── OpenApiSchemaExtension.cs   # Schema processing extensions
│   │   ├── RowPointer.cs               # Excel row navigation
│   │   ├── StringExtensions.cs         # String utilities
│   │   └── XLExtensions.cs             # ClosedXML extensions
│   └── Core/                           # Core domain models (empty)
├── src/openapi2excel.cli/          # CLI tool (openapi2excel.cli.csproj)
│   ├── Program.cs                   # CLI entry point
│   ├── GenerateExcelCommand.cs      # Command implementation
│   └── CustomHelpProvider.cs       # Custom help formatting
└── src/openapi2excel.tests/        # Unit tests (openapi2excel.tests.csproj)
    ├── OpenApiDocumentationGeneratorTest.cs # Core functionality tests
    ├── ExcelCommentsTest.cs         # Comments migration tests
    ├── Investigation/               # Experimental test directory
    ├── Sample/                      # Test data files
    │   ├── Sample1.yaml             # Basic Swagger Petstore sample
    │   ├── sample-api-gw.json       # Complex API Gateway sample
    │   ├── sample-api-gw.xlsx       # Generated Excel baseline
    │   ├── sample-api-gw-with-mappings.xlsx # Excel with existing comments
    │   └── sample-api-gw-workbook-first-mappings.xml # Comment mapping data
    └── xunit.runner.json           # Test runner configuration
```

### Key Files & Dependencies
- **Main Generator**: `src/openapi2excel/OpenApiDocumentationGenerator.cs` - Core conversion logic
- **CLI Interface**: `src/openapi2excel.cli/Program.cs` - Uses Spectre.Console.Cli framework
- **Version**: `semver.txt` - Contains current version (0.1.9)
- **Test Sample**: `src/openapi2excel.tests/Sample/Sample1.yaml` - Swagger Petstore example

### Package Dependencies
- **ClosedXML** (0.105.0): Excel file generation
- **OpenXML** ?????????????????
- **Microsoft.OpenApi** (1.6.25): OpenAPI specification parsing  
- **Spectre.Console** (0.51.1): CLI framework with rich formatting
- **xUnit** (2.9.3): Testing framework

## Common Development Patterns

### Code Style
- **C# 12** with latest language features enabled
- **Nullable reference types** enabled across all projects
- **Implicit usings** enabled
- Follow standard .NET naming conventions

### Known Technical Debt
Current TODOs in codebase (search for "TODO"):
- Language helper refactoring needed in `OpenApiDocumentationOptions.cs`
- Complex example support needed in `OpenApiSchemaExtension.cs`
- Complex default value support needed in `OpenApiSchemaExtension.cs`

### Testing
- **Framework**: xUnit with Visual Studio test adapter
- **Test Data**: (basic functionality) Uses Swagger Petstore sample specification; (comments management) sample-api-gw.json sample-api-gw-with-mappings.xlsx
- **Expected Behavior**: All tests must pass for CI validation
- **Examine Workbook** You can copy an xlsx to a zip and unpack it to examine the content to check the links and patterns of the xml files
- **Coverage**: Provided by Behavioral feature requirements when creating new features
- **Feature Development Approach**: For each major behavior:
	1. Create a failing unit test.
	2. Implement the code to make the test pass.
	3. Refactor as needed, ensuring all tests pass.
- **Parallelism**: deactivated for test running because of a static variable holding all the OpenAPi mapping that must be cleared down with each new test, stopping the testing being isolated.

## Troubleshooting & Validation

### Build Issues
- **Issue**: Restore failures → **Solution**: Ensure .NET 9.0 SDK installed
- **Issue**: Missing project references → **Solution**: Run `dotnet restore` first
- **Issue**: Version conflicts → **Solution**: Delete `bin/` and `obj/` folders, then restore

### Runtime Validation
```powershell
# Quick validation test
dotnet run --project src/openapi2excel.cli --configuration Release -- src/openapi2excel.tests/Sample/Sample1.yaml test-output/validation.xlsx --no-logo
```
Expected output: "Excel file saved to [path]" with no errors.

### File Locations for Quick Reference
- **Projects**: All in `src/` directory
- **Documentation**: `docs/` directory contains technical documentation along with TODO lists
- **Test Output**: Generated Excel files go to `test-output/` directory
- **sourceWorkbook/**: where workflows can find workbooks to use a source for comments to be migrated to new workbooks
- **test-output/**: where tests should create files for review and assert outcomes

### Examine the content of the XLSX archive
Some hard to explain outcomes where the workbook fails to open without error, or does not show the expected output can be investigated by taking a copy of the .xlsx, renaming it to .zip, and unzipping to a folder for its source XML documents.

These can be compared against the archived files of a  working workbook like `Sample/sample-api-gw-with-mappings.xls`.

- **Example Command**: `Copy-Item "test-output\problematic-workbook.xlsx" "test-output\problematic-workbook.zip" ; Expand-Archive -Path "test-output\problematic-workbook.zip" -DestinationPath "test-output\problematic-workbook-unpacked" -Force`


## Development Planning Process

### The `docs/` Folder: Feature Planning & Tracking

The `docs/` directory contains **planning documentation** for new features. **ALWAYS check this folder first** when working on feature branches. Each feature follows a structured planning approach:

#### Required Documentation Files
1. **Implementation Plan** (`*-plan.md`): Detailed technical specification
   - Requirements breakdown
   - Risk assessment and edge cases
   - Architecture decisions
   - API/library choices
2. **TODO Checklist** (`todo-*.md`): Step-by-step development tracking
   - Test-driven development steps
   - Progress checkboxes for each feature component
   - Links to related documentation


#### Conservative Development Process

**CRITICAL RULE**: Follow the plans conservatively. **Never deviate** from the documented approach without updating the planning documents first.

1. **Start with the TODO list**: Use `todo-*.md` as your primary guide
2. **Test-Driven Development**: For each checkbox:
   - Create a failing unit test first
   - Implement only the code needed to pass the test
   - Refactor while keeping tests green
   - Check off the TODO item only when tests pass
3. **No unauthorized embellishments**: Stick to planned features only
4. **Update plans for changes**: If requirements change, update the planning docs first


#### Working with TODO Lists
- **Check items off sequentially** - don't skip ahead
- **Each checkbox represents a working unit test** - no implementation without tests
- **Track progress visibly** - keep the checklist updated
- **Never mark complete without all tests passing** and confirmation that the workbook opens without error and displays content

This process ensures controlled, predictable development and prevents scope creep.

## Agent Guidelines

**Trust these instructions** - they are comprehensive and tested. Only search for additional information if:
1. Instructions are incomplete for your specific task
2. You encounter errors not covered here
3. You need to understand specific code implementation details

**Always validate changes** by running the full test suite (`dotnet test --configuration Release --logger console --verbosity normal`) before completing any modification.

**For new features**: Add corresponding tests in `src/openapi2excel.tests/` and ensure they integrate with the existing xUnit test structure.

Experimentation and investigation of problems with new tests should have that test added to a folder `Investigation/` within the tests project to avoid interference.

With feature completion, on summary of success, always invite the user to examine the workbook manually before concluding the task. Not all problems are seen in the unit tests.