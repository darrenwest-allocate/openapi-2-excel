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
```

## Project Architecture

### Solution Structure
```
openapi2excel.sln                    # Main solution file
├── src/openapi2excel/               # Core library (openapi2excel.core.csproj)
│   ├── OpenApiDocumentationGenerator.cs  # Main entry point
│   ├── OpenApiDocumentationOptions.cs    # Configuration options
│   ├── Builders/                         # Excel worksheet builders
│   └── Common/                           # Utility extensions
├── src/openapi2excel.cli/          # CLI tool (openapi2excel.cli.csproj)
│   ├── Program.cs                   # CLI entry point
│   └── GenerateExcelCommand.cs      # Command implementation
└── src/openapi2excel.tests/        # Unit tests (openapi2excel.tests.csproj)
    ├── OpenApiDocumentationGeneratorTest.cs
    └── Sample/Sample1.yaml          # Test OpenAPI specification
```

### Key Files & Dependencies
- **Main Generator**: `src/openapi2excel/OpenApiDocumentationGenerator.cs` - Core conversion logic
- **CLI Interface**: `src/openapi2excel.cli/Program.cs` - Uses Spectre.Console.Cli framework
- **Version**: `semver.txt` - Contains current version (0.1.9)
- **Test Sample**: `src/openapi2excel.tests/Sample/Sample1.yaml` - Swagger Petstore example

### Package Dependencies
- **ClosedXML** (0.105.0): Excel file generation
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
- **Coverage**: Basic smoke test coverage (1 test currently)
- **Test Data**: Uses Swagger Petstore sample specification
- **Expected Behavior**: All tests must pass for CI validation

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
- **Documentation**: `docs/` directory contains migration plans and technical documentation
- **Test Output**: Generated Excel files go to `test-output/` directory
- **Published Scripts**: `PublishScripts/` contains example Excel outputs

## Development Planning Process

### The `docs/` Folder: Feature Planning & Tracking

The `docs/` directory contains **mandatory planning documentation** for new features. **ALWAYS check this folder first** when working on feature branches. Each feature follows a structured planning approach:

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
3. **Testing Policy** (`unit-testing-policy-*.md`): Test strategy
   - Behavioral test requirements
   - Test structure and organization
   - Expected test artifacts

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

#### Example: Current Feature Branch
The current feature branch follows this pattern with:
- `docs/migrate-unresolved-comments-plan.md` - Complete technical specification
- `docs/todo-migrate-unresolved-comments.md` - Step-by-step TODO checklist
- `docs/unit-testing-policy-migrate-comments.md` - Testing strategy

#### Working with TODO Lists
- **Check items off sequentially** - don't skip ahead
- **Each checkbox represents a working unit test** - no implementation without tests
- **Track progress visibly** - keep the checklist updated
- **Never mark complete without all tests passing**

This process ensures controlled, predictable development and prevents scope creep.

## Agent Guidelines

**Trust these instructions** - they are comprehensive and tested. Only search for additional information if:
1. Instructions are incomplete for your specific task
2. You encounter errors not covered here
3. You need to understand specific code implementation details

**Always validate changes** by running the full test suite (`dotnet test --configuration Release`) before completing any modification.

**For new features**: Add corresponding tests in `src/openapi2excel.tests/` and ensure they integrate with the existing xUnit test structure.
