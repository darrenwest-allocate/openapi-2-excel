# Comment Migration Workflow

A Pull Request that provides an Excel Workbook produced with mappings to the OpenAPI content, and with comments that are unresolved, will, thanks to the workflow **'Migrate Comments from PR Workbook'** be used as the source of comments to be migrated to a new workbook for an updated version of the OpenAPI specification.

## How to Use

1. **Create a Pull Request** with your Excel workbook in this `sourceWorkbook/` directory
2. **Switch to your PR branch** in the GitHub Actions interface
3. **Trigger the GitHub Action workflow** called "Migrate Comments from PR Workbook" from the Actions tab
4. **Provide the OpenAPI specification URL** (the new/updated OpenAPI spec URL)
5. **Wait for completion** - the workflow will automatically comment on your PR with download instructions
6. **Download the migrated workbook** from the workflow artifacts

## Workflow Inputs

- **OpenAPI Spec URL**: Direct URL to the new OpenAPI specification (YAML or JSON format)

That's it! The workflow automatically:
- ✅ Detects the PR associated with your current branch
- ✅ Finds the first Excel workbook in the `sourceWorkbook/` directory  
- ✅ Lets the CLI tool auto-generate the output filename (based on OpenAPI info)
- ✅ Passes the OpenAPI URL directly to the CLI tool (no download needed)
- ✅ Outputs the migrated workbook to the `migration-output/` directory

## What Happens

1. The workflow finds the PR for your current branch
2. Locates Excel workbooks in the `sourceWorkbook/` directory  
3. Uses the CLI tool to migrate comments using the provided OpenAPI URL
4. Uploads the result as a workflow artifact with the CLI-generated filename
5. Comments on your PR with download instructions

## Example

If you upload `my-api-v1.xlsx`, the CLI will auto-generate an appropriate filename based on the OpenAPI specification info (e.g., `PetStore API v1.0.0.xlsx`) containing your migrated comments.

Place your Excel workbook here and follow the steps above to migrate your comments to a new workbook version!
