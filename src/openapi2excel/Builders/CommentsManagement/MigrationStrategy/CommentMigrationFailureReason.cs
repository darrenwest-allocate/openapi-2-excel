namespace openapi2excel.core.Builders.CommentsManagement.MigrationStrategy;

/// <summary>
/// Reasons why a threaded comment migration might fail.
/// </summary>
public enum CommentMigrationState
{
    /// <summary>
    /// The comment has no associated OpenAPI anchor, making it impossible to map to the new workbook.
    /// </summary>
    NoOpenApiAnchorFound,

    /// <summary>
    /// The OpenAPI anchor exists but cannot be found in the new workbook's cell mappings.
    /// </summary>
    OpenApiAnchorNotFoundInNewWorkbook,

    /// <summary>
    /// The target worksheet referenced in the mapping does not exist in the new workbook.
    /// </summary>
    TargetWorksheetNotFound,

    /// <summary>
    /// An unexpected error occurred during the migration process (e.g., file I/O, XML parsing).
    /// </summary>
    UnexpectedErrorDuringMigration,
    
    /// <summary>
    /// Comment successfully migrated with no OpenAPI anchor - placed near title row on existing worksheet.
    /// </summary>
    SuccessfullyMigratedAsNoAnchorComment,
    
    /// <summary>
    /// Comment successfully migrated from missing worksheet - placed on Info sheet.
    /// </summary>
    SuccessfullyMigratedAsNoWorksheetComment,

    Successful,
    
    Unknown
}
