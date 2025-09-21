using Microsoft.OpenApi.Models;

namespace OpenApi2Excel.Core.CustomXml;

/// <summary>
/// Generates referenceable OpenAPI JSON anchors for mapping Excel cells to their OpenAPI source.
/// The anchor syntax is deterministic and unambiguous, allowing programmatic lookup in the OpenAPI document.
/// </summary>
public static class AnchorGenerator
{
    /// <summary>
    /// Generates an anchor for operation information (combines path and operation).
    /// Format: paths.{path}.{method}
    /// Example: paths./pets.get
    /// </summary>
    public static string GenerateOperationInfoAnchor(string path, OpenApiPathItem pathItem, OperationType operationType, OpenApiOperation operation)
    {
        return GenerateOperationAnchor(path, operationType);
    }

    /// <summary>
    /// Generates an anchor for a path and operation.
    /// Format: paths.{path}.{method}
    /// Example: paths./pets.get
    /// </summary>
    public static string GenerateOperationAnchor(string path, OperationType operationType)
    {
        var normalizedPath = NormalizePath(path);
        var method = operationType.ToString().ToLowerInvariant();
        return $"paths.{normalizedPath}.{method}";
    }

    /// <summary>
    /// Generates an anchor for a response within an operation.
    /// Format: paths.{path}.{method}.responses.{status}
    /// Example: paths./pets.get.responses.200
    /// </summary>
    public static string GenerateResponseAnchor(string path, OperationType operationType, string statusCode)
    {
        var operationAnchor = GenerateOperationAnchor(path, operationType);
        return $"{operationAnchor}.responses.{statusCode}";
    }

    /// <summary>
    /// Generates an anchor for a parameter component.
    /// Format: components.parameters.{parameterName}
    /// Example: components.parameters.PetId
    /// </summary>
    public static string GenerateParameterAnchor(string parameterName)
    {
        return $"components.parameters.{parameterName}";
    }

    /// <summary>
    /// Generates an anchor for a schema component.
    /// Format: components.schemas.{schemaName}
    /// Example: components.schemas.Pet
    /// </summary>
    public static string GenerateSchemaAnchor(string schemaName)
    {
        return $"components.schemas.{schemaName}";
    }

    /// <summary>
    /// Generates an anchor for a property within a schema.
    /// Format: components.schemas.{schemaName}.properties.{propertyName}
    /// Example: components.schemas.Pet.properties.name
    /// </summary>
    public static string GenerateSchemaPropertyAnchor(string schemaName, string propertyName)
    {
        return $"components.schemas.{schemaName}.properties.{propertyName}";
    }

    /// <summary>
    /// Generates an anchor for a request body within an operation.
    /// Format: paths.{path}.{method}.requestBody
    /// Example: paths./pets.post.requestBody
    /// </summary>
    public static string GenerateRequestBodyAnchor(string path, OperationType operationType)
    {
        var operationAnchor = GenerateOperationAnchor(path, operationType);
        return $"{operationAnchor}.requestBody";
    }

    /// <summary>
    /// Generates an anchor for a parameter within an operation.
    /// Format: paths.{path}.{method}.parameters.{index}
    /// Example: paths./pets.get.parameters.0
    /// </summary>
    public static string GenerateOperationParameterAnchor(string path, OperationType operationType, int parameterIndex)
    {
        var operationAnchor = GenerateOperationAnchor(path, operationType);
        return $"{operationAnchor}.parameters.{parameterIndex}";
    }

    /// <summary>
    /// Generates an anchor for a security scheme component.
    /// Format: components.securitySchemes.{schemeName}
    /// Example: components.securitySchemes.ApiKeyAuth
    /// </summary>
    public static string GenerateSecuritySchemeAnchor(string schemeName)
    {
        return $"components.securitySchemes.{schemeName}";
    }

    /// <summary>
    /// Normalizes a path for use in anchors by preserving the path structure.
    /// Paths are used as-is with slashes preserved.
    /// </summary>
    private static string NormalizePath(string path)
    {
        // Return path as-is, preserving slashes and other characters
        // This matches the recommendation in the documentation
        return path;
    }

	public static string AppendDetailToAnchor(string mappingAnchor, string detail)
	{
		return $"{mappingAnchor}/@{detail.ToLowerInvariant().Replace(" ", "_")}";
	}
}



