using System.Text.RegularExpressions;
using Microsoft.OpenApi.Models;
using openapi2excel.core.Builders.WorksheetPartsBuilders.Common;

namespace openapi2excel.core.CustomXML;

public class Anchor 
{
	private string? _value;

	public string Value { get
		{
			return _value  ?? string.Empty;
		}
	}

	public Anchor(string value)
	{
		_value = value;
	}

	public override string ToString() => Value;

	internal Anchor With(string code)
	{
		return new Anchor($"{Value}.{code}");
	}
}

/// <summary>
/// Generates referenceable OpenAPI JSON anchors for mapping Excel cells to their OpenAPI source.
/// The anchor syntax is deterministic and unambiguous, allowing programmatic lookup in the OpenAPI document.
/// </summary>
public static partial class AnchorGenerator
{
	/// <summary>
	/// Generates an anchor for operation information (combines path and operation).
	/// Format: paths.{path}.{method}
	/// Example: paths./pets.get
	/// </summary>
	public static Anchor GenerateOperationInfoAnchor(string path, OpenApiPathItem pathItem, OperationType operationType, OpenApiOperation operation)
	{
		return GenerateOperationAnchor(path, operationType);
	}

	/// <summary>
	/// Generates an anchor for a path and operation.
	/// Format: paths.{path}.{method}
	/// Example: paths./pets.get
	/// </summary>
	public static Anchor GenerateOperationAnchor(string path, OperationType operationType)
	{
		var normalizedPath = NormalizePath(new Anchor(path));
		var method = operationType.ToString().ToLowerInvariant();
		return new Anchor($"paths.{normalizedPath}.{method}");
	}

	/// <summary>
	/// Generates an anchor for a response within an operation.
	/// Format: paths.{path}.{method}.responses.{status}
	/// Example: paths./pets.get.responses.200
	/// </summary>
	public static Anchor GenerateResponseAnchor(Anchor path_operationType, string statusCode)
	{
		return new Anchor($"{path_operationType}.responses.{statusCode}");
	}

	/// <summary>
	/// Generates an anchor for a parameter component.
	/// Format: components.parameters.{parameterName}
	/// Example: components.parameters.PetId
	/// </summary>
	public static Anchor GenerateParameterAnchor(string? parameterName)
	{
		return new Anchor($"components.parameters.{parameterName ?? string.Empty}");
	}

	/// <summary>
	/// Generates an anchor for a request body within an operation.
	/// Format: paths.{path}.{method}.requestBody
	/// Example: paths./pets.post.requestBody
	/// </summary>
	public static Anchor GenerateRequestBodyAnchor(Anchor path_operationType)
	{
		return new Anchor($"{path_operationType}.requestBody");
	}

	/// <summary>
	/// Normalizes a path for use in anchors by preserving the path structure.
	/// Paths are used as-is with slashes preserved.
	/// </summary>
	private static Anchor NormalizePath(Anchor path)
	{
		return path;
	}
	
	[GeneratedRegex(@"(&lt;)|(&gt;)|(&amp;)|(&quot;)|(&apos;)")]
	private static partial Regex MatchHtmlEntities();

	[GeneratedRegex(@"[^a-z0-9_]")]
	private static partial Regex MatchNonAlphaNumericUnderScore();
	
	[GeneratedRegex(@"\s+")]
	private static partial Regex MatchWhitespace();

	/// <summary>
	/// Appends a detail to an existing mapping anchor, formatting the detail as lowercase and replacing spaces with underscores.
	/// Format: {mappingAnchor}/@{detail}
	/// Example: paths./pets.get/@summary
	/// </summary>
	/// <param name="mappingAnchor">The base anchor to which the detail will be appended.</param>
	/// <param name="detail">The detail to append, which will be lowercased and have spaces replaced with underscores.</param>
	/// <returns>The combined anchor string with the appended detail.</returns>
	public static Anchor AppendDetailToAnchor(Anchor mappingAnchor, string detail)
	{
		detail = MatchWhitespace().Replace(detail.ToLowerInvariant(), "_");
		detail = MatchHtmlEntities().Replace(detail, string.Empty);
		detail = MatchNonAlphaNumericUnderScore().Replace(detail, string.Empty);
		return new Anchor($"{mappingAnchor}/@{Regex.Replace(detail.ToLowerInvariant(), @"\s+", "_")}");
	}

	internal static Anchor GenerateBodyFormatAnchor(Anchor mappingAnchor)
	{
		return new Anchor($"{mappingAnchor}/@{WorksheetLanguage.Request.BodyFormat.ToLowerInvariant().Replace(" ", string.Empty)}");
	}

	internal static Anchor GeneratePropertyAnchor(Anchor mappingAnchor, string propertyName)
	{
		propertyName = MatchWhitespace().Replace(propertyName.ToLowerInvariant(), "_");
		propertyName = MatchHtmlEntities().Replace(propertyName, string.Empty);
		propertyName = MatchNonAlphaNumericUnderScore().Replace(propertyName, string.Empty);
		return new Anchor($"{mappingAnchor}.{propertyName}");
	}

	internal static Anchor GenerateSchemaDescriptionHeaderAnchor(Anchor mappingAnchor)
	{
		return new Anchor($"{mappingAnchor}/SchemaDescriptionHeader");
	}

	internal static Anchor GenerateHeadersAnchor(Anchor anchor)
	{
		return new Anchor($"{anchor}.headers");
	}

	internal static Anchor GenerateResponseBodyAnchor(Anchor path_operationType)
	{
		return new Anchor($"{path_operationType}.responseBody");
	}

}



