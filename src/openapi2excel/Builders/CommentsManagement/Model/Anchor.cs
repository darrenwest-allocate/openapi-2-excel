namespace openapi2excel.core.Builders.CommentsManagement;

/// <summary>
/// A referenceable anchor used for mapping Excel cells to their OpenAPI source.
/// The anchor syntax is deterministic and unambiguous, allowing programmatic lookup in the OpenAPI document.
/// </summary>
public class Anchor(string value)
{
	 public string Value { get; } = value ?? string.Empty;
	public override string ToString() => Value;

	/// <summary>
	/// Creates a new anchor by appending a code segment to the existing anchor.
	/// </summary>
	internal Anchor With(string code) => new($"{Value}.{code}");
}



