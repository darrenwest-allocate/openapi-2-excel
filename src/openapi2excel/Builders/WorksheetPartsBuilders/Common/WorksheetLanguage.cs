
namespace openapi2excel.core.Builders.WorksheetPartsBuilders.Common;

public static class WorksheetLanguage
{
	public static class Generic
	{
		public const string Name = "Name";
		public const string TitleRow = "/TitleRow";
	}

	public static class Schema
	{
		public const string Title = "SCHEMA";
		public const string ObjectType = "Object type";
		public const string Format = "Format";
		public const string Length = "Length";
		public const string Required = "Required";
		public const string Nullable = "Nullable";
		public const string Range = "Range";
		public const string Pattern = "Pattern";
		public const string Enum = "Enum";
		public const string Deprecated = "Deprecated";
		public const string Default = "Default";
		public const string Example = "Example";
		public const string Description = "Description";
	}

	public static class Response
	{
		public const string Title = "RESPONSE";
		public const string ResponseHeaders = "Response headers";
		public const string HeadersName = "Name";
		
	}

	public static class Request
	{
		public const string Title = "REQUEST";
		public const string BodyFormat = "Body format";
	}

	public static class Operations
	{
		public const string Title = "OPERATION INFORMATION";
		public const string OperationType = "Operation type";
		public const string Id = "Id";
		public const string Path = "Path";
		public const string PathDescription = "Path description";
		public const string PathSummary = "Path summary";
		public const string OperationDescription = "Operation description";
		public const string OperationSummary = "Operation summary";
		public const string Deprecated = "Deprecated";
	}

	public static class Parameters
	{
		public const string Title = "PARAMETERS";
		public const string Name = "Name";
		public const string Location = "Location";
		public const string Serialization = "Serialization";
		public const string Type = "Type";
		public const string ObjectType = "Object type";
		public const string Format = "Format";
		public const string Length = "Length";
		public const string Required = "Required";
		public const string Nullable = "Nullable";
		public const string Range = "Range";
		public const string Pattern = "Pattern";
		public const string Enum = "Enum";
		public const string Deprecated = "Deprecated";
		public const string Default = "Default";
		public const string Example = "Example";
		public const string Description = "Description";
	}
}
