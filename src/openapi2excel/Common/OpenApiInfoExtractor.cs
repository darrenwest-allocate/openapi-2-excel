using Microsoft.OpenApi.Readers;
using System.Text.RegularExpressions;

namespace openapi2excel.core.Common;

public static class OpenApiInfoExtractor
{
    /// <summary>
    /// Extracts title and version information from an OpenAPI document file.
    /// </summary>
    /// <param name="openApiFilePath">Path to the OpenAPI document file (JSON or YAML)</param>
    /// <returns>A tuple containing the title and version, or default values if extraction fails</returns>
    public static async Task<(string title, string version)> ExtractInfoAsync(string openApiFilePath)
    {
        try
        {
            if (!File.Exists(openApiFilePath))
                return ("api", "v1");

            await using var fileStream = File.OpenRead(openApiFilePath);
            return await ExtractInfoAsync(fileStream);
        }
        catch
        {
            return ("api", "v1");
        }
    }

    /// <summary>
    /// Extracts title and version information from an OpenAPI document stream.
    /// </summary>
    /// <param name="openApiStream">Stream containing the OpenAPI document</param>
    /// <returns>A tuple containing the title and version, or default values if extraction fails</returns>
    public static async Task<(string title, string version)> ExtractInfoAsync(Stream openApiStream)
    {
        try
        {
            var readResult = await new OpenApiStreamReader().ReadAsync(openApiStream);

            if (readResult?.OpenApiDocument?.Info == null)
                return ("api", "v1");

            var title = readResult.OpenApiDocument.Info.Title ?? "api";
            var version = readResult.OpenApiDocument.Info.Version ?? "v1";

            return (title, version);
        }
        catch
        {
            return ("api", "v1");
        }
    }

    /// <summary>
    /// Generates a filename in the format {title}_{version}.xlsx with proper sanitization.
    /// </summary>
    /// <param name="title">The API title from OpenAPI info</param>
    /// <param name="version">The API version from OpenAPI info</param>
    /// <returns>A sanitized filename suitable for the filesystem</returns>
    public static string GenerateFilename(string title, string version)
    {
        var sanitizedTitle = SanitizeFilename(title);
        var sanitizedVersion = SanitizeFilename(version).Replace('.', '-');

        // Handle edge cases
        if (string.IsNullOrWhiteSpace(sanitizedTitle))
            sanitizedTitle = "api";
        if (string.IsNullOrWhiteSpace(sanitizedVersion))
            sanitizedVersion = "v1";

        var filename = $"{sanitizedTitle}_{sanitizedVersion}.xlsx";

        // Ensure the filename isn't too long &&   room for directory path
        if (filename.Length > 200)
        {
            var maxTitleLength = 100;
            var maxVersionLength = 50;
            if (sanitizedTitle.Length > maxTitleLength)
                sanitizedTitle = sanitizedTitle.Substring(0, maxTitleLength);
            if (sanitizedVersion.Length > maxVersionLength)
                sanitizedVersion = sanitizedVersion.Substring(0, maxVersionLength);
            filename = $"{sanitizedTitle}_{sanitizedVersion}.xlsx";
        }
        return filename;
    }

    /// <summary>
    /// Sanitizes a string to be safe for use in filenames by removing or replacing invalid characters.
    /// </summary>
    /// <param name="input">The input string to sanitize</param>
    /// <returns>A sanitized string safe for filenames</returns>
    public static string SanitizeFilename(string input)
    {
        if (string.IsNullOrWhiteSpace(input))
            return string.Empty;
        var sanitized = input.Trim();
        foreach (var invalidChar in Path.GetInvalidFileNameChars())
        {
            sanitized = sanitized.Replace(invalidChar, '_');
        }
        // Replace multiple consecutive spaces and underscores with single underscore
        sanitized = Regex.Replace(sanitized, @"[\s_]+", "_");
        // Remove leading/trailing underscores
        sanitized = sanitized.Trim('_');
        // Handle reserved Windows filenames
        var reservedNames = new[] { "CON", "PRN", "AUX", "NUL", "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9", "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9" };
        if (reservedNames.Contains(sanitized.ToUpperInvariant()))
        {
            sanitized = sanitized + "_file";
        }
        return sanitized;
    }
}
