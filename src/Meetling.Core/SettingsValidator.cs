using System;
namespace Meetling.Core;
public static class SettingsValidator
{
    public static string NormalizeBaseUrl(string value)
    {
        if (!Uri.TryCreate((value ?? "").Trim(), UriKind.Absolute, out var uri) || uri.Scheme != Uri.UriSchemeHttps ||
            !string.IsNullOrEmpty(uri.UserInfo) || !string.IsNullOrEmpty(uri.Query) || !string.IsNullOrEmpty(uri.Fragment))
            throw new ArgumentException("Bitte eine absolute HTTPS-Installationsadresse ohne Zugangsdaten, Query oder Fragment eingeben.");
        return uri.AbsoluteUri.TrimEnd('/') + "/";
    }
    public static string RelativePath(string value)
    {
        if (string.IsNullOrWhiteSpace(value) || value.StartsWith("/", StringComparison.Ordinal) || value.Contains("\\") || value.Contains(":") ||
            value.Contains("?") || value.Contains("#") || value.Contains("%") || Array.Exists(value.Split('/'), s => s == "." || s == ".."))
            throw new ArgumentException("API-Pfade müssen relative Pfade ohne führenden Slash oder Traversierung sein.");
        return value;
    }
    public static void ValidateToken(string token)
    {
        if (string.IsNullOrWhiteSpace(token) || Array.Exists(token.ToCharArray(), char.IsControl))
            throw new ArgumentException("Das Anmeldetoken ist leer oder ungültig.");
    }
    public static bool IsExpectedLogin(Uri expected, string candidate)
        => Uri.TryCreate(candidate, UriKind.Absolute, out var actual) && actual.Scheme == Uri.UriSchemeHttps &&
           actual.GetLeftPart(UriPartial.Authority) == expected.GetLeftPart(UriPartial.Authority) &&
           actual.AbsolutePath == expected.AbsolutePath && actual.Query.Length == 0 && actual.Fragment.Length == 0 && actual.UserInfo.Length == 0;
}
