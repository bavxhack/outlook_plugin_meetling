using System;
namespace Meetling.Core;
public static class BrowserNavigationPolicy
{
    // Empty popup documents are part of window.open; credentials still travel over HTTPS.
    public static bool IsAllowed(string candidate) => candidate == "about:blank" ||
        (Uri.TryCreate(candidate, UriKind.Absolute, out var uri) && uri.Scheme == Uri.UriSchemeHttps && uri.UserInfo.Length == 0);
    // Recover only an HTTP redirect to the exact configured login, using its known
    // HTTPS URL. Never follow HTTP or copy callback parameters into a new request.
    public static bool CanReturnViaHttps(Uri expected, string candidate)
        => expected.Scheme == Uri.UriSchemeHttps && expected.IsDefaultPort &&
           Uri.TryCreate(candidate, UriKind.Absolute, out var uri) && uri.Scheme == Uri.UriSchemeHttp && uri.IsDefaultPort &&
           uri.Host.Equals(expected.Host, StringComparison.OrdinalIgnoreCase) && uri.AbsolutePath == expected.AbsolutePath &&
           uri.UserInfo.Length == 0 && uri.Query.Length == 0 && uri.Fragment.Length == 0;
    public static string BlockReason(string candidate)
    {
        if (!Uri.TryCreate(candidate, UriKind.Absolute, out var uri)) return "invalid-uri";
        if (uri.UserInfo.Length != 0) return "userinfo";
        switch (uri.Scheme)
        {
            case "http": return "http";
            case "about": return "about-other";
            case "blob": return "blob";
            case "data": return "data";
            case "file": return "file";
            case "javascript": return "javascript";
            default: return "other-scheme";
        }
    }
    // Only fixed categories reach diagnostics; never expose SSO URLs, codes or session IDs.
    public static string Classify(Uri expected, string candidate)
    {
        if (SettingsValidator.IsExpectedLogin(expected, candidate)) return "plugin-login";
        if (!IsAllowed(candidate)) return "blocked";
        if (candidate == "about:blank") return "empty-document";
        var uri = new Uri(candidate);
        return uri.GetLeftPart(UriPartial.Authority) == expected.GetLeftPart(UriPartial.Authority) ? "installation" : "identity-provider";
    }
}
