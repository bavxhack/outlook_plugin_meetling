using System;
namespace Meetling.Core;
public static class BrowserNavigationPolicy
{
    // Empty popup documents are part of window.open; credentials still travel over HTTPS.
    public static bool IsAllowed(string candidate) => candidate == "about:blank" ||
        (Uri.TryCreate(candidate, UriKind.Absolute, out var uri) && uri.Scheme == Uri.UriSchemeHttps && uri.UserInfo.Length == 0);
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
