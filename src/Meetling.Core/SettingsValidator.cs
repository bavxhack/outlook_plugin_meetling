using System;
using System.Net.Mail;

namespace Meetling.Core;

public static class SettingsValidator
{
    public static string NormalizeBaseUrl(string value)
    {
        var normalized = (value ?? string.Empty).Trim().TrimEnd('/');
        if (!Uri.TryCreate(normalized, UriKind.Absolute, out var uri) ||
            (uri.Scheme != Uri.UriSchemeHttps && uri.Scheme != Uri.UriSchemeHttp) ||
            !string.IsNullOrEmpty(uri.Query) || !string.IsNullOrEmpty(uri.Fragment))
            throw new ArgumentException("Die Meetling-Basis-URL ist ungültig.", nameof(value));
        return normalized;
    }

    public static bool IsInsecureRemoteUrl(string value)
    {
        var uri = new Uri(NormalizeBaseUrl(value));
        return uri.Scheme == Uri.UriSchemeHttp &&
               !uri.IsLoopback &&
               !uri.Host.Equals("localhost", StringComparison.OrdinalIgnoreCase);
    }

    public static bool IsValidEmail(string? value)
    {
        if (value is null)
        {
            return false;
        }

        var normalizedValue = value.Trim();
        if (normalizedValue.Length == 0)
        {
            return false;
        }

        try
        {
            return new MailAddress(normalizedValue).Address.Equals(normalizedValue, StringComparison.OrdinalIgnoreCase);
        }
        catch (FormatException)
        {
            return false;
        }
    }
}
