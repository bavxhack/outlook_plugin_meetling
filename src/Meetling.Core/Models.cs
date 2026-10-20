using System;
namespace Meetling.Core;
public sealed class MeetlingSettings
{
    public string BaseUrl { get; set; } = "";
    public string ApiBasePath { get; set; } = "api/v1/outlook/plugin/";
    public string LoginBasePath { get; set; } = "room/outlook/plugin/";
    public EndpointSettings Endpoints { get; set; } = new EndpointSettings();
    public int TimeoutSeconds { get; set; } = 30;
    [Newtonsoft.Json.JsonIgnore]
    public string Installation => SettingsValidator.NormalizeBaseUrl(BaseUrl);
    public Uri Endpoint(string relative) => Compose(ApiBasePath, relative);
    [Newtonsoft.Json.JsonIgnore]
    public Uri LoginEndpoint => Compose(LoginBasePath, Endpoints.Login);
    private Uri Compose(string prefix, string relative) => new Uri(Installation + SettingsValidator.RelativePath(prefix).TrimEnd('/') + "/" + SettingsValidator.RelativePath(relative));
    public void UpgradeKnownLegacyRoutes()
    {
        // Only the previous shipped prefix is recognized. Never rewrite custom paths
        // or the configuration file itself; tokens remain bound to the installation.
        if (ApiBasePath == "room/outlook/plugin/" || ApiBasePath == "room/outlook/plugin")
        {
            ApiBasePath = "api/v1/outlook/plugin/";
            if (LoginBasePath == "room/outlook/plugin/" && Endpoints != null && Endpoints.Login == "login") Endpoints.Login = "outlooklogin";
        }
    }
    public void Validate()
    {
        _ = Installation;
        if (TimeoutSeconds < 1 || TimeoutSeconds > 300 || Endpoints == null) throw new ArgumentException("Ungültige API-Konfiguration.");
        _ = LoginEndpoint;
        foreach (var endpoint in new[] { Endpoints.Servers, Endpoints.CreateConference, Endpoints.Logout }) _ = Endpoint(endpoint);
    }
}
public sealed class EndpointSettings
{
    public string Login { get; set; } = "outlooklogin";
    public string Servers { get; set; } = "server";
    public string CreateConference { get; set; } = "createConference";
    public string Logout { get; set; } = "logout";
}
public sealed class Server
{
    public Server(string id, string name) { Id = id; Name = name; }
    public string Id { get; }
    public string Name { get; }
    public override string ToString() => Name + " (ID: " + Id + ")";
}
public sealed class LoginResult
{
    public LoginResult(string token, string name) { Token = token; Name = name; }
    public string Token { get; }
    public string Name { get; }
}
public sealed class ConferenceResult
{
    public ConferenceResult(string url, string title) { Url = url; Title = title; }
    public string Url { get; }
    public string Title { get; }
}
public sealed class MeetlingApiException : Exception
{
    public string? Reason { get; }
    public bool InvalidToken => Reason == "No user loged in" || Reason == "No user loged found" || Reason == "User not found" || Reason == "HTTP 401" || Reason == "HTTP 403";
    public MeetlingApiException(string message, string? reason = null) : base(message) { Reason = reason; }
}
