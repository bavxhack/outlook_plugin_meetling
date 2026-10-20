using System;
using System.Collections.Generic;
using System.Globalization;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
namespace Meetling.Core;
public sealed class MeetlingApiClient
{
    private readonly HttpClient http;
    public MeetlingApiClient(HttpClient http) => this.http = http;
    public static HttpClient CreateHttpClient(MeetlingSettings settings) => new HttpClient(new HttpClientHandler { AllowAutoRedirect = false, UseCookies = false }) { Timeout = TimeSpan.FromSeconds(settings.TimeoutSeconds) };
    public static HttpRequestMessage BuildRequest(MeetlingSettings settings, string action, string token, string? serverId = null)
    {
        settings.Validate(); SettingsValidator.ValidateToken(token);
        bool create = action == "create";
        string endpoint = create ? settings.Endpoints.CreateConference : action == "logout" ? settings.Endpoints.Logout : action == "servers" ? settings.Endpoints.Servers : throw new ArgumentException("Unbekannte Aktion.");
        var request = new HttpRequestMessage(create ? HttpMethod.Post : HttpMethod.Get, settings.Endpoint(endpoint));
        request.Headers.TryAddWithoutValidation(create ? "token" : "Authorization", token);
        if (create)
        {
            if (string.IsNullOrWhiteSpace(serverId)) { request.Dispose(); throw new ArgumentException("Server-ID fehlt."); }
            request.Content = new FormUrlEncodedContent(new[] { new KeyValuePair<string,string>("server_id", serverId!) });
        }
        return request;
    }
    private async Task<JToken> Send(MeetlingSettings settings, string action, string token, string? id, CancellationToken ct)
    {
        using (var request = BuildRequest(settings, action, token, id))
        using (var response = await http.SendAsync(request, ct).ConfigureAwait(false))
        {
            if (!response.IsSuccessStatusCode) throw new MeetlingApiException("Meetling HTTP-Fehler " + (int)response.StatusCode + ". Weiterleitungen werden nicht verfolgt.", "HTTP " + (int)response.StatusCode);
            var body = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
            return Parse(body);
        }
    }
    public static JToken Parse(string body)
    {
        if (string.IsNullOrWhiteSpace(body)) throw new MeetlingApiException("Meetling hat eine leere Antwort geliefert.");
        if (body.TrimStart().StartsWith("<", StringComparison.Ordinal)) throw new MeetlingApiException("Meetling liefert HTML statt API-JSON. Bitte Anmeldung/Serverkonfiguration prüfen.");
        JToken json;
        try { json = JToken.Parse(body, new JsonLoadSettings { DuplicatePropertyNameHandling = DuplicatePropertyNameHandling.Error }); }
        catch (JsonException) { throw new MeetlingApiException("Meetling hat ungültiges JSON geliefert."); }
        if (json is JObject error && error["error"]?.Type == JTokenType.Boolean && (bool)error["error"]!)
        {
            var reason = error["reason"]?.Type == JTokenType.String ? (string?)error["reason"] : null;
            // Only documented reasons reach UI; unknown response content may contain secrets.
            var known = new[] { "No user loged in", "No user loged found", "User not found", "Server not found", "Server not allowed for user" };
            throw new MeetlingApiException(Array.IndexOf(known, reason) >= 0 ? "Meetling: " + reason : "Meetling meldet einen fachlichen Fehler.", reason);
        }
        return json;
    }
    private static string Required(JToken json, string key)
    {
        if (json[key]?.Type != JTokenType.String || string.IsNullOrWhiteSpace((string?)json[key])) throw new MeetlingApiException("Meetling-Antwort: Pflichtfeld fehlt oder ist ungültig.");
        return (string)json[key]!;
    }
    public static LoginResult ParseLogin(string body)
    {
        var json = Parse(body);
        if (!(json is JObject)) throw new MeetlingApiException("Ungültiges Login-Schema.");
        var token = Required(json, "token"); SettingsValidator.ValidateToken(token);
        return new LoginResult(token, Required(json, "name"));
    }
    public async Task<IReadOnlyList<Server>> GetServersAsync(MeetlingSettings settings, string token, CancellationToken ct)
        => ParseServers(await Send(settings, "servers", token, null, ct).ConfigureAwait(false));
    public static IReadOnlyList<Server> ParseServers(JToken json)
    {
        if (!(json is JArray array)) throw new MeetlingApiException("Ungültiges Serverlisten-Schema.");
        var ids = new HashSet<string>(StringComparer.Ordinal); var servers = new List<Server>();
        foreach (var item in array)
        {
            // Numeric integer IDs only. Json.NET preserves large integers (BigInteger).
            if (!(item is JObject) || item["id"]?.Type != JTokenType.Integer) throw new MeetlingApiException("Server-ID fehlt oder ist keine Ganzzahl.");
            string id = item["id"]!.ToString(Formatting.None);
            if (!ids.Add(id)) throw new MeetlingApiException("Doppelte Server-ID.");
            servers.Add(new Server(id, Required(item, "name")));
        }
        return servers;
    }
    public async Task<ConferenceResult> CreateConferenceAsync(MeetlingSettings settings, string token, string serverId, CancellationToken ct)
    {
        var json = await Send(settings, "create", token, serverId, ct).ConfigureAwait(false);
        if (!(json is JArray array) || array.Count != 1 || !(array[0] is JObject)) throw new MeetlingApiException("Erwartet wird genau eine Konferenz.");
        var url = Required(array[0], "url"); var title = Required(array[0], "title");
        if (!Uri.TryCreate(url, UriKind.Absolute, out var uri) || uri.Scheme != Uri.UriSchemeHttps || uri.UserInfo.Length != 0) throw new MeetlingApiException("Ungültiger HTTPS-Konferenzlink.");
        return new ConferenceResult(url, title);
    }
    public async Task LogoutAsync(MeetlingSettings settings, string token, CancellationToken ct)
    {
        var json = await Send(settings, "logout", token, null, ct).ConfigureAwait(false);
        if (!(json is JObject) || json["token"]?.Type != JTokenType.Null) throw new MeetlingApiException("Logout wurde nicht bestätigt.");
    }
}
