using System;
using System.Collections.Generic;
using System.Globalization;
using System.Net.Http;
using System.Security.Authentication;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
namespace Meetling.Core;
public sealed class MeetlingApiClient
{
    private readonly HttpClient http;
    private readonly Action<string>? diagnostic;
    public MeetlingApiClient(HttpClient http, Action<string>? diagnostic = null) { this.http = http; this.diagnostic = diagnostic; }
    public static HttpClientHandler CreateTransportHandler() => new HttpClientHandler
    {
        AllowAutoRedirect = false, UseCookies = false,
        // COM-hosted .NET can inherit pre-4.7 TLS defaults from Outlook.
        // Scope modern protocols to this client, leaving Outlook's globals alone.
        // TLS 1.3 = 12288 is the documented value absent from netstandard2.0's enum.
        SslProtocols = SslProtocols.Tls12 | (SslProtocols)12288
    };
    public static HttpClient CreateHttpClient(MeetlingSettings settings) => new HttpClient(CreateTransportHandler()) { Timeout = TimeSpan.FromSeconds(settings.TimeoutSeconds) };
    private void Trace(string message) { try { diagnostic?.Invoke(message); } catch (Exception) { /* Logging must not affect HTTP. */ } }
    public static HttpRequestMessage BuildRequest(MeetlingSettings settings, string action, string token, string? serverId = null)
    {
        settings.Validate(); SettingsValidator.ValidateToken(token);
        bool create = action == "create";
        string endpoint = create ? settings.Endpoints.CreateConference : action == "logout" ? settings.Endpoints.Logout : action == "servers" ? settings.Endpoints.Servers : throw new ArgumentException("Unbekannte Aktion.");
        var request = new HttpRequestMessage(create ? HttpMethod.Post : HttpMethod.Get, settings.Endpoint(endpoint));
        request.Headers.TryAddWithoutValidation("Authorization", "Bearer " + token);
        request.Headers.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
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
        {
            Trace("ApiRequest " + action + " start TLS-offered=1.2/1.3 redirects=off cookies=off");
            try
            {
                using (var response = await http.SendAsync(request, ct).ConfigureAwait(false))
                {
                    Trace("ApiRequest " + action + " HTTP=" + (int)response.StatusCode);
                    if (!response.IsSuccessStatusCode) throw new MeetlingApiException("Meetling HTTP-Fehler " + (int)response.StatusCode + ". Weiterleitungen werden nicht verfolgt.", "HTTP " + (int)response.StatusCode);
                    var body = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                    return Parse(body);
                }
            }
            catch (Exception ex) when (ex is HttpRequestException || ex is TaskCanceledException)
            {
                Trace("ApiRequest " + action + " failed " + TransportFailure.SafeDetails(ex));
                throw;
            }
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
    {
        var servers = ParseServers(await Send(settings, "servers", token, null, ct).ConfigureAwait(false));
        Trace("ApiRequest servers parsed count=" + servers.Count);
        return servers;
    }
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
        // The returned database ID is distinct from the participant URL, title and server ID.
        // Retain the successfully created room even if an older server omits the new ID field.
        var idValue = array[0]["id"];
        string? id = idValue?.Type == JTokenType.Integer ? idValue!.ToString(Formatting.None) : idValue?.Type == JTokenType.String ? (string?)idValue : null;
        if (id == null || id.Length == 0 || !System.Linq.Enumerable.All(id, c => c >= '0' && c <= '9') || !System.Linq.Enumerable.Any(id, c => c != '0')) id = null;
        string? editUrl = id == null ? null : settings.Installation + "room/dashboard?editConference_id=" + Uri.EscapeDataString(id);
        if (id == null) Trace("ApiRequest create missing-or-invalid-conference-id");
        // Phone data is optional: older or incomplete responses must not discard a created room.
        var phone = array[0]["phone"] as JObject;
        var numbers = (phone?["numbers"] as JObject)?["numbers"] as JObject;
        var phoneNumbers = new List<KeyValuePair<string, string>>();
        if (numbers != null)
            foreach (var number in numbers.Properties())
            {
                var value = number.Value.Type == JTokenType.String ? (string?)number.Value : null;
                if (!string.IsNullOrWhiteSpace(number.Name) && value != null && !string.IsNullOrWhiteSpace(value))
                    phoneNumbers.Add(new KeyValuePair<string, string>(number.Name, value.Trim()));
            }
        // Keep the PIN as a string so leading zeroes are retained.
        string? pin = phone?["dialIn"]?.Type == JTokenType.String ? ((string?)phone["dialIn"])?.Trim() : null;
        if (string.IsNullOrWhiteSpace(pin)) pin = null;
        return new ConferenceResult(url, title, id, editUrl, phoneNumbers, pin);
    }
    public async Task LogoutAsync(MeetlingSettings settings, string token, CancellationToken ct)
    {
        var json = await Send(settings, "logout", token, null, ct).ConfigureAwait(false);
        if (!(json is JObject) || json["token"]?.Type != JTokenType.Null) throw new MeetlingApiException("Logout wurde nicht bestätigt.");
    }
}
