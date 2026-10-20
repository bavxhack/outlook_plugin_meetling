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
        if (id == null || id.Length …2992 tokens truncated…n    private readonly OperationGate operations = new OperationGate();
    private bool busy { get => operations.Busy; set { if (value) operations.TryEnter(); else operations.Release(); } }
    private bool signingIn, disposed;
    internal string Status { get; private set; } = "Nicht angemeldet.";
    internal bool CanCreate => !disposed && !busy && !signingIn && configuration.Settings != null && login != null;
    internal bool CanLogin => !disposed && !busy && !signingIn && configuration.Settings != null;
    internal bool CanLogout => !disposed && !busy && (login != null || signingIn);
    internal Action? Changed;
    internal PluginController() { Reload(true); }
    private void SetStatus(string text) { Status = text; Changed?.Invoke(); }
    private void Invalidate()
    {
        epoch.Advance(); lifetime.Cancel(); lifetime.Dispose(); lifetime = new CancellationTokenSource();
        login = null; servers = null; signingIn = false; busy = false;
    }
    internal void Reload(bool initial = false)
    {
        Invalidate(); configuration = settingsStore.Load();
        if (configuration.Settings == null) { SetStatus(configuration.Error ?? "Konfiguration fehlt."); return; }
        if (initial)
        {
            try { login = tokens.Load(configuration.Settings.Installation); }
            catch (Exception) { SetStatus("Gespeicherte Anmeldung nicht lesbar. Bitte erneut anmelden."); return; }
        }
        SetStatus(login == null ? "Nicht angemeldet." : "Angemeldet als " + login.Name + " (noch nicht geprüft)");
    }
    internal void Settings()
    {
        if (busy) return;
        using (var form = new SettingsForm(configuration))
        {
            var result = form.ShowDialog();
            if (result == DialogResult.OK) settingsStore.Save(form.Settings);
            if (result == DialogResult.OK || result == DialogResult.Retry) Reload();
        }
    }
    [MethodImpl(MethodImplOptions.NoInlining)]
    private static Task<LoginResult?> BrowserLogin(Uri endpoint, CancellationToken ct) => LoginWindow.SignInAsync(endpoint, ct);
    internal async Task SignIn()
    {
        if (!CanLogin) return;
        // Explicit user action only. Never called during COM load/startup.
        Invalidate(); signingIn = true; var version = epoch.Value; var ct = lifetime.Token; var settings = configuration.Settings!;
        SetStatus("Anmeldung läuft.");
        try
        {
            var result = await BrowserLogin(settings.LoginEndpoint, ct);
            await ui.InvokeAsync(() =>
            {
                if (disposed || !epoch.IsCurrent(version) || ct.IsCancellationRequested) return;
                if (result == null) { SetStatus("Nicht angemeldet."); return; }
                tokens.Save(settings.Installation, result); // Commit to UI only after durable DPAPI save.
                login = result; servers = null; SetStatus("Angemeldet als " + result.Name);
            });
        }
        catch (Exception ex) { await Error(version, ex, "Anmeldung konnte nicht sicher gespeichert oder geöffnet werden."); }
        finally { await ui.InvokeAsync(() => { if (epoch.IsCurrent(version)) { signingIn = false; Changed?.Invoke(); } }); }
    }
    private async Task<IReadOnlyList<Server>> LoadServers(MeetlingSettings settings, string token, CancellationToken ct, bool refresh)
    {
        if (!refresh && servers != null && DateTime.UtcNow - serversLoaded < TimeSpan.FromMinutes(2)) return servers;
        using (var http = MeetlingApiClient.CreateHttpClient(settings)) return await new MeetlingApiClient(http, DiagnosticLog.WriteInfo).GetServersAsync(settings, token, ct);
    }
    internal async Task RefreshServers()
    {
        if (!CanCreate) return;
        busy = true; int version = epoch.Value; var ct = lifetime.Token;
        SetStatus("Server werden geladen.");
        try
        {
            var list = await LoadServers(configuration.Settings!, login!.Token, ct, true);
            await ui.InvokeAsync(() => { if (!epoch.IsCurrent(version)) return; servers = list; serversLoaded = DateTime.UtcNow; SetStatus(list.Count == 0 ? "Keine Server verfügbar." : "Angemeldet als " + login!.Name + " – " + list.Count + " Server"); });
        }
        catch (Exception ex) { await Error(version, ex, "Server konnten nicht geladen werden."); }
        finally { await ui.InvokeAsync(() => { if (epoch.IsCurrent(version)) { busy = false; Changed?.Invoke(); } }); }
    }
    internal async Task Create(object? application, object control)
    {
        if (!CanCreate) { LocalDialog.Show(Status + " Bitte Einstellungen prüfen und anmelden."); return; }
        busy = true; int version = epoch.Value; var ct = lifetime.Token;
        var settings = configuration.Settings!; string token = login!.Token;
        AppointmentTarget? target = null; ConferenceResult? result = null; bool requested = false;
        try
        {
            target = AppointmentTarget.Capture(application, control); SetStatus("Server werden geladen.");
            var list = await LoadServers(settings, token, ct, false);
            Server? server = null;
            await ui.InvokeAsync(() =>
            {
                if (!epoch.IsCurrent(version)) return;
                servers = list; serversLoaded = DateTime.UtcNow; server = ServerSelection.Choose(list);
                if (server != null && epoch.IsCurrent(version) && !ct.IsCancellationRequested) { target.MarkPending(); requested = true; SetStatus("Konferenz wird erstellt."); }
            });
            if (server == null || !epoch.IsCurrent(version)) return;
            using (var http = MeetlingApiClient.CreateHttpClient(settings)) result = await new MeetlingApiClient(http, DiagnosticLog.WriteInfo).CreateConferenceAsync(settings, token, server.Id, ct);
            await ui.InvokeAsync(() =>
            {
                if (!epoch.IsCurrent(version)) { if (!disposed) ShowLink(result); return; }
                try
                {
                    target.Apply(result); SetStatus("Angemeldet als " + login!.Name);
                    if (result.ModeratorEditUrl == null) LocalDialog.Show("Die Konferenz wurde eingefügt. Der Server hat keine gültige Konferenz-ID geliefert; der Bearbeitungslink konnte deshalb nicht ergänzt werden. Es wurde kein weiterer Raum erstellt.");
                }
                catch (Exception) { ShowLink(result); }
            });
        }
        catch (Exception ex)
        {
            if (result != null) await ui.InvokeAsync(() => { if (!disposed) ShowLink(result); });
            else await Error(version, ex, requested ? "Erstellung fehlgeschlagen oder Ausgang unklar. Der Raum kann bereits existieren. Kein automatischer Wiederholungsversuch." : "Konferenz konnte nicht vorbereitet werden.");
        }
        finally { await ui.InvokeAsync(() => { target?.Dispose(); if (epoch.IsCurrent(version)) { busy = false; if (Status == "Server werden geladen." || Status == "Konferenz wird erstellt.") SetStatus(login == null ? "Nicht angemeldet." : "Angemeldet als " + login.Name); Changed?.Invoke(); } }); }
    }
    private static void ShowLink(ConferenceResult conference)
    {
        using (var form = new Form { Text = "Konferenz erstellt – Links manuell übernehmen", Width = 760, Height = 300 })
        {
            form.Controls.Add(new TextBox { Text = conference.BodyText.Replace("\r", Environment.NewLine), ReadOnly = true, Multiline = true, ScrollBars = ScrollBars.Vertical, Dock = DockStyle.Fill });
            form.Controls.Add(new Label { Text = "Outlook-Einfügung fehlgeschlagen. Bitte diese Links manuell übernehmen. Es wird kein weiterer Raum erstellt.", Dock = DockStyle.Bottom, Height = 60 }); form.ShowDialog();
        }
    }
    internal async Task Logout()
    {
        if (busy) return;
        if (signingIn) { Invalidate(); SetStatus("Nicht angemeldet."); return; }
        if (login == null || configuration.Settings == null) return;
        if (MessageBox.Show("Token für alle Clients mit demselben Wert widerrufen? Die Symfony-Web-Sitzung wird dadurch nicht beendet.", "Meetling – Abmelden", MessageBoxButtons.OKCancel) != DialogResult.OK) return;
        var settings = configuration.Settings; string token = login.Token;
        // Invalidate login and outstanding responses immediately, retain token for retry on failure.
        epoch.Advance(); lifetime.Cancel(); lifetime.Dispose(); lifetime = new CancellationTokenSource(); servers = null;
        int version = epoch.Value; busy = true; SetStatus("Abmeldung läuft.");
        try
        {
            using (var http = MeetlingApiClient.CreateHttpClient(settings)) await new MeetlingApiClient(http, DiagnosticLog.WriteInfo).LogoutAsync(settings, token, lifetime.Token);
            await ui.InvokeAsync(() => { if (epoch.IsCurrent(version)) ClearLocal(settings.Installation, "Nicht angemeldet. Token serverseitig widerrufen."); });
        }
        catch (MeetlingApiException ex) when (ex.Reason == "No user loged found") { await ui.InvokeAsync(() => { if (epoch.IsCurrent(version)) ClearLocal(settings.Installation, "Nicht angemeldet. Token war nicht zugeordnet; kein neuer Widerruf bestätigt."); }); }
        catch (Exception)
        {
            await ui.InvokeAsync(() =>
            {
                if (!epoch.IsCurrent(version)) return;
                SetStatus("Serverseitiger Widerruf nicht bestätigt. Abmelden kann erneut versucht werden.");
                if (MessageBox.Show("Widerruf nicht bestätigt. Nur lokal abmelden? Mit Nein bleibt das Token für einen erneuten Versuch erhalten.", "Meetling", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    ClearLocal(settings.Installation, "Nur lokal abgemeldet. Kein serverseitiger Widerruf bestätigt.");
            });
        }
        finally { await ui.InvokeAsync(() => { if (epoch.IsCurrent(version)) { busy = false; Changed?.Invoke(); } }); }
    }
    private void ClearLocal(string installation, string status)
    {
        bool deleted = true;
        try { tokens.Delete(installation); } catch (Exception ex) { deleted = false; DiagnosticLog.Write("Lokale Tokendatei konnte nicht gelöscht werden", ex); }
        Invalidate();
        SetStatus(deleted ? status : status + " Achtung: Die lokale Tokendatei konnte nicht gelöscht werden; nur der Arbeitsspeicher wurde bereinigt.");
        if (!deleted) LocalDialog.Show("Die Serverantwort wurde verarbeitet. Die lokale Tokendatei konnte nicht gelöscht werden. Bitte vor dem nächsten Outlook-Start den Zugriff auf den Tokenstore prüfen.");
    }
    private Task Error(int version, Exception ex, string context) => ui.InvokeAsync(() =>
    {
        if (disposed || !epoch.IsCurrent(version)) return;
        if (ex is MeetlingApiException api && api.InvalidToken)
        {
            login = null; servers = null;
            try { if (configuration.Settings != null) tokens.Delete(configuration.Settings.Installation); } catch (Exception) { }
            SetStatus("Anmeldung nicht mehr gültig. Bitte erneut anmelden.");
        }
        if (!(ex is MeetlingApiException apiError && apiError.InvalidToken)) SetStatus(context);
        DiagnosticLog.Write(context, ex);
        LocalDialog.Show(context + (ex is HttpRequestException || ex is TaskCanceledException ? "\n" + TransportFailure.UserMessage(ex) : ex is MeetlingApiException || ex is InvalidOperationException ? "\n" + ex.Message : ""));
    });
    public void Dispose() { disposed = true; Invalidate(); lifetime.Cancel(); /* dispatcher remains until queued continuations drain; no Office calls survive the epoch. */ }
}
