using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Meetling.Core;
namespace Meetling.OutlookAddIn;
internal sealed class PluginController : IDisposable
{
    private readonly SettingsStore settingsStore = new SettingsStore();
    private readonly TokenStore tokens = new TokenStore();
    private readonly UiDispatcher ui = new UiDispatcher();
    private readonly SessionEpoch epoch = new SessionEpoch();
    private CancellationTokenSource lifetime = new CancellationTokenSource();
    private ConfigurationSnapshot configuration = new ConfigurationSnapshot(null, false, "Konfiguration fehlt.");
    private LoginResult? login;
    private IReadOnlyList<Server>? servers;
    private DateTime serversLoaded;
    private readonly OperationGate operations = new OperationGate();
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
            var result = await BrowserLogin(settings.Endpoint(settings.Endpoints.Login), ct);
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
        using (var http = MeetlingApiClient.CreateHttpClient(settings)) return await new MeetlingApiClient(http).GetServersAsync(settings, token, ct);
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
            using (var http = MeetlingApiClient.CreateHttpClient(settings)) result = await new MeetlingApiClient(http).CreateConferenceAsync(settings, token, server.Id, ct);
            await ui.InvokeAsync(() =>
            {
                if (!epoch.IsCurrent(version)) { if (!disposed) ShowLink(result.Url); return; }
                try { target.Apply(result); SetStatus("Angemeldet als " + login!.Name); }
                catch (Exception) { ShowLink(result.Url); }
            });
        }
        catch (Exception ex)
        {
            if (result != null) await ui.InvokeAsync(() => { if (!disposed) ShowLink(result.Url); });
            else await Error(version, ex, requested ? "Erstellung fehlgeschlagen oder Ausgang unklar. Der Raum kann bereits existieren. Kein automatischer Wiederholungsversuch." : "Konferenz konnte nicht vorbereitet werden.");
        }
        finally { await ui.InvokeAsync(() => { target?.Dispose(); if (epoch.IsCurrent(version)) { busy = false; if (Status == "Server werden geladen." || Status == "Konferenz wird erstellt.") SetStatus(login == null ? "Nicht angemeldet." : "Angemeldet als " + login.Name); Changed?.Invoke(); } }); }
    }
    private static void ShowLink(string url)
    {
        using (var form = new Form { Text = "Konferenz erstellt – Link manuell übernehmen", Width = 700, Height = 160 })
        {
            form.Controls.Add(new TextBox { Text = url, ReadOnly = true, Dock = DockStyle.Top });
            form.Controls.Add(new Label { Text = "Outlook-Einfügung fehlgeschlagen. Bitte diesen Link manuell übernehmen. Es wird kein weiterer Raum erstellt.", Dock = DockStyle.Bottom, Height = 60 }); form.ShowDialog();
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
            using (var http = MeetlingApiClient.CreateHttpClient(settings)) await new MeetlingApiClient(http).LogoutAsync(settings, token, lifetime.Token);
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
        LocalDialog.Show(context + (ex is MeetlingApiException || ex is InvalidOperationException ? "\n" + ex.Message : ""));
    });
    public void Dispose() { disposed = true; Invalidate(); lifetime.Cancel(); /* dispatcher remains until queued continuations drain; no Office calls survive the epoch. */ }
}
