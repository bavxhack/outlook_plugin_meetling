using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Meetling.Core;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
namespace Meetling.OutlookAddIn;
// WebView2 is loaded only when login is needed, after COM initialization.
internal sealed class LoginWindow : Form
{
    private readonly WebView2 browser = new WebView2 { Dock = DockStyle.Fill };
    private readonly Uri expected;
    private readonly List<Form> popups = new List<Form>();
    private readonly Label status = new Label { Dock = DockStyle.Fill, UseMnemonic = false, Text = "Browser wird gestartet …" };
    private readonly Button returnToMeetling = new Button { Dock = DockStyle.Right, Width = 195, Text = "Zurück zu Meetling", Enabled = false };
    private int navigations, responseErrors;
    private bool httpsReturnUsed;
    private readonly UiDispatcher dispatcher = new UiDispatcher();
    private readonly CancellationToken ct;
    private readonly TaskCompletionSource<LoginResult?> completion = new TaskCompletionSource<LoginResult?>();
    private readonly System.Windows.Forms.Timer timeout = new System.Windows.Forms.Timer { Interval = 300000 };
    private bool finished, reading;
    private LoginWindow(Uri expected, CancellationToken ct)
    {
        this.expected = expected; this.ct = ct;
        Text = "Meetling – Anmelden"; Width = 900; Height = 720; Controls.Add(browser);
        var footer = new Panel { Dock = DockStyle.Bottom, Height = 70, Padding = new Padding(10) };
        footer.Controls.Add(status); footer.Controls.Add(returnToMeetling); Controls.Add(footer); footer.BringToFront();
        returnToMeetling.Click += (s,e) =>
        {
            if (finished || reading || ct.IsCancellationRequested || browser.CoreWebView2 == null) return;
            Trace("manual-return"); SetStatus("Meetling-Anmeldung wird geprüft …");
            browser.CoreWebView2.Navigate(expected.AbsoluteUri);
        };
        FormClosed += (s,e) => { finished = true; timeout.Stop(); completion.TrySetResult(null); };
        timeout.Tick += (s,e) => Fail("Die Anmeldung wurde nach fünf Minuten beendet.");
        Shown += async (s,e) =>
        {
            timeout.Start();
            try { await Initialize(); }
            catch (Exception) { if (!finished) Fail("Browser-Anmeldung nicht verfügbar. Bitte Microsoft Edge WebView2 Evergreen Runtime installieren oder die Installation prüfen."); }
        };
    }
    [MethodImpl(MethodImplOptions.NoInlining)]
    internal static async Task<LoginResult?> SignInAsync(Uri endpoint, CancellationToken ct)
    {
        using (var form = new LoginWindow(endpoint, ct))
        using (ct.Register(() => { try { if (form.IsHandleCreated) form.BeginInvoke((Action)(() => form.Close())); } catch (InvalidOperationException) { } }))
        {
            form.Show();
            return await form.completion.Task;
        }
    }
    private async Task Initialize()
    {
        var environment = await CoreWebView2Environment.CreateAsync(null, Path.Combine(SettingsStore.Root, "WebView2", Guid.NewGuid().ToString("N")));
        if (finished || ct.IsCancellationRequested) return;
        await browser.EnsureCoreWebView2Async(environment);
        if (finished || ct.IsCancellationRequested) return;
        Attach(browser);
        returnToMeetling.Enabled = true;
        Trace("started");
        browser.CoreWebView2.Navigate(expected.AbsoluteUri);
    }
    private void Attach(WebView2 view)
    {
        var core = view.CoreWebView2;
        core.Settings.AreDevToolsEnabled = false;
        core.Settings.AreDefaultContextMenusEnabled = false;
        core.NewWindowRequested += NewWindow;
        core.NavigationStarting += (s,e) =>
        {
            if (finished || ct.IsCancellationRequested) { e.Cancel = true; return; }
            string stage = BrowserNavigationPolicy.Classify(expected, e.Uri);
            e.Cancel = !BrowserNavigationPolicy.IsAllowed(e.Uri);
            if (++navigations <= 200) Trace("navigation " + stage + (e.IsRedirected ? " redirected" : ""));
            if (e.Cancel)
            {
                Trace("navigation-blocked reason=" + BrowserNavigationPolicy.BlockReason(e.Uri));
                if (!httpsReturnUsed && e.IsRedirected && BrowserNavigationPolicy.CanReturnViaHttps(expected, e.Uri))
                {
                    httpsReturnUsed = true;
                    Trace("return-to-configured-https-endpoint");
                    SetStatus("Rückleitung zum konfigurierten HTTPS-Anmeldeendpunkt …");
                    view.BeginInvoke((Action)(() => { if (!finished && !ct.IsCancellationRequested && !view.IsDisposed) view.CoreWebView2.Navigate(expected.AbsoluteUri); }));
                    return;
                }
            }
            SetStatus(e.Cancel ? "Eine nicht unterstützte Weiterleitung wurde blockiert. Bitte Diagnose prüfen." :
                stage == "plugin-login" ? "Meetling-Antwort wird erwartet …" : "SSO-Anmeldung läuft. Nach erfolgreicher MFA bei Bedarf „Zurück zu Meetling“ wählen.");
        };
        core.NavigationCompleted += (s,e) =>
        {
            if (finished || ct.IsCancellationRequested) return;
            if (!e.IsSuccess)
            {
                Trace("navigation-failed " + e.WebErrorStatus);
                SetStatus("Anmeldeseite konnte nicht geladen werden (" + e.WebErrorStatus + "). Bitte Diagnose prüfen.");
            }
            else if (navigations <= 200) Trace("navigation-completed");
        };
        core.DownloadStarting += (s,e) =>
        {
            // A token response must never become a plaintext file on disk.
            e.Cancel = true; e.Handled = true;
            if (!finished) { Trace("download-blocked"); SetStatus("Der Browser wollte eine Datei herunterladen. Anmeldung bleibt geöffnet; bitte Diagnose prüfen."); }
        };
        core.ProcessFailed += (s,e) => { if (!finished) { Trace("browser-process-failed " + e.ProcessFailedKind); Fail("Der Anmeldebrowser wurde unerwartet beendet. Bitte erneut anmelden."); } };
        core.WebResourceResponseReceived += Response;
    }
    private async void NewWindow(object? sender, CoreWebView2NewWindowRequestedEventArgs e)
    {
        // Assign NewWindow through the documented deferral, preserving opener and SSO session.
        e.Handled = true;
        if (finished || ct.IsCancellationRequested) return;
        if (!BrowserNavigationPolicy.IsAllowed(e.Uri)) { Trace("popup-blocked"); SetStatus("Ein SSO-Fenster mit nicht unterstützter Adresse wurde blockiert."); return; }
        if (popups.Count >= 5) { Trace("popup-limit"); SetStatus("Zu viele SSO-Fenster. Bitte unbenutzte Fenster schließen."); return; }
        var deferral = e.GetDeferral();
        Form? popup = null;
        try
        {
            var view = new WebView2 { Dock = DockStyle.Fill };
            popup = new Form { Text = "Meetling – SSO-Anmeldung", Width = 800, Height = 650 };
            popup.Controls.Add(view); popups.Add(popup);
            var window = popup;
            popup.FormClosed += (s,args) => { popups.Remove(window); window.Dispose(); };
            popup.Show(this);
            await view.EnsureCoreWebView2Async(((CoreWebView2)sender!).Environment);
            if (finished || ct.IsCancellationRequested || popup.IsDisposed) return;
            Attach(view);
            view.CoreWebView2.WindowCloseRequested += (s,args) => { if (!window.IsDisposed) window.Close(); };
            e.NewWindow = view.CoreWebView2;
            Trace("popup-opened");
        }
        catch (Exception ex)
        {
            DiagnosticLog.Write("BrowserLogin popup initialization failed", ex);
            if (popup != null && !popup.IsDisposed) popup.Close();
            if (!finished) SetStatus("Das SSO-Fenster konnte nicht geöffnet werden. Bitte Diagnose prüfen.");
        }
        finally { deferral.Complete(); }
    }
    private void SetStatus(string text) { if (!finished && !IsDisposed) status.Text = text; }
    private static void Trace(string stage) => DiagnosticLog.WriteInfo("BrowserLogin " + stage);
    private async void Response(object? sender, CoreWebView2WebResourceResponseReceivedEventArgs e)
    {
        if (finished || ct.IsCancellationRequested) return;
        if (e.Response.StatusCode >= 400 && ++responseErrors <= 30)
            Trace("http-error " + BrowserNavigationPolicy.Classify(expected, e.Request.Uri) + " HTTP=" + e.Response.StatusCode);
        if (reading || e.Request.Method != "GET" || !SettingsValidator.IsExpectedLogin(expected, e.Request.Uri)) return;
        Trace("endpoint-response HTTP=" + e.Response.StatusCode);
        if (e.Response.StatusCode >= 300 && e.Response.StatusCode < 400) return; // SSO redirects are allowed in browser only.
        reading = true;
        string phase = "response-headers";
        try
        {
            if (e.Response.StatusCode != 200) throw new MeetlingApiException("Der Login-Endpunkt meldet einen HTTP-Fehler.");
            if (!e.Response.Headers.Contains("Content-Type"))
            {
                Trace("endpoint-missing-content-type");
                SetStatus("Die Login-Antwort enthält keinen Inhaltstyp. Bitte Diagnose prüfen.");
                return;
            }
            var contentType = e.Response.Headers.GetHeader("Content-Type");
            if (contentType.IndexOf("application/json", StringComparison.OrdinalIgnoreCase) < 0)
            {
                Trace("endpoint-html-or-non-json");
                SetStatus("Der Endpunkt zeigt eine Webseite. Anmeldung abschließen, danach bei Bedarf „Zurück zu Meetling“ wählen.");
                return;
            }
            phase = "response-stream";
            var stream = await e.Response.GetContentAsync();
            phase = "response-read-and-validate";
            var login = await BrowserLoginResponse.ReadAsync(stream, ct);
            await dispatcher.InvokeAsync(() =>
            {
                if (finished || ct.IsCancellationRequested) return;
                if (login == null)
                {
                    Trace("response-stream-unavailable");
                    SetStatus("Der Browser liefert keinen lesbaren Antwortstream. Nach abgeschlossener Anmeldung „Zurück zu Meetling“ wählen.");
                    return;
                }
                Trace("token-response-validated"); finished = true; completion.TrySetResult(login); Close();
            });
        }
        catch (MeetlingApiException ex)
        {
            await dispatcher.InvokeAsync(() => { if (!finished && !ct.IsCancellationRequested) Fail(ex.Reason == "No user loged in" ? "Die Installation liefert keinen angemeldeten Benutzer. Der Server muss einen Anmeldeweg mit Rückkehr zum Login-Endpunkt bereitstellen." : ex.Message); });
        }
        catch (Exception ex)
        {
            // Runtime/stream failures must not discard the authenticated browser session.
            // Log only a fixed phase, exception type and HRESULT, never exception text.
            DiagnosticLog.Write("BrowserLogin " + phase + " failed", ex);
            if (!finished && !ct.IsCancellationRequested)
                await dispatcher.InvokeAsync(() => SetStatus("Login-Antwort nicht lesbar (" + phase + ", 0x" + ex.HResult.ToString("X8") + "). Sitzung bleibt geöffnet; „Zurück zu Meetling“ erneut versuchen."));
        }
        finally { await dispatcher.InvokeAsync(() => { reading = false; if (finished || IsDisposed) dispatcher.Dispose(); }); }
    }
    private void Fail(string message) { if (finished) return; Trace("failed"); finished = true; MessageBox.Show(this, message, "Meetling"); completion.TrySetResult(null); Close(); }
    protected override void Dispose(bool disposing) { if (disposing) { foreach (var popup in popups.ToArray()) { popup.Close(); popup.Dispose(); } popups.Clear(); timeout.Dispose(); browser.Dispose(); if (!reading) dispatcher.Dispose(); } base.Dispose(disposing); }
}
