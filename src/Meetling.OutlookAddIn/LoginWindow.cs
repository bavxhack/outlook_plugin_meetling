using System;
using System.IO;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Meetling.Core;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
namespace Meetling.OutlookAddIn;
// This type and all WebView2 references are reached only through the explicit login action.
internal sealed class LoginWindow : Form
{
    private readonly WebView2 browser = new WebView2 { Dock = DockStyle.Fill };
    private readonly Uri expected;
    private readonly CancellationToken ct;
    private readonly TaskCompletionSource<LoginResult?> completion = new TaskCompletionSource<LoginResult?>();
    private readonly System.Windows.Forms.Timer timeout = new System.Windows.Forms.Timer { Interval = 300000 };
    private bool finished, reading;
    private LoginWindow(Uri expected, CancellationToken ct)
    {
        this.expected = expected; this.ct = ct;
        Text = "Meetling – Anmelden"; Width = 900; Height = 720; Controls.Add(browser);
        FormClosed += (s,e) => { finished = true; timeout.Stop(); completion.TrySetResult(null); };
        timeout.Tick += (s,e) => Fail("Die Anmeldung wurde nach fünf Minuten beendet.");
    }
    [MethodImpl(MethodImplOptions.NoInlining)]
    internal static async Task<LoginResult?> SignInAsync(Uri endpoint, CancellationToken ct)
    {
        using (var form = new LoginWindow(endpoint, ct))
        using (ct.Register(() => { try { if (form.IsHandleCreated) form.BeginInvoke((Action)(() => form.Close())); } catch (InvalidOperationException) { } }))
        {
            form.Show();
            try { await form.Initialize(); } catch (Exception) { if (!form.finished) form.Fail("Browser-Anmeldung nicht verfügbar. Bitte Microsoft Edge WebView2 Evergreen Runtime installieren oder die Installation prüfen."); }
            return await form.completion.Task;
        }
    }
    private async Task Initialize()
    {
        var environment = await CoreWebView2Environment.CreateAsync(null, Path.Combine(SettingsStore.Root, "WebView2", Guid.NewGuid().ToString("N")));
        if (finished || ct.IsCancellationRequested) return;
        await browser.EnsureCoreWebView2Async(environment);
        if (finished || ct.IsCancellationRequested) return;
        browser.CoreWebView2.Settings.AreDevToolsEnabled = false;
        browser.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;
        browser.CoreWebView2.NewWindowRequested += (s,e) => { e.Handled = true; MessageBox.Show(this, "Bitte den Anmeldeprozess im geöffneten Fenster verwenden. Popup-Anmeldungen benötigen eine Prüfung durch Ihre Organisation.", "Meetling"); };
        browser.CoreWebView2.NavigationStarting += (s,e) => { if (!Uri.TryCreate(e.Uri, UriKind.Absolute, out var uri) || uri.Scheme != Uri.UriSchemeHttps) e.Cancel = true; };
        browser.CoreWebView2.WebResourceResponseReceived += Response;
        timeout.Start(); browser.CoreWebView2.Navigate(expected.AbsoluteUri);
    }
    private async void Response(object? sender, CoreWebView2WebResourceResponseReceivedEventArgs e)
    {
        if (finished || reading || ct.IsCancellationRequested || e.Request.Method != "GET" || !SettingsValidator.IsExpectedLogin(expected, e.Request.Uri)) return;
        if (e.Response.StatusCode >= 300 && e.Response.StatusCode < 400) return; // SSO redirects are allowed in browser only.
        reading = true;
        try
        {
            if (e.Response.StatusCode != 200) throw new MeetlingApiException("Der Login-Endpunkt meldet einen HTTP-Fehler.");
            var contentType = e.Response.Headers.GetHeader("Content-Type");
            if (contentType.IndexOf("application/json", StringComparison.OrdinalIgnoreCase) < 0) return; // Installation may show its own HTML sign-in UI.
            using (var stream = await e.Response.GetContentAsync())
            using (var reader = new StreamReader(stream))
            {
                string json = await reader.ReadToEndAsync();
                if (finished || ct.IsCancellationRequested) return;
                var login = MeetlingApiClient.ParseLogin(json);
                finished = true; completion.TrySetResult(login); Close();
            }
        }
        catch (MeetlingApiException ex) { if (!finished) Fail(ex.Reason == "No user loged in" ? "Die Installation liefert keinen angemeldeten Benutzer. Der Server muss einen Anmeldeweg mit Rückkehr zum Login-Endpunkt bereitstellen." : ex.Message); }
        catch (Exception) { if (!finished) Fail("Die Login-Antwort konnte nicht sicher übernommen werden."); }
        finally { reading = false; }
    }
    private void Fail(string message) { if (finished) return; finished = true; MessageBox.Show(this, message, "Meetling"); completion.TrySetResult(null); Close(); }
    protected override void Dispose(bool disposing) { if (disposing) { timeout.Dispose(); browser.Dispose(); } base.Dispose(disposing); }
}
