using Extensibility;
using Office = Meetling.OutlookAddIn.Interop;
using System;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Meetling.OutlookAddIn;

[ComVisible(true)]
[Guid("4BB84EF2-9D83-4710-BDBA-7CBFD6939B48")]
[ProgId("Meetling.OutlookAddIn")]
[ClassInterface(ClassInterfaceType.None)]
[ComDefaultInterface(typeof(IMeetlingRibbonCallbacks))]
public sealed class MeetlingAddIn : IDTExtensibility2, Office.IRibbonExtensibility, IMeetlingRibbonCallbacks
{
    // No settings, filesystem or network work in the constructor.
    private object? application;
    private PluginController? controller;
    private Interop.IRibbonUI? ribbon;
    private PluginController Controller
    {
        get
        {
            if (controller == null) { controller = new PluginController(); controller.Changed = () => { try { ribbon?.Invalidate(); } catch (Exception) { } }; }
            return controller;
        }
    }
    private async void Run(Func<Task> operation)
    {
        try { await operation(); }
        catch (Exception ex) { DiagnosticLog.Write("Meetling-Aktion fehlgeschlagen", ex); LocalDialog.Show("Die Aktion konnte nicht abgeschlossen werden. Bitte die Einstellungen und den Anmeldestatus prüfen."); }
    }
    public void RibbonLoaded(object ui) { ribbon = ui as Interop.IRibbonUI; }
    public string GetStatus(object control) => controller?.Status ?? "Konfiguration und Anmeldung über Einstellungen / Anmelden prüfen.";
    public bool GetCreateEnabled(object control) => controller?.CanCreate ?? false;
    public bool GetLoginEnabled(object control) => controller?.CanLogin ?? true;
    public bool GetLogoutEnabled(object control) => controller?.CanLogout ?? false;
    public void SignIn(object control) => Run(() => Controller.SignIn());
    public void SignOut(object control) => Run(() => Controller.Logout());
    public void RefreshServers(object control) => Run(() => Controller.RefreshServers());
    public void ReloadConfiguration(object control) { try { Controller.Reload(); } catch (Exception ex) { DiagnosticLog.Write("Konfiguration neu laden", ex); LocalDialog.Show("Konfiguration konnte nicht geladen werden."); } }

    public MeetlingAddIn() { }

    public string GetCustomUI(string ribbonId)
    {
        try
        {
            DiagnosticLog.WriteInfo("GetCustomUI RibbonId=" + ribbonId);
            switch (ribbonId)
            {
                case "Microsoft.Outlook.Explorer":
                    return RibbonXml.Replace("<tabs>", "<tabs>" + CalendarTabXml);
                case "Microsoft.Outlook.Appointment":
                case "Microsoft.Outlook.Mail.Compose":
                    return RibbonXml;
                default:
                    DiagnosticLog.WriteInfo("GetCustomUI unsupported RibbonId");
                    return string.Empty;
            }
        }
        catch (Exception ex) { DiagnosticLog.Write("GetCustomUI failed", ex); return string.Empty; }
    }

    public object? GetMeetlingImage(object control)
    {
        try { return RibbonImages.GetPicture(); }
        catch (Exception ex) { DiagnosticLog.Write("GetMeetlingImage failed", ex); return null; }
    }

    public void CreateMeeting(object control)
    {
        if (application == null) { LocalDialog.Show("Outlook ist nicht verbunden."); return; }
        Run(() => Controller.Create(application, control));
    }

    public void TestPlugin(object control)
    {
        try
        {
            DiagnosticLog.WriteInfo("TestPlugin entered");
            LocalDialog.Show("Meetling wurde geladen. Der Button funktioniert." +
                Environment.NewLine + "Version: " + DiagnosticLog.Version);
        }
        catch (Exception ex) { DiagnosticLog.Write("TestPlugin failed", ex); }
    }

    public void OpenSettings(object control)
    {
        try
        {
            if (application == null) { LocalDialog.Show("Meetling API-Konfiguration\nVersion: " + DiagnosticLog.Version + "\nDiagnose: " + DiagnosticLog.LogPath); return; }
            Controller.Settings();
        }
        catch (Exception ex) { DiagnosticLog.Write("OpenSettings failed", ex); LocalDialog.Show("Einstellungen konnten nicht geladen oder gespeichert werden."); }
    }

    public void OnConnection(object app, ext_ConnectMode connectMode, object addInInstance, ref Array custom)
    {
        try { application = app; if (System.Threading.SynchronizationContext.Current == null) System.Threading.SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext()); DiagnosticLog.WriteInfo("OnConnection ConnectMode=" + connectMode); }
        catch (Exception ex) { DiagnosticLog.Write("OnConnection failed", ex); }
    }

    public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
    { controller?.Dispose(); controller = null; ribbon = null; application = null; DiagnosticLog.WriteInfo("OnDisconnection RemoveMode=" + removeMode); }
    public void OnAddInsUpdate(ref Array custom) => DiagnosticLog.WriteInfo("OnAddInsUpdate");
    public void OnStartupComplete(ref Array custom)
    {
        DiagnosticLog.WriteInfo("OnStartupComplete");
        if (application == null) return;
        try { _ = Controller; ribbon?.Invalidate(); } catch (Exception ex) { DiagnosticLog.Write("Configuration initialization failed", ex); }
    }
    public void OnBeginShutdown(ref Array custom)
    { controller?.Dispose(); controller = null; ribbon = null; application = null; DiagnosticLog.WriteInfo("OnBeginShutdown"); }

    // Calendar Home, existing callback and image DISPIDs preserved.
    private const string CalendarTabXml = @"<tab idMso='TabCalendar'>
    <group id='MeetlingCalendarGroup' label='Meetling' insertAfterMso='GroupCalendarNew' getImage='GetMeetlingImage'>
      <button id='MeetlingCalendarStart' label='Neues Meetling' size='large' getImage='GetMeetlingImage'
        getEnabled='GetCreateEnabled' onAction='CreateMeeting' screentip='Meetling-Konferenz in einer nativen Outlook-Besprechung erstellen.'/>
    </group>
  </tab>";

    private const string RibbonXml = @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='RibbonLoaded'>
  <ribbon><tabs><tab id='MeetlingTab' label='Meetling'>
    <group id='MeetlingGroup' label='Meetling'>
      <button id='MeetlingTest' label='Plugin testen' size='large' enabled='true' onAction='TestPlugin'/>
      <button id='MeetlingSettings' label='Einstellungen' enabled='true' onAction='OpenSettings'/>
      <button id='MeetlingConference' label='Neues Meetling' getEnabled='GetCreateEnabled' onAction='CreateMeeting' getImage='GetMeetlingImage'/>
      <button id='MeetlingLogin' label='Anmelden' getEnabled='GetLoginEnabled' onAction='SignIn'/>
      <button id='MeetlingLogout' label='Abmelden' getEnabled='GetLogoutEnabled' onAction='SignOut'/>
      <button id='MeetlingServers' label='Server aktualisieren' getEnabled='GetCreateEnabled' onAction='RefreshServers'/>
      <button id='MeetlingReload' label='Konfiguration neu laden' onAction='ReloadConfiguration'/>
      <labelControl id='MeetlingStatus' getLabel='GetStatus'/>
    </group>
  </tab></tabs></ribbon>
</customUI>";
}
