using Extensibility;
using Office = Meetling.OutlookAddIn.Interop;
using System;
using System.Runtime.InteropServices;

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
        try
        {
            DiagnosticLog.WriteInfo("CreateMeeting entered (mock)");
            AppointmentComposer.Open(application, MockMeetingDraft.Create());
        }
        catch (Exception ex)
        {
            DiagnosticLog.Write("CreateMeeting failed", ex);
            try { LocalDialog.Show("Der Outlook-Termin konnte nicht geöffnet werden. Bitte die Meetling-Diagnose prüfen."); }
            catch (Exception dialogError) { DiagnosticLog.Write("CreateMeeting error dialog failed", dialogError); }
        }
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
            DiagnosticLog.WriteInfo("OpenSettings entered");
            LocalDialog.Show("Basisversion: Die Server- und API-Konfiguration folgt im nächsten Schritt." +
                Environment.NewLine + "Version: " + DiagnosticLog.Version +
                Environment.NewLine + "Diagnose: " + DiagnosticLog.LogPath);
        }
        catch (Exception ex) { DiagnosticLog.Write("OpenSettings failed", ex); }
    }

    public void OnConnection(object app, ext_ConnectMode connectMode, object addInInstance, ref Array custom)
    {
        try { application = app; DiagnosticLog.WriteInfo("OnConnection ConnectMode=" + connectMode); }
        catch (Exception ex) { DiagnosticLog.Write("OnConnection failed", ex); }
    }

    public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
    { application = null; DiagnosticLog.WriteInfo("OnDisconnection RemoveMode=" + removeMode); }
    public void OnAddInsUpdate(ref Array custom) => DiagnosticLog.WriteInfo("OnAddInsUpdate");
    public void OnStartupComplete(ref Array custom) => DiagnosticLog.WriteInfo("OnStartupComplete");
    public void OnBeginShutdown(ref Array custom)
    { application = null; DiagnosticLog.WriteInfo("OnBeginShutdown"); }

    // Calendar Home only; creation is explicitly mocked and performs no API request.
    private const string CalendarTabXml = @"<tab idMso='TabCalendar'>
    <group id='MeetlingCalendarGroup' label='Meetling' insertAfterMso='GroupCalendarNew' getImage='GetMeetlingImage'>
      <button id='MeetlingCalendarStart' label='Neues Meetling' size='large' getImage='GetMeetlingImage'
        enabled='true' onAction='CreateMeeting' screentip='Neue Outlook-Besprechung mit Meetling-Testlink öffnen.'/>
    </group>
  </tab>";

    private const string RibbonXml = @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
  <ribbon><tabs><tab id='MeetlingTab' label='Meetling'>
    <group id='MeetlingGroup' label='Meetling'>
      <button id='MeetlingTest' label='Plugin testen' size='large' enabled='true' onAction='TestPlugin'/>
      <button id='MeetlingSettings' label='Einstellungen' enabled='true' onAction='OpenSettings'/>
      <button id='MeetlingConference' label='Meetling-Konferenz' enabled='false'
        screentip='Wird nach erfolgreichem Test der Basisversion aktiviert.'/>
    </group>
  </tab></tabs></ribbon>
</customUI>";
}
