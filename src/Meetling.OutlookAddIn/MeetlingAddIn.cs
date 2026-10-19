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
    // No field initializers, Outlook RCWs, settings or network dependencies on the startup path.
    public MeetlingAddIn() { }

    public string GetCustomUI(string ribbonId)
    {
        try
        {
            DiagnosticLog.WriteInfo("GetCustomUI RibbonId=" + ribbonId);
            switch (ribbonId)
            {
                case "Microsoft.Outlook.Explorer":
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

    public void TestPlugin(Office.IRibbonControl control)
    {
        try
        {
            DiagnosticLog.WriteInfo("TestPlugin entered");
            LocalDialog.Show("Meetling wurde geladen. Der Button funktioniert." +
                Environment.NewLine + "Version: " + DiagnosticLog.Version);
        }
        catch (Exception ex) { DiagnosticLog.Write("TestPlugin failed", ex); }
    }

    public void OpenSettings(Office.IRibbonControl control)
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
        try { DiagnosticLog.WriteInfo("OnConnection ConnectMode=" + connectMode); }
        catch (Exception ex) { DiagnosticLog.Write("OnConnection failed", ex); }
    }

    public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        => DiagnosticLog.WriteInfo("OnDisconnection RemoveMode=" + removeMode);
    public void OnAddInsUpdate(ref Array custom) => DiagnosticLog.WriteInfo("OnAddInsUpdate");
    public void OnStartupComplete(ref Array custom) => DiagnosticLog.WriteInfo("OnStartupComplete");
    public void OnBeginShutdown(ref Array custom) => DiagnosticLog.WriteInfo("OnBeginShutdown");

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
