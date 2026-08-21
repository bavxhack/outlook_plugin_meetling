using Extensibility;
using Meetling.Core;
using Office = Meetling.OutlookAddIn.Interop;
using Outlook = Microsoft.Office.Interop.Outlook;
using System;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace Meetling.OutlookAddIn;

[ComVisible(true)]
[Guid("4BB84EF2-9D83-4710-BDBA-7CBFD6939B48")]
[ProgId("Meetling.OutlookAddIn")]
[ClassInterface(ClassInterfaceType.None)]
public sealed class MeetlingAddIn : IDTExtensibility2, Office.IRibbonExtensibility
{
    private Outlook.Application? application;
    private readonly SettingsStore settingsStore = new SettingsStore();

    public string GetCustomUI(string ribbonId) => ribbonId == "Microsoft.Outlook.Appointment" ? RibbonXml : string.Empty;

    public void CreateConference(Office.IRibbonControl control)
    {
        if (application == null) return;
        try
        {
            var settings = settingsStore.Load();
            ValidateRequiredSettings(settings);
            var appointmentService = new AppointmentService(application);
            var appointment = appointmentService.ReadCurrent();
            var email = SettingsValidator.IsValidEmail(appointment.OrganizerEmail) ? appointment.OrganizerEmail : settings.OrganizerEmail;
            if (!SettingsValidator.IsValidEmail(email)) throw new InvalidOperationException("Die E-Mail-Adresse des Organisators konnte nicht ermittelt werden. Bitte in den Einstellungen hinterlegen.");
            var request = new ConferenceRequest(email, appointment.Subject, AppointmentRules.CalculateDurationMinutes(appointment.Start, appointment.End), settings.Server, new DateTimeOffset(appointment.Start), settings.KeycloakId);
            using (var httpClient = new HttpClient { Timeout = TimeSpan.FromSeconds(30) })
            {
                var result = new MeetlingApiClient(httpClient).CreateConferenceAsync(settings, request, CancellationToken.None).GetAwaiter().GetResult();
                appointmentService.ApplyConference(result.Uid, result.ParticipantUri);
            }
            MessageBox.Show("Die Meetling-Konferenz wurde erstellt und in den Termin eingefügt.", "Meetling", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex) when (ex is MeetlingApiException || ex is InvalidOperationException || ex is ArgumentException)
        { DiagnosticLog.Write("Konferenz konnte nicht erstellt werden.", ex); MessageBox.Show(ex.Message, "Meetling", MessageBoxButtons.OK, MessageBoxIcon.Error); }
    }

    public void OpenSettings(Office.IRibbonControl control)
    {
        try { using (var form = new SettingsForm(settingsStore.Load())) if (form.ShowDialog() == DialogResult.OK) settingsStore.Save(form.Settings); }
        catch (Exception ex) when (ex is System.IO.IOException || ex is UnauthorizedAccessException || ex is System.Security.Cryptography.CryptographicException)
        { DiagnosticLog.Write("Einstellungen konnten nicht gespeichert werden.", ex); MessageBox.Show("Die Einstellungen konnten nicht gespeichert werden.", "Meetling", MessageBoxButtons.OK, MessageBoxIcon.Error); }
    }

    public void OnConnection(object app, ext_ConnectMode connectMode, object addInInstance, ref Array custom) => application = (Outlook.Application)app;
    public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom) { if (application != null && Marshal.IsComObject(application)) Marshal.FinalReleaseComObject(application); application = null; }
    public void OnAddInsUpdate(ref Array custom) { }
    public void OnStartupComplete(ref Array custom) { }
    public void OnBeginShutdown(ref Array custom) { }

    private static void ValidateRequiredSettings(MeetlingSettings settings)
    {
        SettingsValidator.NormalizeBaseUrl(settings.BaseUrl);
        if (string.IsNullOrWhiteSpace(settings.ApiKey)) throw new InvalidOperationException("Bitte hinterlegen Sie zuerst einen API-Schlüssel.");
        if (string.IsNullOrWhiteSpace(settings.Server)) throw new InvalidOperationException("Bitte hinterlegen Sie zuerst einen Konferenzserver.");
    }

    private const string RibbonXml = @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'><ribbon><tabs><tab idMso='TabAppointment'><group id='MeetlingGroup' label='Meetling'><button id='MeetlingCreate' label='Konferenz erstellen' size='large' imageMso='NewMeetingRequest' onAction='CreateConference'/><button id='MeetlingSettings' label='Einstellungen' imageMso='ControlsGallery' onAction='OpenSettings'/></group></tab></tabs></ribbon></customUI>";
}
