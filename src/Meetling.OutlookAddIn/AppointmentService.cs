using System;
using System.Runtime.InteropServices;
using Meetling.Core;
using Outlook = Microsoft.Office.Interop.Outlook;
namespace Meetling.OutlookAddIn;
// Captured on UI thread before any await; never looks up ActiveInspector after networking.
internal sealed class AppointmentTarget : IDisposable
{
    private object? item;
    private Outlook.Inspector? inspector;
    private readonly bool newItem;
    private AppointmentTarget(object item, Outlook.Inspector inspector, bool newItem) { this.item = item; this.inspector = inspector; this.newItem = newItem; }
    internal static AppointmentTarget Capture(object? application, object? control)
    {
        if (!(application is Outlook.Application outlook)) throw new InvalidOperationException("Outlook ist nicht verbunden.");
        object? context = null;
        try { if (control != null) context = ((dynamic)control).Context; } catch (Exception) { }
        var inspector = context as Outlook.Inspector;
        // Explorer actions always create a new draft, even if another inspector is open.
        if (inspector == null && !(context is Outlook.Explorer)) inspector = outlook.ActiveInspector();
        if (inspector != null)
        {
            object item = inspector.CurrentItem;
            if (!(item is Outlook.AppointmentItem) && !(item is Outlook.MailItem mail && !mail.Sent)) throw new InvalidOperationException("Bitte einen bearbeitbaren Termin oder E-Mail-Entwurf öffnen.");
            var target = new AppointmentTarget(item, inspector, false);
            try { target.EnsureEditable(); target.EnsureNoConference(); return target; } catch { target.Dispose(); throw; }
        }
        var appointment = (Outlook.AppointmentItem)outlook.CreateItem(Outlook.OlItemType.olAppointmentItem);
        appointment.MeetingStatus = Outlook.OlMeetingStatus.olMeeting;
        appointment.Display(false);
        return new AppointmentTarget(appointment, appointment.GetInspector, true);
    }
    private void EnsureEditable()
    {
        if (item is Outlook.AppointmentItem appt && appt.MeetingStatus == Outlook.OlMeetingStatus.olMeetingReceived)
            throw new InvalidOperationException("Eine empfangene Besprechung kann nicht ergänzt werden.");
        dynamic editor = inspector!.WordEditor;
        try { if ((int)editor.ProtectionType != -1 || (bool)editor.ReadOnly) throw new InvalidOperationException("Der Editor ist schreibgeschützt."); }
        finally { Release(editor); }
    }
    private void EnsureNoConference()
    {
        dynamic properties = ((dynamic)item!).UserProperties;
        object? link = null; object? legacyUid = null;
        try { legacyUid = properties.Find(AppointmentRules.UidProperty, true); if (legacyUid != null && !string.IsNullOrWhiteSpace((string)((dynamic)legacyUid).Value)) throw new InvalidOperationException("Dieses Element enthält bereits eine Meetling-Konferenz."); link = properties.Find(AppointmentRules.ParticipantLinkProperty, true); if (link != null && !string.IsNullOrWhiteSpace((string)((dynamic)link).Value)) throw new InvalidOperationException("Dieses Element enthält bereits eine Meetling-Konferenz."); }
        finally { Release(legacyUid); Release(link); Release(properties); }
    }
    internal void MarkPending()
    {
        EnsureNoConference();
        dynamic properties = ((dynamic)item!).UserProperties;
        object? property = null;
        try
        {
            property = properties.Find("MeetlingCreationState", true);
            if (property != null && !string.IsNullOrEmpty((string)((dynamic)property).Value)) throw new InvalidOperationException("Für dieses Element wurde bereits eine Erstellung begonnen. Bei unklarem Ausgang bitte die Installation prüfen; es erfolgt keine automatische Wiederholung.");
            property = property ?? properties.Add("MeetlingCreationState", Outlook.OlUserPropertyType.olText, false);
            ((dynamic)property).Value = "pending";
        }
        finally { Release(property); Release(properties); }
    }
    internal void Apply(ConferenceResult conference)
    {
        EnsureEditable(); EnsureNoConference();
        dynamic editor = inspector!.WordEditor; object? range = null; object? content = null;
        try
        {
            content = editor.Content;
            int end = (int)((dynamic)content).End - 1;
            range = editor.Range(end, end);
            ((dynamic)range).InsertAfter("\r\rMeetling-Konferenz\rTeilnehmen: " + conference.Url + "\r");
            dynamic properties = ((dynamic)item!).UserProperties; object? link = null;
            try { link = properties.Find(AppointmentRules.ParticipantLinkProperty, true) ?? properties.Add(AppointmentRules.ParticipantLinkProperty, Outlook.OlUserPropertyType.olText, false); ((dynamic)link).Value = conference.Url; }
            finally { Release(link); Release(properties); }
            if (newItem) ((Outlook.AppointmentItem)item!).Subject = conference.Title;
        }
        finally { Release(range); Release(content); Release(editor); }
        // No Save or Send. Native fields and formatting remain owned by Outlook/user.
    }
    public void Dispose() { Release(inspector); Release(item); inspector = null; item = null; }
    private static void Release(object? value) { if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value); }
}
