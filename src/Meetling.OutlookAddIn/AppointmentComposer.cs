using System;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Meetling.OutlookAddIn;

internal static class AppointmentComposer
{
    // Only the Outlook UI boundary is replaced by the Windows CI harness.
    internal static Action<object?, MockMeetingDraft> Open = OpenOutlookMeeting;

    private static void OpenOutlookMeeting(object? application, MockMeetingDraft draft)
    {
        var outlook = application as Outlook.Application;
        if (outlook == null) throw new InvalidOperationException("Outlook application unavailable.");
        Outlook.AppointmentItem? item = null;
        try
        {
            item = (Outlook.AppointmentItem)outlook.CreateItem(Outlook.OlItemType.olAppointmentItem);
            item.MeetingStatus = Outlook.OlMeetingStatus.olMeeting;
            item.Subject = draft.Subject;
            var existing = item.Body ?? string.Empty;
            item.Body = string.IsNullOrWhiteSpace(existing) ? draft.BodyText :
                existing + Environment.NewLine + Environment.NewLine + draft.BodyText;
            // Native meeting editor (attendees, date/time and Scheduling Assistant).
            // All Office calls run synchronously on the Ribbon's Outlook UI thread.
            item.Display(false);
        }
        finally
        {
            // Release only the new item reference we own, never Outlook's Application RCW.
            if (item != null && Marshal.IsComObject(item)) Marshal.ReleaseComObject(item);
        }
    }
}
