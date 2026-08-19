using Meetling.Core;
using Outlook = Microsoft.Office.Interop.Outlook;
using System;
using System.Net;
using System.Runtime.InteropServices;

namespace Meetling.OutlookAddIn;

internal sealed class AppointmentSnapshot
{
    public AppointmentSnapshot(string subject, DateTime start, DateTime end, string organizerEmail)
    { Subject = subject; Start = start; End = end; OrganizerEmail = organizerEmail; }
    public string Subject { get; }
    public DateTime Start { get; }
    public DateTime End { get; }
    public string OrganizerEmail { get; }
}

internal sealed class AppointmentService
{
    private readonly Outlook.Application application;
    public AppointmentService(Outlook.Application application) => this.application = application;

    public AppointmentSnapshot ReadCurrent()
    {
        Outlook.Inspector? inspector = null; Outlook.AppointmentItem? item = null;
        try
        {
            inspector = application.ActiveInspector();
            if (inspector == null || !(inspector.CurrentItem is Outlook.AppointmentItem appointment))
                throw new InvalidOperationException("Der geöffnete Outlook-Eintrag ist kein Termin.");
            item = appointment;
            if (HasConference(item)) throw new InvalidOperationException("Dieser Termin enthält bereits eine Meetling-Konferenz.");
            return new AppointmentSnapshot(item.Subject ?? "Outlook-Termin", item.Start, item.End, ResolveSenderAddress(item));
        }
        finally { Release(item); Release(inspector); }
    }

    public void ApplyConference(string uid, Uri participantUri)
    {
        Outlook.Inspector? inspector = null; Outlook.AppointmentItem? item = null; Outlook.UserProperties? properties = null;
        try
        {
            inspector = application.ActiveInspector();
            if (inspector == null || !(inspector.CurrentItem is Outlook.AppointmentItem appointment)) throw new InvalidOperationException("Der geöffnete Outlook-Eintrag ist kein Termin.");
            item = appointment;
            if (HasConference(item)) throw new InvalidOperationException("Dieser Termin enthält bereits eine Meetling-Konferenz.");
            var encoded = WebUtility.HtmlEncode(participantUri.AbsoluteUri);
            if (item.BodyFormat == Outlook.OlBodyFormat.olFormatHTML)
                item.HTMLBody = $"<p><strong>Meetling-Konferenz:</strong> <a href=\"{encoded}\">{encoded}</a></p>" + (item.HTMLBody ?? string.Empty);
            else
                item.Body = $"Meetling-Konferenz: {participantUri.AbsoluteUri}{Environment.NewLine}{Environment.NewLine}" + (item.Body ?? string.Empty);
            properties = item.UserProperties;
            SetProperty(properties, AppointmentRules.UidProperty, uid);
            SetProperty(properties, AppointmentRules.ParticipantLinkProperty, participantUri.AbsoluteUri);
            item.Save();
        }
        finally { Release(properties); Release(item); Release(inspector); }
    }

    private static bool HasConference(Outlook.AppointmentItem item)
    {
        Outlook.UserProperties? properties = null; Outlook.UserProperty? uid = null; Outlook.UserProperty? link = null;
        try
        {
            properties = item.UserProperties;
            uid = properties.Find(AppointmentRules.UidProperty, true); link = properties.Find(AppointmentRules.ParticipantLinkProperty, true);
            return AppointmentRules.HasExistingConference(uid?.Value?.ToString(), link?.Value?.ToString());
        }
        finally { Release(link); Release(uid); Release(properties); }
    }

    private string ResolveSenderAddress(Outlook.AppointmentItem item)
    {
        Outlook.Account? account = null; Outlook.NameSpace? session = null; Outlook.Recipient? currentUser = null; Outlook.AddressEntry? entry = null; Outlook.ExchangeUser? exchangeUser = null;
        try
        {
            account = item.SendUsingAccount;
            if (SettingsValidator.IsValidEmail(account?.SmtpAddress)) return account.SmtpAddress;
            session = application.Session; currentUser = session.CurrentUser; entry = currentUser.AddressEntry;
            if (entry != null && entry.Type == "EX") { exchangeUser = entry.GetExchangeUser(); if (SettingsValidator.IsValidEmail(exchangeUser?.PrimarySmtpAddress)) return exchangeUser.PrimarySmtpAddress; }
            return entry?.Address ?? string.Empty;
        }
        finally { Release(exchangeUser); Release(entry); Release(currentUser); Release(session); Release(account); }
    }

    private static void SetProperty(Outlook.UserProperties properties, string name, string value)
    {
        Outlook.UserProperty? property = null;
        try { property = properties.Find(name, true) ?? properties.Add(name, Outlook.OlUserPropertyType.olText, false, Type.Missing); property.Value = value; }
        finally { Release(property); }
    }

    private static void Release(object? value) { if (value != null && Marshal.IsComObject(value)) Marshal.FinalReleaseComObject(value); }
}
