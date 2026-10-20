using System;

namespace Meetling.OutlookAddIn;

internal sealed class MockMeetingDraft
{
    internal string Subject => "Meetling-Besprechung (Test)";
    internal Uri MeetingLink { get; }
    internal string BodyText => "Meetling-Besprechung" + Environment.NewLine +
        "Testlink (Mock – keine echte Konferenz):" + Environment.NewLine + MeetingLink.AbsoluteUri;

    private MockMeetingDraft(Uri link) => MeetingLink = link;

    internal static MockMeetingDraft Create() => new MockMeetingDraft(
        new Uri("https://meetling.example.invalid/mock/" + Guid.NewGuid().ToString("N")));
}
