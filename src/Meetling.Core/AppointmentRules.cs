using System;

namespace Meetling.Core;

public static class AppointmentRules
{
    public const string UidProperty = "MeetlingUid";
    public const string ParticipantLinkProperty = "MeetlingParticipantLink";

    public static int CalculateDurationMinutes(DateTime start, DateTime end)
    {
        var totalMinutes = (end - start).TotalMinutes;
        if (totalMinutes <= 0 || totalMinutes > int.MaxValue)
            throw new ArgumentOutOfRangeException(nameof(end), "Das Terminende muss nach dem Beginn liegen.");
        return (int)Math.Ceiling(totalMinutes);
    }

    public static bool HasExistingConference(string? uid, string? participantLink) =>
        !string.IsNullOrWhiteSpace(uid) || !string.IsNullOrWhiteSpace(participantLink);
}
