using System;

namespace Meetling.Core;

public sealed class MeetlingSettings
{
    public string BaseUrl { get; set; } = string.Empty;
    public string ApiKey { get; set; } = string.Empty;
    public string Server { get; set; } = string.Empty;
    public string OrganizerEmail { get; set; } = string.Empty;
    public string KeycloakId { get; set; } = string.Empty;
}

public sealed class ConferenceRequest
{
    public ConferenceRequest(string email, string name, int durationMinutes, string server, DateTimeOffset start, string? keycloakId = null)
    {
        Email = email;
        Name = name;
        DurationMinutes = durationMinutes;
        Server = server;
        Start = start;
        KeycloakId = keycloakId;
    }

    public string Email { get; }
    public string Name { get; }
    public int DurationMinutes { get; }
    public string Server { get; }
    public DateTimeOffset Start { get; }
    public string? KeycloakId { get; }
}

public sealed class ConferenceResult
{
    public ConferenceResult(string uid, Uri participantUri) { Uid = uid; ParticipantUri = participantUri; }
    public string Uid { get; }
    public Uri ParticipantUri { get; }
}

public sealed class MeetlingApiException : Exception
{
    public MeetlingApiException(string message) : base(message) { }
    public MeetlingApiException(string message, Exception innerException) : base(message, innerException) { }
}
