using System;
using System.Collections.Generic;
using System.Globalization;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Meetling.Core;

public sealed class MeetlingApiClient
{
    private readonly HttpClient httpClient;

    public MeetlingApiClient(HttpClient httpClient) => this.httpClient = httpClient ?? throw new ArgumentNullException(nameof(httpClient));

    public async Task<ConferenceResult> CreateConferenceAsync(MeetlingSettings settings, ConferenceRequest request, CancellationToken cancellationToken)
    {
        var baseUrl = SettingsValidator.NormalizeBaseUrl(settings.BaseUrl);
        using (var createMessage = BuildCreateRequest(baseUrl, settings.ApiKey, request))
        using (var createResponse = await SendAsync(createMessage, cancellationToken).ConfigureAwait(false))
        {
            var createBody = await createResponse.Content.ReadAsStringAsync().ConfigureAwait(false);
            EnsureSuccess(createResponse, createBody);
            var uid = ReadRequiredString(createBody, "uid", "Die Meetling-Antwort enthält keine UID.");

            using (var infoMessage = BuildInfoRequest(baseUrl, settings.ApiKey, uid))
            using (var infoResponse = await SendAsync(infoMessage, cancellationToken).ConfigureAwait(false))
            {
                var infoBody = await infoResponse.Content.ReadAsStringAsync().ConfigureAwait(false);
                EnsureSuccess(infoResponse, infoBody);
                var link = ReadRequiredString(infoBody, "room_url", "Die Meetling-Antwort enthält keinen Teilnehmerlink.");
                if (!Uri.TryCreate(link, UriKind.Absolute, out var participantUri) ||
                    (participantUri.Scheme != Uri.UriSchemeHttps && participantUri.Scheme != Uri.UriSchemeHttp))
                    throw new MeetlingApiException("Meetling hat keinen gültigen Teilnehmerlink geliefert.");
                return new ConferenceResult(uid, participantUri);
            }
        }
    }

    public static HttpRequestMessage BuildCreateRequest(string baseUrl, string apiKey, ConferenceRequest request)
    {
        var fields = new List<KeyValuePair<string, string>>
        {
            new KeyValuePair<string, string>("email", request.Email),
            new KeyValuePair<string, string>("name", request.Name),
            new KeyValuePair<string, string>("duration", request.DurationMinutes.ToString(CultureInfo.InvariantCulture)),
            new KeyValuePair<string, string>("server", request.Server),
            new KeyValuePair<string, string>("start", request.Start.ToString("o", CultureInfo.InvariantCulture))
        };
        if (request.KeycloakId is { } keycloakId && !string.IsNullOrWhiteSpace(keycloakId))
        {
            fields.Add(new KeyValuePair<string, string>("keycloakId", keycloakId));
        }
        var message = new HttpRequestMessage(HttpMethod.Post, SettingsValidator.NormalizeBaseUrl(baseUrl) + "/api/v1/room")
        {
            Content = new FormUrlEncodedContent(fields)
        };
        message.Headers.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);
        return message;
    }

    public static HttpRequestMessage BuildInfoRequest(string baseUrl, string apiKey, string uid)
    {
        var message = new HttpRequestMessage(HttpMethod.Get,
            SettingsValidator.NormalizeBaseUrl(baseUrl) + "/api/v1/info/" + Uri.EscapeDataString(uid));
        message.Headers.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);
        return message;
    }

    private async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
    {
        try { return await httpClient.SendAsync(request, cancellationToken).ConfigureAwait(false); }
        catch (TaskCanceledException ex) { throw new MeetlingApiException("Zeitüberschreitung bei der Verbindung zu Meetling.", ex); }
        catch (HttpRequestException ex) { throw new MeetlingApiException("Meetling ist nicht erreichbar.", ex); }
    }

    private void EnsureSuccess(HttpResponseMessage response, string json)
    {
        if (!response.IsSuccessStatusCode)
            throw new MeetlingApiException("Meetling meldet einen HTTP-Fehler (" + (int)response.StatusCode + ").");
        var data = Deserialize(json);
        if (data.TryGetValue("error", out var error) && error is bool isError && isError)
            throw new MeetlingApiException("Meetling konnte die Konferenz nicht erstellen.");
    }

    private string ReadRequiredString(string json, string field, string missingMessage)
    {
        var data = Deserialize(json);
        if (!data.TryGetValue(field, out var value) || !(value is string text) || string.IsNullOrWhiteSpace(text))
            throw new MeetlingApiException(missingMessage);
        return text;
    }

    private Dictionary<string, object> Deserialize(string json)
    {
        try
        {
            var serializer = new DataContractJsonSerializer(typeof(Dictionary<string, object>),
                new DataContractJsonSerializerSettings { UseSimpleDictionaryFormat = true });
            using (var stream = new System.IO.MemoryStream(Encoding.UTF8.GetBytes(json)))
                return (Dictionary<string, object>)serializer.ReadObject(stream) ?? throw new InvalidOperationException();
        }
        catch (Exception ex) when (ex is FormatException || ex is InvalidOperationException || ex is System.Runtime.Serialization.SerializationException)
        { throw new MeetlingApiException("Meetling hat eine ungültige JSON-Antwort geliefert.", ex); }
    }
}
