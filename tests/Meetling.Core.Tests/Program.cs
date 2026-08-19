using Meetling.Core;
using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;

internal static class Program
{
    private static int failures;
    private static int Main()
    {
        Run("URL normalisieren", () => Equal("https://example.org", SettingsValidator.NormalizeBaseUrl(" https://example.org/// ")));
        Run("URL ablehnen", () => Throws<ArgumentException>(() => SettingsValidator.NormalizeBaseUrl("ftp://example.org")));
        Run("HTTP remote warnen", () => True(SettingsValidator.IsInsecureRemoteUrl("http://example.org")));
        Run("HTTP localhost erlauben", () => True(!SettingsValidator.IsInsecureRemoteUrl("http://localhost:8080")));
        Run("Dauer berechnen", () => Equal(61, AppointmentRules.CalculateDurationMinutes(new DateTime(2026, 1, 1, 10, 0, 0), new DateTime(2026, 1, 1, 11, 0, 1))));
        Run("Request und Bearer", TestRequest);
        RunAsync("Erfolgsantworten", () => TestApi("{\"error\":false,\"uid\":\"abc\"}", "{\"error\":false,\"room_url\":\"https://meet.example/room\"}", null));
        RunAsync("API-Fehler", () => TestApi("{\"error\":true}", "{}", "Meetling konnte"));
        RunAsync("Ungültiges JSON", () => TestApi("kein json", "{}", "ungültige JSON"));
        RunAsync("Fehlende UID", () => TestApi("{\"error\":false}", "{}", "keine UID"));
        RunAsync("Fehlender Teilnehmerlink", () => TestApi("{\"error\":false,\"uid\":\"abc\"}", "{\"error\":false}", "keinen Teilnehmerlink"));
        Run("Bestehende Konferenz", () => { True(AppointmentRules.HasExistingConference("abc", null)); True(!AppointmentRules.HasExistingConference(null, null)); });
        Console.WriteLine(failures == 0 ? "Alle Tests erfolgreich." : failures + " Test(s) fehlgeschlagen.");
        return failures == 0 ? 0 : 1;
    }

    private static void TestRequest()
    {
        var request = MeetlingApiClient.BuildCreateRequest("https://meet.example/", "secret", new ConferenceRequest("a@example.org", "Planung", 30, "srv", new DateTimeOffset(2026, 1, 2, 10, 0, 0, TimeSpan.FromHours(1)), "kid"));
        Equal("Bearer", request.Headers.Authorization.Scheme); Equal("secret", request.Headers.Authorization.Parameter);
        Equal("https://meet.example/api/v1/room", request.RequestUri.AbsoluteUri);
        var body = request.Content.ReadAsStringAsync().GetAwaiter().GetResult();
        True(body.Contains("email=a%40example.org")); True(body.Contains("duration=30")); True(body.Contains("keycloakId=kid"));
    }

    private static async Task TestApi(string createJson, string infoJson, string expectedError)
    {
        var handler = new QueueHandler(createJson, infoJson); var client = new MeetlingApiClient(new HttpClient(handler));
        try
        {
            var result = await client.CreateConferenceAsync(new MeetlingSettings { BaseUrl = "https://meet.example", ApiKey = "secret" }, new ConferenceRequest("a@example.org", "Test", 30, "srv", DateTimeOffset.Now), CancellationToken.None);
            if (expectedError != null) throw new Exception("Erwarteter Fehler blieb aus."); Equal("abc", result.Uid); Equal("https://meet.example/room", result.ParticipantUri.AbsoluteUri);
        }
        catch (MeetlingApiException ex) { if (expectedError == null || !ex.Message.Contains(expectedError)) throw; }
    }

    private static void Run(string name, Action test) { try { test(); Console.WriteLine("PASS " + name); } catch (Exception ex) { failures++; Console.WriteLine("FAIL " + name + ": " + ex.Message); } }
    private static void RunAsync(string name, Func<Task> test) => Run(name, () => test().GetAwaiter().GetResult());
    private static void True(bool value) { if (!value) throw new Exception("Bedingung nicht erfüllt."); }
    private static void Equal<T>(T expected, T actual) { if (!EqualityComparer<T>.Default.Equals(expected, actual)) throw new Exception($"Erwartet {expected}, erhalten {actual}."); }
    private static void Throws<T>(Action action) where T : Exception { try { action(); } catch (T) { return; } throw new Exception(typeof(T).Name + " wurde nicht ausgelöst."); }

    private sealed class QueueHandler : HttpMessageHandler
    {
        private readonly Queue<string> responses;
        public QueueHandler(params string[] responses) => this.responses = new Queue<string>(responses);
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken) => Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) { Content = new StringContent(responses.Dequeue()) });
    }
}
