using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using Meetling.Core;
using Newtonsoft.Json.Linq;
internal static class Program
{
    private static int checks;
    private static void Check(bool value, string label) { if (!value) throw new Exception(label); checks++; }
    private static void Reject(Action action, string label) { try { action(); } catch (Exception) { checks++; return; } throw new Exception("Accepted: " + label); }
    private static async Task RejectAsync(Func<Task> action, string label) { try { await action(); } catch (Exception) { checks++; return; } throw new Exception("Accepted: " + label); }
    private sealed class Handler : HttpMessageHandler
    {
        internal Func<HttpRequestMessage, Task<HttpResponseMessage>> Respond = _ => Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK));
        internal int Calls;
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken ct) { Calls++; return Respond(request); }
    }
    private static HttpResponseMessage Json(string body, HttpStatusCode status = HttpStatusCode.OK) => new HttpResponseMessage(status) { Content = new StringContent(body) };
    private static async Task<int> Main()
    {
        try
        {
            var settings = new MeetlingSettings { BaseUrl = "https://HOST.example/meetling" };
            Check(settings.Installation == "https://host.example/meetling/", "normalization");
            Check(settings.Endpoint("server").AbsoluteUri == "https://host.example/meetling/room/outlook/plugin/server", "installation subpath");
            foreach (var bad in new[] { "http://host/", "https://user:pw@host/", "https://host/?x", "https://host/#x", "" }) Reject(() => SettingsValidator.NormalizeBaseUrl(bad), "invalid base URL");
            foreach (var bad in new[] { "/server", "../server", "https://other/server", "a\\b", "%2e%2e/a", "server?x" }) Reject(() => settings.Endpoint(bad), "unsafe endpoint");
            foreach (var token in new[] { "", " ", "a\rb", "a\nb", "a\0b" }) Reject(() => MeetlingApiClient.BuildRequest(settings, "servers", token), "bad token");
            foreach (var action in new[] { "servers", "logout", "create" })
            {
                using (var request = MeetlingApiClient.BuildRequest(settings, action, "opaque-token", "9007199254740993"))
                {
                    Check(request.Headers.GetValues(action == "create" ? "token" : "Authorization").Single() == "opaque-token", "exact header");
                    Check(!request.Headers.Contains(action == "create" ? "Authorization" : "token"), "separate header");
                    Check(request.Method == (action == "create" ? HttpMethod.Post : HttpMethod.Get), "method convention");
                    if (action == "create") { Check(await request.Content.ReadAsStringAsync() == "server_id=9007199254740993", "only server ID, no Outlook content"); Check(request.Content.Headers.ContentType.MediaType == "application/x-www-form-urlencoded", "form"); }
                }
            }
            foreach (var reason in new[] { "No user loged in", "No user loged found", "User not found", "Server not found", "Server not allowed for user" })
            {
                try { MeetlingApiClient.Parse("{\"error\":true,\"reason\":\"" + reason + "\"}"); throw new Exception("error accepted"); }
                catch (MeetlingApiException ex) { Check(ex.Reason == reason, "exact reason"); }
            }
            foreach (var bad in new[] { "", " ", "<html>login</html>", "{bad", "{\"token\":\"a\",\"token\":\"b\"}" }) Reject(() => MeetlingApiClient.Parse(bad), "response format");
            foreach (int count in new[] { 0, 1, 2, 4 })
            {
                var array = new JArray(Enumerable.Range(1, count).Select(i => new JObject { ["id"] = i, ["name"] = "Same name" }));
                var servers = MeetlingApiClient.ParseServers(array); Check(servers.Count == count, "server count");
                Check(servers.Select(s => s.Id).Distinct().Count() == count, "ID identity independent of name");
            }
            foreach (var bad in new[] { "{}", "[{}]", "[{\"id\":null,\"name\":\"a\"}]", "[{\"id\":\"\",\"name\":\"a\"}]", "[{\"id\":\"10\",\"name\":\"a\"}]", "[{\"id\":1.5,\"name\":\"a\"}]", "[{\"id\":1,\"name\":\"a\"},{\"id\":1,\"name\":\"b\"}]" }) Reject(() => MeetlingApiClient.ParseServers(JToken.Parse(bad)), "invalid server schema");
            Check(MeetlingApiClient.ParseServers(JToken.Parse("[{\"id\":999999999999999999999999999999,\"name\":\"a\"}]"))[0].Id == "999999999999999999999999999999", "large ID lossless");
            var expected = settings.Endpoint(settings.Endpoints.Login);
            Check(expected.AbsoluteUri == "https://host.example/meetling/room/outlook/plugin/outlooklogin", "default Outlook login endpoint");
            Check(SettingsValidator.IsExpectedLogin(expected, expected.AbsoluteUri), "exact login");
            foreach (var bad in new[] { "https://host.example/meetling/room/outlook/plugin/login", "https://other.example/meetling/room/outlook/plugin/outlooklogin", expected + "/", expected + "?x", expected + "#x", "http://host.example/meetling/room/outlook/plugin/outlooklogin" }) Check(!SettingsValidator.IsExpectedLogin(expected, bad), "reject other login response");
            Check(MeetlingApiClient.ParseLogin("{\"token\":\"opaque\",\"name\":\"Erika\"}").Name == "Erika", "login schema");
            foreach (var bad in new[] { "[]", "{}", "{\"token\":\" \",\"name\":\"a\"}", "{\"token\":123,\"name\":\"a\"}", "{\"token\":\"a\",\"name\":null}" }) Reject(() => MeetlingApiClient.ParseLogin(bad), "invalid login");
            var handler = new Handler(); using (var http = new HttpClient(handler))
            {
                var api = new MeetlingApiClient(http);
                handler.Respond = _ => Task.FromResult(Json("[{\"url\":\"https://host.example/a%2Fb?x=1\",\"title\":\"generated\"}]"));
                var conference = await api.CreateConferenceAsync(settings, "token", "10", CancellationToken.None);
                Check(conference.Url == "https://host.example/a%2Fb?x=1", "URL unchanged");
                foreach (var bad in new[] { "{}", "[]", "[{},{}]", "[{\"url\":\"http://host/a\",\"title\":\"a\"}]", "[{\"url\":\"https://host/a\",\"title\":\"\"}]" }) { handler.Respond = _ => Task.FromResult(Json(bad)); await RejectAsync(() => api.CreateConferenceAsync(settings, "token", "10", CancellationToken.None), "invalid conference"); }
                handler.Respond = _ => Task.FromResult(Json("{\"token\":null}")); await api.LogoutAsync(settings, "token", CancellationToken.None); checks++;
                handler.Respond = _ => Task.FromResult(Json("{}")); await RejectAsync(() => api.LogoutAsync(settings, "token", CancellationToken.None), "unconfirmed logout");
                handler.Respond = _ => Task.FromResult(Json("{\"error\":true,\"reason\":\"No user loged found\"}")); await RejectAsync(() => api.LogoutAsync(settings, "token", CancellationToken.None), "invalid logout token");
                foreach (var status in new[] { HttpStatusCode.Redirect, HttpStatusCode.Unauthorized, HttpStatusCode.InternalServerError }) { handler.Respond = _ => Task.FromResult(Json("[]", status)); await RejectAsync(() => api.GetServersAsync(settings, "token", CancellationToken.None), "HTTP status first"); }
                int before = handler.Calls; handler.Respond = _ => throw new TaskCanceledException(); await RejectAsync(() => api.CreateConferenceAsync(settings, "token", "10", CancellationToken.None), "timeout"); Check(handler.Calls == before + 1, "no retry");
                handler.Respond = _ => throw new HttpRequestException(); await RejectAsync(() => api.LogoutAsync(settings, "token", CancellationToken.None), "network logout");
            }
            string dir = Path.Combine(Path.GetTempPath(), "MeetlingTests-" + Guid.NewGuid()); Directory.CreateDirectory(dir);
            try
            {
                var managed = Path.Combine(dir, "managed.json"); var user = Path.Combine(dir, "user.json"); var store = new ConfigurationStore(managed, user);
                Check(store.Load().Settings == null && !store.Load().Managed, "missing configuration");
                store.SaveUser(settings); Check(store.Load().Settings!.Installation == settings.Installation, "GUI fallback");
                File.WriteAllText(managed, "{\"baseUrl\":\"https://org.example/installation/\"}"); Check(store.Load().Managed && store.Load().Settings!.Installation == "https://org.example/installation/", "managed priority");
                Reject(() => store.SaveUser(settings), "managed lock");
                File.WriteAllText(managed, "{broken"); Check(store.Load().Managed && store.Load().Settings == null, "no GUI fallback for broken policy");
                File.Delete(managed); Check(store.Load().Settings!.Installation == settings.Installation, "fallback after removal");
            }
            finally { Directory.Delete(dir, true); }
            var gate = new OperationGate(); Check(gate.TryEnter() && gate.Busy, "first operation admitted"); Check(!gate.TryEnter(), "double click blocked"); gate.Release(); Check(gate.TryEnter(), "next operation after completion"); gate.Release();
            var epoch = new SessionEpoch(); int old = epoch.Value; epoch.Advance(); Check(!epoch.IsCurrent(old), "stale response discarded");
            Console.WriteLine("PASS: " + checks + " API/configuration/session checks. HTTP simulated; live integration not tested."); return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
    }
}
