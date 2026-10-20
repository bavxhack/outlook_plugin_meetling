using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
namespace Meetling.Core;
public static class BrowserLoginResponse
{
    // WebView2 may return no stream for canceled navigations or downloads.
    // No DOM, cookies, URLs or alternate hosts are used to recover a token.
    public static async Task<LoginResult?> ReadAsync(Stream? stream, CancellationToken ct)
    {
        ct.ThrowIfCancellationRequested();
        if (stream == null) return null;
        using (var reader = new StreamReader(stream))
        {
            var json = await reader.ReadToEndAsync().ConfigureAwait(false);
            ct.ThrowIfCancellationRequested();
            return MeetlingApiClient.ParseLogin(json);
        }
    }
}
