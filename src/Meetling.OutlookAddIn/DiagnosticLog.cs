using System;
using System.IO;

namespace Meetling.OutlookAddIn;

internal static class DiagnosticLog
{
    private const long MaxBytes = 256 * 1024;
    private static readonly object Gate = new object();
    public static string Version => typeof(MeetlingAddIn).Assembly.GetName().Version?.ToString() ?? "unknown";
    public static string LogPath => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "Meetling", "OutlookAddIn", "diagnose.log");

    public static void WriteInfo(string message) => WriteLine(message);
    // Do not include exception messages: future exceptions may contain secrets or item content.
    public static void Write(string message, Exception exception)
    {
        try { WriteLine(message + " " + Meetling.Core.TransportFailure.SafeDetails(exception)); }
        catch (Exception) { /* Diagnostics must never break Office. */ }
    }

    private static void WriteLine(string message)
    {
        try
        {
            lock (Gate)
            {
                var path = LogPath;
                Directory.CreateDirectory(Path.GetDirectoryName(path)!);
                if (File.Exists(path) && new FileInfo(path).Length >= MaxBytes)
                {
                    File.Delete(path + ".1");
                    File.Move(path, path + ".1");
                }
                // Bound each entry, including host-supplied Ribbon IDs.
                var bounded = message.Replace("\r", " ").Replace("\n", " ");
                if (bounded.Length > 1024) bounded = bounded.Substring(0, 1024);
                File.AppendAllText(path, $"{DateTimeOffset.Now:o} Version={Version} " +
                    $"Process={(Environment.Is64BitProcess ? "x64" : "x86")} {bounded}{Environment.NewLine}");
            }
        }
        catch (Exception) { /* Includes IO, access, security and path failures. */ }
    }
}
