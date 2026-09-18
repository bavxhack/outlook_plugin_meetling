using System;
using System.IO;

namespace Meetling.OutlookAddIn;

internal static class DiagnosticLog
{
    public static void Write(string message, Exception exception)
    {
        try
        {
            var directory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Meetling", "OutlookAddIn");
            Directory.CreateDirectory(directory);
            File.AppendAllText(Path.Combine(directory, "diagnose.log"), $"{DateTimeOffset.Now:o} {message} {exception.GetType().Name}: {exception.Message}{Environment.NewLine}");
        }
        catch (IOException) { }
        catch (UnauthorizedAccessException) { }
    }
}
