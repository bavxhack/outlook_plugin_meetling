using System;
using System.IO;

namespace Meetling.OutlookAddIn;

internal static class DiagnosticLog
{
    public static void WriteInfo(string message) => WriteLine(message);

    public static void Write(string message, Exception exception)
        => WriteLine($"{message} {exception.GetType().Name}: {exception.Message}");

    private static void WriteLine(string message)
    {
        try
        {
            var directory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Meetling", "OutlookAddIn");
            Directory.CreateDirectory(directory);
            File.AppendAllText(Path.Combine(directory, "diagnose.log"), $"{DateTimeOffset.Now:o} {message}{Environment.NewLine}");
        }
        catch (IOException) { }
        catch (UnauthorizedAccessException) { }
        catch (System.Security.SecurityException) { }
    }
}
