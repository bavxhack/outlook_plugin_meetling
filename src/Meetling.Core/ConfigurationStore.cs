using System;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
namespace Meetling.Core;
public sealed class ConfigurationSnapshot
{
    public ConfigurationSnapshot(MeetlingSettings? settings, bool managed, string? error) { Settings = settings; Managed = managed; Error = error; }
    public MeetlingSettings? Settings { get; }
    public bool Managed { get; }
    public string? Error { get; }
}
public sealed class ConfigurationStore
{
    public static readonly JsonSerializerSettings JsonSettings = new JsonSerializerSettings { ContractResolver = new CamelCasePropertyNamesContractResolver(), MissingMemberHandling = MissingMemberHandling.Error };
    private readonly string managedPath, userPath;
    public ConfigurationStore(string managedPath, string userPath) { this.managedPath = managedPath; this.userPath = userPath; }
    public ConfigurationSnapshot Load()
    {
        // File.Exists suppresses access errors. GetAttributes preserves them: fail closed.
        bool managed;
        try { File.GetAttributes(managedPath); managed = true; }
        catch (FileNotFoundException) { managed = false; }
        catch (DirectoryNotFoundException) { managed = false; }
        catch (Exception) { return new ConfigurationSnapshot(null, true, "Verwaltete Konfiguration ist nicht lesbar."); }
        var path = managed ? managedPath : userPath;
        try
        {
            var settings = JsonConvert.DeserializeObject<MeetlingSettings>(File.ReadAllText(path), JsonSettings) ?? throw new ArgumentException();
            settings.UpgradeKnownLegacyRoutes(); settings.Validate(); return new ConfigurationSnapshot(settings, managed, null);
        }
        catch (Exception) { return new ConfigurationSnapshot(null, managed, managed ? "Verwaltete Konfiguration ist ungültig oder nicht lesbar." : "Konfiguration fehlt oder ist ungültig."); }
    }
    public void SaveUser(MeetlingSettings settings)
    {
        if (Load().Managed) throw new InvalidOperationException("Von Ihrer Organisation vorgegeben.");
        settings.Validate(); AtomicFile.Write(userPath, System.Text.Encoding.UTF8.GetBytes(JsonConvert.SerializeObject(settings, Formatting.Indented, JsonSettings)));
    }
}
public static class AtomicFile
{
    public static void Write(string path, byte[] bytes)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        string temp = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try
        {
            using (var stream = new FileStream(temp, FileMode.CreateNew, FileAccess.Write, FileShare.None)) { stream.Write(bytes, 0, bytes.Length); stream.Flush(true); }
            if (File.Exists(path)) File.Replace(temp, path, null); else File.Move(temp, path);
        }
        finally { if (File.Exists(temp)) File.Delete(temp); }
    }
}
// UI-thread owned epoch. Changing it invalidates every outstanding result.
public sealed class SessionEpoch
{
    public int Value { get; private set; }
    public int Advance() => ++Value;
    public bool IsCurrent(int value) => value == Value;
}
