using Meetling.Core;
using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Serialization;

namespace Meetling.OutlookAddIn;

internal sealed class SettingsStore
{
    private readonly string directory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Meetling", "OutlookAddIn");
    private string SettingsPath => Path.Combine(directory, "settings.xml");
    private string SecretPath => Path.Combine(directory, "api-key.bin");

    public MeetlingSettings Load()
    {
        var settings = new MeetlingSettings();
        try
        {
            if (File.Exists(SettingsPath))
                using (var stream = File.OpenRead(SettingsPath))
                    settings = (MeetlingSettings)new XmlSerializer(typeof(MeetlingSettings)).Deserialize(stream);
            if (File.Exists(SecretPath))
                settings.ApiKey = Encoding.UTF8.GetString(ProtectedData.Unprotect(File.ReadAllBytes(SecretPath), null, DataProtectionScope.CurrentUser));
        }
        catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException || ex is CryptographicException || ex is InvalidOperationException)
        { DiagnosticLog.Write("Einstellungen konnten nicht gelesen werden.", ex); }
        return settings;
    }

    public void Save(MeetlingSettings settings)
    {
        Directory.CreateDirectory(directory);
        var publicSettings = new MeetlingSettings { BaseUrl = settings.BaseUrl, Server = settings.Server, OrganizerEmail = settings.OrganizerEmail, KeycloakId = settings.KeycloakId };
        using (var stream = File.Create(SettingsPath)) new XmlSerializer(typeof(MeetlingSettings)).Serialize(stream, publicSettings);
        File.WriteAllBytes(SecretPath, ProtectedData.Protect(Encoding.UTF8.GetBytes(settings.ApiKey), null, DataProtectionScope.CurrentUser));
    }
}
