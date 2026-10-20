using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using Meetling.Core;
using Newtonsoft.Json;
namespace Meetling.OutlookAddIn;
internal sealed class SettingsStore
{
    internal static string Root => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Meetling", "OutlookAddIn");
    private readonly ConfigurationStore store = new ConfigurationStore(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "Meetling", "OutlookAddIn", "config.json"), Path.Combine(Root, "config.json"));
    internal ConfigurationSnapshot Load() => store.Load();
    internal void Save(MeetlingSettings settings) => store.SaveUser(settings);
}
internal sealed class TokenStore
{
    private string PathFor(string installation)
    {
        using (var hash = SHA256.Create()) return Path.Combine(SettingsStore.Root, "plugin-token-" + BitConverter.ToString(hash.ComputeHash(Encoding.UTF8.GetBytes(SettingsValidator.NormalizeBaseUrl(installation)))).Replace("-", "") + ".bin");
    }
    private static byte[] Entropy(string installation) => Encoding.UTF8.GetBytes("Meetling.PluginToken.v1|" + SettingsValidator.NormalizeBaseUrl(installation));
    internal LoginResult? Load(string installation)
    {
        var path = PathFor(installation);
        if (!File.Exists(path)) return null;
        var bytes = ProtectedData.Unprotect(File.ReadAllBytes(path), Entropy(installation), DataProtectionScope.CurrentUser);
        try
        {
            var record = JsonConvert.DeserializeObject<TokenRecord>(Encoding.UTF8.GetString(bytes));
            if (record == null || record.Installation != SettingsValidator.NormalizeBaseUrl(installation)) throw new CryptographicException();
            SettingsValidator.ValidateToken(record.Token);
            if (string.IsNullOrWhiteSpace(record.Name)) throw new CryptographicException();
            return new LoginResult(record.Token, record.Name);
        }
        finally { Array.Clear(bytes, 0, bytes.Length); }
    }
    internal void Save(string installation, LoginResult login)
    {
        SettingsValidator.ValidateToken(login.Token);
        var bytes = Encoding.UTF8.GetBytes(JsonConvert.SerializeObject(new TokenRecord { Installation = SettingsValidator.NormalizeBaseUrl(installation), Token = login.Token, Name = login.Name }));
        try { AtomicFile.Write(PathFor(installation), ProtectedData.Protect(bytes, Entropy(installation), DataProtectionScope.CurrentUser)); }
        finally { Array.Clear(bytes, 0, bytes.Length); }
    }
    internal void Delete(string installation) => File.Delete(PathFor(installation));
    private sealed class TokenRecord
    {
        public string Installation { get; set; } = "";
        public string Token { get; set; } = "";
        public string Name { get; set; } = "";
    }
}
