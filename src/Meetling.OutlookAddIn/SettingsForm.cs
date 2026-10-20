using System;
using System.Drawing;
using System.Windows.Forms;
using Meetling.Core;
namespace Meetling.OutlookAddIn;
internal sealed class SettingsForm : Form
{
    private readonly TextBox url;
    internal MeetlingSettings Settings { get; private set; }
    internal SettingsForm(ConfigurationSnapshot snapshot)
    {
        Settings = snapshot.Settings ?? new MeetlingSettings(); Text = "Meetling – Einstellungen"; ClientSize = new Size(590, 185); StartPosition = FormStartPosition.CenterScreen;
        Controls.Add(new Label { Text = snapshot.Managed ? "Von Ihrer Organisation vorgegeben" : "Installationsadresse (HTTPS, gegebenenfalls mit Unterpfad)", Left = 15, Top = 15, Width = 560 });
        url = new TextBox { Text = Settings.BaseUrl, ReadOnly = snapshot.Managed, Left = 15, Top = 43, Width = 560 }; Controls.Add(url);
        Controls.Add(new Label { Text = snapshot.Error ?? "Nach Änderungen bitte erneut anmelden.", Left = 15, Top = 80, Width = 560, Height = 40 });
        var save = new Button { Text = snapshot.Managed ? "Schließen" : "Speichern", Left = 340, Top = 135, Width = 110 };
        save.Click += (s,e) => { try { if (!snapshot.Managed) { Settings.BaseUrl = SettingsValidator.NormalizeBaseUrl(url.Text); Settings.Validate(); DialogResult = DialogResult.OK; } else DialogResult = DialogResult.Cancel; Close(); } catch (Exception) { MessageBox.Show("Bitte eine gültige HTTPS-Installationsadresse eingeben."); } };
        Controls.Add(save);
        var reload = new Button { Text = "Neu laden", Left = 465, Top = 135, Width = 110, DialogResult = DialogResult.Retry }; Controls.Add(reload);
    }
}
