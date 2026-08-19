using Meetling.Core;
using System;
using System.Drawing;
using System.Windows.Forms;

namespace Meetling.OutlookAddIn;

internal sealed class SettingsForm : Form
{
    private readonly TextBox baseUrl = new TextBox();
    private readonly TextBox apiKey = new TextBox { UseSystemPasswordChar = true };
    private readonly TextBox server = new TextBox();
    private readonly TextBox organizerEmail = new TextBox();
    private readonly TextBox keycloakId = new TextBox();
    public MeetlingSettings Settings { get; private set; }

    public SettingsForm(MeetlingSettings settings)
    {
        Settings = settings; Text = "Meetling-Einstellungen"; ClientSize = new Size(520, 250); FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false; MinimizeBox = false; StartPosition = FormStartPosition.CenterScreen;
        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(12), ColumnCount = 2, RowCount = 6 };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 190)); layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        AddRow(layout, 0, "Meetling-Basis-URL", baseUrl, settings.BaseUrl);
        AddRow(layout, 1, "API-Schlüssel", apiKey, settings.ApiKey);
        AddRow(layout, 2, "Konferenzserver", server, settings.Server);
        AddRow(layout, 3, "Organisator-E-Mail (optional)", organizerEmail, settings.OrganizerEmail);
        AddRow(layout, 4, "Keycloak-ID (optional)", keycloakId, settings.KeycloakId);
        var buttons = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, Dock = DockStyle.Fill };
        var save = new Button { Text = "Speichern", DialogResult = DialogResult.None, AutoSize = true };
        var cancel = new Button { Text = "Abbrechen", DialogResult = DialogResult.Cancel, AutoSize = true };
        save.Click += SaveClicked; buttons.Controls.Add(save); buttons.Controls.Add(cancel); layout.Controls.Add(buttons, 1, 5);
        Controls.Add(layout); AcceptButton = save; CancelButton = cancel;
    }

    private static void AddRow(TableLayoutPanel layout, int row, string label, TextBox box, string value)
    { box.Text = value; box.Dock = DockStyle.Fill; layout.Controls.Add(new Label { Text = label, AutoSize = true, Anchor = AnchorStyles.Left }, 0, row); layout.Controls.Add(box, 1, row); }

    private void SaveClicked(object sender, EventArgs e)
    {
        try
        {
            var normalized = SettingsValidator.NormalizeBaseUrl(baseUrl.Text);
            if (string.IsNullOrWhiteSpace(apiKey.Text) || string.IsNullOrWhiteSpace(server.Text)) { MessageBox.Show("API-Schlüssel und Konferenzserver sind erforderlich.", "Meetling", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            if (!string.IsNullOrWhiteSpace(organizerEmail.Text) && !SettingsValidator.IsValidEmail(organizerEmail.Text)) { MessageBox.Show("Die Organisator-E-Mail ist ungültig.", "Meetling", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            if (SettingsValidator.IsInsecureRemoteUrl(normalized) && MessageBox.Show("Die Verbindung verwendet unverschlüsseltes HTTP. Trotzdem speichern?", "Meetling", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) return;
            Settings = new MeetlingSettings { BaseUrl = normalized, ApiKey = apiKey.Text, Server = server.Text.Trim(), OrganizerEmail = organizerEmail.Text.Trim(), KeycloakId = keycloakId.Text.Trim() };
            DialogResult = DialogResult.OK; Close();
        }
        catch (ArgumentException ex) { MessageBox.Show(ex.Message, "Meetling", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
    }
}
