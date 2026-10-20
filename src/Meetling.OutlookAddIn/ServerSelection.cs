using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Meetling.Core;
namespace Meetling.OutlookAddIn;
internal static class ServerSelection
{
    internal static Func<IReadOnlyList<Server>, Server?> Confirm = ShowDialog;
    internal static Server? Choose(IReadOnlyList<Server> servers)
    {
        if (servers.Count == 0) { LocalDialog.Show("Keine Server verfügbar."); return null; }
        if (servers.Count == 1) return servers[0];
        return Confirm(servers);
    }
    private static Server? ShowDialog(IReadOnlyList<Server> servers)
    {
        using (var form = new ServerPickerForm(servers))
            return form.ShowDialog() == DialogResult.OK ? form.SelectedServer : null;
    }
}
// Native WinForms controls, Windows contrast colors and Segoe UI match the Office desktop idiom.
internal sealed class ServerPickerForm : Form
{
    private TableLayoutPanel layout = null!;
    private FlowLayoutPanel serverPanel = null!;
    private readonly List<Button> buttons = new List<Button>();
    private readonly ToolTip tips = new ToolTip();
    private readonly Bitmap cube = RibbonImages.RenderCube(28);
    private readonly Font bodyFont = new Font("Segoe UI", 10F);
    private readonly Font headingFont = new Font("Segoe UI", 16F, FontStyle.Bold);
    internal IReadOnlyList<Button> ServerButtons => buttons;
    internal Server? SelectedServer { get; private set; }
    internal ServerPickerForm(IReadOnlyList<Server> servers)
    {
        Text = "Meetling – Server auswählen";
        AutoScaleDimensions = new SizeF(96F, 96F); AutoScaleMode = AutoScaleMode.Dpi; Font = bodyFont;
        BackColor = SystemColors.Window; ForeColor = SystemColors.WindowText;
        ClientSize = new Size(560, 190 + Math.Min(5, servers.Count) * 88);
        MinimumSize = new Size(460, 320);
        StartPosition = FormStartPosition.CenterParent;
        ShowInTaskbar = false; MinimizeBox = false; MaximizeBox = false;
        var root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 4, Padding = new Padding(24, 20, 24, 16), BackColor = SystemColors.Window };
        layout = root;
        root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize)); root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        Controls.Add(root);
        root.Controls.Add(new Label { Text = "Server auswählen", UseMnemonic = false, AutoSize = true, Font = headingFont, Margin = new Padding(0, 0, 0, 6), ForeColor = SystemInformation.HighContrast ? SystemColors.WindowText : Color.FromArgb(0, 95, 184) }, 0, 0);
        root.Controls.Add(new Label { Text = "Ein Klick auf den Server erstellt dein Meetling.", AutoSize = true, Margin = new Padding(0, 0, 0, 20) }, 0, 1);
        var choices = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true, FlowDirection = FlowDirection.TopDown, WrapContents = false, Margin = Padding.Empty, Padding = new Padding(0, 0, 0, 8) };
        serverPanel = choices;
        root.Controls.Add(choices, 0, 2);
        var duplicates = new HashSet<string>(servers.GroupBy(s => s.Name, StringComparer.CurrentCultureIgnoreCase).Where(g => g.Count() > 1).Select(g => g.Key), StringComparer.CurrentCultureIgnoreCase);
        foreach (var server in servers.OrderBy(s => s.Name, StringComparer.CurrentCultureIgnoreCase).ThenBy(s => s.Id, StringComparer.Ordinal))
        {
            var button = new Button
            {
                Text = "  " + server.Name + (duplicates.Contains(server.Name) ? "\r\n  Server-ID: " + server.Id : ""),
                Tag = server, Height = 76, Width = 480, Margin = new Padding(0, 0, 0, 12), Padding = new Padding(16, 4, 12, 4),
                TextAlign = ContentAlignment.MiddleLeft, ImageAlign = ContentAlignment.MiddleLeft, Image = cube,
                TextImageRelation = TextImageRelation.ImageBeforeText, FlatStyle = FlatStyle.Flat,
                BackColor = SystemColors.Window, ForeColor = SystemColors.WindowText, UseMnemonic = false, AutoEllipsis = true,
                AccessibleName = server.Name, AccessibleDescription = "Konferenzserver, ID " + server.Id,
                TabIndex = buttons.Count
            };
            button.FlatAppearance.BorderSize = 1; button.FlatAppearance.BorderColor = SystemColors.ControlDark;
            button.FlatAppearance.MouseOverBackColor = SystemColors.ControlLight;
            button.FlatAppearance.MouseDownBackColor = SystemColors.Control;
            button.Enter += (s,e) => button.FlatAppearance.BorderColor = SystemColors.Highlight;
            button.Leave += (s,e) => button.FlatAppearance.BorderColor = SystemColors.ControlDark;
            button.Click += (s,e) => { SelectedServer = (Server)button.Tag; DialogResult = DialogResult.OK; Close(); };
            tips.SetToolTip(button, server.Name + "\nServer-ID: " + server.Id);
            choices.Controls.Add(button); buttons.Add(button);
        }
        choices.ClientSizeChanged += (s,e) => { foreach (var button in buttons) button.Width = Math.Max(240, choices.ClientSize.Width - SystemInformation.VerticalScrollBarWidth - 6); };
        var footer = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, FlowDirection = FlowDirection.RightToLeft, Margin = new Padding(0, 12, 0, 0) };
        var cancel = new Button { Text = "Abbrechen", AutoSize = true, MinimumSize = new Size(104, 32), DialogResult = DialogResult.Cancel, FlatStyle = FlatStyle.System };
        footer.Controls.Add(cancel); root.Controls.Add(footer, 0, 3); CancelButton = cancel;
        // No preselected server or default accept button. Click/keyboard activation confirms a specific ID.
    }
    protected override void OnShown(EventArgs e)
    {
        base.OnShown(e);
        FitServerButtons();
    }
    internal void FitServerButtons()
    {
        // Measure after WinForms has applied the host's DPI/font scaling and label wrapping.
        // A fixed height estimate can leave only one card visible in Outlook at higher scaling.
        serverPanel.AutoScrollPosition = Point.Empty;
        foreach (var button in buttons)
            button.Height = Math.Max(button.Height, button.GetPreferredSize(new Size(button.Width, 0)).Height + button.Padding.Vertical);
        PerformLayout(); layout.PerformLayout(); serverPanel.PerformLayout();
        int required = serverPanel.Padding.Vertical + buttons.Sum(b => b.Height + b.Margin.Vertical);
        int chrome = Height - ClientSize.Height;
        int maximum = Math.Max(1, Screen.FromControl(this).WorkingArea.Height - chrome - 32);
        int desired = ClientSize.Height + required - serverPanel.ClientSize.Height;
        ClientSize = new Size(ClientSize.Width, Math.Min(maximum, Math.Max(ClientSize.Height, desired)));
        PerformLayout(); layout.PerformLayout(); serverPanel.PerformLayout();
        // Growing a centered dialog must not put its footer outside the working area.
        var area = Screen.FromControl(this).WorkingArea;
        Top = Math.Max(area.Top, Math.Min(Top, area.Bottom - Height));
        serverPanel.AutoScrollPosition = Point.Empty;
    }
    protected override void Dispose(bool disposing)
    {
        base.Dispose(disposing);
        if (disposing) { tips.Dispose(); cube.Dispose(); bodyFont.Dispose(); headingFont.Dispose(); }
    }
}
