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
        using (var form = new Form { Text = "Meetling – Server auswählen", ClientSize = new Size(500, 115), StartPosition = FormStartPosition.CenterScreen })
        {
            var combo = new ComboBox { Left = 15, Top = 15, Width = 470, DropDownStyle = ComboBoxStyle.DropDownList };
            combo.Items.AddRange(servers.OrderBy(s => s.Name).Cast<object>().ToArray()); combo.SelectedIndex = -1;
            var ok = new Button { Text = "Bestätigen", Left = 260, Top = 65, Width = 110, Enabled = false, DialogResult = DialogResult.OK };
            combo.SelectedIndexChanged += (s,e) => ok.Enabled = combo.SelectedItem is Server;
            form.Controls.Add(combo); form.Controls.Add(ok); form.Controls.Add(new Button { Text = "Abbrechen", Left = 380, Top = 65, DialogResult = DialogResult.Cancel });
            return form.ShowDialog() == DialogResult.OK ? combo.SelectedItem as Server : null;
        }
    }
}
