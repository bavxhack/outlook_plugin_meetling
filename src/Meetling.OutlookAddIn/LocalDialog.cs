using System;
using System.Windows.Forms;

namespace Meetling.OutlookAddIn;

internal static class LocalDialog
{
    // The Windows smoke harness replaces only this managed sink before COM activation.
    // There is no public or COM-visible switch that can disable user dialogs.
    internal static Action<string> Show = text =>
        MessageBox.Show(text, "Meetling", MessageBoxButtons.OK, MessageBoxIcon.Information);
}
