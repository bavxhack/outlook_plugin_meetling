using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using Meetling.OutlookAddIn;
using Meetling.OutlookAddIn.Interop;
using Extensibility;

internal static class Program
{
    private static void Check(bool condition, string message)
    {
        if (!condition) throw new InvalidOperationException(message);
    }

    [STAThread]
    private static int Main()
    {
        object? instance = null;
        IntPtr dispatchPointer = IntPtr.Zero;
        try
        {
            string? shown = null;
            LocalDialog.Show = text => shown = text; // CI must never open a modal dialog.
            var type = Type.GetTypeFromProgID("Meetling.OutlookAddIn", true);
            instance = Activator.CreateInstance(type!);
            Check(instance != null, "COM activation returned null");
            // Acquire default IDispatch from the activated CCW; call raw vtable methods.
            dispatchPointer = Marshal.GetIDispatchForObject(instance!);
            var table = Marshal.ReadIntPtr(dispatchPointer);
            var names = (GetIds)Marshal.GetDelegateForFunctionPointer(
                Marshal.ReadIntPtr(table, 5 * IntPtr.Size), typeof(GetIds));
            var invoke = (Invoke)Marshal.GetDelegateForFunctionPointer(
                Marshal.ReadIntPtr(table, 6 * IntPtr.Size), typeof(Invoke));
            var iid = Guid.Empty;
            foreach (var pair in new[] { new { Name = "TestPlugin", Id = 1 }, new { Name = "OpenSettings", Id = 2 } })
            {
                var ids = new int[1];
                Marshal.ThrowExceptionForHR(names(dispatchPointer, ref iid, new[] { pair.Name }, 1, 0, ids));
                Check(ids[0] == pair.Id, "Unexpected callback DISPID");
                var variant = Marshal.AllocCoTaskMem(24); // VARIANT: 16 bytes x86, 24 bytes x64.
                try
                {
                    for (int i = 0; i < 24; i++) Marshal.WriteByte(variant, i, 0); // Initialize empty VARIANT.
                    Marshal.WriteInt16(variant, 9); // VT_DISPATCH with null pointer -> null IRibbonControl
                    var args = new DISPPARAMS { rgvarg = variant, cArgs = 1 };
                    var exception = new System.Runtime.InteropServices.ComTypes.EXCEPINFO();
                    uint error;
                    shown = null;
                    Marshal.ThrowExceptionForHR(invoke(dispatchPointer, ids[0], ref iid, 0, 1,
                        ref args, IntPtr.Zero, ref exception, out error));
                    Check(shown != null && shown.Contains(DiagnosticLog.Version) && shown.Contains("Version: "),
                        "Callback did not reach local dialog sink");
                    Check(shown!.Contains(pair.Id == 1 ? "Der Button funktioniert." : "API-Konfiguration"),
                        "Wrong callback content");
                    if (pair.Id == 2) Check(shown.Contains("diagnose.log"), "Diagnostic path missing");
                }
                finally { Marshal.FreeCoTaskMem(variant); }
            }

            foreach (var contract in new[] { typeof(IDTExtensibility2), typeof(IRibbonExtensibility), typeof(IMeetlingRibbonCallbacks) })
            {
                var contractId = contract.GUID;
                IntPtr pointer;
                Marshal.ThrowExceptionForHR(Marshal.QueryInterface(dispatchPointer, ref contractId, out pointer));
                Marshal.Release(pointer);
            }
            Check(instance is IDTExtensibility2, "IDTExtensibility2 unavailable");
            Check(instance is IRibbonExtensibility, "IRibbonExtensibility QueryInterface failed");
            var ribbon = (IRibbonExtensibility)instance!;
            XNamespace ns = "http://schemas.microsoft.com/office/2009/07/customui";
            foreach (var id in new[] { "Microsoft.Outlook.Explorer", "Microsoft.Outlook.Appointment", "Microsoft.Outlook.Mail.Compose" })
            {
                var xml = XDocument.Parse(ribbon.GetCustomUI(id)); // before OnConnection
                var tab = xml.Root!.Element(ns + "ribbon")!.Element(ns + "tabs")!.Element(ns + "tab");
                Check((string?)tab!.Attribute("id") == "MeetlingTab" && (string?)tab.Attribute("label") == "Meetling", "Tab missing");
                var group = tab.Element(ns + "group");
                Check((string?)group!.Attribute("label") == "Meetling", "Group missing");
                var buttons = group.Elements(ns + "button").ToArray();
                Check(buttons.Length == 3, "Expected three buttons");
                Check((string?)buttons[0].Attribute("onAction") == "TestPlugin", "Test callback missing");
                Check((string?)buttons[1].Attribute("onAction") == "OpenSettings", "Settings callback missing");
                Check((string?)buttons[2].Attribute("enabled") == "false" && buttons[2].Attribute("onAction") == null, "Conference must be inert");
            }
            Check(ribbon.GetCustomUI("unknown") == "" && ribbon.GetCustomUI(null!) == "", "Unknown Ribbon ID");
            var lifecycle = (IDTExtensibility2)instance!;
            Array custom = new object[0];
            lifecycle.OnConnection(null!, ext_ConnectMode.ext_cm_Startup, null!, ref custom);
            lifecycle.OnStartupComplete(ref custom);
            lifecycle.OnAddInsUpdate(ref custom);
            lifecycle.OnBeginShutdown(ref custom);
            lifecycle.OnDisconnection(ext_DisconnectMode.ext_dm_HostShutdown, ref custom);
            Console.WriteLine("PASS: COM activation, default IDispatch GetIDsOfNames/Invoke, host interfaces, Ribbon structure and lifecycle. Outlook not tested.");
            return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
        finally
        {
            if (dispatchPointer != IntPtr.Zero) Marshal.Release(dispatchPointer);
            if (instance != null && Marshal.IsComObject(instance)) Marshal.FinalReleaseComObject(instance);
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct DISPPARAMS
    {
        public IntPtr rgvarg;
        public IntPtr rgdispidNamedArgs;
        public uint cArgs;
        public uint cNamedArgs;
    }

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int GetIds(IntPtr self, ref Guid iid,
        [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr, SizeParamIndex = 3)] string[] names,
        uint count, uint lcid, [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 3)] int[] ids);

    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int Invoke(IntPtr self, int id, ref Guid iid, uint lcid, ushort flags,
        ref DISPPARAMS args, IntPtr result, ref System.Runtime.InteropServices.ComTypes.EXCEPINFO exception, out uint error);
}
