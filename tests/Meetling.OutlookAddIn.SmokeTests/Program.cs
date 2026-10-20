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
            MockMeetingDraft? created = null;
            AppointmentComposer.Open = (application, draft) => created = draft;
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
            foreach (var pair in new[] { new { Name = "TestPlugin", Id = 1 }, new { Name = "OpenSettings", Id = 2 }, new { Name = "CreateMeeting", Id = 4 } })
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
                    if (pair.Id == 4)
                    {
                        Check(created != null, "Meeting callback did not reach native editor boundary");
                        Check(created!.MeetingLink.Host == "meetling.example.invalid" &&
                            created.BodyText.Contains(created.MeetingLink.AbsoluteUri) &&
                            created.BodyText.Contains("Mock") && created.Subject.Contains("Test"), "Mock meeting content");
                        continue;
                    }
                    Check(shown != null && shown.Contains(DiagnosticLog.Version) && shown.Contains("Version: "),
                        "Callback did not reach local dialog sink");
                    Check(shown!.Contains(pair.Id == 1 ? "Der Button funktioniert." : "API-Konfiguration"),
                        "Wrong callback content");
                    if (pair.Id == 2) Check(shown.Contains("diagnose.log"), "Diagnostic path missing");
                }
                finally { Marshal.FreeCoTaskMem(variant); }
            }

            var imageIds = new int[1];
            Marshal.ThrowExceptionForHR(names(dispatchPointer, ref iid, new[] { "GetMeetlingImage" }, 1, 0, imageIds));
            Check(imageIds[0] == 3, "Image callback DISPID");
            var imageArg = Marshal.AllocCoTaskMem(24);
            var imageResult = Marshal.AllocCoTaskMem(24);
            try
            {
                for (var i = 0; i < 24; i++) { Marshal.WriteByte(imageArg, i, 0); Marshal.WriteByte(imageResult, i, 0); }
                Marshal.WriteInt16(imageArg, 9);
                var imageArgs = new DISPPARAMS { rgvarg = imageArg, cArgs = 1 };
                var imageException = new System.Runtime.InteropServices.ComTypes.EXCEPINFO();
                uint imageError;
                Marshal.ThrowExceptionForHR(invoke(dispatchPointer, imageIds[0], ref iid, 0, 1,
                    ref imageArgs, imageResult, ref imageException, out imageError));
                var picture = Marshal.GetObjectForNativeVariant(imageResult);
                Check(picture != null, "Image callback returned null");
                var picturePointer = Marshal.GetIUnknownForObject(picture);
                try
                {
                    var pictureId = new Guid("7BF80981-BF32-101A-8BBB-00AA00300CAB"); // stdole.IPictureDisp
                    IntPtr pictureDispatch;
                    Marshal.ThrowExceptionForHR(Marshal.QueryInterface(picturePointer, ref pictureId, out pictureDispatch));
                    Marshal.Release(pictureDispatch);
                }
                finally { Marshal.Release(picturePointer); }
            }
            finally { VariantClear(imageResult); Marshal.FreeCoTaskMem(imageArg); Marshal.FreeCoTaskMem(imageResult); }
            using (var icon = RibbonImages.RenderCube(64))
            {
                Check(icon.GetPixel(0, 0).A == 0, "Logo corner is not transparent");
                var fill = icon.GetPixel(32, 10);
                Check(fill.R == 60 && fill.G == 179 && fill.B == 159, "Original logo color changed");
                Check(icon.GetPixel(34, 38).R == 255 && icon.GetPixel(53, 38).R == 255, "Logo eyes missing");
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
                var tabs = xml.Root!.Element(ns + "ribbon")!.Element(ns + "tabs")!;
                var tab = tabs.Elements(ns + "tab").Single(t => (string?)t.Attribute("id") == "MeetlingTab");
                Check((string?)tab!.Attribute("id") == "MeetlingTab" && (string?)tab.Attribute("label") == "Meetling", "Tab missing");
                var group = tab.Element(ns + "group");
                Check((string?)group!.Attribute("label") == "Meetling", "Group missing");
                var calendarTab = tabs.Elements(ns + "tab").SingleOrDefault(t => (string?)t.Attribute("idMso") == "TabCalendar");
                Check((calendarTab != null) == (id == "Microsoft.Outlook.Explorer"), "Calendar group leaked into inspector XML");
                if (calendarTab != null)
                {
                    var calendarGroup = calendarTab.Element(ns + "group")!;
                    Check((string?)calendarGroup.Attribute("insertAfterMso") == "GroupCalendarNew", "Calendar placement");
                    var start = calendarGroup.Element(ns + "button")!;
                    Check((string?)start.Attribute("enabled") == "true" &&
                        (string?)start.Attribute("onAction") == "CreateMeeting", "Calendar click binding");
                    Check((string?)start.Attribute("id") == "MeetlingCalendarStart" &&
                        (string?)start.Attribute("getImage") == "GetMeetlingImage", "Calendar logo binding");
                }
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
            Console.WriteLine("PASS: COM activation, default IDispatch GetIDsOfNames/Invoke, image IPictureDisp, mock meeting editor boundary, host interfaces, Ribbon structure and lifecycle. Outlook not tested.");
            return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
        finally
        {
            if (dispatchPointer != IntPtr.Zero) Marshal.Release(dispatchPointer);
            if (instance != null && Marshal.IsComObject(instance)) Marshal.FinalReleaseComObject(instance);
        }
    }

    [DllImport("oleaut32.dll")]
    private static extern int VariantClear(IntPtr variant);

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
