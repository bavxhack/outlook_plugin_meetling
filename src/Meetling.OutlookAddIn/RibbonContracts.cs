using System.Runtime.InteropServices;

namespace Meetling.OutlookAddIn.Interop;

[ComImport]
[Guid("000C0396-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsDual)]
public interface IRibbonExtensibility
{
    [DispId(1)]
    [return: MarshalAs(UnmanagedType.BStr)]
    string GetCustomUI([MarshalAs(UnmanagedType.BStr)] string ribbonId);
}

[ComImport]
[Guid("000C0395-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsDual)]
public interface IRibbonControl
{
}
