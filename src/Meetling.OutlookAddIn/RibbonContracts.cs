using System.Runtime.InteropServices;

namespace Meetling.OutlookAddIn.Interop;

[ComImport]
[ComVisible(true)]
[Guid("000C0396-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsDual)]
public interface IRibbonExtensibility
{
    [DispId(1)]
    [return: MarshalAs(UnmanagedType.BStr)]
    string GetCustomUI([MarshalAs(UnmanagedType.BStr)] string ribbonId);
}

[ComImport]
[ComVisible(true)]
[Guid("000C0395-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsDual)]
public interface IRibbonControl
{
    [DispId(1)]
    string Id { [return: MarshalAs(UnmanagedType.BStr)] get; }
    [DispId(2)]
    object Context { [return: MarshalAs(UnmanagedType.IDispatch)] get; }
    [DispId(3)]
    string Tag { [return: MarshalAs(UnmanagedType.BStr)] get; }
}


[ComImport, Guid("000C03A7-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
public interface IRibbonUI
{
    [DispId(1)] void Invalidate();
}
