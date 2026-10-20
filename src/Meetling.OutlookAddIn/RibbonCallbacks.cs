using System.Runtime.InteropServices;

namespace Meetling.OutlookAddIn;

[ComVisible(true)]
[Guid("6D6E490A-EC2B-4F53-9AF3-32690E634D19")]
[InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
// Office passes IRibbonControl as IDispatch. No registered Office typelib is needed
// to describe these callbacks; the baseline deliberately never dereferences control.
public interface IMeetlingRibbonCallbacks
{
    [DispId(1)]
    void TestPlugin([In, MarshalAs(UnmanagedType.IDispatch)] object control);

    [DispId(2)]
    void OpenSettings([In, MarshalAs(UnmanagedType.IDispatch)] object control);
    [DispId(3)]
    [return: MarshalAs(UnmanagedType.IDispatch)]
    object? GetMeetlingImage([In, MarshalAs(UnmanagedType.IDispatch)] object control);
    [DispId(4)]
    void CreateMeeting([In, MarshalAs(UnmanagedType.IDispatch)] object control);
}
