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
    [DispId(5)] void SignIn([In, MarshalAs(UnmanagedType.IDispatch)] object control);
    [DispId(6)] void SignOut([In, MarshalAs(UnmanagedType.IDispatch)] object control);
    [DispId(7)] void RefreshServers([In, MarshalAs(UnmanagedType.IDispatch)] object control);
    [DispId(8)] void ReloadConfiguration([In, MarshalAs(UnmanagedType.IDispatch)] object control);
    [DispId(9)] void RibbonLoaded([In, MarshalAs(UnmanagedType.IDispatch)] object ui);
    [DispId(10)] string GetStatus([In, MarshalAs(UnmanagedType.IDispatch)] object control);
    [DispId(11)] bool GetCreateEnabled([In, MarshalAs(UnmanagedType.IDispatch)] object control);
    [DispId(12)] bool GetLoginEnabled([In, MarshalAs(UnmanagedType.IDispatch)] object control);
    [DispId(13)] bool GetLogoutEnabled([In, MarshalAs(UnmanagedType.IDispatch)] object control);
}
