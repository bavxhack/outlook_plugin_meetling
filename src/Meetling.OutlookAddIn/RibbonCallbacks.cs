using System.Runtime.InteropServices;
using Meetling.OutlookAddIn.Interop;

namespace Meetling.OutlookAddIn;

[ComVisible(true)]
[Guid("6D6E490A-EC2B-4F53-9AF3-32690E634D19")]
[InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
public interface IMeetlingRibbonCallbacks
{
    [DispId(1)]
    void TestPlugin([In, MarshalAs(UnmanagedType.Interface)] IRibbonControl control);

    [DispId(2)]
    void OpenSettings([In, MarshalAs(UnmanagedType.Interface)] IRibbonControl control);
}
