using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
namespace Meetling.OutlookAddIn;
internal sealed class UiDispatcher : IDisposable
{
    private readonly Control control;
    private readonly int thread = Thread.CurrentThread.ManagedThreadId;
    internal UiDispatcher() { control = new Control(); _ = control.Handle; }
    internal Task InvokeAsync(Action action)
    {
        var result = new TaskCompletionSource<bool>();
        void Run() { try { action(); result.SetResult(true); } catch (Exception e) { result.SetException(e); } }
        if (control.IsDisposed) { result.SetCanceled(); return result.Task; }
        if (Thread.CurrentThread.ManagedThreadId == thread) Run(); else control.BeginInvoke((Action)Run);
        return result.Task;
    }
    public void Dispose() => control.Dispose();
}
