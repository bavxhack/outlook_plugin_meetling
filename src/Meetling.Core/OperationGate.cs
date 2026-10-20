namespace Meetling.Core;
// The caller owns the UI thread; reset invalidates a generation of work.
public sealed class OperationGate
{
    public bool Busy { get; private set; }
    public bool TryEnter() { if (Busy) return false; Busy = true; return true; }
    public void Release() => Busy = false;
}
