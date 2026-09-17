using System.Diagnostics;

namespace MeetingAlarm.Outlook;

internal sealed class CacheFailureMonitor : TraceListener
{
    private int _failed;

    // MsalCacheHelper logs and swallows some persistence failures. Latch the error
    // without retaining its potentially sensitive text, then fail the public call.
    public override void TraceEvent(TraceEventCache? eventCache, string source, TraceEventType eventType,
        int id, string? message) => Observe(eventType);

    public override void TraceEvent(TraceEventCache? eventCache, string source, TraceEventType eventType,
        int id, string? format, params object?[]? args) => Observe(eventType);

    public override void Write(string? message) { }
    public override void WriteLine(string? message) { }

    private void Observe(TraceEventType eventType)
    {
        if (eventType is TraceEventType.Error or TraceEventType.Critical)
            Interlocked.Exchange(ref _failed, 1);
    }

    internal void ThrowIfFailed()
    {
        if (Volatile.Read(ref _failed) != 0)
            throw new OutlookCalendarException(
                "Outlook's encrypted token cache failed to read or persist. Recreate the connector after correcting the cache.");
    }
}
