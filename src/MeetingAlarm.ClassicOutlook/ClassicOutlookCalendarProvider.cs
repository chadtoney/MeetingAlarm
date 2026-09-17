using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Microsoft.CSharp.RuntimeBinder;
using MeetingAlarm.Core;

namespace MeetingAlarm.ClassicOutlook;

/// <summary>Read-only access to the default calendar of the classic Outlook default profile.</summary>
public sealed class ClassicOutlookCalendarProvider : ICalendarProvider, IDisposable
{
    private readonly object gate = new();
    private readonly AutoResetEvent wake = new(false);
    private readonly CancellationTokenSource lifetime = new();
    private readonly TaskCompletionSource workerStopped = new(TaskCreationOptions.RunContinuationsAsynchronously);
    private readonly Func<ICalendarSnapshotReader> readerFactory;
    private readonly TimeSpan timeout;
    private readonly Thread worker;
    private Work? occupied;
    private bool closed;
    private bool stopped;
    private string? workerFailure;
    private Task? shutdownCompletion;

    internal Task ShutdownCompletion
    {
        get { lock (gate) return shutdownCompletion ?? Task.CompletedTask; }
    }

    public ClassicOutlookCalendarProvider()
        : this(() => new ComSnapshotReader(), TimeSpan.FromSeconds(45)) { }

    internal ClassicOutlookCalendarProvider(Func<ICalendarSnapshotReader> readerFactory, TimeSpan timeout)
    {
        ArgumentNullException.ThrowIfNull(readerFactory);
        if (timeout <= TimeSpan.Zero || timeout > TimeSpan.FromMinutes(10))
            throw new ArgumentOutOfRangeException(nameof(timeout));
        this.readerFactory = readerFactory;
        this.timeout = timeout;
        worker = new Thread(Run) { IsBackground = true, Name = "MeetingAlarm classic Outlook STA" };
        worker.SetApartmentState(ApartmentState.STA);
        worker.Start();
    }

    public static bool IsInstalled => Type.GetTypeFromProgID("Outlook.Application") is not null;

    public Task<IReadOnlyList<Meeting>> GetMeetingsAsync(
        DateTimeOffset from, DateTimeOffset to, CancellationToken cancellationToken)
    {
        lock (gate)
        {
            if (closed)
                return Task.FromException<IReadOnlyList<Meeting>>(new ObjectDisposedException(nameof(ClassicOutlookCalendarProvider)));
            if (cancellationToken.IsCancellationRequested)
                return Task.FromCanceled<IReadOnlyList<Meeting>>(cancellationToken);
            if (to <= from || to - from > TimeSpan.FromDays(366) ||
                from.Year < 1602 || to.Year > 9998)
                return Task.FromException<IReadOnlyList<Meeting>>(new ArgumentOutOfRangeException(
                    nameof(to), "Use an increasing calendar window of at most 366 days, within years 1602–9998."));
            if (workerFailure is not null)
                return Task.FromException<IReadOnlyList<Meeting>>(new ClassicOutlookException(workerFailure));
            if (occupied is not null)
                return Task.FromException<IReadOnlyList<Meeting>>(new ClassicOutlookException(
                    "Classic Outlook is still processing a previous request. Close any Outlook prompt and retry after it finishes."));
            occupied = new Work(from, to, cancellationToken, lifetime.Token, timeout);
            if (!stopped) wake.Set();
            return occupied.Completion.Task;
        }
    }

    public void Dispose()
    {
        lock (gate)
        {
            if (closed) return;
            closed = true;
            occupied?.Completion.TrySetException(new ObjectDisposedException(nameof(ClassicOutlookCalendarProvider)));
            // Do not wait for native COM or execute arbitrary token callbacks on the GUI thread.
            shutdownCompletion = DisposeLifetimeAfterShutdownAsync(lifetime.CancelAsync());
            if (!stopped) wake.Set();
        }
    }

    private async Task DisposeLifetimeAfterShutdownAsync(Task cancellation)
    {
        await workerStopped.Task.ConfigureAwait(false);
        try
        {
            await cancellation.ConfigureAwait(false);
        }
        finally
        {
            lifetime.Dispose();
        }
    }

    private void Run()
    {
        var initialized = false;
        try
        {
            Marshal.ThrowExceptionForHR(Native.OleInitialize(IntPtr.Zero));
            initialized = true;
            while (true)
            {
                Work? work;
                lock (gate)
                {
                    work = occupied;
                    if (closed && work is null) break;
                }
                if (work is null)
                {
                    Native.WaitAndPump(wake);
                    continue;
                }
                IReadOnlyList<Meeting>? snapshot = null;
                Exception? error = null;
                try
                {
                    work.Token.ThrowIfCancellationRequested();
                    using var reader = readerFactory()
                        ?? throw new ClassicOutlookException("The classic Outlook snapshot reader is unavailable.");
                    snapshot = reader.Read(work.From, work.To, work.Token).ToArray();
                    work.Token.ThrowIfCancellationRequested();
                }
                catch (OperationCanceledException) when (work.Token.IsCancellationRequested)
                {
                    work.CompleteCancellation();
                }
                catch (Exception exception) when (IsExpectedFailure(exception))
                {
                    error = Sanitize(exception);
                }
                lock (gate)
                {
                    occupied = null;
                    if (error is not null) work.Completion.TrySetException(error);
                    else if (snapshot is not null) work.Completion.TrySetResult(snapshot);
                }
                work.Dispose();
            }
        }
        catch (Exception exception) when (IsExpectedFailure(exception))
        {
            lock (gate)
            {
                workerFailure = Sanitize(exception).Message;
                occupied?.Completion.TrySetException(new ClassicOutlookException(workerFailure));
                occupied?.Dispose();
                occupied = null;
            }
        }
        finally
        {
            if (initialized) Native.OleUninitialize();
            lock (gate)
            {
                stopped = true;
                wake.Dispose();
            }
            workerStopped.TrySetResult();
        }
    }

    private static bool IsExpectedFailure(Exception exception) => exception is
        ClassicOutlookException or COMException or InvalidComObjectException or
        RuntimeBinderException or InvalidCastException or FormatException or OverflowException or
        ArgumentException or InvalidOperationException or NotSupportedException or
        UnauthorizedAccessException or Win32Exception or RegexMatchTimeoutException or
        System.Reflection.TargetInvocationException;

    private static Exception Sanitize(Exception exception)
    {
        if (exception is ClassicOutlookException) return exception;
        // Deliberately omit the native message, inner exception, subject, body, and links.
        return new ClassicOutlookException(
            $"Classic Outlook could not read a complete calendar snapshot (error 0x{exception.HResult:X8}). " +
            "Check that classic Outlook has an accessible default profile and resolve any Outlook or programmatic-access prompt, then retry.");
    }

    private sealed class Work : IDisposable
    {
        private readonly CancellationToken caller;
        private readonly CancellationToken shutdown;
        private readonly CancellationTokenSource deadline;
        private readonly CancellationTokenSource linked;
        private readonly CancellationTokenRegistration registration;
        internal readonly TaskCompletionSource<IReadOnlyList<Meeting>> Completion =
            new(TaskCreationOptions.RunContinuationsAsynchronously);
        internal DateTimeOffset From { get; }
        internal DateTimeOffset To { get; }
        internal CancellationToken Token => linked.Token;

        internal Work(DateTimeOffset from, DateTimeOffset to, CancellationToken caller,
            CancellationToken shutdown, TimeSpan timeout)
        {
            From = from;
            To = to;
            this.caller = caller;
            this.shutdown = shutdown;
            deadline = new CancellationTokenSource(timeout);
            linked = CancellationTokenSource.CreateLinkedTokenSource(caller, shutdown, deadline.Token);
            registration = linked.Token.Register(CompleteCancellation);
        }

        internal void CompleteCancellation()
        {
            if (shutdown.IsCancellationRequested)
                Completion.TrySetException(new ObjectDisposedException(nameof(ClassicOutlookCalendarProvider)));
            else if (caller.IsCancellationRequested)
                Completion.TrySetCanceled(caller);
            else
                Completion.TrySetException(new ClassicOutlookException(
                    "Classic Outlook did not finish within the connection/read timeout. Resolve any Outlook prompt. " +
                    "The existing request must finish before another request can run."));
        }

        public void Dispose()
        {
            registration.Dispose();
            linked.Dispose();
            deadline.Dispose();
        }
    }

    private static class Native
    {
        [DllImport("ole32.dll")]
        internal static extern int OleInitialize(IntPtr reserved);
        [DllImport("ole32.dll")]
        internal static extern void OleUninitialize();
        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint MsgWaitForMultipleObjectsEx(uint count, IntPtr[] handles,
            uint milliseconds, uint wakeMask, uint flags);
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool PeekMessage(out Message message, IntPtr window, uint min, uint max, uint remove);
        [DllImport("user32.dll")]
        private static extern bool TranslateMessage(ref Message message);
        [DllImport("user32.dll")]
        private static extern IntPtr DispatchMessage(ref Message message);

        [StructLayout(LayoutKind.Sequential)]
        private struct Message
        {
            internal IntPtr Window;
            internal uint Id;
            internal UIntPtr WParam;
            internal IntPtr LParam;
            internal uint Time;
            internal int X;
            internal int Y;
            internal uint Private;
        }

        internal static void WaitAndPump(AutoResetEvent signal)
        {
            var result = MsgWaitForMultipleObjectsEx(1, [signal.SafeWaitHandle.DangerousGetHandle()],
                uint.MaxValue, 0x04FF, 0x0004);
            if (result == uint.MaxValue) throw new Win32Exception(Marshal.GetLastWin32Error());
            // Bound each drain so a busy message queue cannot starve disposal or queued work.
            for (var i = 0; i < 64 && PeekMessage(out var message, IntPtr.Zero, 0, 0, 1); i++)
            {
                TranslateMessage(ref message);
                DispatchMessage(ref message);
            }
        }
    }
}
