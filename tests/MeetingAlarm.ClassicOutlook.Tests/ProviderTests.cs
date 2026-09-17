using System.Diagnostics;
using System.Runtime.InteropServices;
using MeetingAlarm.Core;
using Xunit;

namespace MeetingAlarm.ClassicOutlook.Tests;

public class ProviderTests
{
    private static readonly DateTimeOffset From = new(2026, 9, 8, 0, 0, 0, TimeSpan.Zero);
    private static readonly Meeting Example = new("classic:test", "Test", From, From.AddHours(1),
        MeetingResponse.Accepted, true, IsLocal: false);

    [Fact]
    public async Task ConstructorDoesNotCreateReaderAndCancelledCallDoesNotConnect()
    {
        var created = 0;
        using var provider = new ClassicOutlookCalendarProvider(() =>
        {
            Interlocked.Increment(ref created);
            return new FakeReader((_, _, _) => [Example]);
        }, TimeSpan.FromSeconds(5));
        await Task.Delay(50);
        Assert.Equal(0, Volatile.Read(ref created));
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            provider.GetMeetingsAsync(From, From.AddDays(1), cancellation.Token));
        Assert.Equal(0, Volatile.Read(ref created));
    }

    [Fact]
    public async Task ReadsAndDisposesOnSameDedicatedStaAndReturnsDetachedSnapshot()
    {
        int factoryThread = 0, readThread = 0, disposalThread = 0;
        ApartmentState apartment = ApartmentState.Unknown;
        var source = new List<Meeting> { Example };
        using var provider = new ClassicOutlookCalendarProvider(() =>
        {
            factoryThread = Environment.CurrentManagedThreadId;
            return new FakeReader((from, to, _) =>
            {
                readThread = Environment.CurrentManagedThreadId;
                apartment = Thread.CurrentThread.GetApartmentState();
                Assert.False(Thread.CurrentThread.IsThreadPoolThread);
                Assert.True(Thread.CurrentThread.IsBackground);
                Assert.Equal(From, from);
                Assert.Equal(From.AddDays(1), to);
                return source;
            }, () => disposalThread = Environment.CurrentManagedThreadId);
        }, TimeSpan.FromSeconds(5));
        var snapshot = await provider.GetMeetingsAsync(From, From.AddDays(1), default);
        source.Clear();
        Assert.Single(snapshot);
        Assert.Equal(ApartmentState.STA, apartment);
        Assert.Equal(factoryThread, readThread);
        Assert.Equal(readThread, disposalThread);
    }

    [Fact]
    public async Task NativeErrorIsSanitizedAndNextCallReusesWorker()
    {
        var threads = new List<int>();
        var calls = 0;
        using var provider = new ClassicOutlookCalendarProvider(() => new FakeReader((_, _, _) =>
        {
            threads.Add(Environment.CurrentManagedThreadId);
            if (++calls == 1) throw new COMException("SECRET SUBJECT token=secret", unchecked((int)0x80070005));
            return [Example];
        }), TimeSpan.FromSeconds(5));
        var error = await Assert.ThrowsAsync<ClassicOutlookException>(() =>
            provider.GetMeetingsAsync(From, From.AddDays(1), default));
        Assert.Contains("80070005", error.Message);
        Assert.IsAssignableFrom<IOException>(error);
        Assert.DoesNotContain("SECRET", error.ToString());
        Assert.DoesNotContain("token=", error.ToString());
        Assert.Null(error.InnerException);
        Assert.Single(await provider.GetMeetingsAsync(From, From.AddDays(1), default));
        Assert.Equal(threads[0], threads[1]);
    }

    [Fact]
    public async Task FailureAfterPartialReadNeverReturnsPartialSuccess()
    {
        var disposed = false;
        using var provider = new ClassicOutlookCalendarProvider(() => new FakeReader((_, _, _) =>
        {
            var partial = new List<Meeting> { Example };
            throw new InvalidCastException("secret data");
        }, () => disposed = true), TimeSpan.FromSeconds(5));
        var error = await Assert.ThrowsAsync<ClassicOutlookException>(() =>
            provider.GetMeetingsAsync(From, From.AddDays(1), default));
        Assert.DoesNotContain("secret", error.ToString());
        Assert.True(disposed);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task CancelOrTimeoutReturnsPromptlyButKeepsSingleSlotUntilNativeReturn(bool timeout)
    {
        using var entered = new ManualResetEventSlim();
        using var release = new ManualResetEventSlim();
        using var cancellation = new CancellationTokenSource();
        var calls = 0;
        var threadIds = new List<int>();
        using var provider = new ClassicOutlookCalendarProvider(() => new FakeReader((_, _, _) =>
        {
            threadIds.Add(Environment.CurrentManagedThreadId);
            if (Interlocked.Increment(ref calls) == 1)
            {
                entered.Set();
                Assert.True(release.Wait(TimeSpan.FromSeconds(10)));
            }
            return [Example];
        }), timeout ? TimeSpan.FromMilliseconds(200) : TimeSpan.FromSeconds(5));
        try
        {
            var first = provider.GetMeetingsAsync(From, From.AddDays(1), cancellation.Token);
            Assert.True(entered.Wait(TimeSpan.FromSeconds(5)));
            if (timeout)
            {
                var error = await Assert.ThrowsAsync<ClassicOutlookException>(() => first.WaitAsync(TimeSpan.FromSeconds(5)));
                Assert.IsAssignableFrom<IOException>(error);
                Assert.Contains("timeout", error.Message);
                Assert.Contains("Resolve any Outlook prompt", error.Message);
            }
            else
            {
                cancellation.Cancel();
                await Assert.ThrowsAnyAsync<OperationCanceledException>(() => first.WaitAsync(TimeSpan.FromSeconds(5)));
            }
            var rejected = Enumerable.Range(0, 30)
                .Select(_ => provider.GetMeetingsAsync(From, From.AddDays(1), default)).ToArray();
            foreach (var task in rejected)
                Assert.Contains("previous request", (await Assert.ThrowsAsync<ClassicOutlookException>(() => task)).Message);
            Assert.Equal(1, Volatile.Read(ref calls));
            release.Set();
            IReadOnlyList<Meeting>? next = null;
            for (var attempt = 0; attempt < 100 && next is null; attempt++)
            {
                await Task.Delay(10);
                try { next = await provider.GetMeetingsAsync(From, From.AddDays(1), default); }
                catch (ClassicOutlookException error) when (error.Message.Contains("previous request")) { }
            }
            Assert.NotNull(next);
            Assert.Single(next);
            Assert.Equal(threadIds[0], threadIds[1]);
        }
        finally { release.Set(); }
    }

    [Fact]
    public async Task DisposeIsNonBlockingAndCleanupWaitsOnOwningSta()
    {
        using var entered = new ManualResetEventSlim();
        using var release = new ManualResetEventSlim();
        using var disposed = new ManualResetEventSlim();
        int readThread = 0, disposalThread = 0;
        var provider = new ClassicOutlookCalendarProvider(() => new FakeReader((_, _, _) =>
        {
            readThread = Environment.CurrentManagedThreadId;
            entered.Set();
            Assert.True(release.Wait(TimeSpan.FromSeconds(10)));
            return [Example];
        }, () =>
        {
            disposalThread = Environment.CurrentManagedThreadId;
            disposed.Set();
        }), TimeSpan.FromSeconds(5));
        try
        {
            var task = provider.GetMeetingsAsync(From, From.AddDays(1), default);
            Assert.True(entered.Wait(TimeSpan.FromSeconds(5)));
            var stopwatch = Stopwatch.StartNew();
            provider.Dispose();
            Assert.True(stopwatch.Elapsed < TimeSpan.FromSeconds(1));
            provider.Dispose();
            await Assert.ThrowsAsync<ObjectDisposedException>(() => task);
            await Assert.ThrowsAsync<ObjectDisposedException>(() =>
                provider.GetMeetingsAsync(From, From.AddDays(1), default));
            Assert.False(disposed.IsSet);
            release.Set();
            Assert.True(disposed.Wait(TimeSpan.FromSeconds(5)));
            Assert.Equal(readThread, disposalThread);
            await provider.ShutdownCompletion.WaitAsync(TimeSpan.FromSeconds(5));
        }
        finally
        {
            release.Set();
            provider.Dispose();
        }
    }

    [Fact]
    public async Task LifetimeCleanupWaitsForCancellationCallbacksWithoutBlockingDispose()
    {
        using var entered = new ManualResetEventSlim();
        using var callbackEntered = new ManualResetEventSlim();
        using var callbackRelease = new ManualResetEventSlim();
        var provider = new ClassicOutlookCalendarProvider(() => new FakeReader((_, _, token) =>
        {
            using var registration = token.Register(() =>
            {
                callbackEntered.Set();
                Assert.True(callbackRelease.Wait(TimeSpan.FromSeconds(10)));
            });
            entered.Set();
            Assert.True(callbackEntered.Wait(TimeSpan.FromSeconds(5)));
            return [Example];
        }), TimeSpan.FromSeconds(5));
        try
        {
            var task = provider.GetMeetingsAsync(From, From.AddDays(1), default);
            Assert.True(entered.Wait(TimeSpan.FromSeconds(5)));
            var stopwatch = Stopwatch.StartNew();
            provider.Dispose();
            Assert.True(stopwatch.Elapsed < TimeSpan.FromSeconds(1));
            Assert.True(callbackEntered.Wait(TimeSpan.FromSeconds(5)));
            Assert.False(provider.ShutdownCompletion.IsCompleted);
            await Assert.ThrowsAsync<ObjectDisposedException>(() => task);
            callbackRelease.Set();
            await provider.ShutdownCompletion.WaitAsync(TimeSpan.FromSeconds(5));
            provider.Dispose();
        }
        finally
        {
            callbackRelease.Set();
            provider.Dispose();
        }
    }

    [Fact]
    public async Task InvalidWindowsNeverConnect()
    {
        var created = false;
        using var provider = new ClassicOutlookCalendarProvider(() =>
        {
            created = true;
            return new FakeReader((_, _, _) => []);
        }, TimeSpan.FromSeconds(5));
        await Assert.ThrowsAsync<ArgumentOutOfRangeException>(() =>
            provider.GetMeetingsAsync(From, From, default));
        await Assert.ThrowsAsync<ArgumentOutOfRangeException>(() =>
            provider.GetMeetingsAsync(From, From.AddDays(367), default));
        Assert.False(created);
    }

    private sealed class FakeReader(
        Func<DateTimeOffset, DateTimeOffset, CancellationToken, IReadOnlyList<Meeting>> read,
        Action? dispose = null) : ICalendarSnapshotReader
    {
        public IReadOnlyList<Meeting> Read(DateTimeOffset from, DateTimeOffset to, CancellationToken token) =>
            read(from, to, token);
        public void Dispose() => dispose?.Invoke();
    }
}
