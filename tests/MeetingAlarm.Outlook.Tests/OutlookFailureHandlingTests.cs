using System.Security;
using System.Security.Cryptography;
using System.Text.Json;
using Microsoft.Identity.Client;
using Microsoft.Identity.Client.Extensions.Msal;
using Xunit;

namespace MeetingAlarm.Outlook.Tests;

public sealed class OutlookFailureHandlingTests
{
    public static TheoryData<Exception> ExpectedCacheFailures => new()
    {
        new IOException("private cache path"),
        new UnauthorizedAccessException("private cache path"),
        new SecurityException("private security detail"),
        new CryptographicException("private encrypted data"),
        new MsalCachePersistenceException("private cache path"),
        new JsonException("private cache data")
    };

    public static TheoryData<Exception> UnexpectedCacheFailures => new()
    {
        new Exception("unexpected"),
        new NullReferenceException("programming bug"),
        new InvalidOperationException("programming bug"),
        new ArgumentException("programming bug"),
        new NotSupportedException("programming bug"),
        new IndexOutOfRangeException("programming bug"),
        new ObjectDisposedException("programming bug"),
        new OperationCanceledException(),
        new OutOfMemoryException()
    };

    [Theory]
    [MemberData(nameof(ExpectedCacheFailures))]
    public void RecognizesOnlyExpectedCacheFailures(Exception failure) =>
        Assert.True(OutlookCalendarProvider.IsExpectedCacheFailure(failure));

    [Theory]
    [MemberData(nameof(UnexpectedCacheFailures))]
    public void DoesNotClassifyProgrammingErrorsOrCancellationAsCacheFailures(Exception failure) =>
        Assert.False(OutlookCalendarProvider.IsExpectedCacheFailure(failure));

    [Theory]
    [MemberData(nameof(ExpectedCacheFailures))]
    public void ExpectedCleanupFailureIsSanitizedAndAllResourcesAreReleased(Exception failure)
    {
        using var http = new TrackingDisposable();
        var traceClosed = false;
        var error = Assert.Throws<OutlookCalendarException>(() => OutlookCalendarProvider.ReleaseResources(
            () => throw failure, http, () => traceClosed = true));
        Assert.True(http.Disposed);
        Assert.True(traceClosed);
        Assert.DoesNotContain("private", error.ToString());
        Assert.Null(error.InnerException);
    }

    [Theory]
    [MemberData(nameof(UnexpectedCacheFailures))]
    public void UnexpectedCleanupFailurePropagatesButAllResourcesAreReleased(Exception failure)
    {
        using var http = new TrackingDisposable();
        var traceClosed = false;
        var observed = Record.Exception(() => OutlookCalendarProvider.ReleaseResources(
            () => throw failure, http, () => traceClosed = true));
        Assert.Same(failure, observed);
        Assert.True(http.Disposed);
        Assert.True(traceClosed);
    }

    [Fact]
    public void HttpDisposeFailureStillClosesTraceAndPropagatesProgrammingBug()
    {
        var failure = new InvalidOperationException("synthetic programming bug");
        var http = new TrackingDisposable(failure);
        var cacheUnregistered = false;
        var traceClosed = false;
        var observed = Record.Exception(() => OutlookCalendarProvider.ReleaseResources(
            () => cacheUnregistered = true, http, () => traceClosed = true));
        Assert.Same(failure, observed);
        Assert.True(cacheUnregistered);
        Assert.True(http.Disposed);
        Assert.True(traceClosed);
    }

    [Fact]
    public void MsalCleanupFailureRemainsSanitizedMsalException()
    {
        using var http = new TrackingDisposable();
        var traceClosed = false;
        var error = Assert.Throws<MsalClientException>(() => OutlookCalendarProvider.ReleaseResources(
            () => throw new MsalClientException("synthetic", "private error detail"),
            http, () => traceClosed = true));
        Assert.True(http.Disposed);
        Assert.True(traceClosed);
        Assert.DoesNotContain("private", error.ToString());
        Assert.Null(error.InnerException);
    }

    [Fact]
    public async Task CancelledCreationDoesNotCreateCacheResources()
    {
        var directory = Path.Combine(AppContext.BaseDirectory, "test-cache-" + Guid.NewGuid().ToString("N"));
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => OutlookCalendarProvider.CreateAsync(
            Guid.NewGuid().ToString("D"), "organizations", directory, cancellation.Token));
        Assert.False(Directory.Exists(directory));
    }

    private sealed class TrackingDisposable(Exception? failure = null) : IDisposable
    {
        public bool Disposed { get; private set; }

        public void Dispose()
        {
            Disposed = true;
            if (failure is not null)
                throw failure;
        }
    }
}
