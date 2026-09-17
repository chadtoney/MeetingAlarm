using System.Diagnostics;
using Microsoft.Identity.Client;
using Xunit;

namespace MeetingAlarm.Outlook.Tests;

public sealed class OutlookConfigurationTests
{
    [Fact]
    public void ErrorsMatchHostSafeBoundaryTypes()
    {
        Assert.IsAssignableFrom<HttpRequestException>(new OutlookCalendarException("sanitized"));
        Assert.IsAssignableFrom<MsalException>(new OutlookSignInRequiredException());
    }

    [Theory]
    [InlineData("common")]
    [InlineData("organizations")]
    [InlineData("consumers")]
    [InlineData("181e0386-29b8-4ee4-85c8-228065652ea7")]
    public void AllowsOnlyApprovedTenantForms(string tenant) =>
        OutlookCalendarProvider.ValidateConfiguration(Guid.NewGuid().ToString("D"), tenant, AppContext.BaseDirectory);

    [Theory]
    [InlineData("https://evil.example/tenant")]
    [InlineData("contoso.example")]
    [InlineData("")]
    [InlineData("00000000-0000-0000-0000-000000000000")]
    [InlineData("common/other")]
    public void RejectsInvalidTenant(string tenant) =>
        Assert.Throws<ArgumentException>(() =>
            OutlookCalendarProvider.ValidateConfiguration(Guid.NewGuid().ToString("D"), tenant, AppContext.BaseDirectory));

    [Theory]
    [InlineData("")]
    [InlineData("not-an-app-uuid")]
    [InlineData("00000000-0000-0000-0000-000000000000")]
    public void RejectsInvalidClientId(string clientId) =>
        Assert.Throws<ArgumentException>(() =>
            OutlookCalendarProvider.ValidateConfiguration(clientId, "organizations", AppContext.BaseDirectory));

    [Fact]
    public void RejectsRelativeCacheLocation() =>
        Assert.Throws<ArgumentException>(() =>
            OutlookCalendarProvider.ValidateConfiguration(Guid.NewGuid().ToString("D"), "organizations", "relative"));

    [Fact]
    public void CacheFailureMonitorLatchesAndDoesNotRetainErrorDetails()
    {
        using var monitor = new CacheFailureMonitor();
        monitor.TraceEvent(null, "synthetic", TraceEventType.Information, 0, "ignored");
        monitor.ThrowIfFailed();
        monitor.TraceEvent(null, "synthetic", TraceEventType.Error, 0, "secret-token cache-path");
        var error = Assert.Throws<OutlookCalendarException>(monitor.ThrowIfFailed);
        Assert.DoesNotContain("secret-token", error.ToString());
        monitor.TraceEvent(null, "synthetic", TraceEventType.Information, 0, "success");
        Assert.Throws<OutlookCalendarException>(monitor.ThrowIfFailed);
    }

    [Fact]
    public async Task EmptyDpapiCacheRequiresExplicitSignInWithoutNetwork()
    {
        var directory = Path.Combine(AppContext.BaseDirectory, "test-cache-" + Guid.NewGuid().ToString("N"));
        try
        {
            await using var provider = await OutlookCalendarProvider.CreateAsync(
                Guid.NewGuid().ToString("D"), "organizations", directory, default);
            await Assert.ThrowsAsync<OutlookSignInRequiredException>(() =>
                provider.GetMeetingsAsync(DateTimeOffset.UtcNow, DateTimeOffset.UtcNow.AddHours(1), default));
            await provider.SignOutAsync(default);
        }
        finally
        {
            if (Directory.Exists(directory))
                Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public async Task CorruptDpapiCacheFailsExplicitlyInsteadOfSilentlyFallingBack()
    {
        var directory = Path.Combine(AppContext.BaseDirectory, "test-cache-" + Guid.NewGuid().ToString("N"));
        var clientId = Guid.NewGuid().ToString("D");
        try
        {
            await using var provider = await OutlookCalendarProvider.CreateAsync(clientId, "organizations", directory, default);
            await File.WriteAllTextAsync(Path.Combine(directory, $"outlook-{clientId}-organizations.msalcache"),
                "synthetic invalid encrypted data");
            var error = await Assert.ThrowsAsync<OutlookCalendarException>(() =>
                provider.GetMeetingsAsync(DateTimeOffset.UtcNow, DateTimeOffset.UtcNow.AddHours(1), default));
            Assert.Contains("cache", error.Message);
            Assert.DoesNotContain("synthetic invalid encrypted data", error.ToString());
        }
        finally
        {
            if (Directory.Exists(directory))
                Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public async Task UnwritableCacheLocationFailsExplicitly()
    {
        var file = Path.Combine(AppContext.BaseDirectory, "test-cache-file-" + Guid.NewGuid().ToString("N"));
        try
        {
            await File.WriteAllTextAsync(file, "synthetic file occupying requested directory");
            var error = await Assert.ThrowsAsync<OutlookCalendarException>(() =>
                OutlookCalendarProvider.CreateAsync(Guid.NewGuid().ToString("D"), "organizations", file, default));
            Assert.Contains("cache", error.Message);
            Assert.DoesNotContain(file, error.ToString());
        }
        finally { File.Delete(file); }
    }

    [Fact]
    public async Task DisposedProviderRejectsFurtherAccess()
    {
        var directory = Path.Combine(AppContext.BaseDirectory, "test-cache-" + Guid.NewGuid().ToString("N"));
        try
        {
            var provider = await OutlookCalendarProvider.CreateAsync(
                Guid.NewGuid().ToString("D"), "organizations", directory, default);
            provider.Dispose();
            await provider.DisposeAsync();
            await Assert.ThrowsAsync<ObjectDisposedException>(() => provider.SignOutAsync(default));
            await Assert.ThrowsAsync<ObjectDisposedException>(() =>
                provider.GetMeetingsAsync(DateTimeOffset.UtcNow, DateTimeOffset.UtcNow.AddHours(1), default));
        }
        finally
        {
            if (Directory.Exists(directory))
                Directory.Delete(directory, recursive: true);
        }
    }
}
