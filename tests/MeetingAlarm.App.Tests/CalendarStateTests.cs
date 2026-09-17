using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using MeetingAlarm.App.Services;
using MeetingAlarm.Core;
using Xunit;

namespace MeetingAlarm.App.Tests;

public sealed class CalendarStateTests : IDisposable
{
    private readonly string directory = Path.Combine(Path.GetTempPath(), "MeetingAlarm.Tests", Guid.NewGuid().ToString("N"));
    private static readonly DateTimeOffset Now = new(2026, 9, 8, 20, 0, 0, TimeSpan.Zero);
    private static readonly Meeting Local = new("local:1", "Local test", Now, Now.AddMinutes(30),
        MeetingResponse.Accepted, false, IsLocal: true);
    private static readonly Meeting Graph = new("graphId", "Synthetic calendar test", Now, Now.AddMinutes(30),
        MeetingResponse.Accepted, true);
    private static readonly Meeting Classic = Graph with { Id = "classic:1" };

    [Theory]
    [InlineData(false, CalendarSource.LocalOnly)]
    [InlineData(true, CalendarSource.MicrosoftGraph)]
    public void LegacyStateRetainsItsOriginalSource(bool enabled, CalendarSource expected)
    {
        var state = JsonSerializer.Deserialize<AppState>($"{{\"OutlookEnabled\":{enabled.ToString().ToLowerInvariant()}}}")!;
        Assert.Equal(expected, state.ActiveSource);
    }

    [Theory]
    [InlineData(CalendarSource.LocalOnly)]
    [InlineData(CalendarSource.ClassicOutlook)]
    [InlineData(CalendarSource.MicrosoftGraph)]
    public void ExplicitSourceOverridesLegacyFlag(CalendarSource source)
    {
        Assert.Equal(source, (new AppState { Source = source, OutlookEnabled = true }).ActiveSource);
    }

    [Fact]
    public void SourceSwitchReplacesCacheWithoutDroppingLocalAlarmsOrGraphConfiguration()
    {
        var state = new AppState { LocalMeetings = [Local], ClientId = "saved-id", TenantId = "saved-tenant" }
            .WithCalendar(CalendarSource.MicrosoftGraph, [Graph], Now);
        var switched = state.WithCalendar(CalendarSource.ClassicOutlook, [Classic], Now.AddMinutes(1));
        Assert.Equal(CalendarSource.ClassicOutlook, switched.ActiveSource);
        Assert.False(switched.OutlookEnabled);
        Assert.Equal("saved-id", switched.ClientId);
        Assert.Equal("saved-tenant", switched.TenantId);
        Assert.Equal(2, switched.AllMeetings().Count);
        Assert.Contains(Local, switched.AllMeetings());
        Assert.Contains(Classic, switched.AllMeetings());
        Assert.DoesNotContain(Graph, switched.AllMeetings());
        Assert.Equal(CalendarSource.MicrosoftGraph, state.ActiveSource);
        Assert.Equal(Now.AddMinutes(1), switched.LastSync);
    }

    [Fact]
    public void DisconnectSuppressesCacheButKeepsLocalAlarmsAndPreferences()
    {
        var options = new AlarmOptions { SoundEnabled = false };
        var state = new AppState { LocalMeetings = [Local], Options = options }
            .WithCalendar(CalendarSource.ClassicOutlook, [Classic], Now).WithoutCalendar();
        Assert.False(state.HasCalendarSource);
        Assert.False(state.OutlookEnabled);
        Assert.Empty(state.CachedMeetings);
        Assert.Null(state.LastSync);
        Assert.Equal(Local, Assert.Single(state.AllMeetings()));
        Assert.Same(options, state.Options);
    }

    [Fact]
    public void InactiveLeftoverCacheNeverProducesAlarms()
    {
        var state = new AppState { Source = CalendarSource.LocalOnly, CachedMeetings = [Graph] };
        Assert.Empty(state.AllMeetings());
    }

    [Fact]
    public void SwitchingBackToGraphPreservesPerSourceDecisions()
    {
        var decision = new AlarmDecision(Graph.OccurrenceKey, Graph.End, Dismissed: true);
        var state = new AppState { Decisions = [decision] }
            .WithCalendar(CalendarSource.ClassicOutlook, [Classic], Now)
            .WithCalendar(CalendarSource.MicrosoftGraph, [Graph], Now);
        Assert.True(state.OutlookEnabled);
        Assert.Empty(new AlarmScheduler(state.Decisions).Tick(Now, state.AllMeetings(), state.Options));
    }

    [Theory]
    [InlineData(CalendarSource.LocalOnly)]
    [InlineData((CalendarSource)55)]
    public void RejectsInvalidConnectedSource(CalendarSource source)
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => new AppState().WithCalendar(source, [], Now));
    }

    [Fact]
    public void SourceChoiceAndCacheSurviveEncryptedRoundTrip()
    {
        var store = new ProtectedStateStore(directory);
        var state = new AppState { LocalMeetings = [Local] }
            .WithCalendar(CalendarSource.ClassicOutlook, [Classic], Now);
        store.Save(state);
        var loaded = store.Load();
        Assert.Equal(CalendarSource.ClassicOutlook, loaded.ActiveSource);
        Assert.Equal(state.AllMeetings(), loaded.AllMeetings());
        Assert.DoesNotContain("Synthetic calendar test", Encoding.UTF8.GetString(File.ReadAllBytes(Path.Combine(directory, "state.bin"))));
        Assert.False(File.Exists(Path.Combine(directory, "state.bin.tmp")));
    }

    [Fact]
    public void MissingStateIsLocalOnlyAndDoesNotLaunchAConnector()
    {
        Assert.Equal(CalendarSource.LocalOnly, new ProtectedStateStore(directory).Load().ActiveSource);
    }

    [Fact]
    public void UnknownSourceInProtectedFileFailsExplicitly()
    {
        var store = new ProtectedStateStore(directory);
        store.Save(new AppState { Source = (CalendarSource)999 });
        Assert.Throws<InvalidDataException>(() => store.Load());
    }

    [Fact]
    public void CorruptStateIsNotSilentlyReplaced()
    {
        Directory.CreateDirectory(directory);
        var path = Path.Combine(directory, "state.bin");
        byte[] bytes = [1, 2, 3, 4];
        File.WriteAllBytes(path, bytes);
        Assert.Throws<CryptographicException>(() => new ProtectedStateStore(directory).Load());
        Assert.Equal(bytes, File.ReadAllBytes(path));
    }

    public void Dispose()
    {
        if (Directory.Exists(directory))
            Directory.Delete(directory, recursive: true);
    }
}
