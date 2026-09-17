using MeetingAlarm.Core;
using Xunit;

namespace MeetingAlarm.Core.Tests;

public sealed class AlarmSchedulerTests
{
    private static readonly DateTimeOffset Start = new(2026, 9, 8, 18, 0, 0, TimeSpan.Zero);
    private static readonly Meeting Meeting = new("one", "Meeting", Start, Start.AddHours(1),
        MeetingResponse.Accepted, true);
    private static readonly AlarmOptions Options = new();

    [Fact]
    public void FiresAtStartNotBeforeAndStaysUntilAcknowledged()
    {
        var scheduler = new AlarmScheduler();
        Assert.Empty(scheduler.Tick(Start.AddSeconds(-1), [Meeting], Options));
        Assert.Single(scheduler.Tick(Start, [Meeting], Options));
        Assert.Single(scheduler.Tick(Start.AddMinutes(20), [Meeting], Options));
        Assert.Empty(scheduler.Tick(Meeting.End, [Meeting], Options));
    }

    [Fact]
    public void DismissalSurvivesRestart()
    {
        var scheduler = new AlarmScheduler();
        scheduler.Dismiss(Meeting);
        var restarted = new AlarmScheduler(scheduler.Decisions);
        Assert.Empty(restarted.Tick(Start, [Meeting], Options));
    }

    [Fact]
    public void SnoozeSurvivesRestartAndCanFireOutsideCatchUpWindow()
    {
        var scheduler = new AlarmScheduler();
        scheduler.Snooze(Meeting, Start.AddMinutes(9), TimeSpan.FromMinutes(1));
        scheduler = new AlarmScheduler(scheduler.Decisions);
        Assert.Empty(scheduler.Tick(Start.AddMinutes(9), [Meeting], Options));
        Assert.Single(scheduler.Tick(Start.AddMinutes(10), [Meeting], Options));
    }

    [Fact]
    public void CancellationDeletionAndReschedulingWithdrawActiveAlarm()
    {
        var scheduler = new AlarmScheduler();
        Assert.Single(scheduler.Tick(Start, [Meeting], Options));
        Assert.Empty(scheduler.Tick(Start, [Meeting with { IsCancelled = true }], Options));
        Assert.Single(scheduler.Tick(Start, [Meeting], Options));
        Assert.Empty(scheduler.Tick(Start, [], Options));
        Assert.Single(scheduler.Tick(Start, [Meeting], Options));
        Assert.Empty(scheduler.Tick(Start, [Meeting with { Start = Start.AddHours(1) }], Options));
    }

    [Fact]
    public void RescheduledMeetingGetsFreshAlarmEvenIfOriginalDismissed()
    {
        var scheduler = new AlarmScheduler();
        scheduler.Dismiss(Meeting);
        var moved = Meeting with { Start = Start.AddMinutes(10) };
        Assert.Single(scheduler.Tick(moved.Start, [moved], Options));
    }

    [Fact]
    public void QuietModeSuppressesAlarmsAndExpiredPauseRecovers()
    {
        var scheduler = new AlarmScheduler();
        var options = Options with { QuietUntil = Start.AddMinutes(1) };
        Assert.Empty(scheduler.Tick(Start, [Meeting], options));
        Assert.Single(scheduler.Tick(Start.AddMinutes(1), [Meeting], options));
    }

    [Fact]
    public void ResumeDoesNotDumpHoursOfOldReminders()
    {
        var scheduler = new AlarmScheduler();
        Assert.Single(scheduler.Tick(Start.AddMinutes(4), [Meeting], Options));
        scheduler = new AlarmScheduler();
        Assert.Empty(scheduler.Tick(Start.AddMinutes(6), [Meeting], Options));
    }

    [Theory]
    [InlineData(MeetingResponse.Accepted, true)]
    [InlineData(MeetingResponse.Organizer, true)]
    [InlineData(MeetingResponse.Tentative, false)]
    [InlineData(MeetingResponse.None, false)]
    [InlineData(MeetingResponse.Declined, false)]
    public void DefaultResponseRules(MeetingResponse response, bool eligible)
    {
        Assert.Equal(eligible, AlarmScheduler.IsEligible(Meeting with { Response = response }, Options));
    }

    [Fact]
    public void TentativeAndUnansweredAreExplicitOptIns()
    {
        Assert.True(AlarmScheduler.IsEligible(Meeting with { Response = MeetingResponse.Tentative },
            Options with { IncludeTentative = true }));
        Assert.True(AlarmScheduler.IsEligible(Meeting with { Response = MeetingResponse.None },
            Options with { IncludeUnanswered = true }));
    }

    [Fact]
    public void ExcludesSoloAndAllDayEventsButAllowsLocalAlarms()
    {
        Assert.False(AlarmScheduler.IsEligible(Meeting with { HasOtherAttendees = false }, Options));
        Assert.False(AlarmScheduler.IsEligible(Meeting with { IsAllDay = true }, Options));
        Assert.True(AlarmScheduler.IsEligible(Meeting with { HasOtherAttendees = false, IsLocal = true }, Options));
    }

    [Fact]
    public void PerOccurrenceOverridesDoNotOverrideCancellationOrDecline()
    {
        var options = Options with { Overrides = new() { [Meeting.OccurrenceKey] = MeetingOverride.Include } };
        Assert.True(AlarmScheduler.IsEligible(Meeting with { Response = MeetingResponse.None }, options));
        Assert.False(AlarmScheduler.IsEligible(Meeting with { IsCancelled = true }, options));
        Assert.False(AlarmScheduler.IsEligible(Meeting with { Response = MeetingResponse.Declined }, options));
        Assert.False(AlarmScheduler.IsEligible(Meeting,
            Options with { Overrides = new() { [Meeting.OccurrenceKey] = MeetingOverride.Exclude } }));
    }

    [Fact]
    public void OverlappingMeetingsHaveIndependentDecisions()
    {
        var second = Meeting with { Id = "two" };
        var scheduler = new AlarmScheduler();
        Assert.Equal(2, scheduler.Tick(Start, [Meeting, second], Options).Count);
        scheduler.Dismiss(Meeting);
        Assert.Equal(second, Assert.Single(scheduler.Tick(Start, [Meeting, second], Options)));
    }

    [Fact]
    public void DuplicateSnapshotEntriesDoNotDuplicateAlarm()
    {
        Assert.Single(new AlarmScheduler().Tick(Start, [Meeting, Meeting], Options));
    }

    [Fact]
    public void SameInstantInDifferentTimeZonesHasSameOccurrenceKey()
    {
        Assert.Equal(Meeting.OccurrenceKey, (Meeting with { Start = Start.ToOffset(TimeSpan.FromHours(-5)) }).OccurrenceKey);
    }

    [Fact]
    public void ClockMovingBackwardsWithdrawsAlarmUntilStart()
    {
        var scheduler = new AlarmScheduler();
        Assert.Single(scheduler.Tick(Start, [Meeting], Options));
        Assert.Empty(scheduler.Tick(Start.AddMinutes(-1), [Meeting], Options));
        Assert.Single(scheduler.Tick(Start, [Meeting], Options));
    }

    [Fact]
    public void ExpiredDecisionsArePrunedAndInvalidSnoozeIsRejected()
    {
        var scheduler = new AlarmScheduler();
        scheduler.Dismiss(Meeting);
        scheduler.Tick(Meeting.End, [Meeting], Options);
        Assert.Empty(scheduler.Decisions);
        Assert.Throws<ArgumentOutOfRangeException>(() => scheduler.Snooze(Meeting, Start, TimeSpan.Zero));
    }
}
