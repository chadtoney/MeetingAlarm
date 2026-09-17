using System.Globalization;
using MeetingAlarm.Core;
using Xunit;

namespace MeetingAlarm.ClassicOutlook.Tests;

public class CalendarDataTests
{
    [Theory]
    [InlineData(2, true, false)]
    [InlineData(0, true, false)]
    [InlineData(5, true, false)]
    [InlineData(3, false, false)]
    [InlineData(1, false, false)]
    [InlineData(3, true, true)]
    [InlineData(1, true, true)]
    [InlineData(4, true, false)]
    public void CalendarMeetingsNeverBypassDefaultEligibility(int response, bool attendees, bool expected)
    {
        var start = DateTimeOffset.Parse("2026-09-08T15:00:00Z");
        var meeting = CalendarData.MapMeeting("store", "folder", "entry", "Test",
            start, start.AddHours(1), response, 1, attendees, false, null);
        Assert.False(meeting.IsLocal);
        Assert.Equal(expected, AlarmScheduler.IsEligible(meeting, new AlarmOptions()));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    public void MappingRejectsZeroOrNegativeDuration(int minutes)
    {
        var start = DateTimeOffset.Parse("2026-09-08T15:00:00Z");
        Assert.Throws<ClassicOutlookException>(() => CalendarData.MapMeeting("store", "folder", "entry",
            "Test", start, start.AddMinutes(minutes), 3, 1, true, false, null));
    }

    [Theory]
    [InlineData(0, MeetingResponse.None)]
    [InlineData(1, MeetingResponse.Organizer)]
    [InlineData(2, MeetingResponse.Tentative)]
    [InlineData(3, MeetingResponse.Accepted)]
    [InlineData(4, MeetingResponse.Declined)]
    [InlineData(5, MeetingResponse.None)]
    public void MapsResponse(int value, MeetingResponse expected) =>
        Assert.Equal(expected, CalendarData.MapResponse(value));

    [Fact]
    public void UnknownResponseFails() =>
        Assert.Throws<ClassicOutlookException>(() => CalendarData.MapResponse(6));

    [Theory]
    [InlineData(0, false)]
    [InlineData(1, false)]
    [InlineData(3, false)]
    [InlineData(4, true)]
    [InlineData(5, true)]
    [InlineData(7, true)]
    public void MapsCancelledBit(int status, bool expected) =>
        Assert.Equal(expected, CalendarData.IsCancelled(status));

    [Theory]
    [InlineData(0, false)]
    [InlineData(1, true)]
    [InlineData(2, true)]
    [InlineData(3, true)]
    public void ExcludesOrganizerRecipient(int type, bool expected) =>
        Assert.Equal(expected, CalendarData.IsOtherAttendee(type));

    [Fact]
    public void UtcComDateNeverGetsLocalOffsetApplied()
    {
        var input = new DateTime(2026, 11, 1, 6, 30, 0, DateTimeKind.Unspecified);
        Assert.Equal(new DateTimeOffset(2026, 11, 1, 6, 30, 0, TimeSpan.Zero), CalendarData.AsUtc(input));
    }

    [Theory]
    [InlineData("en-US", "[Start] < '9/9/2026 1:01 AM' AND [End] > '9/8/2026 11:59 PM'")]
    [InlineData("en-GB", "[Start] < '09/09/2026 01:01' AND [End] > '08/09/2026 23:59'")]
    [InlineData("de-DE", "[Start] < '09.09.2026 01:01' AND [End] > '08.09.2026 23:59'")]
    public void QueryUsesRegionalDatesAndNoSeconds(string culture, string expected)
    {
        var zone = TimeZoneInfo.CreateCustomTimeZone("Test", TimeSpan.FromHours(-5), "Test", "Test");
        Assert.Equal(expected, CalendarData.BuildOverlapQuery(
            DateTimeOffset.Parse("2026-09-09T04:59:30Z"),
            DateTimeOffset.Parse("2026-09-09T06:00:01Z"),
            zone, CultureInfo.GetCultureInfo(culture)));
    }

    [Fact]
    public void FallBackAmbiguousBoundariesWidenInsteadOfReversingWindow()
    {
        var zone = TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time");
        var query = CalendarData.BuildOverlapQuery(
            DateTimeOffset.Parse("2026-11-01T05:45:00Z"),
            DateTimeOffset.Parse("2026-11-01T06:15:00Z"),
            zone, CultureInfo.GetCultureInfo("en-US"));
        Assert.Equal("[Start] < '11/2/2026 1:16 AM' AND [End] > '10/31/2026 1:45 AM'", query);
    }

    [Fact]
    public void SpringForwardUsesValidLocalBoundaries()
    {
        var zone = TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time");
        Assert.Equal("[Start] < '3/8/2026 3:31 AM' AND [End] > '3/8/2026 1:30 AM'",
            CalendarData.BuildOverlapQuery(DateTimeOffset.Parse("2026-03-08T06:30:00Z"),
                DateTimeOffset.Parse("2026-03-08T07:30:00Z"), zone, CultureInfo.GetCultureInfo("en-US")));
    }

    [Fact]
    public void ExactOverlapIncludesCrossMidnightAndExcludesTouchingBoundaries()
    {
        var from = DateTimeOffset.Parse("2026-09-09T00:00:00Z");
        var to = from.AddHours(1);
        Assert.True(CalendarData.Overlaps(from.AddHours(-2), from.AddMinutes(1), from, to));
        Assert.False(CalendarData.Overlaps(from.AddHours(-2), from, from, to));
        Assert.False(CalendarData.Overlaps(to, to.AddHours(1), from, to));
    }

    [Fact]
    public void StableIdentityDistinguishesOccurrencesStoresAndDelimiterContent()
    {
        var start = DateTimeOffset.Parse("2026-09-09T00:00:00Z");
        var id = CalendarData.OccurrenceId("store", "folder", "entry", start);
        Assert.StartsWith("classic:", id);
        Assert.Equal(id, CalendarData.OccurrenceId("store", "folder", "entry", start.ToOffset(TimeSpan.FromHours(5))));
        Assert.NotEqual(id, CalendarData.OccurrenceId("store", "folder", "entry", start.AddDays(1)));
        Assert.NotEqual(id, CalendarData.OccurrenceId("other", "folder", "entry", start));
        Assert.NotEqual(CalendarData.OccurrenceId("a:b", "c", "d", start),
            CalendarData.OccurrenceId("a", "b:c", "d", start));
        Assert.Throws<ClassicOutlookException>(() => CalendarData.OccurrenceId("", "folder", "entry", start));
    }
}
