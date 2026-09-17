using System.Text.Json.Serialization;
using MeetingAlarm.Core;

namespace MeetingAlarm.App.Services;

public sealed record AppState
{
    public int Version { get; init; } = 1;
    public AlarmOptions Options { get; init; } = new();
    public List<Meeting> LocalMeetings { get; init; } = [];
    public List<Meeting> CachedMeetings { get; init; } = [];
    public List<AlarmDecision> Decisions { get; init; } = [];
    public DateTimeOffset? LastSync { get; init; }
    public string ClientId { get; init; } = "";
    public string TenantId { get; init; } = "organizations";
    public bool OutlookEnabled { get; init; }
    public CalendarSource? Source { get; init; }

    // Older state files only have OutlookEnabled, which always meant Graph.
    [JsonIgnore]
    public CalendarSource ActiveSource =>
        Source ?? (OutlookEnabled ? CalendarSource.MicrosoftGraph : CalendarSource.LocalOnly);

    [JsonIgnore]
    public bool HasCalendarSource => ActiveSource != CalendarSource.LocalOnly;

    public IReadOnlyList<Meeting> AllMeetings() => LocalMeetings
        .Concat(HasCalendarSource ? CachedMeetings : [])
        .OrderBy(meeting => meeting.Start).ToArray();

    public AppState WithCalendar(CalendarSource source, IReadOnlyList<Meeting> meetings, DateTimeOffset syncedAt)
    {
        if (source == CalendarSource.LocalOnly || !Enum.IsDefined(source))
            throw new ArgumentOutOfRangeException(nameof(source), "Choose a calendar connector.");
        return this with
        {
            Source = source,
            OutlookEnabled = source == CalendarSource.MicrosoftGraph,
            CachedMeetings = meetings.ToList(),
            LastSync = syncedAt
        };
    }

    public AppState WithoutCalendar() => this with
    {
        Source = CalendarSource.LocalOnly, OutlookEnabled = false, CachedMeetings = [], LastSync = null
    };
}
