namespace MeetingAlarm.Core;

public enum MeetingResponse
{
    None,
    Accepted,
    Tentative,
    Declined,
    Organizer
}

public sealed record Meeting(
    string Id,
    string Subject,
    DateTimeOffset Start,
    DateTimeOffset End,
    MeetingResponse Response,
    bool HasOtherAttendees,
    bool IsAllDay = false,
    bool IsCancelled = false,
    string? JoinUrl = null,
    bool IsLocal = false)
{
    public string OccurrenceKey => $"{Id}|{Start.UtcTicks}";
}

public enum MeetingOverride
{
    Default,
    Include,
    Exclude
}

public sealed record AlarmOptions
{
    public bool IncludeTentative { get; init; }
    public bool IncludeUnanswered { get; init; }
    public bool SoundEnabled { get; init; } = true;
    public bool PrivateDisplay { get; init; } = true;
    public DateTimeOffset? QuietUntil { get; init; }
    public Dictionary<string, MeetingOverride> Overrides { get; init; } = [];
}

public sealed record AlarmDecision(
    string OccurrenceKey,
    DateTimeOffset ExpiresAt,
    bool Dismissed = false,
    DateTimeOffset? SnoozeUntil = null);

public interface ICalendarProvider
{
    Task<IReadOnlyList<Meeting>> GetMeetingsAsync(
        DateTimeOffset from, DateTimeOffset to, CancellationToken cancellationToken);
}
