namespace MeetingAlarm.Core;

public sealed class AlarmScheduler
{
    public static readonly TimeSpan CatchUpWindow = TimeSpan.FromMinutes(5);
    private readonly Dictionary<string, AlarmDecision> decisions;
    private readonly HashSet<string> active = [];

    public AlarmScheduler(IEnumerable<AlarmDecision>? savedDecisions = null)
    {
        decisions = (savedDecisions ?? []).ToDictionary(d => d.OccurrenceKey);
    }

    public IReadOnlyCollection<AlarmDecision> Decisions => decisions.Values;

    public static bool IsEligible(Meeting meeting, AlarmOptions options)
    {
        if (meeting.IsCancelled || meeting.IsAllDay || meeting.Response == MeetingResponse.Declined)
            return false;

        var preference = options.Overrides.GetValueOrDefault(meeting.OccurrenceKey);
        if (preference != MeetingOverride.Default)
            return preference == MeetingOverride.Include;

        if (meeting.IsLocal)
            return true;
        if (!meeting.HasOtherAttendees)
            return false;

        return meeting.Response switch
        {
            MeetingResponse.Accepted or MeetingResponse.Organizer => true,
            MeetingResponse.Tentative => options.IncludeTentative,
            MeetingResponse.None => options.IncludeUnanswered,
            _ => false
        };
    }

    // Call with a complete snapshot, not a delta: deleted/canceled/rescheduled
    // occurrences must withdraw any alarm that is already on screen.
    public IReadOnlyList<Meeting> Tick(
        DateTimeOffset now, IReadOnlyList<Meeting> meetings, AlarmOptions options)
    {
        foreach (var key in decisions.Where(p => p.Value.ExpiresAt <= now).Select(p => p.Key).ToArray())
            decisions.Remove(key);

        if (options.QuietUntil > now)
        {
            active.Clear();
            return [];
        }

        var current = meetings
            .Where(m => m.Start <= now && m.End > now && IsEligible(m, options))
            .DistinctBy(m => m.OccurrenceKey)
            .ToDictionary(m => m.OccurrenceKey);
        active.IntersectWith(current.Keys);

        foreach (var (key, meeting) in current)
        {
            decisions.TryGetValue(key, out var decision);
            if (decision?.Dismissed == true || decision?.SnoozeUntil > now)
            {
                active.Remove(key);
                continue;
            }

            if (active.Contains(key) ||
                decision?.SnoozeUntil <= now ||
                now - meeting.Start <= CatchUpWindow)
                active.Add(key);
        }

        return active.Select(key => current[key]).OrderBy(m => m.Start).ToArray();
    }

    public void Dismiss(Meeting meeting)
    {
        decisions[meeting.OccurrenceKey] = new(meeting.OccurrenceKey, meeting.End, Dismissed: true);
        active.Remove(meeting.OccurrenceKey);
    }

    public void Snooze(Meeting meeting, DateTimeOffset now, TimeSpan duration)
    {
        if (duration <= TimeSpan.Zero)
            throw new ArgumentOutOfRangeException(nameof(duration), "Snooze duration must be positive.");
        decisions[meeting.OccurrenceKey] = new(meeting.OccurrenceKey, meeting.End, SnoozeUntil: now + duration);
        active.Remove(meeting.OccurrenceKey);
    }
}
