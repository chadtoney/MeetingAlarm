using System.Globalization;
using System.Net;
using System.Text.RegularExpressions;
using MeetingAlarm.Core;

namespace MeetingAlarm.ClassicOutlook;

public sealed class ClassicOutlookException : IOException
{
    internal ClassicOutlookException(string message) : base(message) { }
}

internal interface ICalendarSnapshotReader : IDisposable
{
    IReadOnlyList<Meeting> Read(DateTimeOffset from, DateTimeOffset to, CancellationToken token);
}

internal static class CalendarData
{
    internal const int MaximumItems = 10_000;

    internal static MeetingResponse MapResponse(int value) => value switch
    {
        0 or 5 => MeetingResponse.None,
        1 => MeetingResponse.Organizer,
        2 => MeetingResponse.Tentative,
        3 => MeetingResponse.Accepted,
        4 => MeetingResponse.Declined,
        _ => throw new ClassicOutlookException("Classic Outlook returned an unknown response status.")
    };

    internal static bool IsCancelled(int status) => (status & 4) != 0;

    internal static bool IsOtherAttendee(int type) => type switch
    {
        0 => false, // olOrganizer
        1 or 2 or 3 => true,
        _ => throw new ClassicOutlookException("Classic Outlook returned an unknown recipient type.")
    };

    internal static DateTimeOffset AsUtc(DateTime value) =>
        new(DateTime.SpecifyKind(value, DateTimeKind.Utc));

    internal static Meeting MapMeeting(string store, string folder, string entry, string? subject,
        DateTimeOffset start, DateTimeOffset end, int responseStatus, int meetingStatus,
        bool hasOtherAttendees, bool isAllDay, string? joinUrl)
    {
        if (end <= start)
            throw new ClassicOutlookException("Classic Outlook returned invalid appointment times.");
        return new Meeting(OccurrenceId(store, folder, entry, start), subject ?? "", start, end,
            MapResponse(responseStatus), hasOtherAttendees, isAllDay, IsCancelled(meetingStatus),
            joinUrl, IsLocal: false);
    }

    internal static string OccurrenceId(string store, string folder, string entry, DateTimeOffset start)
    {
        if (string.IsNullOrWhiteSpace(store) || string.IsNullOrWhiteSpace(folder) ||
            string.IsNullOrWhiteSpace(entry))
            throw new ClassicOutlookException("Classic Outlook returned an incomplete calendar identity.");
        return $"classic:{store.Length}:{store}:{folder.Length}:{folder}:{entry.Length}:{entry}:{start.UtcTicks}";
    }

    internal static string BuildOverlapQuery(DateTimeOffset from, DateTimeOffset to,
        TimeZoneInfo zone, CultureInfo culture)
    {
        var lower = TimeZoneInfo.ConvertTime(from, zone).DateTime;
        var upper = TimeZoneInfo.ConvertTime(to, zone).DateTime;
        // Jet compares local wall times without seconds. Widen ambiguous DST boundaries,
        // then apply the exact half-open UTC overlap test to every returned occurrence.
        if (zone.IsAmbiguousTime(lower)) lower = lower.AddDays(-1);
        if (zone.IsAmbiguousTime(upper)) upper = upper.AddDays(1);
        lower = new DateTime(lower.Ticks - lower.Ticks % TimeSpan.TicksPerMinute);
        upper = new DateTime(upper.Ticks - upper.Ticks % TimeSpan.TicksPerMinute).AddMinutes(1);
        static string Literal(DateTime date, CultureInfo format) =>
            date.ToString("g", format).Replace("'", "''", StringComparison.Ordinal);
        return $"[Start] < '{Literal(upper, culture)}' AND [End] > '{Literal(lower, culture)}'";
    }

    internal static bool Overlaps(DateTimeOffset start, DateTimeOffset end,
        DateTimeOffset from, DateTimeOffset to) => start < to && end > from;
}

internal static class TeamsLinks
{
    private const int MaximumTextLength = 1_048_576;
    private static readonly Regex UrlPattern = new(
        """https://[^\s<>"']{1,8192}""",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.NonBacktracking,
        TimeSpan.FromMilliseconds(200));

    internal static string? Validate(string? value)
    {
        if (string.IsNullOrWhiteSpace(value) || value.Length > 8192) return null;
        value = WebUtility.HtmlDecode(value.Trim());
        return MeetingLink.TryCreate(value, out var link) && link.IsTeams ? link.Address.AbsoluteUri : null;
    }

    internal static string? Extract(string? text)
    {
        if (string.IsNullOrEmpty(text)) return null;
        if (text.Length > MaximumTextLength)
            throw new ClassicOutlookException("A classic Outlook item exceeds the safe link-extraction limit.");
        var decoded = WebUtility.HtmlDecode(text);
        foreach (Match match in UrlPattern.Matches(decoded))
        {
            var value = match.Value.TrimEnd('.', ',', ';', ')', ']', '}');
            if (Validate(value) is { } url) return url;
        }
        return null;
    }
}
