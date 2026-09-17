using System.Globalization;
using System.Runtime.InteropServices;
using MeetingAlarm.Core;

namespace MeetingAlarm.ClassicOutlook;

internal sealed class ComSnapshotReader : ICalendarSnapshotReader
{
    private const string TeamsProperty =
        "http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/SkypeTeamsMeetingUrl";
    private const int MapiNotFound = unchecked((int)0x8004010F);

    public IReadOnlyList<Meeting> Read(DateTimeOffset from, DateTimeOffset to, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        using var scope = new ComScope();
        var type = Type.GetTypeFromProgID("Outlook.Application")
            ?? throw new ClassicOutlookException("Classic Outlook is not installed or its COM registration is unavailable.");
        dynamic outlook = scope.Own(Activator.CreateInstance(type));
        token.ThrowIfCancellationRequested();
        dynamic session = scope.Own(outlook.GetNamespace("MAPI"));
        // GetDefaultFolder connects to the existing default profile; never create a profile,
        // call Logon with credentials, or change Outlook's security configuration.
        dynamic folder = scope.Own(session.GetDefaultFolder(9));
        string folderId = folder.EntryID;
        string storeId = folder.StoreID;
        dynamic items = scope.Own(folder.Items);
        items.Sort("[Start]", false);
        items.IncludeRecurrences = true;
        dynamic restricted = scope.Own(items.Restrict(
            CalendarData.BuildOverlapQuery(from, to, TimeZoneInfo.Local, CultureInfo.CurrentCulture)));
        var result = new List<Meeting>();
        var identities = new HashSet<string>(StringComparer.Ordinal);
        object? current = restricted.GetFirst();
        var count = 0;
        while (current is not null)
        {
            using (var itemScope = new ComScope())
            {
                dynamic item = itemScope.Own(current);
                token.ThrowIfCancellationRequested();
                if (++count > CalendarData.MaximumItems)
                    throw new ClassicOutlookException("Classic Outlook exceeded the calendar snapshot limit. Choose a smaller time window.");
                if ((int)item.Class != 26)
                    throw new ClassicOutlookException("Classic Outlook returned a non-appointment calendar item.");
                var start = CalendarData.AsUtc((DateTime)item.StartUTC);
                var end = CalendarData.AsUtc((DateTime)item.EndUTC);
                if (end <= start)
                    throw new ClassicOutlookException("Classic Outlook returned invalid appointment times.");
                if (CalendarData.Overlaps(start, end, from, to))
                {
                    string entry = item.EntryID;
                    string subject = item.Subject;
                    int response = item.ResponseStatus;
                    int status = item.MeetingStatus;
                    bool allDay = item.AllDayEvent;
                    bool attendees = HasOtherAttendees(item, token);
                    string? url = ReadJoinUrl(item, itemScope);
                    var meeting = CalendarData.MapMeeting(storeId, folderId, entry, subject,
                        start, end, response, status, attendees, allDay, url);
                    if (!identities.Add(meeting.Id))
                        throw new ClassicOutlookException("Classic Outlook returned duplicate calendar occurrences.");
                    result.Add(meeting);
                }
            }
            token.ThrowIfCancellationRequested();
            current = restricted.GetNext();
        }
        token.ThrowIfCancellationRequested();
        return result.OrderBy(meeting => meeting.Start).ToArray();
    }

    private static bool HasOtherAttendees(dynamic item, CancellationToken token)
    {
        using var scope = new ComScope();
        dynamic recipients = scope.Own(item.Recipients);
        int count = recipients.Count;
        if (count < 0 || count > CalendarData.MaximumItems)
            throw new ClassicOutlookException("Classic Outlook exceeded the recipient snapshot limit.");
        for (var i = 1; i <= count; i++)
        {
            token.ThrowIfCancellationRequested();
            using var recipientScope = new ComScope();
            dynamic recipient = recipientScope.Own(recipients.Item(i));
            if (CalendarData.IsOtherAttendee((int)recipient.Type)) return true;
        }
        return false;
    }

    private static string? ReadJoinUrl(dynamic item, ComScope scope)
    {
        dynamic accessor = scope.Own(item.PropertyAccessor);
        string? property = null;
        try
        {
            property = (string)accessor.GetProperty(TeamsProperty);
        }
        catch (COMException exception) when (exception.HResult == MapiNotFound)
        {
            // This named property is absent on many ordinary calendar items.
        }
        if (TeamsLinks.Validate(property) is { } known) return known;
        string? location = item.Location;
        if (TeamsLinks.Extract(location) is { } locationLink) return locationLink;
        // Body can be protected by Outlook programmatic-access policy. A denial is an
        // error for the entire snapshot, not a successful empty/link-less fallback.
        string? body = item.Body;
        return TeamsLinks.Extract(body);
    }

    public void Dispose() { } // Every RCW has operation-local ownership.
}

internal sealed class ComScope : IDisposable
{
    private readonly Stack<object> owned = new();

    internal object Own(object? value)
    {
        if (value is null)
            throw new ClassicOutlookException("Classic Outlook returned an unavailable calendar object.");
        owned.Push(value);
        return value;
    }

    public void Dispose()
    {
        var failed = false;
        while (owned.TryPop(out var value))
        {
            try
            {
                if (Marshal.IsComObject(value))
                    Marshal.ReleaseComObject(value);
            }
            catch (Exception exception) when (exception is COMException or InvalidComObjectException)
            {
                failed = true;
            }
        }
        if (failed)
            throw new ClassicOutlookException("Classic Outlook could not release all calendar objects cleanly.");
    }
}
