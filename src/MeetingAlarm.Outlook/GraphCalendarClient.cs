using System.Globalization;
using System.Net;
using System.Net.Http.Headers;
using System.Text.Json;
using MeetingAlarm.Core;

namespace MeetingAlarm.Outlook;

internal sealed class GraphCalendarClient(HttpClient httpClient)
{
    internal const int MaximumPages = 100;
    internal const int MaximumWindowDays = 31;
    private const int MaximumPageBytes = 8 * 1024 * 1024;
    private const string SelectFields =
        "id,subject,start,end,responseStatus,isOrganizer,isCancelled,isAllDay,attendees,onlineMeeting";

    internal async Task<IReadOnlyList<Meeting>> GetMeetingsAsync(
        DateTimeOffset from, DateTimeOffset to, string accessToken, CancellationToken cancellationToken)
    {
        ValidateWindow(from, to);
        var start = Uri.EscapeDataString(from.UtcDateTime.ToString("O", CultureInfo.InvariantCulture));
        var end = Uri.EscapeDataString(to.UtcDateTime.ToString("O", CultureInfo.InvariantCulture));
        var next = new Uri($"https://graph.microsoft.com/v1.0/me/calendarView?startDateTime={start}" +
            $"&endDateTime={end}&$select={SelectFields}&$top=100");
        var meetings = new List<Meeting>();
        var seenPages = new HashSet<string>(StringComparer.Ordinal);
        var seenMeetings = new HashSet<string>(StringComparer.Ordinal);
        for (var pageNumber = 0; next is not null; pageNumber++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (pageNumber >= MaximumPages || !seenPages.Add(next.AbsoluteUri))
                throw new OutlookCalendarException("Outlook pagination exceeded its safety limit or repeated a page.");
            ValidatePageUri(next);
            using var response = await SendAsync(next, accessToken, cancellationToken).ConfigureAwait(false);
            using var document = await ReadPageAsync(response, cancellationToken).ConfigureAwait(false);
            var page = ParsePage(document.RootElement);
            foreach (var meeting in page.Meetings)
            {
                if (!seenMeetings.Add(meeting.OccurrenceKey))
                    throw new OutlookCalendarException("Outlook returned a duplicate meeting occurrence.");
                meetings.Add(meeting);
            }
            next = page.NextLink;
        }
        return meetings;
    }

    internal static void ValidateWindow(DateTimeOffset from, DateTimeOffset to)
    {
        if (to <= from || to - from > TimeSpan.FromDays(MaximumWindowDays))
            throw new ArgumentOutOfRangeException(nameof(to), "The calendar window must be positive and at most 31 days.");
    }

    internal static void ValidatePageUri(Uri uri)
    {
        if (!uri.IsAbsoluteUri || uri.Scheme != Uri.UriSchemeHttps ||
            !string.Equals(uri.Host, "graph.microsoft.com", StringComparison.OrdinalIgnoreCase) ||
            uri.Port != 443 || uri.UserInfo.Length != 0 || uri.Fragment.Length != 0 ||
            !string.Equals(uri.AbsolutePath, "/v1.0/me/calendarView", StringComparison.Ordinal))
            throw new OutlookCalendarException("Outlook returned an unsafe pagination URL.");
    }

    private async Task<HttpResponseMessage> SendAsync(
        Uri uri, string accessToken, CancellationToken cancellationToken)
    {
        for (var attempt = 0; ; attempt++)
        {
            using var request = new HttpRequestMessage(HttpMethod.Get, uri);
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
            request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            request.Headers.TryAddWithoutValidation("Prefer", "outlook.timezone=\"UTC\", IdType=\"ImmutableId\"");
            HttpResponseMessage response;
            try
            {
                response = await httpClient.SendAsync(
                    request, HttpCompletionOption.ResponseHeadersRead, cancellationToken).ConfigureAwait(false);
            }
            catch (HttpRequestException)
            {
                throw new OutlookCalendarException("The Outlook calendar request could not reach Microsoft Graph.");
            }
            catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
            {
                throw new OutlookCalendarException("The Outlook calendar request timed out.");
            }
            if (response.IsSuccessStatusCode)
                return response;
            var status = response.StatusCode;
            var retryAfter = response.Headers.RetryAfter;
            var delay = retryAfter?.Delta ??
                (retryAfter?.Date is { } retryAt ? retryAt - DateTimeOffset.UtcNow : (TimeSpan?)null);
            response.Dispose();
            if (attempt < 2 && (status == HttpStatusCode.TooManyRequests || status == HttpStatusCode.ServiceUnavailable) &&
                delay is { } wait && wait >= TimeSpan.Zero && wait <= TimeSpan.FromSeconds(30))
            {
                await Task.Delay(wait, cancellationToken).ConfigureAwait(false);
                continue;
            }
            throw new OutlookCalendarException(
                $"Microsoft Graph calendar request failed (HTTP {(int)status}).", status);
        }
    }

    private static async Task<JsonDocument> ReadPageAsync(
        HttpResponseMessage response, CancellationToken cancellationToken)
    {
        if (response.Content.Headers.ContentLength > MaximumPageBytes)
            throw new OutlookCalendarException("The Outlook calendar response exceeded the safety size limit.");
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        timeout.CancelAfter(TimeSpan.FromSeconds(60));
        try
        {
            await using var stream = await response.Content.ReadAsStreamAsync(timeout.Token).ConfigureAwait(false);
            using var buffer = new MemoryStream();
            var chunk = new byte[16 * 1024];
            int count;
            while ((count = await stream.ReadAsync(chunk, timeout.Token).ConfigureAwait(false)) != 0)
            {
                if (buffer.Length + count > MaximumPageBytes)
                    throw new OutlookCalendarException("The Outlook calendar response exceeded the safety size limit.");
                buffer.Write(chunk, 0, count);
            }
            return JsonDocument.Parse(buffer.ToArray());
        }
        catch (JsonException)
        {
            throw InvalidData();
        }
        catch (IOException)
        {
            throw new OutlookCalendarException("The Outlook calendar response could not be read.");
        }
        catch (HttpRequestException)
        {
            throw new OutlookCalendarException("The Outlook calendar response could not be read.");
        }
        catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
        {
            throw new OutlookCalendarException("The Outlook calendar response timed out.");
        }
    }

    internal static (IReadOnlyList<Meeting> Meetings, Uri? NextLink) ParsePage(JsonElement root)
    {
        var values = Required(root, "value", JsonValueKind.Array);
        var meetings = new List<Meeting>();
        foreach (var value in values.EnumerateArray())
            meetings.Add(ParseMeeting(value));
        Uri? nextLink = null;
        if (root.TryGetProperty("@odata.nextLink", out var link))
        {
            if (link.ValueKind != JsonValueKind.String ||
                !Uri.TryCreate(link.GetString(), UriKind.Absolute, out nextLink))
                throw InvalidData();
            ValidatePageUri(nextLink);
        }
        return (meetings, nextLink);
    }

    private static Meeting ParseMeeting(JsonElement value)
    {
        var id = Required(value, "id", JsonValueKind.String).GetString();
        if (string.IsNullOrWhiteSpace(id))
            throw InvalidData();
        var subjectValue = RequiredNullable(value, "subject", JsonValueKind.String);
        var subject = subjectValue.ValueKind == JsonValueKind.Null ? null : subjectValue.GetString();
        var start = ParseUtcTime(Required(value, "start", JsonValueKind.Object));
        var end = ParseUtcTime(Required(value, "end", JsonValueKind.Object));
        if (end <= start)
            throw InvalidData();
        var isOrganizer = RequiredBoolean(value, "isOrganizer");
        var isCancelled = RequiredBoolean(value, "isCancelled");
        var isAllDay = RequiredBoolean(value, "isAllDay");
        var attendees = Required(value, "attendees", JsonValueKind.Array);
        foreach (var attendee in attendees.EnumerateArray())
        {
            if (attendee.ValueKind != JsonValueKind.Object)
                throw InvalidData();
        }
        var responseStatus = Required(value, "responseStatus", JsonValueKind.Object);
        var response = Required(responseStatus, "response", JsonValueKind.String).GetString() switch
        {
            "none" or "notResponded" => MeetingResponse.None,
            "accepted" => MeetingResponse.Accepted,
            "tentativelyAccepted" => MeetingResponse.Tentative,
            "declined" => MeetingResponse.Declined,
            "organizer" => MeetingResponse.Organizer,
            _ => throw InvalidData()
        };
        if (isOrganizer)
            response = MeetingResponse.Organizer;
        else if (response == MeetingResponse.Organizer)
            throw InvalidData();
        string? joinUrl = null;
        var onlineMeeting = RequiredNullable(value, "onlineMeeting", JsonValueKind.Object);
        if (onlineMeeting.ValueKind == JsonValueKind.Object)
        {
            var join = RequiredNullable(onlineMeeting, "joinUrl", JsonValueKind.String);
            if (join.ValueKind == JsonValueKind.String)
            {
                joinUrl = join.GetString();
                if (!MeetingLink.TryCreate(joinUrl, out _))
                    throw InvalidData();
            }
        }
        // An invitee has at least one other participant: the meeting's organizer.
        return new Meeting(id, string.IsNullOrWhiteSpace(subject) ? "Untitled meeting" : subject,
            start, end, response, !isOrganizer || attendees.GetArrayLength() > 0,
            isAllDay, isCancelled, joinUrl);
    }

    private static DateTimeOffset ParseUtcTime(JsonElement value)
    {
        var zone = Required(value, "timeZone", JsonValueKind.String).GetString();
        var text = Required(value, "dateTime", JsonValueKind.String).GetString();
        string[] formats =
        [
            "yyyy-MM-dd'T'HH:mm:ss", "yyyy-MM-dd'T'HH:mm:ss.FFFFFFF",
            "yyyy-MM-dd'T'HH:mm:ss'Z'", "yyyy-MM-dd'T'HH:mm:ss.FFFFFFF'Z'",
            "yyyy-MM-dd'T'HH:mm:sszzz", "yyyy-MM-dd'T'HH:mm:ss.FFFFFFFzzz"
        ];
        if (zone != "UTC" || !DateTimeOffset.TryParseExact(text, formats, CultureInfo.InvariantCulture,
                DateTimeStyles.AssumeUniversal, out var time) || time.Offset != TimeSpan.Zero)
            throw InvalidData();
        return time;
    }

    private static JsonElement Required(JsonElement parent, string name, JsonValueKind kind)
    {
        if (parent.ValueKind != JsonValueKind.Object || !parent.TryGetProperty(name, out var value) ||
            value.ValueKind != kind)
            throw InvalidData();
        return value;
    }

    private static JsonElement RequiredNullable(JsonElement parent, string name, JsonValueKind kind)
    {
        if (parent.ValueKind != JsonValueKind.Object || !parent.TryGetProperty(name, out var value) ||
            (value.ValueKind != kind && value.ValueKind != JsonValueKind.Null))
            throw InvalidData();
        return value;
    }

    private static bool RequiredBoolean(JsonElement parent, string name)
    {
        if (parent.ValueKind != JsonValueKind.Object || !parent.TryGetProperty(name, out var value) ||
            value.ValueKind is not (JsonValueKind.True or JsonValueKind.False))
            throw InvalidData();
        return value.GetBoolean();
    }

    private static OutlookCalendarException InvalidData() =>
        new("Microsoft Graph returned missing or malformed calendar data.");
}
