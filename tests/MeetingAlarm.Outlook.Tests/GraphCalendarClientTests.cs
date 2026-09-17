using System.Net;
using System.Net.Http.Headers;
using System.Text.Json;
using System.Text.Json.Nodes;
using MeetingAlarm.Core;
using Xunit;

namespace MeetingAlarm.Outlook.Tests;

public sealed class GraphCalendarClientTests
{
    private static readonly DateTimeOffset From = new(2026, 9, 8, 0, 0, 0, TimeSpan.Zero);
    private static readonly DateTimeOffset To = From.AddDays(1);

    internal static JsonObject Event(string id = "immutable-occurrence-1") => new()
    {
        ["id"] = id,
        ["subject"] = "Synthetic planning meeting",
        ["start"] = new JsonObject { ["dateTime"] = "2026-09-08T14:00:00.0000000", ["timeZone"] = "UTC" },
        ["end"] = new JsonObject { ["dateTime"] = "2026-09-08T15:00:00.0000000", ["timeZone"] = "UTC" },
        ["responseStatus"] = new JsonObject { ["response"] = "accepted" },
        ["isOrganizer"] = false,
        ["isCancelled"] = false,
        ["isAllDay"] = false,
        ["attendees"] = new JsonArray(new JsonObject()),
        ["onlineMeeting"] = new JsonObject { ["joinUrl"] = "https://teams.microsoft.com/l/meetup-join/synthetic" }
    };

    private static string Page(JsonObject? meeting = null, string? nextLink = null)
    {
        var page = new JsonObject { ["value"] = meeting is null ? new JsonArray() : new JsonArray(meeting) };
        if (nextLink is not null)
            page["@odata.nextLink"] = nextLink;
        return page.ToJsonString();
    }

    private static Meeting Parse(JsonObject meeting)
    {
        using var doc = JsonDocument.Parse(Page(meeting));
        return Assert.Single(GraphCalendarClient.ParsePage(doc.RootElement).Meetings);
    }

    [Theory]
    [InlineData("none", MeetingResponse.None)]
    [InlineData("notResponded", MeetingResponse.None)]
    [InlineData("accepted", MeetingResponse.Accepted)]
    [InlineData("tentativelyAccepted", MeetingResponse.Tentative)]
    [InlineData("declined", MeetingResponse.Declined)]
    public void MapsOwnResponse(string response, MeetingResponse expected)
    {
        var data = Event();
        data["responseStatus"]!["response"] = response;
        Assert.Equal(expected, Parse(data).Response);
    }

    [Fact]
    public void MapsOrganizerFlagsUtcAndJoinUrl()
    {
        var data = Event();
        data["isOrganizer"] = true;
        data["isCancelled"] = true;
        data["isAllDay"] = true;
        var meeting = Parse(data);
        Assert.Equal(MeetingResponse.Organizer, meeting.Response);
        Assert.True(meeting.HasOtherAttendees);
        Assert.True(meeting.IsCancelled);
        Assert.True(meeting.IsAllDay);
        Assert.False(meeting.IsLocal);
        Assert.Equal(TimeSpan.Zero, meeting.Start.Offset);
        Assert.Equal(From.AddHours(14), meeting.Start);
        Assert.Equal(From.AddHours(15), meeting.End);
        Assert.Equal("https://teams.microsoft.com/l/meetup-join/synthetic", meeting.JoinUrl);
        Assert.Equal("immutable-occurrence-1", meeting.Id);
    }

    [Theory]
    [InlineData(true, false)]
    [InlineData(false, true)]
    public void AttendeePresenceIncludesOrganizerForInvitees(bool organizer, bool hasOthers)
    {
        var data = Event();
        data["isOrganizer"] = organizer;
        data["attendees"] = new JsonArray();
        Assert.Equal(hasOthers, Parse(data).HasOtherAttendees);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void DefaultsOnlyBlankSubject(string? subject)
    {
        var data = Event();
        data["subject"] = subject;
        Assert.Equal("Untitled meeting", Parse(data).Subject);
    }

    [Fact]
    public void AllowsExplicitlyNullOnlineMeeting()
    {
        var data = Event();
        data["onlineMeeting"] = null;
        Assert.Null(Parse(data).JoinUrl);
    }

    [Theory]
    [InlineData("id")]
    [InlineData("subject")]
    [InlineData("start")]
    [InlineData("end")]
    [InlineData("responseStatus")]
    [InlineData("isOrganizer")]
    [InlineData("isCancelled")]
    [InlineData("isAllDay")]
    [InlineData("attendees")]
    [InlineData("onlineMeeting")]
    public void RejectsMissingSelectedFields(string field)
    {
        var data = Event();
        data.Remove(field);
        Assert.Throws<OutlookCalendarException>(() => Parse(data));
    }

    [Theory]
    [InlineData("2026-09-08T14:00:00", "UTC", true)]
    [InlineData("2026-09-08T14:00:00.1234567Z", "UTC", true)]
    [InlineData("2026-09-08T14:00:00+00:00", "UTC", true)]
    [InlineData("2026-09-08T14:00:00+01:00", "UTC", false)]
    [InlineData("2026-09-08T14:00:00", "Pacific Standard Time", false)]
    [InlineData("2026-09-08", "UTC", false)]
    [InlineData("not-a-time", "UTC", false)]
    [InlineData("2026-02-30T14:00:00", "UTC", false)]
    [InlineData("2026-09-08T16:00:00", "UTC", false)]
    public void ValidatesUtcTimes(string time, string zone, bool valid)
    {
        var data = Event();
        data["start"]!["dateTime"] = time;
        data["start"]!["timeZone"] = zone;
        if (valid)
            Assert.Equal(TimeSpan.Zero, Parse(data).Start.Offset);
        else
            Assert.Throws<OutlookCalendarException>(() => Parse(data));
    }

    [Theory]
    [InlineData("futureResponse")]
    [InlineData("organizer")]
    public void RejectsUnknownOrContradictoryResponse(string response)
    {
        var data = Event();
        data["responseStatus"]!["response"] = response;
        Assert.Throws<OutlookCalendarException>(() => Parse(data));
    }

    [Theory]
    [InlineData("javascript:alert(1)")]
    [InlineData("file:///C:/data")]
    [InlineData("http://teams.microsoft.com/meeting")]
    [InlineData("https://user:password@teams.microsoft.com/meeting")]
    [InlineData("")]
    public void RejectsUnsafeJoinUrls(string url)
    {
        var data = Event();
        data["onlineMeeting"]!["joinUrl"] = url;
        Assert.Throws<OutlookCalendarException>(() => Parse(data));
    }

    [Fact]
    public void OccurrenceKeysStayStableAndSeparateRecurringOccurrences()
    {
        var first = Parse(Event());
        Assert.Equal(first.OccurrenceKey, Parse(Event()).OccurrenceKey);
        var laterData = Event();
        laterData["start"]!["dateTime"] = "2026-09-09T14:00:00";
        laterData["end"]!["dateTime"] = "2026-09-09T15:00:00";
        Assert.NotEqual(first.OccurrenceKey, Parse(laterData).OccurrenceKey);
    }

    [Fact]
    public async Task PaginatesWithSelectedFieldsUtcAndImmutableIds()
    {
        const string next = "https://graph.microsoft.com/v1.0/me/calendarView?$skiptoken=synthetic";
        using var handler = new StubHandler((request, count) =>
        {
            Assert.Equal("Bearer", request.Headers.Authorization!.Scheme);
            Assert.Equal("synthetic-token", request.Headers.Authorization.Parameter);
            var prefer = Assert.Single(request.Headers.GetValues("Prefer"));
            Assert.Contains("outlook.timezone=\"UTC\"", prefer);
            Assert.Contains("IdType=\"ImmutableId\"", prefer);
            Assert.Equal(HttpMethod.Get, request.Method);
            if (count == 1)
            {
                var uri = Uri.UnescapeDataString(request.RequestUri!.AbsoluteUri);
                Assert.Contains("startDateTime=2026-09-08T00:00:00.0000000Z", uri);
                Assert.Contains("$select=id,subject,start,end,responseStatus,isOrganizer,isCancelled,isAllDay,attendees,onlineMeeting", uri);
                Assert.DoesNotContain(",body", uri);
                return Json(Page(Event("first"), next));
            }
            Assert.Equal(next, request.RequestUri!.AbsoluteUri);
            return Json(Page(Event("second")));
        });
        using var http = new HttpClient(handler);
        var result = await new GraphCalendarClient(http).GetMeetingsAsync(From, To, "synthetic-token", default);
        Assert.Equal(["first", "second"], result.Select(m => m.Id));
        Assert.Equal(2, handler.Count);
    }

    [Theory]
    [InlineData("https://evil.example/v1.0/me/calendarView")]
    [InlineData("http://graph.microsoft.com/v1.0/me/calendarView")]
    [InlineData("https://graph.microsoft.com.evil.example/v1.0/me/calendarView")]
    [InlineData("https://graph.microsoft.com:444/v1.0/me/calendarView")]
    [InlineData("https://user@graph.microsoft.com/v1.0/me/calendarView")]
    [InlineData("https://graph.microsoft.com/v1.0/me/calendarView#fragment")]
    [InlineData("https://graph.microsoft.com/v1.0/me/messages")]
    [InlineData("/v1.0/me/calendarView?$skiptoken=2")]
    public async Task RejectsUnsafePaginationWithoutSendingBearer(string next)
    {
        using var handler = new StubHandler((_, _) => Json(Page(Event(), next)));
        using var http = new HttpClient(handler);
        await Assert.ThrowsAsync<OutlookCalendarException>(() =>
            new GraphCalendarClient(http).GetMeetingsAsync(From, To, "synthetic-token", default));
        Assert.Equal(1, handler.Count);
    }

    [Theory]
    [InlineData(401)]
    [InlineData(403)]
    [InlineData(404)]
    [InlineData(429)]
    [InlineData(500)]
    [InlineData(503)]
    [InlineData(302)]
    public async Task RejectsHttpFailuresWithoutPrivateDetails(int status)
    {
        using var handler = new StubHandler((_, _) => new HttpResponseMessage((HttpStatusCode)status)
        {
            Content = new StringContent("private-calendar-subject secret-token"),
            Headers = { Location = new Uri("https://evil.example") }
        });
        using var http = new HttpClient(handler);
        var error = await Assert.ThrowsAsync<OutlookCalendarException>(() =>
            new GraphCalendarClient(http).GetMeetingsAsync(From, To, "secret-token", default));
        Assert.Equal((HttpStatusCode)status, error.StatusCode);
        Assert.Contains($"HTTP {status}", error.Message);
        Assert.DoesNotContain("private", error.ToString());
        Assert.DoesNotContain("secret-token", error.ToString());
        Assert.Null(error.InnerException);
        Assert.Equal(1, handler.Count);
    }

    [Fact]
    public async Task FailsEntireCallWhenLaterPageFails()
    {
        using var handler = new StubHandler((_, count) => count == 1
            ? Json(Page(Event(), "https://graph.microsoft.com/v1.0/me/calendarView?$skiptoken=2"))
            : new HttpResponseMessage(HttpStatusCode.Forbidden));
        using var http = new HttpClient(handler);
        await Assert.ThrowsAsync<OutlookCalendarException>(() =>
            new GraphCalendarClient(http).GetMeetingsAsync(From, To, "synthetic-token", default));
        Assert.Equal(2, handler.Count);
    }

    [Fact]
    public async Task FailsEntireCallWhenLaterPageIsMalformed()
    {
        using var handler = new StubHandler((_, count) => count == 1
            ? Json(Page(Event(), "https://graph.microsoft.com/v1.0/me/calendarView?$skiptoken=2"))
            : Json("{\"value\":[{\"id\":\"missing-required-data\"}]}"));
        using var http = new HttpClient(handler);
        await Assert.ThrowsAsync<OutlookCalendarException>(() =>
            new GraphCalendarClient(http).GetMeetingsAsync(From, To, "synthetic-token", default));
        Assert.Equal(2, handler.Count);
    }

    [Fact]
    public async Task RejectsDuplicateOccurrencesAcrossPages()
    {
        using var handler = new StubHandler((_, count) => Json(Page(Event(),
            count == 1 ? "https://graph.microsoft.com/v1.0/me/calendarView?$skiptoken=2" : null)));
        using var http = new HttpClient(handler);
        await Assert.ThrowsAsync<OutlookCalendarException>(() =>
            new GraphCalendarClient(http).GetMeetingsAsync(From, To, "synthetic-token", default));
        Assert.Equal(2, handler.Count);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task StopsRepeatedAndRunawayPages(bool repeated)
    {
        using var handler = new StubHandler((_, count) => Json(Page(nextLink:
            $"https://graph.microsoft.com/v1.0/me/calendarView?$skiptoken={(repeated ? 1 : count)}")));
        using var http = new HttpClient(handler);
        await Assert.ThrowsAsync<OutlookCalendarException>(() =>
            new GraphCalendarClient(http).GetMeetingsAsync(From, To, "synthetic-token", default));
        Assert.Equal(repeated ? 2 : GraphCalendarClient.MaximumPages, handler.Count);
    }

    [Fact]
    public async Task RetriesBoundedRetryAfter()
    {
        using var handler = new StubHandler((_, count) => count < 3
            ? new HttpResponseMessage(HttpStatusCode.TooManyRequests)
            {
                Headers = { RetryAfter = new RetryConditionHeaderValue(TimeSpan.Zero) }
            }
            : Json(Page(Event())));
        using var http = new HttpClient(handler);
        Assert.Single(await new GraphCalendarClient(http).GetMeetingsAsync(From, To, "synthetic-token", default));
        Assert.Equal(3, handler.Count);
    }

    [Fact]
    public async Task DoesNotRetryIndefinitely()
    {
        using var handler = new StubHandler((_, _) => new HttpResponseMessage(HttpStatusCode.ServiceUnavailable)
        {
            Headers = { RetryAfter = new RetryConditionHeaderValue(TimeSpan.Zero) }
        });
        using var http = new HttpClient(handler);
        await Assert.ThrowsAsync<OutlookCalendarException>(() =>
            new GraphCalendarClient(http).GetMeetingsAsync(From, To, "synthetic-token", default));
        Assert.Equal(3, handler.Count);
    }

    [Fact]
    public async Task LongRetryAfterFailsRatherThanRetryingEarly()
    {
        using var handler = new StubHandler((_, _) => new HttpResponseMessage(HttpStatusCode.TooManyRequests)
        {
            Headers = { RetryAfter = new RetryConditionHeaderValue(TimeSpan.FromHours(1)) }
        });
        using var http = new HttpClient(handler);
        await Assert.ThrowsAsync<OutlookCalendarException>(() =>
            new GraphCalendarClient(http).GetMeetingsAsync(From, To, "synthetic-token", default));
        Assert.Equal(1, handler.Count);
    }

    [Fact]
    public async Task CancellationInterruptsRetryDelay()
    {
        using var cancellation = new CancellationTokenSource();
        using var handler = new StubHandler((_, _) =>
        {
            cancellation.Cancel();
            return new HttpResponseMessage(HttpStatusCode.TooManyRequests)
            {
                Headers = { RetryAfter = new RetryConditionHeaderValue(TimeSpan.FromSeconds(30)) }
            };
        });
        using var http = new HttpClient(handler);
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            new GraphCalendarClient(http).GetMeetingsAsync(From, To, "synthetic-token", cancellation.Token));
        Assert.Equal(1, handler.Count);
    }

    [Theory]
    [InlineData("{}")]
    [InlineData("{\"value\":null}")]
    [InlineData("{\"value\":[],\"@odata.nextLink\":null}")]
    [InlineData("{\"value\":[null]}")]
    [InlineData("not-json-private-meeting")]
    public async Task RejectsMalformedResponsesWithoutContents(string json)
    {
        using var handler = new StubHandler((_, _) => Json(json));
        using var http = new HttpClient(handler);
        var error = await Assert.ThrowsAsync<OutlookCalendarException>(() =>
            new GraphCalendarClient(http).GetMeetingsAsync(From, To, "synthetic-token", default));
        Assert.DoesNotContain("private-meeting", error.ToString());
    }

    [Fact]
    public async Task RejectsOversizedResponses()
    {
        using var handler = new StubHandler((_, _) => Json(new string(' ', 8 * 1024 * 1024 + 1)));
        using var http = new HttpClient(handler);
        await Assert.ThrowsAsync<OutlookCalendarException>(() =>
            new GraphCalendarClient(http).GetMeetingsAsync(From, To, "synthetic-token", default));
    }

    [Fact]
    public async Task SanitizesNetworkErrors()
    {
        using var handler = new StubHandler((_, _) => throw new HttpRequestException("private-data"));
        using var http = new HttpClient(handler);
        var error = await Assert.ThrowsAsync<OutlookCalendarException>(() =>
            new GraphCalendarClient(http).GetMeetingsAsync(From, To, "synthetic-token", default));
        Assert.DoesNotContain("private-data", error.ToString());
    }

    [Fact]
    public async Task CancellationDoesNotSendRequest()
    {
        using var handler = new StubHandler((_, _) => Json(Page()));
        using var http = new HttpClient(handler);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            new GraphCalendarClient(http).GetMeetingsAsync(From, To, "synthetic-token", cancellation.Token));
        Assert.Equal(0, handler.Count);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    [InlineData(32)]
    public async Task RejectsUnboundedOrInvertedWindowsBeforeRequest(int days)
    {
        using var handler = new StubHandler((_, _) => Json(Page()));
        using var http = new HttpClient(handler);
        await Assert.ThrowsAsync<ArgumentOutOfRangeException>(() =>
            new GraphCalendarClient(http).GetMeetingsAsync(From, From.AddDays(days), "synthetic-token", default));
        Assert.Equal(0, handler.Count);
    }

    private static HttpResponseMessage Json(string json) => new(HttpStatusCode.OK)
    {
        Content = new StringContent(json, System.Text.Encoding.UTF8, "application/json")
    };

    private sealed class StubHandler(Func<HttpRequestMessage, int, HttpResponseMessage> respond) : HttpMessageHandler
    {
        internal int Count { get; private set; }
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            return Task.FromResult(respond(request, ++Count));
        }
    }
}
