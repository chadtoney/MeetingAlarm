using MeetingAlarm.Core;
using Xunit;

namespace MeetingAlarm.Core.Tests;

public sealed class MeetingLinkTests
{
    [Theory]
    [InlineData("https://teams.microsoft.com/l/meetup-join/19%3ameeting_test/0?context=%7B%7D")]
    [InlineData("https://teams.microsoft.com/meet/123?p=private")]
    [InlineData("https://teams.cloud.microsoft/meet/123")]
    [InlineData("https://teams.live.com/meet/123")]
    [InlineData("https://teams.microsoft.us/l/meetup-join/abc/0")]
    public void RecognizesTeamsLinksWithoutChangingJoinParameters(string url)
    {
        var link = MeetingLink.Parse(url);
        Assert.True(link.IsTeams);
        Assert.Equal(url, link.Address.AbsoluteUri);
        Assert.Equal("Join Teams meeting", link.JoinLabel);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("javascript:alert(1)")]
    [InlineData("msteams://meet/123")]
    [InlineData("http://teams.microsoft.com/meet/123")]
    [InlineData("file:///C:/test.exe")]
    [InlineData("https://user:password@teams.microsoft.com/meet/123")]
    public void DoesNotExposeUnsafeLinks(string? url)
    {
        Assert.False(MeetingLink.TryCreate(url, out var result));
        Assert.Null(result);
        Assert.Throws<ArgumentException>(() => MeetingLink.Parse(url));
    }

    [Theory]
    [InlineData("https://example.com/meeting")]
    [InlineData("https://teams.microsoft.com.evil.example/meet/123")]
    [InlineData("https://teams.microsoft.com/chat/123")]
    [InlineData("https://teams.live.com/l/meetup-join/abc")]
    [InlineData("https://teams.microsoft.com/meet/")]
    public void KeepsOtherHttpsLinksUsableButDoesNotLabelThemTeams(string url)
    {
        var link = MeetingLink.Parse(url);
        Assert.False(link.IsTeams);
        Assert.Equal("Join meeting", link.JoinLabel);
    }

    [Fact]
    public void PrivacyModeHidesMeetingIdentifiersAndPasscodesButRetainsTheDestination()
    {
        const string url = "https://teams.microsoft.com/meet/private-meeting?p=private-passcode";
        var link = MeetingLink.Parse(url);
        Assert.Equal("Open meeting on teams.microsoft.com", link.DisplayText(true));
        Assert.Equal(url, link.DisplayText(false));
        Assert.Equal(url, link.Address.AbsoluteUri);
    }
}
