using Xunit;

namespace MeetingAlarm.ClassicOutlook.Tests;

public class TeamsLinksTests
{
    [Theory]
    [InlineData("https://teams.microsoft.com/l/meetup-join/19%3ameeting_abc%40thread.v2/0?context=%7B%7D")]
    [InlineData("https://teams.microsoft.com/meet/12345?p=abc")]
    [InlineData("https://teams.cloud.microsoft/meet/12345")]
    [InlineData("https://teams.microsoft.us/l/meetup-join/abc/0")]
    [InlineData("https://teams.live.com/meet/12345")]
    public void AcceptsRecognizedJoinRoutes(string url) => Assert.Equal(url, TeamsLinks.Validate(url));

    [Theory]
    [InlineData("http://teams.microsoft.com/meet/123")]
    [InlineData("https://evil.example/meet/123")]
    [InlineData("https://teams.microsoft.com.evil.example/meet/123")]
    [InlineData("https://teams.microsoft.com@evil.example/meet/123")]
    [InlineData("https://user@teams.microsoft.com/meet/123")]
    [InlineData("https://teams.microsoft.com:444/meet/123")]
    [InlineData("https://teams.microsoft.com/chat/123")]
    [InlineData("https://teams.microsoft.com/l/meetup-join/")]
    [InlineData("https://teams.live.com/l/meetup-join/abc")]
    [InlineData("https://teams.microsoft.com/meet/123#https://evil.example")]
    [InlineData("javascript:alert(1)")]
    public void RejectsUnsafeOrUnrecognizedUrls(string url) => Assert.Null(TeamsLinks.Validate(url));

    [Fact]
    public void IgnoresArbitraryFirstHttpsUrlAndDecodesHtml()
    {
        Assert.Equal("https://teams.microsoft.com/meet/123?p=abc&x=1",
            TeamsLinks.Extract("""Privacy https://example.com <a href="https://teams.microsoft.com/meet/123?p=abc&amp;x=1">Join</a>"""));
    }

    [Fact]
    public void BoundedTextFailsExplicitly() =>
        Assert.Throws<ClassicOutlookException>(() => TeamsLinks.Extract(new string('a', 1_048_577)));

    [Fact]
    public void OrdinaryBodyWithoutJoinLinkReturnsNull() =>
        Assert.Null(TeamsLinks.Extract("Meet at reception. https://example.com"));
}
