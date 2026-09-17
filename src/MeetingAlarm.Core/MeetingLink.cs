using System.Diagnostics.CodeAnalysis;

namespace MeetingAlarm.Core;

public sealed record MeetingLink
{
    public Uri Address { get; }
    public bool IsTeams { get; }
    public string JoinLabel => IsTeams ? "Join Teams meeting" : "Join meeting";

    private MeetingLink(Uri address)
    {
        Address = address;
        var prefixes = address.IdnHost.ToLowerInvariant() switch
        {
            "teams.microsoft.com" or "teams.cloud.microsoft" or "teams.microsoft.us" =>
                new[] { "/l/meetup-join/", "/meet/" },
            "teams.live.com" => new[] { "/meet/" },
            _ => []
        };
        IsTeams = address.Port == 443 && address.Fragment.Length == 0 &&
            prefixes.Any(prefix => address.AbsolutePath.StartsWith(prefix, StringComparison.Ordinal) &&
                address.AbsolutePath.Length > prefix.Length);
    }

    public string DisplayText(bool privateDisplay) =>
        privateDisplay ? $"Open meeting on {Address.IdnHost}" : Address.AbsoluteUri;

    public static bool TryCreate(string? value, [NotNullWhen(true)] out MeetingLink? link)
    {
        link = null;
        if (!Uri.TryCreate(value, UriKind.Absolute, out var address) ||
            address.Scheme != Uri.UriSchemeHttps || address.UserInfo.Length != 0)
            return false;
        link = new(address);
        return true;
    }

    public static MeetingLink Parse(string? value) =>
        TryCreate(value, out var link) ? link :
            throw new ArgumentException("Join links must be HTTPS web addresses with no embedded credentials.");
}
