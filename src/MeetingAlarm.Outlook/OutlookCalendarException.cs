using System.Net;
using Microsoft.Identity.Client;

namespace MeetingAlarm.Outlook;

public class OutlookCalendarException : HttpRequestException
{
    public OutlookCalendarException(string message, HttpStatusCode? statusCode = null)
        : base(message, null, statusCode) { }
}

public sealed class OutlookSignInRequiredException : MsalUiRequiredException
{
    public OutlookSignInRequiredException()
        : base("outlook_sign_in_required", "Outlook sign-in is required. Connect Outlook explicitly to continue.") { }
}
