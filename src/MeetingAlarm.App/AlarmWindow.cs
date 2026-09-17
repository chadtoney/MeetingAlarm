using MeetingAlarm.App.Services;
using MeetingAlarm.Core;
using System.Runtime.InteropServices;
using Microsoft.UI;
using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Automation;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media;
using Windows.Graphics;
using Windows.ApplicationModel.DataTransfer;

namespace MeetingAlarm.App;

internal sealed class AlarmWindow : Window
{
    private readonly AppController controller;
    private Meeting meeting;
    private bool acknowledged;
    private readonly TextBlock subject = new()
    {
        FontSize = 28, FontWeight = Microsoft.UI.Text.FontWeights.SemiBold,
        TextWrapping = TextWrapping.Wrap
    };
    private readonly TextBlock time = new() { FontSize = 16, TextWrapping = TextWrapping.Wrap };
    private readonly TextBlock error = new()
    {
        TextWrapping = TextWrapping.Wrap, Foreground = new SolidColorBrush(Colors.OrangeRed)
    };
    private readonly Button join = new() { Content = "Join meeting", MinWidth = 140 };
    private readonly HyperlinkButton meetingLink = new() { HorizontalAlignment = HorizontalAlignment.Stretch, Padding = new Thickness(0) };
    private readonly TextBlock linkText = new() { TextTrimming = TextTrimming.CharacterEllipsis, MaxWidth = 440 };
    private readonly Button copyLink = new() { Content = "Copy link" };
    private readonly TextBlock linkStatus = new() { TextWrapping = TextWrapping.Wrap, FontSize = 12 };
    private readonly TextBlock missingLink = new() { TextWrapping = TextWrapping.Wrap, Opacity = 0.75 };
    private readonly StackPanel linkPanel = new() { Spacing = 8 };

    public AlarmWindow(AppController controller, Meeting meeting)
    {
        this.controller = controller;
        this.meeting = meeting;
        Title = "Meeting Alarm - reminder";
        AppIcon.Apply(this);
        var root = new StackPanel { Padding = new Thickness(28), Spacing = 18 };
        var heading = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 10 };
        heading.Children.Add(AppIcon.CreateImage(28));
        heading.Children.Add(new TextBlock
        {
            Text = "YOUR MEETING IS STARTING", FontSize = 13,
            Foreground = new SolidColorBrush(Colors.CornflowerBlue), VerticalAlignment = VerticalAlignment.Center
        });
        root.Children.Add(heading);
        root.Children.Add(subject);
        root.Children.Add(time);
        meetingLink.Content = linkText;
        AutomationProperties.SetAutomationId(meetingLink, "AlarmMeetingLink");
        AutomationProperties.SetAutomationId(copyLink, "CopyMeetingLink");
        meetingLink.Click += (_, _) => OpenMeeting();
        copyLink.Click += (_, _) => CopyMeetingLink();
        var linkRow = new Grid { ColumnSpacing = 12 };
        linkRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        linkRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        linkRow.Children.Add(meetingLink);
        Grid.SetColumn(copyLink, 1);
        linkRow.Children.Add(copyLink);
        linkPanel.Children.Add(linkRow);
        linkPanel.Children.Add(linkStatus);
        root.Children.Add(linkPanel);
        root.Children.Add(missingLink);
        root.Children.Add(new TextBlock
        {
            Text = "This reminder stays until you acknowledge it or the meeting ends.",
            TextWrapping = TextWrapping.Wrap
        });
        var buttons = new StackPanel { Orientation = Orientation.Horizontal, Spacing = 10 };
        join.Style = (Style)Application.Current.Resources["AccentButtonStyle"];
        join.Click += (_, _) => OpenMeeting();
        AutomationProperties.SetAutomationId(join, "JoinMeeting");
        buttons.Children.Add(join);
        var snooze = new Button { Content = "Snooze 1 minute" };
        snooze.Click += (_, _) => Acknowledge(snooze: true);
        buttons.Children.Add(snooze);
        var dismiss = new Button { Content = "Dismiss" };
        dismiss.Click += (_, _) => Acknowledge(snooze: false);
        buttons.Children.Add(dismiss);
        root.Children.Add(buttons);
        root.Children.Add(error);
        Content = new ScrollViewer { Content = root, HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled };
        AppWindow.Resize(new SizeInt32(700, 480));
        var area = DisplayArea.GetFromWindowId(AppWindow.Id, DisplayAreaFallback.Primary).WorkArea;
        int offset = controller.ActiveCount > 1 ? Math.Abs(meeting.Id.GetHashCode() % 4) * 24 : 0;
        AppWindow.Move(new PointInt32(area.X + (area.Width - 700) / 2 + offset,
            area.Y + (area.Height - 480) / 2 + offset));
        if (AppWindow.Presenter is OverlappedPresenter presenter)
        {
            presenter.IsAlwaysOnTop = true;
            presenter.IsMaximizable = false;
            presenter.IsMinimizable = false;
        }
        AppWindow.Closing += (_, args) =>
        {
            if (!acknowledged)
            {
                args.Cancel = true;
                Acknowledge(snooze: false);
            }
        };
        UpdateMeeting(meeting, controller.State.Options.PrivateDisplay);
    }

    public void UpdateMeeting(Meeting updated, bool privateDisplay)
    {
        if (meeting.JoinUrl != updated.JoinUrl)
            linkStatus.Text = "";
        meeting = updated;
        subject.Text = privateDisplay ? "It's meeting time." : meeting.Subject;
        time.Text = $"{meeting.Start.ToLocalTime():h:mm tt} - {meeting.End.ToLocalTime():h:mm tt}";
        bool hasLink = MeetingLink.TryCreate(meeting.JoinUrl, out var link);
        join.Visibility = linkPanel.Visibility = hasLink ? Visibility.Visible : Visibility.Collapsed;
        missingLink.Visibility = hasLink ? Visibility.Collapsed : Visibility.Visible;
        if (link is not null)
        {
            join.Content = link.JoinLabel;
            linkText.Text = link.DisplayText(privateDisplay);
            AutomationProperties.SetName(meetingLink, link.JoinLabel + " link");
            ToolTipService.SetToolTip(meetingLink, privateDisplay ? null : link.Address.AbsoluteUri);
        }
        else
        {
            missingLink.Text = string.IsNullOrWhiteSpace(meeting.JoinUrl)
                ? "No online meeting link was included with this event."
                : "This event's meeting link is unavailable: only HTTPS links without embedded credentials are supported.";
        }
    }

    private void OpenMeeting()
    {
        if (!controller.Join(meeting, acknowledge: true))
            error.Text = controller.Error ?? "Could not open the join link.";
    }

    private void CopyMeetingLink()
    {
        try
        {
            var link = MeetingLink.Parse(meeting.JoinUrl);
            var package = new DataPackage { RequestedOperation = DataPackageOperation.Copy };
            package.SetText(link.Address.AbsoluteUri);
            package.SetWebLink(link.Address);
            if (!Clipboard.SetContentWithOptions(package, new ClipboardContentOptions
                { IsAllowedInHistory = false, IsRoamable = false }))
                throw new InvalidOperationException("Windows could not copy the meeting link. Try again.");
            error.Text = "";
            linkStatus.Text = "Link copied. The alarm remains active.";
        }
        catch (Exception exception) when (exception is COMException or InvalidOperationException or ArgumentException)
        {
            error.Text = "Could not copy the meeting link. Check that the Windows clipboard is available and try again.";
        }
    }

    private void Acknowledge(bool snooze)
    {
        if (!controller.Acknowledge(meeting, snooze))
            error.Text = controller.Error ?? "Could not save your choice. Please try again.";
    }

    public void CloseAcknowledged()
    {
        acknowledged = true;
        Close();
    }
}
