using System.Diagnostics;
using MeetingAlarm.ClassicOutlook;
using MeetingAlarm.App.Services;
using MeetingAlarm.Core;
using Microsoft.UI.Text;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Automation;
using Microsoft.UI.Xaml.Controls;
using Windows.Graphics;

namespace MeetingAlarm.App;

internal sealed class MainWindow : Window
{
    public const string WindowTitle = "Meeting Alarm";
    private readonly AppController controller;
    private readonly TextBlock liveStatus = new() { FontSize = 22, TextWrapping = TextWrapping.Wrap };
    private readonly TextBlock nextMeeting = new() { TextWrapping = TextWrapping.Wrap };
    private readonly TextBlock syncStatus = new() { TextWrapping = TextWrapping.Wrap };
    private readonly TextBlock activeSource = new() { TextWrapping = TextWrapping.Wrap };
    private readonly TextBlock sourceHint = new() { TextWrapping = TextWrapping.Wrap };
    private readonly ComboBox sourcePicker = new()
    {
        Header = "Calendar source",
        ItemsSource = new[] { "Classic Outlook (on this PC)", "Microsoft Graph (cloud)" },
        HorizontalAlignment = HorizontalAlignment.Stretch
    };
    private readonly StackPanel graphSetup = new() { Spacing = 12 };
    private readonly InfoBar error = new() { Severity = InfoBarSeverity.Error, IsClosable = true };
    private readonly InfoBar stale = new() { Severity = InfoBarSeverity.Warning, IsClosable = false };
    private readonly StackPanel meetings = new() { Spacing = 8 };
    private readonly Button quiet = new() { Content = "Quiet for 30 minutes" };
    private readonly Button connect = new() { Content = "Connect Outlook" };
    private readonly Button disconnect = new() { Content = "Disconnect calendar" };
    private readonly Button refresh = new() { Content = "Refresh calendar" };
    private readonly TextBox clientId = new() { Header = "Application (client) ID", PlaceholderText = "Approved desktop application's GUID" };
    private readonly TextBox tenant = new() { Header = "Directory (tenant) ID", PlaceholderText = "Tenant GUID or organizations" };
    private readonly CheckBox sound = new() { Content = "Repeat alarm sound until acknowledged" };
    private readonly CheckBox privacy = new() { Content = "Hide meeting titles and full join URLs on alarms (presentation privacy)" };
    private readonly CheckBox tentative = new() { Content = "Include tentative meetings" };
    private readonly CheckBox unanswered = new() { Content = "Include invitations I have not answered" };
    private readonly CheckBox startup = new() { Content = "Start Meeting Alarm when I sign in to Windows" };
    private bool rendering;
    private long renderedRevision = -1;
    public bool CanHideToTray { get; set; }
    public bool Quitting { get; set; }

    public MainWindow(AppController controller, Action quit)
    {
        this.controller = controller;
        Title = WindowTitle;
        AppIcon.Apply(this);
        var root = new StackPanel { Padding = new Thickness(28), Spacing = 18, MaxWidth = 1080 };
        var heading = Row();
        heading.Children.Add(AppIcon.CreateImage(48));
        heading.Children.Add(new TextBlock
        {
            Text = "Meeting Alarm", FontSize = 34, FontWeight = FontWeights.SemiBold,
            VerticalAlignment = VerticalAlignment.Center
        });
        root.Children.Add(heading);
        root.Children.Add(new TextBlock
        {
            Text = "A meeting reminder that stays until you respond.",
            FontSize = 16, TextWrapping = TextWrapping.Wrap
        });
        root.Children.Add(liveStatus);
        root.Children.Add(nextMeeting);
        var actions = Row();
        actions.Children.Add(Button("Test alarm now", () => controller.TestAlarm(0), "TestAlarmNow"));
        actions.Children.Add(Button("Test in 10 seconds", () => controller.TestAlarm(10), "TestAlarmLater"));
        quiet.Click += (_, _) => controller.ToggleQuiet();
        actions.Children.Add(quiet);
        actions.Children.Add(Button("Hide to tray", () =>
        {
            if (CanHideToTray) AppWindow.Hide();
            else controller.ReportError("The tray icon is unavailable. Keep this window open.");
        }));
        actions.Children.Add(Button("Quit", quit));
        root.Children.Add(actions);
        error.CloseButtonClick += (_, _) => controller.ClearError();
        root.Children.Add(error);
        root.Children.Add(stale);
        root.Children.Add(Heading("Alarm behavior"));
        root.Children.Add(new TextBlock
        {
            Text = "Defaults: accepted meetings and meetings you organize with other attendees. Canceled, declined, all-day, and solo calendar events are excluded. Local alarms always qualify.",
            TextWrapping = TextWrapping.Wrap
        });
        root.Children.Add(sound);
        root.Children.Add(privacy);
        root.Children.Add(tentative);
        root.Children.Add(unanswered);
        root.Children.Add(startup);
        root.Children.Add(new TextBlock
        {
            Text = "Alarms need this app running and the PC awake. They do not override mute, volume, audio routing, the lock screen, or exclusive full-screen apps. Audio pauses while Windows is locked. Quiet mode suppresses both the window and sound; there is no automatic Teams call detection.",
            TextWrapping = TextWrapping.Wrap, Opacity = 0.75
        });
        foreach (var checkbox in new[] { sound, privacy, tentative, unanswered })
        {
            checkbox.Checked += OptionsChanged;
            checkbox.Unchecked += OptionsChanged;
        }
        startup.IsChecked = NativeMethods.StartsAtSignIn();
        startup.Checked += StartupChanged;
        startup.Unchecked += StartupChanged;

        var calendar = new StackPanel { Spacing = 12 };
        calendar.Children.Add(activeSource);
        calendar.Children.Add(syncStatus);
        sourcePicker.SelectedIndex = controller.State.ActiveSource == CalendarSource.MicrosoftGraph ? 1 : 0;
        AutomationProperties.SetAutomationId(sourcePicker, "CalendarSourcePicker");
        sourcePicker.SelectionChanged += (_, _) => UpdateSourceSelection();
        calendar.Children.Add(sourcePicker);
        calendar.Children.Add(sourceHint);
        calendar.Children.Add(new TextBlock
        {
            Text = "Both connectors are available, but only one calendar source is active. Changing this selection takes effect only after Connect succeeds. Local alarms remain enabled. Neither connector depends on Scout.",
            TextWrapping = TextWrapping.Wrap
        });
        clientId.Text = controller.State.ClientId;
        tenant.Text = controller.State.TenantId;
        graphSetup.Children.Add(clientId);
        graphSetup.Children.Add(tenant);
        graphSetup.Children.Add(new TextBlock
        {
            Text = "Graph setup: use an approved Entra desktop app registration with redirect URI http://localhost and delegated Microsoft Graph Calendars.Read. No client secret or tenant resources are created by this app. Corporate policy may require ownership registration and administrator consent. Public Microsoft cloud only; not GCC High/DoD.",
            TextWrapping = TextWrapping.Wrap
        });
        calendar.Children.Add(graphSetup);
        var calendarButtons = Row();
        connect.Style = (Style)Application.Current.Resources["AccentButtonStyle"];
        connect.Click += async (_, _) =>
        {
            if (sourcePicker.SelectedIndex == 0)
                await controller.ConnectClassicAsync();
            else
                await controller.ConnectAsync(clientId.Text, tenant.Text);
        };
        AutomationProperties.SetAutomationId(connect, "ConnectCalendar");
        disconnect.Click += async (_, _) => await controller.DisconnectAsync();
        refresh.Click += async (_, _) => await controller.RefreshAsync();
        calendarButtons.Children.Add(connect);
        calendarButtons.Children.Add(refresh);
        calendarButtons.Children.Add(disconnect);
        calendar.Children.Add(calendarButtons);
        UpdateSourceSelection();
        root.Children.Add(new Expander
        {
            Header = "Calendar connection", Content = calendar, IsExpanded = !controller.State.HasCalendarSource,
            HorizontalAlignment = HorizontalAlignment.Stretch, HorizontalContentAlignment = HorizontalAlignment.Stretch
        });
        root.Children.Add(CreateLocalAlarmPanel());
        root.Children.Add(Heading("Upcoming meetings and local alarms"));
        root.Children.Add(new TextBlock
        {
            Text = "Per-occurrence choices: Default follows your rules; Always includes tentative or solo events; Skip excludes the occurrence. Canceled, declined, and all-day events remain excluded.",
            TextWrapping = TextWrapping.Wrap, Opacity = 0.75
        });
        root.Children.Add(meetings);
        var bottom = Row();
        bottom.Children.Add(Button("Open local data folder", () => controller.Run(() =>
            Process.Start(new ProcessStartInfo(ProtectedStateStore.DataDirectory) { UseShellExecute = true }))));
        bottom.Children.Add(new TextBlock
        {
            Text = "Calendar and alarm state are encrypted for your Windows account.",
            VerticalAlignment = VerticalAlignment.Center, TextWrapping = TextWrapping.Wrap
        });
        root.Children.Add(bottom);
        Content = new ScrollViewer
        {
            Content = root, HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled
        };
        AppWindow.Resize(new SizeInt32(1040, 860));
        AppWindow.Closing += (_, args) =>
        {
            if (Quitting)
                return;
            args.Cancel = true;
            if (CanHideToTray) AppWindow.Hide();
            else quit();
        };
        controller.Changed += Render;
        controller.ClockChanged += UpdateClock;
        Closed += (_, _) =>
        {
            controller.Changed -= Render;
            controller.ClockChanged -= UpdateClock;
        };
        Render();
    }

    private void UpdateSourceSelection()
    {
        bool useClassic = sourcePicker.SelectedIndex == 0;
        graphSetup.Visibility = useClassic ? Visibility.Collapsed : Visibility.Visible;
        connect.Content = useClassic ? "Connect classic Outlook" : "Connect Microsoft Graph";
        sourceHint.Text = useClassic
            ? (ClassicOutlookCalendarProvider.IsInstalled
                ? "Classic Outlook is installed. Reads its default calendar locally, including recurring meetings. No Entra app registration is needed. Classic Outlook must have a configured profile and may open or run in the background; its setup or security prompts require your attention. The app does not change Outlook security settings or close Outlook."
                : "Classic Outlook's COM interface was not found. Install and configure classic Outlook first, or select Microsoft Graph. New Outlook alone does not provide this local interface.")
            : "Reads your primary Outlook calendar directly through Microsoft Graph. Does not require classic Outlook, but does require an approved application/client ID and Microsoft sign-in.";
    }

    private UIElement CreateLocalAlarmPanel()
    {
        var panel = new StackPanel { Spacing = 12 };
        var title = new TextBox { Header = "Alarm title", Text = "Meeting" };
        var date = new CalendarDatePicker { Header = "Date", Date = DateTimeOffset.Now, MinDate = DateTimeOffset.Now.Date };
        var time = new TimePicker { Header = "Start time", Time = DateTime.Now.AddMinutes(5).TimeOfDay };
        var duration = new NumberBox { Header = "Duration (minutes)", Value = 30, Minimum = 1, Maximum = 1440, Width = 180 };
        var joinUrl = new TextBox { Header = "Join link (optional HTTPS URL)" };
        panel.Children.Add(new TextBlock
        {
            Text = "Create an alarm on this PC only. Nothing is added to Outlook or sent to anyone.",
            TextWrapping = TextWrapping.Wrap
        });
        panel.Children.Add(title);
        var when = Row();
        when.Children.Add(date);
        when.Children.Add(time);
        when.Children.Add(duration);
        panel.Children.Add(when);
        panel.Children.Add(joinUrl);
        panel.Children.Add(Button("Schedule local alarm", () => controller.Run(() =>
        {
            if (date.Date is null || !double.IsFinite(duration.Value) || duration.Value < 1 || duration.Value > 1440)
                throw new ArgumentException("Choose a date and a duration between 1 and 1440 minutes.");
            var local = DateTime.SpecifyKind(date.Date.Value.Date + time.Time, DateTimeKind.Unspecified);
            if (TimeZoneInfo.Local.IsInvalidTime(local) || TimeZoneInfo.Local.IsAmbiguousTime(local))
                throw new ArgumentException("That local time is skipped or repeated by daylight saving time. Choose an unambiguous time.");
            var start = new DateTimeOffset(local, TimeZoneInfo.Local.GetUtcOffset(local));
            controller.AddLocal(title.Text, start, start.AddMinutes(duration.Value),
                string.IsNullOrWhiteSpace(joinUrl.Text) ? null : joinUrl.Text.Trim());
        })));
        return new Expander
        {
            Header = "Schedule a local alarm", Content = panel,
            HorizontalAlignment = HorizontalAlignment.Stretch, HorizontalContentAlignment = HorizontalAlignment.Stretch
        };
    }

    private void StartupChanged(object sender, RoutedEventArgs args)
    {
        if (rendering)
            return;
        if (!controller.Run(() => NativeMethods.SetStartAtSignIn(startup.IsChecked == true)))
        {
            rendering = true;
            startup.IsChecked = NativeMethods.StartsAtSignIn();
            rendering = false;
        }
    }

    private void OptionsChanged(object sender, RoutedEventArgs args)
    {
        if (!rendering)
            controller.SetOptions(controller.State.Options with
            {
                SoundEnabled = sound.IsChecked == true,
                PrivateDisplay = privacy.IsChecked == true,
                IncludeTentative = tentative.IsChecked == true,
                IncludeUnanswered = unanswered.IsChecked == true
            });
    }

    private void Render()
    {
        rendering = true;
        sound.IsChecked = controller.State.Options.SoundEnabled;
        privacy.IsChecked = controller.State.Options.PrivateDisplay;
        tentative.IsChecked = controller.State.Options.IncludeTentative;
        unanswered.IsChecked = controller.State.Options.IncludeUnanswered;
        connect.IsEnabled = !controller.IsBusy;
        sourcePicker.IsEnabled = !controller.IsBusy;
        clientId.IsEnabled = !controller.IsBusy;
        tenant.IsEnabled = !controller.IsBusy;
        disconnect.IsEnabled = !controller.IsBusy && controller.State.HasCalendarSource;
        refresh.IsEnabled = disconnect.IsEnabled;
        error.Message = controller.Error ?? "";
        error.IsOpen = controller.Error is not null;
        if (renderedRevision != controller.Revision)
        {
            RenderMeetings();
            renderedRevision = controller.Revision;
        }
        rendering = false;
        UpdateClock();
    }

    private void RenderMeetings()
    {
        meetings.Children.Clear();
        var upcoming = controller.AllMeetings().Where(m => m.End > DateTimeOffset.UtcNow).Take(100).ToArray();
        if (upcoming.Length == 0)
        {
            meetings.Children.Add(new TextBlock
            {
                Text = "No upcoming meetings loaded. Connect a calendar source, schedule a local alarm, or use Test in 10 seconds.",
                TextWrapping = TextWrapping.Wrap
            });
            return;
        }
        foreach (var meeting in upcoming)
        {
            var content = new StackPanel { Spacing = 8 };
            content.Children.Add(new TextBlock
            {
                Text = meeting.Subject, FontSize = 17, FontWeight = FontWeights.SemiBold, TextWrapping = TextWrapping.Wrap
            });
            var choice = controller.State.Decisions.FirstOrDefault(d => d.OccurrenceKey == meeting.OccurrenceKey);
            string status = choice?.Dismissed == true ? "Dismissed" :
                choice?.SnoozeUntil > DateTimeOffset.UtcNow ? $"Snoozed until {choice.SnoozeUntil.Value.ToLocalTime():h:mm:ss tt}" :
                AlarmScheduler.IsEligible(meeting, controller.State.Options) ? "Alarm enabled" : "Excluded";
            content.Children.Add(new TextBlock
            {
                Text = $"{meeting.Start.ToLocalTime():ddd, MMM d  h:mm tt} - {meeting.End.ToLocalTime():h:mm tt} | " +
                    $"{(meeting.IsLocal ? "Local" : meeting.Response)} | {status}",
                TextWrapping = TextWrapping.Wrap, Opacity = 0.75
            });
            var controls = Row();
            var preference = new ComboBox { Width = 155, ItemsSource = new[] { "Default", "Always", "Skip" } };
            AutomationProperties.SetName(preference, "Alarm preference for " + meeting.Subject);
            preference.SelectedIndex = (int)controller.State.Options.Overrides.GetValueOrDefault(meeting.OccurrenceKey);
            preference.SelectionChanged += (_, _) =>
            {
                if (!rendering && preference.SelectedIndex >= 0)
                    controller.SetOverride(meeting, (MeetingOverride)preference.SelectedIndex);
            };
            controls.Children.Add(preference);
            if (!string.IsNullOrWhiteSpace(meeting.JoinUrl))
                controls.Children.Add(Button("Join", () => controller.Join(meeting, acknowledge: false)));
            if (meeting.IsLocal)
                controls.Children.Add(Button("Remove", () => controller.RemoveLocal(meeting)));
            content.Children.Add(controls);
            meetings.Children.Add(new Border
            {
                Padding = new Thickness(16), CornerRadius = new CornerRadius(8), BorderThickness = new Thickness(1),
                BorderBrush = (Microsoft.UI.Xaml.Media.Brush)Application.Current.Resources["CardStrokeColorDefaultBrush"],
                Child = content
            });
        }
    }

    private void UpdateClock()
    {
        var now = DateTimeOffset.UtcNow;
        liveStatus.Text = controller.Quiet
            ? $"Quiet until {controller.State.Options.QuietUntil!.Value.ToLocalTime():h:mm tt}"
            : controller.ActiveCount > 0 ? $"{controller.ActiveCount} meeting alarm(s) waiting for you" : "Ready for your next meeting";
        quiet.Content = controller.Quiet ? "Resume alarms" : "Quiet for 30 minutes";
        var next = controller.AllMeetings().FirstOrDefault(m => m.Start > now && AlarmScheduler.IsEligible(m, controller.State.Options)
            && !controller.State.Decisions.Any(d => d.OccurrenceKey == m.OccurrenceKey && d.Dismissed));
        nextMeeting.Text = next is null ? "No future eligible alarms in the loaded calendar." :
            $"Next: {next.Subject} - {next.Start.ToLocalTime():ddd h:mm tt} (in {Math.Ceiling((next.Start - now).TotalMinutes):0} min)";
        syncStatus.Text = controller.IsBusy ? "Connecting or refreshing..." : controller.ConnectionStatus;
        activeSource.Text = $"Active source: {AppController.SourceName(controller.State.ActiveSource)}";
        if (controller.State.LastSync is { } lastSync)
            syncStatus.Text += $" Last successful sync: {lastSync.ToLocalTime():MMM d, h:mm:ss tt}.";
        stale.IsOpen = controller.State.HasCalendarSource &&
            (controller.State.LastSync is null || now - controller.State.LastSync > TimeSpan.FromMinutes(5));
        stale.Message = "Calendar data is stale. Cached alarms still run, but recent changes or cancellations may be missing. Refresh or reconnect the active calendar source.";
    }

    private static TextBlock Heading(string text) => new()
    {
        Text = text, FontSize = 20, FontWeight = FontWeights.SemiBold
    };

    private static StackPanel Row() => new() { Orientation = Orientation.Horizontal, Spacing = 10 };

    private static Button Button(string text, Action action, string? automationId = null)
    {
        var button = new Button { Content = text };
        if (automationId is not null)
            AutomationProperties.SetAutomationId(button, automationId);
        button.Click += (_, _) => action();
        return button;
    }
}
