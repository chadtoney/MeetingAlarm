using System.ComponentModel;
using System.Diagnostics;
using System.Security;
using System.Security.Cryptography;
using System.Text.Json;
using MeetingAlarm.ClassicOutlook;
using MeetingAlarm.Core;
using MeetingAlarm.Outlook;
using Microsoft.Identity.Client;
using Microsoft.UI.Dispatching;

namespace MeetingAlarm.App.Services;

internal sealed class AppController : IDisposable
{
    private readonly ProtectedStateStore store;
    private readonly AlarmAudio audio;
    private readonly DispatcherQueueTimer timer;
    private readonly CancellationTokenSource lifetime = new();
    private readonly Dictionary<string, AlarmWindow> windows = [];
    private AlarmScheduler scheduler;
    private OutlookCalendarProvider? outlook;
    private ClassicOutlookCalendarProvider? classic;
    private DateTimeOffset nextSync = DateTimeOffset.MinValue;
    private bool busy, disposed, audioFailed;
    private long revision;
    private DateTimeOffset previousTick = DateTimeOffset.UtcNow;
    private IReadOnlyList<Meeting> active = [];

    public AppState State { get; private set; }
    public string? Error { get; private set; }
    public string ConnectionStatus { get; private set; } = "Outlook not connected. Local alarms are ready.";
    public bool IsBusy => busy;
    public int ActiveCount => active.Count;
    public long Revision => revision;
    public bool Quiet => State.Options.QuietUntil > DateTimeOffset.UtcNow;
    public bool DesktopAvailable { get; private set; } = true;
    public event Action? Changed;
    public event Action? ClockChanged;

    public AppController(ProtectedStateStore store, AppState state)
    {
        this.store = store;
        State = state;
        if (state.HasCalendarSource)
            ConnectionStatus = $"{SourceName(state.ActiveSource)} selected. Waiting for calendar refresh.";
        scheduler = new(state.Decisions);
        audio = new();
        timer = DispatcherQueue.GetForCurrentThread().CreateTimer();
        timer.Interval = TimeSpan.FromSeconds(1);
        timer.Tick += (_, _) => Tick();
    }

    public void Start()
    {
        timer.Start();
        Tick();
    }

    public IReadOnlyList<Meeting> AllMeetings() => State.AllMeetings();

    public static string SourceName(CalendarSource source) => source switch
    {
        CalendarSource.ClassicOutlook => "Classic Outlook",
        CalendarSource.MicrosoftGraph => "Microsoft Graph",
        _ => "Local alarms only"
    };

    public void ReportError(string error)
    {
        Error = error;
        Changed?.Invoke();
    }

    public void ClearError()
    {
        Error = null;
        Changed?.Invoke();
    }

    private void Commit(AppState candidate)
    {
        store.Save(candidate);
        State = candidate;
        revision++;
        Changed?.Invoke();
    }

    public bool Run(Action action)
    {
        try
        {
            action();
            return true;
        }
        catch (Exception exception) when (IsExpected(exception))
        {
            ReportError(exception.Message);
            return false;
        }
    }

    public static bool IsExpected(Exception exception) => exception is
        IOException or UnauthorizedAccessException or CryptographicException or JsonException or
        HttpRequestException or MsalException or ArgumentException or InvalidOperationException or
        Win32Exception or SecurityException or FormatException;

    public void SetOptions(AlarmOptions options)
    {
        if (Run(() => Commit(State with { Options = options })))
        {
            audioFailed = false;
            Tick();
        }
    }

    public void ToggleQuiet() => SetOptions(State.Options with
    {
        QuietUntil = Quiet ? null : DateTimeOffset.UtcNow.AddMinutes(30)
    });

    public void SetOverride(Meeting meeting, MeetingOverride preference)
    {
        var overrides = new Dictionary<string, MeetingOverride>(State.Options.Overrides);
        if (preference == MeetingOverride.Default)
            overrides.Remove(meeting.OccurrenceKey);
        else
            overrides[meeting.OccurrenceKey] = preference;
        SetOptions(State.Options with { Overrides = overrides });
    }

    public void AddLocal(string title, DateTimeOffset start, DateTimeOffset end, string? joinUrl = null)
    {
        if (string.IsNullOrWhiteSpace(title))
            throw new ArgumentException("Give the local alarm a title.");
        if (start < DateTimeOffset.UtcNow.AddSeconds(-1) || end <= start)
            throw new ArgumentException("Choose a future start time and an end time after the start.");
        if (!string.IsNullOrWhiteSpace(joinUrl))
            MeetingLink.Parse(joinUrl);
        var meeting = new Meeting("local:" + Guid.NewGuid(), title.Trim(), start, end,
            MeetingResponse.Accepted, false, JoinUrl: joinUrl, IsLocal: true);
        Commit(State with
        {
            LocalMeetings = State.LocalMeetings.Where(m => m.End > DateTimeOffset.UtcNow).Append(meeting).ToList()
        });
        Tick();
    }

    public void TestAlarm(int delaySeconds)
    {
        // A test intentionally uses the same scheduling/persistence path as real alarms.
        var start = DateTimeOffset.UtcNow.AddSeconds(delaySeconds);
        Run(() => AddLocal("Test meeting alarm", start, start.AddMinutes(10)));
        if (Quiet)
            ReportError("Test scheduled, but quiet mode is active. Resume alarms to see it.");
    }

    public void RemoveLocal(Meeting meeting)
    {
        if (Run(() => Commit(State with
            { LocalMeetings = State.LocalMeetings.Where(m => m.Id != meeting.Id).ToList() })))
            Tick();
    }

    public bool Acknowledge(Meeting meeting, bool snooze)
    {
        if (snooze)
            scheduler.Snooze(meeting, DateTimeOffset.UtcNow, TimeSpan.FromMinutes(1));
        else
            scheduler.Dismiss(meeting);

        if (!Run(() => Commit(State with { Decisions = scheduler.Decisions.ToList() })))
        {
            scheduler = new(State.Decisions);
            return false;
        }
        Tick();
        return true;
    }

    public bool Join(Meeting meeting, bool acknowledge)
    {
        if (!Run(() =>
        {
            if (string.IsNullOrWhiteSpace(meeting.JoinUrl))
                throw new InvalidOperationException("This meeting has no online join link.");
            var link = MeetingLink.Parse(meeting.JoinUrl);
            Process.Start(new ProcessStartInfo(link.Address.AbsoluteUri) { UseShellExecute = true });
        }))
            return false;
        return !acknowledge || Acknowledge(meeting, snooze: false);
    }

    public async Task ConnectAsync(string clientId, string tenantId)
    {
        if (busy)
        {
            ReportError("A calendar operation is already running. Wait for it to finish.");
            return;
        }
        busy = true;
        Changed?.Invoke();
        OutlookCalendarProvider? candidate = null;
        try
        {
            candidate = await OutlookCalendarProvider.CreateAsync(clientId.Trim(), tenantId.Trim(),
                Path.Combine(ProtectedStateStore.DataDirectory, "Identity"), lifetime.Token);
            var account = await candidate.SignInAsync(lifetime.Token);
            var now = DateTimeOffset.UtcNow;
            var meetings = await candidate.GetMeetingsAsync(now.AddHours(-12), now.AddDays(7), lifetime.Token);
            lifetime.Token.ThrowIfCancellationRequested();
            Commit(State.WithCalendar(CalendarSource.MicrosoftGraph, meetings, DateTimeOffset.UtcNow) with
            {
                ClientId = clientId.Trim(), TenantId = tenantId.Trim()
            });
            outlook?.Dispose();
            outlook = candidate;
            candidate = null;
            ConnectionStatus = $"Microsoft Graph connected: {account}";
            nextSync = DateTimeOffset.UtcNow.AddMinutes(2);
            Error = null;
        }
        catch (OperationCanceledException)
        {
            if (!disposed)
                ReportError("Calendar sign-in or retrieval was canceled or timed out.");
        }
        catch (Exception exception) when (IsExpected(exception))
        {
            ReportError("Outlook connection failed: " + exception.Message);
        }
        finally
        {
            candidate?.Dispose();
            busy = false;
            if (!disposed)
            {
                Changed?.Invoke();
                Tick();
            }
        }
    }

    public async Task ConnectClassicAsync()
    {
        if (busy)
        {
            ReportError("A calendar operation is already running. Wait for it to finish.");
            return;
        }
        busy = true;
        Changed?.Invoke();
        try
        {
            // Keep the worker after a timeout: COM cannot be force-canceled, so
            // repeated Connect clicks must not create new blocked Outlook workers.
            classic ??= new ClassicOutlookCalendarProvider();
            var now = DateTimeOffset.UtcNow;
            var meetings = await classic.GetMeetingsAsync(now.AddHours(-12), now.AddDays(7), lifetime.Token);
            lifetime.Token.ThrowIfCancellationRequested();
            Commit(State.WithCalendar(CalendarSource.ClassicOutlook, meetings, DateTimeOffset.UtcNow));
            outlook?.Dispose();
            outlook = null;
            ConnectionStatus = "Classic Outlook connected through the default local calendar. Refreshes every 2 minutes.";
            nextSync = DateTimeOffset.UtcNow.AddMinutes(2);
            Error = null;
        }
        catch (OperationCanceledException)
        {
            if (!disposed)
                ReportError("Classic Outlook connection was canceled or timed out. Your previous source is unchanged.");
        }
        catch (Exception exception) when (IsExpected(exception))
        {
            ReportError("Classic Outlook connection failed: " + exception.Message);
        }
        finally
        {
            busy = false;
            if (!disposed)
            {
                Changed?.Invoke();
                Tick();
            }
        }
    }

    public async Task DisconnectAsync()
    {
        if (busy)
        {
            ReportError("Wait for the current calendar operation before disconnecting.");
            return;
        }
        busy = true;
        Changed?.Invoke();
        try
        {
            if (outlook is null && !string.IsNullOrEmpty(State.ClientId))
                outlook = await OutlookCalendarProvider.CreateAsync(State.ClientId, State.TenantId,
                    Path.Combine(ProtectedStateStore.DataDirectory, "Identity"), lifetime.Token);
            if (outlook is not null)
                await outlook.SignOutAsync(lifetime.Token);
            Commit(State.WithoutCalendar());
            outlook?.Dispose();
            outlook = null;
            // The COM worker stays dormant so a previously timed-out call cannot
            // be multiplied by disconnect/reconnect. It never polls on its own.
            ConnectionStatus = "Calendar disconnected. Local alarms remain enabled; classic Outlook was not closed.";
            Error = null;
        }
        catch (OperationCanceledException)
        {
            if (!disposed)
                ReportError("Disconnect was canceled.");
        }
        catch (Exception exception) when (IsExpected(exception))
        {
            ReportError("Could not finish disconnecting: " + exception.Message);
        }
        finally
        {
            busy = false;
            if (!disposed)
            {
                Changed?.Invoke();
                Tick();
            }
        }
    }

    public async Task RefreshAsync()
    {
        if (busy || !State.HasCalendarSource || disposed)
            return;
        busy = true;
        nextSync = DateTimeOffset.UtcNow.AddMinutes(2);
        Changed?.Invoke();
        try
        {
            ICalendarProvider provider;
            if (State.ActiveSource == CalendarSource.ClassicOutlook)
            {
                classic ??= new ClassicOutlookCalendarProvider();
                provider = classic;
            }
            else
            {
                outlook ??= await OutlookCalendarProvider.CreateAsync(State.ClientId, State.TenantId,
                    Path.Combine(ProtectedStateStore.DataDirectory, "Identity"), lifetime.Token);
                provider = outlook;
            }
            var now = DateTimeOffset.UtcNow;
            var meetings = await provider.GetMeetingsAsync(now.AddHours(-12), now.AddDays(7), lifetime.Token);
            lifetime.Token.ThrowIfCancellationRequested();
            Commit(State.WithCalendar(State.ActiveSource, meetings, DateTimeOffset.UtcNow));
            ConnectionStatus = $"{SourceName(State.ActiveSource)} connected. Calendar refreshes every 2 minutes.";
            Error = null;
        }
        catch (OperationCanceledException)
        {
            if (!disposed)
                ReportError("Calendar refresh timed out. Cached alarms remain active.");
        }
        catch (Exception exception) when (IsExpected(exception))
        {
            ConnectionStatus = $"{SourceName(State.ActiveSource)} refresh failed. Cached alarms remain active.";
            ReportError(exception is MsalUiRequiredException
                ? "Microsoft Graph needs sign-in. Select Microsoft Graph and click Connect to renew access."
                : "Calendar refresh failed: " + exception.Message);
        }
        finally
        {
            busy = false;
            if (!disposed)
                Changed?.Invoke();
        }
    }

    private void Tick()
    {
        if (disposed)
            return;
        var now = DateTimeOffset.UtcNow;
        if (now - previousTick > TimeSpan.FromSeconds(30) || now < previousTick.AddSeconds(-5))
            nextSync = DateTimeOffset.MinValue;
        previousTick = now;
        DesktopAvailable = NativeMethods.IsDesktopAvailable();
        active = scheduler.Tick(now, AllMeetings(), State.Options);
        var byKey = active.ToDictionary(m => m.OccurrenceKey);
        foreach (var (key, window) in windows.ToArray())
        {
            if (!byKey.TryGetValue(key, out var meeting))
            {
                windows.Remove(key);
                window.CloseAcknowledged();
            }
            else
                window.UpdateMeeting(meeting, State.Options.PrivateDisplay);
        }
        foreach (var meeting in active)
        {
            if (!windows.ContainsKey(meeting.OccurrenceKey))
            {
                var window = new AlarmWindow(this, meeting);
                windows.Add(meeting.OccurrenceKey, window);
                window.Activate();
            }
        }
        if (!audioFailed)
            audioFailed = !Run(() => audio.SetPlaying(
                active.Count > 0 && State.Options.SoundEnabled && !Quiet && DesktopAvailable));
        ClockChanged?.Invoke();
        if (State.HasCalendarSource && now >= nextSync && !busy)
            _ = RefreshAsync();
    }

    public void Dispose()
    {
        if (disposed)
            return;
        disposed = true;
        timer.Stop();
        lifetime.Cancel();
        foreach (var window in windows.Values.ToArray())
            window.CloseAcknowledged();
        windows.Clear();
        audio.Dispose();
        outlook?.Dispose();
        classic?.Dispose();
        lifetime.Dispose();
    }
}
