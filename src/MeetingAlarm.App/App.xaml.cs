using System.Diagnostics;
using MeetingAlarm.App.Services;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Windows.Graphics;

namespace MeetingAlarm.App;

public partial class App : Application
{
    private Mutex? instance;
    private MainWindow? main;
    private Window? recovery;
    private AppController? controller;
    private TrayIcon? tray;
    private bool handlingFatalError;

    public App()
    {
        UnhandledException += (_, args) =>
        {
            args.Handled = true;
            if (handlingFatalError)
            {
                Exit();
                return;
            }
            handlingFatalError = true;
            ShowRecovery($"Unexpected application error (0x{args.Exception.HResult:X8}): {args.Message}");
        };
        InitializeComponent();
    }

    protected override void OnLaunched(LaunchActivatedEventArgs args)
    {
        instance = new Mutex(true, @"Local\MeetingAlarm", out bool first);
        if (!first)
        {
            var existing = NativeMethods.FindWindow(null, MainWindow.WindowTitle);
            if (existing != 0)
                NativeMethods.PostMessage(existing, NativeMethods.ShowMessage, 0, 0);
            instance.Dispose();
            instance = null;
            Exit();
            return;
        }

        try
        {
            var store = new ProtectedStateStore();
            controller = new(store, store.Load());
            main = new(controller, Quit);
            main.Activate();
            try
            {
                tray = new(WinRT.Interop.WindowNative.GetWindowHandle(main),
                    () => NativeMethods.Show(main), controller.ToggleQuiet, Quit, controller.ReportError);
                main.CanHideToTray = true;
            }
            catch (Exception exception) when (AppController.IsExpected(exception))
            {
                controller.ReportError(exception.Message + " Keep the main window open.");
            }
            controller.Start();
            if (main.CanHideToTray && Environment.GetCommandLineArgs().Contains("--background"))
                main.AppWindow.Hide();
        }
        catch (Exception exception) when (AppController.IsExpected(exception))
        {
            ShowRecovery(exception.Message);
        }
    }

    private void ShowRecovery(string message)
    {
        controller?.Dispose();
        controller = null;
        tray?.Dispose();
        tray = null;
        if (main is not null)
        {
            main.Quitting = true;
            main.Close();
            main = null;
        }
        recovery = new Window { Title = "Meeting Alarm - startup error" };
        var panel = new StackPanel { Padding = new Thickness(24), Spacing = 16 };
        panel.Children.Add(new TextBlock { Text = "Meeting Alarm could not start", FontSize = 24 });
        panel.Children.Add(new TextBlock
        {
            Text = message + "\n\nExisting data has not been replaced. Alarms are NOT running. " +
                "If state.bin is damaged or belongs to a different Windows account, back it up before removing it and restarting.",
            TextWrapping = TextWrapping.Wrap
        });
        var close = new Button { Content = "Quit" };
        close.Click += (_, _) => Quit();
        panel.Children.Add(close);
        recovery.Content = panel;
        recovery.AppWindow.Resize(new SizeInt32(640, 330));
        recovery.Closed += (_, _) => Quit();
        recovery.Activate();
    }

    private void Quit()
    {
        tray?.Dispose();
        tray = null;
        controller?.Dispose();
        controller = null;
        if (main is not null)
        {
            main.Quitting = true;
            main.Close();
            main = null;
        }
        instance?.Dispose();
        instance = null;
        Exit();
    }
}
