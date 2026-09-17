using System.ComponentModel;
using System.Runtime.InteropServices;

namespace MeetingAlarm.App.Services;

internal sealed class TrayIcon : IDisposable
{
    private NativeMethods.NotifyIconData data;
    private readonly NativeMethods.SubclassProc callback;
    private readonly Action show, pause, quit;
    private readonly Action<string> reportError;
    private readonly nint icon;
    private bool disposed;
    private readonly uint taskbarCreated = NativeMethods.RegisterWindowMessage("TaskbarCreated");

    public TrayIcon(nint hwnd, Action show, Action pause, Action quit, Action<string> reportError)
    {
        this.show = show;
        this.pause = pause;
        this.quit = quit;
        this.reportError = reportError;
        callback = WindowProc;
        uint dpi = NativeMethods.GetDpiForWindow(hwnd);
        icon = NativeMethods.LoadImage(0, AppIcon.FilePath, 1,
            NativeMethods.GetSystemMetricsForDpi(49, dpi), NativeMethods.GetSystemMetricsForDpi(50, dpi), 0x0010);
        if (icon == 0)
            throw new Win32Exception(Marshal.GetLastWin32Error(), "Could not load the Meeting Alarm tray icon.");
        data = new()
        {
            Size = (uint)Marshal.SizeOf<NativeMethods.NotifyIconData>(),
            Hwnd = hwnd,
            Id = 1,
            Flags = 1 | 2 | 4,
            CallbackMessage = NativeMethods.TrayMessage,
            Icon = icon,
            Tip = "Meeting Alarm - click to open; right-click for controls",
            Info = "",
            InfoTitle = ""
        };
        if (!NativeMethods.SetWindowSubclass(hwnd, callback, 1, 0))
        {
            NativeMethods.DestroyIcon(icon);
            throw new Win32Exception("Could not initialize tray interactions.");
        }
        if (!NativeMethods.Shell_NotifyIcon(0, ref data))
        {
            NativeMethods.RemoveWindowSubclass(hwnd, callback, 1);
            NativeMethods.DestroyIcon(icon);
            throw new Win32Exception("Could not add the Meeting Alarm tray icon.");
        }
    }

    private nint WindowProc(nint hwnd, uint message, nuint wParam, nint lParam, nuint id, nuint reference)
    {
        if (message == taskbarCreated)
        {
            if (!NativeMethods.Shell_NotifyIcon(0, ref data))
            {
                reportError("The tray icon could not be restored after Explorer restarted. Keep this window open.");
                show();
            }
        }
        else if (message == NativeMethods.ShowMessage)
            show();
        else if (message == NativeMethods.TrayMessage)
        {
            if ((uint)lParam is 0x0202 or 0x0203)
                show();
            else if ((uint)lParam == 0x0205)
                ShowMenu(hwnd);
        }
        return NativeMethods.DefSubclassProc(hwnd, message, wParam, lParam);
    }

    private void ShowMenu(nint hwnd)
    {
        var menu = NativeMethods.CreatePopupMenu();
        if (menu == 0)
        {
            reportError("Cannot open the tray menu. Use the main window controls.");
            show();
            return;
        }
        try
        {
            NativeMethods.AppendMenu(menu, 0, 1, "Open Meeting Alarm");
            NativeMethods.AppendMenu(menu, 0, 2, "Quiet for 30 minutes / resume");
            NativeMethods.AppendMenu(menu, 0, 3, "Quit");
            NativeMethods.GetCursorPos(out var point);
            NativeMethods.SetForegroundWindow(hwnd);
            switch (NativeMethods.TrackPopupMenu(menu, 0x0100 | 0x0002, point.X, point.Y, 0, hwnd, 0))
            {
                case 1: show(); break;
                case 2: pause(); break;
                case 3: quit(); break;
            }
        }
        finally { NativeMethods.DestroyMenu(menu); }
    }

    public void Dispose()
    {
        if (disposed)
            return;
        disposed = true;
        NativeMethods.Shell_NotifyIcon(2, ref data);
        NativeMethods.RemoveWindowSubclass(data.Hwnd, callback, 1);
        NativeMethods.DestroyIcon(icon);
    }
}
