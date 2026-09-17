using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media.Imaging;

namespace MeetingAlarm.App.Services;

internal static class AppIcon
{
    public static string FilePath => Path.Combine(AppContext.BaseDirectory, "Assets", "MeetingAlarm.ico");

    public static void Apply(Window window) => window.AppWindow.SetIcon(FilePath);

    public static Image CreateImage(double size) => new()
    {
        Width = size, Height = size,
        Source = new BitmapImage(new Uri("ms-appx:///Assets/MeetingAlarm.png"))
    };
}
