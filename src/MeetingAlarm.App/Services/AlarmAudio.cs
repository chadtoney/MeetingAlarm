using System.Text;

namespace MeetingAlarm.App.Services;

internal sealed class AlarmAudio : IDisposable
{
    private readonly string path;
    private bool playing;

    public AlarmAudio()
    {
        Directory.CreateDirectory(ProtectedStateStore.DataDirectory);
        path = Path.Combine(ProtectedStateStore.DataDirectory, "alarm.wav");
        CreateTone(path);
    }

    public void SetPlaying(bool enabled)
    {
        if (playing == enabled)
            return;
        if (!NativeMethods.PlaySound(enabled ? path : null, 0, enabled ? 0x00020000u | 0x0001u | 0x0008u | 0x0002u : 0))
            throw new InvalidOperationException("Windows could not play the alarm sound. Check your audio device.");
        playing = enabled;
    }

    private static void CreateTone(string target)
    {
        const int rate = 22050, seconds = 3;
        var sampleCount = rate * seconds;
        using var writer = new BinaryWriter(File.Create(target), Encoding.ASCII);
        writer.Write(Encoding.ASCII.GetBytes("RIFF"));
        writer.Write(36 + sampleCount * 2);
        writer.Write(Encoding.ASCII.GetBytes("WAVEfmt "));
        writer.Write(16);
        writer.Write((short)1);
        writer.Write((short)1);
        writer.Write(rate);
        writer.Write(rate * 2);
        writer.Write((short)2);
        writer.Write((short)16);
        writer.Write(Encoding.ASCII.GetBytes("data"));
        writer.Write(sampleCount * 2);
        for (var index = 0; index < sampleCount; index++)
        {
            double time = (double)index / rate;
            double beat = time % 0.45;
            double envelope = time < 1.35 && beat < 0.22
                ? Math.Min(1, beat / 0.015) * Math.Min(1, (0.22 - beat) / 0.03) : 0;
            double frequency = (int)(time / 0.45) % 2 == 0 ? 740 : 988;
            writer.Write((short)(Math.Sin(2 * Math.PI * frequency * time) * envelope * 0.28 * short.MaxValue));
        }
    }

    public void Dispose()
    {
        if (playing)
            NativeMethods.PlaySound(null, 0, 0);
    }
}
