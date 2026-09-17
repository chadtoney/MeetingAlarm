using System.Security.Cryptography;
using System.Text.Json;
using MeetingAlarm.Core;

namespace MeetingAlarm.App.Services;

public sealed class ProtectedStateStore
{
    public static string DataDirectory { get; } =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "MeetingAlarm");
    private readonly string directory;
    private readonly string path;

    public ProtectedStateStore(string? dataDirectory = null)
    {
        directory = dataDirectory ?? DataDirectory;
        path = Path.Combine(directory, "state.bin");
    }

    public AppState Load()
    {
        if (!File.Exists(path))
            return new();

        var plaintext = ProtectedData.Unprotect(File.ReadAllBytes(path), null, DataProtectionScope.CurrentUser);
        try
        {
            var state = JsonSerializer.Deserialize<AppState>(plaintext)
                ?? throw new InvalidDataException("The saved state is empty.");
            if (state.Version != 1 || state.Options is null || state.LocalMeetings is null ||
                state.CachedMeetings is null || state.Decisions is null || state.Options.Overrides is null ||
                !Enum.IsDefined(state.ActiveSource))
                throw new InvalidDataException("The saved state version or structure is not supported.");
            return state;
        }
        finally
        {
            CryptographicOperations.ZeroMemory(plaintext);
        }
    }

    public void Save(AppState state)
    {
        Directory.CreateDirectory(directory);
        var plaintext = JsonSerializer.SerializeToUtf8Bytes(state);
        byte[] encrypted;
        try
        {
            encrypted = ProtectedData.Protect(plaintext, null, DataProtectionScope.CurrentUser);
        }
        finally
        {
            CryptographicOperations.ZeroMemory(plaintext);
        }

        var temporary = path + ".tmp";
        using (var stream = new FileStream(temporary, FileMode.Create, FileAccess.Write, FileShare.None))
        {
            stream.Write(encrypted);
            stream.Flush(flushToDisk: true);
        }
        File.Move(temporary, path, overwrite: true);
    }
}
