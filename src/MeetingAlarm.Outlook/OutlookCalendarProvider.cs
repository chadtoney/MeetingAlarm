using System.ComponentModel;
using System.Diagnostics;
using System.Security;
using System.Security.Cryptography;
using System.Text.Json;
using MeetingAlarm.Core;
using Microsoft.Identity.Client;
using Microsoft.Identity.Client.Extensions.Msal;

namespace MeetingAlarm.Outlook;

public sealed class OutlookCalendarProvider : ICalendarProvider, IDisposable, IAsyncDisposable
{
    private static readonly string[] Scopes = ["https://graph.microsoft.com/Calendars.Read"];
    private readonly IPublicClientApplication _application;
    private readonly MsalCacheHelper _cache;
    private readonly TraceSource _cacheTrace;
    private readonly CacheFailureMonitor _cacheFailure;
    private readonly HttpClient _httpClient;
    private readonly GraphCalendarClient _graph;
    private readonly SemaphoreSlim _gate = new(1, 1);
    private bool _disposed;

    private OutlookCalendarProvider(IPublicClientApplication application, MsalCacheHelper cache,
        TraceSource cacheTrace, CacheFailureMonitor cacheFailure, HttpClient httpClient)
    {
        _application = application;
        _cache = cache;
        _cacheTrace = cacheTrace;
        _cacheFailure = cacheFailure;
        _httpClient = httpClient;
        _graph = new GraphCalendarClient(_httpClient);
    }

    public static async Task<OutlookCalendarProvider> CreateAsync(
        string clientId, string tenantId, string cacheDirectory, CancellationToken cancellationToken)
    {
        ValidateConfiguration(clientId, tenantId, cacheDirectory);
        cancellationToken.ThrowIfCancellationRequested();
        if (!OperatingSystem.IsWindows())
            throw new PlatformNotSupportedException("Outlook token persistence requires Windows DPAPI.");
        var tenant = tenantId.ToLowerInvariant();
        var cacheTrace = new TraceSource("MeetingAlarm.Outlook.Cache", SourceLevels.Error);
        cacheTrace.Listeners.Clear();
        var cacheFailure = new CacheFailureMonitor();
        cacheTrace.Listeners.Add(cacheFailure);
        MsalCacheHelper? cache = null;
        IPublicClientApplication? application = null;
        HttpClient? httpClient = null;
        var cacheRegistered = false;
        var ownershipTransferred = false;
        try
        {
            CreateCacheDirectory(cacheDirectory);
            var storage = new StorageCreationPropertiesBuilder(
                $"outlook-{Guid.Parse(clientId):D}-{tenant}.msalcache", cacheDirectory).Build();
            cache = await MsalCacheHelper.CreateAsync(storage, cacheTrace).ConfigureAwait(false);
            cache.VerifyPersistence();
            cacheFailure.ThrowIfFailed();
            cancellationToken.ThrowIfCancellationRequested();
            application = PublicClientApplicationBuilder.Create(clientId)
                .WithAuthority(AzureCloudInstance.AzurePublic, tenant)
                .WithRedirectUri("http://localhost")
                .Build();
            httpClient = new HttpClient(new HttpClientHandler { AllowAutoRedirect = false })
            {
                Timeout = TimeSpan.FromSeconds(60)
            };
            cacheRegistered = true;
            cache.RegisterCache(application.UserTokenCache);
            cancellationToken.ThrowIfCancellationRequested();
            var provider = new OutlookCalendarProvider(application, cache, cacheTrace, cacheFailure, httpClient);
            ownershipTransferred = true;
            return provider;
        }
        catch (MsalException)
        {
            throw new MsalClientException("outlook_configuration_failed",
                "Outlook could not initialize authentication. Check the approved app configuration.");
        }
        catch (Exception exception) when (IsExpectedCacheFailure(exception))
        {
            throw new OutlookCalendarException(
                "Outlook could not initialize its encrypted Windows token cache. Check the cache directory and app configuration.");
        }
        finally
        {
            if (!ownershipTransferred)
            {
                ReleaseResources(
                    () =>
                    {
                        if (cacheRegistered)
                            cache!.UnregisterCache(application!.UserTokenCache);
                    },
                    httpClient,
                    cacheTrace.Close);
            }
        }
    }

    internal static void ValidateConfiguration(string clientId, string tenantId, string cacheDirectory)
    {
        if (!Guid.TryParseExact(clientId, "D", out var client) || client == Guid.Empty)
            throw new ArgumentException("An approved application client UUID is required.", nameof(clientId));
        if (!(Guid.TryParseExact(tenantId, "D", out var tenant) && tenant != Guid.Empty) &&
            tenantId is not ("organizations" or "common" or "consumers"))
            throw new ArgumentException("Tenant must be a UUID, organizations, common, or consumers.", nameof(tenantId));
        if (string.IsNullOrWhiteSpace(cacheDirectory) ||
            cacheDirectory.IndexOfAny(Path.GetInvalidPathChars()) >= 0 || !Path.IsPathFullyQualified(cacheDirectory))
            throw new ArgumentException("A fully qualified cache directory is required.", nameof(cacheDirectory));
    }

    private static void CreateCacheDirectory(string cacheDirectory)
    {
        try { Directory.CreateDirectory(cacheDirectory); }
        catch (ArgumentException)
        {
            throw new ArgumentException("The Outlook cache directory path is invalid.", nameof(cacheDirectory));
        }
        catch (NotSupportedException)
        {
            throw new ArgumentException("The Outlook cache directory path format is unsupported.", nameof(cacheDirectory));
        }
    }

    public async Task<string> SignInAsync(CancellationToken cancellationToken)
    {
        await _gate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            ThrowIfDisposed();
            VerifyCache();
            var result = await _application.AcquireTokenInteractive(Scopes)
                .WithUseEmbeddedWebView(false)
                .WithPrompt(Prompt.SelectAccount)
                .ExecuteAsync(cancellationToken).ConfigureAwait(false);
            _cacheFailure.ThrowIfFailed();
            var accounts = await _application.GetAccountsAsync().ConfigureAwait(false);
            _cacheFailure.ThrowIfFailed();
            foreach (var account in accounts)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (account.HomeAccountId.Identifier != result.Account.HomeAccountId.Identifier)
                {
                    await _application.RemoveAsync(account).ConfigureAwait(false);
                    _cacheFailure.ThrowIfFailed();
                }
            }
            VerifyCache();
            return string.IsNullOrWhiteSpace(result.Account.Username) ? "Outlook account" : result.Account.Username;
        }
        catch (MsalException)
        {
            throw new MsalClientException("outlook_sign_in_failed",
                "Outlook sign-in failed or was cancelled. Check the approved app registration and consent.");
        }
        catch (Win32Exception)
        {
            throw new MsalClientException("outlook_browser_failed", "Outlook could not launch the system sign-in browser.");
        }
        catch (Exception exception) when (IsExpectedCacheFailure(exception))
        {
            throw new OutlookCalendarException("Outlook sign-in or encrypted token persistence failed.");
        }
        catch (OutlookCalendarException) { throw; }
        catch (HttpRequestException)
        {
            throw new MsalClientException("outlook_sign_in_network_failed", "Outlook sign-in could not reach Microsoft.");
        }
        finally { _gate.Release(); }
    }

    public async Task SignOutAsync(CancellationToken cancellationToken)
    {
        await _gate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            ThrowIfDisposed();
            VerifyCache();
            var accounts = await _application.GetAccountsAsync().ConfigureAwait(false);
            _cacheFailure.ThrowIfFailed();
            foreach (var account in accounts)
            {
                cancellationToken.ThrowIfCancellationRequested();
                await _application.RemoveAsync(account).ConfigureAwait(false);
                _cacheFailure.ThrowIfFailed();
            }
            VerifyCache();
        }
        catch (MsalException)
        {
            throw new MsalClientException("outlook_sign_out_failed",
                "Outlook could not clear the account from its encrypted token cache.");
        }
        catch (Exception exception) when (IsExpectedCacheFailure(exception))
        {
            throw new OutlookCalendarException("Outlook could not clear the account from its encrypted token cache.");
        }
        finally { _gate.Release(); }
    }

    public async Task<IReadOnlyList<Meeting>> GetMeetingsAsync(
        DateTimeOffset from, DateTimeOffset to, CancellationToken cancellationToken)
    {
        GraphCalendarClient.ValidateWindow(from, to);
        await _gate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            ThrowIfDisposed();
            VerifyCache();
            var accounts = (await _application.GetAccountsAsync().ConfigureAwait(false)).ToArray();
            _cacheFailure.ThrowIfFailed();
            if (accounts.Length != 1)
                throw new OutlookSignInRequiredException();
            var result = await _application.AcquireTokenSilent(Scopes, accounts[0])
                .ExecuteAsync(cancellationToken).ConfigureAwait(false);
            _cacheFailure.ThrowIfFailed();
            return await _graph.GetMeetingsAsync(from, to, result.AccessToken, cancellationToken).ConfigureAwait(false);
        }
        catch (MsalUiRequiredException)
        {
            throw new OutlookSignInRequiredException();
        }
        catch (MsalException)
        {
            throw new MsalClientException("outlook_silent_token_failed",
                "Outlook could not acquire a token silently. Try reconnecting Outlook.");
        }
        catch (Exception exception) when (IsExpectedCacheFailure(exception))
        {
            throw new OutlookCalendarException("Outlook calendar access or encrypted token persistence failed.");
        }
        catch (OutlookCalendarException) { throw; }
        catch (HttpRequestException)
        {
            throw new OutlookCalendarException("Outlook calendar access could not reach Microsoft.");
        }
        finally { _gate.Release(); }
    }

    private void VerifyCache()
    {
        _cacheFailure.ThrowIfFailed();
        try
        {
            _cache.VerifyPersistence();
            _cacheFailure.ThrowIfFailed();
        }
        catch (Exception exception) when (IsExpectedCacheFailure(exception))
        {
            throw new OutlookCalendarException("Outlook's encrypted token cache is unavailable. Calendar access was stopped.");
        }
    }

    private void ThrowIfDisposed() => ObjectDisposedException.ThrowIf(_disposed, this);

    internal static bool IsExpectedCacheFailure(Exception exception) =>
        exception is IOException or UnauthorizedAccessException or SecurityException or
            CryptographicException or MsalCachePersistenceException or JsonException;

    internal static void ReleaseResources(Action unregisterCache, IDisposable? httpClient, Action closeTrace)
    {
        try
        {
            try { unregisterCache(); }
            finally
            {
                try { httpClient?.Dispose(); }
                finally { closeTrace(); }
            }
        }
        catch (MsalException)
        {
            throw new MsalClientException("outlook_cleanup_failed", "Outlook could not detach its authentication cache.");
        }
        catch (Exception exception) when (IsExpectedCacheFailure(exception))
        {
            throw new OutlookCalendarException("Outlook could not finish releasing its local connector resources.");
        }
    }

    public void Dispose() => DisposeAsync().AsTask().GetAwaiter().GetResult();

    public async ValueTask DisposeAsync()
    {
        await _gate.WaitAsync().ConfigureAwait(false);
        try
        {
            if (_disposed)
                return;
            _disposed = true;
            ReleaseResources(() => _cache.UnregisterCache(_application.UserTokenCache), _httpClient, _cacheTrace.Close);
        }
        finally { _gate.Release(); }
    }
}
