using System;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Identity.Client;

namespace TradingAddin.Services;

/// <summary>
/// Microsoft authentication via MSAL (Entra ID / Azure AD).
///
/// Strip this file out entirely if you don't need Entra ID auth -- it's the
/// only piece that pulls in <c>Microsoft.Identity.Client</c>.
/// </summary>
public class AuthService
{
    private readonly AppSettings _settings;
    private readonly IPublicClientApplication? _msal;

    public AuthService(AppSettings settings)
    {
        _settings = settings;
        if (string.IsNullOrWhiteSpace(_settings.TenantId) ||
            string.IsNullOrWhiteSpace(_settings.ClientId))
        {
            // No tenant configured -> run unauthenticated. The back-end is
            // assumed to handle anonymous requests in dev / PoC environments.
            return;
        }

        _msal = PublicClientApplicationBuilder.Create(_settings.ClientId)
            .WithAuthority(AzureCloudInstance.AzurePublic, _settings.TenantId)
            .WithDefaultRedirectUri()
            .Build();
    }

    public bool   IsAuthenticated => CurrentAccount is not null;
    public string UserName => CurrentAccount?.Username ?? "(none)";

    public IAccount? CurrentAccount { get; private set; }

    public async Task<string?> GetAccessTokenAsync()
    {
        if (_msal is null || string.IsNullOrWhiteSpace(_settings.ApiScope)) return null;

        var scopes = new[] { _settings.ApiScope };

        // Try silent first; fall back to interactive if no cached account.
        try
        {
            var accounts = await _msal.GetAccountsAsync();
            CurrentAccount = accounts.FirstOrDefault();

            if (CurrentAccount is not null)
            {
                var silent = await _msal
                    .AcquireTokenSilent(scopes, CurrentAccount)
                    .ExecuteAsync();
                CurrentAccount = silent.Account;
                return silent.AccessToken;
            }
        }
        catch (MsalUiRequiredException) { /* fall through */ }
        catch
        {
            // Any non-UI MSAL error -> treat as not authenticated.
            return null;
        }

        try
        {
            var interactive = await _msal
                .AcquireTokenInteractive(scopes)
                .ExecuteAsync();
            CurrentAccount = interactive.Account;
            return interactive.AccessToken;
        }
        catch
        {
            return null;
        }
    }
}

