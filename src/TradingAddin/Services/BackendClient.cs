using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;

namespace TradingAddin.Services;

/// <summary>
/// Thin client that pushes captured Excel values to a back-end HTTP service.
/// Replace the JSON shape and the endpoint path with whatever your API expects.
/// </summary>
public class BackendClient
{
    private readonly AppSettings _settings;
    private readonly AuthService _auth;
    private readonly HttpClient _http = new();

    public BackendClient(AppSettings settings, AuthService auth)
    {
        _settings = settings;
        _auth = auth;
        _http.BaseAddress = new Uri(_settings.BackendUrl);
        _http.Timeout = TimeSpan.FromSeconds(10);
    }

    public async Task SendCellValueAsync(string cellAddress, object value)
    {
        try
        {
            await EnsureAuthHeaderAsync();
            var payload = new
            {
                cell = cellAddress,
                value,
                ts = DateTimeOffset.UtcNow.ToString("O"),
            };
            // POST /ingest/cell -- adapt to your API.
            var resp = await _http.PostAsJsonAsync("/ingest/cell", payload);
            resp.EnsureSuccessStatusCode();
        }
        catch (Exception ex)
        {
            // Logging only. The add-in never throws into Excel.
            Logging.AddInLogging.CreateLogger<BackendClient>().LogError(ex, "Send failed");
        }
    }

    private async Task EnsureAuthHeaderAsync()
    {
        var token = await _auth.GetAccessTokenAsync();
        if (token is null) return;
        _http.DefaultRequestHeaders.Authorization =
            new AuthenticationHeaderValue("Bearer", token);
    }
}

