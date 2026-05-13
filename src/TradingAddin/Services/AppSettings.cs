using System;
using System.IO;
using System.Text.Json;

namespace TradingAddin.Services;

/// <summary>
/// Per-user settings loaded from <c>%AppData%/TradingAddin/settings.json</c>.
/// All fields have safe defaults so the add-in loads even on a fresh install.
/// </summary>
public class AppSettings
{
    public string BackendUrl  { get; set; } = "http://localhost:8000";
    public string TenantId    { get; set; } = "";
    public string ClientId    { get; set; } = "";
    public string ApiScope    { get; set; } = "";
    public int    PollSeconds { get; set; } = 5;
    public bool   AutoStartWatcher { get; set; } = true;
    public string LogPath     { get; set; } = "";

    public static string SettingsPath =>
        Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "TradingAddin",
            "settings.json");

    public static AppSettings Load()
    {
        try
        {
            if (!File.Exists(SettingsPath)) return new AppSettings();
            var json = File.ReadAllText(SettingsPath);
            return JsonSerializer.Deserialize<AppSettings>(json) ?? new AppSettings();
        }
        catch
        {
            // On any deserialisation error, fall back to defaults rather than refuse to load.
            return new AppSettings();
        }
    }

    public void Save()
    {
        Directory.CreateDirectory(Path.GetDirectoryName(SettingsPath)!);
        var json = JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(SettingsPath, json);
    }
}
