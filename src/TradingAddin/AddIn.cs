using System;
using ExcelDna.Integration;
using ExcelDna.IntelliSense;
using Microsoft.Extensions.Logging;
using TradingAddin.Logging;
using TradingAddin.Services;

namespace TradingAddin;

/// <summary>
/// Entry point of the trading add-in. Wires up authentication, the cell watcher,
/// the back-end API client, and the ribbon. Inspect each service for the
/// extension points; the template ships with non-functional stubs you can swap.
/// </summary>
public class AddIn : IExcelAddIn
{
    public static AddIn? Instance { get; private set; }

    public AppSettings Settings { get; private set; } = new();
    public AuthService? Auth   { get; private set; }
    public BackendClient? Api  { get; private set; }
    public CellChangeWatcher? Watcher { get; private set; }

    private ILogger<AddIn>? _log;

    public void AutoOpen()
    {
        Instance = this;

        _log = AddInLogging.CreateLogger<AddIn>();
        _log.LogInformation("TradingAddin loaded -- {Version}", AddInVersion.Get());

        Settings = AppSettings.Load();
        Auth     = new AuthService(Settings);
        Api      = new BackendClient(Settings, Auth);
        Watcher  = new CellChangeWatcher(Settings, Api);

        IntelliSenseServer.Install();
        Watcher.Start();
    }

    public void AutoClose()
    {
        _log?.LogInformation("TradingAddin unloading");
        Watcher?.Stop();
        IntelliSenseServer.Uninstall();
    }
}
