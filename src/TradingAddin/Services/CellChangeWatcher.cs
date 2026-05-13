using System;
using System.Collections.Generic;
using System.Linq;
using System.Timers;
using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;

namespace TradingAddin.Services;

/// <summary>
/// Polls a configurable set of named ranges or absolute cell addresses on a
/// short interval and pushes the values to the back-end whenever they change.
///
/// Why polling rather than the SheetChange event? Excel raises SheetChange
/// only for user edits, not for cells that recompute via volatile formulas
/// (BDP, RTD feeds). For market data, a poll is more robust than the event.
/// </summary>
public class CellChangeWatcher
{
    private readonly AppSettings _settings;
    private readonly BackendClient _api;
    private readonly Dictionary<string, object?> _lastValues = new();
    private System.Timers.Timer? _timer;

    public bool IsRunning => _timer is not null;
    public int  WatchedCellCount { get; private set; }

    public CellChangeWatcher(AppSettings settings, BackendClient api)
    {
        _settings = settings;
        _api = api;
    }

    public void Start()
    {
        if (_timer is not null) return;
        _timer = new System.Timers.Timer(_settings.PollSeconds * 1000.0);
        _timer.Elapsed += OnTick;
        _timer.AutoReset = true;
        _timer.Enabled = true;
    }

    public void Stop()
    {
        if (_timer is null) return;
        _timer.Stop();
        _timer.Dispose();
        _timer = null;
    }

    private void OnTick(object? sender, ElapsedEventArgs e)
    {
        // Must run on the COM-affine main thread.
        ExcelAsyncUtil.QueueAsMacro(SweepCells);
    }

    private void SweepCells()
    {
        try
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;
            // Replace with your own discovery: named ranges, a config sheet,
            // a sheet-level table, etc. The default scans named ranges that
            // start with "TRD_" by convention.
            var watched = app.ActiveWorkbook?.Names
                .Cast<Excel.Name>()
                .Where(n => n.Name.StartsWith("TRD_", StringComparison.OrdinalIgnoreCase))
                .ToList();
            if (watched is null) return;

            WatchedCellCount = watched.Count;

            foreach (var name in watched)
            {
                var range = name.RefersToRange;
                if (range is null) continue;
                var current = range.Value2;
                var key = name.Name;

                if (_lastValues.TryGetValue(key, out var previous) &&
                    Equals(previous, current)) continue;

                _lastValues[key] = current;
                if (current is not null) _api.SendCellValueAsync(range.Address, current);
            }
        }
        catch
        {
            // Any COM exception is ignored. We'll retry on the next tick.
        }
    }
}
