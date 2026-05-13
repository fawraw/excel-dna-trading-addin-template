using System.Linq;
using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;

namespace TradingAddin;

/// <summary>
/// Ribbon commands kept separate from the Ribbon class so they can be unit-tested
/// and re-used from menus or hotkeys.
/// </summary>
public static class Commands
{
    public static void SendSelection()
    {
        if (AddIn.Instance is null) return;

        var app = (Excel.Application)ExcelDnaUtil.Application;
        var selection = app.Selection as Excel.Range;
        if (selection is null)
        {
            app.StatusBar = "Trading add-in: nothing selected.";
            return;
        }

        var rows = selection.Rows.Count;
        var cols = selection.Columns.Count;
        var cellsSent = 0;

        for (int r = 1; r <= rows; r++)
        {
            for (int c = 1; c <= cols; c++)
            {
                var cell = (Excel.Range)selection.Cells[r, c];
                if (cell.Value2 is null) continue;
                AddIn.Instance.Api?.SendCellValueAsync(cell.Address, cell.Value2);
                cellsSent++;
            }
        }

        app.StatusBar = $"Trading add-in: sent {cellsSent} cell(s).";
    }

    public static void ToggleWatcher()
    {
        if (AddIn.Instance?.Watcher is null) return;
        if (AddIn.Instance.Watcher.IsRunning) AddIn.Instance.Watcher.Stop();
        else AddIn.Instance.Watcher.Start();

        var app = (Excel.Application)ExcelDnaUtil.Application;
        app.StatusBar = AddIn.Instance.Watcher.IsRunning
            ? "Trading add-in: cell watcher started"
            : "Trading add-in: cell watcher stopped";
    }

    public static void ShowDiagnostics()
    {
        var lines = new[]
        {
            $"Version       : {AddInVersion.Get()}",
            $"Backend URL   : {AddIn.Instance?.Settings.BackendUrl ?? "(unset)"}",
            $"Authenticated : {(AddIn.Instance?.Auth?.IsAuthenticated == true ? "yes" : "no")}",
            $"User          : {AddIn.Instance?.Auth?.UserName ?? "(none)"}",
            $"Watcher       : {(AddIn.Instance?.Watcher?.IsRunning == true ? "running" : "stopped")}",
            $"Watched cells : {AddIn.Instance?.Watcher?.WatchedCellCount ?? 0}",
        };
        System.Windows.Forms.MessageBox.Show(
            string.Join("\n", lines),
            "Trading add-in diagnostics",
            System.Windows.Forms.MessageBoxButtons.OK,
            System.Windows.Forms.MessageBoxIcon.Information);
    }
}
