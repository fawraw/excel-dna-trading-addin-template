# Architecture

This template's structure, the load sequence, and where to extend it.

## Layout

```
src/TradingAddin/
    TradingAddin.csproj         .NET 6 windows-only project, Excel-DNA SDK
    TradingAddin-AddIn.dna      Excel-DNA add-in descriptor

    AddIn.cs                    IExcelAddIn entry point (AutoOpen / AutoClose)
    AddInVersion.cs             Single-method helper around Assembly version
    Ribbon.cs                   Custom UI XML and the ribbon callbacks
    Commands.cs                 Logic for ribbon actions, testable in isolation

    Logging/
        AddInLogging.cs         ILoggerFactory wrapper (console sink by default)

    Services/
        AppSettings.cs          JSON-serialised settings under %AppData%/TradingAddin
        AuthService.cs          MSAL public client, silent + interactive flows
        BackendClient.cs        HttpClient against /ingest/cell
        CellChangeWatcher.cs    Timer-driven sweep of named ranges starting with TRD_
```

## Load sequence

```
Excel start
    -> Excel-DNA shim loads TradingAddin.dll
        -> AddIn.AutoOpen()
            -> AppSettings.Load()       (%AppData%/TradingAddin/settings.json)
            -> new AuthService(settings) (MSAL public client if Tenant/ClientId set)
            -> new BackendClient(settings, auth)
            -> new CellChangeWatcher(settings, api)
            -> IntelliSenseServer.Install()
            -> Watcher.Start()
        -> Ribbon.GetCustomUI() called by Excel to draw the tab
```

## Extension points

- **Discovery of which cells to watch**: see `CellChangeWatcher.SweepCells`.
  Default = named ranges whose name starts with `TRD_`. Replace with a config
  sheet, an XML island, a database lookup, etc.

- **What gets sent to the back-end**: see `BackendClient.SendCellValueAsync`.
  Default payload is `{cell, value, ts}`. Adapt to your API's schema.

- **Authentication**: `AuthService` uses MSAL with a public client (interactive
  consent). Switch to `ConfidentialClientApplicationBuilder` for a confidential
  client, or rip the file out entirely if you don't need auth (and delete the
  `Microsoft.Identity.Client` reference from the csproj).

- **Logging**: the default sink is the console (useful when Excel is launched
  from a terminal). Add a file sink in `AddInLogging.cs` for production
  diagnostics.

- **Ribbon**: edit the inline XML in `Ribbon.GetCustomUI` to add groups,
  toggle buttons, dropdowns, edit boxes. See the
  [Microsoft Office RibbonX reference](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-customui/d842006e-5d22-43e1-bb6c-c83fc81550b8)
  for the supported elements.

## Threading notes

Excel COM is single-threaded apartment. Any code that touches an Excel object
must run on the main thread. The watcher dispatches its sweep back to the
main thread via `ExcelAsyncUtil.QueueAsMacro`. Don't call `app.Range[...]`
from a background thread directly -- you'll get an `0x80010001 (RPC_E_CALL_REJECTED)`
exception or worse.

HTTP I/O and logging are fine off-thread.

## Hardening for production

The template intentionally stops at a working skeleton. Before shipping:

- Pin all NuGet versions and produce a `packages.lock.json`.
- Sign the `.xll` and the wrapper installer with an Authenticode certificate.
- Replace the console logger with a rolling file sink under
  `%LocalAppData%/TradingAddin/logs/`.
- Add resilient retry on `BackendClient` (Polly is the obvious choice).
- Decide what happens when the back-end is unreachable for minutes at a time
  (drop, queue locally, surface a UI warning).
- Decide what happens when the user closes Excel mid-send.
- Write end-to-end tests with a mock COM workbook (NSubstitute, Moq).
