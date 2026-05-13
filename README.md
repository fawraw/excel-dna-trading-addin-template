# excel-dna-trading-addin-template

[![Build](https://github.com/fawraw/excel-dna-trading-addin-template/actions/workflows/build.yml/badge.svg)](https://github.com/fawraw/excel-dna-trading-addin-template/actions/workflows/build.yml)
[![License: MIT](https://img.shields.io/github/license/fawraw/excel-dna-trading-addin-template)](LICENSE)
[![Latest release](https://img.shields.io/github/v/release/fawraw/excel-dna-trading-addin-template)](https://github.com/fawraw/excel-dna-trading-addin-template/releases)
![.NET](https://img.shields.io/badge/.NET-6.0-512BD4)
![Excel-DNA](https://img.shields.io/badge/Excel--DNA-1.8-217346)

A production-grade starting point for a trading-floor Excel add-in built on [Excel-DNA](https://excel-dna.net/). Comes with the wiring you actually need (ribbon, MSAL auth, HTTP back-end client, cell watcher, settings file, packaging scripts) and none of the AI boilerplate you don't.

Designed for the case where:

- A trader has Excel open all day with live prices in named cells (Bloomberg BDP, RTD feeds, internal data feeds).
- You want a small native add-in that pushes those cell values to a back-end (HTTP, Kafka, ZeroMQ) so other systems can see them.
- You may or may not need Microsoft Entra ID authentication.

## What's in the box

| File                                                          | Purpose                                          |
|---------------------------------------------------------------|--------------------------------------------------|
| `src/TradingAddin/TradingAddin.csproj`                        | .NET 6 Windows project, Excel-DNA SDK pinned    |
| `src/TradingAddin/AddIn.cs`                                   | `IExcelAddIn` entry point                       |
| `src/TradingAddin/Ribbon.cs`                                  | Custom "Trading" ribbon tab                     |
| `src/TradingAddin/Commands.cs`                                | Ribbon button handlers (testable in isolation)  |
| `src/TradingAddin/Services/AppSettings.cs`                    | Per-user settings under `%AppData%/TradingAddin` |
| `src/TradingAddin/Services/AuthService.cs`                    | MSAL public client (Entra ID / Azure AD)        |
| `src/TradingAddin/Services/BackendClient.cs`                  | HttpClient against `/ingest/cell`               |
| `src/TradingAddin/Services/CellChangeWatcher.cs`              | Timer-driven sweep of named ranges              |
| `src/TradingAddin/Logging/AddInLogging.cs`                    | `ILoggerFactory` with a console sink            |
| `src/TradingAddin/TradingAddin-AddIn.dna`                     | Excel-DNA add-in descriptor                     |
| `scripts/build.ps1`                                           | One-command build + package                     |
| `scripts/install.ps1`                                         | Copy to `%AppData%/Microsoft/AddIns`            |
| `docs/architecture.md`                                        | Load sequence, threading notes, extension points |

## Build

You need .NET 6 SDK and Windows.

```powershell
./scripts/build.ps1
# Output: dist/TradingAddin-AddIn-packed.xll, dist/TradingAddin-AddIn64-packed.xll
```

## Install in Excel

```powershell
./scripts/install.ps1
```

Then in Excel: File -> Options -> Add-ins -> Manage: Excel Add-ins -> Go -> tick `TradingAddin-AddIn64-packed` -> OK.

A "Trading" tab appears in the ribbon with three buttons:

- **Send selection**: push the currently-selected cell range to the back-end.
- **Toggle cell watcher**: start / stop the timer that polls named cells.
- **Diagnostics**: status dump in a modal (version, auth state, watched cell count).

## Configure

Edit `%AppData%/TradingAddin/settings.json` (created on first run):

```json
{
  "BackendUrl":  "https://api.example.internal",
  "TenantId":    "<your-entra-tenant-guid>",
  "ClientId":    "<your-entra-app-guid>",
  "ApiScope":    "api://<your-api-app-id>/.default",
  "PollSeconds": 5,
  "AutoStartWatcher": true,
  "LogPath": ""
}
```

Leave `TenantId` / `ClientId` empty if you don't need authentication; the back-end client will run anonymously.

## Cell discovery

By default, the cell watcher picks up **named ranges whose name starts with `TRD_`**. Example:

| Name         | Refers to            |
|--------------|----------------------|
| `TRD_ZAR_5Y` | `=Sheet1!$B$5`       |
| `TRD_ZAR_10Y`| `=Sheet1!$B$6`       |
| `TRD_FX_ZAR` | `=Sheet1!$D$2`       |

Every `PollSeconds`, the watcher reads each one. If its value changed since last poll, the new value is sent to `POST /ingest/cell`.

Change the prefix or replace the discovery logic entirely in `Services/CellChangeWatcher.cs`.

## Why polling rather than the SheetChange event?

Excel raises `SheetChange` only when a user **types** into a cell. Cells that recompute via volatile formulas (`BDP`, `RTD`, custom DDE) don't fire the event. For market-data capture, a poll is more robust. Tune `PollSeconds` against your back-end's throughput.

## Distribute to a fleet

Three common paths:

1. **Group policy**: push the `.xll` to `%AppData%/Microsoft/AddIns/` and the registry key `HKCU\Software\Microsoft\Office\<version>\Excel\Add-in Manager` that references it.
2. **MSI / MSIX installer**: wrap the .xll in WiX or Advanced Installer.
3. **SCCM** / Intune: script the same registry edits as group policy.

`install.ps1` does the per-user variant manually. For an enterprise rollout, lift the logic into your standard package format.

## Hardening before shipping

See [`docs/architecture.md`](docs/architecture.md) for the production checklist (signing, logging, retry, resilience).

## License

MIT. See [LICENSE](LICENSE).
