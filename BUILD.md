# MyOlap 1.5.0 — Build and Packaging Guide

This document describes how to build MyOlap from source and how the distributable
installer is produced. It accompanies the source package handed to third parties.

- **Product:** MyOlap — Excel OLAP Analytics Add-in
- **Version:** 1.5.0
- **Publisher:** GoLive Systems Ltd
- **Repository:** <https://github.com/emilesd/MS-Excel-ERP-Analytics>

---

## 1. What MyOlap is

MyOlap is an in-process Microsoft Excel add-in packaged as an XLL via
[Excel-DNA](https://excel-dna.net/). It adds a **MyOlap** ribbon tab that provides
multi-dimensional (OLAP-style) analysis directly on a worksheet: models, dimension
hierarchies, drill down/up, pivot, member filtering, POV slicing, data import from
CSV/XLSX/XLS/TXT, and PDF export.

All data is stored locally in a single SQLite file per user. There is no server
component and no network communication.

---

## 2. Build prerequisites

| Requirement | Version | Notes |
|---|---|---|
| Windows | 10 or 11, 64-bit | Build host |
| .NET SDK | 8.0 (x64) | <https://dotnet.microsoft.com/download/dotnet/8.0> |
| Inno Setup | 6.x | Only needed to build the installer: `winget install JRSoftware.InnoSetup` |
| Microsoft Excel | 2016 or later, 64-bit | Only needed to run/test, not to build |

The build restores NuGet packages from nuget.org and requires network access on a
first build.

---

## 3. Building the add-in

```bash
dotnet build MyOlap/MyOlap.csproj -c Release
```

`ExcelDnaPack` runs automatically as part of the Release build and produces
self-contained packed XLLs.

### Build output

Under `MyOlap/bin/Release/net8.0-windows/`:

| File | Purpose |
|---|---|
| `publish/MyOlap-AddIn64-packed.xll` | The add-in for 64-bit Excel; all managed assemblies are embedded |
| `publish/MyOlap-AddIn-packed.xll` | The add-in for 32-bit Excel |
| `runtimes/win-x64/native/e_sqlite3.dll` | Native SQLite engine (x64) |
| `runtimes/win-x86/native/e_sqlite3.dll` | Native SQLite engine (x86) |
| `MyOlap.deps.json`, `MyOlap-AddIn.deps.json`, `MyOlap-AddIn64.deps.json` | Dependency manifests read by the Excel-DNA .NET host |
| `MyOlap.runtimeconfig.json` | Target framework and roll-forward policy |

> The packed XLL embeds managed assemblies but **not** native ones. `e_sqlite3.dll`
> must sit next to the XLL at runtime; `SqliteRepository` installs a
> `DllImportResolver` that loads it from the XLL's own directory.

---

## 4. Building the installer

```powershell
powershell -ExecutionPolicy Bypass -File Packaging\build-installer.ps1
```

The script builds in Release, stages the payload into `Packaging\payload`, downloads
the .NET 8 Desktop Runtime into `Packaging\prereq` (cached after the first run), and
compiles `Packaging\MyOlap.iss` into:

```
Packaging\output\MyOlap-Setup-1.5.0.exe
```

### What the installer does

1. Installs per user into `%LOCALAPPDATA%\MyOlap\AddIn` — **no administrator rights**
   are required for the add-in itself.
2. Detects Office bitness (Click-to-Run `Platform` value, falling back to the Excel
   `InstallRoot` path) and installs the matching XLL and native SQLite build.
3. Installs the bundled .NET 8 Desktop Runtime (x64) only if it is missing. This is
   the only step that elevates.
4. Registers the add-in under
   `HKCU\Software\Microsoft\Office\<ver>\Excel\Options\OPEN<n>`, reusing an existing
   MyOlap slot or taking the first free one so the list stays contiguous (Excel stops
   scanning at the first gap).
5. Clears Excel's `Resiliency\DisabledItems` list so a reinstall is not silently
   suppressed after an earlier crash.

Uninstall removes the files and the registry entry, closes the gap it leaves in the
`OPEN` list, and **preserves the user's `myolap.db` data file**.

---

## 5. Repository layout

```text
MyOlap/                  Excel add-in (net8.0-windows, WinForms + WPF, Excel-DNA)
  AddIn.cs               Excel-DNA entry point
  Core/                  OLAP engine: view state, dimension trees, undo, model management
  Data/                  SQLite persistence, schema, file import
  Ribbon/                Ribbon XML, callbacks, grid rendering via COM reflection
  UI/                    WinForms dialogs
  Reports/               Report building and PDF export
Packaging/               Inno Setup script and installer build script
DbQuery/                 Dev helper: ad-hoc queries against myolap.db
ReadXlsx/                Dev helper: inspect XLSX structure
TestConsole/             Scripted end-to-end tests that drive the engine without Excel
TestData/                Sample CSV/XLSX fixtures
Docs/                    Product documentation
```

Run the test harness with:

```bash
dotnet run --project TestConsole
```

---

## 6. Runtime footprint

| Item | Location |
|---|---|
| Add-in binaries | `%LOCALAPPDATA%\MyOlap\AddIn` |
| User data (SQLite) | `%LOCALAPPDATA%\MyOlap\myolap.db` |
| Excel registration | `HKCU\Software\Microsoft\Office\<ver>\Excel\Options\OPEN<n>` |
| Uninstall entry | `HKCU\...\Windows\CurrentVersion\Uninstall\{8E1F2A64-...}_is1` |

Nothing is written outside the user profile except the .NET runtime prerequisite,
which is a standard Microsoft machine-wide component.

---

## 7. Third-party dependencies

| Package | Version | License |
|---|---|---|
| ExcelDna.AddIn | 1.9.0 | Zlib |
| Microsoft.Data.Sqlite.Core | 8.0.13 | MIT |
| SQLitePCLRaw.bundle_e_sqlite3 | 2.1.10 | Apache-2.0 |
| CsvHelper | 33.1.0 | MS-PL OR Apache-2.0 |
| PdfSharpCore | 1.3.67 | MIT |
| Svg | 3.4.7 | MS-PL |
| **EPPlus** | **8.5.0** | **Polyform Noncommercial 1.0.0** |

> **Licensing note:** EPPlus 5 and later are distributed under the Polyform
> Noncommercial license. Commercial distribution requires a paid EPPlus commercial
> license from EPPlus Software AB, or replacing EPPlus with an alternative
> spreadsheet library. EPPlus is used for reading and writing XLSX/XLS files during
> data and dimension import.

---

## 8. Code signing

The installer and the XLL are currently **unsigned**. Consequences:

- Windows SmartScreen shows a "Windows protected your PC" warning that the user must
  click through.
- On machines with **Smart App Control** enabled (default on many clean Windows 11
  installs), the unsigned installer is **blocked outright with no override**.

Production distribution requires an Authenticode signature — either an OV/EV code
signing certificate or Microsoft Trusted Signing — applied to both
`MyOlap-Setup-<version>.exe` and the packed XLL before packing.
