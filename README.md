# MyOlap - Excel OLAP Analytics Add-in

An MS Excel add-in for OLAP-style analytics. Pull data from ERP/Excel/CSV/TXT files, clean it, visualize it, and generate reports — all within Excel.

---

## Features

- **OLAP Grid** — Multi-dimensional data view with row/column axes
- **Drill Down / Drill Up** — Navigate hierarchical data structures
- **Pivot (Swap Row/Col)** — Instantly transpose row and column dimensions
- **Keep / Remove Selected** — Filter members on any axis
- **Undo** — Step back through view state history
- **Pick Member** — Slice data by any dimension/member
- **Data Loading** — Import from CSV, XLSX, XLS, TXT files
- **Dimension Loading** — Bulk-load dimension members from spreadsheets
- **PDF Export** — Generate formatted PDF reports
- **Up to 12 Dimensions** — 5 pre-defined (View, Year, Period, Version, Measure) + 7 custom
- **Local SQLite Database** — No server required, all data stored locally, no SQLite installation needed

## Tech Stack

| Component | Technology |
|-----------|-----------|
| Add-in Framework | Excel-DNA 1.9 |
| Language | C# / .NET 8.0 |
| Database | SQLite (Microsoft.Data.Sqlite) — bundled, no install needed |
| Excel File I/O | EPPlus |
| CSV Parsing | CsvHelper |
| PDF Generation | PdfSharpCore |
| UI | WinForms dialogs + RibbonX |

## Prerequisites

- **Windows 10** or later (64-bit)
- **Microsoft Excel 2016** or later (64-bit)
- **.NET 8.0 SDK (x64)** — required to build the project
  Download: <https://dotnet.microsoft.com/en-us/download/dotnet/8.0>
  Choose **SDK x64** under Windows installers

> **Note:** SQLite is bundled with the add-in — no separate database installation is required.

## Quick Start (Clone & Run)

After installing the .NET 8 SDK, a fresh clone needs just **two commands**:

```bash
git clone https://github.com/emilesd/MS-Excel-ERP-Analytics.git
cd MS-Excel-ERP-Analytics
LaunchMyOlap.bat
```

`LaunchMyOlap.bat` is self-bootstrapping:

1. Closes any running Excel and clears resiliency state
2. Builds `MyOlap/MyOlap.csproj` in Release
3. Deploys the build output to `%LOCALAPPDATA%\MyOlap\`
4. Registers the XLL under `HKCU\Software\Microsoft\Office\16.0\Excel\Options\OPEN`
5. Launches Excel

Run `LaunchMyOlap.bat` again any time you change source code — it rebuilds and redeploys before launching.

## Manual Build (if you don't want to use the launcher)

```bash
dotnet build MyOlap/MyOlap.csproj -c Release
xcopy /s /y /i MyOlap\bin\Release\net8.0-windows\* %LOCALAPPDATA%\MyOlap\
```

Then either run `LaunchMyOlap.bat` or register the XLL yourself via Excel's *File > Options > Add-Ins > Manage: Excel Add-ins > Go > Browse...* and pick `%LOCALAPPDATA%\MyOlap\MyOlap-AddIn64.xll`.

## Quick Reference

| Action | How |
|--------|-----|
| Create a model | MyOlap tab > **Select Model** > **New Model...** |
| Add dimensions & members | MyOlap tab > **Manage Model** |
| Bulk-load a dimension from a file | MyOlap tab > **Manage Model** > **Load Dimension** |
| Import fact data | MyOlap tab > **Load Data** (supports CSV, XLSX, XLS, TXT) |
| Refresh the grid | MyOlap tab > **Refresh Data** |
| Drill into a member | Click a header cell > **Drill Down** |
| Roll up | Click a header cell > **Drill Up** |
| Swap rows/columns | MyOlap tab > **Swap Row/Col** |
| Filter to one member | Click a header cell > **Keep Selected** |
| Remove a member | Click a header cell > **Remove Selected** |
| Undo any change | MyOlap tab > **Undo Last** |
| Slice by dimension | MyOlap tab > **Pick Member** |
| Export report | MyOlap tab > **Export PDF** |

## Project Structure

```text
MS-Excel-ERP-Analytics/
├── README.md                     # This file
├── LaunchMyOlap.bat              # One-click build + deploy + launch
├── TestGuide.txt                 # Step-by-step testing guide (13 scenarios)
├── .gitignore
│
├── MyOlap/                       # The Excel add-in (main project)
│   ├── MyOlap.csproj             # net8.0-windows, ExcelDna.AddIn 1.9
│   ├── MyOlap.sln                # Solution file
│   ├── AddIn.cs                  # Excel-DNA entry point + automated self-tests
│   ├── UserGuide.md
│   ├── install.bat
│   ├── Core/                     # OLAP engine
│   │   ├── OlapEngine.cs         # Central logic: drill, pivot, filter, grid build
│   │   ├── ViewState.cs          # Which dimensions/members are on each axis
│   │   ├── ModelManager.cs       # Create/edit models, dimensions, members
│   │   ├── DimensionTree.cs      # Hierarchical parent-child dimension tree
│   │   └── UndoManager.cs        # View state undo stack
│   ├── Data/                     # Data layer
│   │   ├── Schema.cs             # Domain models (Model, Dimension, Member, Fact)
│   │   ├── SqliteRepository.cs   # SQLite data access (CRUD, queries, aggregation)
│   │   └── DataLoader.cs         # File import: CSV, XLSX, XLS, TXT
│   ├── Ribbon/
│   │   └── MyOlapRibbon.cs       # Custom ribbon + grid rendering (COM reflection)
│   ├── UI/                       # Dialog forms
│   │   ├── ModelBrowserForm.cs
│   │   ├── ManageStructureForm.cs
│   │   ├── DataLoadForm.cs
│   │   ├── LoadDimensionForm.cs
│   │   ├── MemberPickerForm.cs
│   │   ├── DrillOptionsForm.cs
│   │   └── SettingsForm.cs
│   └── Reports/
│       ├── ReportBuilder.cs
│       └── PdfExporter.cs
│
├── DbQuery/                      # Helper console: ad-hoc queries against myolap.db
│   ├── DbQuery.csproj
│   └── Program.cs
├── ReadXlsx/                     # Helper console: inspect XLSX structure
│   ├── ReadXlsx.csproj
│   └── Program.cs
├── TestConsole/                  # Helper console: scripted end-to-end tests
│   ├── TestConsole.csproj
│   └── Program.cs
│
├── TestData/                     # Test fixtures
│   ├── SampleData.csv
│   ├── Sales_Data Sample.xlsx
│   ├── Sales_Dimensions.xlsx
│   ├── MyAnalysisBook.xlsx
│   ├── BU_Structure.pdf          # Dimension structure references
│   └── Measure_Structure.pdf
│
└── Docs/                         # Product documentation
    ├── MyOlap Product Brief v1.3.pptx
    ├── MyOlap Version 1.5 - Features and Fixes List.xlsx
    ├── MyOlap - Updates and Fixes.docx
    ├── Product_Structure.pdf
    ├── Year_Structure.pdf
    └── feedback.pdf
```

## Helper Projects

In addition to the main add-in (`MyOlap/`), three small .NET 8 console apps live at the repo root for development and diagnostics:

- **`DbQuery/`** — opens the current `myolap.db` SQLite file and runs ad-hoc queries from the command line. Useful for inspecting what the add-in actually persisted.
- **`ReadXlsx/`** — dumps the structure (sheets, ranges, sample rows) of an XLSX file via EPPlus. Useful when debugging the **Load Data** wizard's column mapping.
- **`TestConsole/`** — scripted end-to-end test harness that drives the `OlapEngine` / `SqliteRepository` directly without Excel in the loop. Fast smoke tests.

Build/run any of them with:

```bash
dotnet run --project DbQuery
dotnet run --project ReadXlsx
dotnet run --project TestConsole
```

## Testing

See **[TestGuide.txt](TestGuide.txt)** for a comprehensive step-by-step testing guide with **13 test scenarios** covering model creation, dimension/member management, data loading, drill, pivot, filter, undo, slicing, PDF export, settings, and reopening models.

For unattended/scripted checks, build and run `TestConsole`.

## Local Development Layout

When you run `LaunchMyOlap.bat`, the runtime files end up at:

- `%LOCALAPPDATA%\MyOlap\` — the deployed add-in (XLL + DLLs + DNA + runtime DB)
- `%LOCALAPPDATA%\MyOlap\myolap.db` — the user's local data store (SQLite)

The repository also git-ignores two folders that are sometimes useful during development:

- `AddIn/` — an alternate staging folder if you prefer Excel to load from inside the repo tree
- `Deployed/` — a snapshot of release artifacts to hand to end-users

Neither is committed; both are recreated from build output when you need them.

## Troubleshooting

| Problem | Solution |
|---------|----------|
| `LaunchMyOlap.bat` says ".NET SDK not found" | Install .NET 8 SDK x64, then re-run the script |
| Build fails with NuGet errors | Run `dotnet restore MyOlap/MyOlap.csproj` and try again |
| MyOlap tab doesn't appear in Excel | Close Excel, run `LaunchMyOlap.bat` (it clears resiliency state) |
| .NET runtime error on startup | Install .NET 8.0 Desktop Runtime (x64) in addition to the SDK |
| `dotnet` command not found | Install .NET 8.0 SDK and open a fresh Command Prompt |
| Excel disabled the add-in | Re-run `LaunchMyOlap.bat` (clears resiliency registry) |
| Grid shows "Model is ready" | Model has no data yet — use **Load Data** first |
| Want to inspect what's in the local DB | `dotnet run --project DbQuery` |

## License

Prototype — for evaluation purposes.
