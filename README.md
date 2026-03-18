# MyOlap v2.2 - Excel OLAP Analytics Add-in

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
  Download: https://dotnet.microsoft.com/en-us/download/dotnet/8.0  
  Choose **SDK x64** under Windows installers

> **Note:** SQLite is bundled with the add-in — no separate database installation is required.

## Installation (Step by Step)

### Step 1: Install .NET 8.0 SDK

1. Go to https://dotnet.microsoft.com/en-us/download/dotnet/8.0
2. Under **SDK**, download the **x64 Windows installer**
3. Run the installer and follow the prompts
4. Verify by opening **Command Prompt** and running:
   ```
   dotnet --version
   ```
   You should see a version like `8.0.xxx`

### Step 2: Clone or Download the Repository

```bash
git clone https://github.com/emilesd/MS-Excel-ERP-Analytics.git
cd MS-Excel-ERP-Analytics
```

Or download as ZIP from GitHub and extract.

### Step 3: Build the Project

Open **Command Prompt** in the project folder and run:

```bash
dotnet build MyOlap/MyOlap.csproj -c Release
```

Wait for the `Build succeeded` message.

### Step 4: Deploy the Add-in

Copy the build output to your local AppData folder:

```bash
xcopy /s /y /i MyOlap\bin\Release\net8.0-windows\* %LOCALAPPDATA%\MyOlap\
```

### Step 5: Launch

Double-click **LaunchMyOlap.bat** in the project folder.

Excel will open with the **"MyOlap v2.2"** tab in the ribbon.

> **Subsequent launches:** Just run `LaunchMyOlap.bat` again. You only need to repeat Steps 3-4 if the source code changes.

## Quick Reference

| Action | How |
|--------|-----|
| Create a model | MyOlap tab > **Select Model** > **New Model...** |
| Add dimensions & members | MyOlap tab > **Manage Model** |
| Import data | MyOlap tab > **Load Data** (supports CSV, XLSX, XLS, TXT) |
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

```
MS-Excel-ERP-Analytics/
├── README.md                     # This file
├── LaunchMyOlap.bat              # One-click launcher (clears cache, opens Excel)
├── TestGuide.txt                 # Step-by-step testing guide (13 scenarios)
├── TestData/
│   └── SampleData.csv            # Sample OLAP data (62 rows, multi-dimensional)
│
├── MyOlap/                       # Source code
│   ├── MyOlap.csproj             # Project file & NuGet dependencies
│   ├── AddIn.cs                  # Excel-DNA entry point + automated self-tests
│   │
│   ├── Core/                     # OLAP engine
│   │   ├── OlapEngine.cs         # Central logic: drill, pivot, filter, grid build
│   │   ├── ViewState.cs          # Tracks which dimensions/members are on each axis
│   │   ├── ModelManager.cs       # Create/edit models, dimensions, members
│   │   ├── DimensionTree.cs      # Hierarchical parent-child dimension tree
│   │   └── UndoManager.cs        # View state undo stack
│   │
│   ├── Data/                     # Data layer
│   │   ├── Schema.cs             # Domain models (Model, Dimension, Member, Fact)
│   │   ├── SqliteRepository.cs   # SQLite data access (CRUD, queries, aggregation)
│   │   └── DataLoader.cs         # File import: CSV, XLSX, XLS, TXT
│   │
│   ├── Ribbon/                   # Excel UI
│   │   └── MyOlapRibbon.cs       # Custom ribbon + grid rendering (COM reflection)
│   │
│   ├── UI/                       # Dialog forms
│   │   ├── ModelBrowserForm.cs   # Browse & select models
│   │   ├── ManageStructureForm.cs# Add/edit dimensions & members
│   │   ├── DataLoadForm.cs       # Data import wizard with column mapping
│   │   ├── MemberPickerForm.cs   # Select member for slicing
│   │   ├── DrillOptionsForm.cs   # Choose drill mode (children / next level)
│   │   └── SettingsForm.cs       # Model display settings
│   │
│   └── Reports/                  # Reporting
│       ├── ReportBuilder.cs      # Build report data from grid
│       └── PdfExporter.cs        # Generate formatted PDF output
```

## Testing

See **[TestGuide.txt](TestGuide.txt)** for a comprehensive step-by-step testing guide with **13 test scenarios** covering:

1. Create a New Model
2. Add Custom Dimensions & Members
3. Load Sample Data from CSV
4. View the OLAP Grid
5. Drill Down / Drill Up
6. Swap Row/Col (Pivot)
7. Keep Selected / Remove Selected
8. Undo Last
9. Pick Member (Change Slice)
10. Export to PDF
11. Settings
12. Reopen an Existing Model

## Troubleshooting

| Problem | Solution |
|---------|----------|
| `LaunchMyOlap.bat` says "Add-in not found" | Run the build and deploy steps (Steps 3-4 above) |
| MyOlap tab doesn't appear in Excel | Close Excel, run `LaunchMyOlap.bat` (clears cache) |
| .NET runtime error on startup | Install .NET 8.0 Desktop Runtime (x64) |
| `dotnet` command not found | Install .NET 8.0 SDK and restart Command Prompt |
| Excel disabled the add-in | Run `LaunchMyOlap.bat` (clears resiliency registry) |
| Grid shows "Model is ready" | Model has no data yet — use **Load Data** first |

## License

Prototype — for evaluation purposes.
