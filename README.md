# MyOlap v2.2 - Excel OLAP Analytics Add-in

An MS Excel add-in for OLAP-style analytics. Pull data from ERP/Excel/CSV/TXT files, clean it, visualize it, and generate reports — all within Excel.

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
- **Local SQLite Database** — No server required, all data stored locally

## Tech Stack

| Component | Technology |
|-----------|-----------|
| Add-in Framework | Excel-DNA 1.9 |
| Language | C# / .NET 8.0 |
| Database | SQLite (Microsoft.Data.Sqlite) |
| Excel File I/O | EPPlus |
| CSV Parsing | CsvHelper |
| PDF Generation | PdfSharpCore |
| UI | WinForms dialogs + RibbonX |

## Prerequisites

- Windows 10+ (64-bit)
- Microsoft Excel 2016+ (64-bit)
- [.NET 8.0 Desktop Runtime (x64)](https://dotnet.microsoft.com/en-us/download/dotnet/8.0)

## Quick Start

```bash
# Build
dotnet build MyOlap/MyOlap.csproj -c Release

# Install (copy to AppData)
xcopy /s /y MyOlap\bin\Release\net8.0-windows\* %LOCALAPPDATA%\MyOlap\

# Launch
LaunchMyOlap.bat
```

## Project Structure

```
├── LaunchMyOlap.bat              # One-click launcher
├── TestGuide.txt                 # Step-by-step testing guide (13 scenarios)
├── TestData/
│   └── SampleData.csv            # Sample OLAP data for testing
├── MyOlap/
│   ├── MyOlap.csproj             # Project file
│   ├── AddIn.cs                  # Excel-DNA entry point + self-test functions
│   ├── Core/
│   │   ├── OlapEngine.cs         # Central OLAP logic (drill, pivot, filter)
│   │   ├── ViewState.cs          # View state model (axes, visible members)
│   │   ├── ModelManager.cs       # Model/dimension/member CRUD
│   │   ├── DimensionTree.cs      # Hierarchical dimension tree
│   │   └── UndoManager.cs        # Undo stack
│   ├── Data/
│   │   ├── Schema.cs             # Domain models (Model, Dimension, Member, etc.)
│   │   ├── SqliteRepository.cs   # SQLite data access layer
│   │   └── DataLoader.cs         # CSV/Excel/TXT file import
│   ├── Ribbon/
│   │   └── MyOlapRibbon.cs       # Ribbon UI + grid rendering (COM reflection)
│   ├── UI/
│   │   ├── ModelBrowserForm.cs   # Model selection dialog
│   │   ├── ManageStructureForm.cs# Dimension/member management
│   │   ├── DataLoadForm.cs       # Data import wizard
│   │   ├── MemberPickerForm.cs   # Member selection for slicing
│   │   ├── DrillOptionsForm.cs   # Drill mode selection
│   │   └── SettingsForm.cs       # Model settings
│   └── Reports/
│       ├── ReportBuilder.cs      # Report data model
│       └── PdfExporter.cs        # PDF generation
```

## Testing

See [TestGuide.txt](TestGuide.txt) for a comprehensive step-by-step testing guide with 13 test scenarios covering all features.

## License

Prototype — for evaluation purposes.
