# MyOlap v1.0 — Excel OLAP Add-in

An MS Excel add-in that provides OLAP-style analysis, hierarchical dimension navigation, and report generation — all locally on the user's workstation with no external server dependencies.

## Quick Start

### Building from Source

```
cd MyOlap
dotnet build -c Release
```

The build produces packed `.xll` files in `bin\Release\net8.0-windows\publish\`:
- `MyOlap-AddIn-packed.xll` — 32-bit Excel
- `MyOlap-AddIn64-packed.xll` — 64-bit Excel

### Installing

1. Build the project (or use the pre-built `.xll`)
2. Run `install.bat` from the `bin\Release\net8.0-windows\` folder, OR:
3. Open Excel → File → Options → Add-ins → Excel Add-ins → Browse → select the `.xll` file

The **MyOlap** tab will appear on the Excel ribbon.

## Architecture

| Layer | Technology |
|-------|-----------|
| Excel Integration | Excel-DNA 1.9 (.NET 8) |
| Database | SQLite (local, file-based) |
| Excel I/O | EPPlus 8.5 |
| CSV Parsing | CsvHelper 33 |
| PDF Export | PdfSharpCore 1.3 |

All data is stored in `%LOCALAPPDATA%\MyOlap\myolap.db`.

## Project Structure

```
MyOlap/
├── AddIn.cs                  # Excel-DNA entry point
├── Ribbon/
│   └── MyOlapRibbon.cs       # Ribbon UI + all button callbacks
├── Core/
│   ├── OlapEngine.cs         # Central OLAP query engine
│   ├── ModelManager.cs       # Model creation and structure management
│   ├── DimensionTree.cs      # Hierarchical tree builder
│   ├── ViewState.cs          # View state snapshots (for undo)
│   └── UndoManager.cs        # 3-level undo stack
├── Data/
│   ├── Schema.cs             # Data models (OlapModel, Dimension, Member, etc.)
│   ├── SqliteRepository.cs   # SQLite data access layer
│   └── DataLoader.cs         # File import (xlsx/csv/txt)
├── UI/
│   ├── ModelBrowserForm.cs   # Model selection dialog
│   ├── MemberPickerForm.cs   # Hierarchy tree picker
│   ├── ManageStructureForm.cs# Dimension/member editor
│   ├── SettingsForm.cs       # Retrieval options
│   ├── DrillOptionsForm.cs   # Drill-down mode selector
│   └── DataLoadForm.cs       # File import with column mapping
└── Reports/
    ├── ReportBuilder.cs      # Grid-to-report converter
    └── PdfExporter.cs        # PDF rendering
```
