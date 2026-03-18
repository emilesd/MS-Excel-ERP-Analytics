# MyOlap v1.0 — User Guide

## 1. Installation

1. Obtain the `MyOlap-AddIn64-packed.xll` file (for 64-bit Excel) or `MyOlap-AddIn-packed.xll` (for 32-bit Excel).
2. Open Microsoft Excel (Microsoft 365 desktop).
3. Go to **File > Options > Add-ins**.
4. At the bottom, select **Excel Add-ins** from the "Manage" dropdown and click **Go...**.
5. Click **Browse...**, navigate to the `.xll` file, and click **OK**.
6. The **MyOlap** tab will appear on the Excel ribbon.

Alternatively, simply double-click the `.xll` file to load it into the current Excel session.

---

## 2. Creating a Model

### From the Ribbon
1. Click **Select Model** on the MyOlap ribbon tab.
2. In the Model Browser dialog, click **New Model...**.
3. Enter a name for your model and click OK.
4. The model is created with 5 pre-defined dimensions:
   - **View** (Actual, Budget, Forecast)
   - **Version**
   - **Time**
   - **Year**
   - **Measure** (Fact)

### From a Workbook
You can also create a model from an Excel workbook with two sheets:
- **Dimensions** sheet: columns `Name | Type | SortOrder`
- **Members** sheet: columns `DimensionName | MemberName | ParentName | Description | Level`

---

## 3. Managing Model Structure

1. Select a model first (Select Model button).
2. Click **Manage Model** on the ribbon.
3. The Manage Structure dialog shows:
   - **Left panel**: List of all dimensions (pre-defined dimensions are tagged with their type).
   - **Right panel**: Hierarchical tree of members for the selected dimension.
4. Use **Add Dimension** to add user-defined dimensions (up to 12 total).
5. Select a dimension and use **Add Member** to create members. If a tree node is selected, the new member becomes its child.
6. Use **Remove** to delete a member and its children.

---

## 4. Loading Data

1. Click **Load Data** on the ribbon.
2. Click **...** to browse for a data file (`.xlsx`, `.csv`, or `.txt`).
3. The column headers from the file appear in the mapping grid.
4. For each column, assign:
   - **Role**: `Dimension`, `Value (Numeric)`, `Value (Text)`, or `(Skip)`.
   - **Map To Dimension**: which model dimension the column maps to.
5. At least one column must be assigned as a Value.
6. Click **Load Data** — records are imported into the model.
7. Members not already in the model are auto-created during import.

---

## 5. Viewing Data (Select Model + Refresh)

1. Click **Select Model** and choose your model.
2. The default view displays:
   - First user-defined dimension on **rows**.
   - Measure dimension on **columns**.
   - All other dimensions in the **point-of-view** (first root member selected).
3. Click **Refresh Data** at any time to re-query and re-render the grid.

---

## 6. Navigating Dimensions

### Pick Member
1. Click **Pick Member** on the ribbon.
2. Select a dimension from the dropdown.
3. Browse the hierarchy tree and select a member.
4. Choose whether to place it on **Rows** or **Columns**.
5. Click OK — the grid updates with the new member on the chosen axis.

### Drill Down
1. Click on a member cell (row or column header) in the worksheet.
2. Click **Drill Down** on the ribbon.
3. Choose the drill mode:
   - **Next Generation**: shows immediate children only.
   - **All Generations**: shows the full subtree.
   - **Base Generation Only**: shows only leaf (bottom-level) members.
4. The member is replaced by its children in the grid.

### Drill Up
1. Click on a member cell in the worksheet.
2. Click **Drill Up** — the member is replaced by its parent.
3. If already at the top, it resets to all root members.

---

## 7. View Manipulation

### Swap Row/Col (Pivot)
Click **Swap Row/Col** to transpose rows and columns, providing a pivoted view of the same data.

### Keep Selected
1. Click on a member cell.
2. Click **Keep Selected** — all other members in that dimension are removed from the view, keeping only the selected one.

### Remove Selected
1. Click on a member cell.
2. Click **Remove Selected** — only the selected member is removed from the view.

### Undo Last
Click **Undo Last** to revert to the previous view. You can undo up to **3 times**.

---

## 8. Settings

Click **Settings** on the ribbon to configure:

| Setting | Description |
|---------|-------------|
| Omit Empty Rows | Hides rows where all values are empty |
| Omit Empty Columns | Hides columns where all values are empty |
| Member Display | Show **Name Only**, **Description Only**, or **Name and Description** |

Settings are saved per model and applied on the next Refresh.

---

## 9. Exporting Reports to PDF

1. Set up your desired view using the navigation tools above.
2. Click **Export PDF** on the ribbon.
3. Choose a file name and location.
4. A formatted PDF report is generated containing:
   - Report title (model name)
   - Generation timestamp
   - Row and column headers
   - Data values formatted to 2 decimal places

---

## 10. Model Structure Limits

| Feature | Limit |
|---------|-------|
| Dimensions per model | Up to 12 (5 pre-defined + 7 user-defined) |
| Pre-defined dimensions | View, Version, Time, Year, Measure |
| Hierarchy levels | Unlimited |
| Filters per member | Up to 5 |
| Open models | 1 at a time (prototype) |
| Undo levels | 3 |

---

## 11. Data Storage

All data is stored locally in a SQLite database at:
```
%LOCALAPPDATA%\MyOlap\myolap.db
```
No external server connection is required. The database file can be backed up by simply copying it.
