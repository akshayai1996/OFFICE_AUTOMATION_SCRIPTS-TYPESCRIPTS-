# 📊 Office Scripts Automation Suite for Excel

> **Production-grade TypeScript automations for Microsoft Excel (Excel for Web & Microsoft 365 Desktop). Standardize print settings, build intelligent workbook navigation indexes, filter formatting attributes, and eliminate repetitive spreadsheet chores.**

---

## ⚡ Highlights

* 📑 **Interactive Sheet Index**: Auto-generates an interactive navigation hub with `=HYPERLINK()` jumps, instant **Show / Hide** tab visibility toggles, and zero data loss on rebuilds.
* 🖨️ **Master Page Layout Cloner**: Copies comprehensive print setup (orientation, margins, paper size, multi-page headers & footers, and zoom scaling) from a master sheet to dozens of worksheets in one click.
* 🔍 **Strikethrough Filter Engine**: Solves native Excel's missing ability to filter or sort by strikethrough formatting by auto-generating filterable audit columns.
* 📄 **Global Single-Page Fit**: Instant batch configuration to fit every worksheet in a workbook onto a single page (1 page wide by 1 page tall) for clean PDF exports and printing.
* 🚀 **Zero Dependencies**: Pure TypeScript using native `ExcelScript` APIs—no third-party add-ins, COM libraries, or admin-installed extensions required.

---

## ⚠️ Prerequisites & Environment

> ℹ️ **Office Scripts run directly within Excel on the Web and Microsoft 365 desktop apps connected to OneDrive or SharePoint.**

* **Supported Excel Versions**: Excel for Microsoft 365 (Windows, Mac, Web).
* **Automate Tab**: Ensure the **Automate** tab is visible in your Excel ribbon. (Available for commercial/educational Microsoft 365 licenses with OneDrive/SharePoint storage).
* **Runtime**: Excel JavaScript/TypeScript V8 Cloud & Desktop Sandbox.

---

## 🖥️ System Architecture & Workflow

```text
┌────────────────────────────────────────────────────────────────────────┐
│               EXCEL OFFICE SCRIPTS RUNTIME ARCHITECTURE                │
├────────────────────────────────────────────────────────────────────────┤
│  Microsoft 365 Cloud / Excel for the Web / Desktop (V8 Engine)         │
│  ┌────────────────────────────────────────────────────────────┐        │
│  │             Office Scripts TypeScript Sandbox              │        │
│  │  - Synchronous Execution API (ExcelScript.Workbook)        │        │
│  │  - Batch Formatting & Formula Insertion                    │        │
│  │  - Zero External Dependencies & Zero Local Add-ins         │        │
│  └─────────────────────────────┬──────────────────────────────┘        │
│                                │                                       │
│          ┌─────────────────────┼─────────────────────┐                 │
│          ▼                     ▼                     ▼                 │
│  ┌───────────────┐     ┌───────────────┐     ┌───────────────┐         │
│  │ INDEX_SHEET   │     │ MASTER_SHEET  │     │ STRIKETHROUGH │         │
│  │ Dynamic Index │     │ Page Setup    │     │ Format Audit  │         │
│  │ & Visibility  │     │ Cloner        │     │ & AutoFilter  │         │
│  └───────┬───────┘     └───────┬───────┘     └───────┬───────┘         │
│          │                     │                     │                 │
│          └─────────────────────┼─────────────────────┘                 │
│                                ▼                                       │
│  ┌────────────────────────────────────────────────────────────┐        │
│  │           Workbook Target Layers (Worksheets & Ranges)     │        │
│  └────────────────────────────────────────────────────────────┘        │
└────────────────────────────────────────────────────────────────────────┘
```

Office Scripts operate directly against the Excel document object model via asynchronous/synchronous batches. The scripts in this repository are engineered to minimize round-trip API calls, using bulk methods like `Range.copyFrom()` and array-based formula writes for maximum execution speed.

---

## 🧠 Script Deep Dive & Capabilities

### 1. 📑 `INDEX_SHEET.TS` — Dynamic Sheet Navigation & Visibility Manager

Managing workbooks with 20 to 100+ worksheets can become overwhelming. `INDEX_SHEET.TS` generates and maintains a centralized table of contents on a dedicated `Sheet Index` tab.

```text
┌──────────────────────────────────────────────────────────────────────────┐
│          DYNAMIC INDEX SHEET SCHEMA & PRESERVATION ARCHITECTURE          │
├──────────────────────────────────────────────────────────────────────────┤
│  Col A       Col B          Col C      Col D       Col E+ (Preserved)    │
│ ┌─────────┬──────────────┬──────────┬───────────┬───────────────────┐    │
│ │  Sr No  │  Sheet Name  │   Link   │  Visible  │ Custom Notes, etc │    │
│ ├─────────┼──────────────┼──────────┼───────────┼───────────────────┤    │
│ │    1    │ Summary      │ Open --> │ Show/Hide │ Dept Owner / Tags │    │
│ │    2    │ Q1_Revenue   │ Open --> │ Show/Hide │ Finalized 2026    │    │
│ │    3    │ Raw_Logs     │ Open --> │ Hide      │ Archived Audit    │    │
│ └─────────┴──────────────┴──────────┴───────────┴───────────────────┘    │
│   Footer / Notes Rows Below Table (Preserved verbatim on rebuild)        │
│  ─────────────────────────────────────────────────────────────────       │
│  Rebuild Pipeline:                                                       │
│  1. Snapshot: Read metadata, custom columns & row heights in bulk        │
│  2. Temp Rename: Relabel old index to collision-free _IDX_[ts]           │
│  3. Recreate: Build clean table with live sheets + formula links         │
│  4. Restore: Range.copyFrom() formats & merges preserved user data       │
│  5. Sync: Apply SheetVisibility (visible/hidden) across workbook         │
└──────────────────────────────────────────────────────────────────────────┘
```

#### Key Capabilities:
* **Interactive Visibility Control**: Generates an in-cell dropdown list (`Show` / `Hide`) in Column D. Changing a status and re-running the script immediately shows or hides corresponding worksheets.
* **Non-Destructive Rebuilds**:
  * Preserves any custom columns added to the right of column D (e.g., status, department, owners, notes).
  * Preserves footer notes and calculations placed underneath the table.
  * Preserves custom column widths and row heights.
* **Blazing Fast Formatting Restoration**: Uses single-call `Range.copyFrom(..., ExcelScript.RangeCopyType.formats)` per row instead of reading/writing cell-by-cell attributes. This slashes API overhead from 3,000+ calls down to ~100 on large workbooks.
* **Collision-Proof Execution**: Automatically cleans up orphaned temporary rebuild sheets and prevents name collision bugs via randomized timestamp tokens (`_IDX_[timestamp]`).

---

### 2. 🖨️ `MASTER_SHEET_COPIER.TS` — Universal Print Layout Replicator

Excel does not provide a native one-click way to apply print setup and custom headers/footers across disparate worksheets. `MASTER_SHEET_COPIER.TS` reads the layout rules of your first worksheet (Sheet `0`) and clones them to every other worksheet in the workbook.

```text
┌────────────────────────────────────────────────────────────────────────┐
│                 MASTER PRINT SETUP PROPAGATION ENGINE                  │
├────────────────────────────────────────────────────────────────────────┤
│  [Master Worksheet: Sheet[0]]                                          │
│  ├── Orientation & Paper Size (Portrait/Landscape, A4/Letter)          │
│  ├── Centering Flags (Horizontal & Vertical Center)                    │
│  ├── Margins (Top, Bottom, Left, Right, Header, Footer)                │
│  ├── Print Gridlines & Headings (Row/Column Headers)                   │
│  ├── Scaling / Zoom (Fit to N pages Wide/Tall OR Fixed Scale %)        │
│  └── Header/Footer Groups:                                             │
│      ├── Default for All Pages (Left / Center / Right)                 │
│      ├── First Page Unique (Left / Center / Right)                     │
│      └── Even Pages Unique (Left / Center / Right)                     │
│                                │                                       │
│                                ▼  Deep Copy Loop                       │
│  [Target Sheets: Sheet[1] through Sheet[N]]                            │
│  └── Exact clone applied to every sheet in a single execution pass     │
└────────────────────────────────────────────────────────────────────────┘
```

#### Replicated Print Properties:
* **Orientation & Page Setup**: `Orientation` (Portrait/Landscape), `PaperSize` (Letter, A4, Legal, etc.), and `CenterHorizontally` / `CenterVertically`.
* **Margins**: Millimeter-accurate sync for `TopMargin`, `BottomMargin`, `LeftMargin`, `RightMargin`, `HeaderMargin`, and `FooterMargin`.
* **View & Print Headings**: Synchronizes `PrintGridlines` and `PrintHeadings` across all tabs.
* **Scaling Rules**: Replicates either percentage zoom (`scale`) or multi-page constraints (`horizontalFitToPages` / `verticalFitToPages`).
* **Headers & Footers Hierarchy**:
  * Default Header / Footer (Left, Center, Right sections)
  * First Page Header / Footer (Left, Center, Right sections)
  * Even Pages Header / Footer (Left, Center, Right sections)
  * Preserves active `HeaderFooterState` (`Default`, `FirstPage`, `OddEven`, or `All`).

---

### 3. 🔍 `strikethrough_filter.ts` — Strikethrough Audit & Filter Engine

In Excel, users frequently use strikethrough formatting to mark tasks as completed, reconcile accounting lines, or flag deprecated rows. However, **native Excel filters cannot filter or sort by font strikethrough**.

```text
┌────────────────────────────────────────────────────────────────────────┐
│                STRIKETHROUGH FILTER DETECTION PIPELINE                 │
├────────────────────────────────────────────────────────────────────────┤
│  1. User selects any cell/range in the target column                   │
│  2. Script identifies active column & auto-detects header row          │
│  3. Iterates used rows -> inspects cell.getFormat().getFont()          │
│  4. Checks font.getStrikethrough() boolean state                       │
│  5. Creates helper column 'Strikethrough?' with 'Yes' / 'No'           │
│  6. Refreshes AutoFilter across full range including helper col        │
│                                                                        │
│  Result: You can instantly filter by 'Yes' or 'No' in native Excel     │
└────────────────────────────────────────────────────────────────────────┘
```

#### Problem vs. Solution:

| Feature / Capability | Native Excel UI | `strikethrough_filter.ts` |
| :--- | :--- | :--- |
| **Filter by Strikethrough Font** | ❌ Not Supported | ✅ **Instant One-Click Filter** |
| **Header Row Detection** | Manual selection required | ✅ **Auto-detects header offsets** |
| **Data Integrity** | Manual checking prone to errors | ✅ **Adds explicit `Yes`/`No` audit column** |
| **AutoFilter Refresh** | Manual range re-selection | ✅ **Auto-expands filter to include helper** |

#### How to use:
1. Select any cell or column range containing the text with strikethroughs.
2. Run the script.
3. A new column named `Strikethrough?` will appear at the right of your table populated with `Yes` or `No`.
4. Use standard Excel AutoFilter dropdowns to isolate or review your strikethrough entries.

---

### 4. 📄 `FITSHEETSONONEPAGE.TS` — Global Single-Page Scaler

When exporting large multi-tab financial models or executive summaries to PDF, manual page scaling across 30+ sheets is tedious and error-prone.

#### Core Functionality:
* Iterates through every worksheet via `workbook.getWorksheets()`.
* Configures `pageLayout.setZoom({ horizontalFitToPages: 1, verticalFitToPages: 1 })`.
* Guarantees all tables print or export to PDF neatly fitted to a single page width and height.

---

## 🛠️ Step-by-Step Installation & Execution Guide

Running these scripts requires no external tools or installations:

```text
Step 1: Open your workbook in Excel (Web or Desktop M365).
Step 2: Click on the [Automate] tab in the top navigation ribbon.
Step 3: Click [New Script] (opens the Code Editor sidebar).
Step 4: Delete any default placeholder code.
Step 5: Copy the code from the desired .TS file in this repository and paste it into the editor.
Step 6: Click [Save Script] and give it an intuitive name (e.g., "Build Sheet Index").
Step 7: Click [Run] to execute the automation.
```

> 💡 **Tip**: In the **Automate** pane, you can click the ellipsis (`...`) next to any script and select **Add in workbook**. This creates a clickable macro button directly on your worksheet for one-touch team execution!

---

## 📋 Script Matrix & Quick Reference

```text
┌──────────────────────────┬─────────────────────────┬──────────────────────────┐
│ Script File              │ Core Purpose            │ Execution Scope          │
├──────────────────────────┼─────────────────────────┼──────────────────────────┤
│ INDEX_SHEET.TS           │ Dynamic navigation tab  │ Entire Workbook          │
│ MASTER_SHEET_COPIER.TS   │ Clone print settings    │ Sheet 1 -> Sheets 2..N   │
│ strikethrough_filter.ts  │ Filter strikethrough    │ Active Selection/Sheet   │
│ FITSHEETSONONEPAGE.TS    │ 1x1 page print fit      │ All Worksheets           │
└──────────────────────────┴─────────────────────────┴──────────────────────────┘
```

---

## 📁 Repository Structure

```text
┌────────────────────────────────────────────────────────────────────────┐
│                          REPOSITORY FILE TREE                          │
├────────────────────────────────────────────────────────────────────────┤
│  OFFICE_AUTOMATION_SCRIPTS-TYPESCRIPTS-/                               │
│  ├── INDEX_SHEET.TS             # Interactive navigation builder       │
│  ├── MASTER_SHEET_COPIER.TS     # Print layout propagation engine      │
│  ├── strikethrough_filter.ts    # Font strikethrough filter audit      │
│  ├── FITSHEETSONONEPAGE.TS      # Batch 1-page scaling config          │
│  ├── LICENSE                    # MIT License                          │
│  └── README.md                  # Comprehensive Documentation          │
└────────────────────────────────────────────────────────────────────────┘
```

---

## ❓ Troubleshooting & FAQs

### Q: Why is the Automate tab missing in my Excel?
> **Answer**: The Automate tab requires an active Microsoft 365 commercial, education, or enterprise license. It also requires the workbook to be saved in OneDrive for Business or SharePoint. Personal Microsoft accounts or on-premises older Excel versions do not support Office Scripts.

### Q: Does running `INDEX_SHEET.TS` delete my sheet names or custom data?
> **Answer**: No. `INDEX_SHEET.TS` is non-destructive. It reads the current sheets, preserves any extra columns you added to the right of column D, preserves notes below the table, and restores your custom cell formatting using batch format cloning.

### Q: Can I run these scripts automatically via Power Automate?
> **Answer**: Yes! All scripts in this repository follow the standard `function main(workbook: ExcelScript.Workbook)` signature, making them 100% compatible with the **Run script** action in Microsoft Power Automate cloud flows.

---

## 📄 License

This repository is licensed under the [MIT License](LICENSE). You are free to use, modify, and distribute these scripts across personal and commercial projects.

---

**Automate Excel with Confidence. Zero Add-ins Required.**
