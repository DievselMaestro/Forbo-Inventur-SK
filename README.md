# INVENTORY PROGRAM SK – FORBO MALACKY WAREHOUSE MANAGEMENT

## Overview

Desktop application for warehouse inventory at **Forbo SK – Malacky**.
Supports **Rolls only** (no granulate) with barcode/QR scanner integration.

- Version: **1.0 SK**
- Platform: Windows 11, Python 3.11+
- Main script: `inventur_app_sk.py`

---

## System Requirements

| Component | Requirement |
|-----------|-------------|
| Operating System | Windows 11 (or Windows 10) |
| Python | 3.11 or higher |
| Disk space | At least 500 MB free |
| Hardware | Barcode / QR scanner (Keyboard Wedge) |

Python dependencies: `tkinter`, `pandas`, `openpyxl`

---

## Installation

### Step 1 – Copy files
Copy all program files into a folder on your computer.

### Step 2 – Install Python
1. Double-click `install_python.bat`
2. The script downloads Python and installs all required modules automatically
3. Wait until "Installation complete" is displayed

### Step 3 – Prepare the master table
Place your master table Excel file (single sheet, see column list below) anywhere accessible.
On first launch the program will ask you to locate this file.

**Required columns in the master table:**

| Column | Description |
|--------|-------------|
| `Charge` | Batch number (text, leading zeros preserved) |
| `Material` | Material number |
| `Materialkurztext` | Material description |
| `Werk` | Plant |
| `Lagerort` | Storage location (master) |
| `Länge m` | Stage 0 length in metres |
| `Breite mm` | Stage 0 width in mm |
| `Frei verwendbar` | Free usable area (m²) |
| `Lnge1` / `Brte1` | Stage 1 length (m) / width (mm) |
| `Lnge2` / `Brte2` | Stage 2 length (m) / width (mm) |

> **Note:** Area (m²) is calculated automatically from `Länge m × Breite mm / 1000`. There is no separate area column in the master table.

### Step 4 – Start the program
1. Double-click `start_inventur.bat`
2. The application opens maximised automatically

---

## First-Time Setup

On the first launch, a dialog will appear asking you to locate the master table file.
Use the **Settings** button at any time to change the master table path or the export folder.

Settings are saved to `config/settings_sk.json`.

---

## Usage Guide

### Basic workflow

1. **Scan a QR code or barcode**
   - The scan field is always focused
   - Scan with the scanner or type the batch number manually, then press **ENTER**

2. **Roll found (green)**
   - Roll data is displayed automatically
   - Scan or type the **Shelf Location** and press **ENTER** → cursor jumps to Measured Width
   - Enter the **Measured Width (mm)** and press **ENTER** → entry is saved automatically
   - Optionally add a **Remark** before pressing ENTER on the width field
   - Alternatively click **Save** at any point

3. **Roll not found**
   - A manual entry dialog opens
   - Pre-filled fields (from QR code): Batch No., Location (QR), all dimension stages
   - Fields to fill in manually: **Material No.** (mandatory), Description, **Shelf Location** (mandatory), **Measured Width (mm)** (mandatory), Remarks
   - Click **Save** or press **ENTER**

4. **Duplicate protection**
   - Already scanned batch numbers are detected automatically
   - A warning is shown if the same batch is scanned twice
   - Prevents duplicate entries in the inventory

### QR Code Format

The program parses semicolon-delimited QR codes:

```
LÖrt;Charge;Lnge0;Brte0;Lnge1;Brte1;Lnge2;Brte2
```

Three dimension stages are supported (Stage 0, Stage 1, Stage 2), each with Length and Width in mm.
If the scanned value contains no semicolons, the entire string is treated as the batch number.

### Keyboard Shortcuts

| Key | Action |
|-----|--------|
| `ENTER` on Shelf Location | Jump to Measured Width field |
| `ENTER` on Measured Width | Save the entry |
| `ENTER` on Remarks | Save the entry |
| `ESC` | Cancel current scan |
| `Ctrl+Z` | Undo last entry |
| `Ctrl+S` | Manual save |
| `F11` | Toggle fullscreen |

---

## Output File

The program writes a single Excel file: **`Inventory_Rolls_SK.xlsx`**

The file contains two sheets:

| Sheet | Content |
|-------|---------|
| `Inventory` | All successfully matched rolls |
| `Not_Found` | Manually entered rolls not found in the master table |

### Column structure (both sheets)

| Column | Description |
|--------|-------------|
| Timestamp | Date and time of scan |
| Plant | Plant code from master table |
| Location (Master) | Storage location from master table |
| Location (QR) | Storage location from QR code |
| Material No. | Material number |
| Description | Material description |
| Batch No. | Batch number (formatted as text) |
| Length S0–S2 (mm) | Roll length per stage |
| Width S0–S2 (mm) | Roll width per stage |
| Area (m2) | Calculated: `Länge m × Breite mm / 1000` |
| Free Usable | Free usable area (m²) from master table column `Frei verwendbar` |
| Shelf Location | Shelf location entered during scan |
| Measured Width (mm) | Width measured during scan (control value) |
| Remarks | Optional remark |

> **Note:** The Batch No. column is always formatted as text to preserve leading zeros (e.g. `0618639923`).

---

## Auto-Save

The program saves automatically after every scan.
No data is lost if the application closes unexpectedly.

---

## Export / Backup

Click **Export / Backup** to create a timestamped copy of the current inventory file:

```
backups/Inventory_Rolls_SK_Backup_YYYYMMDD_HHMMSS.xlsx
```

The original file is not modified.

---

## Session Resume

When the program starts, it automatically loads any existing `Inventory_Rolls_SK.xlsx` from the export folder.
All previously scanned rolls are restored and duplicate protection remains active.

---

## Annual Reset

At the start of a new inventory cycle:

1. Rename or archive `Inventory_Rolls_SK.xlsx` (e.g. `Inventory_Rolls_SK_2025.xlsx`)
2. Replace the master table with the new file
3. Start the program — a fresh inventory file will be created automatically

---

## File Structure

```
inventur_programm_sk/
├── inventur_app_sk.py        # Main application (Rolls only, SK version)
├── install_python.bat        # Python installation script
├── start_inventur.bat        # Application launcher
├── requirements.txt          # Python dependencies
├── README.md                 # This documentation
├── icon.ico                  # Application icon (optional)
├── Inventory_Rolls_SK.xlsx   # Inventory output file (auto-created)
├── backups/                  # Timestamped backup files
└── config/
    ├── settings_sk.json      # Application settings
    └── inventory_sk.log      # Log file
```

---

## Configuration

Settings file: `config/settings_sk.json`

```json
{
  "auto_save": true,
  "arbeitstabelle_path": "C:/path/to/master_table.xlsx",
  "export_path": "C:/path/to/output/folder",
  "vollbild": true
}
```

| Setting | Description |
|---------|-------------|
| `auto_save` | Save after every scan (recommended: `true`) |
| `arbeitstabelle_path` | Full path to the master table Excel file |
| `export_path` | Folder where the inventory file and backups are saved |
| `vollbild` | Start maximised (`true` recommended) |

---

## Log File

All activity is logged to `config/inventory_sk.log`:

- Application start / stop
- Master table load results
- Every scanned batch number
- Errors and warnings

---

## Troubleshooting

### Application does not start
- Check Python installation: run `python --version` in a command prompt
- Re-run `install_python.bat`

### Master table not found
- Use the **Settings** dialog to set the correct file path
- Make sure the file is not open in Excel when the program loads

### Scanner does not work
- Test the scanner in a text editor — it should type characters and send ENTER
- If ENTER is not sent automatically, press it manually after each scan
- Or use the **Scan** button in the UI

### Excel file locked / save error
- Close all Excel windows that may have the file open
- Check that the export folder exists and you have write permissions

### Leading zeros disappear in Excel
- This is handled automatically — the Batch No. column is always text-formatted
- If you open the file in Excel and it converts numbers, select the column and format as Text

### Duplicate warning appears unexpectedly
- Check the scanned rolls list — the batch may already be recorded
- Right-click the entry and select **Delete** if needed, then rescan

### Performance is slow
- Close other applications
- Very large master tables (> 5 000 rows) may slow down the initial load

---

## Support

If problems persist:
1. Check `config/inventory_sk.log` for error details
2. Restart the application
3. Contact IT support and attach the log file

---

**Developed for Forbo Movement Systems — Malacky Warehouse**
*Version 1.0 SK — March 2026*
