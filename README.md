# INVENTORY PROGRAM – FORBO WAREHOUSE MANAGEMENT

## Overview

Desktop application for warehouse inventory at **Forbo**.
Supports **three warehouse modes**: **Lager 1 (SK)**, **Zert**, and **KMAT**.

- Version: **2.1**
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

### Step 3 – Prepare the master tables
Place your master table Excel files anywhere accessible.
On first launch the program will ask you to locate the relevant file.

**Required columns – Lager 1 (SK) master table:**

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

**Required columns – Zert master table:**

| Column | Description |
|--------|-------------|
| `Charge` | Batch number (text) |
| `Materialnummer` | Material number |
| `Materialkurztext` | Material description |
| `MArt` | Material type |
| `Werk` | Plant |
| `LOrt` | Storage location |
| `BME` | Unit of measure |
| `Frei verwendbar` | Free usable quantity |
| `Länge` | Length (mm) |
| `Breite` | Width (mm) |
| `ADV` | ADV description |

**Required columns – KMAT master table:**

| Column | Description |
|--------|-------------|
| `Kauf` | Customer order number (Kaufnummer) |
| `POS` | Position number within the order (e.g. 10, 20, 30) |
| `Materialnummer` | Material number |
| `Materialkurztext` | Material description |
| `Werk` | Plant |
| `Lagerort` | Storage location |
| `BME` | Unit of measure |
| `Frei verwendbar` | Free usable quantity (optional) |

> One Kaufnummer can have multiple positions. The combination of Kaufnummer + POS uniquely identifies a product.

### Step 4 – Start the program
1. Double-click `start_inventur.bat`
2. The warehouse selection dialog opens — select **Lager 1 (SK)**, **Zert**, or **KMAT**
3. The application opens maximised automatically

---

## First-Time Setup

On the first launch, a dialog will appear asking you to locate the master table file.
Use the **Settings** button at any time to change all paths.

Settings are saved to `config/settings_sk.json`.

---

## Warehouse Selection

Every time the application starts, a selection dialog appears:

| Button | Description |
|--------|-------------|
| **Lager 1 (SK)** | Forbo SK – Malacky warehouse (Rolls, Fach + Width input) |
| **Zert** | Zert warehouse (Charge-based, quantity input only) |
| **KMAT** | KMAT warehouse (Customer order + Position, quantity input) |

Closing the dialog without a selection exits the application.

---

## Usage Guide

### Lager 1 (SK) – Basic workflow

1. **Scan a QR code or barcode**
   - The scan field is always focused
   - Scan with the scanner or type the batch number manually, then press **ENTER**

2. **Roll found**
   - Roll data is displayed automatically (dimensions, area, free usable)
   - Enter the **Shelf Location** and press **ENTER** → cursor jumps to Measured Width
   - Enter the **Measured Width (mm)** and press **ENTER** → entry is saved automatically
   - Optionally add a **Remark** before saving
   - Alternatively click **Save**

3. **Roll not found**
   - A manual entry dialog opens
   - Pre-filled from QR code: Batch No., Location (QR), all dimension stages
   - Fill in manually: **Material No.** (mandatory), Description, **Shelf Location** (mandatory), **Measured Width (mm)** (mandatory), Remarks

4. **Duplicate protection**
   - Already scanned batch numbers are detected automatically and blocked

### Zert – Basic workflow

1. **Scan a QR code or barcode**
   - Same semicolon-delimited format; the batch number is extracted automatically

2. **Charge found**
   - Article data is displayed (Material No., Description, Length, Width, Free Usable, UOM)
   - Enter only the **Recorded Quantity** (mandatory) and press **ENTER** or click **Save**
   - No shelf location or measured width required

3. **Charge not found**
   - A simplified dialog opens with: Charge (pre-filled), **Material No.** (mandatory), **Recorded Quantity** (mandatory), Remarks (optional)

4. **Duplicate protection**
   - Applies within the Zert session separately from SK

### KMAT – Basic workflow

1. **Scan a barcode**
   - The scanner reads only the **Kaufnummer** (customer order number), e.g. `17131209`
   - If the Kaufnummer does not exist in the master table, an error message is shown and the scan is reset — no manual entry is possible

2. **Select Position**
   - A dialog opens automatically showing all available positions for that Kaufnummer
   - Select the correct **POS** from the dropdown and click **OK**

3. **Product found**
   - Product data is displayed (Material No., Description, Plant, Location, UOM, Free Usable)
   - Enter the **Recorded Quantity** (mandatory) and press **ENTER** or click **Save**
   - Optionally add a **Remark** before saving

4. **Duplicate protection**
   - The combination of Kaufnummer + POS is checked — re-scanning the same combination is blocked

### QR Code Format (SK and Zert modes)

The program parses semicolon-delimited QR codes:

```
Locat;Charge;Lnge0;Brte0;Lnge1;Brte1;Lnge2;Brte2
```

If the scanned value contains no semicolons, the entire string is treated as the batch number.

### Keyboard Shortcuts

| Key | Action |
|-----|--------|
| `ENTER` | Advance to next field / save entry |
| `ESC` | Cancel current scan |
| `Ctrl+Z` | Undo last entry |
| `Ctrl+S` | Manual save |
| `F11` | Toggle fullscreen |

---

## Output Files

### Lager 1 (SK): `Inventory_Rolls_SK.xlsx`

| Sheet | Content |
|-------|---------|
| `Inventory` | All successfully matched rolls |
| `Not_Found` | Manually entered rolls not found in master table |

**Column structure:**

| Column | Description |
|--------|-------------|
| Timestamp | Date and time of scan |
| Plant | Plant code |
| Location (Master) | Storage location from master table |
| Location (QR) | Storage location from QR code |
| Material No. | Material number |
| Description | Material description |
| Batch No. | Batch number (text-formatted) |
| Length S0–S2 (mm) | Roll length per stage |
| Width S0–S2 (mm) | Roll width per stage |
| Area (m2) | Calculated: `Länge m × Breite mm / 1000` |
| Free Usable | Free usable area from master table |
| Shelf Location | Shelf location entered during scan |
| Measured Width (mm) | Width measured during scan |
| Remarks | Optional remark |

### Zert: `Inventory_Zert.xlsx`

| Sheet | Content |
|-------|---------|
| `Inventory` | All successfully matched entries |
| `Not_Found` | Manually entered entries not found in master table |

**Column structure:**

| Column | Description |
|--------|-------------|
| Timestamp | Date and time of scan |
| Material No. | Material number |
| Description | Material description |
| Mat. Type | Material type |
| Plant | Plant code |
| Location | Storage location |
| Charge | Batch number (text-formatted) |
| UOM | Unit of measure |
| Free Usable | Free usable quantity from master table |
| Length (mm) | Length from master table |
| Width (mm) | Width from master table |
| ADV | ADV description |
| Recorded Quantity | Quantity entered during scan |
| Remarks | Optional remark |

> **Note:** The Charge column is always formatted as text to preserve leading zeros (e.g. `0618570243`).

### KMAT: `Inventory_KMAT.xlsx`

| Sheet | Content |
|-------|---------|
| `Inventory` | All successfully matched entries |
| `Not_Found` | Not used (no manual entry in KMAT mode) |

**Column structure:**

| Column | Description |
|--------|-------------|
| Timestamp | Date and time of scan |
| Plant | Plant code |
| Location | Storage location |
| Material No. | Material number |
| Description | Material description |
| Kauf-Nr. | Customer order number (Kaufnummer) |
| POS | Position number within the order |
| UOM | Unit of measure |
| Free Usable | Free usable quantity from master table |
| Recorded Quantity | Quantity entered during scan |
| Remarks | Optional remark |

---

## Auto-Save

The program saves automatically after every scan.
No data is lost if the application closes unexpectedly.

---

## Export / Backup

Click **Export / Backup** to create a timestamped copy of the current inventory file:

- SK mode: `backups/Inventory_Rolls_SK_Backup_YYYYMMDD_HHMMSS.xlsx`
- Zert mode: `backups/Inventory_Zert_Backup_YYYYMMDD_HHMMSS.xlsx`
- KMAT mode: `backups/Inventory_KMAT_Backup_YYYYMMDD_HHMMSS.xlsx`

The original file is not modified.

---

## Session Resume

When the program starts and a warehouse is selected, it automatically loads any existing inventory file from the export folder. All previously scanned entries are restored and duplicate protection remains active.

---

## Annual Reset

At the start of a new inventory cycle:

1. Rename or archive the current inventory file (e.g. `Inventory_Rolls_SK_2025.xlsx`)
2. Replace the master table with the new file
3. Start the program — a fresh inventory file will be created automatically

---

## File Structure

```
inventur-programm-f/
├── inventur_app_sk.py        # Main application
├── install_python.bat        # Python installation script
├── start_inventur.bat        # Application launcher
├── requirements.txt          # Python dependencies
├── README.md                 # This documentation
├── icon.ico                  # Application icon (optional)
├── Inventory_Rolls_SK.xlsx   # SK inventory output (auto-created)
├── Inventory_Zert.xlsx       # Zert inventory output (auto-created)
├── Inventory_KMAT.xlsx       # KMAT inventory output (auto-created)
├── backups/                  # Timestamped backup files
├── data/
│   ├── Arbeitstabelle_Rollen_St012.XLSX   # SK master table
│   ├── Arbeitstabelle_ZERT_v2.xlsx        # Zert master table
│   └── Arbeitstabelle_KMAT.xlsx           # KMAT master table
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
  "arbeitstabelle_path": "C:/path/to/sk_master_table.xlsx",
  "export_path": "C:/path/to/output/folder",
  "arbeitstabelle_zert_path": "C:/path/to/zert_master_table.xlsx",
  "export_zert_path": "C:/path/to/zert/output/folder",
  "arbeitstabelle_kmat_path": "C:/path/to/kmat_master_table.xlsx",
  "export_kmat_path": "C:/path/to/kmat/output/folder",
  "vollbild": true
}
```

| Setting | Description |
|---------|-------------|
| `auto_save` | Save after every scan (recommended: `true`) |
| `arbeitstabelle_path` | Full path to the SK master table Excel file |
| `export_path` | Folder for SK inventory file and backups |
| `arbeitstabelle_zert_path` | Full path to the Zert master table Excel file |
| `export_zert_path` | Folder for Zert inventory file and backups |
| `arbeitstabelle_kmat_path` | Full path to the KMAT master table Excel file |
| `export_kmat_path` | Folder for KMAT inventory file and backups |
| `vollbild` | Start maximised (`true` recommended) |

---

## Log File

All activity is logged to `config/inventory_sk.log`:

- Application start / stop
- Warehouse mode selected
- Master table load results
- Every scanned batch number / Kaufnummer
- Errors and warnings

---

## Troubleshooting

### Application does not start
- Check Python installation: run `python --version` in a command prompt
- Re-run `install_python.bat`

### Warehouse selection dialog closes immediately
- The application exits if the dialog is closed without a selection — this is intentional
- Simply start the application again and click a warehouse button

### Master table not found
- Use the **Settings** dialog to set the correct file path
- Make sure the file is not open in Excel when the program loads

### Scanner does not work
- Test the scanner in a text editor — it should type characters and send ENTER
- If ENTER is not sent automatically, press it manually after each scan

### Excel file locked / save error
- Close all Excel windows that may have the file open
- Check that the export folder exists and you have write permissions

### Leading zeros disappear in Excel
- This is handled automatically — the Charge/Batch No. column is always text-formatted

### Duplicate warning appears unexpectedly
- Check the scanned items list — the entry may already be recorded
- Right-click the entry and select **Delete** if needed, then rescan

### KMAT: Kundenauftrag not found
- Verify that the Kaufnummer exists in the KMAT master table
- Make sure the correct master table file is configured in Settings
- No manual entry is possible in KMAT mode — only entries present in the master table can be recorded

---

## Support

If problems persist:
1. Check `config/inventory_sk.log` for error details
2. Restart the application
3. Contact IT support and attach the log file

---

**Developed for Forbo Movement Systems**
*Version 2.1 – March 2026*
