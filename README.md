# INVENTORY PROGRAM – FORBO WAREHOUSE MANAGEMENT

## Overview

Desktop application for warehouse inventory at **Forbo**.
Supports **eight warehouse modes**: **HALB**, **ZERT**, **KMAT**, **WIP**, **Assembling**, **Service**, **Warehouse** and **Granules**.

- Version: **3.0** (shown in the window title, the header and the start dialog)
- Platform: Windows 11, Python 3.11+
- Main script: `inventur_app.py`

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

**Required columns – HALB master table:**

| Column | Description |
|--------|-------------|
| `Batch` | Batch number (text, leading zeros preserved) |
| `Material` | Material number |
| `Material Description` | Material description |
| `Plant` | Plant |
| `Storage Location` | Storage location (master) |
| `storage bin` | Shelf / storage bin |
| `length m` | Stage 0 length in metres |
| `width mm` | Stage 0 width in mm |
| `Unrestricted` | Free usable area (m²) |
| `length1` / `width1` | Stage 1 length (m) / width (mm) |
| `length2` / `width2` | Stage 2 length (m) / width (mm) |

**Required columns – ZERT master table:**

| Column | Description |
|--------|-------------|
| `Batch` | Batch number (text) |
| `Material Number` | Material number |
| `Material Description` | Material description |
| `Plant` | Plant |
| `Storage Location` | Storage location |
| `Base Unit of Measure` | Unit of measure |
| `Unrestricted` | Free usable quantity |
| `Length` | Length (mm) |
| `Width` | Width (mm) |
| `ADV` | ADV description |

**Required columns – KMAT master table:**

| Column | Description |
|--------|-------------|
| `Special stock number` | Customer order number |
| `POS` | Position number within the order (e.g. 10, 20, 30) |
| `Material Number` | Material number |
| `Material Description` | Material description |
| `Plant` | Plant |
| `Storage Location` | Storage location |
| `Base Unit of Measure` | Unit of measure |
| `Unrestricted` | Free usable quantity (optional) |

> One Special stock number can have multiple positions. The combination of Special stock number + POS uniquely identifies a product.

**Required columns – WIP master table:**

| Column | Description |
|--------|-------------|
| `Plant` | Plant |
| `Sales Order` | Sales order number (scanned by the user) |
| `Sales Order Item` | Sales order item |
| `Order` | Manufacturing order number (unique, selected from dropdown) |
| `Material Number` | Material number |
| `Material description` | Material description |
| `Quantity Delivered (GMEIN)` | Delivered quantity |
| `Unit of measure (=GMEIN)` | Unit of measure |
| `Basic finish date` | Planned finish date |
| `Order quantity (GMEIN)` | Order quantity |
| `System Status` | Current system status |

> The WIP master table must be provided as a single-sheet Excel file (.xlsx). The program always reads the **first sheet**, regardless of its name.

> One Sales Order can contain multiple Orders. Each Order number is unique across the entire master table.

**Assembling / Service / Warehouse master tables:** same columns as the HALB master table.
Each mode has its own master table and export folder in **Settings**.

**Required columns – Granules master table** (German SAP headers or English headers are both accepted):

| German column | English column | Description |
|---------------|----------------|-------------|
| `Werk` | `Plant` | Plant |
| `LOrt` | `Storage Location` | Storage location |
| `Materialnummer` | `Material Number` | Material number |
| `Materialkurztext` | `Material Description` | Material description |
| `Charge` | `Batch` | Batch number (scanned, **required**) |
| `Frei verwendbar` | `Unrestricted` | Free usable quantity |
| `BME` | `Base Unit of Measure` | Unit of measure (KG) |

> The Granules master table is read from the **first sheet**. Leading zeros of the batch number are ignored when searching (`0618604585` = `618604585`).

### Step 4 – Start the program
1. Double-click `start_inventur.bat`
2. The warehouse selection dialog opens — select one of the eight warehouse modes
3. The application opens maximised automatically

---

## First-Time Setup

On the first launch, a dialog will appear asking you to locate the master table file.
Use the **Settings** button at any time to change all paths.

Settings are saved to `config/settings.json`. The Settings dialog has one tab per warehouse mode.
An existing `config/settings_sk.json` from an older version is taken over automatically on the first start.

---

## Warehouse Selection

Every time the application starts, a selection dialog appears with eight buttons arranged in two rows:

| Button | Description |
|--------|-------------|
| **HALB** | HALB warehouse (Rolls, Shelf Location + Width input) |
| **ZERT** | ZERT warehouse (Charge-based, quantity input only) |
| **KMAT** | KMAT warehouse (Customer order + Position, quantity input) |
| **WIP** | WIP warehouse (Sales Order + Order selection, quantity input) |
| **ASSEMBLING** | Assembling – works like HALB, output `Inventory_Assembling.xlsx` |
| **SERVICE** | Service – works like HALB, output `Inventory_Service.xlsx` |
| **WAREHOUSE** | Warehouse – works like HALB, output `Inventory_Warehouse.xlsx` |
| **GRANULES** | Granules (batch barcode, counted quantity in KG) |

Closing the dialog without a selection exits the application.

---

## Usage Guide

### HALB – Basic workflow

1. **Scan a QR code or barcode**
   - The scan field is always focused and displayed in large text for easy reading
   - Scan with the scanner or type the batch number manually, then press **ENTER**

2. **Roll found**
   - Roll data is displayed first (dimensions, area, free usable)
   - The input fields appear below the roll data
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

### ZERT – Basic workflow

1. **Scan a QR code or barcode**
   - Same semicolon-delimited format; the batch number is extracted automatically

2. **Charge found**
   - Article data is displayed (Material No., Description, Length, Width, Free Usable, UOM)
   - Enter only the **Recorded Quantity** (mandatory) and press **ENTER** or click **Save**
   - No shelf location or measured width required

3. **Charge not found**
   - A simplified dialog opens with: Charge (pre-filled), **Material No.** (mandatory), **Recorded Quantity** (mandatory), Remarks (optional)

4. **Duplicate protection**
   - Applies within the ZERT session separately from HALB

### KMAT – Basic workflow

1. **Scan a barcode**
   - The scanner reads only the **Special stock number** (customer order number), e.g. `17131209`
   - If the Special stock number does not exist in the master table, an error message is shown and the scan is reset — no manual entry is possible

2. **Select Position**
   - A dialog opens automatically showing all available positions for that Special stock number
   - Select the correct **POS** from the dropdown and click **OK**

3. **Product found**
   - Product data is displayed (Material No., Description, Plant, Location, UOM, Free Usable)
   - Enter the **Recorded Quantity** (mandatory) and press **ENTER** or click **Save**
   - Optionally add a **Remark** before saving

4. **Duplicate protection**
   - The combination of Special stock number + POS is checked — re-scanning the same combination is blocked

### WIP – Basic workflow

1. **Scan a barcode**
   - The scanner reads the **Sales Order** number, e.g. `16494221`
   - If the Sales Order does not exist in the master table, or all its Orders have already been recorded, an error message is shown

2. **Select Order**
   - A dialog opens automatically showing all **not yet recorded** Orders for that Sales Order
   - Select the correct **Order** from the dropdown and click **OK**
   - Already recorded Orders are automatically hidden from the list
   - Scanning the same Sales Order again will only show the remaining unrecorded Orders

3. **Product found**
   - Product data is displayed (Material No., Description, Plant, Order Qty, UOM, Basic Finish Date)
   - Enter the **Recorded Quantity** (mandatory) and press **ENTER** or click **Save**
   - Optionally add a **Remark** before saving

4. **Duplicate protection**
   - Each Order number is unique — once recorded it no longer appears in the dropdown

### Assembling / Service / Warehouse – Basic workflow

Identical to the HALB workflow. Each mode uses its own master table and writes its own
output file `Inventory_<Mode>.xlsx`.

### GRANULES – Basic workflow

1. **Scan the batch barcode**
   - The barcode contains the batch number only, e.g. `0618604585`
2. **Batch found** — the following is displayed:
   - Batch No., Material No., Description, Quantity (free usable, KG)
   - Enter the **Counted Qty (KG)** (mandatory, number) and optionally a **Remark**
   - Press **ENTER** or click **Save**
3. **Batch not found**
   - A manual entry dialog opens: **Material No.** (mandatory), Description, **Counted Qty (KG)** (mandatory), Remarks
   - The entry is saved to the `Not_Found` sheet
4. **Duplicate protection** — the same batch cannot be recorded twice

### QR Code Format (HALB-type and ZERT modes)

The program parses semicolon-delimited QR codes:

```
Locat;Charge;Lnge0;Brte0;Lnge1;Brte1;Lnge2;Brte2
```

If the scanned value contains no semicolons, the entire string is treated as the batch number.

### Editing an existing entry

If a value was entered incorrectly (wrong shelf location, wrong width, wrong quantity, wrong remark), it can be corrected without deleting and re-scanning:

1. **Right-click** the entry in the **SCANNED ITEMS** list
2. Select **Edit entry** from the context menu
3. A dialog opens with the current values pre-filled
4. Correct the desired field(s) and click **Save**

Editable fields per mode:

| Mode | Editable fields |
|------|----------------|
| HALB / Assembling / Service / Warehouse | Shelf Location, Measured Width (mm), Remarks |
| ZERT | Recorded Quantity, Remarks |
| KMAT | Recorded Quantity, Remarks |
| WIP  | Recorded Quantity, Remarks |
| GRANULES | Counted Qty (KG), Remarks |

The batch number / charge / order number is shown for reference but cannot be changed.
The Excel file is updated automatically after saving.

### Toolbar Buttons

| Button | Action |
|--------|--------|
| **Export / Backup** | Create a timestamped backup of the current inventory file |
| **Settings** | Open the settings dialog (file paths, display options) |
| **Fullscreen (F11)** | Toggle fullscreen mode |
| **Quit** | Exit the application |
| **Switch Warehouse** | Return to the warehouse selection dialog (hidden by default — enable in Settings) |

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

### HALB: `Inventory_HALB.xlsx`

Assembling, Service and Warehouse write the same structure to `Inventory_Assembling.xlsx`,
`Inventory_Service.xlsx` and `Inventory_Warehouse.xlsx`.

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
| Area (m2) | Calculated: `length m × width mm / 1000` |
| Free Usable | Free usable area from master table |
| Shelf Location | Shelf location entered during scan |
| Measured Width (mm) | Width measured during scan |
| Remarks | Optional remark |

### ZERT: `Inventory_ZERT.xlsx`

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
| Plant | Plant code |
| Location | Storage location |
| Batch | Batch number (text-formatted) |
| Base Unit of Measure | Unit of measure |
| Unrestricted | Free usable quantity from master table |
| Length | Length from master table |
| Width | Width from master table |
| ADV | ADV description |
| Recorded Quantity | Quantity entered during scan |
| Remarks | Optional remark |

> **Note:** The Batch column is always formatted as text to preserve leading zeros (e.g. `0618570243`).

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
| Special stock number | Customer order number |
| POS | Position number within the order |
| UOM | Unit of measure |
| Free Usable | Free usable quantity from master table |
| Recorded Quantity | Quantity entered during scan |
| Remarks | Optional remark |

### WIP: `Inventory_WIP.xlsx`

| Sheet | Content |
|-------|---------|
| `Inventory` | All recorded WIP entries |

**Column structure:**

| Column | Description |
|--------|-------------|
| Timestamp | Date and time of scan |
| Plant | Plant code |
| Sales Order | Sales order number (scanned) |
| Order | Manufacturing order number (selected from dropdown) |
| Material No. | Material number |
| Description | Material description |
| UOM | Unit of measure |
| Order Qty | Order quantity from master table |
| Recorded Quantity | Quantity entered during scan |
| Remarks | Optional remark |

### GRANULES: `Inventory_Granules.xlsx`

| Sheet | Content |
|-------|---------|
| `Inventory` | All batches found in the master table |
| `Not_Found` | Batches entered manually (not in the master table) |

**Column structure:**

| Column | Description |
|--------|-------------|
| Timestamp | Date and time of scan |
| Plant | Plant code |
| Location | Storage location |
| Material No. | Material number |
| Description | Material description |
| Batch No. | Batch number (text, leading zeros kept) |
| Free Usable | Free usable quantity from master table |
| UOM | Unit of measure |
| Counted Qty (KG) | Quantity entered during scan |
| Difference (KG) | Counted Qty − Free Usable (empty for not-found batches) |
| Remarks | Optional remark |

---

## Auto-Save

The program saves automatically after every scan.
No data is lost if the application closes unexpectedly.

---

## Export / Backup

Click **Export / Backup** to create a timestamped copy of the current inventory file:

- HALB mode: `backups/Inventory_HALB_Backup_YYYYMMDD_HHMMSS.xlsx`
- ZERT mode: `backups/Inventory_ZERT_Backup_YYYYMMDD_HHMMSS.xlsx`
- KMAT mode: `backups/Inventory_KMAT_Backup_YYYYMMDD_HHMMSS.xlsx`
- WIP mode:  `backups/Inventory_WIP_Backup_YYYYMMDD_HHMMSS.xlsx`
- Assembling / Service / Warehouse: `backups/Inventory_<Mode>_Backup_YYYYMMDD_HHMMSS.xlsx`
- GRANULES mode: `backups/Inventory_Granules_Backup_YYYYMMDD_HHMMSS.xlsx`

The original file is not modified.

---

## Session Resume

When the program starts and a warehouse is selected, it automatically loads any existing inventory file from the export folder. All previously scanned entries are restored and duplicate protection remains active.

---

## Annual Reset

At the start of a new inventory cycle:

1. Rename or archive the current inventory file (e.g. `Inventory_HALB_2025.xlsx`)
2. Replace the master table with the new file
3. Start the program — a fresh inventory file will be created automatically

---

## File Structure

```
inventur-programm-f/
├── inventur_app.py           # Main application
├── install_python.bat        # Python installation script
├── start_inventur.bat        # Application launcher
├── requirements.txt          # Python dependencies
├── README.md                 # This documentation
├── icon.ico                  # Application icon (optional)
├── Inventory_HALB.xlsx       # HALB inventory output (auto-created)
├── Inventory_ZERT.xlsx       # ZERT inventory output (auto-created)
├── Inventory_KMAT.xlsx       # KMAT inventory output (auto-created)
├── Inventory_WIP.xlsx        # WIP inventory output (auto-created)
├── Inventory_<Mode>.xlsx     # Assembling / Service / Warehouse / Granules output (auto-created)
├── backups/                  # Timestamped backup files
├── Daten/
│   ├── Arbeitstabelle_Rollen_St012_EN.XLSX   # HALB master table
│   ├── Arbeitstabelle_ZERT_EN.xlsx           # ZERT master table
│   ├── Arbeitstabelle_KMAT_EN.xlsx           # KMAT master table
│   ├── Arbeitstabelle_WIP.xlsx               # WIP master table
│   └── Arbeitstabelle_Granulate.xlsx         # Granules master table
└── config/
    ├── settings.json         # Application settings
    └── inventory.log         # Log file
```

---

## Configuration

Settings file: `config/settings.json`

```json
{
  "auto_save": true,
  "arbeitstabelle_path": "C:/path/to/halb_master_table.xlsx",
  "export_path": "C:/path/to/output/folder",
  "arbeitstabelle_zert_path": "C:/path/to/zert_master_table.xlsx",
  "export_zert_path": "C:/path/to/zert/output/folder",
  "arbeitstabelle_kmat_path": "C:/path/to/kmat_master_table.xlsx",
  "export_kmat_path": "C:/path/to/kmat/output/folder",
  "arbeitstabelle_wip_path": "C:/path/to/wip_master_table.xlsx",
  "export_wip_path": "C:/path/to/wip/output/folder",
  "arbeitstabelle_assembling_path": "C:/path/to/assembling_master_table.xlsx",
  "export_assembling_path": "C:/path/to/assembling/output/folder",
  "arbeitstabelle_service_path": "C:/path/to/service_master_table.xlsx",
  "export_service_path": "C:/path/to/service/output/folder",
  "arbeitstabelle_warehouse_path": "C:/path/to/warehouse_master_table.xlsx",
  "export_warehouse_path": "C:/path/to/warehouse/output/folder",
  "arbeitstabelle_granules_path": "C:/path/to/granules_master_table.xlsx",
  "export_granules_path": "C:/path/to/granules/output/folder",
  "vollbild": true,
  "show_switch_warehouse": false
}
```

| Setting | Description |
|---------|-------------|
| `auto_save` | Save after every scan (recommended: `true`) |
| `arbeitstabelle_path` | Full path to the HALB master table Excel file |
| `export_path` | Folder for HALB inventory file and backups |
| `arbeitstabelle_zert_path` | Full path to the ZERT master table Excel file |
| `export_zert_path` | Folder for ZERT inventory file and backups |
| `arbeitstabelle_kmat_path` | Full path to the KMAT master table Excel file |
| `export_kmat_path` | Folder for KMAT inventory file and backups |
| `arbeitstabelle_wip_path` | Full path to the WIP master table Excel file |
| `export_wip_path` | Folder for WIP inventory file and backups |
| `arbeitstabelle_<mode>_path` | Master table for `assembling`, `service`, `warehouse`, `granules` |
| `export_<mode>_path` | Output folder for `assembling`, `service`, `warehouse`, `granules` |
| `vollbild` | Start maximised (`true` recommended) |
| `show_switch_warehouse` | Show the **Switch Warehouse** button in the toolbar (default: `false`) |

---

## Log File

All activity is logged to `config/inventory.log`:

- Application start / stop
- Warehouse mode selected
- Master table load results
- Every scanned batch number / Sales Order / Kaufnummer
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
- Right-click the entry and select **Edit entry** to correct it, or **Delete entry** to remove and rescan

### KMAT: Special stock number not found
- Verify that the Special stock number exists in the KMAT master table
- Make sure the correct master table file is configured in Settings
- No manual entry is possible in KMAT mode — only entries present in the master table can be recorded

### WIP: Sales Order not found / no orders shown
- Verify that the Sales Order number exists in the WIP master table
- If all Orders under that Sales Order have already been recorded, the message "Sales Order not found or all orders have already been recorded" is shown — this is correct behaviour
- Make sure the correct WIP master table file is configured in Settings
- The program reads the **first sheet** of the Excel file regardless of its name

---

## Support

If problems persist:
1. Check `config/inventory.log` for error details
2. Restart the application
3. Contact IT support and attach the log file

---

---

## Changelog

### Version 3.0 – October 2026
- **Feature – New warehouse modes:** **Assembling**, **Service** and **Warehouse** (work like HALB, each with its own master table, export folder and `Inventory_<Mode>.xlsx`).
- **Feature – Granules mode:** scan batch barcode, enter counted quantity in KG. Batches not in the master table are entered manually and saved to the `Not_Found` sheet. Output contains a `Difference (KG)` column. German SAP column headers are accepted.
- **UI – Version number** shown in window title, header and start dialog.
- **UI – Settings dialog** with one tab per warehouse mode.
- **UI – English wording:** "Charge" replaced by "Batch No." in all visible texts (ZERT Excel column now `Batch No.`; older files with `Charge` are still loaded).
- **General:** All site-specific references removed. Files renamed: `inventur_app.py`, `start_inventur.bat`, `config/settings.json` (old settings are migrated automatically), `config/inventory.log`.

### Version 2.6 – June 2026
- **UI – Full English translation:** All remaining German text in dialogs, error messages, and labels has been replaced with English. The startup warehouse selection dialog no longer shows bilingual text.
- **UI – "Kauf-Nr." renamed to "Purchase No.":** All visible labels, column headers, and error messages now use "Purchase No." consistently (internal variable names and Excel column headers updated accordingly).
- **UI – WIP unit label:** "Stk" (Stück) changed to "pcs" in the WIP input panel.
- **Feature – Switch Warehouse button:** A new **Switch Warehouse** button can be shown in the toolbar via **Settings → Show 'Switch Warehouse' button**. It is hidden by default and intended for presentations. Clicking it asks for confirmation and restarts the application, showing the warehouse selection dialog again.

### Version 2.5 – May 2026
- **Bug fix – WIP Edit entry:** "Entry not found in data" error when trying to edit a WIP entry is resolved. The lookup now correctly searches `sales_order` + `order` in the WIP data list instead of falling back to the HALB data.
- **Bug fix – WIP Delete entry:** Same root cause fixed for delete — WIP entries are now removed from the correct data list.
- **UI – WIP scan dialog (Order selection popup):** All font sizes doubled (labels, combobox input, dropdown list). Window enlarged from 400×200 to 600×350 px.
- **UI – WIP info panel:** All label and value font sizes doubled (9→18, 10→20).
- **UI – WIP input panel:** All label and entry field font sizes doubled (10→20, quantity entry 12→24).
- **UI – KMAT scan dialog (Position selection popup):** All font sizes doubled (labels, combobox input, dropdown list). Window enlarged from 400×200 to 600×350 px.

### Version 2.4 – May 2026
- WIP warehouse mode added (Sales Order + Order selection, quantity input).

---

**Developed for Forbo Movement Systems**
*Version 3.0 – October 2026*
