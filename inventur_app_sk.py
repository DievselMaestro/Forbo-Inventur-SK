#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
INVENTORY PROGRAM FOR FORBO SK - MALACKY WAREHOUSE MANAGEMENT
==============================================================

Desktop application for warehouse inventory with barcode scanner integration.
Supports Rolls only (no granulate).
Developed for Windows 11, Python 3.11+

Date: March 2026
Version: 1.0 SK
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from datetime import datetime
import os
import sys
import shutil
from pathlib import Path
import json
import logging


# ---------------------------------------------------------------------------
# Settings Dialog
# ---------------------------------------------------------------------------

class SettingsDialog:
    """Dialog for configuring file paths (master table + export folder)."""

    def __init__(self, parent, current_config):
        self.result = None
        self.config = current_config.copy()

        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Settings")
        self.dialog.geometry("620x280")
        self.dialog.resizable(True, False)
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.dialog.geometry(
            "+%d+%d" % (parent.winfo_rootx() + 60, parent.winfo_rooty() + 60)
        )

        self._build_widgets()
        self.dialog.wait_window()

    def _build_widgets(self):
        frame = ttk.Frame(self.dialog, padding="15")
        frame.pack(fill=tk.BOTH, expand=True)

        # --- Master Table path ---
        ttk.Label(frame, text="Master Table file (*.xlsx):", font=("Arial", 10, "bold")).grid(
            row=0, column=0, sticky=tk.W, pady=(0, 4))

        self.master_path_var = tk.StringVar(
            value=self.config.get("arbeitstabelle_path", ""))
        master_entry = ttk.Entry(frame, textvariable=self.master_path_var, width=55)
        master_entry.grid(row=1, column=0, sticky=(tk.W, tk.E), padx=(0, 8))

        ttk.Button(frame, text="Browse...", command=self._browse_master).grid(
            row=1, column=1, sticky=tk.W)

        # --- Export folder ---
        ttk.Label(frame, text="Export output folder:", font=("Arial", 10, "bold")).grid(
            row=2, column=0, sticky=tk.W, pady=(16, 4))

        self.export_path_var = tk.StringVar(
            value=self.config.get("export_path", ""))
        export_entry = ttk.Entry(frame, textvariable=self.export_path_var, width=55)
        export_entry.grid(row=3, column=0, sticky=(tk.W, tk.E), padx=(0, 8))

        ttk.Button(frame, text="Browse...", command=self._browse_export).grid(
            row=3, column=1, sticky=tk.W)

        frame.columnconfigure(0, weight=1)

        # --- Buttons ---
        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=4, column=0, columnspan=2, pady=(24, 0))

        ttk.Button(btn_frame, text="Save", command=self._save, width=12).pack(
            side=tk.LEFT, padx=(0, 12))
        ttk.Button(btn_frame, text="Cancel", command=self._cancel, width=12).pack(
            side=tk.LEFT)

        self.dialog.bind("<Return>", lambda e: self._save())
        self.dialog.bind("<Escape>", lambda e: self._cancel())

    def _browse_master(self):
        path = filedialog.askopenfilename(
            title="Select Master Table file",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if path:
            self.master_path_var.set(path)

    def _browse_export(self):
        path = filedialog.askdirectory(title="Select Export output folder")
        if path:
            self.export_path_var.set(path)

    def _save(self):
        self.config["arbeitstabelle_path"] = self.master_path_var.get().strip()
        self.config["export_path"] = self.export_path_var.get().strip()
        self.result = self.config
        self.dialog.destroy()

    def _cancel(self):
        self.result = None
        self.dialog.destroy()


# ---------------------------------------------------------------------------
# NotFoundDialogSK
# ---------------------------------------------------------------------------

class NotFoundDialogSK:
    """Dialog for rolls not found in the master table.
    Pre-fills Batch No., Location and all QR dimensions; user supplies Material,
    Description, Shelf Location (mandatory) and Remarks (optional).
    Always treated as a roll — no type selector.
    """

    def __init__(self, parent, charge, qr_data):
        self.result = None

        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Roll Not Found - Manual Entry")
        self.dialog.geometry("560x640")
        self.dialog.resizable(True, True)
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.dialog.geometry(
            "+%d+%d" % (parent.winfo_rootx() + 60, parent.winfo_rooty() + 60)
        )

        self._qr = qr_data  # dict from parse_qr_code
        self._build_widgets(charge)
        self.dialog.wait_window()

    def _build_widgets(self, charge):
        canvas = tk.Canvas(self.dialog, borderwidth=0)
        scrollbar = ttk.Scrollbar(self.dialog, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        inner = ttk.Frame(canvas, padding="15")
        win_id = canvas.create_window((0, 0), window=inner, anchor="nw")

        def _on_frame_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _on_canvas_configure(event):
            canvas.itemconfig(win_id, width=event.width)

        inner.bind("<Configure>", _on_frame_configure)
        canvas.bind("<Configure>", _on_canvas_configure)

        # Title
        ttk.Label(
            inner,
            text=f"Roll not found in master table!\nBatch No.: {charge}",
            font=("Arial", 12, "bold"),
            foreground="red",
            justify=tk.CENTER,
        ).pack(pady=(0, 16))

        # Pre-filled / read-only info
        info_frame = ttk.LabelFrame(inner, text="Data from QR Code (read-only)", padding="10")
        info_frame.pack(fill=tk.X, pady=(0, 12))
        info_frame.columnconfigure(1, weight=1)

        pre_filled = [
            ("Batch No.:", str(self._qr.get("charge", charge))),
            ("Location (QR):", str(self._qr.get("lort", ""))),
            ("Stage 0 - Length (mm):", str(self._qr.get("lnge0", ""))),
            ("Stage 0 - Width (mm):", str(self._qr.get("brte0", ""))),
            ("Stage 1 - Length (mm):", str(self._qr.get("lnge1", ""))),
            ("Stage 1 - Width (mm):", str(self._qr.get("brte1", ""))),
            ("Stage 2 - Length (mm):", str(self._qr.get("lnge2", ""))),
            ("Stage 2 - Width (mm):", str(self._qr.get("brte2", ""))),
        ]

        for r, (lbl, val) in enumerate(pre_filled):
            ttk.Label(info_frame, text=lbl, font=("Arial", 9, "bold")).grid(
                row=r, column=0, sticky=tk.W, pady=1)
            ttk.Label(info_frame, text=val, font=("Arial", 9)).grid(
                row=r, column=1, sticky=tk.W, padx=(8, 0), pady=1)

        # Manual fields
        manual_frame = ttk.LabelFrame(inner, text="Manual Input", padding="10")
        manual_frame.pack(fill=tk.X, pady=(0, 12))

        self.material_var = tk.StringVar()
        self.kurztext_var = tk.StringVar()
        self.fach_var = tk.StringVar()
        self.brte_meas_var = tk.StringVar()
        self.remarks_var = tk.StringVar()

        fields = [
            ("Material No. *:", self.material_var, True),
            ("Description:", self.kurztext_var, False),
            ("Shelf Location *:", self.fach_var, True),
            ("Measured Width (mm) *:", self.brte_meas_var, True),
            ("Remarks:", self.remarks_var, False),
        ]

        self._entries = {}
        for i, (label, var, _required) in enumerate(fields):
            field_frame = ttk.Frame(manual_frame)
            field_frame.pack(fill=tk.X, pady=4)
            ttk.Label(field_frame, text=label, font=("Arial", 9)).pack(anchor=tk.W)
            entry = ttk.Entry(field_frame, textvariable=var, width=50, font=("Arial", 10))
            entry.pack(fill=tk.X)
            self._entries[label] = entry
            if i == 0:
                entry.focus_set()

        # Buttons
        btn_frame = ttk.Frame(inner)
        btn_frame.pack(pady=(16, 4))
        ttk.Button(btn_frame, text="Save", command=self._save, width=12).pack(
            side=tk.LEFT, padx=(0, 12))
        ttk.Button(btn_frame, text="Cancel", command=self._cancel, width=12).pack(
            side=tk.LEFT)

        self.dialog.bind("<Escape>", lambda e: self._cancel())

    def _save(self):
        material = self.material_var.get().strip()
        fach = self.fach_var.get().strip()
        brte_meas = self.brte_meas_var.get().strip()

        if not material:
            messagebox.showerror("Error", "Material is a required field.", parent=self.dialog)
            return
        if not fach:
            messagebox.showerror("Error", "Shelf Location is a required field.", parent=self.dialog)
            return
        if not brte_meas:
            messagebox.showerror("Error", "Measured Width (mm) is a required field.", parent=self.dialog)
            return

        self.result = {
            "charge": str(self._qr.get("charge", "")),
            "lort_master": "",
            "lort_qr": str(self._qr.get("lort", "")),
            "material": material,
            "kurztext": self.kurztext_var.get().strip(),
            "werk": "",
            "lnge0": self._qr.get("lnge0", 0),
            "brte0": self._qr.get("brte0", 0),
            "lnge1": self._qr.get("lnge1", 0),
            "brte1": self._qr.get("brte1", 0),
            "lnge2": self._qr.get("lnge2", 0),
            "brte2": self._qr.get("brte2", 0),
            "flache": "",
            "frei_verw": "",
            "fach": fach,
            "brte_meas": brte_meas,
            "remarks": self.remarks_var.get().strip(),
            "status": "not_found",
        }
        self.dialog.destroy()

    def _cancel(self):
        self.result = None
        self.dialog.destroy()


# ---------------------------------------------------------------------------
# Main Application Class
# ---------------------------------------------------------------------------

class InventurAppSK:
    """Inventory application for Forbo SK (Malacky) — Rolls only."""

    # ------------------------------------------------------------------
    # Initialisation
    # ------------------------------------------------------------------

    def __init__(self):
        self.root = tk.Tk()

        # Base paths (EXE-compatible)
        self.base_dir = Path(self.get_base_path())
        self.config_dir = self.base_dir / "config"
        self.config_dir.mkdir(exist_ok=True)

        # Logging (needs config_dir first)
        self.setup_logging()

        # Config (may override paths)
        self.load_config()

        # Resolve arbeitstabelle / export paths from config
        self._resolve_paths()

        # Data
        self.init_data()

        # UI
        self.setup_ui()

        # Check paths & load master table
        self.check_paths_on_startup()

        # Load existing session
        self.load_existing_cz()

        # Shortcuts
        self.bind_shortcuts()

    # ------------------------------------------------------------------
    # Path / Base helpers
    # ------------------------------------------------------------------

    def get_base_path(self):
        """Return base path — works for both .py script and PyInstaller .exe."""
        if getattr(sys, "frozen", False):
            return os.path.dirname(sys.executable)
        return os.path.dirname(os.path.abspath(__file__))

    def _resolve_paths(self):
        """Derive runtime paths from config."""
        arb = self.config.get("arbeitstabelle_path", "")
        self.arbeitstabelle_path = Path(arb) if arb else None

        exp = self.config.get("export_path", "")
        if exp:
            self.export_path = Path(exp)
        else:
            self.export_path = self.base_dir

        self.inventur_path = self.export_path / "Inventory_Rolls_SK.xlsx"
        backup_dir = self.export_path / "backups"
        try:
            backup_dir.mkdir(parents=True, exist_ok=True)
        except Exception:
            pass

    # ------------------------------------------------------------------
    # Logging
    # ------------------------------------------------------------------

    def setup_logging(self):
        """Configure logging to config/inventory_sk.log."""
        log_file = self.config_dir / "inventory_sk.log"
        logging.basicConfig(
            level=logging.INFO,
            format="%(asctime)s - %(levelname)s - %(message)s",
            handlers=[
                logging.FileHandler(log_file, encoding="utf-8"),
                logging.StreamHandler(),
            ],
        )
        self.logger = logging.getLogger(__name__)
        self.logger.info("Inventory SK application started")

    # ------------------------------------------------------------------
    # Config
    # ------------------------------------------------------------------

    def load_config(self):
        """Load config from config/settings_sk.json."""
        config_path = self.config_dir / "settings_sk.json"
        default_config = {
            "auto_save": True,
            "arbeitstabelle_path": "",
            "export_path": "",
            "vollbild": True,
        }
        try:
            if config_path.exists():
                with open(config_path, "r", encoding="utf-8-sig") as f:
                    self.config = json.load(f)
                # Merge any missing keys
                for k, v in default_config.items():
                    self.config.setdefault(k, v)
            else:
                self.config = default_config
                self.save_config()
        except Exception as e:
            self.config = default_config
            self.logger.error(f"Error loading config: {e}")

    def save_config(self):
        """Save config to config/settings_sk.json."""
        try:
            config_path = self.config_dir / "settings_sk.json"
            with open(config_path, "w", encoding="utf-8") as f:
                json.dump(self.config, f, indent=2, ensure_ascii=False)
        except Exception as e:
            self.logger.error(f"Error saving config: {e}")

    # ------------------------------------------------------------------
    # Data initialisation
    # ------------------------------------------------------------------

    def init_data(self):
        """Initialise data structures."""
        self.df_rollen = None
        self.inventur_data = []       # found rolls
        self.not_found_data = []      # not-found rolls
        self.current_scan = None
        self.current_qr_data = None
        self.undo_stack = []

    # ------------------------------------------------------------------
    # QR parsing
    # ------------------------------------------------------------------

    def parse_qr_code(self, raw):
        """Parse semicolon-delimited QR code.

        Format: Lort;Charge;Lnge0;Brte0;Lnge1;Brte1;Lnge2;Brte2
        Fallback: treat whole string as Charge.

        Returns dict with keys:
            lort, charge, lnge0, brte0, lnge1, brte1, lnge2, brte2
        """
        raw = raw.strip()
        if ";" in raw:
            parts = raw.split(";")

            def _int(val, default=0):
                try:
                    return int(val.strip())
                except (ValueError, AttributeError):
                    return default

            return {
                "lort": parts[0].strip() if len(parts) > 0 else "",
                "charge": parts[1].strip() if len(parts) > 1 else raw,
                "lnge0": _int(parts[2]) if len(parts) > 2 else 0,
                "brte0": _int(parts[3]) if len(parts) > 3 else 0,
                "lnge1": _int(parts[4]) if len(parts) > 4 else 0,
                "brte1": _int(parts[5]) if len(parts) > 5 else 0,
                "lnge2": _int(parts[6]) if len(parts) > 6 else 0,
                "brte2": _int(parts[7]) if len(parts) > 7 else 0,
            }
        else:
            return {
                "lort": "",
                "charge": raw,
                "lnge0": 0,
                "brte0": 0,
                "lnge1": 0,
                "brte1": 0,
                "lnge2": 0,
                "brte2": 0,
            }

    # ------------------------------------------------------------------
    # Master table
    # ------------------------------------------------------------------

    def load_arbeitstabelle(self):
        """Load SK master table from the configured path (single sheet)."""
        if self.arbeitstabelle_path is None or not self.arbeitstabelle_path.exists():
            self.df_rollen = None
            self.logger.warning("Master table not loaded (path not set or file missing)")
            return

        try:
            self.df_rollen = pd.read_excel(
                self.arbeitstabelle_path,
                dtype={"Charge": str},
            )
            if "Charge" in self.df_rollen.columns:
                self.df_rollen["Charge"] = self.df_rollen["Charge"].astype(str)

            count = len(self.df_rollen)
            self.logger.info(f"Master table loaded: {count} rows from {self.arbeitstabelle_path}")
            self.status_var.set(f"Master table loaded: {count} rolls")

            # Refresh header info label
            self._update_header_info()

        except Exception as e:
            messagebox.showerror("Error", f"Error loading master table:\n{e}")
            self.logger.error(f"Error loading master table: {e}")
            self.df_rollen = None

    # ------------------------------------------------------------------
    # Charge lookup
    # ------------------------------------------------------------------

    def suche_charge(self, charge):
        """Look up a charge in df_rollen. Returns row dict or None."""
        if self.df_rollen is None:
            return None
        matches = self.df_rollen[self.df_rollen["Charge"] == str(charge)]
        if not matches.empty:
            return matches.iloc[0].to_dict()
        return None

    # ------------------------------------------------------------------
    # Startup path check
    # ------------------------------------------------------------------

    def check_paths_on_startup(self):
        """Show a warning / prompt to locate file if master path is not set."""
        arb_missing = (
            self.arbeitstabelle_path is None
            or not self.arbeitstabelle_path.exists()
        )

        if arb_missing:
            answer = messagebox.askyesno(
                "Master Table Not Found",
                "The master table file has not been configured or could not be found.\n\n"
                "Would you like to locate it now?",
            )
            if answer:
                path = filedialog.askopenfilename(
                    title="Select Master Table file",
                    filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                )
                if path:
                    self.config["arbeitstabelle_path"] = path
                    self.save_config()
                    self._resolve_paths()
                    self.load_arbeitstabelle()
                else:
                    self._disable_scan("No master table selected. Scanning disabled.")
            else:
                self._disable_scan("Master table not configured. Scanning disabled.")
        else:
            self.load_arbeitstabelle()

        # Default export path to base_dir if not set
        if not self.config.get("export_path", ""):
            self.config["export_path"] = str(self.base_dir)
            self.save_config()
            self._resolve_paths()

    def _disable_scan(self, msg):
        self.scan_entry.config(state="disabled")
        self.status_var.set(f"WARNING: {msg}")
        self.logger.warning(msg)

    # ------------------------------------------------------------------
    # Settings dialog
    # ------------------------------------------------------------------

    def open_settings_dialog(self):
        """Open the settings dialog and apply any changes."""
        dlg = SettingsDialog(self.root, self.config)
        if dlg.result:
            self.config.update(dlg.result)
            self.save_config()
            self._resolve_paths()
            # Re-enable scan if it was disabled
            self.scan_entry.config(state="normal")
            self.load_arbeitstabelle()
            messagebox.showinfo("Settings Saved", "Settings have been saved.\nMaster table reloaded.")

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def setup_ui(self):
        """Build the main UI."""
        self.root.title("INVENTORY Forbo SK - Malacky Warehouse Management")
        self.root.geometry("1280x820")
        self.root.configure(bg="#f0f0f0")

        try:
            icon_path = Path(__file__).parent / "icon.ico"
            self.root.iconbitmap(str(icon_path))
        except Exception:
            pass

        self.root.state("zoomed")

        self.main_frame = ttk.Frame(self.root, padding="10")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        self.main_frame.columnconfigure(0, weight=1)

        self.create_header()
        self.create_scan_section()
        self.create_current_scan_section()
        self.create_list_section()
        self.create_button_section()
        self.create_status_bar()

        self.scan_entry.focus_set()

    # --- Header ---

    def create_header(self):
        header_frame = ttk.Frame(self.main_frame)
        header_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 16))
        header_frame.columnconfigure(1, weight=1)

        ttk.Label(header_frame, text="[F]", font=("Arial", 20, "bold"),
                  foreground="#1f4e79").grid(row=0, column=0, padx=(0, 16))

        ttk.Label(
            header_frame,
            text="INVENTORY Forbo SK - Malacky Warehouse Management",
            font=("Arial", 18, "bold"),
            foreground="#1f4e79",
        ).grid(row=0, column=1, sticky=tk.W)

        self.header_info_var = tk.StringVar(value="No master table loaded")
        ttk.Label(header_frame, textvariable=self.header_info_var,
                  font=("Arial", 10)).grid(row=0, column=2, padx=(20, 0))

    def _update_header_info(self):
        if self.df_rollen is not None:
            self.header_info_var.set(f"DB: {len(self.df_rollen)} rolls in master table")
        else:
            self.header_info_var.set("No master table loaded")

    # --- Scan section ---

    def create_scan_section(self):
        scan_frame = ttk.LabelFrame(self.main_frame, text="BARCODE / QR SCAN", padding="10")
        scan_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        scan_frame.columnconfigure(0, weight=1)

        self.scan_var = tk.StringVar()
        self.scan_entry = ttk.Entry(scan_frame, textvariable=self.scan_var,
                                    font=("Arial", 14), width=50)
        self.scan_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10))

        ttk.Button(scan_frame, text="Scan", command=self.process_scan).grid(row=0, column=1)

        self.scan_entry.bind("<Return>", lambda e: self.process_scan())

        ttk.Label(
            scan_frame,
            text="Scan field is always focused. Scan a QR code or type + ENTER.",
            font=("Arial", 9),
            foreground="gray",
        ).grid(row=1, column=0, columnspan=2, pady=(5, 0))

    # --- Current scan section ---

    def create_current_scan_section(self):
        self.current_frame = ttk.LabelFrame(
            self.main_frame, text="CURRENT SCAN", padding="10")
        self.current_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        self.current_frame.columnconfigure(1, weight=1)
        self.current_frame.columnconfigure(3, weight=1)

        self._create_display_labels()
        self._create_input_panel()

        self.current_frame.grid_remove()

    def _create_display_labels(self):
        """Create read-only display labels for scan data."""
        # Row 0: Batch No. | Material
        ttk.Label(self.current_frame, text="Batch No.:", font=("Arial", 11, "bold")).grid(
            row=0, column=0, sticky=tk.W, pady=2)
        self.lbl_charge = ttk.Label(self.current_frame, text="", font=("Arial", 11))
        self.lbl_charge.grid(row=0, column=1, sticky=tk.W, padx=(8, 20), pady=2)

        ttk.Label(self.current_frame, text="Material:", font=("Arial", 11, "bold")).grid(
            row=0, column=2, sticky=tk.W, pady=2)
        self.lbl_material = ttk.Label(self.current_frame, text="", font=("Arial", 11))
        self.lbl_material.grid(row=0, column=3, sticky=tk.W, padx=(8, 0), pady=2)

        # Row 1: Description
        ttk.Label(self.current_frame, text="Description:", font=("Arial", 11, "bold")).grid(
            row=1, column=0, sticky=tk.W, pady=2)
        self.lbl_kurztext = ttk.Label(self.current_frame, text="", font=("Arial", 11))
        self.lbl_kurztext.grid(row=1, column=1, columnspan=3, sticky=tk.W, padx=(8, 0), pady=2)

        # Row 2: Plant | Location (Master)
        ttk.Label(self.current_frame, text="Plant:", font=("Arial", 11, "bold")).grid(
            row=2, column=0, sticky=tk.W, pady=2)
        self.lbl_werk = ttk.Label(self.current_frame, text="", font=("Arial", 11))
        self.lbl_werk.grid(row=2, column=1, sticky=tk.W, padx=(8, 20), pady=2)

        ttk.Label(self.current_frame, text="Location (Master):", font=("Arial", 11, "bold")).grid(
            row=2, column=2, sticky=tk.W, pady=2)
        self.lbl_lort_master = ttk.Label(self.current_frame, text="", font=("Arial", 11))
        self.lbl_lort_master.grid(row=2, column=3, sticky=tk.W, padx=(8, 0), pady=2)

        # Row 3: Location QR | Area
        ttk.Label(self.current_frame, text="Location (QR):", font=("Arial", 11, "bold")).grid(
            row=3, column=0, sticky=tk.W, pady=2)
        self.lbl_lort_qr = ttk.Label(self.current_frame, text="", font=("Arial", 11))
        self.lbl_lort_qr.grid(row=3, column=1, sticky=tk.W, padx=(8, 20), pady=2)

        ttk.Label(self.current_frame, text="Area (m2):", font=("Arial", 11, "bold")).grid(
            row=3, column=2, sticky=tk.W, pady=2)
        self.lbl_flache = ttk.Label(self.current_frame, text="", font=("Arial", 11))
        self.lbl_flache.grid(row=3, column=3, sticky=tk.W, padx=(8, 0), pady=2)

        # Row 4: Free Usable
        ttk.Label(self.current_frame, text="Free Usable:", font=("Arial", 11, "bold")).grid(
            row=4, column=0, sticky=tk.W, pady=2)
        self.lbl_frei = ttk.Label(self.current_frame, text="", font=("Arial", 11))
        self.lbl_frei.grid(row=4, column=1, sticky=tk.W, padx=(8, 0), pady=2)

        # Row 5: Dimensions grid
        dim_frame = ttk.LabelFrame(self.current_frame, text="Dimensions (mm)", padding="6")
        dim_frame.grid(row=5, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=(8, 4))
        dim_frame.columnconfigure(1, weight=1)
        dim_frame.columnconfigure(3, weight=1)
        dim_frame.columnconfigure(5, weight=1)

        ttk.Label(dim_frame, text="Stage 0", font=("Arial", 10, "bold")).grid(
            row=0, column=0, columnspan=2, sticky=tk.W, padx=(0, 16))
        ttk.Label(dim_frame, text="Stage 1", font=("Arial", 10, "bold")).grid(
            row=0, column=2, columnspan=2, sticky=tk.W, padx=(0, 16))
        ttk.Label(dim_frame, text="Stage 2", font=("Arial", 10, "bold")).grid(
            row=0, column=4, columnspan=2, sticky=tk.W)

        # Stage 0
        ttk.Label(dim_frame, text="Length:", font=("Arial", 9)).grid(
            row=1, column=0, sticky=tk.W, padx=(0, 4))
        self.lbl_lnge0 = ttk.Label(dim_frame, text="", font=("Arial", 10, "bold"), width=8)
        self.lbl_lnge0.grid(row=1, column=1, sticky=tk.W, padx=(0, 16))

        # Stage 1
        ttk.Label(dim_frame, text="Length:", font=("Arial", 9)).grid(
            row=1, column=2, sticky=tk.W, padx=(0, 4))
        self.lbl_lnge1 = ttk.Label(dim_frame, text="", font=("Arial", 10, "bold"), width=8)
        self.lbl_lnge1.grid(row=1, column=3, sticky=tk.W, padx=(0, 16))

        # Stage 2
        ttk.Label(dim_frame, text="Length:", font=("Arial", 9)).grid(
            row=1, column=4, sticky=tk.W, padx=(0, 4))
        self.lbl_lnge2 = ttk.Label(dim_frame, text="", font=("Arial", 10, "bold"), width=8)
        self.lbl_lnge2.grid(row=1, column=5, sticky=tk.W)

        ttk.Label(dim_frame, text="Width:", font=("Arial", 9)).grid(
            row=2, column=0, sticky=tk.W, padx=(0, 4))
        self.lbl_brte0 = ttk.Label(dim_frame, text="", font=("Arial", 10, "bold"), width=8)
        self.lbl_brte0.grid(row=2, column=1, sticky=tk.W, padx=(0, 16))

        ttk.Label(dim_frame, text="Width:", font=("Arial", 9)).grid(
            row=2, column=2, sticky=tk.W, padx=(0, 4))
        self.lbl_brte1 = ttk.Label(dim_frame, text="", font=("Arial", 10, "bold"), width=8)
        self.lbl_brte1.grid(row=2, column=3, sticky=tk.W, padx=(0, 16))

        ttk.Label(dim_frame, text="Width:", font=("Arial", 9)).grid(
            row=2, column=4, sticky=tk.W, padx=(0, 4))
        self.lbl_brte2 = ttk.Label(dim_frame, text="", font=("Arial", 10, "bold"), width=8)
        self.lbl_brte2.grid(row=2, column=5, sticky=tk.W)

        ttk.Separator(self.current_frame, orient="horizontal").grid(
            row=6, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=8)

    def _create_input_panel(self):
        """Create user input widgets (Fach + Measured Width + Remarks)."""
        self.fach_var = tk.StringVar()
        self.brte_meas_var = tk.StringVar()
        self.remarks_var = tk.StringVar()
        self.input_widgets = {}

        self.input_container = ttk.Frame(self.current_frame)
        self.input_container.grid(row=7, column=0, columnspan=4,
                                   sticky=(tk.W, tk.E), pady=4)
        self.input_container.columnconfigure(1, weight=1)

        # Fach (mandatory) — row 0
        lbl_fach = ttk.Label(self.input_container, text="Shelf Location *:",
                              font=("Arial", 11, "bold"))
        lbl_fach.grid(row=0, column=0, sticky=tk.W, pady=5)
        entry_fach = ttk.Entry(self.input_container, textvariable=self.fach_var,
                                font=("Arial", 12), width=20)
        entry_fach.grid(row=0, column=1, sticky=tk.W, padx=(10, 0), pady=5)
        entry_fach.bind("<Return>", lambda e: self.input_widgets["brte_meas_entry"].focus_set())
        self.input_widgets["fach_entry"] = entry_fach

        # Measured Width (mandatory) — row 1
        lbl_brte_meas = ttk.Label(self.input_container, text="Measured Width (mm) *:",
                                   font=("Arial", 11, "bold"))
        lbl_brte_meas.grid(row=1, column=0, sticky=tk.W, pady=5)
        entry_brte_meas = ttk.Entry(self.input_container, textvariable=self.brte_meas_var,
                                     font=("Arial", 12), width=10)
        entry_brte_meas.grid(row=1, column=1, sticky=tk.W, padx=(10, 0), pady=5)
        entry_brte_meas.bind("<Return>", self.save_current_scan)
        self.input_widgets["brte_meas_entry"] = entry_brte_meas

        # Remarks (optional) — row 2
        lbl_remarks = ttk.Label(self.input_container, text="Remarks (optional):",
                                 font=("Arial", 11, "bold"))
        lbl_remarks.grid(row=2, column=0, sticky=tk.W, pady=5)
        entry_remarks = ttk.Entry(self.input_container, textvariable=self.remarks_var,
                                   font=("Arial", 12), width=50)
        entry_remarks.grid(row=2, column=1, sticky=tk.W, padx=(10, 0), pady=5)
        entry_remarks.bind("<Return>", self.save_current_scan)
        self.input_widgets["remarks_entry"] = entry_remarks

        # Save button — row 3
        save_btn = ttk.Button(self.input_container, text="Save",
                               command=self.save_current_scan)
        save_btn.grid(row=3, column=1, sticky=tk.W, padx=(10, 0), pady=8)
        self.input_widgets["save_button"] = save_btn

    # --- List section ---

    def create_list_section(self):
        list_frame = ttk.LabelFrame(self.main_frame, text="SCANNED ROLLS", padding="10")
        list_frame.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(1, weight=1)
        self.main_frame.rowconfigure(3, weight=1)

        count_frame = ttk.Frame(list_frame)
        count_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 8))

        self.count_label = ttk.Label(count_frame, text="0 rolls", font=("Arial", 12, "bold"))
        self.count_label.pack(side=tk.LEFT)

        columns = ("Time", "Batch No.", "Material", "Shelf Location", "Status")
        self.tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=12)

        for col in columns:
            self.tree.heading(col, text=col)

        self.tree.column("Time", width=90, minwidth=70)
        self.tree.column("Batch No.", width=130, minwidth=100)
        self.tree.column("Material", width=110, minwidth=80)
        self.tree.column("Shelf Location", width=100, minwidth=70)
        self.tree.column("Status", width=130, minwidth=100)

        self.tree.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        v_scroll = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.tree.yview)
        v_scroll.grid(row=1, column=1, sticky=(tk.N, tk.S))
        self.tree.configure(yscrollcommand=v_scroll.set)

        h_scroll = ttk.Scrollbar(list_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        h_scroll.grid(row=2, column=0, sticky=(tk.W, tk.E))
        self.tree.configure(xscrollcommand=h_scroll.set)

        self.tree.bind("<Button-3>", self.show_context_menu)

    # --- Button section ---

    def create_button_section(self):
        btn_frame = ttk.Frame(self.main_frame)
        btn_frame.grid(row=4, column=0, sticky=(tk.W, tk.E), pady=(8, 0))

        ttk.Button(btn_frame, text="Export / Backup", command=self.export_backup).grid(
            row=0, column=0, padx=(0, 10))
        ttk.Button(btn_frame, text="Settings", command=self.open_settings_dialog).grid(
            row=0, column=1, padx=(0, 10))
        ttk.Button(btn_frame, text="Fullscreen (F11)", command=self.toggle_fullscreen).grid(
            row=0, column=2, padx=(0, 10))
        ttk.Button(btn_frame, text="Quit", command=self.quit_app).grid(
            row=0, column=3)

    # --- Status bar ---

    def create_status_bar(self):
        self.status_var = tk.StringVar(value="Ready to scan...")
        status_bar = ttk.Label(
            self.root, textvariable=self.status_var,
            relief=tk.SUNKEN, anchor=tk.W, font=("Arial", 9))
        status_bar.grid(row=1, column=0, sticky=(tk.W, tk.E))

    # ------------------------------------------------------------------
    # Shortcuts
    # ------------------------------------------------------------------

    def bind_shortcuts(self):
        self.root.bind("<Control-z>", lambda e: self.undo_last_action())
        self.root.bind("<Control-Z>", lambda e: self.undo_last_action())
        self.root.bind("<Control-s>", lambda e: self._manual_save())
        self.root.bind("<Control-S>", lambda e: self._manual_save())
        self.root.bind("<Escape>", lambda e: self._reset_scan())
        self.root.bind("<F11>", lambda e: self.toggle_fullscreen())
        self.root.bind("<FocusIn>", self.ensure_scan_focus)

    # ------------------------------------------------------------------
    # Scan processing
    # ------------------------------------------------------------------

    def process_scan(self):
        """Process a scanned QR / barcode."""
        raw = self.scan_var.get().strip()
        if not raw:
            self.status_var.set("Please scan or enter a barcode / QR code.")
            return

        qr = self.parse_qr_code(raw)
        self.current_qr_data = qr
        charge = qr["charge"]

        self.logger.info(f"Scan received: raw='{raw}' -> charge='{charge}'")

        # Duplicate check
        if self._is_already_scanned(charge):
            messagebox.showwarning(
                "Already Scanned",
                f"Batch No. '{charge}' has already been scanned!\n\n"
                "Please check the list of scanned items.",
            )
            self._reset_scan()
            return

        row_data = self.suche_charge(charge)

        if row_data is not None:
            self.show_found_rolle(row_data, qr)
        else:
            self.show_not_found_dialog(charge, qr)

    def _is_already_scanned(self, charge):
        for item in self.inventur_data + self.not_found_data:
            if str(item.get("charge", "")) == str(charge):
                return True
        return False

    # ------------------------------------------------------------------
    # Show found roll
    # ------------------------------------------------------------------

    def show_found_rolle(self, row_data, qr):
        """Display found roll data and show the input panel."""
        def _fmt_int(val):
            """Format numeric value as integer string (remove .0)."""
            if val == "" or val is None:
                return ""
            try:
                return str(int(float(val)))
            except (ValueError, TypeError):
                return str(val)

        def _fmt_float2(val):
            """Format float to 2 decimal places, empty if missing."""
            if val == "" or val is None:
                return ""
            try:
                return f"{float(val):.2f}"
            except (ValueError, TypeError):
                return str(val)

        # Compute Area (m2) from Länge m × Breite mm / 1000
        laenge_m = row_data.get("Länge m", None)
        breite_mm = row_data.get("Breite mm", None)
        if laenge_m not in (None, "") and breite_mm not in (None, ""):
            try:
                flache_val = f"{float(laenge_m) * float(breite_mm) / 1000:.2f}"
            except (ValueError, TypeError):
                flache_val = ""
        else:
            flache_val = ""

        self.current_scan = {
            "charge": qr["charge"],
            "lort_master": _fmt_int(row_data.get("Lagerort", "") or ""),
            "lort_qr": qr["lort"],
            "material": _fmt_int(row_data.get("Material", "") or ""),
            "kurztext": str(row_data.get("Materialkurztext", "") or ""),
            "werk": _fmt_int(row_data.get("Werk", "") or ""),
            "lnge0": qr["lnge0"],
            "brte0": qr["brte0"],
            "lnge1": qr["lnge1"],
            "brte1": qr["brte1"],
            "lnge2": qr["lnge2"],
            "brte2": qr["brte2"],
            "flache": flache_val,
            "frei_verw": _fmt_float2(row_data.get("Frei verwendbar", "") or ""),
            "status": "found",
        }

        # Update display labels
        self.lbl_charge.config(text=self.current_scan["charge"])
        self.lbl_material.config(text=self.current_scan["material"])
        self.lbl_kurztext.config(text=self.current_scan["kurztext"])
        self.lbl_werk.config(text=self.current_scan["werk"])
        self.lbl_lort_master.config(text=self.current_scan["lort_master"])
        self.lbl_lort_qr.config(text=self.current_scan["lort_qr"])
        self.lbl_flache.config(text=str(self.current_scan["flache"]))
        self.lbl_frei.config(text=str(self.current_scan["frei_verw"]))
        self.lbl_lnge0.config(text=str(self.current_scan["lnge0"]))
        self.lbl_brte0.config(text=str(self.current_scan["brte0"]))
        self.lbl_lnge1.config(text=str(self.current_scan["lnge1"]))
        self.lbl_brte1.config(text=str(self.current_scan["brte1"]))
        self.lbl_lnge2.config(text=str(self.current_scan["lnge2"]))
        self.lbl_brte2.config(text=str(self.current_scan["brte2"]))

        self.current_frame.config(text="ROLL FOUND")
        self.current_frame.grid()

        # Reset input fields
        self.fach_var.set("")
        self.brte_meas_var.set("")
        self.remarks_var.set("")
        self.input_widgets["fach_entry"].focus_set()

        self.scan_var.set("")
        self.status_var.set(f"Roll found: {self.current_scan['kurztext']} | Enter Shelf Location")

    # ------------------------------------------------------------------
    # Not found dialog
    # ------------------------------------------------------------------

    def show_not_found_dialog(self, charge, qr):
        """Show dialog for rolls not in master table."""
        dlg = NotFoundDialogSK(self.root, charge, qr)

        if dlg.result:
            data = dlg.result.copy()
            data["zeitstempel"] = datetime.now().strftime("%d.%m.%Y %H:%M:%S")

            self.not_found_data.append(data)

            self.undo_stack.append(("add_not_found", data.copy()))
            if len(self.undo_stack) > 50:
                self.undo_stack.pop(0)

            if self.config.get("auto_save", True):
                self.save_cz_excel()

            self.update_list()
            total = len(self.inventur_data) + len(self.not_found_data)
            self.status_var.set(
                f"Not-found roll saved. Total: {total} "
                f"({len(self.inventur_data)} found, {len(self.not_found_data)} not found)"
            )
            self.logger.info(f"Not-found roll saved: {data['charge']}")
        else:
            self.logger.info("Not-found dialog cancelled")

        self._reset_scan()

    # ------------------------------------------------------------------
    # Save current scan
    # ------------------------------------------------------------------

    def save_current_scan(self, event=None):
        """Validate and save the current found roll."""
        if not self.current_scan:
            return

        fach = self.fach_var.get().strip()
        if not fach:
            messagebox.showwarning(
                "Required Field",
                "Shelf Location is mandatory. Please enter a value.",
            )
            self.input_widgets["fach_entry"].focus_set()
            return

        brte_meas = self.brte_meas_var.get().strip()
        if not brte_meas:
            messagebox.showwarning(
                "Required Field",
                "Measured Width (mm) is mandatory. Please enter a value.",
            )
            self.input_widgets["brte_meas_entry"].focus_set()
            return

        self.current_scan["fach"] = fach
        self.current_scan["brte_meas"] = brte_meas
        self.current_scan["remarks"] = self.remarks_var.get().strip()
        self.current_scan["zeitstempel"] = datetime.now().strftime("%d.%m.%Y %H:%M:%S")

        self.inventur_data.append(self.current_scan.copy())

        self.undo_stack.append(("add_found", self.current_scan.copy()))
        if len(self.undo_stack) > 50:
            self.undo_stack.pop(0)

        if self.config.get("auto_save", True):
            self.save_cz_excel()

        self.update_list()

        total = len(self.inventur_data) + len(self.not_found_data)
        self.status_var.set(
            f"Roll saved. Total: {total} "
            f"({len(self.inventur_data)} found, {len(self.not_found_data)} not found)"
        )
        self.logger.info(f"Roll saved: {self.current_scan['charge']}")

        self._reset_scan()

    # ------------------------------------------------------------------
    # Reset scan
    # ------------------------------------------------------------------

    def _reset_scan(self):
        self.current_scan = None
        self.current_qr_data = None
        self.scan_var.set("")
        self.fach_var.set("")
        self.brte_meas_var.set("")
        self.remarks_var.set("")
        self.current_frame.grid_remove()
        self.scan_entry.focus_set()
        self.status_var.set("Ready to scan...")

    # ------------------------------------------------------------------
    # List update
    # ------------------------------------------------------------------

    def update_list(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        all_items = []
        for d in self.inventur_data:
            all_items.append((d, "Found"))
        for d in self.not_found_data:
            all_items.append((d, "Not Found"))

        all_items.sort(key=lambda x: x[0].get("zeitstempel", ""), reverse=True)

        for d, status in all_items:
            ts = d.get("zeitstempel", "")
            time_part = ts.split(" ")[1] if " " in ts else ts
            item_id = self.tree.insert(
                "", "end",
                values=(
                    time_part,
                    d.get("charge", ""),
                    d.get("material", ""),
                    d.get("fach", ""),
                    status,
                ),
            )
            if status == "Not Found":
                self.tree.set(item_id, "Status", "NOT FOUND")
            else:
                self.tree.set(item_id, "Status", "Found")

        total = len(self.inventur_data) + len(self.not_found_data)
        self.count_label.config(
            text=f"{total} rolls ({len(self.inventur_data)} found, "
                 f"{len(self.not_found_data)} not found)"
        )

    # ------------------------------------------------------------------
    # Excel save
    # ------------------------------------------------------------------

    INVENTORY_HEADERS = [
        "Timestamp",
        "Plant",              # was: Werk
        "Location (Master)",  # was: LÖrt (Master)
        "Location (QR)",      # was: LÖrt (QR)
        "Material No.",       # was: Material
        "Description",        # was: Materialkurztext
        "Batch No.",          # was: Charge
        "Length S0 (mm)",     # was: Lnge
        "Width S0 (mm)",      # was: Brte
        "Length S1 (mm)",     # was: Lnge1
        "Width S1 (mm)",      # was: Brte1
        "Length S2 (mm)",     # was: Lnge2
        "Width S2 (mm)",      # was: Brte2
        "Area (m2)",          # was: Flache
        "Free Usable",        # was: +Frei verw.
        "Shelf Location",     # was: Shelf (Fach)
        "Measured Width (mm)",  # control value entered during scan
        "Remarks",
    ]

    def _row_from_item(self, d):
        def _clean(v):
            if v is None:
                return ""
            s = str(v)
            if s.lower() == "nan":
                return ""
            return s

        return [
            _clean(d.get("zeitstempel", "")),
            _clean(d.get("werk", "")),
            _clean(d.get("lort_master", "")),
            _clean(d.get("lort_qr", "")),
            _clean(d.get("material", "")),
            _clean(d.get("kurztext", "")),
            _clean(d.get("charge", "")),
            d.get("lnge0", ""),
            d.get("brte0", ""),
            d.get("lnge1", ""),
            d.get("brte1", ""),
            d.get("lnge2", ""),
            d.get("brte2", ""),
            _clean(d.get("flache", "")),
            _clean(d.get("frei_verw", "")),
            _clean(d.get("fach", "")),
            _clean(d.get("brte_meas", "")),
            _clean(d.get("remarks", "")),
        ]

    def _format_charge_col(self, ws, data_count):
        """Format Batch No. column (col 7 = G) as text."""
        charge_col = get_column_letter(7)
        for r in range(2, data_count + 2):
            ws[f"{charge_col}{r}"].number_format = "@"

    def save_cz_excel(self):
        """Write Inventory_Rolls_SK.xlsx with Inventory and Not_Found sheets."""
        try:
            wb = Workbook()
            if "Sheet" in wb.sheetnames:
                wb.remove(wb["Sheet"])

            # Inventory sheet
            ws_inv = wb.create_sheet("Inventory")
            ws_inv.append(self.INVENTORY_HEADERS)
            for d in self.inventur_data:
                ws_inv.append(self._row_from_item(d))
            self._format_charge_col(ws_inv, len(self.inventur_data))

            # Not_Found sheet
            ws_nf = wb.create_sheet("Not_Found")
            ws_nf.append(self.INVENTORY_HEADERS)
            for d in self.not_found_data:
                ws_nf.append(self._row_from_item(d))
            self._format_charge_col(ws_nf, len(self.not_found_data))

            # Ensure export dir exists
            self.export_path.mkdir(parents=True, exist_ok=True)
            wb.save(self.inventur_path)
            self.logger.info(f"Excel saved: {self.inventur_path}")

        except Exception as e:
            messagebox.showerror("Save Error", f"Error saving Excel file:\n{e}")
            self.logger.error(f"Error saving Excel: {e}")

    # ------------------------------------------------------------------
    # Load existing session
    # ------------------------------------------------------------------

    def load_existing_cz(self):
        """Resume a previous session by loading Inventory_Rolls_SK.xlsx."""
        if not self.inventur_path.exists():
            return

        loaded = 0
        try:
            df_inv = pd.read_excel(
                self.inventur_path, sheet_name="Inventory", dtype={"Batch No.": str})
            for _, row in df_inv.iterrows():
                self.inventur_data.append(self._row_to_dict(row, status="found"))
                loaded += 1
        except Exception as e:
            self.logger.error(f"Error loading Inventory sheet: {e}")

        try:
            df_nf = pd.read_excel(
                self.inventur_path, sheet_name="Not_Found", dtype={"Batch No.": str})
            for _, row in df_nf.iterrows():
                self.not_found_data.append(self._row_to_dict(row, status="not_found"))
                loaded += 1
        except Exception:
            pass  # Sheet may not exist yet

        if loaded:
            self.update_list()
            self.status_var.set(
                f"Session resumed: {len(self.inventur_data)} found, "
                f"{len(self.not_found_data)} not found."
            )
            self.logger.info(f"Existing session loaded: {loaded} rows")

    def _row_to_dict(self, row, status):
        def _str(v):
            if pd.isna(v):
                return ""
            s = str(v)
            return "" if s.lower() == "nan" else s

        def _int_or_empty(v):
            try:
                return int(v)
            except (ValueError, TypeError):
                return 0

        return {
            "zeitstempel": _str(row.get("Timestamp", "")),
            "werk": _str(row.get("Plant", "")),
            "lort_master": _str(row.get("Location (Master)", "")),
            "lort_qr": _str(row.get("Location (QR)", "")),
            "material": _str(row.get("Material No.", "")),
            "kurztext": _str(row.get("Description", "")),
            "charge": _str(row.get("Batch No.", "")),
            "lnge0": _int_or_empty(row.get("Length S0 (mm)", 0)),
            "brte0": _int_or_empty(row.get("Width S0 (mm)", 0)),
            "lnge1": _int_or_empty(row.get("Length S1 (mm)", 0)),
            "brte1": _int_or_empty(row.get("Width S1 (mm)", 0)),
            "lnge2": _int_or_empty(row.get("Length S2 (mm)", 0)),
            "brte2": _int_or_empty(row.get("Width S2 (mm)", 0)),
            "flache": _str(row.get("Area (m2)", "")),
            "frei_verw": _str(row.get("Free Usable", "")),
            "fach": _str(row.get("Shelf Location", "")),
            "brte_meas": _str(row.get("Measured Width (mm)", "")),
            "remarks": _str(row.get("Remarks", "")),
            "status": status,
        }

    # ------------------------------------------------------------------
    # Export / Backup
    # ------------------------------------------------------------------

    def export_backup(self):
        """Save a timestamped backup of the current Excel file."""
        try:
            # Always write the latest data first
            self.save_cz_excel()

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_dir = self.export_path / "backups"
            backup_dir.mkdir(parents=True, exist_ok=True)
            backup_file = backup_dir / f"Inventory_Rolls_SK_Backup_{timestamp}.xlsx"

            if self.inventur_path.exists():
                shutil.copy2(self.inventur_path, backup_file)
                messagebox.showinfo(
                    "Backup Created",
                    f"Backup saved to:\n{backup_file}",
                )
                self.logger.info(f"Backup created: {backup_file}")
            else:
                messagebox.showwarning("Warning", "No data file to backup yet.")
        except Exception as e:
            messagebox.showerror("Export Error", f"Error creating backup:\n{e}")
            self.logger.error(f"Backup error: {e}")

    # ------------------------------------------------------------------
    # Undo
    # ------------------------------------------------------------------

    def undo_last_action(self):
        if not self.undo_stack:
            self.status_var.set("Nothing to undo")
            return

        action, data = self.undo_stack.pop()
        charge = data.get("charge", "")

        if action == "add_found":
            self.inventur_data = [
                d for d in self.inventur_data if d.get("charge") != charge]
        elif action == "add_not_found":
            self.not_found_data = [
                d for d in self.not_found_data if d.get("charge") != charge]

        if self.config.get("auto_save", True):
            self.save_cz_excel()
        self.update_list()
        self.status_var.set(f"Undo: removed entry for batch '{charge}'")
        self.logger.info(f"Undo {action}: {charge}")

    # ------------------------------------------------------------------
    # Context menu
    # ------------------------------------------------------------------

    def show_context_menu(self, event):
        selection = self.tree.selection()
        if not selection:
            return
        item_id = selection[0]

        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label="Delete entry",
                         command=lambda: self._delete_entry(item_id))
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def _delete_entry(self, item_id):
        if not messagebox.askyesno("Confirm Delete",
                                   "Delete this entry permanently?"):
            return

        values = self.tree.item(item_id, "values")
        if len(values) >= 2:
            charge = values[1]
            self.inventur_data = [
                d for d in self.inventur_data if d.get("charge") != charge]
            self.not_found_data = [
                d for d in self.not_found_data if d.get("charge") != charge]

            self.save_cz_excel()
            self.update_list()
            self.status_var.set(f"Entry deleted: {charge}")
            self.logger.info(f"Entry deleted: {charge}")

    # ------------------------------------------------------------------
    # Misc helpers
    # ------------------------------------------------------------------

    def toggle_fullscreen(self):
        current = self.root.attributes("-fullscreen")
        self.root.attributes("-fullscreen", not current)
        if not current:
            self.status_var.set("Fullscreen ON — press F11 to exit")
        else:
            self.status_var.set("Fullscreen OFF")

    def ensure_scan_focus(self, event=None):
        if hasattr(self, "scan_entry") and not self.current_scan:
            self.root.after(100, lambda: self.scan_entry.focus_set())

    def _manual_save(self):
        self.save_cz_excel()
        self.status_var.set("Manually saved.")

    def quit_app(self):
        if messagebox.askyesno("Quit", "Are you sure you want to quit?"):
            self.logger.info("Application closed by user")
            self.root.quit()

    # ------------------------------------------------------------------
    # Run
    # ------------------------------------------------------------------

    def run(self):
        self.root.mainloop()


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    try:
        app = InventurAppSK()
        app.run()
    except Exception as e:
        messagebox.showerror("Critical Error", f"Unexpected error:\n{e}")
        logging.error(f"Critical error: {e}")
        raise


if __name__ == "__main__":
    main()
