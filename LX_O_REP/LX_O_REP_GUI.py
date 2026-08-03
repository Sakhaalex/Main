import sys
import os
# Ensure the script's own directory is on sys.path so sibling imports work
# regardless of which directory python is invoked from.
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import re
import csv
import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# --- Dependency Check (no subprocesses / no ghost popups) ---
_dep_root = tk.Tk()
_dep_root.withdraw()
try:
    import xlsxwriter
    import openpyxl
except ImportError as e:
    messagebox.showerror(
        "Missing Dependency",
        f"Missing package: {e.name}\nPlease run: pip install xlsxwriter openpyxl"
    )
    _dep_root.destroy()
    sys.exit(1)
_dep_root.destroy()
# -------------------------------------------------------------

from LX_CommCentre import LX_CommCentre
from LX_Orchestrator import Orchestrator

# ================== Palette ==================
CREME_BG    = "#FDFBF7"   # warm off-white  (window bg)
CREME_ALT   = "#F4F0EA"   # warm creme      (card / alt bg)
CARD_WHITE  = "#FFFFFF"   # card surface
BORDER_CLR  = "#E2DDD5"   # thin separator / card border
DARK_SLATE  = "#2C3E50"   # high-contrast button
ACCENT_GOLD = "#B5860D"   # validate / action accent
WARN_BG     = "#FFF3CD"   # warning banner background
WARN_FG     = "#856404"   # warning banner text
PASS_GREEN  = "#1D6F42"
FAIL_RED    = "#C0392B"
WARN_AMBER  = "#E67E22"

# ================== Typography ===============
FONT_HEADER    = ("Segoe UI", 14, "bold")
FONT_SUBHEADER = ("Segoe UI", 10, "bold")
FONT_BODY      = ("Segoe UI",  9)
FONT_MONO      = ("Consolas",  9)
# =============================================

# Cell coordinate pattern: 1–3 uppercase letters followed by digits (e.g. A1, AFG3)
_CELL_COORD_RE = re.compile(r"^[A-Z]{1,3}[0-9]+$")


def _resolve_csv(filename: str) -> str | None:
    """
    Deterministic 3-location waterfall search for a CSV file.
    Returns the first existing absolute path, or None if not found.

    Search order:
      1. Current Working Directory (os.getcwd())
      2. Script Directory          (os.path.dirname(__file__))
      3. Root parent folder        (os.path.abspath(os.path.join(__file__, "../../")))
    """
    search_dirs = [
        os.getcwd(),
        os.path.dirname(os.path.abspath(__file__)),
        os.path.abspath(os.path.join(os.path.abspath(__file__), "../../")),
    ]
    for d in search_dirs:
        candidate = os.path.join(d, filename)
        if os.path.isfile(candidate):
            return candidate
    return None


class O_REP_GUI:
    def __init__(self, root):
        self.root = root
        self.root.title("LX Report Generator (O_REP)")
        self.root.geometry("1060x740")
        self.root.configure(bg=CREME_BG)

        # --- Auto-resolve both CSV paths ---
        self.record_path   = _resolve_csv("Record.csv")
        self.contract_path = _resolve_csv("Record_Contract.csv")

        self.csv_data       = []
        self.csv_fieldnames = []
        self._warning_bar   = None  # reference to the warning banner widget

        self._build_styles()
        self._build_ui()
        self.load_csv()

    # ------------------------------------------------------------------
    # ttk Styles
    # ------------------------------------------------------------------

    def _build_styles(self):
        style = ttk.Style()
        style.theme_use("clam")

        style.configure("TFrame",        background=CREME_BG)
        style.configure("TLabel",        background=CREME_BG,  font=FONT_BODY)
        style.configure("TButton",       background=CREME_ALT, font=FONT_BODY)
        style.configure("TLabelframe",   background=CREME_BG,  font=FONT_SUBHEADER)
        style.configure("TLabelframe.Label", background=CREME_BG, font=FONT_SUBHEADER,
                        foreground=DARK_SLATE)

        # Treeview — increased row height, creme background
        style.configure("Treeview",
                        background=CARD_WHITE,
                        fieldbackground=CARD_WHITE,
                        rowheight=24,
                        font=FONT_BODY)
        style.configure("Treeview.Heading",
                        background=CREME_ALT,
                        font=FONT_SUBHEADER,
                        foreground=DARK_SLATE)
        style.map("Treeview",
                  background=[("selected", DARK_SLATE)],
                  foreground=[("selected", "white")])

        # Separator
        style.configure("TSeparator", background=BORDER_CLR)

    # ------------------------------------------------------------------
    # UI Construction
    # ------------------------------------------------------------------

    def _build_ui(self):
        # ---- Top bar: title strip ----
        title_bar = tk.Frame(self.root, bg=DARK_SLATE, pady=10)
        title_bar.pack(fill="x")
        tk.Label(title_bar, text="LX Report Generator",
                 font=FONT_HEADER, bg=DARK_SLATE, fg="white").pack(side="left", padx=18)
        tk.Label(title_bar, text="O_REP  v2.0",
                 font=FONT_BODY, bg=DARK_SLATE, fg=ACCENT_GOLD).pack(side="right", padx=18)

        # ---- Config row (PCB ID / Globals / Paths) ----
        top_frame = tk.Frame(self.root, bg=CREME_BG)
        top_frame.pack(fill="x", padx=14, pady=(10, 0))

        self._build_pcb_section(top_frame)
        self._build_globals_section(top_frame)
        self._build_paths_section(top_frame)

        # ---- Thin separator ----
        tk.Frame(self.root, bg=BORDER_CLR, height=1).pack(fill="x", padx=14, pady=8)

        # ---- Warning banner (shown only when CSV missing) ----
        self._warning_bar = tk.Frame(self.root, bg=WARN_BG, pady=6)
        # packed conditionally in load_csv()

        # ---- Mid: CSV / Treeview ----
        mid_lf = tk.LabelFrame(self.root, text="Source of Truth  ─  Record.csv",
                               bg=CREME_BG, font=FONT_SUBHEADER,
                               fg=DARK_SLATE, bd=1, relief="solid")
        mid_lf.pack(fill="both", expand=True, padx=14, pady=(0, 6))

        # Button row above treeview
        btn_row = tk.Frame(mid_lf, bg=CREME_BG)
        btn_row.pack(fill="x", padx=6, pady=6)

        self._make_small_btn(btn_row, "⟳  Reload CSV",       self.load_csv,  CREME_ALT).pack(side="left", padx=(0, 6))
        self._make_small_btn(btn_row, "💾  Save CSV Changes", self.save_csv,  CREME_ALT).pack(side="left")

        # Path label
        self.path_label = tk.Label(btn_row, text="", font=FONT_MONO,
                                    bg=CREME_BG, fg="#888888")
        self.path_label.pack(side="right", padx=6)

        # Treeview + scrollbars
        tv_frame = tk.Frame(mid_lf, bg=CARD_WHITE, bd=1, relief="solid")
        tv_frame.pack(fill="both", expand=True, padx=6, pady=(0, 6))

        columns = ("Index", "LX", "Type", "M", "D_File_Name")
        self.tree = ttk.Treeview(tv_frame, columns=columns, show="headings",
                                  selectmode="browse")
        col_widths = {"Index": 60, "LX": 50, "Type": 80, "M": 80, "D_File_Name": 200}
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=col_widths.get(col, 100), anchor="center")

        vsb = ttk.Scrollbar(tv_frame, orient="vertical",   command=self.tree.yview)
        hsb = ttk.Scrollbar(tv_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tv_frame.rowconfigure(0, weight=1)
        tv_frame.columnconfigure(0, weight=1)

        self.tree.bind("<Double-1>", self.toggle_lx)
        # Alternating row tags
        self.tree.tag_configure("odd",  background=CREME_BG)
        self.tree.tag_configure("even", background=CARD_WHITE)

        # ---- Bottom: action buttons + log ----
        bot_frame = tk.Frame(self.root, bg=CREME_ALT, pady=8)
        bot_frame.pack(fill="x", padx=14, pady=(0, 10))

        # Action buttons side by side
        action_row = tk.Frame(bot_frame, bg=CREME_ALT)
        action_row.pack(fill="x", padx=8, pady=(0, 6))

        self.btn_validate = tk.Button(
            action_row,
            text="✔  VALIDATE MATRIX",
            command=self.validate_matrix,
            bg=ACCENT_GOLD, fg="white",
            font=("Segoe UI", 11, "bold"),
            padx=18, pady=8, relief="flat",
            activebackground="#8B6300", activeforeground="white"
        )
        self.btn_validate.pack(side="left", padx=(0, 10))

        self.btn_start = tk.Button(
            action_row,
            text="▶  START REPORT GENERATION",
            command=self.run_generation,
            bg=DARK_SLATE, fg="white",
            font=("Segoe UI", 11, "bold"),
            padx=18, pady=8, relief="flat",
            activebackground="#1A252F", activeforeground="white"
        )
        self.btn_start.pack(side="left")

        # Log area
        log_frame = tk.Frame(bot_frame, bg=CREME_ALT)
        log_frame.pack(fill="x", padx=8)
        tk.Label(log_frame, text="System Log", font=FONT_SUBHEADER,
                 bg=CREME_ALT, fg=DARK_SLATE).pack(anchor="w", pady=(0, 2))
        self.log_area = tk.Text(log_frame, height=7,
                                 bg=CARD_WHITE, fg="#333333",
                                 font=FONT_MONO,
                                 bd=1, relief="solid",
                                 state="disabled")
        log_vsb = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_area.yview)
        self.log_area.configure(yscrollcommand=log_vsb.set)
        self.log_area.pack(side="left", fill="x", expand=True)
        log_vsb.pack(side="right", fill="y")

    def _build_pcb_section(self, parent):
        card = self._card(parent, "PCB_ID Configuration")
        card.pack(side="left", fill="both", expand=True, padx=(0, 6))

        for i, (lbl, attr) in enumerate([
            ("Project:",  "entry_proj"),
            ("Version:",  "entry_ver"),
            ("Board No:", "entry_brd"),
        ]):
            tk.Label(card, text=lbl, font=FONT_BODY, bg=CARD_WHITE, anchor="e"
                     ).grid(row=i, column=0, padx=(10, 4), pady=4, sticky="e")
            e = tk.Entry(card, width=16, font=FONT_BODY,
                          bg=CREME_BG, relief="solid", bd=1)
            e.grid(row=i, column=1, padx=(0, 10), pady=4)
            setattr(self, attr, e)

    def _build_globals_section(self, parent):
        card = self._card(parent, "Global Variables")
        card.pack(side="left", fill="both", expand=True, padx=(0, 6))

        for i, (lbl, attr) in enumerate([
            ("Tester:",      "entry_tester"),
            ("Reviewed By:", "entry_reviewer"),
        ]):
            tk.Label(card, text=lbl, font=FONT_BODY, bg=CARD_WHITE, anchor="e"
                     ).grid(row=i, column=0, padx=(10, 4), pady=4, sticky="e")
            e = tk.Entry(card, width=16, font=FONT_BODY,
                          bg=CREME_BG, relief="solid", bd=1)
            e.grid(row=i, column=1, padx=(0, 10), pady=4)
            setattr(self, attr, e)

    def _build_paths_section(self, parent):
        card = self._card(parent, "Path Mapping (P0 – P9)")
        card.pack(side="left", fill="both", expand=True)

        self.path_entries = []
        for i in range(2):  # Compact display: P0, P1
            tk.Label(card, text=f"P{i}:", font=FONT_BODY, bg=CARD_WHITE
                     ).grid(row=i, column=0, padx=(10, 4), pady=4)
            e = tk.Entry(card, width=18, font=FONT_BODY,
                          bg=CREME_BG, relief="solid", bd=1)
            e.grid(row=i, column=1, padx=(0, 2), pady=4)
            btn = tk.Button(card, text="…", font=FONT_BODY,
                             bg=CREME_ALT, padx=4,
                             command=lambda ent=e: self.browse_path(ent))
            btn.grid(row=i, column=2, padx=(2, 10), pady=4)
            self.path_entries.append(e)

    def _card(self, parent, title="") -> tk.Frame:
        """Create a titled card frame (white bg, grey border)."""
        outer = tk.LabelFrame(parent, text=title,
                               bg=CARD_WHITE, fg=DARK_SLATE,
                               font=FONT_SUBHEADER,
                               bd=1, relief="solid")
        return outer

    @staticmethod
    def _make_small_btn(parent, text, cmd, bg) -> tk.Button:
        return tk.Button(parent, text=text, command=cmd,
                          font=FONT_BODY, bg=bg,
                          fg=DARK_SLATE, padx=10, pady=4,
                          relief="groove",
                          activebackground=DARK_SLATE, activeforeground="white")

    # ------------------------------------------------------------------
    # Warning Banner
    # ------------------------------------------------------------------

    def _show_warning_banner(self, msg: str):
        """Display a fixed amber banner above the treeview when CSV is missing."""
        # Clear any old content
        for w in self._warning_bar.winfo_children():
            w.destroy()

        tk.Label(self._warning_bar, text=f"⚠  {msg}",
                 font=FONT_SUBHEADER, bg=WARN_BG, fg=WARN_FG).pack(side="left", padx=14)

        tk.Button(self._warning_bar, text="Browse File…",
                  command=self.browse_record_csv,
                  bg=WARN_FG, fg="white",
                  font=FONT_BODY, padx=8, pady=2).pack(side="left", padx=6)

        self._warning_bar.pack(fill="x", padx=14, pady=4, before=self.root.winfo_children()[4])

    def _hide_warning_banner(self):
        self._warning_bar.pack_forget()

    def browse_record_csv(self):
        path = filedialog.askopenfilename(
            title="Select Record.csv",
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
        )
        if path:
            self.record_path = path
            self._hide_warning_banner()
            self.load_csv()

    # ------------------------------------------------------------------
    # CSV Logic
    # ------------------------------------------------------------------

    def browse_path(self, entry_widget):
        path = filedialog.askdirectory()
        if path:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, path)

    def log(self, msg: str):
        self.log_area.config(state="normal")
        self.log_area.insert(tk.END, msg + "\n")
        self.log_area.see(tk.END)
        self.log_area.config(state="disabled")
        self.root.update_idletasks()

    def load_csv(self):
        """Load Record.csv; shows warning banner if file cannot be located."""
        for item in self.tree.get_children():
            self.tree.delete(item)

        if not self.record_path or not os.path.isfile(self.record_path):
            self._show_warning_banner(
                "Record.csv not found — browsed automatically in CWD, script dir, and parent root."
            )
            self.path_label.config(text="No file loaded")
            self.log("[WARN] Record.csv could not be auto-detected. Use 'Browse File' to locate it.")
            return

        self._hide_warning_banner()
        self.path_label.config(text=self.record_path)

        with open(self.record_path, "r", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            self.csv_fieldnames = reader.fieldnames or []
            self.csv_data = list(reader)

        for i, row in enumerate(self.csv_data):
            tag = "odd" if i % 2 else "even"
            self.tree.insert("", "end", iid=i, tags=(tag,), values=(
                row.get("Index",      ""),
                row.get("LX",         ""),
                row.get("Type",       ""),
                row.get("M",          ""),
                row.get("D_File_Name",""),
            ))

        self.log(f"[OK] Loaded {len(self.csv_data)} row(s) from: {self.record_path}")

    def toggle_lx(self, event):
        item = self.tree.selection()
        if not item:
            return
        item = item[0]
        values = list(self.tree.item(item, "values"))
        values[1] = "" if values[1].upper() == "X" else "X"
        self.tree.item(item, values=values)
        self.csv_data[int(item)]["LX"] = values[1]

    def save_csv(self):
        if not self.csv_data or not self.record_path:
            self.log("[WARN] Nothing to save — no CSV loaded.")
            return
        with open(self.record_path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=self.csv_fieldnames)
            writer.writeheader()
            writer.writerows(self.csv_data)
        self.log(f"[OK] Saved changes to {self.record_path}")

    # ------------------------------------------------------------------
    # Validate Matrix
    # ------------------------------------------------------------------

    def validate_matrix(self):
        """
        Integrity scan of Record.csv against Record_Contract.csv.

        Checks:
          1. Missing required parameter values (Type, D_File_Name).
          2. Broken mirror index pointers (M column Mx references valid Index).
          3. Valid cell coordinate formatting (e.g. A1, AFG3) for cell range fields.
        """
        if not self.csv_data:
            messagebox.showwarning("Validate Matrix",
                                   "No CSV data loaded. Please load Record.csv first.")
            return

        self.log("")
        self.log("═" * 60)
        self.log("  VALIDATE MATRIX — Integrity Scan")
        self.log("═" * 60)

        # Build set of all valid Index values for mirror-pointer checks
        valid_indices = set()
        for row in self.csv_data:
            idx = row.get("Index", "").strip()
            if idx:
                valid_indices.add(idx)

        # Build set of valid Type aliases from Record_Contract.csv (if available)
        valid_types: set | None = None
        if self.contract_path and os.path.isfile(self.contract_path):
            valid_types = set()
            with open(self.contract_path, "r", encoding="utf-8-sig") as f:
                for crow in csv.DictReader(f):
                    alias = crow.get("Alias", "").strip()
                    if alias:
                        valid_types.add(alias)
            self.log(f"  Contract aliases loaded: {sorted(valid_types)}")
        else:
            self.log("  [WARN] Record_Contract.csv not found — Type alias validation skipped.")

        self.log("─" * 60)

        cell_fields = [
            "S_Start_Cell_Range", "S_End_Cell_Range",
            "D_Start_Cell_Range", "D_End_Cell_Range",
        ]
        required_fields = ["Type", "D_File_Name"]

        total = len(self.csv_data)
        issues = 0

        for i, row in enumerate(self.csv_data):
            row_num   = i + 1
            row_index = row.get("Index", "?")
            row_type  = row.get("Type",  "").strip()
            row_issues = []

            # --- Check 1: Required fields ---
            for field in required_fields:
                val = row.get(field, "").strip()
                if not val:
                    row_issues.append(f"MISSING required field '{field}'")

            # --- Check 2: Type alias against contract ---
            if valid_types is not None and row_type and row_type not in valid_types:
                row_issues.append(f"UNKNOWN Type alias '{row_type}' (not in Record_Contract.csv)")

            # --- Check 3: Mirror index pointer ---
            mirror = row.get("M", "").strip()
            if mirror:
                # Valid values: digit-only index OR mirror references like M1, M2
                # Mirror references must point to an existing Index value
                if re.match(r"^M\d+$", mirror, re.IGNORECASE):
                    pointed = mirror[1:]  # e.g. "M2" → "2"
                    if pointed not in valid_indices:
                        row_issues.append(
                            f"BROKEN mirror pointer '{mirror}' → Index '{pointed}' does not exist"
                        )

            # --- Check 4: Cell coordinate format ---
            for cf in cell_fields:
                val = row.get(cf, "").strip()
                if val and not _CELL_COORD_RE.match(val):
                    row_issues.append(
                        f"INVALID cell coordinate '{val}' in field '{cf}' "
                        f"(expected pattern like A1, AFG3)"
                    )

            # --- Report row result ---
            if row_issues:
                issues += len(row_issues)
                self.log(f"  [FAIL] Row {row_num} (Index={row_index}, Type={row_type})")
                for iss in row_issues:
                    self.log(f"         ✗ {iss}")
            else:
                self.log(f"  [PASS] Row {row_num} (Index={row_index}, Type={row_type})")

        self.log("─" * 60)
        if issues == 0:
            self.log(f"  ✔  All {total} row(s) passed validation — Matrix Integrity: CLEAN")
        else:
            self.log(f"  ✗  {issues} issue(s) found across {total} row(s) — Matrix Integrity: REVIEW NEEDED")
        self.log("═" * 60)
        self.log("")

    # ------------------------------------------------------------------
    # Report Generation
    # ------------------------------------------------------------------

    def run_generation(self):
        self.save_csv()
        self.btn_start.config(state=tk.DISABLED)
        self.btn_validate.config(state=tk.DISABLED)
        self.log("─" * 60)
        self.log("  Starting Report Generation…")
        self.log("─" * 60)

        pcb_id = f"{self.entry_proj.get()}_{self.entry_ver.get()}_{self.entry_brd.get()}"
        LX_CommCentre.update_pcb_id(pcb_id)
        LX_CommCentre.update_credentials(self.entry_tester.get(), self.entry_reviewer.get())

        for i, entry in enumerate(self.path_entries):
            LX_CommCentre.set_path(i, entry.get())

        def execute():
            orch = Orchestrator(log_callback=self.log)
            orch.run_sequence(self.record_path)
            self.root.after(0, lambda: self.btn_start.config(state=tk.NORMAL))
            self.root.after(0, lambda: self.btn_validate.config(state=tk.NORMAL))
            self.root.after(0, lambda: messagebox.showinfo(
                "Complete", "Report Generation Finished."))

        threading.Thread(target=execute, daemon=True).start()


def main():
    root = tk.Tk()
    O_REP_GUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
