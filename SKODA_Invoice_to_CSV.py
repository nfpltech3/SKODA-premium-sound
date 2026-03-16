"""
SKODA Invoice to CSV Converter
Extracts line-item data from Premium Sound Solutions invoice PDFs
and exports to a CSV (.csv) file.

Author : Nagarkot Forwarders
Version: 1.0
"""

import pdfplumber
import pandas as pd
import re
import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# ──────────────────────────────────────────────────────────────
# TRY to load Pillow for logo support; degrade gracefully
# ──────────────────────────────────────────────────────────────
try:
    from PIL import Image, ImageTk
    HAS_PIL = True
except ImportError:
    HAS_PIL = False


# ═══════════════════════════════════════════════════════════════
#  EXTRACTION ENGINE
# ═══════════════════════════════════════════════════════════════

class InvoiceExtractor:
    """
    Parses Premium Sound Solutions (PSS) structured invoice PDFs.

    Expected PDF layout (per item block):
        <Pos> <ArticleNo> <Qty> <QtyUnit> <UnitPrice> <PriceUnit> <Amount> <Currency>
        <Description>                          Net weight: <NetWt> <NetWtUnit>
        YOUR REFERENCE: <Ref>
        COMMODITY CODE: <HSCode>
        Country of origin <Country>
    
    Global fields (once per invoice):
        Doc. Nr.  / Doc. Date   → Invoice No / Invoice Date
        Description: <Name>     → Name (usually on last page)
    """

    # ── regex patterns ────────────────────────────────────────
    # Item line:  10 11067501 3,960 PC 271.08 USD/100 PC 10,734.77 USD
    RE_ITEM = re.compile(
        r"^\s*(\d{1,4})\s+"           # Pos (1-4 digits)
        r"(\d{5,12})\s+"              # Article No
        r"([\d,]+)\s+"                # Qty  (may contain commas)
        r"(PC|PCS|PAC|EA|SET)\s+"     # Qty Unit
        r"([\d,.]+)\s+"               # Unit Price
        r"(\S+(?:\s+\S+)?)\s+"        # Price per piece (e.g. USD/100 PC)
        r"([\d,.]+)\s+"               # Amount
        r"([A-Z]{3})"                 # Currency
    )
    RE_DOC_NR      = re.compile(r"^(00\d{8,})\s+(\d{2}\.\d{2}\.\d{4})")
    RE_NET_WT      = re.compile(r"Net\s+weight:\s*([\d,.]+)\s*(\w+)", re.IGNORECASE)
    RE_YOUR_REF    = re.compile(r"YOUR\s+REFERENCE:\s*(.+)", re.IGNORECASE)
    RE_COMMODITY   = re.compile(r"COMMODITY\s+CODE:\s*(\S+)", re.IGNORECASE)
    RE_ORIGIN      = re.compile(r"Country\s+of\s+origin\s+(.+)", re.IGNORECASE)
    RE_DESCRIPTION = re.compile(r"^Description:\s*(.+)", re.IGNORECASE)

    DEFAULT_VALUE  = "(AUTOMOTIVE PARTS FOR CAPTIVE CONSUMPTION)"

    # ── public API ────────────────────────────────────────────

    def extract(self, pdf_path: str) -> list[dict]:
        """Return a list of dicts, one per line-item, from the given PDF."""

        rows: list[dict] = []
        inv_no = inv_date = name = None
        current: dict | None = None

        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if not text:
                    continue

                lines = text.split("\n")
                for line in lines:
                    stripped = line.strip()
                    if not stripped:
                        continue

                    # ── Doc Nr / Date (once) ──────────────
                    if inv_no is None:
                        m = self.RE_DOC_NR.search(stripped)
                        if m:
                            inv_no   = m.group(1)
                            inv_date = m.group(2)
                            continue

                    # ── Item line ────────────────────────
                    m = self.RE_ITEM.match(stripped)
                    if m:
                        # save previous item if any
                        if current is not None:
                            rows.append(current)
                        current = {
                            "Invoice No":       inv_no  or "",
                            "Invoice Date":     inv_date or "",
                            "Your Reference":   "",
                            "Name":             "",
                            "Article No":       m.group(2),
                            "Description":      "",
                            "Default":          self.DEFAULT_VALUE,
                            "HS Code":          "",
                            "Country of Origin":"",
                            "Qty":              m.group(3).replace(",", ""),
                            "Qty Unit":         m.group(4),
                            "Unit Price":       m.group(5).replace(",", ""),
                            "Price per piece":  m.group(6),
                            "Net wt":           "",
                            "Net wt unit":      "",
                            "Amount":           m.group(7).replace(",", ""),
                            "Currency":         m.group(8),
                        }
                        continue

                    # ── Subsequent detail lines (attach to current item) ─
                    if current is not None:
                        # Description + Net weight on the same line
                        m_nw = self.RE_NET_WT.search(stripped)
                        if m_nw:
                            current["Net wt"]      = m_nw.group(1).replace(",", "")
                            current["Net wt unit"] = m_nw.group(2)
                            # Text before "Net weight:" is the description
                            desc_part = stripped[:m_nw.start()].strip()
                            if desc_part and not current["Description"]:
                                # Remove leading dot artifacts from PDF
                                desc_part = desc_part.lstrip(". ").strip()
                                if desc_part:
                                    current["Description"] = desc_part
                            continue

                        # YOUR REFERENCE
                        m_ref = self.RE_YOUR_REF.match(stripped)
                        if m_ref:
                            # Strip ALL whitespace from the reference
                            current["Your Reference"] = re.sub(r"\s+", "", m_ref.group(1))
                            continue

                        # COMMODITY CODE
                        m_hs = self.RE_COMMODITY.match(stripped)
                        if m_hs:
                            current["HS Code"] = m_hs.group(1).strip()
                            continue

                        # Country of origin
                        m_co = self.RE_ORIGIN.match(stripped)
                        if m_co:
                            current["Country of Origin"] = m_co.group(1).strip()
                            continue

                    # ── Global: Description (Name) ─────────
                    m_name = self.RE_DESCRIPTION.match(stripped)
                    if m_name:
                        name = m_name.group(1).strip()

        # flush last item
        if current is not None:
            rows.append(current)

        # back-fill Name onto all rows
        if name:
            for r in rows:
                r["Name"] = name

        return rows


# ═══════════════════════════════════════════════════════════════
#  GUI APPLICATION
# ═══════════════════════════════════════════════════════════════

class App:
    """Tkinter front-end following the Nagarkot brand standard."""

    # ── Brand colours ─────────────────────────────────────────
    BG          = "#F4F6F8"
    PANEL_BG    = "#FFFFFF"
    FG          = "#1E1E1E"
    BRAND_BLUE  = "#1F3F6E"
    HOVER_BLUE  = "#2A528F"
    ACCENT_RED  = "#D8232A"
    MUTED_GRAY  = "#6B7280"
    BORDER_GRAY = "#E5E7EB"

    COLUMN_SCHEMA = [
        "Invoice No", "Invoice Date", "Your Reference", "Name",
        "Article No", "Description", "Default", "HS Code",
        "Country of Origin", "Qty", "Qty Unit", "Unit Price",
        "Price per piece", "Net wt", "Net wt unit", "Amount", "Currency",
    ]

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Nagarkot Forwarders — SKODA Invoice to CSV")
        try:
            self.root.state('zoomed')
        except tk.TclError:
            self.root.attributes('-zoomed', True)
        self.root.configure(bg=self.BG)
        self.root.minsize(900, 600)

        self.extractor = InvoiceExtractor()
        self.file_paths: list[str] = []
        self.extracted_rows: list[dict] = []

        self._setup_styles()
        self._create_header()
        self._create_main()
        self._create_footer()

    # ── Theming ───────────────────────────────────────────────

    def _setup_styles(self):
        s = ttk.Style()
        s.theme_use("clam")

        s.configure("TFrame", background=self.BG)
        # Panel frames
        s.configure("Panel.TFrame", background=self.PANEL_BG)
        s.configure("TLabelframe", background=self.PANEL_BG, foreground=self.BRAND_BLUE, font=("Arial", 10, "bold"), bordercolor=self.BORDER_GRAY)
        s.configure("TLabelframe.Label", background=self.PANEL_BG, foreground=self.BRAND_BLUE, font=("Arial", 10, "bold"))
        
        s.configure("TLabel", background=self.BG, foreground=self.FG, font=("Arial", 10))
        s.configure("Panel.TLabel", background=self.PANEL_BG, foreground=self.FG, font=("Arial", 10))
        s.configure("TButton", font=("Arial", 10, "bold"))

        s.configure("Header.TLabel", font=("Helvetica", 18, "bold"), foreground=self.BRAND_BLUE, background=self.BG)
        s.configure("SubHeader.TLabel", font=("Helvetica", 10), foreground=self.MUTED_GRAY, background=self.BG)

        # Primary button
        s.configure("Primary.TButton", background=self.BRAND_BLUE, foreground="white", borderwidth=0)
        s.map("Primary.TButton", background=[("active", self.HOVER_BLUE)])
        
        # Secondary button
        s.configure("Secondary.TButton", background=self.PANEL_BG, foreground=self.BRAND_BLUE, bordercolor=self.BORDER_GRAY)
        s.map("Secondary.TButton", background=[("active", self.BORDER_GRAY)])
        
        # Status-specific labels
        s.configure("Success.TLabel", background=self.BG, foreground=self.BRAND_BLUE, font=("Arial", 10, "bold"))
        s.configure("Error.TLabel", background=self.BG, foreground=self.ACCENT_RED, font=("Arial", 10, "bold"))

        # Treeview formatting
        s.configure("Treeview", background=self.PANEL_BG, fieldbackground=self.PANEL_BG, foreground=self.FG)
        s.configure("Treeview.Heading", background=self.BORDER_GRAY, foreground=self.BRAND_BLUE, font=("Arial", 9, "bold"))

    # ── Header (Logo left · Title centre) ──────────────────────

    def _create_header(self):
        hdr = ttk.Frame(self.root, style="TFrame")
        hdr.pack(fill="x", pady=20)
        
        title_ctr = ttk.Frame(hdr, style="TFrame")
        title_ctr.pack(side="top")
        
        ttk.Label(title_ctr, text="INVOICE TO CSV CONVERTER", style="Header.TLabel").pack(anchor="center")
        ttk.Label(title_ctr, text="Premium Sound Solutions — SKODA", style="SubHeader.TLabel").pack(anchor="center")

        # Logo on the left
        logo_placed = False
        logo_lbl = None
        if HAS_PIL:
            logo_candidates = [
                resource_path("Nagarkot Logo.png"),
                os.path.join(os.path.dirname(__file__), "Nagarkot Logo.png"),
                r"c:\projects\Skoda 1702\Pre alert documents_Plant Pune -1\Skoda Auto AS - 1\Nagarkot Logo.png",
            ]
            for lp in logo_candidates:
                if os.path.isfile(lp):
                    try:
                        img = Image.open(lp)
                        h_pct = 20 / float(img.size[1])
                        w_new = int(float(img.size[0]) * h_pct)
                        img = img.resize((w_new, 20), Image.Resampling.LANCZOS)
                        self._logo_img = ImageTk.PhotoImage(img)
                        logo_lbl = ttk.Label(hdr, image=self._logo_img, style="TLabel")
                        logo_placed = True
                        break
                    except Exception as e:
                        pass
        if not logo_placed:
            logo_lbl = ttk.Label(hdr, text="NAGARKOT", font=("Arial", 12, "bold"), foreground=self.BRAND_BLUE, style="TLabel")
        
        logo_lbl.place(relx=0.0, x=20, rely=0.5, anchor="w")

    # ── Main content area ─────────────────────────────────────

    def _create_main(self):
        main = ttk.Frame(self.root, padding="20", style="TFrame")
        main.pack(fill="both", expand=True)

        # ╌╌ Input section ╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌
        inp = ttk.LabelFrame(main, text="Select Invoice PDFs", padding="12", style="TLabelframe")
        inp.pack(fill="x", pady=(0, 20))

        # Toolbar row
        toolbar = ttk.Frame(inp, style="Panel.TFrame")
        toolbar.pack(fill="x", pady=(0, 5))

        ttk.Button(toolbar, text="📂  Select Files",
                   command=self._browse_pdfs, style="Secondary.TButton").pack(side="left", padx=(0, 10))
        ttk.Button(toolbar, text="❌  Clear List",
                   command=self._clear_files, style="Secondary.TButton").pack(side="left")

        # Listbox with scrollbar
        list_frame = ttk.Frame(inp, style="Panel.TFrame")
        list_frame.pack(fill="x", pady=5)

        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side="right", fill="y")

        self.file_listbox = tk.Listbox(
            list_frame, height=4, yscrollcommand=scrollbar.set,
            selectmode=tk.EXTENDED, bd=1, relief="solid",
            font=("Consolas", 10), bg=self.PANEL_BG, fg=self.FG, highlightthickness=1,
            highlightcolor=self.BRAND_BLUE, highlightbackground=self.BORDER_GRAY)
        self.file_listbox.pack(side="left", fill="x", expand=True)
        scrollbar.config(command=self.file_listbox.yview)

        self.file_count_var = tk.StringVar(value="0 file(s) selected")
        ttk.Label(inp, textvariable=self.file_count_var,
                  font=("Arial", 9), style="Panel.TLabel", foreground=self.MUTED_GRAY).pack(anchor="w")

        # ╌╌ Output folder ╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌
        out_frame = ttk.LabelFrame(main, text="Output Folder", padding="12", style="TLabelframe")
        out_frame.pack(fill="x", pady=(0, 20))

        row_out = ttk.Frame(out_frame, style="Panel.TFrame")
        row_out.pack(fill="x")

        ttk.Button(row_out, text="📁  Change Folder",
                   command=self._browse_output, style="Secondary.TButton").pack(side="left", padx=(0, 10))

        self.output_dir_var = tk.StringVar(value="(same as invoice folder)")
        ttk.Label(row_out, textvariable=self.output_dir_var,
                  font=("Consolas", 10), style="Panel.TLabel", foreground=self.FG).pack(side="left", fill="x", expand=True)

        # ╌╌ Preview table ╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌╌
        pf = ttk.LabelFrame(main, text="Data Preview", padding="12", style="TLabelframe")
        pf.pack(fill="both", expand=True)

        # Horizontal + vertical scrollbars
        xsb = ttk.Scrollbar(pf, orient="horizontal")
        ysb = ttk.Scrollbar(pf, orient="vertical")

        cols = self.COLUMN_SCHEMA
        self.tree = ttk.Treeview(pf, columns=cols, show="headings",
                                 xscrollcommand=xsb.set, yscrollcommand=ysb.set, style="Treeview")
        xsb.config(command=self.tree.xview)
        ysb.config(command=self.tree.yview)

        # Column widths
        col_widths = {
            "Invoice No": 100, "Invoice Date": 85, "Your Reference": 110,
            "Name": 100, "Article No": 85, "Description": 240,
            "Default": 140, "HS Code": 100, "Country of Origin": 100,
            "Qty": 60, "Qty Unit": 55, "Unit Price": 75,
            "Price per piece": 100, "Net wt": 75, "Net wt unit": 60,
            "Amount": 85, "Currency": 55,
        }
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=col_widths.get(c, 80), minwidth=50)

        self.tree.pack(side="left", fill="both", expand=True)
        ysb.pack(side="right", fill="y")
        xsb.pack(side="bottom", fill="x")

    # ── Footer ────────────────────────────────────────────────

    def _create_footer(self):
        b_frame = tk.Frame(self.root, height=1, bg=self.BORDER_GRAY)
        b_frame.pack(side="bottom", fill="x")

        ftr = ttk.Frame(self.root, padding="15 10", style="TFrame")
        ftr.pack(side="bottom", fill="x")

        ttk.Label(ftr, text="Nagarkot Forwarders Pvt. Ltd. ©",
                  font=("Arial", 9), style="TLabel", foreground=self.MUTED_GRAY).pack(side="left")

        self.run_btn = ttk.Button(ftr, text="⚙  Process & Export",
                                  command=self._process, style="Primary.TButton")
        self.run_btn.pack(side="right", padx=(10, 0))

        self.status_var = tk.StringVar(value="Ready")
        self.status_lbl = ttk.Label(ftr, textvariable=self.status_var, style="TLabel", foreground=self.FG)
        self.status_lbl.pack(side="right", padx=10)

    # ── Callbacks ─────────────────────────────────────────────

    def _browse_pdfs(self):
        filenames = filedialog.askopenfilenames(
            title="Select Invoice PDFs",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if filenames:
            for f in filenames:
                if f not in self.file_paths:
                    self.file_paths.append(f)
                    self.file_listbox.insert(tk.END, os.path.basename(f))
            self._update_count()
            # Auto-set output folder to the first file's directory (if not already changed)
            if self.output_dir_var.get() == "(same as invoice folder)":
                self.output_dir_var.set(os.path.dirname(self.file_paths[0]))

    def _clear_files(self):
        self.file_paths.clear()
        self.file_listbox.delete(0, tk.END)
        self.tree.delete(*self.tree.get_children())
        self._update_count()

    def _update_count(self):
        self.file_count_var.set(f"{len(self.file_paths)} file(s) selected")

    def _browse_output(self):
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder:
            self.output_dir_var.set(folder)

    def _process(self):
        """Extract → preview → export."""

        if not self.file_paths:
            messagebox.showwarning("No files", "Please select at least one PDF file.")
            return

        self.status_var.set("Processing…")
        self.status_lbl.configure(style="TLabel")
        self.root.update_idletasks()

        # clear previous preview
        self.tree.delete(*self.tree.get_children())
        self.extracted_rows.clear()

        try:
            all_rows = []
            processed = 0
            failed_files = []

            for pdf_path in self.file_paths:
                try:
                    if not os.path.isfile(pdf_path):
                        failed_files.append(os.path.basename(pdf_path))
                        continue
                    rows = self.extractor.extract(pdf_path)
                    if rows:
                        src = os.path.basename(pdf_path)
                        for r in rows:
                            r["Source File"] = src
                        all_rows.extend(rows)
                        processed += 1
                    else:
                        failed_files.append(os.path.basename(pdf_path))
                except Exception as e:
                    failed_files.append(os.path.basename(pdf_path))
                    print(f"Error processing {pdf_path}: {e}")

            if not all_rows:
                self.status_var.set("No data found")
                self.status_lbl.configure(style="Error.TLabel")
                messagebox.showinfo("Info", "No line-item data could be extracted from the selected PDFs.")
                return

            self.extracted_rows = all_rows

            # populate preview
            for r in all_rows:
                vals = [r.get(c, "") for c in self.COLUMN_SCHEMA]
                self.tree.insert("", "end", values=vals)

            # ── Export to CSV ─────────────────────────────────
            # Put Source File as last column
            export_cols = self.COLUMN_SCHEMA + ["Source File"]
            df = pd.DataFrame(all_rows, columns=export_cols)
            # Preserve leading zeros: wrap as ="value" so Excel treats as text
            df["Invoice No"] = df["Invoice No"].apply(lambda x: f'="{x}"')

            # Cast numeric columns for clean values
            numeric_cols = ["Article No", "HS Code", "Qty",
                            "Unit Price", "Net wt", "Amount"]
            for col in numeric_cols:
                df[col] = pd.to_numeric(df[col], errors="coerce")

            # Build output path using chosen output folder
            out_dir = self.output_dir_var.get()
            if out_dir == "(same as invoice folder)":
                out_dir = os.path.dirname(self.file_paths[0])
            ts       = datetime.now().strftime("%Y%m%d_%H%M%S")
            out_name = f"SKODA_Invoice_{len(self.file_paths)}files_{ts}.csv"
            out_path = os.path.join(out_dir, out_name)

            df.to_csv(out_path, index=False, encoding="utf-8-sig")

            self.status_var.set(f"Exported → {out_name}")
            self.status_lbl.configure(style="Success.TLabel")

            msg = f"✅ {len(all_rows)} item(s) from {processed} file(s) extracted and saved.\n\n"
            msg += f"File: {out_name}\nLocation: {out_dir}"
            if failed_files:
                msg += f"\n\n⚠ Failed/empty: {', '.join(failed_files)}"
            messagebox.showinfo("Success", msg)

        except Exception as exc:
            self.status_var.set("Error")
            self.status_lbl.configure(style="Error.TLabel")
            messagebox.showerror("Processing Error", f"An error occurred:\n\n{exc}")
            import traceback
            traceback.print_exc()


# ═══════════════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════════════

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
