"""
Skoda Export CSV Generator for Logisys
Parses Skoda commercial invoice PDFs and generates Logisys-compatible CSV.

Author : Nagarkot Forwarders Pvt Ltd
Version: 1.0
"""

import os
import re
import sys
import csv
import logging
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext
from typing import Optional

import pdfplumber
from PIL import Image, ImageTk

# ============================================================================
#  LOGGING
# ============================================================================
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)


# ============================================================================
#  RESOURCE PATH (PyInstaller support)
# ============================================================================
def resource_path(relative_path: str) -> str:
    """Get absolute path to resource, works for dev and for PyInstaller."""
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


# ============================================================================
#  LOGISYS TEMPLATE HEADER (114 columns — loaded from template CSV)
# ============================================================================

def _load_header_from_template() -> list[str]:
    """Read the fixed header row from ExportCSVTemplate.csv."""
    template_path = resource_path("ExportCSVTemplate.csv")
    if os.path.exists(template_path):
        with open(template_path, "r", encoding="utf-8") as f:
            reader = csv.reader(f)
            header = next(reader)
            return header
    # Fallback: hardcoded 114-column header
    return [
        "Inv No", "Date", "TOI", "Product Desc", "Quantity", "Unit", "Currency",
        "Rate", "RITC", "PMV", "PMV Currency", "Exim Scheme", "NFEI Category",
        "Manufacturer", "Manufacturer Addr", "Manufacturer CodeType",
        "Manufacturer Code", "Manufacturer Cntry", "Manufacturer State",
        "Manuf Postal Code", "SourceState", "TransitCntry", "End Use",
        "IGST Pay Status", "IGST Taxable Amt", "IGST Rate", "Reward Item",
        "DEEC Regn No", "DEEC Regn Date", "DEEC Item SNo(Part E)",
        "DEEC Export Qty", "DEEC Item SNo(Part C)", "DEEC Item Type",
        "DEEC Item Desc", "DEEC Qty", "DEEC Unit", "EPCG Regn No",
        "EPCG Regn Date", "EPCG Item SNo(Part E)", "EPCG Export Qty",
        "EPCG Item SNo(Part C)", "EPCG Item Type", "EPCG Item Desc",
        "EPCG Qty", "EPCG Unit", "DBK Sr. No", "DBK QTY", "DBK Under",
        "DFIA Regn No", "DFIA Regn Date", "DFIA Exp Item SNo",
        "DFIA Export Qty", "DFIA Imp Item SNo", "DFIA Item Type",
        "DFIA Item Desc", "DFIA Qty", "DFIA Unit", "DFRC Regn No",
        "DFRC Prod Group", "DFRC SION Sr No", "DFRC SION IONorm SNo",
        "DFRC Qty", "DFRC Unit", "DFRC Item Type", "DFRC Item Description",
        "DFRC Tech. Details", "Re-Exp BE No", "Re-Exp BE Date",
        "Re-Exp Inv SNo", "Re-Exp Item SNo", "Re-Exp Imp Port Code",
        "Re-Exp Manual BE", "Re-Exp BE Item Desc.", "Re-Exp Qty Imported",
        "Re-Exp Unit", "Re-Exp Assessable Value", "Re-Exp Tot Duty Paid",
        "Re-Exp Duty Paid On", "Re-Exp Qty Exported", "Re-Exp Unit",
        "Re-Exp Tech. Details", "Re-Exp Input Credit Availed",
        "Re-Exp Personal Use Item", "Re-Exp Othr Identifying Params",
        "Re-Exp Obligation No.", "Re-Exp DBK Amt. Claimed",
        "Re-Exp Item Un-Used", "Re-Exp Commissioner Permission",
        "Re-Exp Board No.", "Re-Exp Board Date", "Re-Exp MODVAT Availed",
        "Re-Exp MODVAT Reversed", "Product Material Code", "IGST Amount",
        "TP Exporter Name", " TP Exporter IE Code", "TP Exporter BrS_No.",
        "TP Exporter  Reg Type", "TP Exporter  Reg No", "TP Exporter Addr",
        "SQC Qty", "SQC Unit", "CompCessRate", "Origin District",
        "PTAFTA Code", "PerNoOfUnitsForItemRate", "RoDTEP Status ",
        "RoDTEP QTY", "RoDTEP Unit", "Medicinal Plant", "Formulation",
        "Surface Material in Contact", "GST Allowed", "LGD Code",
    ]

LOGISYS_HEADER: list[str] = _load_header_from_template()


# ============================================================================
#  PDF PARSER — Skoda Commercial Invoice
# ============================================================================

def _clean_number(val: str) -> str:
    """Remove thousand separators from European-formatted numbers: 7,393.00 → 7393.00"""
    if val is None:
        return ""
    return val.replace(",", "")


def _merge_description_lines(german: Optional[str], english: Optional[str]) -> str:
    """
    Combine German + English description into a single clean description.
    We prefer the English description for Logisys.
    """
    parts: list[str] = []
    if english:
        # Remove newlines from multi-line descriptions
        eng_clean = " ".join(english.replace("\n", " ").split())
        parts.append(eng_clean)
    if not parts and german:
        ger_clean = " ".join(german.replace("\n", " ").split())
        parts.append(ger_clean)
    return " ".join(parts).strip()


def _clean_drawing_no(raw: Optional[str]) -> str:
    """Clean part number / drawing number, joining all lines into a single string.

    Multi-line part numbers in the PDF look like:
        '6JM 419 091 A\\nVKH'  → '6JM 419 091 A VKH'
    The second line is a variant/color code that is part of the full part number.
    """
    if not raw:
        return ""
    # Join all lines with a space to capture variant/color codes
    parts = [line.strip() for line in raw.split("\n") if line.strip()]
    return " ".join(parts)


def parse_skoda_invoice(pdf_path: str) -> dict:
    """
    Parse a Skoda commercial invoice PDF and return structured data.

    Returns:
        dict with keys:
            - invoice_no: str
            - invoice_date: str (DD-MM-YYYY)
            - currency: str
            - items: list[dict] — one dict per line item
    """
    result = {
        "invoice_no": "",
        "invoice_date": "",
        "currency": "EUR",
        "items": [],
    }

    with pdfplumber.open(pdf_path) as pdf:
        # --- Extract header info from Page 1 ---
        page1_text = pdf.pages[0].extract_text() or ""

        # Invoice number — try multiple patterns
        # Pattern 1: from page header table "Rechnung\nInvoice\n100008431"
        inv_match = re.search(r"(?:Invoice|Rechnung)\n(\d{9,})", page1_text)
        if inv_match:
            result["invoice_no"] = inv_match.group(1)
        else:
            # Pattern 2: inline "Invoice ... <number>"
            inv_match2 = re.search(r"(?:Invoice|Rechnung)[\s\S]*?(\d{9,})", page1_text)
            if inv_match2:
                result["invoice_no"] = inv_match2.group(1)

        # Invoice date (format in PDF: DD-MM-YYYY or YYYY-MM-DD)
        date_match = re.search(r"(\d{2}-\d{2}-\d{4})\s*/\s*\d{4}-\d{2}-\d{2}", page1_text)
        if date_match:
            result["invoice_date"] = date_match.group(1)
        else:
            # Try alternate date pattern
            date_match2 = re.search(r"(\d{2}\.\d{2}\.\d{4})", page1_text)
            if date_match2:
                d = date_match2.group(1)
                result["invoice_date"] = d.replace(".", "-")

        # Currency
        curr_match = re.search(r"(EUR|USD|GBP|INR)", page1_text)
        if curr_match:
            result["currency"] = curr_match.group(1)

        # --- Extract item data from all pages ---
        for page_num, page in enumerate(pdf.pages):
            tables = page.extract_tables()
            if not tables:
                continue

            for table in tables:
                if not table or len(table) < 5:
                    continue

                # Check if this is the item table (has 15 columns)
                # The header row should mention "Pos-No" or "PosNr"
                header_text = str(table[0][0]) if table[0] and table[0][0] else ""
                if "Pos" not in header_text and "PosNr" not in header_text:
                    continue

                # Skip only the header row (row 0); data rows are identified by
                # their 6-digit position number in column 0.
                for row_idx in range(1, len(table)):
                    row = table[row_idx]

                    # Skip rows that don't have 15 columns
                    if not row or len(row) < 15:
                        continue

                    # A valid data row starts with a position number (6 digits)
                    pos_no = str(row[0]).strip() if row[0] else ""
                    if not re.match(r"^\d{6}$", pos_no):
                        continue

                    # Extract fields
                    drawing_no_raw = row[1] or ""
                    german_desc = row[2] or ""
                    coo = str(row[3]).strip() if row[3] else ""

                    # HS Code extraction — handle pdfplumber splitting
                    # across columns 4 and 5.
                    # Normal: col4="" col5="87089900"
                    # Split:  col4="8" col5="7089900"  → needs merge
                    col4_val = str(row[4]).strip() if row[4] else ""
                    col5_val = str(row[5]).strip() if row[5] else ""

                    if col4_val and col4_val.isdigit() and col5_val.isdigit():
                        # Concatenate split HS code
                        hs_code = col4_val + col5_val
                    else:
                        hs_code = col5_val

                    # Validate HS code should be 8 digits
                    if hs_code and not re.match(r"^\d{8}$", hs_code):
                        logger.warning(
                            "Item %s: HS code '%s' is not 8 digits — check PDF extraction.",
                            pos_no, hs_code,
                        )

                    dbk_code = str(row[6]).strip() if row[6] else ""
                    quantity = str(row[9]).strip() if row[9] else ""
                    unit_raw = str(row[10]).strip() if row[10] else ""
                    net_weight = str(row[12]).strip() if row[12] else ""
                    price_per_100 = str(row[13]).strip() if row[13] else ""
                    total_price = str(row[14]).strip() if row[14] else ""

                    # Get English description from the next row(s)
                    english_desc = ""
                    if row_idx + 1 < len(table):
                        next_row = table[row_idx + 1]
                        if next_row and len(next_row) >= 3:
                            # Position 0 should be None (continuation row)
                            if next_row[0] is None:
                                english_desc = str(next_row[2]) if next_row[2] else ""

                    # Also check row+2 for multi-line English descriptions
                    if row_idx + 2 < len(table):
                        next_row2 = table[row_idx + 2]
                        if next_row2 and len(next_row2) >= 3:
                            if next_row2[0] is None and next_row2[2]:
                                extra = str(next_row2[2]).strip()
                                if extra and not re.match(r"^\d{6}$", extra):
                                    if english_desc:
                                        english_desc += " " + extra
                                    else:
                                        english_desc = extra

                    # Build description
                    part_no = _clean_drawing_no(drawing_no_raw)
                    desc = _merge_description_lines(german_desc, english_desc)

                    # Build final product description with part number
                    if desc and part_no:
                        product_desc = f"Parts & Components of Passenger Car - Part no. {part_no}  {desc}"
                    elif desc:
                        product_desc = desc
                    elif part_no:
                        product_desc = f"Part no. {part_no}"
                    else:
                        product_desc = ""

                    # Calculate per-unit rate: Price/100 ÷ 100
                    try:
                        p100 = float(_clean_number(price_per_100))
                        rate = round(p100 / 100, 2)
                    except (ValueError, ZeroDivisionError):
                        rate = 0.0

                    # Map unit: PCE → PCS
                    unit_map = {"PCE": "PCS", "KG": "KGS", "M": "MTR"}
                    unit = unit_map.get(unit_raw, unit_raw)

                    item = {
                        "pos_no": pos_no,
                        "part_no": part_no,
                        "product_desc": product_desc,
                        "coo": coo,
                        "hs_code": hs_code,
                        "dbk_code": dbk_code,
                        "quantity": quantity,
                        "unit": unit,
                        "net_weight": _clean_number(net_weight),
                        "price_per_100": _clean_number(price_per_100),
                        "rate": rate,
                        "total_price": _clean_number(total_price),
                    }
                    result["items"].append(item)

    logger.info(
        "Parsed invoice %s: %d items found.",
        result["invoice_no"], len(result["items"])
    )
    return result


# ============================================================================
#  HS CODE → UNIT LOOKUP (from HSN Code Master)
# ============================================================================

# Default filename for the HS Code Master Excel
HS_MASTER_FILENAME = "HSN Code - Master_Data_Feb_2026.xlsx"


def load_hs_unit_map(excel_path: str) -> dict[str, str]:
    """
    Load HS Code Master Excel and build a mapping of 8-digit HS codes to UQC.

    The master file structure:
        Column A (Code)        — HS code (6 or 8 digits)
        Column B (Description) — Item description
        Column C (UQC)         — Unit of quantity (KGS / NOS)

    Only 8-digit codes with a non-empty UQC are included.
    Returns dict like {"87089400": "KGS", "85122010": "NOS", ...}
    """
    try:
        import openpyxl
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        ws = wb.active
        mapping: dict[str, str] = {}

        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or row[0] is None:
                continue
            code = str(row[0]).strip()
            uqc = str(row[2]).strip().upper() if (len(row) >= 3 and row[2]) else ""
            # Only include 8-digit HS codes with a valid UQC
            if len(code) == 8 and uqc:
                mapping[code] = uqc

        wb.close()
        logger.info("Loaded %d HS code → UQC mappings from %s", len(mapping), excel_path)
        return mapping
    except Exception as e:
        logger.error("Failed to load HS unit map: %s", e)
        return {}


# ============================================================================
#  LOGISYS ROW BUILDER
# ============================================================================

def build_logisys_rows(
    invoice_data: dict,
    toi: str,
    exim_scheme: str,
    end_use: str,
    igst_pay_status: str,
    igst_rate: str,
    dbk_under: str,
    origin_district: str,
    ptafta_code: str,
    hs_unit_map: dict[str, str],
) -> list[list[str]]:
    """
    Convert parsed invoice data into Logisys-format rows (107 columns each).
    """
    rows: list[list[str]] = []

    inv_no = invoice_data["invoice_no"]
    inv_date = invoice_data["invoice_date"]
    currency = invoice_data["currency"]

    for item in invoice_data["items"]:
        # Initialize empty row with 107 columns
        row = [""] * len(LOGISYS_HEADER)

        coo = item["coo"]
        hs_code = item["hs_code"]
        dbk_code = item["dbk_code"]
        quantity = item["quantity"]
        net_weight = item["net_weight"]

        # Determine Exim Scheme based on COO
        exim = exim_scheme if exim_scheme != "auto" else ("19" if coo == "IN" else "00")

        # Determine Reward Item based on COO
        reward_item = "YES" if coo == "IN" else "NO"

        # DBK Sr. No = DBK code + "B"
        dbk_sr_no = f"{dbk_code}B" if dbk_code and dbk_code != "0" else ""

        # Determine SQC/RoDTEP unit from HS code lookup
        sqc_unit = hs_unit_map.get(hs_code, "KGS")  # Default KGS
        if sqc_unit == "KGS":
            sqc_qty = net_weight
            rodtep_qty = net_weight
        elif sqc_unit in ("NOS", "PCS"):
            sqc_qty = quantity
            rodtep_qty = quantity
            sqc_unit = "NOS"
        else:
            sqc_qty = net_weight
            rodtep_qty = net_weight

        # Fill the row — column indices per ExportCSVTemplate.csv
        row[0] = inv_no                              # [0]  Inv No
        row[1] = inv_date                            # [1]  Date
        row[2] = toi                                 # [2]  TOI
        row[3] = item["product_desc"]                # [3]  Product Desc
        row[4] = quantity                            # [4]  Quantity
        row[5] = item["unit"]                        # [5]  Unit
        row[6] = currency                            # [6]  Currency
        row[7] = str(item["rate"])                   # [7]  Rate (per unit)
        row[8] = hs_code                             # [8]  RITC (HS Code)
        # [9-10]: PMV, PMV Currency — blank
        row[11] = exim                               # [11] Exim Scheme
        # [12]: NFEI Category — blank
        # [13-21]: Manufacturer info — blank
        row[22] = end_use                            # [22] End Use
        row[23] = igst_pay_status                    # [23] IGST Pay Status
        # [24]: IGST Taxable Amt — auto (blank for now)
        row[25] = igst_rate                          # [25] IGST Rate
        row[26] = reward_item                        # [26] Reward Item
        # [27-44]: DEEC, EPCG — blank
        row[45] = dbk_sr_no                          # [45] DBK Sr. No
        row[46] = quantity                           # [46] DBK QTY
        row[47] = dbk_under                          # [47] DBK Under
        # [48-99]: DFIA, DFRC, Re-Exp, IGST Amt, TP — blank
        row[100] = sqc_qty                           # [100] SQC Qty
        row[101] = sqc_unit                          # [101] SQC Unit
        # [102]: CompCessRate — blank
        row[103] = origin_district                   # [103] Origin District
        row[104] = ptafta_code                       # [104] PTAFTA Code
        # [105]: PerNoOfUnitsForItemRate — blank
        row[106] = "Y"                               # [106] RoDTEP Status
        row[107] = rodtep_qty                        # [107] RoDTEP QTY
        row[108] = sqc_unit                          # [108] RoDTEP Unit
        # [109-113]: Medicinal, Formulation, etc. — blank

        rows.append(row)

    return rows


def write_logisys_csv(rows: list[list[str]], output_path: str) -> None:
    """Write the Logisys CSV with the fixed header and data rows."""
    with open(output_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(LOGISYS_HEADER)
        writer.writerows(rows)
    logger.info("CSV written to %s with %d rows.", output_path, len(rows))


# ============================================================================
#  GUI APPLICATION
# ============================================================================

class SkodaExportApp:
    """Main GUI application for Skoda Export CSV Generator."""

    def __init__(self, master: tk.Tk) -> None:
        self.master = master
        self.master.title("Skoda Export CSV Generator — Logisys")
        self.master.geometry("1100x800")
        self.master.configure(bg="#ffffff")
        self.master.minsize(900, 700)

        # --- State ---
        self.pdf_path: Optional[str] = None
        self.hs_map_path: Optional[str] = None
        self.hs_unit_map: dict[str, str] = {}
        self.parsed_data: Optional[dict] = None

        # --- Tkinter Variables ---
        self.var_toi = tk.StringVar(value="C&F")
        self.var_exim_scheme = tk.StringVar(value="19")
        self.var_end_use = tk.StringVar(value="GNX100")
        self.var_igst_pay = tk.StringVar(value="Export Under Bond - Not Paid")
        self.var_igst_rate = tk.StringVar(value="")
        self.var_dbk_under = tk.StringVar(value="Actual")
        self.var_origin_district = tk.StringVar(value="490")
        self.var_ptafta = tk.StringVar(value="GSTP")

        # --- Build UI ---
        self._setup_styles()
        self._create_header()
        self._create_body()
        self._create_footer()

        # --- Auto-load HS Code Master on startup ---
        self._auto_load_hs_master()

    # ------------------------------------------------------------------ #
    #  STYLES
    # ------------------------------------------------------------------ #
    def _setup_styles(self) -> None:
        style = ttk.Style()
        style.theme_use("clam")

        bg = "#ffffff"
        primary = "#0056b3"

        style.configure(".", background=bg, font=("Segoe UI", 10))
        style.configure("TLabel", background=bg, foreground="#000000")
        style.configure("TFrame", background=bg)
        style.configure("TLabelframe", background=bg, bordercolor="#e0e0e0")
        style.configure(
            "TLabelframe.Label",
            background=bg, foreground=primary,
            font=("Segoe UI", 11, "bold"),
        )

        # Treeview
        style.configure(
            "Treeview",
            background="#ffffff", fieldbackground="#ffffff",
            foreground="#000000", rowheight=24, font=("Segoe UI", 9),
        )
        style.configure(
            "Treeview.Heading",
            background="#f0f4f8", foreground="#333",
            font=("Segoe UI", 9, "bold"),
        )
        style.map(
            "Treeview",
            background=[("selected", primary)],
            foreground=[("selected", "white")],
        )

        # Buttons
        style.configure(
            "Primary.TButton",
            font=("Segoe UI", 10, "bold"),
            background=primary, foreground="white", borderwidth=0,
        )
        style.map("Primary.TButton", background=[("active", "#004494")])

        style.configure(
            "Secondary.TButton",
            font=("Segoe UI", 10),
            background="#f0f0f0", foreground="#333", borderwidth=1,
        )
        style.map("Secondary.TButton", background=[("active", "#e0e0e0")])

        style.configure(
            "Success.TButton",
            font=("Segoe UI", 11, "bold"),
            background="#28a745", foreground="white", borderwidth=0,
        )
        style.map("Success.TButton", background=[("active", "#218838")])

        # Radiobuttons
        style.configure("TRadiobutton", background=bg, font=("Segoe UI", 9))

    # ------------------------------------------------------------------ #
    #  HEADER
    # ------------------------------------------------------------------ #
    def _create_header(self) -> None:
        header = tk.Frame(self.master, bg="#ffffff", height=80)
        header.pack(fill="x", padx=20, pady=10)
        header.grid_columnconfigure(0, weight=1)
        header.grid_columnconfigure(1, weight=2)
        header.grid_columnconfigure(2, weight=1)

        # Logo
        try:
            logo_path = resource_path("Nagarkot Logo.png")
            if os.path.exists(logo_path):
                pil_img = Image.open(logo_path)
                h = 22
                ratio = h / float(pil_img.size[1])
                w = int(float(pil_img.size[0]) * ratio)
                pil_img = pil_img.resize((w, h), Image.Resampling.LANCZOS)
                self._logo_img = ImageTk.PhotoImage(pil_img)
                tk.Label(
                    header, image=self._logo_img, bg="#ffffff", anchor="w",
                ).grid(row=0, column=0, sticky="w")
            else:
                tk.Label(
                    header, text="[Logo]", bg="#ffffff", fg="red",
                ).grid(row=0, column=0, sticky="w")
        except Exception:
            tk.Label(
                header, text="[Logo]", bg="#ffffff", fg="red",
            ).grid(row=0, column=0, sticky="w")

        # Title
        title_f = tk.Frame(header, bg="#ffffff")
        title_f.grid(row=0, column=1, sticky="ns")
        tk.Label(
            title_f, text="Skoda Export CSV Generator",
            font=("Helvetica", 18, "bold"), bg="#ffffff", fg="#000",
        ).pack(anchor="center")
        tk.Label(
            title_f, text="Logisys Import File from Skoda Invoice PDF",
            font=("Helvetica", 10), bg="#ffffff", fg="#777",
        ).pack(anchor="center")

        tk.Frame(header, bg="#ffffff").grid(row=0, column=2)

        ttk.Separator(self.master, orient="horizontal").pack(fill="x", padx=20)

    # ------------------------------------------------------------------ #
    #  BODY
    # ------------------------------------------------------------------ #
    def _create_body(self) -> None:
        body = tk.Frame(self.master, bg="#ffffff")
        body.pack(fill="both", expand=True, padx=20, pady=10)

        # ---- LEFT PANEL (scrollable) ----
        left_outer = tk.Frame(body, bg="#ffffff", width=380)
        left_outer.pack_propagate(False)
        left_outer.pack(side="left", fill="both")

        # Canvas + Scrollbar for scrollable left panel
        left_canvas = tk.Canvas(left_outer, bg="#ffffff", highlightthickness=0)
        left_scrollbar = ttk.Scrollbar(left_outer, orient="vertical", command=left_canvas.yview)
        left = tk.Frame(left_canvas, bg="#ffffff")

        left.bind(
            "<Configure>",
            lambda e: left_canvas.configure(scrollregion=left_canvas.bbox("all")),
        )
        left_canvas.create_window((0, 0), window=left, anchor="nw", width=360)
        left_canvas.configure(yscrollcommand=left_scrollbar.set)

        left_scrollbar.pack(side="right", fill="y")
        left_canvas.pack(side="left", fill="both", expand=True)

        # Enable mouse-wheel scrolling on the left panel
        def _on_mousewheel(event: tk.Event) -> None:
            left_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        left_canvas.bind_all("<MouseWheel>", _on_mousewheel)

        # Group 1: File Selection
        grp1 = ttk.LabelFrame(left, text="  1. Input Files  ", padding=10)
        grp1.pack(fill="x", pady=(0, 10))

        ttk.Button(
            grp1, text="📄  Select Invoice PDF",
            command=self._pick_pdf, style="Secondary.TButton",
        ).pack(fill="x", pady=3)
        self.lbl_pdf = ttk.Label(
            grp1, text="No PDF selected", foreground="#888",
        )
        self.lbl_pdf.pack(anchor="w", pady=2)

        ttk.Separator(grp1, orient="horizontal").pack(fill="x", pady=5)

        # HS Code Master status label
        self.lbl_hs = ttk.Label(
            grp1, text="HS Master: Loading...", foreground="#888",
        )
        self.lbl_hs.pack(anchor="w", pady=2)

        ttk.Button(
            grp1, text="🔄  Reload / Change HS Code Master",
            command=self._pick_hs_map, style="Secondary.TButton",
        ).pack(fill="x", pady=3)

        # Group 2: Manual Options
        grp2 = ttk.LabelFrame(left, text="  2. Export Options  ", padding=10)
        grp2.pack(fill="x", pady=(0, 10))

        # -- TOI --
        self._radio_group(grp2, "Terms of Invoice (TOI):",
                          self.var_toi, ["C&F", "FOB", "CIF", "C&I"])

        ttk.Separator(grp2, orient="horizontal").pack(fill="x", pady=5)

        # -- Exim Scheme --
        self._radio_group(grp2, "Exim Scheme:",
                          self.var_exim_scheme, ["19", "00"])

        ttk.Separator(grp2, orient="horizontal").pack(fill="x", pady=5)

        # -- End Use --
        self._radio_group(grp2, "End Use:",
                          self.var_end_use, ["GNX100", "GNX200"])

        ttk.Separator(grp2, orient="horizontal").pack(fill="x", pady=5)

        # -- IGST Pay Status --
        self._radio_group(
            grp2, "IGST Pay Status:",
            self.var_igst_pay,
            ["Export Under Bond - Not Paid", "Export Against Payment"],
        )

        ttk.Separator(grp2, orient="horizontal").pack(fill="x", pady=5)

        # -- IGST Rate --
        self._radio_group(grp2, "IGST Rate:",
                          self.var_igst_rate, ["18", "5", ""])

        ttk.Separator(grp2, orient="horizontal").pack(fill="x", pady=5)

        # -- DBK Under --
        self._radio_group(grp2, "DBK Under:",
                          self.var_dbk_under, ["Actual", "Provision"])

        ttk.Separator(grp2, orient="horizontal").pack(fill="x", pady=5)

        # -- Origin District --
        f_od = tk.Frame(grp2, bg="#ffffff")
        f_od.pack(fill="x", pady=3)
        ttk.Label(f_od, text="Origin District:").pack(side="left")
        ttk.Entry(f_od, textvariable=self.var_origin_district, width=10).pack(
            side="left", padx=10,
        )

        # -- PTAFTA Code --
        self._radio_group(grp2, "PTAFTA Code:",
                          self.var_ptafta, ["GSTP", "NCPTI"])

        # Group 3: Action
        grp3 = ttk.LabelFrame(left, text="  3. Generate  ", padding=10)
        grp3.pack(fill="x", pady=(0, 10))

        ttk.Button(
            grp3, text="▶  Parse Invoice",
            command=self._parse_invoice, style="Primary.TButton",
        ).pack(fill="x", pady=3)

        ttk.Button(
            grp3, text="💾  Export Logisys CSV",
            command=self._export_csv, style="Success.TButton",
        ).pack(fill="x", pady=3)

        # ---- RIGHT PANEL ----
        right = tk.Frame(body, bg="#ffffff")
        right.pack(side="right", fill="both", expand=True, padx=(15, 0))

        # Data Preview Treeview — all 114 columns matching Logisys template
        ttk.Label(
            right, text="Extracted Items Preview (114 Columns)",
            font=("Segoe UI", 10, "bold"),
        ).pack(anchor="w", pady=(0, 5))

        tree_frame = tk.Frame(right, bg="#ffffff")
        tree_frame.pack(fill="both", expand=True)

        # Use all 114 header columns from the template
        cols = tuple(LOGISYS_HEADER)
        self.tree = ttk.Treeview(tree_frame, columns=cols, show="headings", height=15)

        # Default column width — wider for key fields
        wide_cols = {"Product Desc": 250, "IGST Pay Status": 180, "End Use": 80}
        for c in cols:
            self.tree.heading(c, text=c)
            w = wide_cols.get(c, 90)
            self.tree.column(c, width=w, minwidth=60, anchor="center")
        self.tree.column("Product Desc", anchor="w")

        # Vertical scrollbar
        scrollbar_y = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar_y.set)
        scrollbar_y.pack(side="right", fill="y")

        # Horizontal scrollbar for 114 columns
        scrollbar_x = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(xscrollcommand=scrollbar_x.set)
        scrollbar_x.pack(side="bottom", fill="x")

        self.tree.pack(fill="both", expand=True)

        # Log
        ttk.Label(
            right, text="Log", font=("Segoe UI", 10, "bold"),
        ).pack(anchor="w", pady=(10, 3))

        self.log_box = scrolledtext.ScrolledText(
            right, height=8, font=("Consolas", 9), bg="#f9f9f9", bd=1,
        )
        self.log_box.pack(fill="x")

    # ------------------------------------------------------------------ #
    #  FOOTER
    # ------------------------------------------------------------------ #
    def _create_footer(self) -> None:
        footer = tk.Frame(self.master, bg="#f0f0f0", height=30)
        footer.pack(fill="x", side="bottom")

        tk.Label(
            footer, text="© Nagarkot Forwarders Pvt Ltd",
            bg="#f0f0f0", fg="#666", font=("Segoe UI", 8),
        ).pack(side="left", padx=10, pady=5)

        self.lbl_status = tk.Label(
            footer, text="Ready", bg="#f0f0f0",
            fg="#0056b3", font=("Segoe UI", 8, "bold"),
        )
        self.lbl_status.pack(side="right", padx=10, pady=5)

    # ------------------------------------------------------------------ #
    #  HELPERS
    # ------------------------------------------------------------------ #
    def _radio_group(
        self, parent: tk.Widget, label: str,
        var: tk.StringVar, options: list[str],
    ) -> None:
        """Create a labeled row of radio buttons."""
        f = tk.Frame(parent, bg="#ffffff")
        f.pack(fill="x", pady=3)
        ttk.Label(f, text=label).pack(anchor="w")
        rf = tk.Frame(f, bg="#ffffff")
        rf.pack(anchor="w", padx=15)
        for opt in options:
            display = opt if opt else "None"
            ttk.Radiobutton(rf, text=display, variable=var, value=opt).pack(
                side="left", padx=5,
            )

    def _log(self, msg: str) -> None:
        self.log_box.insert("end", msg + "\n")
        self.log_box.see("end")
        self.master.update_idletasks()

    def _status(self, msg: str) -> None:
        self.lbl_status.config(text=msg)
        self.master.update_idletasks()

    # ------------------------------------------------------------------ #
    #  FILE PICKERS
    # ------------------------------------------------------------------ #
    def _pick_pdf(self) -> None:
        path = filedialog.askopenfilename(
            title="Select Skoda Invoice PDF",
            filetypes=[("PDF Files", "*.pdf")],
        )
        if path:
            self.pdf_path = path
            self.lbl_pdf.config(
                text=os.path.basename(path), foreground="#28a745",
            )
            self._log(f"PDF selected: {os.path.basename(path)}")

    def _auto_load_hs_master(self) -> None:
        """Try to auto-load the HS Code Master Excel from the app directory."""
        master_path = resource_path(HS_MASTER_FILENAME)
        if os.path.exists(master_path):
            self.hs_map_path = master_path
            self.hs_unit_map = load_hs_unit_map(master_path)
            self.lbl_hs.config(
                text=f"HS Master: {len(self.hs_unit_map)} codes loaded ✓",
                foreground="#28a745",
            )
            self._log(f"HS Code Master auto-loaded: {len(self.hs_unit_map)} codes.")
        else:
            self.lbl_hs.config(
                text="HS Master not found — click Reload to load manually",
                foreground="#cc0000",
            )
            self._log(f"WARNING: '{HS_MASTER_FILENAME}' not found in app directory.")

    def _pick_hs_map(self) -> None:
        path = filedialog.askopenfilename(
            title="Select HS Code Master Excel",
            filetypes=[("Excel Files", "*.xlsx *.xls")],
        )
        if path:
            self.hs_map_path = path
            self.hs_unit_map = load_hs_unit_map(path)
            self.lbl_hs.config(
                text=f"HS Master: {len(self.hs_unit_map)} codes loaded ✓",
                foreground="#28a745",
            )
            self._log(f"HS map loaded: {len(self.hs_unit_map)} entries.")

    # ------------------------------------------------------------------ #
    #  PARSE
    # ------------------------------------------------------------------ #
    def _parse_invoice(self) -> None:
        if not self.pdf_path:
            messagebox.showwarning("Warning", "Please select an invoice PDF first.")
            return

        self._status("Parsing invoice...")
        self._log("--- Parsing PDF ---")

        try:
            self.parsed_data = parse_skoda_invoice(self.pdf_path)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to parse PDF:\n{e}")
            self._log(f"ERROR: {e}")
            self._status("Parse failed")
            return

        # Build Logisys rows for preview (uses current GUI settings)
        preview_rows = build_logisys_rows(
            invoice_data=self.parsed_data,
            toi=self.var_toi.get(),
            exim_scheme=self.var_exim_scheme.get(),
            end_use=self.var_end_use.get(),
            igst_pay_status=self.var_igst_pay.get(),
            igst_rate=self.var_igst_rate.get(),
            dbk_under=self.var_dbk_under.get(),
            origin_district=self.var_origin_district.get(),
            ptafta_code=self.var_ptafta.get(),
            hs_unit_map=self.hs_unit_map,
        )

        # Populate treeview with all 114 columns
        self.tree.delete(*self.tree.get_children())
        for row in preview_rows:
            self.tree.insert("", "end", values=row)

        count = len(self.parsed_data["items"])
        self._log(
            f"Invoice {self.parsed_data['invoice_no']} parsed: "
            f"{count} items found."
        )
        self._status(f"Parsed: {count} items")
        messagebox.showinfo("Success", f"Parsed {count} items from invoice.")

    # ------------------------------------------------------------------ #
    #  EXPORT
    # ------------------------------------------------------------------ #
    def _export_csv(self) -> None:
        if not self.parsed_data or not self.parsed_data["items"]:
            messagebox.showwarning(
                "Warning",
                "No data to export. Parse an invoice first.",
            )
            return

        self._status("Building Logisys CSV...")

        rows = build_logisys_rows(
            invoice_data=self.parsed_data,
            toi=self.var_toi.get(),
            exim_scheme=self.var_exim_scheme.get(),
            end_use=self.var_end_use.get(),
            igst_pay_status=self.var_igst_pay.get(),
            igst_rate=self.var_igst_rate.get(),
            dbk_under=self.var_dbk_under.get(),
            origin_district=self.var_origin_district.get(),
            ptafta_code=self.var_ptafta.get(),
            hs_unit_map=self.hs_unit_map,
        )

        # Ask where to save
        default_name = f"Logisys_Export_{self.parsed_data['invoice_no']}.csv"
        save_path = filedialog.asksaveasfilename(
            title="Save Logisys CSV",
            defaultextension=".csv",
            initialfile=default_name,
            filetypes=[("CSV Files", "*.csv")],
        )
        if not save_path:
            return

        try:
            write_logisys_csv(rows, save_path)
            self._log(f"CSV exported: {save_path} ({len(rows)} rows)")
            self._status("Export complete!")
            messagebox.showinfo(
                "Export Complete",
                f"Logisys CSV saved successfully!\n\n"
                f"File: {os.path.basename(save_path)}\n"
                f"Items: {len(rows)}",
            )
        except Exception as e:
            messagebox.showerror("Error", f"Failed to write CSV:\n{e}")
            self._log(f"EXPORT ERROR: {e}")
            self._status("Export failed")


# ============================================================================
#  MAIN
# ============================================================================

if __name__ == "__main__":
    root = tk.Tk()
    app = SkodaExportApp(root)
    root.mainloop()
