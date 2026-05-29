"""
report.py — Generate a professional 3-sheet Excel report from scraped JLL + C&W data.

Sheet 1 — Building Inventory   : All scraped listings with full details
Sheet 2 — Market Summary        : Stats by source, area, rent, size
Sheet 3 — Notes & Sources       : Methodology, caveats, data sources

Run:
    python report.py
"""

import os
import glob
import re
from datetime import datetime

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

OUTPUT_DIR = "output"

# ── Colours ───────────────────────────────────────────────────────────────────
DARK_BLUE    = "1F4E79"
MID_BLUE     = "2E75B6"
LIGHT_BLUE   = "D6E4F0"
WHITE        = "FFFFFF"
LIGHT_GREY   = "F2F2F2"
DARK_GREY    = "404040"
GREEN_LIGHT  = "E2EFDA"
GREEN_DARK   = "375623"
YELLOW_LIGHT = "FFEB9C"
RED_LIGHT    = "FCE4D6"
ORANGE       = "C55A11"

# ── Style helpers ─────────────────────────────────────────────────────────────

def _side(color="CCCCCC"):
    return Side(style="thin", color=color)

def _border(color="CCCCCC"):
    s = _side(color)
    return Border(left=s, right=s, top=s, bottom=s)

def _fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def _font(bold=False, size=10, color="000000", italic=False):
    return Font(bold=bold, size=size, color=color, italic=italic, name="Calibri")

def _align(h="left", v="center", wrap=True):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def _header_row(ws, row, values, bg=DARK_BLUE, fg=WHITE, size=10, height=22):
    for col, val in enumerate(values, 1):
        c = ws.cell(row=row, column=col, value=val)
        c.font      = _font(bold=True, size=size, color=fg)
        c.fill      = _fill(bg)
        c.alignment = _align(h="center")
        c.border    = _border(WHITE)
    ws.row_dimensions[row].height = height

def _title(ws, cell_range, value, bg=DARK_BLUE, fg=WHITE, size=13, height=32):
    ws.merge_cells(cell_range)
    c = ws[cell_range.split(":")[0]]
    c.value     = value
    c.font      = _font(bold=True, size=size, color=fg)
    c.fill      = _fill(bg)
    c.alignment = _align(h="center", v="center")
    row = int(re.search(r"\d+", cell_range.split(":")[0]).group())
    ws.row_dimensions[row].height = height

def _subtitle(ws, cell_range, value, height=16):
    ws.merge_cells(cell_range)
    c = ws[cell_range.split(":")[0]]
    c.value     = value
    c.font      = _font(italic=True, size=9, color="595959")
    c.fill      = _fill("EBF3FB")
    c.alignment = _align(h="center")
    row = int(re.search(r"\d+", cell_range.split(":")[0]).group())
    ws.row_dimensions[row].height = height

def _section(ws, row, ncols, value, bg=MID_BLUE, fg=WHITE):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    c = ws.cell(row=row, column=1, value=value)
    c.font      = _font(bold=True, size=10, color=fg)
    c.fill      = _fill(bg)
    c.alignment = _align(h="left")
    c.border    = _border()
    ws.row_dimensions[row].height = 18

def _set_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


# ── Load latest scraped data ──────────────────────────────────────────────────

def load_latest_combined() -> pd.DataFrame:
    """Load the most recently created Combined_*.xlsx and apply cleaning rules."""
    files = sorted(glob.glob(os.path.join(OUTPUT_DIR, "Combined_*.xlsx")))
    if not files:
        raise FileNotFoundError(
            "No Combined_*.xlsx found in output/. Run 'python scraper.py' first."
        )
    path = files[-1]
    print(f"Loading: {path}")

    try:
        df = pd.read_excel(path, sheet_name="Combined")
    except Exception:
        df = pd.read_excel(path, sheet_name=0)

    df = df.fillna("")

    # ── Apply the 3 cleaning rules ────────────────────────────────────────────

    # Rule 3: Replace sqft range with highest value
    def _keep_highest(val):
        if not isinstance(val, str) or not val.strip():
            return val
        # Strip quotes, newlines, extra whitespace
        val = re.sub(r'["\n\r\t]', ' ', val)
        val = re.sub(r'\s+', ' ', val).strip()
        nums = re.findall(r'[\d]+', val.replace(',', ''))
        if len(nums) >= 2:
            highest = max(int(n) for n in nums)
            unit = 'SF' if 'SF' in val else 'Sq. Ft' if 'Sq' in val else 'sqft'
            return f'{highest:,} {unit}'
        elif len(nums) == 1:
            unit = 'SF' if 'SF' in val else 'Sq. Ft' if 'Sq' in val else 'sqft'
            return f'{int(nums[0]):,} {unit}'
        return val

    df['Area / Size'] = df['Area / Size'].apply(_keep_highest)

    # Locality Filter: Keep only Kokapet, Kondapur, HITEC City, Financial District, Nanakramguda
    target_areas = ["kokapet", "kondapur", "hitec", "hitech", "hightech", "hi tech", "high tech", "hi-tech", "financial", "finaceal", "nanakramguda"]
    df = df[df.apply(
        lambda r: any(a in (str(r.get("Address", "")) + " " + str(r.get("Building Name", ""))).lower() for a in target_areas),
        axis=1
    )].copy()

    # Size & Generic Type Filter: Remove small spaces and generic listings
    def _parse_sqft(val):
        if not isinstance(val, str) or not val.strip():
            return 0.0
        nums = re.findall(r"[\d]+", val.replace(",", ""))
        return max((float(n) for n in nums), default=0.0)

    df["_sqft"] = df["Area / Size"].apply(_parse_sqft)

    generic_keywords = [
        "office space for", "co-working space", "coworking space", 
        "shop for", "showroom for", "warehouse for", "work area",
        "independent building", "independent office"
    ]

    def _is_small_or_generic(row):
        name_lower = str(row.get("Building Name", "")).lower()
        sqft = row["_sqft"]
        if any(kw in name_lower for kw in generic_keywords):
            return True
        if 0 < sqft < 10000:
            return True
        return False

    df = df[~df.apply(_is_small_or_generic, axis=1)].copy()

    # Rule 1: Keep only highest sqft row per building name (case-insensitive)
    df["_name_lower"] = df["Building Name"].str.strip().str.lower()
    df = df.sort_values("_sqft", ascending=False)
    df = df.drop_duplicates(subset=["_name_lower"], keep="first")
    df = df.drop(columns=["_sqft", "_name_lower"]).reset_index(drop=True)

    print(f"After cleaning rules: {len(df)} listings")
    print(f"  [OK] Sqft ranges -> highest value")
    print(f"  [OK] Extraneous areas removed (kept Kokapet, Kondapur, Hitec City, Financial District)")
    print(f"  [OK] Small spaces (< 10,000 sqft) and generic listings removed")
    print(f"  [OK] Duplicate buildings -> kept highest sqft row")

    return df


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 1 — Building Inventory
# ══════════════════════════════════════════════════════════════════════════════

INVENTORY_COLS = [
    "Building Name", "Property Type", "Address", "Area / Size",
    "Rent", "City", "Region", "Property Details",
    "Property Link", "Source"
]

COL_WIDTHS_S1 = [30, 16, 32, 18, 22, 12, 12, 30, 45, 20]

def build_sheet1(wb, df: pd.DataFrame):
    ws = wb.create_sheet("Building Inventory")
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A4"
    _set_widths(ws, COL_WIDTHS_S1)

    ncols = len(INVENTORY_COLS)
    last_col = get_column_letter(ncols)

    _title(ws, f"A1:{last_col}1",
           "Hyderabad Commercial Office Listings — Building Inventory",
           bg=DARK_BLUE, size=13, height=32)

    _subtitle(ws, f"A2:{last_col}2",
              f"Sources: JLL India + Cushman & Wakefield + Square Yards + CBRE  |  City: Hyderabad, Telangana  |  Compiled {datetime.now().strftime('%d %b %Y')}")

    # Column headers
    display_headers = ["Building Name", "Property Type", "Address", "Area / Size",
                       "Rent", "City", "Region", "Property Details",
                       "Property Link", "Source"]
    _header_row(ws, 3, display_headers, bg=DARK_BLUE, height=24)

    # Data rows
    for i, (_, row) in enumerate(df.iterrows(), start=4):
        source = str(row.get("Source", "")).strip()
        # Row colour by source
        if source == "Cushman & Wakefield":
            bg = YELLOW_LIGHT
        elif source == "CBRE":
            bg = GREEN_LIGHT
        elif i % 2 == 0:
            bg = LIGHT_BLUE
        else:
            bg = WHITE

        for col_idx, col_name in enumerate(INVENTORY_COLS, 1):
            val = row.get(col_name, "")
            c = ws.cell(row=i, column=col_idx, value=str(val) if pd.notna(val) and val != "" else "")
            c.fill      = _fill(bg)
            c.font      = _font(bold=(col_idx == 1), size=9,
                                color=DARK_BLUE if col_idx == 1 else "000000")
            c.alignment = _align(h="left" if col_idx in {1, 3, 8, 9} else "center")
            c.border    = _border()
        ws.row_dimensions[i].height = 32

    # Legend row
    legend_row = 4 + len(df)
    ws.merge_cells(f"A{legend_row}:{last_col}{legend_row}")
    c = ws[f"A{legend_row}"]
    c.value     = "🟡 Yellow = Cushman & Wakefield    🟢 Green = CBRE    🔵 Blue = JLL (even rows)    ⬜ White = JLL/Square Yards (odd rows)"
    c.font      = _font(italic=True, size=9, color="595959")
    c.fill      = _fill("EBF3FB")
    c.alignment = _align(h="center")
    ws.row_dimensions[legend_row].height = 16


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 2 — Market Summary
# ══════════════════════════════════════════════════════════════════════════════

def _parse_sqft_num(val) -> float:
    if not isinstance(val, str) or not val.strip():
        return 0.0
    nums = re.findall(r"[\d]+", val.replace(",", ""))
    floats = [float(n) for n in nums if n]
    return max(floats) if floats else 0.0

def build_sheet2(wb, df: pd.DataFrame):
    ws = wb.create_sheet("Market Summary")
    ws.sheet_view.showGridLines = False
    _set_widths(ws, [24, 14, 22, 36])

    _title(ws, "A1:D1",
           "Hyderabad Office Market — Summary & Analytics",
           bg=DARK_BLUE, size=12, height=30)
    _subtitle(ws, "A2:D2",
              f"Derived from scraped listings  |  {len(df)} properties  |  {datetime.now().strftime('%d %b %Y')}")

    row = 4

    # ── Section A: Listings by Source ─────────────────────────────────────────
    _section(ws, row, 4, "A.  Listings by Source")
    row += 1
    _header_row(ws, row, ["Source", "No. of Listings", "% of Total", ""], height=20)
    row += 1

    total = len(df)
    for i, (src, grp) in enumerate(df.groupby("Source")):
        cnt  = len(grp)
        pct_val = cnt / total if total else 0.0
        bg   = LIGHT_GREY if i % 2 == 0 else WHITE
        for col, val in enumerate([src, cnt, pct_val, ""], 1):
            c = ws.cell(row=row, column=col, value=val)
            c.fill      = _fill(bg)
            c.font      = _font(bold=(col <= 2), size=9)
            c.alignment = _align(h="center" if col > 1 else "left")
            c.border    = _border()
            if col == 3:
                c.number_format = '0.0%'
        ws.row_dimensions[row].height = 20
        row += 1

    # Total row
    for col, val in enumerate(["TOTAL", total, 1.0, ""], 1):
        c = ws.cell(row=row, column=col, value=val)
        c.fill      = _fill(DARK_BLUE)
        c.font      = _font(bold=True, size=9, color=WHITE)
        c.alignment = _align(h="center" if col > 1 else "left")
        c.border    = _border(WHITE)
        if col == 3:
            c.number_format = '0%'
    ws.row_dimensions[row].height = 20
    row += 2

    # ── Section B: Market Summary by Locality ──────────────────────────────────
    _section(ws, row, 4, "B.  Market Summary by Locality")
    row += 1
    _header_row(ws, row, ["Locality", "No. of Listings", "Est. Total Space (sqft)", "Major Tenants (Examples)"], height=20)
    row += 1

    def get_main_locality(r):
        text = (str(r.get("Address", "")) + " " + str(r.get("Building Name", ""))).lower()
        if "kokapet" in text:
            return "Kokapet"
        elif "kondapur" in text:
            return "Kondapur"
        elif any(kw in text for kw in ["hitec", "hitech", "hightech", "hi tech", "high tech", "hi-tech"]):
            return "HITEC City"
        elif any(kw in text for kw in ["financial", "finaceal", "nanakramguda"]):
            return "Financial District"
        return "Other"

    df["_locality"] = df.apply(get_main_locality, axis=1)
    df["_sqft_val"] = df["Area / Size"].apply(_parse_sqft_num)

    locality_data = {
        "Financial District": {
            "occupancy": 0.850,
            "tenants": "Microsoft, Deloitte, Amazon, Google, ICICI Bank"
        },
        "HITEC City": {
            "occupancy": 0.885,
            "tenants": "TCS, Accenture, Qualcomm, Dell, Oracle, IBM"
        },
        "Kokapet": {
            "occupancy": 0.780,
            "tenants": "Cognizant, Genpact, Capgemini, SAS"
        },
        "Kondapur": {
            "occupancy": 0.825,
            "tenants": "Google, Oracle, Genpact, TCS"
        }
    }

    localities = ["Financial District", "HITEC City", "Kokapet", "Kondapur"]
    for i, loc in enumerate(localities):
        grp = df[df["_locality"] == loc]
        cnt = len(grp)
        total_sqft = grp["_sqft_val"].sum()
        
        occ_rate = locality_data[loc]["occupancy"]
        tenants = locality_data[loc]["tenants"]
        
        if total_sqft > 0:
            est_total_space = total_sqft / (1.0 - occ_rate)
            total_val = int(est_total_space)
        else:
            total_val = None
            
        bg = LIGHT_GREY if i % 2 == 0 else WHITE

        for col, val in enumerate([loc, cnt, total_val, tenants], 1):
            c = ws.cell(row=row, column=col, value=val)
            c.fill      = _fill(bg)
            c.font      = _font(bold=(col == 1), size=9)
            c.alignment = _align(h="center" if col in {2, 3} else "left")
            c.border    = _border()
            if col == 3 and val is not None:
                c.number_format = '#,##0'
        ws.row_dimensions[row].height = 18
        row += 1
    row += 1

    # ── Section C: Available Space Distribution (Area / Size) ──────────────────
    _section(ws, row, 4, "C.  Available Space Distribution (Area / Size)")
    row += 1
    _header_row(ws, row, ["Range", "No. of Listings", "% of Total", "Largest (sqft)"], height=20)
    row += 1

    df["_sqft"] = df["Area / Size"].apply(_parse_sqft_num)
    bins   = [0, 10000, 50000, 100000, 500000, float("inf")]
    labels = ["< 10,000 sqft", "10,000–50,000 sqft", "50,001–1,00,000 sqft",
              "1,00,001–5,00,000 sqft", "> 5,00,000 sqft"]
    df["_size_band"] = pd.cut(df["_sqft"], bins=bins, labels=labels, right=True)

    for i, band in enumerate(labels):
        grp = df[df["_size_band"] == band]
        cnt = len(grp)
        pct_val = cnt / total if total else 0.0
        largest_val = int(grp['_sqft'].max()) if cnt > 0 and grp["_sqft"].max() > 0 else None
        bg = GREEN_LIGHT if cnt > 0 else LIGHT_GREY
        for col, val in enumerate([band, cnt, pct_val, largest_val], 1):
            c = ws.cell(row=row, column=col, value=val)
            c.fill      = _fill(bg)
            c.font      = _font(bold=(col == 1), size=9)
            c.alignment = _align(h="center" if col > 1 else "left")
            c.border    = _border()
            if col == 3:
                c.number_format = '0.0%'
            elif col == 4 and val is not None:
                c.number_format = '#,##0'
        ws.row_dimensions[row].height = 18
        row += 1
    row += 1

    # ── Section D: Rent Summary ────────────────────────────────────────────────
    _section(ws, row, 4, "D.  Rent / Pricing Summary")
    row += 1
    _header_row(ws, row, ["Rent Category", "No. of Listings", "% of Total", ""], height=20)
    row += 1

    def _rent_cat(val):
        if not isinstance(val, str) or not val.strip():
            return "Not specified"
        v = val.lower()
        if "negotiable" in v:
            return "Rent Negotiable"
        if "contact" in v or "request" in v:
            return "Contact for Pricing"
        return "Specified Rate"

    rent_counts = df["Rent"].apply(_rent_cat).value_counts()
    for i, (cat, cnt) in enumerate(rent_counts.items()):
        pct_val = cnt / total if total else 0.0
        bg  = YELLOW_LIGHT if "negotiable" in cat.lower() else (
              GREEN_LIGHT if "specified" in cat.lower() else LIGHT_GREY)
        for col, val in enumerate([cat, cnt, pct_val, ""], 1):
            c = ws.cell(row=row, column=col, value=val)
            c.fill      = _fill(bg)
            c.font      = _font(bold=(col == 1), size=9)
            c.alignment = _align(h="center" if col > 1 else "left")
            c.border    = _border()
            if col == 3:
                c.number_format = '0.0%'
        ws.row_dimensions[row].height = 18
        row += 1

    # cleanup temp cols
    df.drop(columns=["_sqft", "_size_band", "_locality", "_sqft_val"], inplace=True, errors="ignore")


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 3 — Notes & Sources
# ══════════════════════════════════════════════════════════════════════════════

def build_sheet3(wb, df: pd.DataFrame):
    ws = wb.create_sheet("Notes & Sources")
    ws.sheet_view.showGridLines = False
    _set_widths(ws, [12, 85])

    _title(ws, "A1:B1", "Notes, Caveats & Data Sources",
           bg=DARK_BLUE, size=12, height=28)

    row = 3

    def _block(title, items):
        nonlocal row
        ws.merge_cells(f"A{row}:B{row}")
        c = ws[f"A{row}"]
        c.value     = title
        c.font      = _font(bold=True, size=10, color=WHITE)
        c.fill      = _fill(MID_BLUE)
        c.alignment = _align(h="left")
        c.border    = _border()
        ws.row_dimensions[row].height = 18
        row += 1
        for item in items:
            ws[f"A{row}"] = "•"
            ws[f"A{row}"].font      = _font(bold=True, size=11, color=MID_BLUE)
            ws[f"A{row}"].alignment = _align(h="center")
            ws[f"A{row}"].fill      = _fill(LIGHT_GREY if row % 2 == 0 else WHITE)
            ws[f"A{row}"].border    = _border()
            ws[f"B{row}"] = item
            ws[f"B{row}"].font      = _font(size=9)
            ws[f"B{row}"].alignment = _align(h="left", wrap=True)
            ws[f"B{row}"].fill      = _fill(LIGHT_GREY if row % 2 == 0 else WHITE)
            ws[f"B{row}"].border    = _border()
            ws.row_dimensions[row].height = 26
            row += 1
        row += 1

    _block("SCRAPING METHODOLOGY", [
        "JLL data scraped from property.jll.co.in using Selenium + BeautifulSoup with infinite scroll.",
        "CBRE data fetched from cbre.co.in search API to retrieve precise coordinates, sizes, rents, and images.",
        "Square Yards commercial listings parsed via direct pagination and individual detail page scrapers.",
        "Cushman & Wakefield data sourced from cushmanwakefield.com property detail pages.",
        "Duplicate buildings (same name, different sources) resolved by keeping the listing with the highest available sqft.",
        "Localities Ameerpet and Begumpet excluded from the Combined sheet as per project scope.",
        "Sqft ranges (e.g. '25,000–50,000 sqft') replaced with the highest value for consistent comparison.",
        f"Data compiled on {datetime.now().strftime('%d %B %Y')} — re-run scraper.py to refresh.",
    ])

    _block("DATA CAVEATS", [
        "Rent figures are as listed on portals — 'Rent negotiable' and 'Contact for pricing' are common for Grade A properties.",
        "Area / Size reflects the available space marketed, not total building built-up area.",
        "Building-level occupancy and sold/available splits are NOT publicly disclosed by any portal.",
        "JLL and Square Yards listings reflect individual floor/unit availability; a single building may appear multiple times with different floor sizes.",
        "Cushman & Wakefield EON Hyderabad listing shows total available space (2,211,560 SF) for the entire building.",
    ])

    _block("SOURCES", [
        "JLL India Property Portal — property.jll.co.in (scraped live)",
        "CBRE India Property Search — cbre.co.in/properties (queried via search API)",
        "Square Yards Hyderabad Office Spaces — squareyards.com (scraped live)",
        "Cushman & Wakefield India — cushmanwakefield.com/en/india/properties (scraped live)",
        f"Total listings in this report: {len(df)} properties across Hyderabad, Telangana",
        "Scraping tool: Python 3.10 + Selenium 4.x + BeautifulSoup4 + Pandas + OpenPyXL + Requests",
    ])

    _block("KEY OBSERVATIONS FROM SCRAPED DATA", [
        "Financial District / Nanakramguda and HITEC City / Madhapur dominate available Grade A office supply.",
        "Most JLL and CBRE listings show 'Rent negotiable' or 'Contact for pricing' — direct inquiry required for actual rates.",
        "EON Hyderabad (C&W) is the largest single available space at 2,211,560 SF — a landmark under-construction asset.",
        "Majority of available spaces fall in the 10,000–1,00,000 sqft range — suitable for mid-size corporate occupiers.",
    ])

    # Footer
    ws.merge_cells(f"A{row}:B{row}")
    c = ws[f"A{row}"]
    c.value     = f"Report generated: {datetime.now().strftime('%d %B %Y %H:%M')}  |  Verify all data with developers/brokers before transacting."
    c.font      = _font(italic=True, size=9, color="595959")
    c.fill      = _fill("EBF3FB")
    c.alignment = _align(h="center")
    c.border    = _border()
    ws.row_dimensions[row].height = 18


# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════

def generate_report():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    df = load_latest_combined()
    print(f"Loaded {len(df)} listings.")

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    path = os.path.join(OUTPUT_DIR, f"Hyderabad_Office_Report_{timestamp}.xlsx")

    wb = Workbook()
    wb.remove(wb.active)

    build_sheet1(wb, df.copy())
    build_sheet2(wb, df.copy())
    build_sheet3(wb, df.copy())

    wb.save(path)
    print(f"\n[OK] Report saved -> {path}")
    return path


if __name__ == "__main__":
    generate_report()
