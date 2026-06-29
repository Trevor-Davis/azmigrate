"""
Generate the Kettering assessment deck from source workbooks.

Expected input files in --input-dir:
  - Discovery.xlsx
  - Strategy_Lift_and_shift.xlsx
  - Strategy_PaaS_Preferred.xlsx

Usage:
  python generate_kettering_deck.py --input-dir "C:\\path\\to\\detailed-report" --output "Kettering-Slides-generated.pptx"

Notes:
  - Values are recalculated from the source files each run.
  - Discovery.xlsx may be MIP/encrypted or an older Excel container. If pandas cannot
    read it directly, the script attempts to use local Excel via pywin32 to save a
    temporary readable .xlsx copy. Install pywin32 if needed:
      python -m pip install pywin32
"""

from __future__ import annotations

import argparse
import ast
import csv
import io
import math
import re
import shutil
import tempfile
import urllib.request
from pathlib import Path

import pandas as pd
from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt


SLIDE_W = 13.333
SLIDE_H = 7.5

NAVY = RGBColor(30, 39, 97)
SLATE = RGBColor(54, 69, 79)
MUTED = RGBColor(100, 116, 139)
TEAL = RGBColor(2, 128, 144)
TEAL_DARK = RGBColor(25, 103, 128)
MINT = RGBColor(2, 195, 154)
GREEN = RGBColor(34, 139, 34)
RED = RGBColor(220, 38, 38)
ORANGE = RGBColor(255, 128, 0)
WHITE = RGBColor(255, 255, 255)
BG = RGBColor(248, 250, 252)
LINE = RGBColor(218, 226, 238)
ALT = RGBColor(241, 245, 249)
PINK = RGBColor(244, 114, 182)


DB_RESOURCE_TYPES = {
    "microsoft.applicationmigration/mongosites/mongoinstances",
    "microsoft.applicationmigration/pgsqlsites/pgsqlinstances",
    "microsoft.mysqldiscovery/mysqlsites/mysqlservers",
    "microsoft.offazure/mastersites/sqlsites/sqlservers",
}
WEBAPP_RESOURCE_TYPES = {
    "microsoft.offazure/mastersites/webappsites/iiswebapplications",
    "microsoft.offazure/mastersites/webappsites/tomcatwebapplications",
}
FILESHARE_RESOURCE_TYPE = "microsoft.applicationmigration/storagesites/fileshares"
MACHINE_RESOURCE_TYPE = "microsoft.offazure/vmwaresites/machines"

SKU_MEMORY_GB = {
    "Standard_D2as_v4": 8,
    "Standard_D2as_v5": 8,
    "Standard_D2ds_v4": 8,
    "Standard_D4as_v5": 16,
    "Standard_D8as_v5": 32,
    "Standard_D8s_v5": 32,
    "Standard_D16as_v5": 64,
    "Standard_D48as_v5": 192,
    "Standard_E2as_v5": 16,
    "Standard_E2bs_v5": 16,
    "Standard_E2s_v6": 16,
    "Standard_E4as_v5": 32,
    "Standard_E4bs_v5": 32,
    "Standard_E4s_v6": 32,
    "Standard_E8as_v5": 64,
    "Standard_E8bs_v5": 64,
    "Standard_E8s_v6": 64,
    "Standard_E16bs_v5": 128,
    "Standard_E16s_v6": 128,
    "Standard_E20as_v5": 160,
    "Standard_E20s_v6": 160,
    "Standard_E32as_v5": 256,
    "Standard_E32bs_v5": 256,
    "Standard_E48bs_v5": 384,
    "Standard_E64s_v6": 512,
    "Standard_E96s_v6": 768,
    "Standard_E128s_v6": 1024,
    "Standard_E192is_v6": 1832,
    "Standard_FX12mds": 252,
    "Standard_M16bs_v3": 128,
}


def norm_server(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip().upper()
    if not text:
        return ""
    text = text.split(".")[0]
    if "-" in text and re.match(r"^[A-Z]{2,6}[A-Z0-9]*\d{2,}", text):
        text = text.split("-")[0]
    return text.strip()


def recommended_memory_gb(sku) -> float | None:
    if pd.isna(sku):
        return None
    text = str(sku).strip()
    if not text or text.lower() == "unknown":
        return None
    if text in SKU_MEMORY_GB:
        return float(SKU_MEMORY_GB[text])
    match = re.match(r"^Standard_([DEFL])(\d+)[A-Za-z0-9_]*$", text)
    if match:
        family = match.group(1)
        n = int(match.group(2))
        multiplier = {"D": 4, "E": 8, "F": 2, "L": 8}[family]
        return float(n * multiplier)
    match = re.match(r"^Standard_B(\d+)(m?s)?$", text)
    if match:
        n = int(match.group(1))
        suffix = (match.group(2) or "").lower()
        if suffix == "ms":
            return float(n * 4)
        return float(n * 2)
    match = re.match(r"^Standard_M(\d+)[A-Za-z0-9_]*$", text)
    if match:
        return float(int(match.group(1)) * 8)
    return None


def money_k(value: float) -> str:
    return f"${value / 1000:.1f}K"


def tb_from_gb(value: float) -> str:
    return f"{value / 1024:.2f} TB"


def tb_from_gib(value: float) -> str:
    return f"{value / 1024:.1f} TB"


def pb_from_gb(value: float) -> str:
    return f"{value / 1024 / 1024:.2f} PB"


def safe_div(numerator: float, denominator: float) -> float:
    return numerator / denominator if denominator else 0


def pct(numerator: float, denominator: float) -> str:
    return f"{safe_div(numerator, denominator):.1%}"


def delta_pct(recommended: float, onprem: float) -> float:
    return safe_div(recommended - onprem, onprem) * 100


def delta_text(delta: float) -> str:
    return f"{abs(delta):.1f}% {'decrease' if delta < 0 else 'increase'}"


def read_excel_direct(path: Path, sheet: str) -> pd.DataFrame:
    suffix = path.suffix.lower()
    engine = "openpyxl" if suffix in {".xlsx", ".xlsm"} else None
    return pd.read_excel(path, sheet_name=sheet, engine=engine)


def convert_with_excel(path: Path) -> Path:
    try:
        import win32com.client  # type: ignore
    except ImportError as exc:
        raise RuntimeError(
            f"Could not read {path.name} directly and pywin32 is not installed. "
            "Install it with: python -m pip install pywin32"
        ) from exc

    out = Path(tempfile.mkdtemp()) / f"{path.stem}_converted.xlsx"
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = None
    try:
        wb = excel.Workbooks.Open(str(path), 0, True)
        wb.SaveAs(str(out), 51)
    finally:
        if wb is not None:
            wb.Close(False)
        excel.Quit()
    return out


def log(message: str) -> None:
    print(message, flush=True)


AZURE_VMS_CSV_URL = "https://raw.githubusercontent.com/Trevor-Davis/azmigrate/refs/heads/main/AzureVMs.csv"
_GITHUB_SKU_CACHE: dict[str, float] | None = None


def fetch_github_sku_map() -> dict[str, float]:
    """Download and parse the GitHub AzureVMs.csv into {sku: memory_gb}.

    Cached for the life of the process. Returns {} if the fetch fails so the
    rest of the script can still run with built-in fallbacks.
    """
    global _GITHUB_SKU_CACHE
    if _GITHUB_SKU_CACHE is not None:
        return _GITHUB_SKU_CACHE
    log(f"  Fetching SKU lookup from GitHub: {AZURE_VMS_CSV_URL}")
    mapping: dict[str, float] = {}
    try:
        req = urllib.request.Request(AZURE_VMS_CSV_URL, headers={"User-Agent": "azmigrate-script"})
        with urllib.request.urlopen(req, timeout=20) as resp:
            raw = resp.read().decode("utf-8-sig")
        reader = csv.DictReader(io.StringIO(raw))
        if not reader.fieldnames or "Name" not in reader.fieldnames or "MemoryGB" not in reader.fieldnames:
            log(f"    WARNING: CSV missing expected 'Name'/'MemoryGB' columns; got {reader.fieldnames}")
            _GITHUB_SKU_CACHE = {}
            return _GITHUB_SKU_CACHE
        for row in reader:
            name = (row.get("Name") or "").strip()
            mem = (row.get("MemoryGB") or "").strip()
            if not name or not mem:
                continue
            try:
                mapping[name] = float(mem)
            except ValueError:
                continue
        log(f"    Loaded {len(mapping):,} SKU entries from GitHub")
    except Exception as exc:
        log(f"    WARNING: Could not fetch GitHub SKU list ({exc}); falling back to built-in lookup only")
    _GITHUB_SKU_CACHE = mapping
    return _GITHUB_SKU_CACHE


def _col_letter(n: int) -> str:
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def ensure_recommended_memory(lift_path: Path) -> None:
    """In Strategy_Lift_and_shift.xlsx, create/refresh:
       - sheet 'Azure SKUs' with [Azure SKU, Memory (GB)] for every SKU in
         Server_to_AzureVM!RECOMMENDED_COMPUTE_SKU
       - column 'Recommended Memory' in Server_to_AzureVM (right after
         RECOMMENDED_COMPUTE_SKU) populated by VLOOKUP into 'Azure SKUs'

    Uses Excel automation so formulas are evaluated and saved.
    """
    try:
        import win32com.client  # type: ignore
    except ImportError as exc:
        raise RuntimeError(
            "pywin32 required to update the workbook. Install with: "
            "python -m pip install pywin32"
        ) from exc

    log(f"  Updating {lift_path.name}: Azure SKUs sheet + Recommended Memory column")
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = None
    try:
        wb = excel.Workbooks.Open(str(lift_path))
        try:
            ws = wb.Worksheets("Server_to_AzureVM")
        except Exception as exc:
            raise RuntimeError("Server_to_AzureVM sheet not found in lift workbook") from exc

        used = ws.UsedRange
        last_col = used.Columns.Count
        last_row = used.Rows.Count

        # Map header -> column index
        def header_map():
            m = {}
            cols = ws.UsedRange.Columns.Count
            for c in range(1, cols + 1):
                v = ws.Cells(1, c).Value
                if v is not None:
                    m[str(v).strip()] = c
            return m

        headers = header_map()
        if "RECOMMENDED_COMPUTE_SKU" not in headers:
            raise RuntimeError("RECOMMENDED_COMPUTE_SKU column not found")

        # Remove any pre-existing 'Recommended Memory' column (anywhere) so we
        # always re-insert it cleanly right after RECOMMENDED_COMPUTE_SKU.
        if "Recommended Memory" in headers:
            ws.Columns(headers["Recommended Memory"]).Delete()
            headers = header_map()

        sku_col = headers["RECOMMENDED_COMPUTE_SKU"]
        sku_letter = _col_letter(sku_col)

        # Collect unique SKUs (skip blanks and 'Unknown')
        skus = set()
        for r in range(2, last_row + 1):
            v = ws.Cells(r, sku_col).Value
            if v is None:
                continue
            text = str(v).strip()
            if text and text.lower() != "unknown":
                skus.add(text)

        # Preserve any user-supplied memory values already present in the
        # existing Azure SKUs sheet — the workbook is the source of truth.
        existing_memory: dict[str, float] = {}
        try:
            existing_ws = wb.Worksheets("Azure SKUs")
        except Exception:
            existing_ws = None
        if existing_ws is not None:
            try:
                ex_used = existing_ws.UsedRange
                ex_rows = ex_used.Rows.Count
                for r in range(2, ex_rows + 1):
                    sku_val = existing_ws.Cells(r, 1).Value
                    mem_val = existing_ws.Cells(r, 2).Value
                    if sku_val is None:
                        continue
                    key = str(sku_val).strip()
                    if not key:
                        continue
                    try:
                        if mem_val is not None and str(mem_val).strip() != "":
                            existing_memory[key] = float(mem_val)
                    except (TypeError, ValueError):
                        pass
            except Exception:
                pass
            try:
                existing_ws.Delete()
            except Exception:
                pass

        # (Re)create the 'Azure SKUs' lookup sheet, preserving user-entered values
        github_map = fetch_github_sku_map()
        sku_ws = wb.Worksheets.Add(After=ws)
        sku_ws.Name = "Azure SKUs"
        sku_ws.Cells(1, 1).Value = "Azure SKU"
        sku_ws.Cells(1, 2).Value = "Memory (GB)"
        blanks: list[str] = []
        for i, sku in enumerate(sorted(skus), start=2):
            sku_ws.Cells(i, 1).Value = sku
            if sku in existing_memory:
                sku_ws.Cells(i, 2).Value = existing_memory[sku]
                continue
            if sku in github_map:
                sku_ws.Cells(i, 2).Value = float(github_map[sku])
                continue
            mem = recommended_memory_gb(sku)
            if mem is not None:
                sku_ws.Cells(i, 2).Value = float(mem)
            else:
                blanks.append(sku)

        # Insert 'Recommended Memory' column right after RECOMMENDED_COMPUTE_SKU
        target_col = sku_col + 1
        ws.Columns(target_col).Insert()
        ws.Cells(1, target_col).Value = "Recommended Memory"
        target_letter = _col_letter(target_col)

        # Fill VLOOKUP formula for all data rows
        if last_row >= 2:
            formula = (
                f"=IFERROR(VLOOKUP({sku_letter}2,'Azure SKUs'!A:B,2,FALSE),\"\")"
            )
            rng = ws.Range(f"{target_letter}2:{target_letter}{last_row}")
            rng.Formula = formula  # Excel auto-fills relative references

        excel.Calculate()
        wb.Save()
        log("    Azure SKUs sheet + Recommended Memory column updated")
        if blanks:
            log(f"    WARNING: {len(blanks)} SKU(s) have no memory value yet — fill them in the 'Azure SKUs' sheet:")
            for s in blanks:
                log(f"      - {s}")
            log("    These rows will show as 0 in totals until you fill the values. Re-run the script after editing.")
    finally:
        if wb is not None:
            wb.Close(False)
        excel.Quit()


def read_excel_any(path: Path, sheet: str) -> pd.DataFrame:
    log(f"  - Reading {path.name} [{sheet}]")
    try:
        return read_excel_direct(path, sheet)
    except Exception:
        log(f"    direct read failed, converting via Excel automation...")
        converted = convert_with_excel(path)
        return pd.read_excel(converted, sheet_name=sheet, engine="openpyxl")


def require_columns(df: pd.DataFrame, sheet: str, columns: list[str]) -> None:
    missing = [c for c in columns if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns in {sheet}: {missing}")


def add_text(slide, text, x, y, w, h, size=12, color=SLATE, bold=False, font="Aptos", align=None):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = str(text)
    if align:
        p.alignment = align
    run = p.runs[0] if p.runs else p.add_run()
    run_font = run.font
    run_font.name = font
    run_font.size = Pt(size)
    run_font.color.rgb = color
    run_font.bold = bold
    return box


def add_panel(slide, x, y, w, h, title, accent=TEAL, title_size=13):
    rect = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    rect.fill.solid()
    rect.fill.fore_color.rgb = WHITE
    rect.line.color.rgb = LINE
    rect.line.width = Pt(1)
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(0.08))
    bar.fill.solid()
    bar.fill.fore_color.rgb = accent
    bar.line.fill.background()
    add_text(slide, title, x + 0.15, y + 0.14, w - 0.3, 0.3, title_size, NAVY, True)
    return rect


def add_card(slide, x, y, w, h, value, label, sub="", accent=TEAL, value_size=26, label_size=12, sub_size=12):
    rect = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    rect.fill.solid()
    rect.fill.fore_color.rgb = WHITE
    rect.line.color.rgb = LINE
    rect.line.width = Pt(1)
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(0.08))
    bar.fill.solid()
    bar.fill.fore_color.rgb = accent
    bar.line.fill.background()
    add_text(slide, value, x + 0.17, y + 0.30, w - 0.34, 0.40, value_size, NAVY, True)
    add_text(slide, label, x + 0.17, y + 0.82, w - 0.34, 0.30, label_size, SLATE, True)
    if sub:
        add_text(slide, sub, x + 0.17, y + 1.16, w - 0.34, 0.35, sub_size, MUTED)
    return rect


def add_small_card(slide, x, y, w, h, value, label, accent=TEAL, value_size=14, label_size=8):
    rect = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    rect.fill.solid()
    rect.fill.fore_color.rgb = WHITE
    rect.line.color.rgb = LINE
    rect.line.width = Pt(1)
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(0.06))
    bar.fill.solid()
    bar.fill.fore_color.rgb = accent
    bar.line.fill.background()
    add_text(slide, value, x + 0.12, y + 0.20, w - 0.24, 0.28, value_size, NAVY, True)
    add_text(slide, label, x + 0.12, y + 0.58, w - 0.24, 0.22, label_size, SLATE, True)
    return rect


def add_table(slide, rows, x, y, w, h, font_size=9, header_color=NAVY, col_widths=None):
    table_shape = slide.shapes.add_table(len(rows), len(rows[0]), Inches(x), Inches(y), Inches(w), Inches(h))
    table = table_shape.table
    if col_widths:
        for idx, cw in enumerate(col_widths):
            table.columns[idx].width = Inches(cw)
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            cell = table.cell(r, c)
            cell.text = str(val)
            cell.margin_left = cell.margin_right = Inches(0.04)
            cell.margin_top = cell.margin_bottom = Inches(0.02)
            p = cell.text_frame.paragraphs[0]
            p.font.name = "Aptos"
            p.font.size = Pt(font_size)
            p.font.color.rgb = WHITE if r == 0 else SLATE
            p.font.bold = r == 0
            cell.fill.solid()
            cell.fill.fore_color.rgb = header_color if r == 0 else (WHITE if r % 2 else ALT)
    return table_shape


def color_delta_text(cell, text: str, size: int = 11):
    cell.text = text
    p = cell.text_frame.paragraphs[0]
    p.font.size = Pt(size)
    p.font.color.rgb = SLATE
    m = re.search(r"(\d+(?:\.\d+)?% (?:increase|decrease))", text)
    if not m:
        return
    start = m.start(1)
    end = m.end(1)
    # PowerPoint text run coloring is easiest by splitting runs.
    p.text = ""
    before = text[:start]
    delta = text[start:end]
    after = text[end:]
    if before:
        r = p.add_run()
        r.text = before
        r.font.size = Pt(size)
        r.font.color.rgb = SLATE
    r = p.add_run()
    r.text = delta
    r.font.size = Pt(size)
    r.font.bold = True
    r.font.color.rgb = GREEN if "decrease" in delta else RED
    if after:
        r = p.add_run()
        r.text = after
        r.font.size = Pt(size)
        r.font.color.rgb = SLATE


def create_donut_image(path: Path, smb: int, nfs: int, title="Protocol Distribution"):
    width, height, scale = 520, 360, 4
    img = Image.new("RGB", (width * scale, height * scale), (255, 255, 255))
    draw = ImageDraw.Draw(img)
    try:
        title_font = ImageFont.truetype("arialbd.ttf", 28 * scale)
        legend_font = ImageFont.truetype("arial.ttf", 22 * scale)
    except Exception:
        title_font = ImageFont.load_default()
        legend_font = ImageFont.load_default()

    navy = (31, 78, 121)
    teal = (25, 103, 128)
    orange = (255, 128, 0)
    gray = (89, 89, 89)
    white = (255, 255, 255)

    bb = draw.textbbox((0, 0), title, font=title_font)
    draw.text(((width * scale - (bb[2] - bb[0])) / 2, 8 * scale), title, fill=navy, font=title_font)

    cx, cy = 260 * scale, 168 * scale
    r, inner = 118 * scale, 49 * scale
    bbox = [cx - r, cy - r, cx + r, cy + r]
    total = max(smb + nfs, 1)
    nfs_angle = 360 * nfs / total
    start = -90
    draw.pieslice(bbox, start, start + 360, fill=teal)
    draw.pieslice(bbox, start, start + nfs_angle, fill=orange)
    inner_bbox = [cx - inner, cy - inner, cx + inner, cy + inner]
    draw.ellipse(inner_bbox, fill=white)
    draw.ellipse(bbox, outline=white, width=3 * scale)
    draw.ellipse(inner_bbox, outline=white, width=3 * scale)
    for angle in (start, start + nfs_angle):
        rad = math.radians(angle)
        draw.line(
            (
                cx + inner * math.cos(rad),
                cy + inner * math.sin(rad),
                cx + r * math.cos(rad),
                cy + r * math.sin(rad),
            ),
            fill=white,
            width=4 * scale,
        )

    legend_y = 314 * scale
    box = 12 * scale
    items = [(f"NFS ({nfs})", orange), (f"SMB ({smb})", teal)]
    widths = []
    for label, _ in items:
        bb = draw.textbbox((0, 0), label, font=legend_font)
        widths.append(box + 8 * scale + (bb[2] - bb[0]))
    total_w = sum(widths) + 34 * scale
    x = (width * scale - total_w) / 2
    for (label, color), item_w in zip(items, widths):
        draw.rectangle([x, legend_y, x + box, legend_y + box], fill=color, outline=white, width=2 * scale)
        draw.text((x + box + 8 * scale, legend_y - 5 * scale), label, fill=gray, font=legend_font)
        x += item_w + 34 * scale

    img = img.resize((width, height), Image.Resampling.LANCZOS)
    img.save(path)
    return path


def create_readiness_image(path: Path, readiness_counts: dict[str, int]):
    categories = ["Unknown", "Ready With Conditions", "Not Ready", "Ready"]
    values = [int(readiness_counts.get(category, 0)) for category in categories]
    max_value = max(values) or 1
    total = sum(values) or 1

    width, height, scale = 760, 320, 4
    img = Image.new("RGB", (width * scale, height * scale), (255, 255, 255))
    draw = ImageDraw.Draw(img)
    try:
        label_font = ImageFont.truetype("arial.ttf", 18 * scale)
        value_font = ImageFont.truetype("arialbd.ttf", 18 * scale)
    except Exception:
        label_font = ImageFont.load_default()
        value_font = ImageFont.load_default()

    colors = {
        "Unknown": (148, 163, 184),
        "Ready With Conditions": (255, 128, 0),
        "Not Ready": (220, 38, 38),
        "Ready": (34, 139, 34),
    }
    label_x = 26 * scale
    bar_x = 270 * scale
    bar_w = 315 * scale
    row_h = 60 * scale
    y0 = 42 * scale
    for idx, (category, value) in enumerate(zip(categories, values)):
        y = y0 + idx * row_h
        label = category
        draw.text((label_x, y - 10 * scale), label, fill=(54, 69, 79), font=label_font)
        draw.rounded_rectangle(
            [bar_x, y, bar_x + bar_w, y + 18 * scale],
            radius=9 * scale,
            fill=(241, 245, 249),
        )
        filled = max(2 * scale if value else 0, int(bar_w * value / max_value))
        if filled:
            draw.rounded_rectangle(
                [bar_x, y, bar_x + filled, y + 18 * scale],
                radius=9 * scale,
                fill=colors[category],
            )
        draw.text(
            (bar_x + bar_w + 18 * scale, y - 11 * scale),
            f"{value:,} ({value / total:.1%})",
            fill=(30, 39, 97),
            font=value_font,
        )

    img = img.resize((width, height), Image.Resampling.LANCZOS)
    img.save(path)
    return path


def discovery_sets(discovery: pd.DataFrame):
    require_columns(discovery, "Discovery.xlsx/Data", ["resourceType", "parentResourceName"])
    rows = discovery[["resourceType", "parentResourceName"]].copy()
    rows["resourceType"] = rows["resourceType"].astype(str).str.strip()
    rows["server"] = rows["parentResourceName"].map(norm_server)
    files = set(rows.loc[rows["resourceType"].eq(FILESHARE_RESOURCE_TYPE), "server"]) - {""}
    web = set(rows.loc[rows["resourceType"].isin(WEBAPP_RESOURCE_TYPES), "server"]) - {""}
    db = set(rows.loc[rows["resourceType"].isin(DB_RESOURCE_TYPES), "server"]) - {""}
    return files, web, db


def load_metrics(input_dir: Path):
    discovery_path = input_dir / "Discovery.xlsx"
    lift_path = input_dir / "Strategy_Lift_and_shift.xlsx"
    paas_path = input_dir / "Strategy_PaaS_Preferred.xlsx"
    for path in (discovery_path, lift_path, paas_path):
        if not path.exists():
            raise FileNotFoundError(path)

    discovery = read_excel_any(discovery_path, "Data")
    ensure_recommended_memory(lift_path)
    lift = read_excel_any(lift_path, "Server_to_AzureVM")
    fs = read_excel_any(paas_path, "FileShares_to_Azure_Files")

    require_columns(discovery, "Discovery.xlsx/Data", ["resourceType", "parentResourceName"])
    require_columns(
        lift,
        "Strategy_Lift_and_shift.xlsx/Server_to_AzureVM",
        [
            "SERVER_NAME",
            "ONPREM_STORAGE_GB",
            "RECOMMENDED_STORAGE_SIZE_GB",
            "ONPREM_CORES_COUNT",
            "RECOMMENDED_NUMBER_OF_CORES",
            "ONPREM_MEMORY_MB",
            "RECOMMENDED_COMPUTE_SKU",
            "STORAGE_UTILIZATION_PERCENT",
            "ONPREM_CPU_USAGE_PERCENT",
            "ONPREM_MEMORY_USAGE_PERCENT",
            "TOTAL_MONTHLY_COST_USD",
        ],
    )
    require_columns(
        fs,
        "Strategy_PaaS_Preferred.xlsx/FileShares_to_Azure_Files",
        ["SERVER", "PROTOCOL", "READINESS", "AZURE FILES TARGET PROVISIONED SIZE (GIB)", "ESTIMATED MONTHLY COST (USD)"],
    )

    fileshare_servers, web_servers, db_servers = discovery_sets(discovery)
    fileshare_only = fileshare_servers - web_servers - db_servers

    machines = discovery[discovery["resourceType"].eq(MACHINE_RESOURCE_TYPE)].copy()
    vm_total = len(machines)
    power = machines.get("powerOnStatus", pd.Series(dtype=object)).fillna("Unknown").astype(str)
    powered_on = int(power.str.contains("on", case=False, na=False).sum())
    powered_off = int(power.str.contains("off", case=False, na=False).sum())

    fileshares = discovery[discovery["resourceType"].eq(FILESHARE_RESOURCE_TYPE)].copy()
    fs_total = len(fileshares)

    machine_os = {}
    if "resourceName" in machines.columns and "osType" in machines.columns:
        for _, row in machines.iterrows():
            machine_os[norm_server(row.get("resourceName"))] = str(row.get("osType", "")).strip()
    os_counts = {"Windows": 0, "RHEL": 0, "Other/Unknown": 0}
    for _, row in fileshares.iterrows():
        os_type = machine_os.get(norm_server(row.get("parentResourceName")), "")
        if "win" in os_type.lower():
            os_counts["Windows"] += 1
        elif "linux" in os_type.lower() or "rhel" in os_type.lower() or "red" in os_type.lower():
            os_counts["RHEL"] += 1
        else:
            os_counts["Other/Unknown"] += 1

    db_counts = {
        "SQL Server": int(discovery["resourceType"].eq("microsoft.offazure/mastersites/sqlsites/sqlservers").sum()),
        "MongoDB": int(discovery["resourceType"].eq("microsoft.applicationmigration/mongosites/mongoinstances").sum()),
        "MySQL": int(discovery["resourceType"].eq("microsoft.mysqldiscovery/mysqlsites/mysqlservers").sum()),
        "PostgreSQL": int(discovery["resourceType"].eq("microsoft.applicationmigration/pgsqlsites/pgsqlinstances").sum()),
    }
    db_total = sum(db_counts.values())

    sql_rows = discovery[discovery["resourceType"].eq("microsoft.offazure/mastersites/sqlsites/sqlservers")]
    sql_instance_count = int(len(sql_rows))
    sql_server_count = int(sql_rows["parentResourceName"].dropna().astype(str).str.strip().replace("", pd.NA).dropna().nunique())

    web_counts = {
        "IIS Web Applications": int(discovery["resourceType"].eq("microsoft.offazure/mastersites/webappsites/iiswebapplications").sum()),
        "Tomcat Web Applications": int(discovery["resourceType"].eq("microsoft.offazure/mastersites/webappsites/tomcatwebapplications").sum()),
    }
    web_total = sum(web_counts.values())

    onprem_storage_gb = pd.to_numeric(lift["ONPREM_STORAGE_GB"], errors="coerce").fillna(0).sum()
    rec_storage_gb = pd.to_numeric(lift["RECOMMENDED_STORAGE_SIZE_GB"], errors="coerce").fillna(0).sum()
    onprem_cores = pd.to_numeric(lift["ONPREM_CORES_COUNT"], errors="coerce").fillna(0).sum()
    rec_cores = pd.to_numeric(lift["RECOMMENDED_NUMBER_OF_CORES"], errors="coerce").fillna(0).sum()
    onprem_mem_gb = pd.to_numeric(lift["ONPREM_MEMORY_MB"], errors="coerce").fillna(0).sum() / 1024
    if "Recommended Memory" not in lift.columns:
        lift["Recommended Memory"] = lift["RECOMMENDED_COMPUTE_SKU"].map(recommended_memory_gb)
    else:
        rec_mem = pd.to_numeric(lift["Recommended Memory"], errors="coerce")
        missing = rec_mem.isna()
        if missing.any():
            derived_mem = pd.to_numeric(
                lift.loc[missing, "RECOMMENDED_COMPUTE_SKU"].map(recommended_memory_gb),
                errors="coerce",
            )
            rec_mem.loc[missing] = derived_mem
        lift["Recommended Memory"] = rec_mem
    rec_mem_gb = pd.to_numeric(lift["Recommended Memory"], errors="coerce").fillna(0).sum()
    storage_util = pd.to_numeric(lift["STORAGE_UTILIZATION_PERCENT"], errors="coerce").dropna().mean()
    cpu_util = pd.to_numeric(lift["ONPREM_CPU_USAGE_PERCENT"], errors="coerce").dropna().mean()
    mem_util = pd.to_numeric(lift["ONPREM_MEMORY_USAGE_PERCENT"], errors="coerce").dropna().mean()

    fs["server_norm"] = fs["SERVER"].map(norm_server)
    fs_only_rows = fs[fs["server_norm"].isin(fileshare_only)].copy()
    fs_only_size_gib = pd.to_numeric(fs_only_rows["AZURE FILES TARGET PROVISIONED SIZE (GIB)"], errors="coerce").fillna(0).sum()
    fs_only_cost = pd.to_numeric(fs_only_rows["ESTIMATED MONTHLY COST (USD)"], errors="coerce").fillna(0).sum()
    lift["server_norm"] = lift["SERVER_NAME"].map(norm_server)
    lift_only = lift[lift["server_norm"].isin(fileshare_only)]
    lift_only_cost = pd.to_numeric(lift_only["TOTAL_MONTHLY_COST_USD"], errors="coerce").fillna(0).sum()

    protocol_counts = fs["PROTOCOL"].fillna("Unknown").astype(str).str.strip().replace("", "Unknown").value_counts().to_dict()
    readiness_counts = fs["READINESS"].fillna("Blank").astype(str).str.strip().replace("", "Blank").value_counts().to_dict()

    return {
        "vm_total": vm_total,
        "powered_on": powered_on,
        "powered_off": powered_off,
        "fileshare_total": fs_total,
        "fileshare_servers": len(fileshare_servers),
        "fileshare_only_servers": len(fileshare_only),
        "fileshare_only_shares": len(fs_only_rows),
        "fileshare_only_size_tb": fs_only_size_gib / 1024,
        "fileshare_only_azure_cost": fs_only_cost,
        "fileshare_only_lift_cost": lift_only_cost,
        "protocol_counts": protocol_counts,
        "readiness_counts": readiness_counts,
        "os_counts": os_counts,
        "db_counts": db_counts,
        "db_total": db_total,
        "web_counts": web_counts,
        "web_total": web_total,
        "sql_instance_count": sql_instance_count,
        "sql_server_count": sql_server_count,
        "onprem_storage_gb": onprem_storage_gb,
        "rec_storage_gb": rec_storage_gb,
        "onprem_cores": onprem_cores,
        "rec_cores": rec_cores,
        "onprem_mem_gb": onprem_mem_gb,
        "rec_mem_gb": rec_mem_gb,
        "storage_util": storage_util,
        "cpu_util": cpu_util,
        "mem_util": mem_util,
    }


def new_deck():
    prs = Presentation()
    prs.slide_width = Inches(SLIDE_W)
    prs.slide_height = Inches(SLIDE_H)
    return prs


def blank_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = BG
    return slide


def slide_title(slide, title, subtitle=""):
    add_text(slide, title, 0.5, 0.35, 8.7, 0.55, 32, NAVY, True, "Aptos Display")
    if subtitle:
        add_text(slide, subtitle, 0.52, 0.92, 9.0, 0.3, 12, MUTED)


def add_source(slide, text, size=9):
    add_text(slide, text, 0.6, 6.92, 12.1, 0.18, size, MUTED)


def add_bar(slide, x, y, w, h, value, max_value, color):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    bg.fill.solid()
    bg.fill.fore_color.rgb = ALT
    bg.line.fill.background()
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w * safe_div(value, max_value)), Inches(h))
    bar.fill.solid()
    bar.fill.fore_color.rgb = color
    bar.line.fill.background()
    return bar


def add_vm_power_slide(prs, m):
    slide = blank_slide(prs)
    slide_title(slide, "VM Power State Summary")

    max_count = max(m["powered_on"], m["powered_off"], 1)
    chart_x, chart_y, chart_w, chart_h = 1.25, 1.55, 10.9, 1.2
    tick_max = int(math.ceil(max_count / 250.0) * 250) + 250
    for tick in range(0, tick_max + 1, 250):
        x = chart_x + chart_w * safe_div(tick, tick_max)
        line = slide.shapes.add_shape(MSO_SHAPE.LINE_INVERSE, Inches(x), Inches(chart_y), Inches(0), Inches(chart_h))
        line.line.color.rgb = LINE
        line.line.width = Pt(0.5)
        add_text(slide, f"{tick:,}", x - 0.25, chart_y + chart_h + 0.18, 0.5, 0.15, 9, MUTED, align=PP_ALIGN.CENTER)
    add_bar(slide, chart_x, chart_y + 0.2, chart_w, 0.28, m["powered_on"], max_count, TEAL)
    add_bar(slide, chart_x, chart_y + 0.75, chart_w * safe_div(m["powered_off"], max_count), 0.28, m["powered_off"], m["powered_off"] or 1, PINK)
    add_text(slide, f"Powered On ({m['powered_on']:,})", 4.75, 3.03, 1.8, 0.17, 10, SLATE, True, align=PP_ALIGN.CENTER)
    add_text(slide, f"Powered Off ({m['powered_off']:,})", 6.52, 3.03, 1.85, 0.17, 10, SLATE, True, align=PP_ALIGN.CENTER)

    add_panel(slide, 0.75, 3.92, 11.1, 2.25, "VM power-state counts", title_size=17)
    rows = [
        ["Power state", "VM count", "Share of VMs"],
        ["Powered On", f"{m['powered_on']:,}", pct(m["powered_on"], m["vm_total"])],
        ["Powered Off", f"{m['powered_off']:,}", pct(m["powered_off"], m["vm_total"])],
        ["Total VMs", f"{m['vm_total']:,}", "100.0%"],
    ]
    add_table(slide, rows, 0.95, 4.58, 6.0, 1.46, 12, col_widths=[2.5, 1.8, 1.7])
    add_card(slide, 7.55, 4.52, 1.95, 1.35, f"{m['powered_on']:,}", "Powered On", "", TEAL, 26, label_size=10)
    add_card(slide, 9.65, 4.52, 1.95, 1.35, f"{m['powered_off']:,}", "Powered Off", "", PINK, 26, label_size=10)
    add_source(slide, "Source: Discovery.xlsx, Data sheet, resourceType = microsoft.offazure/vmwaresites/machines.", size=9)


def add_vm_utilization_slide(prs, m):
    slide = blank_slide(prs)
    slide_title(slide, "VM Utilization Summary", "On-premises footprint compared with recommended Azure sizing; lower is better for capacity change.")
    add_text(slide, "On-premises footprint", 2.85, 1.28, 5.83, 0.3, 17, NAVY, True)
    card_w, card_h = 2.08, 1.55
    add_card(slide, 0.55, 1.55, card_w, card_h, f"{m['vm_total']:,}", "Assessed VMs", "", TEAL, 26)
    add_card(slide, 2.85, 1.55, card_w, card_h, pb_from_gb(m["onprem_storage_gb"]), "Storage", f"{m['storage_util']:.1f}% Consumed", TEAL, 26)
    add_card(slide, 5.15, 1.55, card_w, card_h, f"{m['onprem_cores']:,.0f}", "Cores", f"{m['cpu_util']:.1f}% CPU Utilization", MINT, 26)
    add_card(slide, 7.45, 1.55, card_w, card_h, tb_from_gib(m["onprem_mem_gb"]), "Memory", f"{m['mem_util']:.1f}% Memory Utilization", MINT, 26)
    add_text(slide, "Recommended Azure sizing", 2.85, 3.96, 5.83, 0.3, 17, NAVY, True)
    storage_delta = (m["rec_storage_gb"] - m["onprem_storage_gb"]) / m["onprem_storage_gb"] * 100
    cores_delta = (m["rec_cores"] - m["onprem_cores"]) / m["onprem_cores"] * 100
    mem_delta = (m["rec_mem_gb"] - m["onprem_mem_gb"]) / m["onprem_mem_gb"] * 100
    add_card(slide, 2.85, 4.25, card_w, card_h, pb_from_gb(m["rec_storage_gb"]), "Storage", f"{storage_delta:+.1f}% vs on-prem", TEAL, 26)
    add_card(slide, 5.15, 4.25, card_w, card_h, f"{m['rec_cores']:,.0f}", "Cores", f"{cores_delta:+.1f}% vs on-prem", MINT, 26)
    add_card(slide, 7.45, 4.25, card_w, card_h, tb_from_gib(m["rec_mem_gb"]), "Memory", f"{mem_delta:+.1f}% vs on-prem", MINT, 26)

    # Comparison bar chart: on-prem vs recommended Azure sizing, per metric.
    panel_x, panel_y, panel_w, panel_h = 9.75, 1.55, 3.4, 4.25
    add_panel(slide, panel_x, panel_y, panel_w, panel_h, "On-prem vs Azure", title_size=13)
    legend_y = panel_y + 0.55
    sw_x = panel_x + 0.2
    sw = 0.18
    s1 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(sw_x), Inches(legend_y), Inches(sw), Inches(0.12))
    s1.fill.solid(); s1.fill.fore_color.rgb = TEAL; s1.line.fill.background()
    add_text(slide, "On-prem", sw_x + sw + 0.05, legend_y - 0.04, 0.8, 0.2, 9, SLATE)
    s2 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(sw_x + 1.25), Inches(legend_y), Inches(sw), Inches(0.12))
    s2.fill.solid(); s2.fill.fore_color.rgb = MINT; s2.line.fill.background()
    add_text(slide, "Azure", sw_x + 1.25 + sw + 0.05, legend_y - 0.04, 0.8, 0.2, 9, SLATE)

    metrics = [
        ("Storage", pb_from_gb(m["onprem_storage_gb"]), pb_from_gb(m["rec_storage_gb"]), m["onprem_storage_gb"], m["rec_storage_gb"]),
        ("Cores",   f"{m['onprem_cores']:,.0f}",        f"{m['rec_cores']:,.0f}",        m["onprem_cores"],     m["rec_cores"]),
        ("Memory",  tb_from_gib(m["onprem_mem_gb"]),    tb_from_gib(m["rec_mem_gb"]),    m["onprem_mem_gb"],    m["rec_mem_gb"]),
    ]
    bars_x = panel_x + 0.2
    bars_w = 1.8
    label_w = panel_w - 0.4 - bars_w - 0.05
    group_top = panel_y + 1.05
    group_h = (panel_h - 1.2) / len(metrics)
    for i, (label, on_text, az_text, on_val, az_val) in enumerate(metrics):
        gy = group_top + i * group_h
        add_text(slide, label, bars_x, gy, panel_w - 0.4, 0.22, 11, NAVY, True)
        bar_max = max(on_val, az_val, 1)
        b1y = gy + 0.32
        b2y = b1y + 0.34
        add_bar(slide, bars_x, b1y, bars_w, 0.22, on_val, bar_max, TEAL)
        add_text(slide, on_text, bars_x + bars_w + 0.05, b1y - 0.03, label_w, 0.25, 9, NAVY, True)
        add_bar(slide, bars_x, b2y, bars_w, 0.22, az_val, bar_max, MINT)
        add_text(slide, az_text, bars_x + bars_w + 0.05, b2y - 0.03, label_w, 0.25, 9, NAVY, True)

    add_source(
        slide,
        "Source: Strategy_Lift_and_shift.xlsx, Server_to_AzureVM tab. Recommended storage, cores, and memory are compared to on-premises totals.",
        size=7,
    )


def add_fileshare_os_slide(prs, m):
    slide = blank_slide(prs)
    slide_title(slide, "Fileshares by Host OS Category")
    windows = m["os_counts"].get("Windows", 0)
    rhel = m["os_counts"].get("RHEL", 0)
    other = m["os_counts"].get("Other/Unknown", 0)
    breakdown = [("Windows", windows), ("Linux / RHEL", rhel), ("Other / Unknown", other)]
    add_card(slide, 0.85, 1.25, 2.35, 1.35, f"{m['fileshare_total']:,}", "Total fileshares", "", TEAL, value_size=34, label_size=11)
    max_count = max((c for _, c in breakdown), default=0) or 1
    for i, (label, count) in enumerate(breakdown):
        y = 1.52 + i * 0.32
        add_text(slide, label, 3.35, y, 1.4, 0.18, 10, SLATE)
        add_bar(slide, 4.85, y + 0.03, 6.80, 0.15, count, max_count, TEAL)
        add_text(slide, f"{count:,}", 11.80, y, 0.5, 0.18, 10, NAVY, True, align=PP_ALIGN.RIGHT)
    rows = [["Host OS type", "OS category", "Fileshares", "Share"]]
    for label, count in [("Windows", windows), ("Linux", rhel), ("Other/Unknown", other)]:
        if count:
            rows.append([label, "RHEL" if label == "Linux" else label, f"{count:,}", pct(count, m["fileshare_total"])])
    rows.append(["Total", "", f"{m['fileshare_total']:,}", "100.0%"])
    add_panel(slide, 1.8, 3.25, 9.2, 2.6, "Fileshare count by host operating system", title_size=17)
    add_table(slide, rows, 2.0, 3.92, 8.8, 1.8, 12, col_widths=[2.5, 2.3, 2.0, 2.0])
    add_source(
        slide,
        "Source: Discovery.xlsx, Data sheet. Fileshares are resourceType = microsoft.applicationmigration/storagesites/fileshares; OS is joined from host machine records.",
        size=8,
    )


def add_db_slide(prs, m):
    slide = blank_slide(prs)
    slide_title(slide, "Database Resources by Type")
    add_card(slide, 0.85, 1.25, 2.35, 1.35, f"{m['db_total']:,}", "Total database resources", "", TEAL, value_size=34, label_size=11)
    max_count = max(m["db_counts"].values()) or 1
    for i, (label, count) in enumerate(m["db_counts"].items()):
        y = 1.52 + i * 0.25
        add_text(slide, label, 3.35, y, 1.2, 0.15, 9, SLATE)
        add_bar(slide, 4.7, y + 0.02, 6.95, 0.12, count, max_count, TEAL)
        add_text(slide, f"{count:,}", 11.85, y, 0.45, 0.15, 9, NAVY, True, align=PP_ALIGN.RIGHT)
    rows = [["Database type", "Count", "Share"]]
    for k, v in m["db_counts"].items():
        rows.append([k, f"{v:,}", pct(v, m["db_total"])])
    rows.append(["Total", f"{m['db_total']:,}", "100.0%"])
    add_panel(slide, 1.8, 3.25, 9.2, 2.6, "Database count by type", title_size=17)
    add_table(slide, rows, 2.0, 3.92, 8.8, 1.8, 12, col_widths=[4.6, 2.0, 2.0])
    add_source(slide, "Source: Discovery.xlsx, Data sheet. Database rows identified from database-related values in resourceType.", size=8)


def add_webapp_slide(prs, m):
    slide = blank_slide(prs)
    slide_title(slide, "Web Applications by Type", "Counted by webapp rows")
    add_card(slide, 0.85, 1.25, 2.35, 1.35, f"{m['web_total']:,}", "Total webapps", "", TEAL, value_size=34, label_size=11)
    max_count = max(m["web_counts"].values()) or 1
    for i, (label, count) in enumerate(m["web_counts"].items()):
        y = 1.52 + i * 0.28
        add_text(slide, label, 3.35, y, 1.7, 0.18, 10, SLATE)
        add_bar(slide, 5.15, y + 0.03, 6.50, 0.15, count, max_count, TEAL)
        add_text(slide, f"{count:,}", 11.80, y, 0.5, 0.18, 10, NAVY, True, align=PP_ALIGN.RIGHT)
    rows = [["Webapp type", "Webapp rows", "Share"]]
    for k, v in m["web_counts"].items():
        rows.append([k, f"{v:,}", pct(v, m["web_total"])])
    rows.append(["Total", f"{m['web_total']:,}", "100.0%"])
    add_panel(slide, 1.8, 3.25, 9.2, 2.6, "Webapp count by type", title_size=17)
    add_table(slide, rows, 2.0, 3.92, 8.8, 1.8, 12, col_widths=[4.6, 2.0, 2.0])
    add_source(
        slide,
        "Source: Discovery.xlsx, Data sheet. Webapp rows identified from resourceType; duplicate parentResourceName values are counted separately.",
        size=8,
    )


def add_consolidated_slide(prs, m):
    slide = blank_slide(prs)
    slide_title(slide, "Consolidated Infrastructure Summary")
    add_panel(slide, 0.4, 1.1, 2.6, 2.0, "VM Power State Chart", title_size=14)
    max_count = max(m["powered_on"], m["powered_off"], 1)
    for i, (label, count, color) in enumerate([("Powered On", m["powered_on"], TEAL), ("Powered Off", m["powered_off"], PINK)]):
        y = 1.6 + i * 0.66
        add_text(slide, label, 0.6, y, 1.0, 0.22, 8, SLATE, True)
        add_bar(slide, 1.55, y + 0.03, 1.1, 0.15, count, max_count, color)
        add_text(slide, f"{count:,} ({pct(count, m['vm_total'])})", 0.6, y + 0.26, 2.15, 0.3, 10, NAVY, True)

    add_panel(slide, 3.15, 1.1, 9.4, 2.0, "VM Utilization Summary", title_size=14)
    storage_delta = delta_pct(m["rec_storage_gb"], m["onprem_storage_gb"])
    cores_delta = delta_pct(m["rec_cores"], m["onprem_cores"])
    mem_delta = delta_pct(m["rec_mem_gb"], m["onprem_mem_gb"])
    rows = [
        ["", "On-premises total", "On-premises Utilization", "Recommended Azure sizing"],
        ["Assessed VMs", f"{m['vm_total']:,}", "", ""],
        ["Storage", pb_from_gb(m["onprem_storage_gb"]), f"{m['storage_util']:.1f}% Consumed", f"{pb_from_gb(m['rec_storage_gb'])} ({delta_text(storage_delta)})"],
        ["Cores", f"{m['onprem_cores']:,.0f}", f"{m['cpu_util']:.1f}% CPU Utilization", f"{m['rec_cores']:,.0f} ({delta_text(cores_delta)})"],
        ["Memory", tb_from_gib(m["onprem_mem_gb"]), f"{m['mem_util']:.1f}% Memory Utilization", f"{tb_from_gib(m['rec_mem_gb'])} ({delta_text(mem_delta)})"],
    ]
    tbl_shape = add_table(slide, rows, 3.42, 1.56, 9.33, 1.25, 11, col_widths=[1.3, 2.1, 2.55, 3.05])
    table = tbl_shape.table
    for r in (2, 3, 4):
        color_delta_text(table.cell(r, 3), table.cell(r, 3).text)

    add_panel(slide, 0.4, 3.45, 3.9, 2.4, "Fileshare count by host OS", title_size=14)
    windows = m["os_counts"].get("Windows", 0)
    rhel = m["os_counts"].get("RHEL", 0)
    other = m["os_counts"].get("Other/Unknown", 0)
    fs_rows = [["Host OS type", "OS category", "Fileshares", "Share"], ["Windows", "Windows", f"{windows:,}", pct(windows, m["fileshare_total"])]]
    if rhel:
        fs_rows.append(["Linux", "RHEL", f"{rhel:,}", pct(rhel, m["fileshare_total"])])
    if other:
        fs_rows.append(["Other", "Unknown", f"{other:,}", pct(other, m["fileshare_total"])])
    fs_rows.append(["Total", "", f"{m['fileshare_total']:,}", "100.0%"])
    add_table(slide, fs_rows, 0.55, 4.05, 3.6, 1.6, 11, col_widths=[1.1, 1.0, 0.85, 0.65])

    add_panel(slide, 4.55, 3.45, 3.75, 2.4, "Database count by type", title_size=14)
    db_rows = [["Database type", "Count", "Share"]]
    for k, v in m["db_counts"].items():
        db_rows.append([k, f"{v:,}", pct(v, m["db_total"])])
    db_rows.append(["Total", f"{m['db_total']:,}", "100.0%"])
    add_table(slide, db_rows, 4.7, 4.05, 3.45, 1.7, 11, col_widths=[1.85, 0.85, 0.75])

    add_panel(slide, 8.7, 3.45, 3.85, 2.4, "Webapp count by type", title_size=14)
    web_rows = [["Webapp type", "Webapp", "Share"]]
    for k, v in m["web_counts"].items():
        web_rows.append([k, f"{v:,}", pct(v, m["web_total"])])
    web_rows.append(["Total", f"{m['web_total']:,}", "100.0%"])
    add_table(slide, web_rows, 8.85, 4.05, 3.55, 1.6, 11, col_widths=[1.95, 0.85, 0.75])


def add_fileshare_readiness_slide(prs, m, output_dir: Path):
    slide = blank_slide(prs)
    slide_title(slide, "Fileshare Readiness", "Discovery-derived fileshare-only server scope with Azure Files readiness, sizing, and monthly cost comparison.")
    add_card(slide, 0.52, 1.55, 2.35, 1.35, f"{m['fileshare_only_servers']:,}", "Fileshare-only servers", "", TEAL, value_size=20, label_size=10)
    add_card(slide, 3.02, 1.55, 2.35, 1.35, f"{m['fileshare_only_shares']:,}", "Fileshare-only shares", "", MINT, value_size=20, label_size=10)

    add_panel(slide, 0.55, 3.25, 5.1, 3.15, "", title_size=1)
    donut = create_donut_image(output_dir / "protocol-distribution.png", m["protocol_counts"].get("SMB", 0), m["protocol_counts"].get("NFS", 0))
    slide.shapes.add_picture(str(donut), Inches(0.72), Inches(3.5), Inches(4.58), Inches(2.85))

    add_panel(slide, 6.05, 1.55, 6.3, 1.75, "Fileshare-only Azure sizing and cost", title_size=16)
    add_small_card(slide, 6.17, 2.10, 1.42, 1.02, f"{m['fileshare_only_size_tb']:.1f} TB", "Azure Files size", TEAL, value_size=15, label_size=8)
    add_small_card(slide, 7.70, 2.10, 1.42, 1.02, money_k(m["fileshare_only_azure_cost"]), "Azure Files / mo", GREEN, value_size=15, label_size=8)
    add_small_card(slide, 9.23, 2.10, 1.42, 1.02, money_k(m["fileshare_only_lift_cost"]), "Lift & shift / mo", RED, value_size=15, label_size=8)
    delta_value = m["fileshare_only_lift_cost"] - m["fileshare_only_azure_cost"]
    delta_color = GREEN if delta_value >= 0 else RED
    add_small_card(slide, 10.75, 2.10, 1.34, 1.02, money_k(delta_value), "Monthly delta", delta_color, value_size=15, label_size=8)

    add_panel(slide, 6.05, 3.25, 6.3, 3.15, "Azure Files readiness distribution", title_size=16)
    readiness = create_readiness_image(output_dir / "azure-files-readiness-distribution.png", m["readiness_counts"])
    slide.shapes.add_picture(str(readiness), Inches(6.04), Inches(3.97), Inches(6.25), Inches(2.64))
    add_text(
        slide,
        "Sources: Discovery.xlsx Data sheet; Strategy_PaaS_Preferred.xlsx FileShares_to_Azure_Files; Strategy_Lift_and_shift.xlsx Server_to_AzureVM.",
        0.53,
        7.01,
        10.97,
        0.3,
        7,
        MUTED,
    )


def add_sql_readiness_slide(prs, m):
    slide = blank_slide(prs)
    slide_title(slide, "SQL Readiness")
    add_card(slide, 0.85, 1.25, 2.55, 1.55, f"{m['sql_instance_count']:,}", "SQL Server Instances", "", TEAL, value_size=40, label_size=12)
    add_card(slide, 3.65, 1.25, 2.55, 1.55, f"{m['sql_server_count']:,}", "Servers running SQL", "", MINT, value_size=40, label_size=12)

    instances_per_server = m["sql_instance_count"] / m["sql_server_count"] if m["sql_server_count"] else 0
    add_panel(slide, 6.45, 1.25, 6.0, 1.55, "Density", title_size=14)
    add_text(slide, f"{instances_per_server:.1f}", 6.65, 1.75, 2.0, 0.7, 40, NAVY, True)
    add_text(slide, "Average SQL Server instances per host", 8.8, 1.95, 3.5, 0.6, 11, SLATE, True)

    rows = [
        ["Metric", "Count"],
        ["SQL Server Instances", f"{m['sql_instance_count']:,}"],
        ["Unique servers running SQL", f"{m['sql_server_count']:,}"],
    ]
    add_panel(slide, 1.8, 3.25, 9.2, 2.6, "SQL Server inventory", title_size=17)
    add_table(slide, rows, 2.0, 3.92, 8.8, 1.8, 12, col_widths=[5.6, 3.2])
    add_source(
        slide,
        "Source: Discovery.xlsx, Data sheet. SQL Server Instances = rows where resourceType = microsoft.offazure/mastersites/sqlsites/sqlservers. "
        "Servers running SQL = distinct parentResourceName values for those same rows.",
        size=8,
    )
    slide = blank_slide(prs)
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = NAVY

    # Accent stripe
    stripe = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(5.05), Inches(SLIDE_W), Inches(0.10))
    stripe.fill.solid(); stripe.fill.fore_color.rgb = TEAL; stripe.line.fill.background()

    add_text(slide, "Azure Migrate Interpreted", 0.6, 2.6, SLIDE_W - 1.2, 1.3, 60, WHITE, True, font="Aptos Display", align=PP_ALIGN.CENTER)
    add_text(slide, "Discovery and assessment summary", 0.6, 4.05, SLIDE_W - 1.2, 0.5, 20, RGBColor(202, 220, 252), False, align=PP_ALIGN.CENTER)


def build_deck(input_dir: Path, output: Path):
    log(f"Input  : {input_dir}")
    log(f"Output : {output}")
    output.parent.mkdir(parents=True, exist_ok=True)
    log("Loading source workbooks...")
    m = load_metrics(input_dir)
    log("Source data loaded:")
    log(f"  VMs            : {m['vm_total']:,} ({m['powered_on']:,} on / {m['powered_off']:,} off)")
    log(f"  Fileshares     : {m['fileshare_total']:,} ({m['fileshare_only_servers']:,} fileshare-only servers / {m['fileshare_only_shares']:,} shares)")
    log(f"  Databases      : {m['db_total']:,}")
    log(f"  Web apps       : {m['web_total']:,}")
    log("Building slides...")
    prs = new_deck()
    log("  [1/9] Title")
    add_title_slide(prs)
    log("  [2/9] Consolidated Infrastructure Summary")
    add_consolidated_slide(prs, m)
    log("  [3/9] Fileshare Readiness")
    add_fileshare_readiness_slide(prs, m, output.parent)
    log("  [4/9] VM Power State Summary")
    add_vm_power_slide(prs, m)
    log("  [5/9] VM Utilization Summary")
    add_vm_utilization_slide(prs, m)
    log("  [6/9] Fileshares by Host OS Category")
    add_fileshare_os_slide(prs, m)
    log("  [7/9] Database Resources by Type")
    add_db_slide(prs, m)
    log("  [8/9] SQL Readiness")
    add_sql_readiness_slide(prs, m)
    log("  [9/9] Web Applications by Type")
    add_webapp_slide(prs, m)
    log("Saving deck...")
    prs.save(output)
    log(f"Done. Wrote {output}")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--input-dir", required=True, type=Path, help="Directory containing the three source workbooks.")
    parser.add_argument("--output", required=True, type=Path, help="Output .pptx path.")
    args = parser.parse_args()
    build_deck(args.input_dir, args.output)


if __name__ == "__main__":
    main()
