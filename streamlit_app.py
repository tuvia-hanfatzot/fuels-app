import streamlit as st
from io import BytesIO
from copy import copy
import re
import base64

from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
from openpyxl.cell.cell import MergedCell
from openpyxl.utils.cell import range_boundaries

from openpyxl.styles import PatternFill, Font
from openpyxl.chart import PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.styles import PatternFill, Font


# ============================================================
# Streamlit config + UI styling
# ============================================================
st.set_page_config(page_title="Total Distribution → Calculations", layout="centered")

st.markdown(
    """
<style>
    .block-container { padding-top: 2rem; max-width: 900px; }
    .title { font-size: 2rem; font-weight: 800; margin-bottom: .25rem; }
    .subtitle { color: #6b7280; margin-bottom: 1.25rem; }
    .card {
        border: 1px solid rgba(0,0,0,.08);
        border-radius: 16px;
        padding: 16px 18px;
        background: rgba(255,255,255,.6);
    }
    .small { font-size: .9rem; color: #6b7280; }
</style>
""",
    unsafe_allow_html=True,
)

st.markdown('<div class="title">Total Distribution Combiner + Fuels Cleaner</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="subtitle">Upload multiple Total Distribution files → combine first → then run your calculations → download final.</div>',
    unsafe_allow_html=True,
)


# ============================================================
# Settings
# ============================================================
TARGET_SHEETS = {
    "UNOPS": "UNOPS Total Distribution",
    "WFP": "WFP Total Distribution",
}

HEADER_ROWS = 2
DATA_START_ROW = 3

# If you still only want to freeze specific columns in source sheets you can keep this,
# but we now copy VALUES for *all* cells into the master anyway.
FREEZE_COLS = (3, 4)  # C, D  (kept for compatibility)

# After combining, delete columns F–I (6..9) before calculations
DELETE_COMBINED_COLS_F_TO_I_BEFORE_CALC = True


# ============================================================
# Small UI helper (optional loader video)
# ============================================================
def show_small_loader_video(placeholder, path: str, width_px: int = 220):
    try:
        with open(path, "rb") as f:
            b64 = base64.b64encode(f.read()).decode("utf-8")
        placeholder.markdown(
            f"""
            <div style="display:flex;justify-content:center;margin:12px 0;">
              <video width="{width_px}" autoplay loop muted playsinline>
                <source src="data:video/mp4;base64,{b64}" type="video/mp4">
              </video>
            </div>
            """,
            unsafe_allow_html=True,
        )
    except Exception:
        # If video missing, just ignore (no crash)
        placeholder.empty()


# ============================================================
# COMBINER HELPERS
# ============================================================
def _norm(s: str) -> str:
    s = (s or "").strip().lower()
    s = re.sub(r"\s+", " ", s)
    return s


def _find_target_sheet(wb):
    sheetnames = list(wb.sheetnames)
    norm_map = {name: _norm(name) for name in sheetnames}

    targets = {
        "UNOPS": _norm("UNOPS Total Distribution"),
        "WFP": _norm("WFP Total Distribution"),
    }

    # 1) Normalised exact match
    for kind, target_norm in targets.items():
        for real_name, real_norm in norm_map.items():
            if real_norm == target_norm:
                return kind, real_name

    # 2) Broader match
    for real_name, real_norm in norm_map.items():
        if "total" in real_norm and "distribution" in real_norm:
            if "unops" in real_norm:
                return "UNOPS", real_name
            if "wfp" in real_norm:
                return "WFP", real_name

    return None, None


def _row_has_data_row(ws, r, max_col):
    for c in range(1, max_col + 1):
        if ws.cell(r, c).value not in (None, ""):
            return True
    return False


def _first_data_row(ws, start=DATA_START_ROW, max_scan=200):
    end = min(ws.max_row, start + max_scan)
    for r in range(start, end + 1):
        if _row_has_data_row(ws, r, ws.max_column):
            return r
    return start


def _last_data_row(ws, start_row, max_col):
    for r in range(ws.max_row, start_row - 1, -1):
        if _row_has_data_row(ws, r, max_col):
            return r
    return start_row - 1


def _delete_last_text_row_in_sheet(ws, start_row, end_row, max_col):
    """
    Delete the last row that contains any value (end_row already points to last data row).
    Returns new end_row.
    """
    if end_row < start_row:
        return end_row
    ws.delete_rows(end_row, 1)
    return end_row - 1


def _copy_dimensions(src_ws, dst_ws, max_col, header_rows):
    # Column widths
    for c in range(1, max_col + 1):
        L = get_column_letter(c)
        if L in src_ws.column_dimensions:
            w = src_ws.column_dimensions[L].width
            if w is not None:
                dst_ws.column_dimensions[L].width = w

    # Header row heights
    for r in range(1, header_rows + 1):
        if r in src_ws.row_dimensions and src_ws.row_dimensions[r].height is not None:
            dst_ws.row_dimensions[r].height = src_ws.row_dimensions[r].height


def _safe_value(ws_vals, r, c):
    try:
        return ws_vals.cell(r, c).value
    except Exception:
        return None


def _copy_block_values_only(src_ws, src_ws_vals, dst_ws, src_row_start, src_row_end, dst_row_start, max_col):
    """
    Copy block while:
    - keeping formatting (styles, row heights, merges)
    - writing VALUES (not formulas) using src_ws_vals (data_only=True)
    """
    row_delta = dst_row_start - src_row_start

    for r in range(src_row_start, src_row_end + 1):
        if r in src_ws.row_dimensions and src_ws.row_dimensions[r].height is not None:
            dst_ws.row_dimensions[r + row_delta].height = src_ws.row_dimensions[r].height

        for c in range(1, max_col + 1):
            src_cell = src_ws.cell(r, c)
            dst_cell = dst_ws.cell(r + row_delta, c)

            # style
            if src_cell.has_style:
                dst_cell._style = copy(src_cell._style)
            dst_cell.number_format = src_cell.number_format
            dst_cell.protection = copy(src_cell.protection)
            dst_cell.alignment = copy(src_cell.alignment)
            dst_cell.font = copy(src_cell.font)
            dst_cell.border = copy(src_cell.border)
            dst_cell.fill = copy(src_cell.fill)

            # value (not formula)
            val = _safe_value(src_ws_vals, r, c)
            if val is None:
                val = src_cell.value
                if isinstance(val, str) and val.startswith("="):
                    val = None
            dst_cell.value = val

    # merges fully inside copied block
    for mr in list(src_ws.merged_cells.ranges):
        min_r, min_c, max_r, max_c = mr.min_row, mr.min_col, mr.max_row, mr.max_col
        if min_r >= src_row_start and max_r <= src_row_end and max_c <= max_col:
            dst_ws.merge_cells(
                start_row=min_r + row_delta,
                start_column=min_c,
                end_row=max_r + row_delta,
                end_column=max_c,
            )


def _merge_anchor(ws, r, c):
    for mr in ws.merged_cells.ranges:
        if mr.min_row <= r <= mr.max_row and mr.min_col <= c <= mr.max_col:
            return mr.min_row, mr.min_col
    return r, c


def _freeze_formulas_to_values(ws, ws_vals):
    """
    Kept (for safety) – but master copy is already values-only.
    This just reduces risk of weird formulas propagating inside each source ws before copy.
    """
    max_row = ws.max_row
    done = set()

    for r in range(DATA_START_ROW, max_row + 1):
        for c in FREEZE_COLS:
            ar, ac = _merge_anchor(ws, r, c)
            if (ar, ac) in done:
                continue

            target = ws.cell(ar, ac)
            if isinstance(target, MergedCell):
                continue

            cached = ws_vals.cell(ar, ac).value
            if cached is not None:
                target.value = cached

            done.add((ar, ac))


def _insert_wfp_column_a(ws, ws_vals, data_start, data_end):
    ws.insert_cols(1)
    ws_vals.insert_cols(1)

    style_source = ws.cell(data_start, 2)  # B
    for r in range(data_start, data_end + 1):
        has_data = False
        for c in range(2, ws.max_column + 1):
            if ws.cell(r, c).value not in (None, ""):
                has_data = True
                break
        if not has_data:
            continue

        a_cell = ws.cell(r, 1)
        a_cell.value = "WFP"

        if style_source is not None and style_source.has_style:
            a_cell._style = copy(style_source._style)
            a_cell.number_format = style_source.number_format
            a_cell.protection = copy(style_source.protection)
            a_cell.alignment = copy(style_source.alignment)
            a_cell.font = copy(style_source.font)
            a_cell.border = copy(style_source.border)
            a_cell.fill = copy(style_source.fill)


def _row_has_data_for_merge(ws, r, max_col):
    for c in range(2, max_col + 1):
        if ws.cell(r, c).value not in (None, ""):
            return True
    return False


def _merge_wfp_blocks_in_master(master_ws, max_col):
    max_row = master_ws.max_row

    # unmerge anything intersecting col A in data area
    to_unmerge = []
    for mr in list(master_ws.merged_cells.ranges):
        min_c, min_r, max_c, max_r = range_boundaries(str(mr))
        if min_c <= 1 <= max_c and max_r >= DATA_START_ROW:
            to_unmerge.append(str(mr))
    for rng in to_unmerge:
        master_ws.unmerge_cells(rng)

    r = DATA_START_ROW
    while r <= max_row:
        if master_ws.cell(r, 1).value == "WFP" and _row_has_data_for_merge(master_ws, r, max_col):
            start = r
            end = r
            r += 1
            while (
                r <= max_row
                and master_ws.cell(r, 1).value == "WFP"
                and _row_has_data_for_merge(master_ws, r, max_col)
            ):
                end = r
                r += 1

            if end > start:
                top = master_ws.cell(start, 1)
                top_style = copy(top._style)
                top_font = copy(top.font)
                top_alignment = copy(top.alignment)
                top_border = copy(top.border)
                top_fill = copy(top.fill)
                top_number_format = top.number_format
                top_protection = copy(top.protection)

                master_ws.merge_cells(start_row=start, start_column=1, end_row=end, end_column=1)

                top = master_ws.cell(start, 1)
                top.value = "WFP"
                top._style = top_style
                top.font = top_font
                top.alignment = top_alignment
                top.border = top_border
                top.fill = top_fill
                top.number_format = top_number_format
                top.protection = top_protection
        else:
            r += 1


def _clear_merges_on_cols(ws, cols=(1, 2), from_row=DATA_START_ROW):
    kept = []
    for mr in list(ws.merged_cells.ranges):
        min_c, min_r, max_c, max_r = range_boundaries(str(mr))

        if max_r < from_row:
            kept.append(mr)
            continue

        intersects_cols = any(min_c <= col <= max_c for col in cols)
        if not intersects_cols:
            kept.append(mr)

    ws.merged_cells.ranges = kept


def _merge_down_by_blanks(ws, col, max_col, from_row=DATA_START_ROW):
    """
    Merge downwards starting at a value cell, until next value appears,
    as long as the row is a real data row.
    """
    r = from_row
    while r <= ws.max_row:
        v = ws.cell(r, col).value
        if v in (None, "") or not _row_has_data_row(ws, r, max_col):
            r += 1
            continue

        start = r
        end = r
        rr = r + 1

        while rr <= ws.max_row:
            if not _row_has_data_row(ws, rr, max_col):
                break

            next_val = ws.cell(rr, col).value
            if next_val not in (None, ""):
                break

            end = rr
            rr += 1

        if end > start:
            top = ws.cell(start, col)
            top_style = copy(top._style)
            top_font = copy(top.font)
            top_alignment = copy(top.alignment)
            top_border = copy(top.border)
            top_fill = copy(top.fill)
            top_number_format = top.number_format
            top_protection = copy(top.protection)

            ws.merge_cells(start_row=start, start_column=col, end_row=end, end_column=col)

            top = ws.cell(start, col)
            top.value = v
            top._style = top_style
            top.font = top_font
            top.alignment = top_alignment
            top.border = top_border
            top.fill = top_fill
            top.number_format = top_number_format
            top.protection = top_protection

        r = end + 1


def _clear_merges_intersecting_cols(ws, col_start, col_end):
    kept = []
    for mr in list(ws.merged_cells.ranges):
        min_c, min_r, max_c, max_r = range_boundaries(str(mr))
        if not (max_c < col_start or min_c > col_end):
            continue
        kept.append(mr)
    ws.merged_cells.ranges = kept


def _delete_cols_safe(ws, col_start, col_end):
    _clear_merges_intersecting_cols(ws, col_start, col_end)
    for col in range(col_end, col_start - 1, -1):
        if col <= ws.max_column:
            ws.delete_cols(col, 1)


def build_combined_workbook_bytes(uploads, status=None):
    """
    Returns (combined_bytes_io, combined_workbook_object).
    """
    tasks = []

    # Load each uploaded workbook twice: normal + data_only cache
    for up in uploads:
        raw = BytesIO(up.getvalue())
        try:
            wb = load_workbook(raw, data_only=False)
            raw.seek(0)
            wb_vals = load_workbook(raw, data_only=True)
        except Exception as e:
            if status:
                status.warning(f"Skipped '{up.name}': could not read as .xlsx ({e})")
            continue

        kind, sheetname = _find_target_sheet(wb)
        if kind is None:
            if status:
                status.warning(f"Skipped '{up.name}': no matching Total Distribution sheet found. Sheets: {wb.sheetnames}")
            continue

        tasks.append({"name": up.name, "kind": kind, "sheet": sheetname, "wb": wb, "wb_vals": wb_vals})

    if not tasks:
        raise RuntimeError("No valid files found.")

    # UNOPS first, then WFP
    tasks.sort(key=lambda x: 0 if x["kind"] == "UNOPS" else 1)

    # Pre-pass: compute bounds + max cols
    max_cols = []
    bounds = []
    for t in tasks:
        ws = t["wb"][t["sheet"]]
        start_row = _first_data_row(ws, DATA_START_ROW)
        end_row = _last_data_row(ws, start_row, ws.max_column)
        bounds.append((start_row, end_row))
        max_cols.append(ws.max_column)

    master_max_col = max(max_cols) if max_cols else 1

    # Master workbook
    master_wb = Workbook()
    master_ws = master_wb.active
    master_ws.title = "Distribution Summary"

    wrote_header = False
    next_write_row = DATA_START_ROW

    for idx, t in enumerate(tasks):
        ws = t["wb"][t["sheet"]]
        ws_vals = t["wb_vals"][t["sheet"]]

        start_row, end_row = bounds[idx]
        if end_row < start_row:
            if status:
                status.warning(f"⚠️ {t['kind']} | {t['name']} has no data rows.")
            continue

        # WFP insert A + stamp WFP on data rows
        if t["kind"] == "WFP":
            _insert_wfp_column_a(ws, ws_vals, start_row, end_row)
            master_max_col = max(master_max_col, ws.max_column)

        # Recompute end_row after any structural change
        end_row = _last_data_row(ws, start_row, ws.max_column)
        if end_row < start_row:
            if status:
                status.warning(f"⚠️ {t['kind']} | {t['name']} has no data rows after WFP insert.")
            continue

        # Delete last text row BEFORE combining
        end_row = _delete_last_text_row_in_sheet(ws, start_row, end_row, ws.max_column)
        if end_row < start_row:
            if status:
                status.warning(f"⚠️ {t['kind']} | {t['name']} became empty after removing last text row.")
            continue

        # Freeze some formulas in-source (optional safety)
        _freeze_formulas_to_values(ws, ws_vals)

        # Copy header once
        if not wrote_header:
            _copy_dimensions(ws, master_ws, max_col=master_max_col, header_rows=HEADER_ROWS)
            _copy_block_values_only(ws, ws_vals, master_ws, 1, HEADER_ROWS, 1, max_col=master_max_col)
            wrote_header = True

        # Copy data values-only
        _copy_block_values_only(ws, ws_vals, master_ws, start_row, end_row, next_write_row, max_col=master_max_col)
        next_write_row += (end_row - start_row + 1)

    # Merge WFP blocks A
    _merge_wfp_blocks_in_master(master_ws, max_col=master_ws.max_column)

    # Rebuild merges A/B by blanks
    _clear_merges_on_cols(master_ws, cols=(1, 2), from_row=DATA_START_ROW)
    _merge_down_by_blanks(master_ws, col=1, max_col=master_ws.max_column, from_row=DATA_START_ROW)
    _merge_down_by_blanks(master_ws, col=2, max_col=master_ws.max_column, from_row=DATA_START_ROW)

    # Ensure visible
    for c in range(1, master_ws.max_column + 1):
        master_ws.column_dimensions[get_column_letter(c)].hidden = False
    for r in range(1, master_ws.max_row + 1):
        master_ws.row_dimensions[r].hidden = False
    master_ws.auto_filter.ref = None

    # Delete F–I on combined (if requested)
    if DELETE_COMBINED_COLS_F_TO_I_BEFORE_CALC:
        _delete_cols_safe(master_ws, 6, 9)
    
    out = BytesIO()
    master_wb.save(out)
    out.seek(0)
    return out


# ============================================================
# CALCULATIONS HELPERS (your cleaner logic, wrapped into a function)
# ============================================================
def copy_cell_style(src, dst):
    dst.font = copy(src.font)
    dst.fill = copy(src.fill)
    dst.border = copy(src.border)
    dst.alignment = copy(src.alignment)
    dst.number_format = src.number_format
    dst.protection = copy(src.protection)


def unmerge_and_fill(ws, col_min: int, col_max: int):
    merges = []
    for mr in ws.merged_cells.ranges:
        if not (mr.max_col < col_min or mr.min_col > col_max):
            merges.append(mr)

    for mr in merges:
        tl = ws.cell(row=mr.min_row, column=mr.min_col)
        value = tl.value
        style = {
            "font": copy(tl.font),
            "fill": copy(tl.fill),
            "border": copy(tl.border),
            "alignment": copy(tl.alignment),
            "number_format": tl.number_format,
            "protection": copy(tl.protection),
        }

        ws.unmerge_cells(str(mr))

        for r in range(mr.min_row, mr.max_row + 1):
            for c in range(mr.min_col, mr.max_col + 1):
                cell = ws.cell(row=r, column=c)
                cell.value = value
                cell.font = copy(style["font"])
                cell.fill = copy(style["fill"])
                cell.border = copy(style["border"])
                cell.alignment = copy(style["alignment"])
                cell.number_format = style["number_format"]
                cell.protection = copy(style["protection"])


def snapshot_row(ws, r):
    data = []
    for c in range(1, ws.max_column + 1):
        cell = ws.cell(r, c)
        data.append(
            {
                "value": cell.value,
                "font": copy(cell.font),
                "fill": copy(cell.fill),
                "border": copy(cell.border),
                "alignment": copy(cell.alignment),
                "number_format": cell.number_format,
                "protection": copy(cell.protection),
            }
        )
    return data


def restore_row(ws, r, row_data):
    for idx, d in enumerate(row_data, start=1):
        cell = ws.cell(r, idx)
        cell.value = d["value"]
        cell.font = d["font"]
        cell.fill = d["fill"]
        cell.border = d["border"]
        cell.alignment = d["alignment"]
        cell.number_format = d["number_format"]
        cell.protection = d["protection"]


def safe_float(v):
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        s = v.strip()
        if s == "":
            return 0.0
        s = s.replace(",", "")
        try:
            return float(s)
        except Exception:
            return 0.0
    return 0.0


def norm_header(x):
    return "" if x is None else " ".join(str(x).split()).strip().upper()


def rgb_fill(r, g, b):
    return PatternFill(fill_type="solid", fgColor=f"FF{r:02X}{g:02X}{b:02X}")

def style_headers_black_only_with_text(ws, header_row=1):
    header_fill = PatternFill(fill_type="solid", fgColor="000000")  # black
    header_font = Font(color="FFFFFF", bold=True)                  # white

    for c in range(1, ws.max_column + 1):
        cell = ws.cell(row=header_row, column=c)

        # Only style if header has real text
        if cell.value is not None and str(cell.value).strip() != "":
            cell.fill = header_fill
            cell.font = header_font

def _font_without_bold(f: Font) -> Font:
    if f is None:
        return Font(bold=False)
    return Font(
        name=f.name,
        sz=f.sz,
        b=False,
        i=f.italic,
        u=f.underline,
        strike=f.strike,
        color=f.color,
        vertAlign=f.vertAlign,
        outline=f.outline,
        shadow=f.shadow,
        condense=f.condense,
        extend=f.extend,
        charset=f.charset,
        family=f.family,
        scheme=f.scheme,
    )


def remove_bold_except_header(ws, header_row=1):
    """
    Remove bold from all cells except header row.
    Keeps other font properties intact.
    """
    max_r = ws.max_row
    max_c = ws.max_column
    for r in range(1, max_r + 1):
        if r == header_row:
            continue
        for c in range(1, max_c + 1):
            cell = ws.cell(row=r, column=c)
            f = cell.font
            if f is not None and f.bold:
                cell.font = _font_without_bold(f)

def run_calculations_on_combined_bytes(combined_bytes: BytesIO, progress=None, status=None) -> BytesIO:
    """
    Takes the combined workbook bytes (with sheet 'Master'),
    runs your calculations pipeline on it,
    returns final output BytesIO.
    """
    combined_bytes.seek(0)
    wb = load_workbook(combined_bytes, data_only=False)
    combined_bytes.seek(0)
    wb_cache = load_workbook(combined_bytes, data_only=True)

    # We already have only one sheet: "Master"
    TARGET_SHEET = "Distribution Summary"
    sheet_map = {name.strip(): name for name in wb.sheetnames}
    sheet_map_cache = {name.strip(): name for name in wb_cache.sheetnames}

    if TARGET_SHEET not in sheet_map or TARGET_SHEET not in sheet_map_cache:
        raise RuntimeError(f'Sheet "{TARGET_SHEET}" not found. Sheets found: {list(sheet_map.keys())}')

    keep_name = sheet_map[TARGET_SHEET]
    keep_name_cache = sheet_map_cache[TARGET_SHEET]

    # keep only Master in both (defensive)
    for name in list(wb.sheetnames):
        if name != keep_name:
            wb.remove(wb[name])
    for name in list(wb_cache.sheetnames):
        if name != keep_name_cache:
            wb_cache.remove(wb_cache[name])

    ws = wb[keep_name]
    ws_cache = wb_cache[keep_name_cache]
    ws.title = TARGET_SHEET

    if progress:
        progress.progress(10)
    if status:
        status.info("Unmerging A–C…")

    # Unmerge A–C
    unmerge_and_fill(ws, col_min=1, col_max=3)

    if progress:
        progress.progress(18)
    if status:
        status.info("Freezing formulas into values…")

    # Freeze any remaining formulas into values using cache
    max_row = ws.max_row
    max_col = ws.max_column
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            is_formula = (cell.data_type == "f") or (isinstance(cell.value, str) and cell.value.startswith("="))
            if is_formula:
                cell.value = ws_cache.cell(row=r, column=c).value

    if progress:
        progress.progress(26)

    # Delete row 2
    ws.delete_rows(2)

    # Delete rows containing 'TOTAL' in A,B,C (kept from your original)
    if status:
        status.info("Removing TOTAL rows…")

    rows_to_delete = []
    for row in range(1, ws.max_row + 1):
        for col in (1, 2, 3):
            val = ws.cell(row=row, column=col).value
            if val is not None and "TOTAL" in str(val).upper():
                rows_to_delete.append(row)
                break
    for rr in sorted(set(rows_to_delete), reverse=True):
        ws.delete_rows(rr)

    if progress:
        progress.progress(34)

    header_row = 1

    # Unmerge D–F before Fuel Sum
    unmerge_and_fill(ws, col_min=4, col_max=6)  # D..F

    # Add Fuel sum in F (values)
    if status:
        status.info("Building Fuel sum…")

    col_d, col_e, col_f = 4, 5, 6  # D,E,F

    ref_header = ws.cell(row=header_row, column=col_d)
    fuel_header = ws.cell(row=header_row, column=col_f)

    fuel_header.value = "Fuel sum"
    fuel_header.font = copy(ref_header.font)
    fuel_header.fill = copy(ref_header.fill)
    fuel_header.border = copy(ref_header.border)
    fuel_header.alignment = copy(ref_header.alignment)
    fuel_header.number_format = ref_header.number_format
    fuel_header.protection = copy(ref_header.protection)

    if "E" in ws.column_dimensions and ws.column_dimensions["E"].width is not None:
        ws.column_dimensions["F"].width = ws.column_dimensions["E"].width
    elif "D" in ws.column_dimensions and ws.column_dimensions["D"].width is not None:
        ws.column_dimensions["F"].width = ws.column_dimensions["D"].width

    for r in range(header_row + 1, ws.max_row + 1):
        d_cell = ws.cell(row=r, column=col_d)
        e_cell = ws.cell(row=r, column=col_e)
        f_cell = ws.cell(row=r, column=col_f)

        d_val = d_cell.value
        e_val = e_cell.value

        if d_val is None and e_val is None:
            f_cell.value = None
        else:
            try:
                d_num = 0.0 if d_val is None else float(d_val)
                e_num = 0.0 if e_val is None else float(e_val)
                f_cell.value = d_num + e_num
            except Exception:
                f_cell.value = None

        style_src = d_cell if d_cell.value is not None else e_cell
        f_cell.font = copy(style_src.font)
        f_cell.fill = copy(style_src.fill)
        f_cell.border = copy(style_src.border)
        f_cell.alignment = copy(style_src.alignment)
        f_cell.number_format = style_src.number_format
        f_cell.protection = copy(style_src.protection)

    if progress:
        progress.progress(45)

    # Delete rows where Fuel sum is 0 or empty
    if status:
        status.info("Removing empty/zero fuel rows…")

    rows_remove = []
    for r in range(header_row + 1, ws.max_row + 1):
        val = ws.cell(row=r, column=col_f).value
        if val is None or (isinstance(val, str) and val.strip() == ""):
            rows_remove.append(r)
            continue
        try:
            if float(val) == 0.0:
                rows_remove.append(r)
        except Exception:
            pass
    for rr in sorted(set(rows_remove), reverse=True):
        ws.delete_rows(rr)

    if progress:
        progress.progress(52)

    # Delete columns D and E
    ws.delete_cols(5)  # E
    ws.delete_cols(4)  # D
    # Fuel sum moved to D.

    if progress:
        progress.progress(56)

    # Insert Description Sum as column E
    if status:
        status.info("Building Description Sum…")

    desc_col = 5  # E
    ws.insert_cols(desc_col)

    ref_h = ws.cell(row=header_row, column=1)
    desc_h = ws.cell(row=header_row, column=desc_col)

    desc_h.value = "Description Sum"
    desc_h.font = copy(ref_h.font)
    desc_h.fill = copy(ref_h.fill)
    desc_h.border = copy(ref_h.border)
    desc_h.alignment = copy(ref_h.alignment)
    desc_h.number_format = ref_h.number_format
    desc_h.protection = copy(ref_h.protection)

    if "C" in ws.column_dimensions and ws.column_dimensions["C"].width is not None:
        ws.column_dimensions[get_column_letter(desc_col)].width = ws.column_dimensions["C"].width

    for r in range(header_row + 1, ws.max_row + 1):
        a = ws.cell(row=r, column=1).value
        b = ws.cell(row=r, column=2).value
        c = ws.cell(row=r, column=3).value

        a_s = "" if a is None else str(a)
        b_s = "" if b is None else str(b)
        c_s = "" if c is None else str(c)

        cell = ws.cell(row=r, column=desc_col)
        cell.value = f"{a_s},{b_s},{c_s}"

        src = ws.cell(row=r, column=1)
        cell.font = copy(src.font)
        cell.fill = copy(src.fill)
        cell.border = copy(src.border)
        cell.alignment = copy(src.alignment)
        cell.number_format = src.number_format
        cell.protection = copy(src.protection)

    if progress:
        progress.progress(64)

    # Unified Fuel as values
    if status:
        status.info("Building Unified Fuel…")

    unified_col = 6  # F
    ws.insert_cols(unified_col)

    ref_unified_header = ws.cell(row=header_row, column=4)  # D header
    unified_header = ws.cell(row=header_row, column=unified_col)
    unified_header.value = "Unified Fuel"
    unified_header.font = copy(ref_unified_header.font)
    unified_header.fill = copy(ref_unified_header.fill)
    unified_header.border = copy(ref_unified_header.border)
    unified_header.alignment = copy(ref_unified_header.alignment)
    unified_header.number_format = ref_unified_header.number_format
    unified_header.protection = copy(ref_unified_header.protection)

    d_letter = get_column_letter(4)
    f_letter = get_column_letter(unified_col)
    if d_letter in ws.column_dimensions and ws.column_dimensions[d_letter].width is not None:
        ws.column_dimensions[f_letter].width = ws.column_dimensions[d_letter].width

    totals = {}
    for r in range(header_row + 1, ws.max_row + 1):
        key = ws.cell(row=r, column=5).value  # E
        fuel = safe_float(ws.cell(row=r, column=4).value)  # D
        totals[key] = totals.get(key, 0.0) + fuel

    for r in range(header_row + 1, ws.max_row + 1):
        target = ws.cell(row=r, column=unified_col)  # F
        style_src = ws.cell(row=r, column=4)         # D
        key = ws.cell(row=r, column=5).value         # E

        target.value = totals.get(key, 0.0)
        target.font = copy(style_src.font)
        target.fill = copy(style_src.fill)
        target.border = copy(style_src.border)
        target.alignment = copy(style_src.alignment)
        target.number_format = style_src.number_format
        target.protection = copy(style_src.protection)

    if progress:
        progress.progress(72)

    # Sorting
    if status:
        status.info("Sorting…")

    # Find INTERVENTION + AGENCY columns
    intervention_col = None
    agency_col = None

    for c in range(1, ws.max_column + 1):
        h = ws.cell(row=1, column=c).value
        nh = norm_header(h)
        if nh == "INTERVENTION":
            intervention_col = c
        elif nh == "AGENCY":
            agency_col = c

    if intervention_col is None:
        intervention_col = 1

    if agency_col is None:
        raise RuntimeError('Header "AGENCY" not found.')

    # Normalise: fold UN-OHCHR into INGOs
    for r in range(2, ws.max_row + 1):
        v = ws.cell(row=r, column=intervention_col).value
        if v is None:
            continue
        t = str(v).strip()
        if "UN-OHCHR" in t.upper():
            ws.cell(row=r, column=intervention_col).value = "INGOs"

    REGULAR_INTERVENTIONS = {
        "TELECOMMUNICATIONS",
        "HEALTH",
        "WASH",
        "INGOS",
        "UN-OHCHR",
        "WFP",
        "LOGISTICS",   # IMPORTANT: exclude from prefix+convert logic
    }

    def _clean_str(v):
        return "" if v is None else str(v).strip()

    converted_desc_keys = set()  # <-- stable IDs for converted rows
    rows_data = []
    max_r = ws.max_row  # cache for speed

    for r in range(2, max_r + 1):
        iv_raw = _clean_str(ws.cell(row=r, column=intervention_col).value)
        iv_up = iv_raw.upper()

        is_converted = False

        # Apply your rule ONLY when intervention is not regular
        if iv_raw and (iv_up not in REGULAR_INTERVENTIONS):
            agency_cell = ws.cell(row=r, column=agency_col)
            agency_raw = _clean_str(agency_cell.value)

            prefix = f"{iv_raw} - "
            if not agency_raw.startswith(prefix):
                agency_cell.value = f"{iv_raw} - {agency_raw}" if agency_raw else f"{iv_raw} -"

            # Convert INTERVENTION to INGOs
            ws.cell(row=r, column=intervention_col).value = "INGOs"
            is_converted = True

        # After any conversion, read the final values used for sorting + tracking
        fuel_val = safe_float(ws.cell(row=r, column=unified_col).value)
        iv_after = _clean_str(ws.cell(row=r, column=intervention_col).value)
        desc_key = ws.cell(row=r, column=desc_col).value

        # Track converted rows by stable key (survives sorting/deletions)
        if is_converted:
            converted_desc_keys.add(desc_key)

        rows_data.append({
            "fuel": fuel_val,
            "intervention": iv_after,
            "is_converted": is_converted,
            "orig_index": r,
            "desc_key": desc_key,
            "row": snapshot_row(ws, r),
        })


    # ONE sort (tuple key) so behaviour is deterministic:
    # 1) intervention A–Z
    # 2) inside INGOs: real INGOs first, converted/prefixed last
    # 3) fuel DESC only for rows that are NOT converted prefixed INGOs
    # 4) stable fallback by orig_index
    def _sort_key(x):
        iv = (x["intervention"] or "").strip().lower()
        is_ingos = (iv == "ingos")

        # 0 = purple (real INGOs), 1 = white (converted/prefixed)
        converted_rank = 1 if (is_ingos and x["is_converted"]) else 0

        # Sort fuel DESC inside BOTH purple and white groups (independently)
        fuel_rank = -x["fuel"] if is_ingos else -x["fuel"]

        return (iv, converted_rank, fuel_rank, x["orig_index"])

    rows_data.sort(key=_sort_key)

    # Restore rows and rebuild "converted rows" positions AFTER sorting
    write_row = 2
    for obj in rows_data:
        restore_row(ws, write_row, obj["row"])
        write_row += 1


    if progress:
        progress.progress(80)

    # Remove duplicates by Description Sum
    if status:
        status.info("Removing duplicates…")

    seen = set()
    dup_rows = []
    for r in range(2, ws.max_row + 1):
        val = ws.cell(row=r, column=desc_col).value
        key = "" if val is None else str(val).strip()
        if key in seen:
            dup_rows.append(r)
        else:
            seen.add(key)
    for rr in sorted(dup_rows, reverse=True):
        ws.delete_rows(rr)

    if progress:
        progress.progress(86)

    # Total Sum Per Category in G
    if status:
        status.info("Building Total Sum Per Category…")

    total_cat_col = 7  # G
    if ws.max_column >= total_cat_col:
        ws.insert_cols(total_cat_col)

    ref_total_header = ws.cell(row=header_row, column=unified_col)  # F header
    total_header = ws.cell(row=header_row, column=total_cat_col)    # G header
    total_header.value = "Total Sum Per Category"
    total_header.font = copy(ref_total_header.font)
    total_header.fill = copy(ref_total_header.fill)
    total_header.border = copy(ref_total_header.border)
    total_header.alignment = copy(ref_total_header.alignment)
    total_header.number_format = ref_total_header.number_format
    total_header.protection = copy(ref_total_header.protection)

    ref_letter = get_column_letter(unified_col)
    new_letter = get_column_letter(total_cat_col)
    if ref_letter in ws.column_dimensions and ws.column_dimensions[ref_letter].width is not None:
        ws.column_dimensions[new_letter].width = ws.column_dimensions[ref_letter].width

    cat_totals = {}
    for r in range(2, ws.max_row + 1):
        cat = ws.cell(row=r, column=1).value
        cat_key = "" if cat is None else str(cat).strip()
        fuel_val = safe_float(ws.cell(row=r, column=unified_col).value)
        cat_totals[cat_key] = cat_totals.get(cat_key, 0.0) + fuel_val

    for r in range(2, ws.max_row + 1):
        cat = ws.cell(row=r, column=1).value
        cat_key = "" if cat is None else str(cat).strip()

        target = ws.cell(row=r, column=total_cat_col)
        style_src = ws.cell(row=r, column=unified_col)

        target.value = cat_totals.get(cat_key, 0.0)
        target.font = copy(style_src.font)
        target.fill = copy(style_src.fill)
        target.border = copy(style_src.border)
        target.alignment = copy(style_src.alignment)
        target.number_format = style_src.number_format
        target.protection = copy(style_src.protection)

    if progress:
        progress.progress(92)

    # Merge Total Sum Per Category (G) by INTERVENTION
    if status:
        status.info('Merging "Total Sum Per Category" by INTERVENTION…')

    INTERVENTION_COL = intervention_col
    TOTAL_CAT_COL = total_cat_col

    def _norm_intervention(v):
        return "" if v is None else str(v).strip().upper()

    start = 2
    while start <= ws.max_row:
        key = _norm_intervention(ws.cell(row=start, column=INTERVENTION_COL).value)

        end = start
        while end + 1 <= ws.max_row and _norm_intervention(ws.cell(row=end + 1, column=INTERVENTION_COL).value) == key:
            end += 1

        if end > start:
            top_cell = ws.cell(row=start, column=TOTAL_CAT_COL)
            top_val = top_cell.value
            top_style = {
                "font": copy(top_cell.font),
                "fill": copy(top_cell.fill),
                "border": copy(top_cell.border),
                "alignment": copy(top_cell.alignment),
                "number_format": top_cell.number_format,
                "protection": copy(top_cell.protection),
            }

            ws.merge_cells(
                start_row=start,
                start_column=TOTAL_CAT_COL,
                end_row=end,
                end_column=TOTAL_CAT_COL,
            )

            merged_top = ws.cell(row=start, column=TOTAL_CAT_COL)
            merged_top.value = top_val
            merged_top.font = copy(top_style["font"])
            merged_top.fill = copy(top_style["fill"])
            merged_top.border = copy(top_style["border"])
            merged_top.alignment = copy(top_style["alignment"])
            merged_top.number_format = top_style["number_format"]
            merged_top.protection = copy(top_style["protection"])

        start = end + 1

    # OPTIONAL: rename LOGISTICS -> UN Agencies
    for r in range(2, ws.max_row + 1):
        v = ws.cell(row=r, column=intervention_col).value
        if v is None:
            continue
        if str(v).strip().upper() == "LOGISTICS":
            ws.cell(row=r, column=intervention_col).value = "UN Agencies"

    # Color A–G based on INTERVENTION
    if status:
        status.info("Applying colours…")

    def _is_converted_row(r: int) -> bool:
        iv = ws.cell(row=r, column=intervention_col).value
        if ("" if iv is None else str(iv).strip().upper()) != "INGOS":
            return False
        return ws.cell(row=r, column=desc_col).value in converted_desc_keys

    fills = {
        "TELECOMMUNICATIONS": rgb_fill(213, 243, 251),
        "HEALTH": rgb_fill(0, 176, 80),
        "WASH": rgb_fill(250, 178, 138),
        "INGOs": rgb_fill(190, 158, 242),
        "WFP": rgb_fill(44, 195, 236),
        "UN_AGENCIES": rgb_fill(0, 176, 240),
    }

    COLOR_MIN_COL = 1
    COLOR_MAX_COL = 7

    for r in range(2, ws.max_row + 1):
        intervention_val = ws.cell(row=r, column=intervention_col).value
        intervention_text = "" if intervention_val is None else str(intervention_val).strip()
        intervention_up = intervention_text.upper()

        row_fill = None

        if intervention_up == "WFP":
            row_fill = fills["WFP"]
        elif intervention_up == "UN AGENCIES":
            row_fill = fills["UN_AGENCIES"]
        elif intervention_up == "TELECOMMUNICATIONS":
            row_fill = fills["TELECOMMUNICATIONS"]
        elif intervention_up == "HEALTH":
            row_fill = fills["HEALTH"]
        elif intervention_up == "WASH":
            row_fill = fills["WASH"]
        elif intervention_up == "INGOS":
            # Converted/prefixed rows must remain white; real INGOs must be purple
            if _is_converted_row(r):
                row_fill = None
            else:
                row_fill = fills["INGOs"]


        # CRITICAL: never write None into .fill
        if row_fill is None:
            continue

        for c in range(COLOR_MIN_COL, COLOR_MAX_COL + 1):
            ws.cell(row=r, column=c).fill = row_fill


    # Summary sheet (Sector Summary) + pie chart
    totals_by_intervention = {}
    for r in range(2, ws.max_row + 1):
        iv = ws.cell(row=r, column=intervention_col).value
        iv_key = "" if iv is None else str(iv).strip()
        if iv_key == "":
            continue
        totals_by_intervention[iv_key.upper()] = totals_by_intervention.get(iv_key.upper(), 0.0) + safe_float(
            ws.cell(row=r, column=unified_col).value
        )

    sector_rows = [
        ("TELECOMMUNICATIONS", "תקשורת"),
        ("HEALTH", "בריאות"),
        ("WASH", "סניטציה"),
        ("INGOS", "ארגונים לא ממשלתיים"),
        ("UN AGENCIES", 'סוכנויות או"ם'),
        ("WFP", "WFP"),
    ]

    SUMMARY_SHEET_NAME = "Sector Summary"
    if SUMMARY_SHEET_NAME in wb.sheetnames:
        wb.remove(wb[SUMMARY_SHEET_NAME])
    ws_sum = wb.create_sheet(SUMMARY_SHEET_NAME)

    ws_sum["A1"].value = "סקטור"
    ws_sum["B1"].value = "כמות דלק (ליטר)"

    header_fill = PatternFill("solid", fgColor="000000")
    header_font = Font(color="FFFFFF", bold=True)
    for cell_ref in ("A1", "B1", "C1"):
        cell = ws_sum[cell_ref]
        cell.fill = header_fill
        cell.font = header_font

    hdr_src = ws.cell(row=1, column=1)
    for addr in ("A1", "B1"):
        c = ws_sum[addr]
        c.font = copy(hdr_src.font)
        c.fill = copy(hdr_src.fill)
        c.border = copy(hdr_src.border)
        c.alignment = copy(hdr_src.alignment)
        c.number_format = hdr_src.number_format
        c.protection = copy(hdr_src.protection)

    ws_sum.column_dimensions["A"].width = 28
    ws_sum.column_dimensions["B"].width = 18

    num_src = ws.cell(row=2, column=unified_col)
    litres_number_format = num_src.number_format

    summary_fills = {
        "TELECOMMUNICATIONS": fills["TELECOMMUNICATIONS"],
        "HEALTH": fills["HEALTH"],
        "WASH": fills["WASH"],
        "INGOS": fills["INGOs"],
        "UN AGENCIES": fills.get("UN_AGENCIES", rgb_fill(0, 176, 240)),
        "WFP": fills["WFP"],
    }

    style_row_by_intervention = {}
    for r in range(2, ws.max_row + 1):
        iv = ws.cell(row=r, column=intervention_col).value
        k = "" if iv is None else str(iv).strip().upper()
        if k and k not in style_row_by_intervention:
            style_row_by_intervention[k] = r

    row_i = 2
    for key_en, label_he in sector_rows:
        key_u = key_en.strip().upper()

        src_r = style_row_by_intervention.get(key_u)

        a_cell = ws_sum.cell(row=row_i, column=1)
        b_cell = ws_sum.cell(row=row_i, column=2)

        a_cell.value = label_he
        b_cell.value = totals_by_intervention.get(key_u, 0.0)

        if src_r is not None:
            src_a = ws.cell(row=src_r, column=intervention_col)
            src_b = ws.cell(row=src_r, column=unified_col)
            copy_cell_style(src_a, a_cell)
            copy_cell_style(src_b, b_cell)
        else:
            copy_cell_style(hdr_src, a_cell)
            copy_cell_style(hdr_src, b_cell)
            b_cell.number_format = litres_number_format

        fill_obj = summary_fills.get(key_u)
        if fill_obj:
            a_cell.fill = fill_obj
            b_cell.fill = fill_obj

        row_i += 1

    FIRST_ROW = 2
    LAST_ROW = FIRST_ROW + len(sector_rows) - 1

    ws_sum["C1"].value = "אחוז"
    copy_cell_style(ws_sum["A1"], ws_sum["C1"])
    ws_sum.column_dimensions["C"].width = 12

    for r in range(FIRST_ROW, LAST_ROW + 1):
        ws_sum.cell(row=r, column=3).value = f"=B{r}/SUM($B${FIRST_ROW}:$B${LAST_ROW})"
        copy_cell_style(ws_sum.cell(row=r, column=2), ws_sum.cell(row=r, column=3))
        ws_sum.cell(row=r, column=3).number_format = "0.0%"

    pie = PieChart()
    data = Reference(ws_sum, min_col=2, min_row=1, max_row=LAST_ROW)
    cats = Reference(ws_sum, min_col=1, min_row=FIRST_ROW, max_row=LAST_ROW)
    pie.add_data(data, titles_from_data=True)
    pie.set_categories(cats)

    pie.title = ""
    pie.dataLabels = DataLabelList()
    pie.dataLabels.showCatName = True
    pie.dataLabels.showPercent = True
    pie.dataLabels.showVal = False
    pie.dataLabels.showSerName = False
    pie.dataLabels.showLeaderLines = True
    pie.dataLabels.separator = "\n"
    pie.legend = None

    slice_colors = [
        "D5F3FB",
        "00B050",
        "FAB28A",
        "BE9EF2",
        "00B0F0",
        "2CC3EC",
    ]

    ser = pie.series[0]
    ser.dPt = []
    for i, hx in enumerate(slice_colors[: len(sector_rows)]):
        dp = DataPoint(idx=i)
        dp.graphicalProperties.solidFill = hx
        ser.dPt.append(dp)

    ws_sum.add_chart(pie, "E2")

    if progress:
        progress.progress(100)

    # Remove ALL bold on Distribution Summary except header row
    remove_bold_except_header(ws, header_row=1)

    # Force header style on both sheets
    style_headers_black_only_with_text(ws, header_row=1)      # Distribution Summary
    style_headers_black_only_with_text(ws_sum, header_row=1)  # Sector Summary

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out

# ============================================================
# UI: Upload multiple files → run combine → run calculations → download
# ============================================================
left, right = st.columns([2, 1])

with left:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    uploads = st.file_uploader("Upload Total Distribution .xlsx files", type=["xlsx"], accept_multiple_files=True)
    st.markdown(
        '<div class="small">Combines UNOPS + WFP Total Distribution sheets first (values-only), then runs the Fuels Cleaner calculations on the combined result.</div>',
        unsafe_allow_html=True,
    )
    st.markdown("</div>", unsafe_allow_html=True)

with right:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown("**Pipeline**")
    st.markdown(
        """
1) Combine files (values-only)  
2) Trim last text row per file  
3) Rebuild merges A/B  
4) Optional: delete F–I  
5) Run Fuels Cleaner calculations  
6) Download final workbook
"""
    )
    st.markdown("</div>", unsafe_allow_html=True)


def _reset_cached_output():
    st.session_state.pop("final_bytes", None)
    st.session_state.pop("final_name", None)
    st.session_state.pop("last_upload_key", None)


if "last_upload_key" not in st.session_state:
    st.session_state.last_upload_key = None


run_btn = st.button("Run combine + calculations", type="primary", disabled=not uploads)

if uploads:
    upload_key = "|".join([f"{u.name}:{u.size}" for u in uploads])
    if st.session_state.last_upload_key != upload_key:
        st.session_state.last_upload_key = upload_key
        _reset_cached_output()

if "final_bytes" in st.session_state and st.session_state.get("final_bytes"):
    st.success("Done! Download your final file below.")
    st.download_button(
        "Download final Excel",
        data=st.session_state.final_bytes,
        file_name=st.session_state.get("final_name", "Fuels summary.xlsx"),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    st.stop()

if run_btn:
    progress = st.progress(0)
    status = st.empty()

    loader = st.empty()
    show_small_loader_video(loader, "dog_running.mp4", width_px=220)

    try:
        status.info("Step 1/2: Combining files…")
        progress.progress(5)

        combined_bytes = build_combined_workbook_bytes(uploads, status=status)

        progress.progress(25)
        status.info("Step 2/2: Running calculations on combined workbook…")

        final_bytes = run_calculations_on_combined_bytes(combined_bytes, progress=progress, status=status)

        # Cache final output (prevents rerun on download click)
        st.session_state.final_bytes = final_bytes.getvalue()
        st.session_state.final_name = "Fuels summary.xlsx"

        status.success("Done! Download your final file below.")
        st.download_button(
            "Download final Excel",
            data=st.session_state.final_bytes,
            file_name=st.session_state.final_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    except Exception as e:
        status.empty()
        st.exception(e)

    finally:
        loader.empty()
