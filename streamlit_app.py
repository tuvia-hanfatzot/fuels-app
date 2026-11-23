"""Combine fuel distribution workbooks and build the weekly Excel report.

Run with: streamlit run fuel_report.py
Keep the approval workbook beside this file. generate_report() also accepts
file-like uploads directly and can be used without the Streamlit interface.
"""

import base64
import hashlib
import os
import re
from contextlib import closing, contextmanager
from copy import copy
from datetime import date, datetime, timedelta
from difflib import SequenceMatcher, get_close_matches
from functools import lru_cache
from io import BytesIO

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import Cell, MergedCell
from openpyxl.compat.strings import safe_string
from openpyxl.chart import PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.styles.cell_style import StyleArray
from openpyxl.styles.numbers import is_date_format, is_timedelta_format
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import range_boundaries
from openpyxl.utils.datetime import from_excel


TARGET_SHEETS = {
    "UNOPS": "UNOPS Total Distribution",
    "WFP": "WFP Total Distribution",
}

HEADER_ROWS = 2
DATA_START_ROW = 3


FREEZE_COLS = (3, 4)
DELETE_COMBINED_COLS_F_TO_I_BEFORE_CALC = True


STATUS_COL_NAME = "סטטוס"
UNKNOWN_ENTITIES_STATUS = "ארגונים/גופים לא מוכרים (INGOs/חברות מקומיות)"
DISTRIBUTION_SHEET_NAME = "Distribution Summary"
FUZZY_THRESHOLD = 0.78

APPROVAL_SHEETS = {
    "UNOPS": ["UNOPS Total Distribution", "UNOPS Total Distribution "],
    "WFP": ["WFP Total Distribution"],
}


BASE_GENERATED_PREFIXES = {
    "NNGOS",
    "NNGO",
    "FINANCIAL INSTITUTIONS",
    "FINANCIAL INSTITUTION",
    "UN-OHCHR",
    "UN OHCHR",
    "UN - SISTERS LOGISTICS",
    "UN SISTERS LOGISTICS",
    "UN - SISTERS AGENCIES",
    "UN SISTERS AGENCIES",
    "PROTECTION",
    "SHELTER",
    "EDUCATION",
    "NUTRITION",
    "LOGISTICS",
    "FOOD SECURITY",
    "LOCAL AUTHORITIES",
    "CCCM",
}

STATUS_FILL_RULES = [
    ("מסורב", "F4B084"),
    ("מוקפא", "B4C6E7"),
    ("מאפיות", "FFE699"),
    ("מאושר", "C6E0B4"),
    ("שותפים", "E2F0D9"),
    ("בריאות", "C6E0B4"),
    ("תקשורת", "DDEBF7"),
    ("סניטציה", "FCE4D6"),
    ("סוכנויות", "D9EAD3"),
    ("ארגונים/גופים לא מוכרים", "C044E8"),
    ("חברות מקומיות", "C044E8"),
    ("לא מוכרים", "C044E8"),
    ("NOT FOUND", "D9D9D9"),
    ("UNKNOWN", "D9D9D9"),
]


APPROVAL_FILE_NAME = "חלוקת דלקים לארגונים מאושרים, מסורבים ומוקפאים.xlsx"
APPROVAL_STORE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), APPROVAL_FILE_NAME) if "__file__" in globals() else APPROVAL_FILE_NAME


def workbook_bytes(workbook):
    """Serialise a completed workbook and rewind it for the next reader."""
    output = BytesIO()
    workbook.save(output)
    output.seek(0)
    return output


def _prepare_excel_values(ws):
    """Keep the numeric precision and cell types of the former Excel save/reload steps."""
    date_formats = {}
    for cell in ws._cells.values():
        value = cell.value
        if cell.data_type != "n" or type(value) not in (float, int):
            continue
        if isinstance(value, float) or abs(value) >= 10**15:
            text = safe_string(value)
            if not text:
                cell.value = None
                continue
            value = float(text) if "." in text or "e" in text.lower() else int(text)
            cell.value = value
        fmt = cell.number_format
        if fmt not in date_formats:
            date_formats[fmt] = (is_date_format(fmt), is_timedelta_format(fmt))
        is_date, is_duration = date_formats[fmt]
        if is_date:
            try:
                cell.value = from_excel(value, ws.parent.epoch, timedelta=is_duration)
            except (OverflowError, ValueError):
                cell.value = "#VALUE!"


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


# Combining source worksheets


def _norm(s: str) -> str:
    return " ".join((s or "").lower().split())


def _find_target_sheet(wb):
    names = {name: _norm(name) for name in wb.sheetnames}
    for org, target in TARGET_SHEETS.items():
        for name, normalised in names.items():
            if normalised == _norm(target):
                return org, name
    for name, normalised in names.items():
        if "total" in normalised and "distribution" in normalised:
            for org in TARGET_SHEETS:
                if org.lower() in normalised:
                    return org, name
    return None, None


def _row_has_data_row(ws, r, max_col):
    return any(ws.cell(r, c).value not in (None, "") for c in range(1, max_col + 1))


def _first_data_row(ws, start=DATA_START_ROW, max_scan=200):
    # Inspect existing cells only. A format applied far down a sheet can inflate its dimensions.
    stop = start + max_scan
    return min((cell.row for cell in ws._cells.values()
                if start <= cell.row <= stop and cell.value not in (None, "")), default=start)


def _last_data_row(ws, start_row, max_col):
    return max((cell.row for cell in ws._cells.values()
                if cell.row >= start_row and cell.column <= max_col and cell.value not in (None, "")),
               default=start_row - 1)


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


def _copy_style_between_workbooks(src, dst):
    # Style IDs are local to a workbook. Register each component in the destination.
    dst._style = copy(src._style)
    dst.font = copy(src.font)
    dst.fill = copy(src.fill)
    dst.border = copy(src.border)
    dst.alignment = copy(src.alignment)
    dst.number_format = src.number_format
    dst.protection = copy(src.protection)


def _copy_block_values_only(src_ws, src_ws_vals, dst_ws, src_row_start, src_row_end,
                            dst_row_start, max_col, style_map=None, formula_values=None):
    """Copy values, styles and fully contained merges. Share the mapping across blocks."""
    if style_map is None:
        style_map = {}
    offset = dst_row_start - src_row_start
    for r in range(src_row_start, src_row_end + 1):
        height = src_ws.row_dimensions.get(r)
        if height is not None and height.height is not None:
            dst_ws.row_dimensions[r + offset].height = height.height

        for c in range(1, max_col + 1):
            src = src_ws.cell(r, c)
            dst = dst_ws.cell(r + offset, c)
            style_id = src.style_id
            if style_id not in style_map:
                _copy_style_between_workbooks(src, dst)
                style_map[style_id] = copy(dst._style)
            dst._style = copy(style_map[style_id])
            value = src.value
            if src_ws_vals is not None or src.data_type == "f":
                cached = (_safe_value(src_ws_vals, r, c) if src_ws_vals is not None
                          else (formula_values or {}).get((r, c)))
                if cached is not None or (isinstance(value, str) and value.startswith("=")):
                    value = cached
            dst.value = value

    for merged in src_ws.merged_cells.ranges:
        if src_row_start <= merged.min_row and merged.max_row <= src_row_end and merged.max_col <= max_col:
            dst_ws.merge_cells(start_row=merged.min_row + offset, end_row=merged.max_row + offset,
                               start_column=merged.min_col, end_column=merged.max_col)


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
    rows_with_data = {cell.row for cell in ws._cells.values()
                      if data_start <= cell.row <= data_end and cell.value not in (None, "")}
    # A formatted cell far below the data should not turn a sparse sheet into
    # a full rectangle. Move existing cells right, from the last column back.
    for sheet in (ws, ws_vals):
        if sheet is not None:
            for row, col in sorted(sheet._cells, key=lambda pos: pos[1], reverse=True):
                sheet._move_cell(row, col, row_offset=0, col_offset=1)
    style = ws.cell(data_start, 2)._style
    for r in rows_with_data:
        cell = ws.cell(r, 1)
        cell.value = "WFP"
        cell._style = copy(style)


def _row_has_data_for_merge(ws, r, max_col):
    return any(ws.cell(r, c).value not in (None, "") for c in range(2, max_col + 1))


def _merge_column_group(ws, col, first, last):
    if first < last:
        anchor = ws.cell(first, col)
        style = copy(anchor._style)
        ws.merge_cells(start_row=first, end_row=last, start_column=col, end_column=col)
        anchor._style = style


def _merge_wfp_blocks_in_master(master_ws, max_col):
    _clear_merges_on_cols(master_ws, cols=(1,), from_row=DATA_START_ROW)
    last = master_ws.max_row
    first = None
    for r in range(DATA_START_ROW, last + 2):
        is_wfp = (r <= last and master_ws.cell(r, 1).value == "WFP"
                  and _row_has_data_for_merge(master_ws, r, max_col))
        if is_wfp and first is None:
            first = r
        elif not is_wfp and first is not None:
            _merge_column_group(master_ws, 1, first, r - 1)
            first = None


def _clear_merges_on_cols(ws, cols=(1, 2), from_row=DATA_START_ROW):
    for merged in list(ws.merged_cells.ranges):
        if merged.max_row >= from_row and any(merged.min_col <= c <= merged.max_col for c in cols):
            ws.unmerge_cells(str(merged))


def _merge_down_by_blanks(ws, col, max_col, from_row=DATA_START_ROW):
    """A value starts a group. Empty data rows and the next value end it."""
    first = None
    last = ws.max_row
    for r in range(from_row, last + 2):
        has_data = r <= last and _row_has_data_row(ws, r, max_col)
        has_label = has_data and ws.cell(r, col).value not in (None, "")
        if first is not None and (has_label or not has_data):
            _merge_column_group(ws, col, first, r - 1)
            first = None
        if has_label:
            first = r


def _clear_merges_intersecting_cols(ws, col_start, col_end):
    for merged in list(ws.merged_cells.ranges):
        if merged.min_col <= col_end and merged.max_col >= col_start:
            ws.unmerge_cells(str(merged))


def _delete_cols_safe(ws, col_start, col_end):
    _clear_merges_intersecting_cols(ws, col_start, col_end)
    count = min(ws.max_column, col_end) - col_start + 1
    if count > 0:
        ws.delete_cols(col_start, count)


def _ensure_real_cell(ws, row, col):
    """
    Converts a stale MergedCell into a normal writable Cell.
    This is only needed because old code directly edited ws.merged_cells.ranges.
    """
    cell = ws.cell(row=row, column=col)

    if isinstance(cell, MergedCell):
        ws._cells.pop((row, col), None)
        cell = ws.cell(row=row, column=col)

    return cell


def _read_distribution_tasks(uploads, status):
    tasks = []
    for upload in uploads:
        try:
            with closing(load_workbook(BytesIO(upload.getvalue()), read_only=True, keep_links=False)) as metadata:
                kind, sheet = _find_target_sheet(metadata)
                if kind:
                    tasks.append((kind, sheet, upload))
                elif status is not None:
                    status.warning(f"Skipped '{upload.name}': no matching Total Distribution sheet found. "
                                   f"Sheets: {metadata.sheetnames}")
        except Exception as exc:
            if status is not None:
                status.warning(f"Skipped '{upload.name}': could not read as .xlsx ({exc})")
    if not tasks:
        raise RuntimeError("No valid files found.")
    # Sorting is stable, so uploads from the same organisation keep their order.
    return sorted(tasks, key=lambda item: item[0] != "UNOPS")


def _cached_formula_values(data, ws, last_row, column_offset=0):
    """Stream just the selected sheet's cached formula results when it contains formulas."""
    columns_by_row = {}
    for cell in ws._cells.values():
        if cell.data_type == "f" and cell.row <= last_row:
            columns_by_row.setdefault(cell.row, []).append(cell.column)
    if not columns_by_row:
        return {}

    first = min(columns_by_row)
    width = max(max(cols) for cols in columns_by_row.values())
    result = {}
    with closing(load_workbook(BytesIO(data), read_only=True, data_only=True, keep_links=False)) as values_book:
        rows = values_book[ws.title].iter_rows(min_row=first, max_row=max(columns_by_row),
                                               max_col=width, values_only=True)
        for r, values in enumerate(rows, first):
            for col in columns_by_row.get(r, ()):
                result[r, col + column_offset] = values[col - 1]
    return result


def build_combined_workbook(uploads, status=None):
    """Build the master in memory, processing one distribution workbook at a time."""
    tasks = _read_distribution_tasks(uploads, status)
    master = Workbook()
    sheet = master.active
    sheet.title = DISTRIBUTION_SHEET_NAME
    next_row = DATA_START_ROW
    widest = 0
    padding = []
    header_widths = {}

    readable_source = False
    for kind, sheet_name, upload in tasks:
        data = upload.getvalue()
        try:
            source_book = load_workbook(BytesIO(data), keep_links=False)
        except Exception as exc:
            if status is not None:
                status.warning(f"Skipped '{upload.name}': could not read as .xlsx ({exc})")
            continue
        with closing(source_book) as source:
            readable_source = True
            ws = source[sheet_name]
            offset = int(kind == "WFP")
            width = ws.max_column + offset
            widest = max(widest, width)
            first = _first_data_row(ws)
            footer = _last_data_row(ws, first, width)
            if footer <= first:
                reason = "has no data rows" if footer < first else "became empty after removing last text row"
                if status is not None:
                    status.warning(f"⚠️ {kind} | {upload.name} {reason}.")
                continue

            # The last populated row is the source footer. Exclude it from the copy.
            last = footer - 1
            cached = _cached_formula_values(data, ws, last, offset)
            if kind == "WFP":
                _insert_wfp_column_a(ws, None, first, footer)

            style_map = {}
            block_start = next_row
            if next_row == DATA_START_ROW:
                header_widths = {key: dim.width for key, dim in ws.column_dimensions.items()}
                _copy_dimensions(ws, sheet, width, HEADER_ROWS)
                _copy_block_values_only(ws, None, sheet, 1, HEADER_ROWS, 1, width, style_map, cached)
                block_start = 1
            _copy_block_values_only(ws, None, sheet, first, last, next_row, width, style_map, cached)
            next_row += last - first + 1

            # A narrower input may have a custom default style even in its empty columns.
            if 0 not in style_map:
                blank = Cell(ws)
                template = Cell(sheet)
                _copy_style_between_workbooks(blank, template)
                style_map[0] = template._style
            padding.append((block_start, next_row - 1, width, style_map[0]))

    if not readable_source:
        raise RuntimeError("No valid files found.")
    if next_row > DATA_START_ROW:
        sheet.cell(1, widest)
        for start, end, width, style in padding:
            if width < widest and style and any(style):
                for row in sheet.iter_rows(min_row=start, max_row=end, min_col=width + 1, max_col=widest):
                    for cell in row:
                        cell._style = copy(style)
        for c in range(1, widest + 1):
            letter = get_column_letter(c)
            if header_widths.get(letter) is not None:
                sheet.column_dimensions[letter].width = header_widths[letter]

    width = sheet.max_column
    _merge_wfp_blocks_in_master(sheet, width)
    _clear_merges_on_cols(sheet)
    for col in (1, 2):
        _merge_down_by_blanks(sheet, col, width)
    for c in range(1, width + 1):
        sheet.column_dimensions[get_column_letter(c)].hidden = False
    for r in range(1, sheet.max_row + 1):
        sheet.row_dimensions[r].hidden = False
    if DELETE_COMBINED_COLS_F_TO_I_BEFORE_CALC:
        _delete_cols_safe(sheet, 6, 9)
    return master


def build_combined_workbook_bytes(uploads, status=None):
    with closing(build_combined_workbook(uploads, status)) as workbook:
        return workbook_bytes(workbook)


# Distribution calculations


def _copy_formatting(source, existing):
    # Named-style, quote-prefix and pivot flags belong to the destination cell.
    style = copy(source) if source is not None else StyleArray()
    style.xfId = existing.xfId if existing else 0
    style.quotePrefix = existing.quotePrefix if existing else 0
    style.pivotButton = existing.pivotButton if existing else 0
    return style


def copy_cell_style(src, dst):
    """Copy the six formatting components without sharing a mutable style array."""
    if src is None or dst is None:
        return
    if src.parent.parent is not dst.parent.parent:
        # The public helper also accepts cells from two independent workbooks.
        for attr in ("font", "fill", "border", "alignment", "number_format", "protection"):
            setattr(dst, attr, copy(getattr(src, attr)))
        return
    dst._style = _copy_formatting(src._style, dst._style)


def unmerge_and_fill(ws, col_min: int, col_max: int):
    for merged in list(ws.merged_cells.ranges):
        if merged.max_col < col_min or merged.min_col > col_max:
            continue
        anchor = ws.cell(merged.min_row, merged.min_col)
        value = anchor.value
        ws.unmerge_cells(str(merged))
        for row in ws.iter_rows(min_row=merged.min_row, max_row=merged.max_row,
                                min_col=merged.min_col, max_col=merged.max_col):
            for cell in row:
                cell.value = value
                copy_cell_style(anchor, cell)


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


def safe_float(value):
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        s = value.strip().replace(",", "")
        if not s:
            return 0.0
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
    plain_fonts = {}
    for row in ws.iter_rows():
        for cell in row:
            if cell.row == header_row or not cell.font.bold:
                continue
            font_id = cell._style.fontId
            if font_id not in plain_fonts:
                plain_fonts[font_id] = _font_without_bold(cell.font)
            cell.font = plain_fonts[font_id]


def _find_header_col_by_names(ws, header_row, header_names):
    """Find a column by one of several header names using the existing normaliser."""
    wanted = {norm_header(name) for name in header_names}
    for c in range(1, ws.max_column + 1):
        if norm_header(ws.cell(row=header_row, column=c).value) in wanted:
            return c
    return None


def _strip_leading_usage_marker(value):
    """
    New UNOPS format can prefix Description values with:
      - Internal use -
      - external use
    Return (clean_value, usage_type). usage_type is 'internal', 'external', or None.
    """
    if value is None:
        return value, None

    text_value = str(value)
    pattern = re.compile(r"^\s*(internal\s+use|external\s+use)\b\s*[-–—:]?\s*", re.IGNORECASE)
    match = pattern.match(text_value)
    if not match:
        return value, None

    usage_type = "internal" if match.group(1).strip().lower().startswith("internal") else "external"
    clean_value = pattern.sub("", text_value, count=1).strip()
    return clean_value, usage_type


def _compact_alpha_num(value) -> str:
    if value is None:
        return ""
    return re.sub(r"[^A-Z0-9]", "", str(value).upper())


def _contains_ohchr(value) -> bool:
    return "OHCHR" in _compact_alpha_num(value)


def _is_wfp_un_agency(agency) -> bool:
    """Recognise WFP itself and its Food-sector label, excluding partner agencies."""
    return _compact_alpha_num(agency) in {"WFP", "FOODWFP"}


def _is_regular_cluster_for_internal_use(value) -> bool:
    """
    Internal-use rows should become LOGISTICS only when the original cluster is a
    regular cluster. If the cluster is irregular (for example, UN - Sisters Agencies),
    keep it untouched so the existing irregular-cluster logic converts it to INGOs
    with the generated prefix: <cluster> - <agency>.
    """
    return norm_header(value) in {
        "ETC",
        "HEALTH",
        "WASH",
        "INGOS",
        "UN-OHCHR",
        "UN OHCHR",
        "UN AGENCIES",
        "LOGISTICS",
    }


def normalise_distribution_description_usage_markers(ws, header_row=1):
    """
    Normalise the new UNOPS Description format back to the old working format.

    - Removes leading 'Internal use -' / 'external use' from Description.
    - Internal-use rows become LOGISTICS only when the original cluster is regular.
      Irregular clusters stay untouched so the existing logic later converts them
      to INGOs and prefixes the Agency with the original cluster.
    - OHCHR rows are always treated as LOGISTICS so they later become UN Agencies,
      not a generated irregular entry such as 'Protection - OHCHR'.
    - Food / WFP (including the generated Food - WFP label) becomes LOGISTICS,
      then UN Agencies. Rows from the dedicated WFP source remain WFP.
    """
    description_col = _find_header_col_by_names(ws, header_row, ["DESCRIPTION", "Description"])
    cluster_col = _find_header_col_by_names(ws, header_row, ["CLUSTER", "Cluster", "INTERVENTION", "Intervention"])
    agency_col = _find_header_col_by_names(ws, header_row, ["AGENCY", "Agency"])

    if description_col is None:
        return

    for r in range(header_row + 1, ws.max_row + 1):
        cluster_cell = ws.cell(row=r, column=cluster_col) if cluster_col is not None else None
        agency_value = ws.cell(row=r, column=agency_col).value if agency_col is not None else None
        cluster_value = cluster_cell.value if cluster_cell is not None else None

        if cluster_cell is not None and (_contains_ohchr(cluster_value) or _contains_ohchr(agency_value)):
            cluster_cell.value = "LOGISTICS"
            cluster_value = cluster_cell.value

        # Food - WFP belongs to UN Agencies. Normalise before aggregation,
        # sorting, category totals, colours, and Sector Summary are generated.
        if cluster_cell is not None and norm_header(cluster_value) != "WFP" and (
            _compact_alpha_num(agency_value) == "FOODWFP"
            or (norm_header(cluster_value) == "FOOD" and _compact_alpha_num(agency_value) == "WFP")
        ):
            if norm_header(cluster_value) == "FOOD" and _compact_alpha_num(agency_value) == "WFP":
                ws.cell(row=r, column=agency_col).value = "Food - WFP"
            cluster_cell.value = "LOGISTICS"
            cluster_value = cluster_cell.value

        desc_cell = ws.cell(row=r, column=description_col)
        cleaned_desc, usage_type = _strip_leading_usage_marker(desc_cell.value)

        if usage_type is None:
            continue

        desc_cell.value = cleaned_desc

        if usage_type == "internal" and cluster_cell is not None:
            cluster_norm = norm_header(cluster_value)

            # Do not convert WFP rows; the new marker is intended for UNOPS Total Distribution.
            # Irregular clusters must keep the original cluster so the old irregular-cluster
            # conversion still produces: <cluster> - <agency>, then Cluster = INGOs.
            if cluster_norm != "WFP" and _is_regular_cluster_for_internal_use(cluster_value):
                cluster_cell.value = "LOGISTICS"


def _delete_rows_together(ws, removed):
    """Compact the surviving runs once, then delete the unused tail."""
    if not removed:
        return
    removed = sorted(set(removed))
    last = ws.max_row
    right = get_column_letter(ws.max_column)
    start = removed[0] + 1
    for count, stop in enumerate([*removed[1:], last + 1], 1):
        if start < stop:
            ws.move_range(f"A{start}:{right}{stop - 1}", rows=-count)
        start = stop + 1
    ws.delete_rows(last - len(removed) + 1, len(removed))


def _copy_width(ws, source, target):
    dim = ws.column_dimensions.get(source)
    if dim is not None and dim.width is not None:
        ws.column_dimensions[target].width = dim.width


def _add_fuel_columns(ws):
    """Combine diesel and petrol, then sum the surviving rows by description key."""
    unmerge_and_fill(ws, 4, 6)
    ws["F1"] = "Fuel sum"
    copy_cell_style(ws["D1"], ws["F1"])
    _copy_width(ws, "E" if "E" in ws.column_dimensions else "D", "F")

    empty_rows = []
    for diesel, petrol, fuel in ws.iter_rows(min_row=2, min_col=4, max_col=6):
        try:
            litres = (0.0 if diesel.value is None else float(diesel.value))
            litres += 0.0 if petrol.value is None else float(petrol.value)
        except (ValueError, TypeError, OverflowError):
            litres = 0.0
        if litres == 0:
            empty_rows.append(fuel.row)
        else:
            fuel.value = litres
            copy_cell_style(diesel if diesel.value is not None else petrol, fuel)
    _delete_rows_together(ws, empty_rows)

    ws.delete_cols(4, 2)
    ws.insert_cols(5, 2)
    ws["E1"] = "Description Sum"
    ws["F1"] = "Unified Fuel"
    copy_cell_style(ws["A1"], ws["E1"])
    copy_cell_style(ws["D1"], ws["F1"])
    _copy_width(ws, "C", "E")
    _copy_width(ws, "D", "F")

    totals = {}
    for cluster, agency, desc, fuel, key, unified in ws.iter_rows(min_row=2, max_col=6):
        key.value = ",".join("" if cell.value is None else str(cell.value) for cell in (cluster, agency, desc))
        copy_cell_style(cluster, key)
        copy_cell_style(fuel, unified)
        totals[key.value] = totals.get(key.value, 0.0) + safe_float(fuel.value)
    for key, unified in ws.iter_rows(min_row=2, min_col=5, max_col=6):
        unified.value = totals[key.value]


def _label(value):
    return "" if value is None else str(value).strip()


def _sort_distribution_rows(ws, intervention_col, agency_col):
    """Sort by cluster, original/converted INGOs and descending fuel, then keep the first key."""
    regular = {"ETC", "HEALTH", "WASH", "INGOS", "UN-OHCHR", "WFP", "UN AGENCIES", "LOGISTICS"}
    rows = []
    converted_keys = set()
    for row in ws.iter_rows(min_row=2):
        intervention = row[intervention_col - 1]
        agency = row[agency_col - 1]
        if _contains_ohchr(intervention.value) or _contains_ohchr(agency.value):
            intervention.value = "LOGISTICS"
        label = _label(intervention.value)
        converted = bool(label and label.upper() not in regular)
        key = row[4].value
        if converted:
            prefix = f"{label} - "
            if not _label(agency.value).startswith(prefix):
                agency.value = (prefix + _label(agency.value)).rstrip()
            intervention.value = "INGOs"
            converted_keys.add(key)

        order = (_label(intervention.value).lower(), converted, -safe_float(row[5].value))
        rows.append((order, _label(key), row))

    # The stable sort keeps the original order on ties. Snapshot only the winners
    # before overwriting cells that may also be the source of a later row.
    seen = set()
    winners = []
    ordered_rows = sorted(rows, key=lambda item: item[0])
    for position, (_, key, row) in enumerate(ordered_rows, 2):
        if key in seen:
            continue
        seen.add(key)
        # The previous sort left Excel's cell-specific flags at each sorted position.
        cells = [(cell.value, _copy_formatting(cell._style, ws.cell(position, cell.column)._style))
                 for cell in row]
        winners.append(cells)
    for r, values in enumerate(winners, 2):
        for c, (value, style) in enumerate(values, 1):
            cell = ws.cell(r, c)
            cell.value = value
            cell._style = style
    if len(winners) < len(rows):
        ws.delete_rows(2 + len(winners), len(rows) - len(winners))
    return converted_keys


SECTORS = (
    ("ETC", "תקשורת", "D5F3FB"),
    ("HEALTH", "בריאות", "00B050"),
    ("WASH", "סניטציה", "FAB28A"),
    ("INGOS", "ארגונים לא ממשלתיים", "BE9EF2"),
    ("UN AGENCIES", 'סוכנויות או"ם', "00B0F0"),
    ("WFP", "WFP", "2CC3EC"),
)


def _format_categories(ws, intervention_col, converted_keys, fills):
    if ws.max_column >= 7:
        ws.insert_cols(7)
    ws["G1"] = "Total Sum Per Category"
    copy_cell_style(ws["F1"], ws["G1"])
    _copy_width(ws, "F", "G")

    totals = {}
    last = ws.max_row
    for r in range(2, last + 1):
        category = _label(ws.cell(r, 1).value)
        totals[category] = totals.get(category, 0.0) + safe_float(ws.cell(r, 6).value)
    group_start = 2
    group_name = None
    for r in range(2, last + 1):
        total = ws.cell(r, 7)
        total.value = totals[_label(ws.cell(r, 1).value)]
        copy_cell_style(ws.cell(r, 6), total)
        category = _label(ws.cell(r, intervention_col).value).upper()
        if group_name is not None and category != group_name:
            _merge_column_group(ws, 7, group_start, r - 1)
            group_start = r
        group_name = category
    _merge_column_group(ws, 7, group_start, last)

    sector_totals = {}
    first_rows = {}
    for r in range(2, last + 1):
        category = _label(ws.cell(r, intervention_col).value).upper()
        if category == "LOGISTICS":
            ws.cell(r, intervention_col).value = "UN Agencies"
            category = "UN AGENCIES"
        if not category:
            continue
        first_rows.setdefault(category, r)
        sector_totals[category] = sector_totals.get(category, 0.0) + safe_float(ws.cell(r, 6).value)
        if category in fills and not (category == "INGOS" and ws.cell(r, 5).value in converted_keys):
            for col in range(1, 8):
                ws.cell(r, col).fill = fills[category]
    return sector_totals, first_rows


def _add_sector_summary(wb, ws, intervention_col, totals, first_rows, fills):
    if "Sector Summary" in wb.sheetnames:
        wb.remove(wb["Sector Summary"])
    summary = wb.create_sheet("Sector Summary")
    summary.append(["סקטור", "כמות דלק (ליטר)", "אחוז"])
    for header in summary[1]:
        copy_cell_style(ws["A1"], header)
    for col, width in (("A", 28), ("B", 18), ("C", 12)):
        summary.column_dimensions[col].width = width
    last = len(SECTORS) + 1
    for r, (key, label, _) in enumerate(SECTORS, 2):
        a, b, percent = (summary.cell(r, c) for c in (1, 2, 3))
        a.value, b.value = label, totals.get(key, 0.0)
        source_row = first_rows.get(key)
        if source_row is None:
            copy_cell_style(ws["A1"], a)
            copy_cell_style(ws["A1"], b)
            b.number_format = ws["F2"].number_format
        else:
            copy_cell_style(ws.cell(source_row, intervention_col), a)
            copy_cell_style(ws.cell(source_row, 6), b)
        a.fill = b.fill = fills[key]
        percent.value = f"=B{r}/SUM($B$2:$B${last})"
        copy_cell_style(b, percent)
        percent.number_format = "0.0%"

    pie = PieChart()
    pie.add_data(Reference(summary, min_col=2, min_row=1, max_row=last), titles_from_data=True)
    pie.set_categories(Reference(summary, min_col=1, min_row=2, max_row=last))
    pie.title = ""
    pie.dataLabels = DataLabelList()
    pie.dataLabels.showCatName = pie.dataLabels.showPercent = pie.dataLabels.showLeaderLines = True
    pie.dataLabels.showVal = pie.dataLabels.showSerName = False
    pie.dataLabels.separator = "\n"
    pie.legend = None
    for i, (_, _, colour) in enumerate(SECTORS):
        point = DataPoint(idx=i)
        point.graphicalProperties.solidFill = colour
        pie.series[0].dPt.append(point)
    summary.add_chart(pie, "E2")
    return summary


def calculate_distribution(wb, progress=None, status=None):
    """Clean the combined sheet, aggregate distribution rows and create Sector Summary."""
    def update(percent, message):
        if progress is not None:
            progress.progress(percent)
        if status is not None:
            status.info(message)

    names = {name.strip(): name for name in wb.sheetnames}
    if DISTRIBUTION_SHEET_NAME not in names:
        raise RuntimeError(f'Sheet "{DISTRIBUTION_SHEET_NAME}" not found. Sheets found: {list(names)}')
    ws = wb[names[DISTRIBUTION_SHEET_NAME]]
    for sheet in list(wb):
        if sheet is not ws:
            wb.remove(sheet)
    ws.title = DISTRIBUTION_SHEET_NAME

    update(10, "Unmerging A–C and removing TOTAL rows…")
    unmerge_and_fill(ws, 1, 3)
    ws.delete_rows(2)
    removed = [row[0].row for row in ws.iter_rows(max_col=3)
               if any("TOTAL" in str(cell.value).upper() for cell in row)]
    _delete_rows_together(ws, removed)
    normalise_distribution_description_usage_markers(ws)
    update(34, "Calculating fuel totals…")
    _add_fuel_columns(ws)

    headers = {norm_header(cell.value): cell.column for cell in ws[1] if cell.value is not None}
    intervention_col = headers.get("INTERVENTION", 1)
    if "AGENCY" not in headers:
        raise RuntimeError('Header "AGENCY" not found.')
    update(72, "Sorting and removing duplicate descriptions…")
    converted_keys = _sort_distribution_rows(ws, intervention_col, headers["AGENCY"])
    update(86, "Calculating and formatting sector totals…")
    fills = {key: PatternFill("solid", fgColor="FF" + colour) for key, _, colour in SECTORS}
    totals, first_rows = _format_categories(ws, intervention_col, converted_keys, fills)
    summary = _add_sector_summary(wb, ws, intervention_col, totals, first_rows, fills)
    remove_bold_except_header(ws)
    style_headers_black_only_with_text(ws)
    style_headers_black_only_with_text(summary)
    if progress is not None:
        progress.progress(100)
    return wb


def run_calculations_on_combined_bytes(combined_bytes: BytesIO, progress=None, status=None) -> BytesIO:
    combined_bytes.seek(0)
    with closing(load_workbook(combined_bytes)) as workbook:
        calculate_distribution(workbook, progress, status)
        return workbook_bytes(workbook)


# Latest-day workbook extraction and Fuel Dashboard


def _search_sheet(ws, text, exact):
    target = text.strip().upper() if exact else text.upper()
    for r, row in enumerate(ws.iter_rows(values_only=True), 1):
        for c, value in enumerate(row, 1):
            if value is None:
                continue
            label = str(value).upper()
            if (label.strip() == target if exact else target in label):
                return r, c
    return None, None


def find_cell_exact(ws, text: str):
    return _search_sheet(ws, text, exact=True)


def find_cell_containing(ws, text: str):
    return _search_sheet(ws, text, exact=False)


def normalise_date(v):
    if v is None:
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        try:
            return from_excel(v).date()
        except (ValueError, OverflowError, TypeError, AttributeError):
            return None
    if isinstance(v, date):
        return v
    if isinstance(v, str):
        s = v.strip()
        # In these files, 05/11/2026 means May 11, 2026.
        for fmt in ("%d-%b-%Y", "%m/%d/%Y", "%d/%m/%Y", "%Y-%m-%d"):
            try:
                return datetime.strptime(s, fmt).date()
            except Exception:
                pass
    return None


def get_sheet_by_normalised_name(wb, wanted_names):
    wanted = {x.strip().upper() for x in wanted_names}
    for name in wb.sheetnames:
        if name.strip().upper() in wanted:
            return wb[name]
    return None


def unmerge_and_fill_all(ws):
    for mr in list(ws.merged_cells.ranges):
        value = ws.cell(mr.min_row, mr.min_col).value
        ws.unmerge_cells(str(mr))
        for r in range(mr.min_row, mr.max_row + 1):
            for c in range(mr.min_col, mr.max_col + 1):
                ws.cell(r, c).value = value


def write_number_or_blank(cell, value):
    if value is None or value == 0:
        cell.value = None
    else:
        cell.value = value
        cell.number_format = "#,##0"


def add_log(logs, message):
    if logs is not None:
        logs.append(message)


def cell_ref(row, col):
    return f"{get_column_letter(col)}{row}"


def _debug_value(value, max_len=90):
    """Compact repr() so hidden spaces/characters are visible in error messages."""
    if value is None:
        return "None"
    text = repr(value)
    if len(text) > max_len:
        text = text[: max_len - 3] + "..."
    return text


def _describe_row(ws, row, start_col=1, end_col=None, include_blanks=False, value_at=None):
    if row is None or row < 1 or row > ws.max_row:
        return "(row unavailable)"

    end_col = ws.max_column if end_col is None else min(end_col, ws.max_column)
    start_col = max(1, start_col)
    parts = []
    for c in range(start_col, end_col + 1):
        value = value_at(row, c) if value_at else ws.cell(row, c).value
        if not include_blanks and (value is None or str(value).strip() == ""):
            continue
        parts.append(f"{cell_ref(row, c)}={_debug_value(value)}")
    return ", ".join(parts) if parts else "(no non-empty cells)"


def _closest_cells(ws, target_text, max_results=6, value_at=None):
    """Return a few cells whose text most closely resembles the missing label."""
    target = norm_header(target_text)
    scored = []
    for r, row in enumerate(ws.iter_rows(values_only=True), 1):
        for c, value in enumerate(row, 1):
            if value_at is not None and value is None:
                value = value_at(r, c)
            if value is None or str(value).strip() == "":
                continue
            candidate = norm_header(value)
            score = SequenceMatcher(None, target, candidate).ratio()
            if score >= 0.40:
                scored.append((score, r, c, value))

    scored.sort(key=lambda x: (-x[0], x[1], x[2]))
    if not scored:
        return "(no similar cells found)"

    return "; ".join(
        f"{cell_ref(r, c)}={_debug_value(value)} ({score:.0%} similar)"
        for score, r, c, value in scored[:max_results]
    )


def _fuel_structure_error(
    uploaded_file,
    wb,
    ws,
    *,
    stage,
    expected,
    similar_to=None,
    reference=None,
    header_row=None,
    scan_start_col=None,
    scan_end_col=None,
    value_at=None,
):
    """Build a detailed error that points to the exact workbook area that failed."""
    lines = [
        "Fuel Dashboard structure error",
        f"File: {uploaded_file.name}",
        f"Detected organisation: {_identify_org(uploaded_file.name, wb)}",
        f"Stage: {stage}",
        f"Sheet: {ws.title if ws is not None else '(sheet not found)'}",
        f"Expected: {expected}",
        f"Workbook sheets: {list(wb.sheetnames)}",
    ]

    if reference:
        lines.append(f"Reference: {reference}")

    if ws is not None:
        lines.append(f"Sheet size: {ws.max_row} rows x {ws.max_column} columns")

    if header_row is not None:
        start_col = 1 if scan_start_col is None else scan_start_col
        end_col = ws.max_column if scan_end_col is None else scan_end_col
        lines.append(
            f"Scanned area: row {header_row}, columns "
            f"{get_column_letter(max(1, start_col))}:{get_column_letter(min(end_col, ws.max_column))}"
        )
        lines.append(
            f"Values on scanned row: {_describe_row(ws, header_row, start_col, end_col, include_blanks=True, value_at=value_at)}"
        )

        nearby = []
        for r in range(max(1, header_row - 2), min(ws.max_row, header_row + 2) + 1):
            nearby.append(f"Row {r}: {_describe_row(ws, r, start_col, end_col, value_at=value_at)}")
        lines.append("Nearby rows:\n  " + "\n  ".join(nearby))

    similarity_target = similar_to or expected
    lines.append(f"Closest similar cells: {_closest_cells(ws, similarity_target, value_at=value_at) if ws is not None else '(not available)'}")
    lines.append(
        "Likely causes: the header was renamed, moved to another row, merged differently, "
        "or contains extra/hidden characters."
    )
    return RuntimeError("\n".join(lines))


def _identify_org(filename, workbook=None):
    name = filename.upper()
    for org in ("UNOPS", "WFP"):
        if org in name:
            return org
    if workbook is not None:
        names = {n.strip().upper() for n in workbook.sheetnames}
        if any("UNOPS TOTAL DISTRIBUTION" in n for n in names) or "FUEL RECEIPTS SUMMARY" in names:
            return "UNOPS"
        if any("WFP TOTAL DISTRIBUTION" in n for n in names):
            return "WFP"
        if "SUMMARY" in names:
            summary = get_sheet_by_normalised_name(workbook, ["Summary"])
            for row in summary.iter_rows(min_row=1, max_row=35, max_col=6, values_only=True):
                if any(isinstance(v, str) and v.strip().upper() == "UNOPS" for v in row):
                    return "UNOPS"
    raise RuntimeError(f"Could not identify UNOPS or WFP for '{filename}'. Check the workbook structure.")


def detect_org(filename: str, uploaded_file=None) -> str:
    if uploaded_file is None or any(org in filename.upper() for org in TARGET_SHEETS):
        return _identify_org(filename)
    with closing(load_workbook(BytesIO(uploaded_file.getvalue()), read_only=True, data_only=True, keep_links=False)) as wb:
        return _identify_org(filename, wb)


def _horizontal_values(ws):
    """Read across merged labels, leaving lower rows and the original worksheet untouched."""
    expanded = {}
    for merged in ws.merged_cells.ranges:
        value = ws.cell(merged.min_row, merged.min_col).value
        for c in range(merged.min_col + 1, merged.max_col + 1):
            expanded[merged.min_row, c] = value

    def value_at(r, c):
        if (r, c) in expanded:
            return expanded[r, c]
        return ws.cell(r, c).value

    return value_at


def _read_current_fuel(wb, uploaded_file, org):
    if "Summary" not in wb.sheetnames:
        raise RuntimeError(
            "Fuel Dashboard structure error\n"
            f"File: {uploaded_file.name}\n"
            f"Detected organisation: {org}\n"
            "Stage: Current fuel by station\n"
            'Expected sheet: "Summary"\n'
            f"Workbook sheets found: {list(wb.sheetnames)}"
        )

    ws = wb["Summary"]

    value_at = _horizontal_values(ws)

    header_row, station_col = find_cell_containing(ws, "Fuel Station - Address")
    if not header_row:
        raise _fuel_structure_error(
            uploaded_file,
            wb,
            ws,
            value_at=value_at,
            stage="Current fuel by station",
            expected='cell containing "Fuel Station - Address"',
            similar_to="Fuel Station - Address",
        )

    diesel_col = None
    petrol_col = None
    scan_start_col = station_col + 1
    scan_end_col = ws.max_column

    for c in range(scan_start_col, scan_end_col + 1):
        val = value_at(header_row, c)
        txt = "" if val is None else str(val).strip().upper()
        if txt in ("DIESEL", "DIESEL (LITERS)"):
            diesel_col = c
        elif txt in ("PETROL", "PETROL (LITERS)"):
            petrol_col = c
        if diesel_col and petrol_col:
            break

    if diesel_col is None:
        raise _fuel_structure_error(
            uploaded_file,
            wb,
            ws,
            value_at=value_at,
            stage="Current fuel by station -> Diesel/Petrol header detection",
            expected='header equal to "Diesel" or "Diesel (Liters)"',
            similar_to="Diesel",
            reference=(
                f'"Fuel Station - Address" found at {cell_ref(header_row, station_col)}; '
                "Diesel is expected somewhere to its right on the same row"
            ),
            header_row=header_row,
            scan_start_col=scan_start_col,
            scan_end_col=scan_end_col,
        )

    if petrol_col is None:
        raise _fuel_structure_error(
            uploaded_file,
            wb,
            ws,
            value_at=value_at,
            stage="Current fuel by station -> Diesel/Petrol header detection",
            expected='header equal to "Petrol" or "Petrol (Liters)"',
            similar_to="Petrol",
            reference=(
                f'"Fuel Station - Address" found at {cell_ref(header_row, station_col)}; '
                "Petrol is expected somewhere to its right on the same row"
            ),
            header_row=header_row,
            scan_start_col=scan_start_col,
            scan_end_col=scan_end_col,
        )

    width = ws.max_column
    rows = []
    seen = set()
    for r in range(header_row + 1, ws.max_row + 1):
        row_values = (value_at(r, c) for c in range(1, width + 1))
        row_text = " ".join(str(value) for value in row_values if value is not None).upper()

        if "CURRENT FUEL IN STORAGE" in row_text:
            break

        station = value_at(r, station_col)
        if station is None or str(station).strip() == "":
            continue

        station_name = str(station).strip()
        diesel = safe_float(value_at(r, diesel_col))
        petrol = safe_float(value_at(r, petrol_col))
        key = (org, station_name, diesel, petrol)
        if key in seen:
            continue
        seen.add(key)

        rows.append({
            "Organisation": org,
            "Fuel Station - Address": station_name,
            "Diesel": diesel,
            "Petrol": petrol,
            "Total Fuel": diesel + petrol,
        })

    return rows


def _read_daily_entries(wb, uploaded_file, org, selected_dates):
    ws = get_sheet_by_normalised_name(wb, ["Comulative Fuel Entry", "Cumulative Fuel Entry"])
    if ws is None:
        raise RuntimeError(f'"Comulative Fuel Entry" not found in "{uploaded_file.name}".')

    value_at = _horizontal_values(ws)

    date_row, date_col = find_cell_containing(ws, "Date")
    diesel_row, diesel_col = find_cell_containing(ws, "Total Diesel Received")
    benzene_row, benzene_col = find_cell_containing(ws, "Total Benzene Received")

    if not date_row:
        raise RuntimeError(f'"Date" not found in "{uploaded_file.name}".')
    if not diesel_row:
        raise RuntimeError(f'"Total Diesel Received" not found in "{uploaded_file.name}".')
    if not benzene_row:
        raise RuntimeError(f'"Total Benzene Received" not found in "{uploaded_file.name}".')

    # Find the Note column near the Diesel/Benzene amount columns.
    # In the WFP file, this Note header is merged with row 2, so we scan the nearby header area.
    note_col = None
    if org == "WFP":
        min_header_row = max(1, min(date_row, diesel_row, benzene_row) - 2)
        max_header_row = min(ws.max_row, max(date_row, diesel_row, benzene_row) + 2)

        # Prefer columns close to Diesel/Benzene
        search_start_col = max(1, min(diesel_col, benzene_col) - 2)
        search_end_col = min(ws.max_column, max(diesel_col, benzene_col) + 5)

        for r in range(min_header_row, max_header_row + 1):
            for c in range(search_start_col, search_end_col + 1):
                val = value_at(r, c)
                txt = "" if val is None else str(val).strip().upper()
                if txt == "NOTE":
                    note_col = c
                    break
            if note_col:
                break

    selected_dates = set(selected_dates)
    results = {}

    for r in range(date_row + 1, ws.max_row + 1):
        current_date = normalise_date(value_at(r, date_col))
        if current_date is None or current_date not in selected_dates:
            continue

        # WFP-only exclusion:
        # If Note says "From UNOPS", ignore this row completely.
        if org == "WFP" and note_col is not None:
            note_value = value_at(r, note_col)
            note_text = "" if note_value is None else str(note_value).strip().upper()

            if "FROM UNOPS" in note_text:
                continue

        diesel = safe_float(value_at(r, diesel_col))
        benzene = safe_float(value_at(r, benzene_col))

        # Add to existing date total instead of replacing it,
        # in case the same date appears more than once.
        results[current_date] = results.get(current_date, 0.0) + diesel + benzene

    return org, results


def _read_fuel_used(wb, uploaded_file, org, selected_dates, logs=None):
    add_log(logs, f"========== Fuel Used | {uploaded_file.name} | detected org: {org} ==========")

    ws = get_sheet_by_normalised_name(
        wb,
        [
            "Comulative Ditribution Summary",
            "Comulative Distribution Summary",
            "Cumulative Ditribution Summary",
            "Cumulative Distribution Summary",
        ],
    )
    if ws is None:
        raise RuntimeError(f'"Comulative Ditribution Summary" not found in "{uploaded_file.name}".')

    date_row, date_col = find_cell_exact(ws, "Date")
    sum_row, sum_col = find_cell_exact(ws, "Sum")
    if not date_row:
        raise RuntimeError(f'Exact cell "Date" not found in "{uploaded_file.name}".')
    if not sum_row:
        raise RuntimeError(f'Exact cell "Sum" not found in "{uploaded_file.name}".')

    sum_start_col = sum_col
    sum_end_col = sum_col
    for mr in list(ws.merged_cells.ranges):
        if mr.min_row <= sum_row <= mr.max_row and mr.min_col <= sum_col <= mr.max_col:
            sum_start_col = mr.min_col
            sum_end_col = mr.max_col
            break

    add_log(logs, f"Exact Date found at {cell_ref(date_row, date_col)}")
    add_log(logs, f"Exact Sum found at {cell_ref(sum_row, sum_col)}")
    add_log(logs, f"Using ONLY Sum section columns: {get_column_letter(sum_start_col)}:{get_column_letter(sum_end_col)}")

    unmerge_and_fill_all(ws)
    selected_dates = set(selected_dates)
    add_log(logs, f"Selected dates: {[d.strftime('%m/%d/%Y') for d in selected_dates]}")

    total_used = 0.0
    active_diesel_col = None
    active_petrol_col = None
    carried_date = None
    carried_date_row = -10

    for r in range(sum_row + 1, ws.max_row + 1):
        row_date = normalise_date(ws.cell(r, date_col).value)
        # WFP has one date per daily row and undated grand-total/footer rows.
        # NEVER carry a WFP date into later rows: that counted its entire
        # cumulative total as part of the final selected day.
        # UNOPS may put the numbers on the immediately following row only.
        if org == "WFP":
            current_date = row_date
        else:
            if row_date is not None:
                carried_date = row_date
                carried_date_row = r
            current_date = (row_date if row_date is not None else
                            carried_date if r == carried_date_row + 1 else None)
        header_pair_found = False

        for c in range(sum_start_col, sum_end_col):
            left_val = ws.cell(r, c).value
            right_val = ws.cell(r, c + 1).value
            left_txt = "" if left_val is None else str(left_val).strip().upper()
            right_txt = "" if right_val is None else str(right_val).strip().upper()

            if left_txt in ("DIESEL", "DIESIL") and right_txt == "PETROL":
                active_diesel_col = c
                active_petrol_col = c + 1
                header_pair_found = True
                add_log(logs, f"SUM header found row {r}: {cell_ref(r, c)}='{left_val}', {cell_ref(r, c + 1)}='{right_val}'")
                break

        if header_pair_found:
            continue
        if current_date not in selected_dates:
            continue
        if active_diesel_col is None or active_petrol_col is None:
            add_log(logs, f"Skipped row {r}: date matched but no Diesel/Petrol header under exact Sum yet.")
            continue

        diesel_raw = ws.cell(r, active_diesel_col).value
        petrol_raw = ws.cell(r, active_petrol_col).value
        diesel_txt = "" if diesel_raw is None else str(diesel_raw).strip().upper()
        petrol_txt = "" if petrol_raw is None else str(petrol_raw).strip().upper()
        if diesel_txt in ("DIESEL", "DIESIL") or petrol_txt == "PETROL":
            continue

        diesel_num = safe_float(diesel_raw)
        petrol_num = safe_float(petrol_raw)
        combined = diesel_num + petrol_num
        total_used += combined
        add_log(logs, f"{org} SUM value row {r}, date {current_date.strftime('%m/%d/%Y')}: Diesel {cell_ref(r, active_diesel_col)}={diesel_raw} -> {diesel_num}, Petrol {cell_ref(r, active_petrol_col)}={petrol_raw} -> {petrol_num}, combined={combined}")

    add_log(logs, f"FINAL USED TOTAL for {org}: {total_used}")
    add_log(logs, "")
    return org, total_used


def _read_storage(wb, uploaded_file, org, logs=None):
    add_log(logs, f"========== Fuel Storage | {uploaded_file.name} | detected org: {org} ==========")

    if "Summary" not in wb.sheetnames:
        raise RuntimeError(f'"Summary" sheet not found in "{uploaded_file.name}".')

    ws = wb["Summary"]
    # Only anchors have values on the untouched Summary sheet.
    storage_row, storage_col = find_cell_containing(ws, "Current Fuel In storage")
    if not storage_row:
        raise RuntimeError(f'"Current Fuel In storage (L)" not found in "{uploaded_file.name}".')

    add_log(logs, f"Current Fuel In storage found at {cell_ref(storage_row, storage_col)}")
    diesel_col = None
    for c in range(storage_col + 1, ws.max_column + 1):
        raw_value = ws.cell(storage_row, c).value
        add_log(logs, f"Storage scan {cell_ref(storage_row, c)} raw='{raw_value}'")
        if raw_value is None or str(raw_value).strip() == "":
            continue
        diesel_col = c
        break

    if diesel_col is None:
        add_log(logs, f"No diesel value found for {org}. Storage total = 0")
        return org, 0

    petrol_col = diesel_col + 1
    diesel_raw = ws.cell(storage_row, diesel_col).value
    petrol_raw = ws.cell(storage_row, petrol_col).value if petrol_col <= ws.max_column else None
    diesel = safe_float(diesel_raw)
    petrol = safe_float(petrol_raw)

    add_log(logs, f"Storage values for {org}: Diesel {cell_ref(storage_row, diesel_col)}={diesel_raw} -> {diesel}, Petrol {cell_ref(storage_row, petrol_col)}={petrol_raw} -> {petrol}, combined={diesel + petrol}")
    add_log(logs, f"FINAL STORAGE TOTAL for {org}: {diesel + petrol}")
    add_log(logs, "")
    return org, diesel + petrol


@contextmanager
def _open_fuel_source(upload):
    with closing(load_workbook(BytesIO(upload.getvalue()), data_only=True, keep_links=False)) as wb:
        yield wb, _identify_org(upload.name, wb)


def extract_current_fuel_from_file(uploaded_file):
    with _open_fuel_source(uploaded_file) as (wb, org):
        return _read_current_fuel(wb, uploaded_file, org)


def extract_daily_fuel_entries(uploaded_file, selected_dates):
    with _open_fuel_source(uploaded_file) as (wb, org):
        return _read_daily_entries(wb, uploaded_file, org, selected_dates)


def extract_fuel_used(uploaded_file, selected_dates, logs=None):
    with _open_fuel_source(uploaded_file) as (wb, org):
        return _read_fuel_used(wb, uploaded_file, org, selected_dates, logs)


def extract_fuel_storage(uploaded_file, logs=None):
    with _open_fuel_source(uploaded_file) as (wb, org):
        return _read_storage(wb, uploaded_file, org, logs)


@contextmanager
def _fuel_step(upload, org, description):
    try:
        yield
    except Exception as exc:
        raise RuntimeError(
            f"Fuel Dashboard processing failed\nFile: {upload.name}\nDetected organisation: {org}\n"
            f"Sub-step: {description}\n\nDetails:\n{exc}"
        ) from exc


def add_fuel_dashboard_sheet(
    wb,
    latest_uploads,
    selected_dates,
    logs=None,
    sheet_name="Fuel Dashboard",
):
    """Adds the Fuel Dashboard as a new sheet to an existing workbook."""
    all_rows = []
    daily_entries = {}
    fuel_used_by_org = {}
    fuel_storage_by_org = {}

    for upload in latest_uploads:
        with _open_fuel_source(upload) as (source, org):
            with _fuel_step(upload, org, "1/4 - Read current fuel by station"):
                all_rows.extend(_read_current_fuel(source, upload, org))
            with _fuel_step(upload, org, "2/4 - Read cumulative fuel entries"):
                _, daily_entries[org] = _read_daily_entries(source, upload, org, selected_dates)
            with _fuel_step(upload, org, "3/4 - Calculate fuel used"):
                _, fuel_used_by_org[org] = _read_fuel_used(source, upload, org, selected_dates, logs)
            with _fuel_step(upload, org, "4/4 - Read current fuel in storage"):
                _, fuel_storage_by_org[org] = _read_storage(source, upload, org, logs)

    if sheet_name in wb.sheetnames:
        wb.remove(wb[sheet_name])

    ws = wb.create_sheet(sheet_name)

    header_fill = PatternFill(fill_type="solid", fgColor="000000")
    header_font = Font(color="FFFFFF", bold=True)
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center")

    # SECTION 1 - Hebrew Fuel Summary Table A1:D4
    headers3 = ["דלק במאגרים", "דלק שנוצל", "דלק שנכנס", ""]
    header_colours = {1: "A3B18A", 2: "F4B183", 3: "A8D5C8", 4: "FFFFFF"}

    for c, h in enumerate(headers3, start=1):
        cell = ws.cell(1, c)
        cell.value = h
        cell.fill = PatternFill(fill_type="solid", fgColor=header_colours[c])
        cell.font = Font(bold=True, size=14)
        cell.border = border
        cell.alignment = center

    total_entered_wfp = sum(daily_entries.get("WFP", {}).values())
    total_entered_unops = sum(daily_entries.get("UNOPS", {}).values())

    table_rows = [
        {
            "org": "UNOPS",
            "entered": total_entered_unops,
            "used": fuel_used_by_org.get("UNOPS", 0),
            "storage": fuel_storage_by_org.get("UNOPS", 0),
            "org_fill": "F47C25",
        },
        {
            "org": "WFP",
            "entered": total_entered_wfp,
            "used": fuel_used_by_org.get("WFP", 0),
            "storage": fuel_storage_by_org.get("WFP", 0),
            "org_fill": "00A9E0",
        },
        {
            "org": 'סה"כ',
            "entered": total_entered_unops + total_entered_wfp,
            "used": fuel_used_by_org.get("UNOPS", 0) + fuel_used_by_org.get("WFP", 0),
            "storage": fuel_storage_by_org.get("UNOPS", 0) + fuel_storage_by_org.get("WFP", 0),
            "org_fill": "FFFFFF",
        },
    ]

    for r, row in enumerate(table_rows, start=2):
        write_number_or_blank(ws.cell(r, 1), row["storage"])
        write_number_or_blank(ws.cell(r, 2), row["used"])
        write_number_or_blank(ws.cell(r, 3), row["entered"])
        org_cell = ws.cell(r, 4)
        org_cell.value = row["org"]
        org_cell.fill = PatternFill(fill_type="solid", fgColor=row["org_fill"])

        for c in range(1, 5):
            cell = ws.cell(r, c)
            cell.border = border
            cell.font = Font(bold=True, size=14)
            cell.alignment = center
            cell.number_format = "#,##0"

    # SECTION 2 - Daily Fuel Entry A7:C...
    daily_start_row = 7
    daily_start_col = 1
    for i, h in enumerate(["Date", "WFP", "UNOPS"]):
        cell = ws.cell(daily_start_row, daily_start_col + i)
        cell.value = h
        cell.fill = header_fill
        cell.font = header_font
        cell.border = border
        cell.alignment = center

    row_num = daily_start_row + 1
    for d in selected_dates:
        ws.cell(row_num, daily_start_col).value = d.strftime("%d/%m/%Y")
        write_number_or_blank(ws.cell(row_num, daily_start_col + 1), daily_entries.get("WFP", {}).get(d, 0))
        write_number_or_blank(ws.cell(row_num, daily_start_col + 2), daily_entries.get("UNOPS", {}).get(d, 0))
        for c in range(daily_start_col, daily_start_col + 3):
            cell = ws.cell(row_num, c)
            cell.border = border
            cell.alignment = center
            cell.number_format = "#,##0"
        row_num += 1

    ws.cell(row_num, daily_start_col).value = "Total"
    ws.cell(row_num, daily_start_col + 1).value = f"=SUM(B{daily_start_row + 1}:B{row_num - 1})"
    ws.cell(row_num, daily_start_col + 2).value = f"=SUM(C{daily_start_row + 1}:C{row_num - 1})"
    for c in range(daily_start_col, daily_start_col + 3):
        cell = ws.cell(row_num, c)
        cell.border = border
        cell.font = Font(bold=True)
        cell.alignment = center
        cell.number_format = "#,##0"

    # SECTION 3 - Fuel Station Overview E7:I...
    station_start_row = 7
    station_start_col = 5
    station_headers = ["Organisation", "Fuel Station - Address", "Diesel", "Petrol", "Total Fuel"]
    for i, h in enumerate(station_headers):
        cell = ws.cell(station_start_row, station_start_col + i)
        cell.value = h
        cell.fill = header_fill
        cell.font = header_font
        cell.border = border
        cell.alignment = center

    for r_idx, row in enumerate(all_rows, start=station_start_row + 1):
        ws.cell(r_idx, station_start_col).value = row["Organisation"]
        ws.cell(r_idx, station_start_col + 1).value = row["Fuel Station - Address"]
        write_number_or_blank(ws.cell(r_idx, station_start_col + 2), row["Diesel"])
        write_number_or_blank(ws.cell(r_idx, station_start_col + 3), row["Petrol"])
        write_number_or_blank(ws.cell(r_idx, station_start_col + 4), row["Total Fuel"])

        for c in range(station_start_col, station_start_col + 5):
            cell = ws.cell(r_idx, c)
            cell.border = border
            cell.alignment = center
            cell.number_format = "#,##0"

    widths = {"A": 18, "B": 18, "C": 18, "D": 18, "E": 18, "F": 32, "G": 15, "H": 15, "I": 15}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    for r in range(1, ws.max_row + 1):
        ws.row_dimensions[r].height = 24
    for r in range(1, 5):
        ws.row_dimensions[r].height = 35
    ws.freeze_panes = "A7"

    return wb


# Approval matching


def _norm_prefix(value) -> str:
    if value is None:
        return ""
    s = str(value).strip().upper()
    s = s.replace("’", "'").replace("`", "'")
    s = s.replace("‐", "-").replace("‑", "-").replace("–", "-").replace("—", "-")
    s = re.sub(r"[^A-Z0-9א-ת]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


@lru_cache(maxsize=32)
def _prefix_rules(prefixes):
    names = frozenset(_norm_prefix(p) for p in prefixes if p) - {""}
    patterns = [re.escape(p) for p in sorted(names, key=len, reverse=True)]
    literal = re.compile(r"^\s*(?:" + "|".join(patterns) + r")\s*[-–—:]\s*", re.IGNORECASE) if patterns else None
    return names, literal


def strip_generated_prefix(value, prefixes=None) -> str:
    """Strip successive category prefixes only when they end in a dash or colon."""
    text = _label(value)
    if not text:
        return ""
    names, literal = _prefix_rules(frozenset(prefixes or BASE_GENERATED_PREFIXES))
    while text:
        parts = re.split(r"\s*[-–—:]\s*", text, maxsplit=1)
        if len(parts) == 1:
            break
        if _norm_prefix(parts[0]) in names:
            text = parts[1].strip()
            continue
        match = literal.match(text) if literal is not None else None
        if match is None:
            break
        text = text[match.end():].strip()
    return text


NAME_REPLACEMENTS = {
    "INTL": "INTERNATIONAL",
    "INT": "INTERNATIONAL",
    "ORG": "ORGANIZATION",
    "ASSOC": "ASSOCIATION",
    "PROGRAMME": "PROGRAM",
    "SOCAITY": "SOCIETY",
    "CLOBAL": "GLOBAL",
    "PH": "PROJECT HOPE",
    "ACF": "ACTION AGAINST HUNGER",
    "SI": "SOLIDARITES INTERNATIONALE",
    "GC": "GLOBAL COMMUNITIES",
    "MG": "MED GLOBAL",
    "PIB": "PALESTINE ISLAMIC BANK",
}


def norm_text(value, prefixes=None) -> str:
    """Normalise names while ignoring cleaner-added category prefixes."""
    if value is None:
        return ""

    s = strip_generated_prefix(value, prefixes=prefixes).upper()
    s = s.replace("’", "'").replace("`", "'")
    s = s.replace("‐", "-").replace("‑", "-").replace("–", "-").replace("—", "-")
    s = s.replace("&", " AND ")

    # Usually acronyms in brackets differ between files; remove them to match the full name.
    s = re.sub(r"\([^)]*\)", " ", s)
    s = re.sub(r"[^A-Z0-9א-ת]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()

    words = [NAME_REPLACEMENTS.get(w, w) for w in s.split()]
    return " ".join(words)


def _join_key(*parts):
    return " | ".join(part for part in parts if part)


def compact_key(*parts, prefixes=None) -> str:
    return _join_key(*(norm_text(part, prefixes=prefixes) for part in parts))


def clean_header(value) -> str:
    return norm_text(value).replace(" ", "")


def best_fuzzy_match(query_key: str, choices: list[str]):
    """Keep the original five-candidate shortlist, score and threshold."""
    if not query_key or not choices:
        return None, 0.0

    shortlist = get_close_matches(
        query_key,
        choices,
        n=5,
        cutoff=0.50,
    )

    if not shortlist:
        return None, 0.0

    best_key = None
    best_score = 0.0

    for choice in shortlist:
        score = SequenceMatcher(
            None,
            query_key,
            choice,
        ).ratio()

        if score > best_score:
            best_score = score
            best_key = choice

    if best_score >= FUZZY_THRESHOLD:
        return best_key, best_score

    return None, best_score


def get_sheet_by_names(wb, possible_names):
    name_map = {name.strip().upper(): name for name in wb.sheetnames}
    for wanted in possible_names:
        real = name_map.get(wanted.strip().upper())
        if real:
            return wb[real]
    return None


def find_header_row_and_columns(ws, wanted_headers, max_scan_rows=20):
    aliases = {key: {clean_header(item) for item in names} for key, names in wanted_headers.items()}
    required = set(wanted_headers) - {"Status"}
    limit = min(max_scan_rows, ws.max_row or max_scan_rows)
    for r, row in enumerate(ws.iter_rows(max_row=limit, values_only=True), 1):
        columns = {}
        for c, value in enumerate(row, 1):
            if value is None:
                continue
            header = clean_header(value)
            for key, options in aliases.items():
                if header and header in options:
                    columns[key] = c
        if required.issubset(columns):
            return r, columns
    return None, {}


def status_fill(status):
    txt = "" if status is None else str(status).strip().upper()
    for needle, colour in STATUS_FILL_RULES:
        if needle.upper() in txt:
            return PatternFill(fill_type="solid", fgColor=colour)
    return PatternFill(fill_type="solid", fgColor="D9D9D9")


def _shift_merge_range_for_insert(range_string: str, insert_col: int) -> str:
    """Return the same merge range after a column insertion, preserving existing merges."""
    min_col, min_row, max_col, max_row = range_boundaries(range_string)

    # If the full merged range is to the right of the inserted column, move it right.
    if min_col >= insert_col:
        min_col += 1
        max_col += 1
    # If a merged range crosses the insertion point, expand it by one column.
    elif min_col < insert_col <= max_col:
        max_col += 1

    return (
        f"{get_column_letter(min_col)}{min_row}:"
        f"{get_column_letter(max_col)}{max_row}"
    )


def _insert_column_preserve_merges(ws, insert_col: int):
    """openpyxl does not reliably shift merged ranges on insert_cols; do it manually."""
    if insert_col > ws.max_column:
        return  # An appended column cannot shift any cells or merges.
    merge_ranges = [str(mr) for mr in ws.merged_cells.ranges]
    for rng in merge_ranges:
        ws.unmerge_cells(rng)

    ws.insert_cols(insert_col, 1)

    for rng in merge_ranges:
        shifted = _shift_merge_range_for_insert(rng, insert_col)
        try:
            ws.merge_cells(shifted)
        except Exception:
            # If a corrupted/overlapping merge exists in an input file, skip only that merge.
            pass


def _last_text_header_col(ws, header_row, exclude_status=True):
    """Return the last header column that contains visible text.

    This is intentionally based on the header row, not ws.max_column, because
    generated Fuels Summary workbooks can contain charts/empty formatted columns
    to the right.
    """
    last_col = 1
    for c in range(1, ws.max_column + 1):
        value = ws.cell(header_row, c).value
        if value is None or str(value).strip() == "":
            continue
        if exclude_status and str(value).strip() == STATUS_COL_NAME:
            continue
        last_col = max(last_col, c)
    return last_col


def get_or_insert_status_column(ws, header_row, description_col=None):
    """Add סטטוס AFTER the last real header column.

    Earlier versions inserted סטטוס next to DESCRIPTION. In this workbook that
    shifts the Unified Fuel / Total Sum Per Category area and can make Unified
    Fuel appear merged by cluster. Appending סטטוס after the last text header
    keeps all calculation columns exactly where they are.
    """
    existing_status_cols = [
        c for c in range(1, ws.max_column + 1)
        if str(ws.cell(header_row, c).value).strip() == STATUS_COL_NAME
    ]

    last_text_col = _last_text_header_col(ws, header_row, exclude_status=True)
    desired_status_col = last_text_col + 1

    # If the workbook was already created with this fixed version, reuse the column.
    for c in existing_status_cols:
        if c == desired_status_col or c > last_text_col:
            return c, False

    # If a previous/bad version inserted סטטוס inside the calculation columns, stop
    # rather than producing a misleading dashboard from already-shifted/merged data.
    if existing_status_cols:
        raise RuntimeError(
            "This workbook already contains a סטטוס column inside the calculation area. "
            "Please upload the original Fuels Summary workbook and rerun this fixed version."
        )

    status_col = desired_status_col
    _insert_column_preserve_merges(ws, status_col)

    status_header = ws.cell(header_row, status_col)
    status_header.value = STATUS_COL_NAME
    style_source_col = description_col if description_col else last_text_col
    copy_cell_style(ws.cell(header_row, style_source_col), status_header)
    status_header.alignment = Alignment(horizontal="center", vertical="center")
    ws.column_dimensions[get_column_letter(status_col)].width = 22
    return status_col, True


def read_approval_rows(approval_file):
    rows = []
    prefixes = set(BASE_GENERATED_PREFIXES)
    with closing(load_workbook(BytesIO(approval_file.getvalue()), read_only=True, data_only=True, keep_links=False)) as wb:
        for org, names in APPROVAL_SHEETS.items():
            ws = get_sheet_by_names(wb, names)
            if ws is None:
                continue
            # Some Excel writers omit the used range or record it as A1.
            ws.reset_dimensions()
            category = "Cluster" if org == "UNOPS" else "Intervention"
            headers = {category: [category], "Description": ["Description"], "Status": [STATUS_COL_NAME, "Status"]}
            if org == "UNOPS":
                headers["Agency"] = ["Agency"]
            header_row, cols = find_header_row_and_columns(ws, headers)
            if header_row is None:
                continue
            data = ws.iter_rows(min_row=header_row + 1, max_col=max(cols.values()), values_only=True)
            for r, values in enumerate(data, header_row + 1):
                cluster = values[cols[category] - 1]
                agency = values[cols["Agency"] - 1] if "Agency" in cols else None
                desc = values[cols["Description"] - 1]
                status = values[cols["Status"] - 1] if "Status" in cols else None
                if not any((cluster, agency, desc, status)):
                    continue
                prefixes.add(str(cluster).strip())
                rows.append({
                    "Source Org": org, "Approved Cluster/Intervention": cluster,
                    "Approved Agency": agency, "Approved Description": desc,
                    "Review Status": status or "Unknown", "Approval Sheet Row": r,
                })
    if not rows:
        raise RuntimeError("No approval rows were found in the approval workbook.")
    return rows, prefixes


def build_approval_index(approval_file):
    rows, prefixes = read_approval_rows(approval_file)
    rules = frozenset(prefixes)
    index = {org: {"exact": {}, "choices": [], "fuzzy_cache": {}} for org in APPROVAL_SHEETS}
    for row in rows:
        org = row["Source Org"]
        cluster = norm_text(row["Approved Cluster/Intervention"], rules)
        desc = norm_text(row["Approved Description"], rules)
        if org == "UNOPS":
            agency = norm_text(row["Approved Agency"], rules)
            keys = {_join_key(agency, desc), agency, desc, _join_key(cluster, agency), _join_key(cluster, agency, desc)}
        else:
            keys = {_join_key(cluster, desc), desc, cluster}
        for key in keys:
            if key:
                index[org]["exact"].setdefault(key, row)
    for group in index.values():
        group["choices"] = list(group["exact"])
    return index, prefixes, pd.DataFrame(rows)


def detect_row_source_org(cluster_value) -> str:
    # In the final Fuels Summary workbook, WFP rows have Cluster = WFP.
    return "WFP" if norm_text(cluster_value) == "WFP" else "UNOPS"


def review_summary_row(cluster, agency, desc, approval_index, prefixes):
    # Final Fuels Summary can contain internal UN-agency operational rows under
    # Cluster = UN Agencies. These should be reviewed as approved UN agency fuel
    # instead of being matched as a regular NGO/organisation row.
    agency_compact = _compact_alpha_num(agency)
    if norm_text(cluster) == "UN AGENCIES" and (
        _is_wfp_un_agency(agency) or "OHCHR" in agency_compact
    ):
        return {
            "Detected Org": "UNOPS",
            "Review Status": 'סוכנויות או"ם מאושרות',
            "Match Type": "Rule",
            "Match Score": 1.0,
            "Matched Cluster/Intervention": cluster,
            "Matched Agency": agency,
            "Matched Description": desc,
            "Approval Sheet Row": None,
        }

    org = detect_row_source_org(cluster)
    exact = approval_index[org]["exact"]
    choices = approval_index[org]["choices"]
    fuzzy_cache = approval_index[org]["fuzzy_cache"]

    rules = frozenset(prefixes or BASE_GENERATED_PREFIXES)
    agency_key = norm_text(agency, rules)
    desc_key = norm_text(desc, rules)
    full = _join_key(agency_key, desc_key)
    if org == "WFP":
        candidates = [full, desc_key, agency_key]
    else:
        cluster_key = norm_text(cluster, rules)
        candidates = [full, agency_key, desc_key, _join_key(cluster_key, agency_key),
                      _join_key(cluster_key, agency_key, desc_key)]
    candidate_keys = list(dict.fromkeys(candidates))

    for key in candidate_keys:
        if key and key in exact:
            match = exact[key]
            return {
                "Detected Org": org,
                "Review Status": match["Review Status"],
                "Match Type": "Exact",
                "Match Score": 1.0,
                "Matched Cluster/Intervention": match["Approved Cluster/Intervention"],
                "Matched Agency": match["Approved Agency"],
                "Matched Description": match["Approved Description"],
                "Approval Sheet Row": match["Approval Sheet Row"],
            }

    # Fuzzy match using all candidate keys. This handles small typos and non-exact formats.
    best_key = None
    best_score = 0.0
    for query_key in candidate_keys:
        if not query_key:
            continue

        cached = fuzzy_cache.get(query_key)

        if cached is None:
            cached = best_fuzzy_match(
                query_key,
                choices,
            )
            fuzzy_cache[query_key] = cached

        fuzzy_key, score = cached

        if score > best_score:
            best_score = score
            best_key = fuzzy_key

    if best_key:
        match = exact[best_key]
        return {
            "Detected Org": org,
            "Review Status": match["Review Status"],
            "Match Type": "Fuzzy",
            "Match Score": round(best_score, 3),
            "Matched Cluster/Intervention": match["Approved Cluster/Intervention"],
            "Matched Agency": match["Approved Agency"],
            "Matched Description": match["Approved Description"],
            "Approval Sheet Row": match["Approval Sheet Row"],
        }

    no_match_status = UNKNOWN_ENTITIES_STATUS if norm_text(cluster) == "INGOS" else "Not Found"

    return {
        "Detected Org": org,
        "Review Status": no_match_status,
        "Match Type": "No Match",
        "Match Score": round(best_score, 3),
        "Matched Cluster/Intervention": None,
        "Matched Agency": None,
        "Matched Description": None,
        "Approval Sheet Row": None,
    }


# Final worksheet formatting


def _has_hebrew(value) -> bool:
    if value is None:
        return False
    return any("\u0590" <= ch <= "\u05FF" for ch in str(value))


def _is_number_like(value) -> bool:
    if value is None:
        return False
    if isinstance(value, (int, float)):
        return True
    txt = str(value).strip().replace(",", "").replace("%", "")
    if txt == "":
        return False
    try:
        float(txt)
        return True
    except Exception:
        return False


def _font_with(cell, name=None, size=None, color=None, bold=None, italic=None):
    f = cell.font or Font()
    return Font(
        name=name if name is not None else f.name,
        sz=size if size is not None else f.sz,
        b=bold if bold is not None else f.bold,
        i=italic if italic is not None else f.italic,
        u=f.underline,
        strike=f.strike,
        color=color if color is not None else f.color,
        vertAlign=f.vertAlign,
        outline=f.outline,
        shadow=f.shadow,
        condense=f.condense,
        extend=f.extend,
        charset=f.charset,
        family=f.family,
        scheme=f.scheme,
    )


def _thin_black_border():
    side = Side(style="thin", color="000000")
    return Border(left=side, right=side, top=side, bottom=side)


def _safe_unmerge_intersecting_columns(ws, cols_to_delete):
    cols = set(cols_to_delete)
    for mr in list(ws.merged_cells.ranges):
        if any(mr.min_col <= c <= mr.max_col for c in cols):
            ws.unmerge_cells(str(mr))


def _delete_columns_by_headers(ws, header_row, header_names):
    """Delete columns whose header text exactly matches any value in header_names."""
    names = {str(x).strip().upper() for x in header_names}
    cols = []
    for c in range(1, ws.max_column + 1):
        v = ws.cell(header_row, c).value
        if v is not None and str(v).strip().upper() in names:
            cols.append(c)
    if not cols:
        return
    _safe_unmerge_intersecting_columns(ws, cols)
    for c in sorted(cols, reverse=True):
        ws.delete_cols(c, 1)


def _last_nonempty_header_col(ws, header_row=1):
    last = 1
    for c in range(1, ws.max_column + 1):
        v = ws.cell(header_row, c).value
        if v is not None and str(v).strip() != "":
            last = c
    return last


def _set_autofilter_to_real_headers(ws, header_row=1):
    last = _last_nonempty_header_col(ws, header_row)
    if last >= 1 and ws.max_row >= header_row:
        ws.auto_filter.ref = f"A{header_row}:{get_column_letter(last)}{ws.max_row}"
    else:
        ws.auto_filter.ref = None


def _find_header_col_exact(ws, header_text, header_row=1):
    target = str(header_text).strip().upper()
    for c in range(1, ws.max_column + 1):
        v = ws.cell(header_row, c).value
        if v is not None and str(v).strip().upper() == target:
            return c
    return None


def _auto_fit_columns(ws, min_width=10, max_width=55):
    for c in range(1, ws.max_column + 1):
        letter = get_column_letter(c)
        max_len = 0
        for r in range(1, ws.max_row + 1):
            value = ws.cell(r, c).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        if max_len:
            ws.column_dimensions[letter].width = min(max(max_len + 3, min_width), max_width)


def _row_has_any_real_value(ws, row_num, max_col=None):
    max_col = max_col or ws.max_column
    for c in range(1, max_col + 1):
        value = ws.cell(row_num, c).value
        if value is not None and str(value).strip() != "":
            return True
    return False


def remerge_total_sum_per_category(ws, header_row=1):
    """Rebuild consecutive category merges from row-level Unified Fuel."""
    total_col = _find_header_col_exact(ws, "Total Sum Per Category", header_row)
    if total_col is None:
        return
    cluster_col = (_find_header_col_exact(ws, "Cluster", header_row)
                   or _find_header_col_exact(ws, "INTERVENTION", header_row) or 1)
    fuel_col = _find_header_col_exact(ws, "Unified Fuel", header_row)
    last, width = ws.max_row, ws.max_column
    for merged in list(ws.merged_cells.ranges):
        if merged.min_col <= total_col <= merged.max_col and merged.max_row > header_row:
            anchor = ws.cell(merged.min_row, total_col)
            value, style = anchor.value, copy(anchor._style)
            ws.unmerge_cells(str(merged))
            for r in range(merged.min_row, merged.max_row + 1):
                cell = _ensure_real_cell(ws, r, total_col)
                cell.value = value
                cell._style = copy(style)

    if fuel_col is not None:
        totals = {}
        for r in range(header_row + 1, last + 1):
            if _row_has_any_real_value(ws, r, width):
                key = _label(ws.cell(r, cluster_col).value)
                totals[key] = totals.get(key, 0.0) + safe_float(ws.cell(r, fuel_col).value)
        for r in range(header_row + 1, last + 1):
            key = _label(ws.cell(r, cluster_col).value)
            if key in totals:
                _ensure_real_cell(ws, r, total_col).value = totals[key]

    first = None
    previous = None
    for r in range(header_row + 1, last + 2):
        populated = r <= last and _row_has_any_real_value(ws, r, width)
        category = _label(ws.cell(r, cluster_col).value).upper() if populated else None
        if first is not None and category != previous:
            _merge_column_group(ws, total_col, first, r - 1)
            first = None
        if first is None and populated:
            first = r
        previous = category


def style_distribution_summary_final(ws):
    """Final requested styling for Distribution Summary."""
    # Reapply the merge first, then style; no borders are applied to the
    # Total Sum Per Category column, but all other requested styling is applied.
    remerge_total_sum_per_category(ws, header_row=1)

    border = _thin_black_border()
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    total_cat_col = _find_header_col_exact(ws, "Total Sum Per Category", header_row=1)

    fonts = {}
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell, MergedCell):
                continue

            # Style only cells that actually contain values/text; do not format the full sheet.
            if cell.value is None or str(cell.value).strip() == "":
                continue

            cell.alignment = center
            is_header = cell.row == 1
            font_key = (cell._style.fontId, is_header)
            if font_key not in fonts:
                fonts[font_key] = _font_with(cell, name="Calibri", size=11,
                                             color="FFFFFF" if is_header else "000000",
                                             bold=is_header, italic=False)
            cell.font = fonts[font_key]

            # Apply all borders only outside Total Sum Per Category.
            if total_cat_col is None or cell.column != total_cat_col:
                cell.border = border

    _set_autofilter_to_real_headers(ws, header_row=1)


def style_sector_summary_final(ws):
    """Final requested styling for Sector Summary."""
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    un_agencies_fill = PatternFill(fill_type="solid", fgColor="013D59")

    for row in ws.iter_rows():
        for cell in row:
            if cell.value is None:
                continue
            cell.alignment = center

            color = "FFFFFF" if cell.row == 1 else "000000"
            font_name = "Aduma" if (_has_hebrew(cell.value) or _is_number_like(cell.value) or (isinstance(cell.value, str) and cell.value.startswith("="))) else None
            cell.font = _font_with(cell, name=font_name, size=20, color=color, bold=False, italic=False)

    # Header row text should be white.
    for c in range(1, ws.max_column + 1):
        cell = ws.cell(1, c)
        if cell.value is not None and str(cell.value).strip() != "":
            cell.font = _font_with(cell, color="FFFFFF", bold=False, italic=False)

    # סוכנויות או"ם row: dark blue fill #013D59 and white text/value.
    for r in range(1, ws.max_row + 1):
        val = ws.cell(r, 1).value
        if val is not None and str(val).strip() == 'סוכנויות או"ם':
            for c in range(1, min(ws.max_column, 3) + 1):
                cell = ws.cell(r, c)
                cell.fill = un_agencies_fill
                cell.font = _font_with(cell, color="FFFFFF", bold=False, italic=False)

    _auto_fit_columns(ws, min_width=12, max_width=60)


def style_fuel_dashboard_final(ws):
    """Final requested styling for Fuel Dashboard."""
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    width = ws.max_column

    # Top table A1:D4 -> Aduma 16 for Hebrew/numbers, no bold.
    for r in range(1, min(ws.max_row, 4) + 1):
        for c in range(1, min(width, 4) + 1):
            cell = ws.cell(r, c)
            cell.alignment = center
            if _has_hebrew(cell.value) or _is_number_like(cell.value):
                cell.font = _font_with(cell, name="Aduma", size=16, bold=False, italic=False)
            else:
                cell.font = _font_with(cell, size=16, bold=False, italic=False)

    # Bottom tables from row 7 down -> size 12, Aduma for Hebrew/numbers.
    for r in range(7, ws.max_row + 1):
        for c in range(1, width + 1):
            cell = ws.cell(r, c)
            cell.alignment = center
            if _has_hebrew(cell.value) or _is_number_like(cell.value):
                cell.font = _font_with(cell, name="Aduma", size=12, bold=False, italic=False)
            elif cell.value is not None:
                cell.font = _font_with(cell, size=12, bold=False, italic=False)


def style_status_dashboard_final(ws):
    """Final requested styling for סטטוס לפי אישורים."""
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = center
            if _has_hebrew(cell.value) or _is_number_like(cell.value):
                cell.font = _font_with(cell, name="Aduma", size=12, bold=False, italic=False)
            elif cell.value is not None:
                cell.font = _font_with(cell, size=12, bold=False, italic=False)


def move_sheet_before(wb, sheet_name, before_sheet_name):
    if sheet_name not in wb.sheetnames or before_sheet_name not in wb.sheetnames:
        return
    sheet = wb[sheet_name]
    wb._sheets.remove(sheet)
    before_idx = wb.sheetnames.index(before_sheet_name)
    wb._sheets.insert(before_idx, sheet)


def apply_final_workbook_layout(wb):
    if "Distribution Summary" in wb.sheetnames:
        style_distribution_summary_final(wb["Distribution Summary"])
    if "Sector Summary" in wb.sheetnames:
        style_sector_summary_final(wb["Sector Summary"])
    if "Fuel Dashboard" in wb.sheetnames:
        style_fuel_dashboard_final(wb["Fuel Dashboard"])
        move_sheet_before(wb, "Fuel Dashboard", "Sector Summary")
    if "סטטוס לפי אישורים" in wb.sheetnames:
        style_status_dashboard_final(wb["סטטוס לפי אישורים"])


# Hebrew approval status dashboard


def _text_contains(value, *needles) -> bool:
    txt = "" if value is None else str(value).strip().upper()
    return any(str(n).strip().upper() in txt for n in needles if n)


def _normalised_cluster(value) -> str:
    return "" if value is None else str(value).strip().upper()


def _find_required_dashboard_cols(ws, header_row=1):
    cols = {}
    for c in range(1, ws.max_column + 1):
        h = ws.cell(header_row, c).value
        if h is None:
            continue
        h_txt = str(h).strip().upper()
        if h_txt == "CLUSTER":
            cols["Cluster"] = c
        elif h_txt == "AGENCY":
            cols["Agency"] = c
        elif h_txt == "DESCRIPTION":
            cols["Description"] = c
        elif h_txt == STATUS_COL_NAME:
            cols["Status"] = c
        elif h_txt == "UNIFIED FUEL":
            cols["Unified Fuel"] = c
    missing = [k for k in ["Cluster", "Status", "Unified Fuel"] if k not in cols]
    if missing:
        raise RuntimeError(f"Cannot build dashboard. Missing columns: {', '.join(missing)}")
    return cols


def _unified_fuel_column_has_bad_merges(ws, fuel_col, header_row=1):
    """Detect the old bug where Unified Fuel was merged by cluster after inserting סטטוס."""
    bad = []
    for mr in ws.merged_cells.ranges:
        if mr.min_col <= fuel_col <= mr.max_col and mr.max_row > header_row:
            bad.append(str(mr))
    return bad


def _dashboard_category(cluster, agency, status):
    """Map each reviewed row into the requested Hebrew dashboard buckets."""
    c = _normalised_cluster(cluster)
    s = "" if status is None else str(status).strip()
    su = s.upper()

    # Special rule: operational rows under UN Agencies for WFP/OHCHR are approved
    # UN agency fuel, not routine/partner fuel and not unknown entities.
    if c == "UN AGENCIES" and (_is_wfp_un_agency(agency) or "OHCHR" in _compact_alpha_num(agency)):
        return "un_approved"

    # Dedicated sector/status buckets first.
    if "תקשורת" in s:
        return "communications"
    if "בריאות" in s:
        return "health"
    if "סניטציה" in s:
        return "wash_unassigned"
    if "מאפיות" in s:
        return "wfp_bakeries"

    # WFP-specific grouping.
    if c == "WFP":
        if "פעילות" in s:
            return "wfp_routine"
        if "שותפים" in s or "ארגון" in s:
            return "wfp_partners"

    # UN buckets.
    if "סוכנויות" in s and ("לא מפורט" in s or "לא מוכר" in s or "UN SISTERS" in su):
        return "un_unrecognized"
    if "סוכנויות" in s and "מאושרות" in s:
        return "un_approved"

    # Unknown/non-listed organisations.
    if "לא מוכרים" in s or "לא מוכרות" in s or "חברות מקומיות" in s or "NOT FOUND" in su:
        return "unknown_entities"

    # NGO approval review buckets. If a status contains both מסורב and מוקפא
    # (for example TDH), count it as מסורב.
    if "מסורב" in s:
        return "ngo_disapproved"
    if "מוקפא" in s:
        return "ngo_frozen"
    if "מאושר" in s:
        return "ngo_approved"

    return "unknown_entities"


def _write_dashboard_row(ws, row, sector, detail, amount, total, fill_hex, font_hex="FFFFFF", merge_detail=False):
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    fill = PatternFill(fill_type="solid", fgColor=fill_hex)

    if merge_detail:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
        c = ws.cell(row, 1)
        c.value = detail
    else:
        ws.cell(row, 1).value = sector
        ws.cell(row, 2).value = detail
    for col in (1, 2):
        cell = ws.cell(row, col)
        cell.fill = fill
        cell.font = Font(color=font_hex, bold=True, size=12)
        cell.alignment = center
        cell.border = border

    qty_cell = ws.cell(row, 3)
    pct_cell = ws.cell(row, 4)
    qty_cell.value = amount if amount else None
    pct_cell.value = (amount / total) if total else None

    for cell in (qty_cell, pct_cell):
        cell.fill = fill if cell.column == 3 else PatternFill(fill_type="solid", fgColor="D9D9D9")
        cell.font = Font(color=(font_hex if cell.column == 3 else "000000"), bold=True, size=12)
        cell.alignment = center
        cell.border = border
    qty_cell.number_format = "#,##0"
    pct_cell.number_format = "0.0%"


def add_status_dashboard_sheet(wb, source_ws, header_row=1, sheet_name="סטטוס לפי אישורים"):
    """
    Create the requested summary sheet using row-level Unified Fuel values.
    The function intentionally sums Unified Fuel per reviewed row, NOT Total Sum Per Category.
    """
    cols = _find_required_dashboard_cols(source_ws, header_row=header_row)
    fuel_col = cols["Unified Fuel"]
    bad_merges = _unified_fuel_column_has_bad_merges(source_ws, fuel_col, header_row=header_row)
    if bad_merges:
        raise RuntimeError(
            "Unified Fuel is merged in the reviewed file, so row-level dashboard values cannot be trusted. "
            "This version fixes the merge issue for new outputs. Please rerun it on the original Fuels Summary file. "
            f"Problematic ranges: {', '.join(bad_merges[:8])}"
        )

    sums = {
        "ngo_approved": 0.0,
        "ngo_disapproved": 0.0,
        "ngo_frozen": 0.0,
        "wash_unassigned": 0.0,
        "wfp_routine": 0.0,
        "wfp_partners": 0.0,
        "wfp_bakeries": 0.0,
        "un_approved": 0.0,
        "un_unrecognized": 0.0,
        "health": 0.0,
        "unknown_entities": 0.0,
        "communications": 0.0,
        "unclassified": 0.0,
    }

    source_rows = []
    for r in range(header_row + 1, source_ws.max_row + 1):
        cluster = source_ws.cell(r, cols["Cluster"]).value
        agency = source_ws.cell(r, cols.get("Agency")).value if cols.get("Agency") else None
        status = source_ws.cell(r, cols["Status"]).value
        fuel = safe_float(source_ws.cell(r, fuel_col).value)
        if fuel == 0 and not status:
            continue
        category = _dashboard_category(cluster, agency, status)
        sums[category] = sums.get(category, 0.0) + fuel
        source_rows.append({
            "Workbook Row": r,
            "Cluster": cluster,
            "Agency": agency,
            "Status": status,
            "Unified Fuel": fuel,
            "Dashboard Category": category,
        })

    if sheet_name in wb.sheetnames:
        wb.remove(wb[sheet_name])
    ws = wb.create_sheet(sheet_name)
    ws.sheet_view.rightToLeft = False

    # Fixed row order of the Hebrew dashboard.
    headers = ["סקטור", "פירוט", "כמות", "אחוז"]
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(1, c)
        cell.value = h
        cell.fill = PatternFill(fill_type="solid", fgColor="262626")
        cell.font = Font(color="FFFFFF", bold=True, size=13)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(
            left=Side(style="thin", color="000000"),
            right=Side(style="thin", color="000000"),
            top=Side(style="thin", color="000000"),
            bottom=Side(style="thin", color="000000"),
        )

    row_specs = [
        ("NGO", "ארגונים מאושרים", "ngo_approved", "2B8C65", "FFFFFF", False),
        ("NGO", "ארגונים מסורבים", "ngo_disapproved", "FF0000", "FFFFFF", False),
        ("NGO", "ארגונים מוקפאים", "ngo_frozen", "FF0000", "FFFFFF", False),
        (None, "סניטציה ללא שיוך ארגוני", "wash_unassigned", "7563B8", "FFFFFF", True),
        ("WFP", "פעילות שוטפת", "wfp_routine", "4FB3D8", "FFFFFF", False),
        ("WFP", "ארגונים שותפים", "wfp_partners", "4FB3D8", "FFFFFF", False),
        ("WFP", "מאפיות", "wfp_bakeries", "4FB3D8", "FFFFFF", False),
        ('או"ם', 'סוכנויות או"ם מאושרות', "un_approved", "FFC000", "FFFFFF", False),
        ('או"ם', 'סוכנויות או"ם לא מוכרות (UN Sisters)', "un_unrecognized", "FF0000", "FFFFFF", False),
        (None, "בריאות", "health", "6A1BDA", "FFFFFF", True),
        (None, UNKNOWN_ENTITIES_STATUS, "unknown_entities", "C044E8", "FFFFFF", True),
        (None, "תקשורת", "communications", "21A789", "FFFFFF", True),
    ]
    if sums.get("unclassified", 0):
        row_specs.append((None, "לא מסווג", "unclassified", "808080", "FFFFFF", True))

    total = sum(sums.get(key, 0.0) for _, _, key, *_ in row_specs)

    start_row = 2
    for i, spec in enumerate(row_specs, start=start_row):
        sector, detail, key, fill_hex, font_hex, merge_detail = spec
        _write_dashboard_row(ws, i, sector, detail, sums.get(key, 0.0), total, fill_hex, font_hex, merge_detail)

    # Merge group labels vertically (NGO/WFP/UN) after writing row borders/fills.
    group_ranges = [(2, 4, "NGO", "2B8C65"), (6, 8, "WFP", "4FB3D8"), (9, 10, 'או"ם', "FFC000")]
    for r1, r2, label, fill_hex in group_ranges:
        ws.merge_cells(start_row=r1, start_column=1, end_row=r2, end_column=1)
        cell = ws.cell(r1, 1)
        cell.value = label
        cell.fill = PatternFill(fill_type="solid", fgColor=fill_hex)
        cell.font = Font(color="FFFFFF", bold=True, size=16)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        for r in range(r1, r2 + 1):
            ws.cell(r, 1).border = Border(
                left=Side(style="thin", color="000000"),
                right=Side(style="thin", color="000000"),
                top=Side(style="thin", color="000000"),
                bottom=Side(style="thin", color="000000"),
            )

    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 42
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 12
    for r in range(1, ws.max_row + 1):
        ws.row_dimensions[r].height = 28

    return sums, pd.DataFrame(source_rows)


# Weekly Fuel Overview


FUEL_OVERVIEW_SHEET_NAME = "Fuel Overview"

# Requested statuses, together with spelling/plural variants
# that may appear in the final workbook.
FUEL_OVERVIEW_INCLUDED_STATUSES = {
    "ארגונים/גופים לא מוכרים (INGOs/חברות מקומיות)",
    "ארגון מוקפא",
    "ארגון מסורב",
    'סוכנויות או"ם לא מפורטות (UN Sisters)',
    "ארגונים שותפים",

    # Possible variants in the approval workbook:
    "ארגונים מוקפאים",
    "ארגונים מסורבים",
    'סוכנויות או"ם לא מוכרות (UN Sisters)',
}


def _overview_normalise_text(value) -> str:
    if value is None:
        return ""

    text = str(value).strip()

    # Normalise different quotation-mark styles.
    text = (
        text.replace("״", '"')
        .replace("“", '"')
        .replace("”", '"')
    )

    text = re.sub(r"\s+", " ", text)
    return text


def _overview_status_is_included(value) -> bool:
    status = _overview_normalise_text(value)
    return status in FUEL_OVERVIEW_INCLUDED_STATUSES or "מסורב" in status or "מוקפא" in status


def _overview_rounded_percentage(numerator, denominator):
    return round(numerator / denominator * 100, 1) / 100 if denominator else 0.0


def _overview_litre_text(value):
    """
    Create real text containing both the number and ליטר.

    Because this is text rather than a custom number format,
    copying the cell will also copy the word ליטר.
    """
    return f'\u200fליטר {value:,.0f}'


def _overview_find_header_column(ws, possible_names, header_row=1):
    """
    Find a column by its header instead of relying on a fixed letter.
    """
    wanted = {
        _overview_normalise_text(name).upper()
        for name in possible_names
    }

    for c in range(1, ws.max_column + 1):
        current_header = _overview_normalise_text(
            ws.cell(row=header_row, column=c).value
        ).upper()

        if current_header in wanted:
            return c

    raise RuntimeError(
        f'Could not find any of these headers in "{ws.title}": '
        f'{possible_names}'
    )


def _overview_get_used_fuel(ws_dashboard, organisation):
    used_col = _overview_find_header_column(ws_dashboard, ["דלק שנוצל"])
    top_rows = ws_dashboard.iter_rows(min_row=2, max_row=min(ws_dashboard.max_row, 10),
                                       max_col=min(ws_dashboard.max_column, 10))
    for row in top_rows:
        if any(_overview_normalise_text(cell.value).upper() == organisation.upper() for cell in row):
            return safe_float(ws_dashboard.cell(row[0].row, used_col).value)
    raise RuntimeError(f'Could not find "{organisation}" in the Fuel Dashboard summary table.')


def _calculate_fuel_overview_metrics(wb):
    required = [DISTRIBUTION_SHEET_NAME, "Fuel Dashboard"]
    missing = [name for name in required if name not in wb.sheetnames]
    if missing:
        raise RuntimeError("Cannot create Fuel Overview. Missing sheets: " + ", ".join(missing))

    summary, dashboard = (wb[name] for name in required)
    cluster_col = _overview_find_header_column(summary, ["Cluster", "INTERVENTION"])
    fuel_col = _overview_find_header_column(summary, ["Unified Fuel"])
    status_col = _overview_find_header_column(summary, [STATUS_COL_NAME])
    filtered = {"UNOPS": 0.0, "WFP": 0.0}
    for row in summary.iter_rows(min_row=2, max_col=max(cluster_col, fuel_col, status_col), values_only=True):
        if _overview_status_is_included(row[status_col - 1]):
            cluster = _overview_normalise_text(row[cluster_col - 1]).upper()
            org = "WFP" if cluster == "WFP" else "UNOPS"
            filtered[org] += safe_float(row[fuel_col - 1])

    used = {org: _overview_get_used_fuel(dashboard, org) for org in filtered}
    metrics = {}
    for org in filtered:
        key = org.lower()
        metrics[f"{key}_filtered_fuel"] = filtered[org]
        metrics[f"{key}_used_fuel"] = used[org]
        metrics[f"{key}_percentage"] = _overview_rounded_percentage(filtered[org], used[org])
    metrics["total_used_fuel"] = sum(used.values())
    metrics["total_filtered_fuel"] = sum(filtered.values())
    metrics["total_percentage"] = _overview_rounded_percentage(metrics["total_filtered_fuel"], metrics["total_used_fuel"])
    return metrics


def _overview_fill_background(ws, cell_range, fill_colour):
    fill = PatternFill("solid", fgColor=fill_colour)
    left, top, right, bottom = range_boundaries(cell_range)
    for row in ws.iter_rows(min_row=top, max_row=bottom, min_col=left, max_col=right):
        for cell in row:
            cell.fill = fill


def _overview_merge_text(ws, cell_range, value, font_size, bold=True, font_name="Arial", font_colour="000000"):
    left, top, _, _ = range_boundaries(cell_range)
    ws.merge_cells(cell_range)
    cell = ws.cell(top, left)
    cell.value = value
    cell.font = Font(name=font_name, size=font_size, bold=bold, color=font_colour)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    return cell


def _overview_create_box(ws, cell_range, value, fill_colour, font_size, bold=True,
                         number_format=None, border_style="medium", font_name="Arial",
                         font_colour="000000", reading_order=0):
    """Draw a bordered, merged value box. Reading order 2 is right-to-left."""
    left, top, right, bottom = range_boundaries(cell_range)
    fill = PatternFill("solid", fgColor=fill_colour)
    font = Font(name=font_name, size=font_size, bold=bold, color=font_colour)
    alignment = Alignment(horizontal="center", vertical="center", wrap_text=True, readingOrder=reading_order)
    edge, inside = Side(style=border_style, color="000000"), Side(style=None)
    for row in ws.iter_rows(min_row=top, max_row=bottom, min_col=left, max_col=right):
        for cell in row:
            cell.fill, cell.font, cell.alignment = fill, font, alignment
            cell.border = Border(left=edge if cell.column == left else inside,
                                 right=edge if cell.column == right else inside,
                                 top=edge if cell.row == top else inside,
                                 bottom=edge if cell.row == bottom else inside)
    ws.merge_cells(cell_range)
    anchor = ws.cell(top, left)
    anchor.value = value
    anchor.font, anchor.alignment = font, alignment
    if number_format:
        anchor.number_format = number_format
    return anchor


def add_fuel_overview_sheet(
    wb,
    sheet_name=FUEL_OVERVIEW_SHEET_NAME,
):
    """
    Create the visual weekly fuel overview as another Excel sheet.
    """
    metrics = _calculate_fuel_overview_metrics(wb)

    if sheet_name in wb.sheetnames:
        wb.remove(wb[sheet_name])

    ws = wb.create_sheet(sheet_name)

    ws.sheet_view.showGridLines = False
    ws.sheet_view.rightToLeft = False
    ws.sheet_view.zoomScale = 85

    # Page-printing settings.
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.print_area = "A1:N26"

    ws.page_margins.left = 0.2
    ws.page_margins.right = 0.2
    ws.page_margins.top = 0.25
    ws.page_margins.bottom = 0.25

    # Fuel Overview colours.
    background_colour = "D9D9D9"
    red_colour = "F43D3D"
    yellow_colour = "FFB800"
    separator_colour = "F3C4AA"
    percentage_colour = "E7E7E7"
    summary_colour = "E7E7E7"

    # Overall background.
    _overview_fill_background(
        ws,
        "A1:N27",
        background_colour,
    )

    # Column widths.
    column_widths = {
        "A": 5,
        "B": 9,
        "C": 12,
        "D": 12,
        "E": 12,
        "F": 9,
        "G": 4,
        "H": 4,
        "I": 9,
        "J": 9,
        "K": 12,
        "L": 12,
        "M": 12,
        "N": 9,
    }

    for column_letter, width in column_widths.items():
        ws.column_dimensions[column_letter].width = width

    for first, last, height in [
        (1, 1, 8), (2, 3, 30), (4, 4, 8), (5, 7, 23), (8, 8, 8), (9, 11, 23),
        (12, 12, 8), (13, 15, 27), (16, 17, 8), (18, 20, 23), (21, 22, 10),
        (23, 26, 24), (27, 27, 20),
    ]:
        for r in range(first, last + 1):
            ws.row_dimensions[r].height = height

    for org, title_range, left, right in (("wfp", "B2:F3", "C", "E"), ("unops", "J2:N3", "K", "M")):
        _overview_merge_text(ws, title_range, org.upper(), font_size=30)
        _overview_create_box(ws, f"{left}5:{right}7", _overview_litre_text(metrics[f"{org}_filtered_fuel"]),
                             red_colour, 20, font_name="Aduma", reading_order=2)
        _overview_create_box(ws, f"{left}8:{right}8", "", separator_colour, 1, border_style="thin")
        _overview_create_box(ws, f"{left}9:{right}11", _overview_litre_text(metrics[f"{org}_used_fuel"]),
                             yellow_colour, 20, font_name="Aduma", reading_order=2)
        _overview_create_box(ws, f"{left}13:{right}15", metrics[f"{org}_percentage"],
                             percentage_colour, 36, number_format="0.0%", border_style="thin", font_name="Aduma")

    # Explanation boxes
    # ========================================================

    _overview_create_box(
        ws,
        "A18:G20",
        "כמות הדלק שחולקה לארגונים "
        "לא מוכרים/מסורבים השבוע",
        red_colour,
        font_size=14,
    )

    _overview_create_box(
        ws,
        "H18:N20",
        "סך כמות הדלק שחולק השבוע",
        yellow_colour,
        font_size=14,
    )

    # ========================================================
    # Bottom summary
    # ========================================================

    summary_text = (
        f"בשבוע החולף חולקו "
        f"{metrics['total_used_fuel']:,.0f} ליטר דלק, "
        f"מתוכם {metrics['total_percentage']:.1%} "
        f"חולקו לארגונים מסורבים ולא מוכרים"
    )

    _overview_create_box(
        ws,
        "A23:N26",
        "\u200f" + summary_text,
        summary_colour,
        font_size=20,
        border_style="thin",
        font_name="Aduma",
        reading_order=2,
    )

    # Hidden calculation table for checking the source values.
    hidden_values = [
        ["Metric", "Value"],
        ["WFP filtered fuel", metrics["wfp_filtered_fuel"]],
        ["WFP used fuel", metrics["wfp_used_fuel"]],
        ["WFP percentage", metrics["wfp_percentage"]],
        ["UNOPS filtered fuel", metrics["unops_filtered_fuel"]],
        ["UNOPS used fuel", metrics["unops_used_fuel"]],
        ["UNOPS percentage", metrics["unops_percentage"]],
        ["Total used fuel", metrics["total_used_fuel"]],
        ["Total filtered fuel", metrics["total_filtered_fuel"]],
        ["Total percentage", metrics["total_percentage"]],
    ]

    for row_number, row_values in enumerate(
        hidden_values,
        start=1,
    ):
        for column_number, value in enumerate(
            row_values,
            start=16,
        ):
            ws.cell(
                row=row_number,
                column=column_number,
            ).value = value

    ws.column_dimensions["P"].hidden = True
    ws.column_dimensions["Q"].hidden = True

    return ws


# Report assembly and compatibility entry points


def apply_approval_status(wb, approval_index, prefixes, include_match_details=False, source_name="workbook"):
    """Add row-level approval statuses and the final dashboards to the live workbook."""
    if DISTRIBUTION_SHEET_NAME not in wb.sheetnames:
        raise RuntimeError(f'Could not find sheet "{DISTRIBUTION_SHEET_NAME}" in {source_name}.')

    ws = wb[DISTRIBUTION_SHEET_NAME]
    headers = {
        "Cluster": ["Cluster"],
        "Agency": ["AGENCY", "Agency"],
        "Description": ["DESCRIPTION", "Description"],
    }
    header_row, cols = find_header_row_and_columns(ws, headers, max_scan_rows=5)
    if not header_row:
        raise RuntimeError('Could not find Cluster / AGENCY / DESCRIPTION headers in Distribution Summary.')

    status_col, _ = get_or_insert_status_column(ws, header_row, cols["Description"])

    detail_headers = ["Match Type", "Match Score", "Matched Cluster/Intervention", "Matched Agency", "Matched Description", "Approval Sheet Row"]
    detail_cols = {}
    if include_match_details:
        start_col = ws.max_column + 1
        for i, h in enumerate(detail_headers):
            c = start_col + i
            detail_cols[h] = c
            cell = ws.cell(header_row, c)
            cell.value = h
            copy_cell_style(ws.cell(header_row, status_col), cell)
            ws.column_dimensions[get_column_letter(c)].width = 24

    review_rows = []
    prefixes = frozenset(prefixes or BASE_GENERATED_PREFIXES)
    centre = Alignment(horizontal="center", vertical="center", wrap_text=True)
    status_font = Font(bold=True, color="000000")
    status_fills = {}
    for r in range(header_row + 1, ws.max_row + 1):
        cluster = ws.cell(r, cols["Cluster"]).value
        agency = ws.cell(r, cols["Agency"]).value
        desc = ws.cell(r, cols["Description"]).value
        if not any([cluster, agency, desc]):
            continue

        review = review_summary_row(cluster, agency, desc, approval_index, prefixes)
        status_cell = ws.cell(r, status_col)
        status_cell.value = review["Review Status"]

        # Use the row's existing style, then color only the status cell for easy review.
        copy_cell_style(ws.cell(r, cols["Description"]), status_cell)
        label = review["Review Status"]
        if label not in status_fills:
            status_fills[label] = status_fill(label)
        status_cell.fill = status_fills[label]
        status_cell.font = status_font
        status_cell.alignment = centre

        if include_match_details:
            for h in detail_headers:
                c = detail_cols[h]
                ws.cell(r, c).value = review[h]
                copy_cell_style(status_cell, ws.cell(r, c))
                ws.cell(r, c).fill = PatternFill(fill_type=None)
                ws.cell(r, c).font = Font(bold=False, color="000000")

        review_rows.append({
            "Workbook Row": r,
            "Detected Org": review["Detected Org"],
            "Cluster": cluster,
            "Agency": agency,
            "Description": desc,
            "Review Status": review["Review Status"],
            "Match Type": review["Match Type"],
            "Match Score": review["Match Score"],
            "Matched Cluster/Intervention": review["Matched Cluster/Intervention"],
            "Matched Agency": review["Matched Agency"],
            "Matched Description": review["Matched Description"],
            "Approval Sheet Row": review["Approval Sheet Row"],
        })

    # Header styling for the new status column.
    header_cell = ws.cell(header_row, status_col)
    header_cell.value = STATUS_COL_NAME
    header_cell.fill = PatternFill(fill_type="solid", fgColor="000000")
    header_cell.font = Font(color="FFFFFF", bold=True)
    header_cell.alignment = Alignment(horizontal="center", vertical="center")

    # The filter should cover only columns with real headers.
    ws.freeze_panes = "A2"
    _set_autofilter_to_real_headers(ws, header_row=header_row)

    # Build the requested Hebrew dashboard sheet from row-level Unified Fuel values
    # BEFORE deleting helper columns.
    add_status_dashboard_sheet(wb, ws, header_row=header_row)

    # Remove internal helper columns from the final Distribution Summary.
    _delete_columns_by_headers(
        ws,
        header_row,
        ["Fuel sum", "Description Sum"],
    )
    _set_autofilter_to_real_headers(
        ws,
        header_row=header_row,
    )

    # Remove legacy QA sheet if present.
    if "Status Review Log" in wb.sheetnames:
        wb.remove(wb["Status Review Log"])

    # Create the visual Fuel Overview sheet.
    # It requires Fuel Dashboard because the used-fuel values
    # are taken from that sheet.
    if "Fuel Dashboard" in wb.sheetnames:
        add_fuel_overview_sheet(
            wb,
            sheet_name=FUEL_OVERVIEW_SHEET_NAME,
        )

    # Apply the existing final workbook formatting.
    apply_final_workbook_layout(wb)

    if FUEL_OVERVIEW_SHEET_NAME in wb.sheetnames:
        overview = wb[FUEL_OVERVIEW_SHEET_NAME]
        wb.move_sheet(overview, offset=len(wb.sheetnames) - 1 - wb.index(overview))

    return pd.DataFrame(review_rows)


def add_status_to_fuels_summary_workbook(fuels_summary_file, approval_index, prefixes, include_match_details=False):
    with closing(load_workbook(BytesIO(fuels_summary_file.getvalue()))) as workbook:
        review = apply_approval_status(workbook, approval_index, prefixes, include_match_details,
                                       source_name=fuels_summary_file.name)
        return workbook_bytes(workbook), review


def generate_report(uploads, latest_day_uploads, selected_dates, approval_source,
                    progress=None, status=None, include_match_details=False):
    """Run the complete report and return (Excel buffer, row-level review, dashboard log)."""
    if approval_source is None:
        raise RuntimeError(f"Approval/status workbook was not found next to the app file: {APPROVAL_FILE_NAME}")
    if latest_day_uploads and not selected_dates:
        raise RuntimeError("Please select valid week start and end dates for the Fuel Dashboard.")

    def step(percent, message):
        if progress is not None:
            progress.progress(percent)
        if status is not None:
            status.info(message)

    logs = []
    step(5, "Step 1/4: Combining Total Distribution files…")
    with closing(build_combined_workbook(uploads, status)) as workbook:
        step(25, "Step 2/4: Running Fuels Cleaner calculations…")
        _prepare_excel_values(workbook[DISTRIBUTION_SHEET_NAME])
        calculate_distribution(workbook, status=status)
        if latest_day_uploads:
            step(80, "Step 3/4: Adding Fuel Dashboard sheet…")
            add_fuel_dashboard_sheet(workbook, latest_day_uploads, selected_dates, logs)
        elif status is not None:
            status.warning("No latest-day UNOPS/WFP files uploaded, so Fuel Dashboard was not added.")

        step(90, "Step 4/4: Adding approval status review and סטטוס לפי אישורים sheet…")
        for sheet in workbook:
            _prepare_excel_values(sheet)
        index, prefixes, _ = build_approval_index(approval_source)
        review = apply_approval_status(workbook, index, prefixes, include_match_details)
        step(98, "Writing the downloadable Excel file…")
        output = workbook_bytes(workbook)
    if progress is not None:
        progress.progress(100)
    return output, review, logs


class LocalWorkbookFile:
    def __init__(self, path):
        self.path = path
        self.name = os.path.basename(path)
        self.size = os.path.getsize(path) if os.path.exists(path) else 0

    def getvalue(self):
        with open(self.path, "rb") as f:
            return f.read()


class BytesWorkbookFile:
    def __init__(self, name, data):
        self.name = name
        self._data = data
        self.size = len(data)

    def getvalue(self):
        return self._data


def get_saved_approval_file():
    if os.path.exists(APPROVAL_STORE_PATH):
        return LocalWorkbookFile(APPROVAL_STORE_PATH)
    return None


def save_uploaded_approval_file(uploaded_file):
    with open(APPROVAL_STORE_PATH, "wb") as f:
        f.write(uploaded_file.getvalue())
    return LocalWorkbookFile(APPROVAL_STORE_PATH)


# Streamlit interface


def _reset_cached_output(state=None):
    if state is None:
        import streamlit as st
        state = st.session_state
    for key in ("final_bytes", "final_name", "debug_logs", "status_review_summary", "last_upload_key"):
        state.pop(key, None)


def _upload_identity(upload):
    # Streamlit identifies the actual upload, even if a replacement has the same name and size.
    identifier = getattr(upload, "file_id", None)
    if identifier is None:
        identifier = hashlib.sha256(upload.getvalue()).hexdigest()
    return upload.name, upload.size, identifier


def _show_download(st):
    state = st.session_state
    if state.get("status_review_summary"):
        with st.expander("Status review summary"):
            st.dataframe(pd.DataFrame(state.status_review_summary), use_container_width=True)
    if state.get("debug_logs"):
        with st.expander("Fuel Dashboard debug logs"):
            st.text("\n".join(state.debug_logs))
    st.download_button("Download final Excel", data=state.final_bytes, file_name=state.final_name,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       use_container_width=True)


def main():
    import streamlit as st

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

    left, right = st.columns([2, 1])

    with left:
        st.markdown('<div class="card">', unsafe_allow_html=True)

        saved_approval = get_saved_approval_file()
        if saved_approval is not None:
            st.caption(f"Using local approval workbook: {saved_approval.name}")
        else:
            st.warning(f"Approval workbook was not found next to the app file: {APPROVAL_FILE_NAME}")

        uploads = st.file_uploader(
            "Upload Total Distribution .xlsx files",
            type=["xlsx"],
            accept_multiple_files=True,
            key="distribution_uploads",
        )
        latest_day_uploads = st.file_uploader(
            "Upload latest-day UNOPS + WFP files for Fuel Dashboard",
            type=["xlsx"],
            accept_multiple_files=True,
            key="latest_day_uploads",
        )
        st.markdown(
            '<div class="small">The app reads the approval workbook directly from the code folder. The final workbook will include Distribution Summary, Sector Summary, סטטוס לפי אישורים, and the optional Fuel Dashboard and Fuel Overview.</div>',
            unsafe_allow_html=True,
        )
        st.markdown("</div>", unsafe_allow_html=True)

    with right:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown("**Pipeline**")
        st.markdown(
            """
1) Combine Total Distribution files  
2) Run Fuels Cleaner calculations  
3) Create Distribution Summary  
4) Create Sector Summary  
5) Add Fuel Dashboard sheet from latest-day files  
6) Add סטטוס review from local approval list  
7) Add סטטוס לפי אישורים summary table  
8) Download final workbook
"""
        )
        st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("### Fuel Dashboard week dates")
    col1, col2 = st.columns(2)
    with col1:
        start_date = st.date_input("Week start date")
    with col2:
        end_date = st.date_input("Week end date")


    selected_dates = []
    if start_date and end_date:
        selected_dates = [start_date + timedelta(days=n) for n in range((end_date - start_date).days + 1)]

    run_btn = st.button("Run combine + calculations", type="primary", disabled=not uploads)
    approval_version = None
    if saved_approval is not None:
        stat = os.stat(saved_approval.path)
        approval_version = (stat.st_size, stat.st_mtime_ns)
    upload_key = (tuple(_upload_identity(u) for u in uploads or []),
                  tuple(_upload_identity(u) for u in latest_day_uploads or []),
                  start_date, end_date, approval_version)
    if upload_key != st.session_state.get("last_upload_key"):
        _reset_cached_output(st.session_state)
        st.session_state.last_upload_key = upload_key

    if st.session_state.get("final_bytes"):
        st.success("Done! Download your final file below.")
        _show_download(st)
        return
    if not run_btn:
        return

    progress = st.progress(0)
    status = st.empty()
    loader = st.empty()
    loader.info("Processing files…")
    try:
        output, review, logs = generate_report(uploads, latest_day_uploads, selected_dates,
                                               saved_approval, progress, status)
        if review.empty:
            counts = []
        else:
            columns = ["Detected Org", "Review Status", "Match Type"]
            counts = review.groupby(columns, dropna=False).size().reset_index(name="Count").to_dict("records")
        st.session_state.status_review_summary = counts
        st.session_state.debug_logs = logs
        st.session_state.final_name = "Fuels summary.xlsx"
        st.session_state.final_bytes = output.getvalue()
        status.success("Done! Download your final file below.")
        _show_download(st)
    except Exception as exc:
        status.empty()
        st.error("Processing failed. The diagnostic below shows the exact file and workbook area that caused the problem.")
        st.code(str(exc), language=None)
        with st.expander("Technical traceback"):
            st.exception(exc)
    finally:
        loader.empty()


if __name__ == "__main__":
    main()
