import streamlit as st
from io import BytesIO
from copy import copy
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill
from openpyxl.chart import PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.styles import PatternFill, Font
import base64

st.set_page_config(page_title="Excel Cleaner", layout="centered")

# ----------------------------
# UI (redesign + helpers)
# ----------------------------
st.markdown("""
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
""", unsafe_allow_html=True)

st.markdown('<div class="title">Fuels Data Cleaner</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Upload an Excel file — get back a cleaned, sorted, deduplicated, colour-coded sheet.</div>', unsafe_allow_html=True)

TARGET_SHEET = "UNOPS Total Distribution"

left, right = st.columns([2, 1])

with left:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    uploaded = st.file_uploader("Excel file (.xlsx)", type=["xlsx"], label_visibility="collapsed")
    st.markdown('<div class="small">Keeps only “UNOPS Total Distribution”, cleans & enriches it, then returns a new file.</div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

with right:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown("**What it does**")
    st.markdown("""
- Keeps target sheet only  
- Unmerges A–C, fills values  
- Removes row 2 + TOTAL rows  
- Fuel Sum, Description Sum, Unified Fuel  
- Sorts + removes duplicates  
- Adds Total Sum Per Category  
- Colours rows (A–G) by INTERVENTION
""")
    st.markdown('</div>', unsafe_allow_html=True)

def copy_cell_style(src, dst):
    dst.font = copy(src.font)
    dst.fill = copy(src.fill)
    dst.border = copy(src.border)
    dst.alignment = copy(src.alignment)
    dst.number_format = src.number_format
    dst.protection = copy(src.protection)

# ----------------------------
# Session-state cache (prevents re-processing on download click)
# ----------------------------
def _reset_cached_output():
    st.session_state.pop("cleaned_bytes", None)
    st.session_state.pop("cleaned_name", None)

if "last_upload_key" not in st.session_state:
    st.session_state.last_upload_key = None

# ----------------------------
# Excel helpers
# ----------------------------
def unmerge_and_fill(ws, col_min: int, col_max: int):
    """Unmerge any merged range intersecting columns [col_min..col_max], then fill value+style across the range."""
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
    """Store value + formatting for full row."""
    data = []
    for c in range(1, ws.max_column + 1):
        cell = ws.cell(r, c)
        data.append({
            "value": cell.value,
            "font": copy(cell.font),
            "fill": copy(cell.fill),
            "border": copy(cell.border),
            "alignment": copy(cell.alignment),
            "number_format": cell.number_format,
            "protection": copy(cell.protection),
        })
    return data

def restore_row(ws, r, row_data):
    """Write row back with formatting."""
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
    """Convert common Excel-ish values into float, else 0."""
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
    """Normalize header strings (handles extra spaces/newlines)."""
    return "" if x is None else " ".join(str(x).split()).strip().upper()

def rgb_fill(r, g, b):
    """openpyxl uses ARGB hex, so prefix FF for fully-opaque."""
    return PatternFill(fill_type="solid", fgColor=f"FF{r:02X}{g:02X}{b:02X}")

def show_small_loader_video(placeholder, path: str, width_px: int = 220):
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
        unsafe_allow_html=True
    )

# ----------------------------
# Processing
# ----------------------------
if uploaded:
    # If the upload changed, clear cached output so we re-process once
    upload_key = f"{uploaded.name}:{uploaded.size}"
    if st.session_state.last_upload_key != upload_key:
        st.session_state.last_upload_key = upload_key
        _reset_cached_output()

    # If we've already processed this upload, don't re-run heavy code
    if "cleaned_bytes" in st.session_state:
        st.success("Done! Download your cleaned file below.")
        st.download_button(
            "Download cleaned Excel",
            data=st.session_state.cleaned_bytes,
            file_name=st.session_state.get("cleaned_name", "cleaned.xlsx"),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.stop()

    progress = st.progress(0)
    status = st.empty()

    loader = st.empty()
    show_small_loader_video(loader, "dog_running.mp4", width_px=220)  # <-- change 220

    try:
        status.info("Loading workbook…")
        wb = load_workbook(uploaded, data_only=False)
        uploaded.seek(0)
        wb_cache = load_workbook(uploaded, data_only=True)
        progress.progress(10)

        sheet_map = {name.strip(): name for name in wb.sheetnames}
        sheet_map_cache = {name.strip(): name for name in wb_cache.sheetnames}

        if TARGET_SHEET not in sheet_map or TARGET_SHEET not in sheet_map_cache:
            status.error(
                f'Sheet "{TARGET_SHEET}" not found.\n\n'
                f"Sheets found: {list(sheet_map.keys())}"
            )
        else:
            keep_name = sheet_map[TARGET_SHEET]
            keep_name_cache = sheet_map_cache[TARGET_SHEET]

            # Keep only the target sheet (both workbooks)
            for name in list(wb.sheetnames):
                if name != keep_name:
                    wb.remove(wb[name])
            for name in list(wb_cache.sheetnames):
                if name != keep_name_cache:
                    wb_cache.remove(wb_cache[name])

            ws = wb[keep_name]
            ws_cache = wb_cache[keep_name_cache]
            ws.title = TARGET_SHEET
            progress.progress(15)

            # ---------- UNMERGE A–C + COPY VALUES & STYLES ----------
            status.info("Unmerging A–C…")
            unmerge_and_fill(ws, col_min=1, col_max=3)  # A..C
            progress.progress(22)

            # ---------- CONVERT ALL EXISTING FORMULAS TO VALUES (cached) ----------
            status.info("Freezing formulas into values…")
            max_row = ws.max_row
            max_col = ws.max_column
            for r in range(1, max_row + 1):
                for c in range(1, max_col + 1):
                    cell = ws.cell(row=r, column=c)
                    is_formula = (cell.data_type == "f") or (
                        isinstance(cell.value, str) and cell.value.startswith("=")
                    )
                    if is_formula:
                        cell.value = ws_cache.cell(row=r, column=c).value
            progress.progress(30)

            # ---------- DELETE ROW 2 ----------
            ws.delete_rows(2)

            # ---------- DELETE ROWS CONTAINING 'TOTAL' IN A,B,C ----------
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
            progress.progress(38)

            header_row = 1

            # ---------- Unmerge D–F before Fuel Sum ----------
            unmerge_and_fill(ws, col_min=4, col_max=6)  # D..F

            # ---------- ADD "Fuel sum" IN F AS VALUES ----------
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
            progress.progress(48)

            # ---------- DELETE ROWS WHERE FUEL SUM IS 0 OR EMPTY ----------
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
            progress.progress(55)

            # ---------- DELETE COLUMNS D AND E ----------
            ws.delete_cols(5)  # E
            ws.delete_cols(4)  # D
            # Now Fuel sum moved from F -> D.

            # ---------- INSERT "Description Sum" AS COLUMN E ----------
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
            progress.progress(65)

            # ---------- UNIFIED FUEL AS VALUES ----------
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
                target = ws.cell(row=r, column=unified_col)   # F
                style_src = ws.cell(row=r, column=4)          # D
                key = ws.cell(row=r, column=5).value          # E

                target.value = totals.get(key, 0.0)
                target.font = copy(style_src.font)
                target.fill = copy(style_src.fill)
                target.border = copy(style_src.border)
                target.alignment = copy(style_src.alignment)
                target.number_format = style_src.number_format
                target.protection = copy(style_src.protection)
            progress.progress(75)

            # ---------- SORTING ----------
            status.info("Sorting…")
            intervention_col = None
            for c in range(1, ws.max_column + 1):
                h = ws.cell(row=1, column=c).value
                if norm_header(h) == "INTERVENTION":
                    intervention_col = c
                    break
            if intervention_col is None:
                intervention_col = 1

            # --- Normalise: fold UN-OHCHR into INGOs (so all calculations treat it as INGOs) ---
            for r in range(2, ws.max_row + 1):
                v = ws.cell(row=r, column=intervention_col).value
                if v is None:
                    continue
                t = str(v).strip()
                if "UN-OHCHR" in t.upper():
                    ws.cell(row=r, column=intervention_col).value = "INGOs"

            rows_data = []
            for r in range(2, ws.max_row + 1):
                fuel_val = safe_float(ws.cell(row=r, column=unified_col).value)
                intervention_val = str(ws.cell(row=r, column=intervention_col).value or "")
                rows_data.append({
                    "fuel": fuel_val,
                    "intervention": intervention_val,
                    "row": snapshot_row(ws, r)
                })

            rows_data.sort(key=lambda x: -x["fuel"])
            rows_data.sort(key=lambda x: x["intervention"].strip().lower())

            write_row = 2
            for obj in rows_data:
                restore_row(ws, write_row, obj["row"])
                write_row += 1
            progress.progress(82)

            # ---------- REMOVE DUPLICATES BY Description Sum ----------
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
            progress.progress(88)

            # ---------- Total Sum Per Category in G ----------
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

            ref_letter = get_column_letter(unified_col)     # F
            new_letter = get_column_letter(total_cat_col)   # G
            if ref_letter in ws.column_dimensions and ws.column_dimensions[ref_letter].width is not None:
                ws.column_dimensions[new_letter].width = ws.column_dimensions[ref_letter].width

            cat_totals = {}
            for r in range(2, ws.max_row + 1):
                cat = ws.cell(row=r, column=1).value
                cat_key = "" if cat is None else str(cat).strip()
                fuel_val = safe_float(ws.cell(row=r, column=unified_col).value)  # F
                cat_totals[cat_key] = cat_totals.get(cat_key, 0.0) + fuel_val

            for r in range(2, ws.max_row + 1):
                cat = ws.cell(row=r, column=1).value
                cat_key = "" if cat is None else str(cat).strip()

                target = ws.cell(row=r, column=total_cat_col)      # G
                style_src = ws.cell(row=r, column=unified_col)     # F

                target.value = cat_totals.get(cat_key, 0.0)
                target.font = copy(style_src.font)
                target.fill = copy(style_src.fill)
                target.border = copy(style_src.border)
                target.alignment = copy(style_src.alignment)
                target.number_format = style_src.number_format
                target.protection = copy(style_src.protection)
            progress.progress(94)
            
            # =====================================================
            # NEW: MERGE "Total Sum Per Category" (G) BY INTERVENTION
            # Merges G for each consecutive INTERVENTION group
            # (sorting already grouped INTERVENTION together)
            # =====================================================
            status.info('Merging "Total Sum Per Category" by INTERVENTION…')

            INTERVENTION_COL = intervention_col
            TOTAL_CAT_COL = total_cat_col  # should be 7 (G)

            def _norm_intervention(v):
                return "" if v is None else str(v).strip().upper()

            start = 2
            while start <= ws.max_row:
                key = _norm_intervention(ws.cell(row=start, column=INTERVENTION_COL).value)

                end = start
                while end + 1 <= ws.max_row and _norm_intervention(ws.cell(row=end + 1, column=INTERVENTION_COL).value) == key:
                    end += 1

                # Merge only if group has 2+ rows
                if end > start:
                    # Take style + value from the first row in the group
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

                    # Merge range G(start):G(end)
                    ws.merge_cells(start_row=start, start_column=TOTAL_CAT_COL,
                                   end_row=end, end_column=TOTAL_CAT_COL)

                    # Re-apply value + style to the merged (top-left) cell
                    merged_top = ws.cell(row=start, column=TOTAL_CAT_COL)
                    merged_top.value = top_val
                    merged_top.font = copy(top_style["font"])
                    merged_top.fill = copy(top_style["fill"])
                    merged_top.border = copy(top_style["border"])
                    merged_top.alignment = copy(top_style["alignment"])
                    merged_top.number_format = top_style["number_format"]
                    merged_top.protection = copy(top_style["protection"])

                start = end + 1

            # ============================
            # 1) OPTIONAL: rename LOGISTICS -> UN Agencies
            # Put this BEFORE the colouring loop (right before: status.info("Applying colours…"))
            # ============================
            for r in range(2, ws.max_row + 1):
                v = ws.cell(row=r, column=intervention_col).value
                if v is None:
                    continue
                if str(v).strip().upper() == "LOGISTICS":
                    ws.cell(row=r, column=intervention_col).value = "UN Agencies"

            # ---------- COLOR A–G BASED ON INTERVENTION ----------
            status.info("Applying colours…")
            fills = {
                "TELECOMMUNICATIONS": rgb_fill(213, 243, 251),
                "HEALTH": rgb_fill(0, 176, 80),
                "WASH": rgb_fill(250, 178, 138),
                "INGOs": rgb_fill(190, 158, 242),
                "WFP": rgb_fill(44, 195, 236),
                "UN_AGENCIES": rgb_fill(0, 176, 240),
            }

            COLOR_MIN_COL = 1  # A
            COLOR_MAX_COL = 7  # G

            for r in range(2, ws.max_row + 1):
                intervention_val = ws.cell(row=r, column=intervention_col).value
                intervention_text = "" if intervention_val is None else str(intervention_val).strip()
                intervention_up = intervention_text.upper()

                row_fill = None

                # Priority: WFP
                if intervention_up == "WFP":
                    row_fill = fills["WFP"]

                # LOGISTICS renamed to UN Agencies and uses the *old* UN-OHCHR colour
                elif intervention_up == "UN AGENCIES":
                    row_fill = fills["UN_AGENCIES"]

                elif intervention_up == "TELECOMMUNICATIONS":
                    row_fill = fills["TELECOMMUNICATIONS"]
                elif intervention_up == "HEALTH":
                    row_fill = fills["HEALTH"]
                elif intervention_up == "WASH":
                    row_fill = fills["WASH"]
                elif intervention_up == "INGOS":
                    row_fill = fills["INGOs"]

                if row_fill is None:
                    continue

                for c in range(COLOR_MIN_COL, COLOR_MAX_COL + 1):
                    ws.cell(row=r, column=c).fill = row_fill

            progress.progress(100)

        # =====================================================
        # NEW: Create summary sheet (סקטור / כמות דלק (ליטר))
        # Uses totals per INTERVENTION (sum of Unified Fuel column F)
        # Keeps same colours per sector
        # =====================================================

        # 1) Build totals per INTERVENTION from the final cleaned sheet
        totals_by_intervention = {}
        for r in range(2, ws.max_row + 1):
            iv = ws.cell(row=r, column=intervention_col).value
            iv_key = "" if iv is None else str(iv).strip()
            if iv_key == "":
                continue
            totals_by_intervention[iv_key.upper()] = totals_by_intervention.get(iv_key.upper(), 0.0) + safe_float(
                ws.cell(row=r, column=unified_col).value  # Unified Fuel (F)
            )

        # 2) Define the required display order + Hebrew translations
        sector_rows = [
            ("TELECOMMUNICATIONS", "תקשורת"),
            ("HEALTH", "בריאות"),
            ("WASH", "סניטציה"),
            ("INGOS", "ארגונים לא ממשלתיים"),
            ("UN AGENCIES", 'סוכנויות או"ם'),
            ("WFP", "WFP"),
        ]

        # 3) Create/replace the sheet
        SUMMARY_SHEET_NAME = "Sector Summary"
        if SUMMARY_SHEET_NAME in wb.sheetnames:
            wb.remove(wb[SUMMARY_SHEET_NAME])
        ws_sum = wb.create_sheet(SUMMARY_SHEET_NAME)

        # 4) Headers (Hebrew)
        ws_sum["A1"].value = "סקטור"
        ws_sum["B1"].value = "כמות דלק (ליטר)"

        # Black background + white text
        header_fill = PatternFill("solid", fgColor="000000")
        header_font = Font(color="FFFFFF", bold=True)

        for cell_ref in ("A1", "B1", "C1"):   # C1 exists for %
            cell = ws_sum[cell_ref]
            cell.fill = header_fill
            cell.font = header_font

        # Optional: copy header styling from your main sheet header row (A1 style)
        hdr_src = ws.cell(row=1, column=1)
        for addr in ("A1", "B1"):
            c = ws_sum[addr]
            c.font = copy(hdr_src.font)
            c.fill = copy(hdr_src.fill)
            c.border = copy(hdr_src.border)
            c.alignment = copy(hdr_src.alignment)
            c.number_format = hdr_src.number_format
            c.protection = copy(hdr_src.protection)

        # Column widths (nice readable)
        ws_sum.column_dimensions["A"].width = 28
        ws_sum.column_dimensions["B"].width = 18

        # Number format for litres column (copy from Unified Fuel style if you want)
        num_src = ws.cell(row=2, column=unified_col)  # any body cell in Unified Fuel column
        litres_number_format = num_src.number_format

        # 5) Reuse the SAME colour fills you already defined for the main sheet
        # Make sure these keys match your current colour logic
        summary_fills = {
            "TELECOMMUNICATIONS": fills["TELECOMMUNICATIONS"],
            "HEALTH": fills["HEALTH"],
            "WASH": fills["WASH"],
            "INGOS": fills["INGOs"],
            "UN AGENCIES": fills.get("UN_AGENCIES", rgb_fill(0, 176, 240)),  # fallback just in case
            "WFP": fills["WFP"],
        }

        # 6.1) Map: INTERVENTION -> first row index (used as style source)
        style_row_by_intervention = {}
        for r in range(2, ws.max_row + 1):
            iv = ws.cell(row=r, column=intervention_col).value
            k = "" if iv is None else str(iv).strip().upper()
            if k and k not in style_row_by_intervention:
                style_row_by_intervention[k] = r

        # 6.2) Write summary rows using the main-sheet row styles
        row_i = 2
        for key_en, label_he in sector_rows:
            key_u = key_en.strip().upper()

            # pick a source row from the main sheet
            src_r = style_row_by_intervention.get(key_u)

            a_cell = ws_sum.cell(row=row_i, column=1)
            b_cell = ws_sum.cell(row=row_i, column=2)

            # values
            a_cell.value = label_he
            b_cell.value = totals_by_intervention.get(key_u, 0.0)

            if src_r is not None:
                # copy style from main sheet:
                # - column A style from the INTERVENTION cell (text look)
                # - column B style from Unified Fuel cell (number look)
                src_a = ws.cell(row=src_r, column=intervention_col)
                src_b = ws.cell(row=src_r, column=unified_col)

                copy_cell_style(src_a, a_cell)
                copy_cell_style(src_b, b_cell)
            else:
                # fallback (if category not found in data): at least keep header-ish formatting
                copy_cell_style(hdr_src, a_cell)
                copy_cell_style(hdr_src, b_cell)
                b_cell.number_format = litres_number_format

            # apply colours to A+B (same sector colour)
            fill_obj = summary_fills.get(key_u)
            if fill_obj:
                a_cell.fill = fill_obj
                b_cell.fill = fill_obj

            row_i += 1

        # ============================
        # EXCEL PIE CHART (inside the workbook) + % table under it
        # ============================

        FIRST_ROW = 2
        LAST_ROW = FIRST_ROW + len(sector_rows) - 1  # 7 if you have 6 sectors

        # Helper % column in C (so we can also show % under the chart)
        ws_sum["C1"].value = "אחוז"
        copy_cell_style(ws_sum["A1"], ws_sum["C1"])  # header style like the other headers
        ws_sum.column_dimensions["C"].width = 12

        for r in range(FIRST_ROW, LAST_ROW + 1):
            # =B2/SUM($B$2:$B$7)
            ws_sum.cell(row=r, column=3).value = f"=B{r}/SUM($B${FIRST_ROW}:$B${LAST_ROW})"
            # Copy number styling from the litres column, but format as percent
            copy_cell_style(ws_sum.cell(row=r, column=2), ws_sum.cell(row=r, column=3))
            ws_sum.cell(row=r, column=3).number_format = "0.0%"

        # Pie chart: values from B2:B7, labels from A2:A7
        pie = PieChart()
        data = Reference(ws_sum, min_col=2, min_row=1, max_row=LAST_ROW)      # includes header "כמות דלק (ליטר)"
        cats = Reference(ws_sum, min_col=1, min_row=FIRST_ROW, max_row=LAST_ROW)
        pie.add_data(data, titles_from_data=True)
        pie.set_categories(cats)

        pie.title = ""
        pie.dataLabels = DataLabelList()
        pie.dataLabels.showCatName = True     # show the sector name (Hebrew text in col A)
        pie.dataLabels.showPercent = True     # show %
        pie.dataLabels.showVal = False        # hide 78,674
        pie.dataLabels.showSerName = False    # hide "כמות דלק (ליטר)"
        pie.dataLabels.showLeaderLines = True
        pie.dataLabels.separator = "\n"       # put name and % on two lines
        pie.legend = None                     # hide list of categories

        # Match slice colours to your sector colours (hex, no '#')
        slice_colors = [
            "D5F3FB",  # TELECOMMUNICATIONS
            "00B050",  # HEALTH
            "FAB28A",  # WASH
            "BE9EF2",  # INGOs
            "00B0F0",  # UN Agencies
            "2CC3EC",  # WFP
        ]

        # Apply slice colours (1 series in a pie chart)
        ser = pie.series[0]
        ser.dPt = []
        for i, hx in enumerate(slice_colors[:len(sector_rows)]):
            dp = DataPoint(idx=i)
            dp.graphicalProperties.solidFill = hx
            ser.dPt.append(dp)

        # Position chart to the right of the table (adjust if you want)
        ws_sum.add_chart(pie, "E2")

        # ---------- Save + download ----------
        output = BytesIO()
        wb.save(output)
        output.seek(0)

        status.success("Done! Download your cleaned file below.")
        st.download_button(
            "Download cleaned Excel",
            data=output,
            file_name="cleaned.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    except Exception as e:
        status.empty()
        st.exception(e)

    finally:
        # ✅ Always remove the loader (even if an error happens)
        loader.empty()
