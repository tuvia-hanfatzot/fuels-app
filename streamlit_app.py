import streamlit as st
from io import BytesIO
from copy import copy
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill

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

# ----------------------------
# Processing
# ----------------------------
if uploaded:
    progress = st.progress(0)
    status = st.empty()

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
            with st.spinner("Processing…"):
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

                # ---------- COLOR A–G BASED ON INTERVENTION ----------
                status.info("Applying colours…")
                fills = {
                    "TELECOMMUNICATIONS": rgb_fill(213, 243, 251),
                    "HEALTH": rgb_fill(0, 176, 80),
                    "WASH_OR_LOGISTICS": rgb_fill(250, 178, 138),
                    "INGOs": rgb_fill(190, 158, 242),
                    "UNOCHR_SUBSTR": rgb_fill(0, 176, 240),   # contains "UN-OHCHR"
                    "WFP": rgb_fill(44, 195, 236),
                }

                COLOR_MIN_COL = 1  # A
                COLOR_MAX_COL = 7  # G

                for r in range(2, ws.max_row + 1):
                    intervention_val = ws.cell(row=r, column=intervention_col).value
                    intervention_text = "" if intervention_val is None else str(intervention_val).strip()
                    intervention_up = intervention_text.upper()

                    row_fill = None
                    if "UN-OHCHR" in intervention_up:
                        row_fill = fills["UNOCHR_SUBSTR"]
                    elif intervention_up == "TELECOMMUNICATIONS":
                        row_fill = fills["TELECOMMUNICATIONS"]
                    elif intervention_up == "HEALTH":
                        row_fill = fills["HEALTH"]
                    elif intervention_up == "WASH" or intervention_up == "LOGISTICS":
                        row_fill = fills["WASH_OR_LOGISTICS"]
                    elif intervention_up == "INGOS":
                        row_fill = fills["INGOs"]
                    elif intervention_up == "WFP":
                        row_fill = fills["WFP"]

                    if row_fill is None:
                        continue

                    for c in range(COLOR_MIN_COL, COLOR_MAX_COL + 1):
                        ws.cell(row=r, column=c).fill = row_fill

                progress.progress(100)

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
