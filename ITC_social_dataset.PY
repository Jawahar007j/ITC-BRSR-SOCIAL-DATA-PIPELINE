import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SRC = '/mnt/user-data/uploads/ITC_Social_Dataset_2026.xlsx'
OUT = '/home/claude/work/ITC_Social_Dataset_2026_Refined.xlsx'

# ---- Metadata for each of the 35 tables: (sheet_name, title, principle) ----
META = {
    1:  ("Waste_Management", "Waste Management by Category — Re-used / Recycled / Safely Disposed", "Principle 2"),
    2:  ("Reclaimed_Packaging", "Reclaimed Products & Packaging as % of Total Products Sold", "Principle 2"),
    3:  ("Principle_3_Marker", "BRSR Principle 3 — Employee Well-being (Section Marker)", None),
    4:  ("Employee_Benefits", "Employees Covered by Insurance & Welfare Benefits", "Principle 3"),
    5:  ("Worker_Benefits", "Workers Covered by Insurance & Welfare Benefits", "Principle 3"),
    6:  ("Wellbeing_Cost", "Cost Incurred on Well-being Measures (% of Total Revenue)", "Principle 3"),
    7:  ("Statutory_Benefits", "Statutory Benefits Coverage — PF, Gratuity, ESI", "Principle 3"),
    8:  ("Parental_Leave", "Parental Leave — Return to Work & Retention Rate", "Principle 3"),
    9:  ("Grievance_Mechanism", "Grievance Redressal Mechanism — Availability & Description", "Principle 3"),
    10: ("Trade_Union_Membership", "Trade Union / Association Membership — Employees & Workers", "Principle 3"),
    11: ("Training_Coverage", "Training Coverage — Health & Safety and Skill Upgradation", "Principle 3"),
    12: ("Performance_Reviews", "Performance & Career Development Review Coverage", "Principle 3"),
    13: ("Safety_Incidents", "Safety Incident Metrics — LTIFR, Injuries, Fatalities", "Principle 3"),
    14: ("Safety_Complaints", "Complaints on Working Conditions & Health and Safety", "Principle 3"),
    15: ("Safety_Assessment", "Health & Safety / Working Conditions — Assessment Coverage", "Principle 3"),
    16: ("Rehabilitation", "Rehabilitation of Affected Employees / Workers", "Principle 3"),
    17: ("ValueChain_Safety", "Value Chain Partners — Health & Safety Assessment Coverage", "Principle 3"),
    18: ("Principle_4_Marker", "BRSR Principle 4 — Stakeholder Engagement (Section Marker)", None),
    19: ("Stakeholder_Engagement", "Stakeholder Engagement — Groups, Channels & Frequency", "Principle 4"),
    20: ("Principle_5_Marker", "BRSR Principle 5 — Human Rights (Section Marker)", None),
    21: ("HR_Training_Coverage", "Human Rights Training Coverage — Employees & Workers", "Principle 5"),
    22: ("Minimum_Wage", "Minimum Wage Compliance — Employees & Workers", "Principle 5"),
    23: ("Median_Remuneration", "Median Remuneration by Gender — BoD, KMP, Employees, Workers", "Principle 5"),
    24: ("Gender_Pay_Gap", "Gender Pay Gap — Gross Wages Paid to Females (% of Total)", "Principle 5"),
    25: ("HR_Complaints", "Human Rights Complaints — Filed & Pending Resolution", "Principle 5"),
    26: ("POSH_Complaints", "POSH Complaints Summary (Sexual Harassment at Workplace Act)", "Principle 5"),
    27: ("HR_Assessment_Internal", "Human Rights Assessment Coverage — Plants & Offices", "Principle 5"),
    28: ("HR_Assessment_ValueChain", "Human Rights Assessment Coverage — Value Chain Partners", "Principle 5"),
    29: ("Principle_6_Marker", "BRSR Principle 6 — Environment (Section Marker)", None),
    30: ("Energy_Consumption", "Energy Consumption — Renewable vs Non-Renewable Sources", "Principle 6"),
    31: ("Resettlement_RR", "Resettlement & Rehabilitation — Project Affected Families", "Principle 6/8"),
    32: ("Sourcing_by_Location", "Sourcing / Presence by Location Type", "Principle 8"),
    33: ("Sourcing_MSME_India", "Sourcing from MSMEs / Small Producers & Within India", "Principle 8"),
    34: ("CSR_Aspirational_Dist", "CSR Spend in Aspirational Districts (Rs. Lakhs)", "Principle 8"),
    35: ("CSR_Beneficiaries", "CSR Project Beneficiaries & Vulnerable Group Coverage (FY 2025-26)", "Principle 8"),
}

# ---- Styles ----
FONT_NAME = "Arial"
TITLE_FONT = Font(name=FONT_NAME, size=13, bold=True, color="FFFFFF")
SUBTITLE_FONT = Font(name=FONT_NAME, size=9, italic=True, color="595959")
HEADER_FONT = Font(name=FONT_NAME, size=10, bold=True, color="FFFFFF")
SECTION_FONT = Font(name=FONT_NAME, size=10, bold=True, italic=True, color="1F3864")
LABEL_FONT = Font(name=FONT_NAME, size=10, bold=True, color="000000")
DATA_FONT = Font(name=FONT_NAME, size=10, color="000000")

TITLE_FILL = PatternFill("solid", fgColor="1F3864")
HEADER_FILL = PatternFill("solid", fgColor="4472C4")
SECTION_FILL = PatternFill("solid", fgColor="D9E2F3")
ALT_FILL = PatternFill("solid", fgColor="F2F2F2")

THIN = Side(style="thin", color="B7B7B7")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT_WRAP = Alignment(horizontal="left", vertical="center", wrap_text=True)
RIGHT = Alignment(horizontal="right", vertical="center", wrap_text=True)


def is_blank(v):
    return v is None or (isinstance(v, str) and v.strip() == "")


def load_tables():
    wb = openpyxl.load_workbook(SRC, data_only=True)
    tables = {}
    for name in wb.sheetnames:
        idx = int(name.split("_")[1])
        ws = wb[name]
        rows = list(ws.iter_rows(values_only=True))
        # rows[0] is header ('PDF_Page', 0, 1, ...) -> skip
        data_rows = [list(r[1:]) for r in rows[1:]]
        pages = sorted(set(r[0] for r in rows[1:] if r[0] is not None))
        tables[idx] = {"rows": data_rows, "pages": pages}
    return tables


def _cell_is_data_like(v):
    if v is None:
        return False
    if not isinstance(v, str):
        return True
    s = v.strip()
    if s == "":
        return False
    low = s.lower()
    if low in ("na", "n/a", "-", "–", "not applicable"):
        return True
    s_no_fy = FY_PATTERN.sub("", s)
    if any(ch.isdigit() for ch in s_no_fy):
        return True
    if len(s) > 180:
        return True
    return False


def classify_rows(rows):
    """Split rows into header rows (pure label text, no data-like cells) vs data rows.
    The first row of a table is always treated as a header row."""
    header_count = 0
    for i, r in enumerate(rows):
        rest = r[1:]
        non_blank_rest = [c for c in rest if not is_blank(c)]
        if i == 0:
            header_count += 1
            continue
        if not non_blank_rest:
            break
        row_is_data = any(_cell_is_data_like(c) for c in non_blank_rest) or _cell_is_data_like(r[0])
        if row_is_data:
            break
        header_count += 1
    return rows[:header_count], rows[header_count:]


def _row_is_category(r):
    col0 = r[0]
    return isinstance(col0, str) and col0.strip() not in ("",)


def write_merges_for_row(ws, excel_row, values, start_col=1):
    """Merge consecutive None cells following a non-None cell within a single row."""
    n = len(values)
    c = 0
    while c < n:
        if not is_blank(values[c]):
            span_end = c
            k = c + 1
            while k < n and is_blank(values[k]):
                span_end = k
                k += 1
            if span_end > c:
                ws.merge_cells(start_row=excel_row, start_column=start_col + c,
                                end_row=excel_row, end_column=start_col + span_end)
            c = k
        else:
            c += 1


def autofit(ws, n_cols, widths):
    for i in range(n_cols):
        letter = get_column_letter(i + 1)
        ws.column_dimensions[letter].width = widths[i]


def build_principle_marker_sheet(wb, idx, meta):
    sheet_name, title, _ = meta
    ws = wb.create_sheet(sheet_name)
    ws.sheet_view.showGridLines = False
    ws.merge_cells("A1:F3")
    cell = ws["A1"]
    cell.value = title
    cell.font = Font(name=FONT_NAME, size=16, bold=True, color="FFFFFF")
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 60
    for r in range(1, 4):
        for c in range(1, 7):
            ws.cell(row=r, column=c).fill = PatternFill("solid", fgColor="1F3864")
    for i in range(6):
        ws.column_dimensions[get_column_letter(i + 1)].width = 18
    ws.freeze_panes = None


def build_table_sheet(wb, idx, meta, table):
    sheet_name, title, principle = meta
    rows = table["rows"]
    pages = table["pages"]
    n_cols = max(len(r) for r in rows)
    rows = [r + [None] * (n_cols - len(r)) for r in rows]

    header_rows, data_rows = classify_rows(rows)

    ws = wb.create_sheet(sheet_name)
    ws.sheet_view.showGridLines = False

    total_cols = n_cols
    last_col_letter = get_column_letter(total_cols)

    # --- Title bar ---
    ws.merge_cells(f"A1:{last_col_letter}1")
    t = ws["A1"]
    t.value = f"Table {idx} — {title}"
    t.font = TITLE_FONT
    t.fill = TITLE_FILL
    t.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 26

    # --- Subtitle bar ---
    ws.merge_cells(f"A2:{last_col_letter}2")
    page_str = ", ".join(str(p) for p in pages) if pages else "N/A"
    s = ws["A2"]
    principle_str = f"BRSR {principle} | " if principle else ""
    s.value = f"{principle_str}Source: ITC Report and Accounts 2026, Page {page_str} | FY 2025-26 vs FY 2024-25 (as applicable)"
    s.font = SUBTITLE_FONT
    s.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[2].height = 16

    excel_row = 4  # leave row 3 blank as spacer

    # --- Header rows ---
    header_start_row = excel_row
    for r in header_rows:
        for c_idx, val in enumerate(r):
            cell = ws.cell(row=excel_row, column=c_idx + 1)
            if not is_blank(val):
                cell.value = val.replace("\n", " ") if isinstance(val, str) else val
            cell.font = HEADER_FONT
            cell.fill = HEADER_FILL
            cell.alignment = CENTER
            cell.border = BORDER
        write_merges_for_row(ws, excel_row, r)
        ws.row_dimensions[excel_row].height = 30
        excel_row += 1
    header_end_row = excel_row - 1

    # --- Data rows ---
    data_start_row = excel_row
    alt = False
    for r in data_rows:
        non_blank_rest = [c for c in r[1:] if not is_blank(c)]
        is_section_header = is_blank(r[0]) is False and len(non_blank_rest) == 0 and not is_blank(r[0])
        col0_blank = is_blank(r[0])
        row_has_data = any(not is_blank(c) for c in r)
        if not row_has_data:
            continue

        for c_idx, val in enumerate(r):
            cell = ws.cell(row=excel_row, column=c_idx + 1)
            if not is_blank(val):
                cell.value = val.replace("\n", " ") if isinstance(val, str) else val
            cell.border = BORDER
            if c_idx == 0:
                cell.alignment = LEFT_WRAP
            elif isinstance(val, str) and len(val) > 25:
                cell.alignment = LEFT_WRAP
            else:
                cell.alignment = RIGHT if c_idx > 0 else LEFT_WRAP

        if is_section_header:
            # merge across remaining columns, style as section divider
            ws.merge_cells(start_row=excel_row, start_column=2, end_row=excel_row, end_column=total_cols)
            for c_idx in range(total_cols):
                cell = ws.cell(row=excel_row, column=c_idx + 1)
                cell.font = SECTION_FONT
                cell.fill = SECTION_FILL
            ws.cell(row=excel_row, column=1).font = SECTION_FONT
        else:
            for c_idx in range(total_cols):
                cell = ws.cell(row=excel_row, column=c_idx + 1)
                if c_idx == 0:
                    cell.font = LABEL_FONT
                else:
                    cell.font = DATA_FONT
                if alt:
                    cell.fill = ALT_FILL

        # taller rows for long wrapped text
        max_chars = max((len(v) for v in r if isinstance(v, str)), default=0)
        if max_chars > 150:
            ws.row_dimensions[excel_row].height = min(300, 15 * (max_chars // 90 + 1))

        excel_row += 1
        alt = not alt

    data_end_row = excel_row - 1


    # --- Column widths ---
    widths = []
    for c_idx in range(total_cols):
        max_len = 12
        for r in header_rows + data_rows:
            v = r[c_idx] if c_idx < len(r) else None
            if isinstance(v, str):
                pieces = v.replace("\n", " ")
                max_len = max(max_len, min(len(pieces), 45))
        widths.append(max(12, min(max_len + 2, 48)))
    if total_cols > 0:
        widths[0] = max(widths[0], 30)
    autofit(ws, total_cols, widths)

    ws.freeze_panes = ws.cell(row=header_end_row + 1, column=2)


import re

LABEL_COLS_OVERRIDE = {13: 2, 34: 2}  # tables with 2 leading identity columns
FY_PATTERN = re.compile(r"FY\s?20\d{2}-\d{2}")


REF_CODE_PATTERN = re.compile(r"\s*\([A-F](?:/[A-F])?\)\s*$")


def clean_text(v):
    if isinstance(v, str):
        s = v.replace("\n", " ").strip()
        return s
    return v


def strip_ref_code(s):
    return REF_CODE_PATTERN.sub("", s).strip()


def parse_numeric(v):
    if isinstance(v, (int, float)):
        return float(v)
    if not isinstance(v, str):
        return None
    s = v.strip()
    low = s.lower()
    if low in ("", "-", "na", "n/a", "not applicable", "–"):
        return None
    s_clean = s.replace(",", "")
    m = re.search(r"[-+]?\d+\.?\d*", s_clean)
    if not m:
        return None
    try:
        return float(m.group(0))
    except ValueError:
        return None


def fill_headers_hierarchical(header_rows, n_cols):
    """Forward-fill each header level, but reset the carry whenever the
    combination of parent (higher) levels changes at that column — this
    prevents a sub-header label from bleeding into an unrelated group."""
    n_levels = len(header_rows)
    filled = [[None] * n_cols for _ in range(n_levels)]
    for lvl in range(n_levels):
        last = None
        last_parent_key = object()
        for c in range(n_cols):
            parent_key = tuple(filled[p][c] for p in range(lvl))
            if parent_key != last_parent_key:
                last = None
                last_parent_key = parent_key
            v = header_rows[lvl][c] if c < len(header_rows[lvl]) else None
            if not is_blank(v):
                last = clean_text(v)
            filled[lvl][c] = last
    return filled


def build_metric_tables(idx, rows):
    """Returns an ordered list of mini cross-tab tables, one per distinct
    metric/particular: {metric, unit, row_order:[...], fy_order:[...], data:{(row,fy):(num)}}"""
    n_cols = max(len(r) for r in rows)
    rows = [list(r) + [None] * (n_cols - len(r)) for r in rows]
    header_rows, data_rows = classify_rows(rows)
    label_cols = LABEL_COLS_OVERRIDE.get(idx, 1)

    filled_headers = fill_headers_hierarchical(header_rows, n_cols)

    col_desc = []
    for c in range(label_cols, n_cols):
        levels, prev = [], None
        for lvl in range(len(filled_headers)):
            lv = filled_headers[lvl][c]
            if lv is not None and lv != prev:
                cleaned = strip_ref_code(lv)
                if cleaned:
                    levels.append(cleaned)
            prev = lv
        col_desc.append(levels)

    col_fy, col_metric = [], []
    FILLER = {"current financial year", "previous financial year", "current year", "previous year"}
    for levels in col_desc:
        fy_val, rest = None, []
        for lv in levels:
            m = FY_PATTERN.search(lv)
            if m:
                fy_val = m.group(0)
            elif lv.strip().lower() not in FILLER:
                rest.append(lv)
        col_fy.append(fy_val)
        col_metric.append(" - ".join(rest) if rest else "Value")

    tables = {}   # metric -> dict
    order = []    # preserve first-seen order of metrics
    group_label = None
    last_parent = None
    SUBROW_LABELS = {"male", "female", "total"}
    last_labels = [None] * label_cols

    for r in data_rows:
        rest_vals = [r[c] for c in range(label_cols, n_cols)]
        non_blank_rest = [v for v in rest_vals if not is_blank(v)]
        label_vals = [r[c] for c in range(label_cols)]
        any_label = any(not is_blank(v) for v in label_vals)

        if len(non_blank_rest) == 0 and any_label:
            group_label = clean_text(label_vals[0])
            last_parent = None
            last_labels = [None] * label_cols
            continue
        if not non_blank_rest:
            continue

        cur_labels = []
        for i, v in enumerate(label_vals):
            if not is_blank(v):
                last_labels[i] = clean_text(v)
            cur_labels.append(last_labels[i])

        primary_label = cur_labels[0] if cur_labels else None
        clean_primary = re.sub(r"^[-–]\s*", "", primary_label).strip() if primary_label else primary_label
        norm = clean_primary.strip().lower() if clean_primary else ""
        extra_labels = [x for x in cur_labels[1:] if x]

        if norm in SUBROW_LABELS and last_parent:
            row_key_parts = ([group_label] if group_label else []) + [last_parent, clean_primary] + extra_labels
        else:
            row_key_parts = ([group_label] if group_label else []) + [clean_primary] + extra_labels
            if clean_primary and norm not in SUBROW_LABELS:
                last_parent = clean_primary
        row_key = " – ".join([p for p in row_key_parts if p]) or "Value"

        for c, val in enumerate(rest_vals):
            if is_blank(val):
                continue
            metric = col_metric[c]
            fy = col_fy[c] or "Value"
            val_text = clean_text(val)
            if metric not in tables:
                tables[metric] = {"metric": metric, "unit": None, "row_order": [], "fy_order": [], "data": {}}
                order.append(metric)
            tbl = tables[metric]
            if row_key not in tbl["row_order"]:
                tbl["row_order"].append(row_key)
            if fy not in tbl["fy_order"]:
                tbl["fy_order"].append(fy)
            num = parse_numeric(val_text)
            if "%" in val_text:
                tbl["unit"] = "%"
            elif tbl["unit"] is None and num is not None:
                tbl["unit"] = "Number"
            tbl["data"][(row_key, fy)] = num

    return [tables[m] for m in order]


def build_tables_sheet(wb, idx, meta, table):
    sheet_name, title, principle = meta
    rows = table["rows"]
    pages = table["pages"]
    metric_tables = build_metric_tables(idx, rows)
    metric_tables = [t for t in metric_tables if any(v is not None for v in t["data"].values())]
    if not metric_tables:
        return

    base = sheet_name[:24] if len(sheet_name) > 24 else sheet_name
    ws = wb.create_sheet(f"{base}_Tables")
    ws.sheet_view.showGridLines = False
    max_cols = max(2, max(len(t["fy_order"]) + 1 for t in metric_tables))
    last_col_letter = get_column_letter(max_cols)

    ws.merge_cells(f"A1:{last_col_letter}1")
    t0 = ws["A1"]
    t0.value = f"Table {idx} — {title} (Chart-Ready Tables)"
    t0.font = TITLE_FONT
    t0.fill = TITLE_FILL
    t0.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 26

    page_str = ", ".join(str(p) for p in pages) if pages else "N/A"
    principle_str = f"BRSR {principle} | " if principle else ""
    ws.merge_cells(f"A2:{last_col_letter}2")
    s0 = ws["A2"]
    s0.value = f"{principle_str}Source: ITC Report and Accounts 2026, Page {page_str} | Each block below is a self-contained table — select it and Insert > Chart directly."
    s0.font = SUBTITLE_FONT
    s0.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[2].height = 16

    excel_row = 4
    max_label_len = 12

    for tbl in metric_tables:
        metric = tbl["metric"]
        unit = tbl["unit"] or "Number"
        fy_order = tbl["fy_order"]
        row_order = tbl["row_order"]
        n_data_cols = len(fy_order)
        this_last_col = get_column_letter(1 + n_data_cols)

        ws.merge_cells(start_row=excel_row, start_column=1, end_row=excel_row, end_column=1 + n_data_cols)
        mcell = ws.cell(row=excel_row, column=1, value=f"{metric}  (Unit: {unit})")
        mcell.font = Font(name=FONT_NAME, size=10, bold=True, color="FFFFFF")
        mcell.fill = PatternFill("solid", fgColor="2E5395")
        mcell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[excel_row].height = 20
        excel_row += 1

        header_r = excel_row
        ws.cell(row=header_r, column=1, value="Category")
        for j, fy in enumerate(fy_order):
            ws.cell(row=header_r, column=2 + j, value=fy)
        for c_idx in range(1 + n_data_cols):
            cell = ws.cell(row=header_r, column=c_idx + 1)
            cell.font = HEADER_FONT
            cell.fill = HEADER_FILL
            cell.alignment = CENTER
            cell.border = BORDER
        excel_row += 1

        for r_i, row_key in enumerate(row_order):
            max_label_len = max(max_label_len, len(row_key))
            cell0 = ws.cell(row=excel_row, column=1, value=row_key)
            cell0.font = LABEL_FONT
            cell0.border = BORDER
            cell0.alignment = LEFT_WRAP
            if r_i % 2 == 1:
                cell0.fill = ALT_FILL
            for j, fy in enumerate(fy_order):
                num = tbl["data"].get((row_key, fy))
                cell = ws.cell(row=excel_row, column=2 + j)
                cell.border = BORDER
                cell.font = DATA_FONT
                cell.alignment = RIGHT
                if r_i % 2 == 1:
                    cell.fill = ALT_FILL
                if num is not None:
                    if unit == "%":
                        cell.value = num / 100.0
                        cell.number_format = "0.0%"
                    else:
                        cell.value = num
                        cell.number_format = "#,##0.##"
            excel_row += 1

        excel_row += 1  # spacer between tables

    widths = [max(18, min(max_label_len + 2, 50))] + [14] * (max_cols - 1)
    autofit(ws, max_cols, widths)


def main():
    tables = load_tables()
    wb = Workbook()
    wb.remove(wb.active)

    # Cover sheet
    cover = wb.create_sheet("Cover")
    cover.sheet_view.showGridLines = False
    cover.merge_cells("A1:F3")
    c = cover["A1"]
    c.value = "ITC Limited — BRSR Social Performance Dataset (Refined)"
    c.font = Font(name=FONT_NAME, size=18, bold=True, color="FFFFFF")
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for r in range(1, 4):
        for cc in range(1, 7):
            cover.cell(row=r, column=cc).fill = PatternFill("solid", fgColor="1F3864")
    cover.row_dimensions[1].height = 70

    cover["A5"] = "Source:"
    cover["B5"] = "ITC Limited, Report and Accounts 2026 (Business Responsibility & Sustainability Report — Social Metrics)"
    cover["A6"] = "Coverage:"
    cover["B6"] = "FY 2025-26 (Current Financial Year) vs FY 2024-25 (Previous Financial Year), as disclosed"
    cover["A7"] = "Scope:"
    cover["B7"] = "BRSR Principles 3–6 & 8 — Employee Well-being, Stakeholder Engagement, Human Rights, Environment, CSR"
    cover["A8"] = "Note:"
    cover["B8"] = "All figures reproduced exactly as extracted from the source report. No values have been altered, estimated, or recalculated."
    for row in range(5, 9):
        cover.cell(row=row, column=1).font = Font(name=FONT_NAME, bold=True, size=10)
        cover.cell(row=row, column=2).font = Font(name=FONT_NAME, size=10)
        cover.cell(row=row, column=2).alignment = Alignment(wrap_text=True, vertical="center")
    cover.column_dimensions["A"].width = 12
    cover.column_dimensions["B"].width = 95
    for row in range(5, 9):
        cover.row_dimensions[row].height = 30

    cover["A10"] = "Sheet Index"
    cover["A10"].font = Font(name=FONT_NAME, bold=True, size=12)
    r = 11
    cover.cell(row=r, column=1, value="No.").font = Font(name=FONT_NAME, bold=True)
    cover.cell(row=r, column=2, value="Sheet Name").font = Font(name=FONT_NAME, bold=True)
    cover.cell(row=r, column=3, value="Contents").font = Font(name=FONT_NAME, bold=True)
    for cc in (1, 2, 3):
        cover.cell(row=r, column=cc).fill = HEADER_FILL
        cover.cell(row=r, column=cc).font = HEADER_FONT
    r += 1
    for idx in range(1, 36):
        sheet_name, title, principle = META[idx]
        cover.cell(row=r, column=1, value=idx)
        cover.cell(row=r, column=2, value=sheet_name)
        cover.cell(row=r, column=3, value=title)
        r += 1
    cover.column_dimensions["C"].width = 70

    for idx in range(1, 36):
        meta = META[idx]
        if meta[2] is None and "Marker" in meta[0]:
            build_principle_marker_sheet(wb, idx, meta)
        else:
            build_table_sheet(wb, idx, meta, tables[idx])
            build_tables_sheet(wb, idx, meta, tables[idx])

    wb.save(OUT)
    print("Saved:", OUT)


if __name__ == "__main__":
    main()
