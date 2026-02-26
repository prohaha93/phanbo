import pandas as pd
import numpy as np
import pulp
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from io import BytesIO          # ADDED: để tạo luồng nhị phân cho file kết quả

# ============================================================
# COLOR PALETTE (from sample file)
# ============================================================
C_DARK_BLUE   = "FF1F4E79"  # STS header fill
C_MID_BLUE    = "FF2E75B6"  # BAY header fill
C_LIGHT_BLUE  = "FF9DC3E6"  # BLOCK header fill (MATRIX)
C_PALE_BLUE   = "FFD6E4F0"  # Hour cell fill / title fill
C_ALT_ROW     = "FFEBF3FB"  # Alternating row fill (odd)
C_WHITE       = "FFFFFFFF"  # Even row fill
C_YELLOW      = "FFFFF2CC"  # TOTAL cell fill (row/col totals)
C_GREEN       = "FF375623"  # Grand TOTAL header fill
C_TITLE_BG    = "FFDEEAF1"  # Title row background (DETAIL)
C_TITLE_BG_M  = "FFD6E4F0"  # Title row background (MATRIX)
C_ORANGE_FILL = "FFFCE4D6"  # WC cell fill
C_ORANGE_FONT = "FF833C00"  # WC cell font color
C_GREY_FONT   = "FFBFBFBF"  # Dash "—" font color
C_HEADER_BG   = "FFDEEAF1"  # Title background DETAIL

FONT_NAME = "Calibri"

def _font(bold=False, color="FF000000", size=10):
    return Font(name=FONT_NAME, bold=bold, color=color, size=size)

def _fill(color):
    return PatternFill("solid", fgColor=color)

def _align(h="center", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def _thin_border():
    s = Side(border_style="thin", color="FF000000")
    return Border(left=s, right=s, top=s, bottom=s)

def _thick_border():
    s = Side(border_style="medium", color="FF000000")
    return Border(left=s, right=s, top=s, bottom=s)

def _style(ws, coord, value=None, bold=False, font_color="FF000000",
           fill_color=None, align="center", wrap=False, border=True, size=10):
    cell = ws[coord] if isinstance(coord, str) else coord
    if value is not None:
        cell.value = value
    cell.font = _font(bold=bold, color=font_color, size=size)
    if fill_color:
        cell.fill = _fill(fill_color)
    cell.alignment = _align(h=align, wrap=wrap)
    if border:
        cell.border = _thin_border()
    return cell

# ============================================================
# HÀM CHÍNH: run_optimization
# Nhận đầu vào là file Excel (đường dẫn hoặc file-like object)
# Trả về: (bytes của file kết quả, tổng số dòng phân bổ, giá trị hàm mục tiêu)
# ============================================================
def run_optimization(input_file):
    """
    Thực hiện toàn bộ quá trình đọc dữ liệu, tối ưu và tạo file Excel kết quả.
    input_file: đường dẫn (string) hoặc đối tượng file-like (BytesIO) chứa file Excel đầu vào.
    Trả về: (excel_bytes, total_rows, objective_value)
    """
    # ============================================================
    # 1. Read and parse original data (sử dụng input_file thay vì tên cứng)
    # ============================================================
    # MODIFIED: dùng pd.ExcelFile trên input_file (chấp nhận cả string và file-like)
    xls = pd.ExcelFile(input_file)

    df1 = pd.read_excel(xls, sheet_name='MOVEHOUR-WEIGHTCLASS', header=None)
    sts_bay_map = {}
    for col in range(2, df1.shape[1]):
        sts = df1.iloc[0, col]
        bay = df1.iloc[1, col]
        if pd.notna(sts) and pd.notna(bay):
            sts_bay_map[col] = (str(sts).strip(), str(bay).strip())

    demands = {}
    current_hour = None
    for idx in range(2, df1.shape[0]):
        row = df1.iloc[idx]
        hour = row[0]
        if pd.isna(hour):
            hour = current_hour
        else:
            current_hour = hour
        weight = row[1]
        if pd.isna(weight):
            continue
        weight = int(weight)
        for col in range(2, df1.shape[1]):
            qty = row[col]
            if pd.notna(qty) and qty != '':
                qty = int(qty)
                if qty > 0:
                    sts, bay = sts_bay_map[col]
                    key = (hour, sts, bay)
                    if key not in demands:
                        demands[key] = {}
                    demands[key][weight] = demands[key].get(weight, 0) + qty

    job_keys = list(demands.keys())
    jobs_by_hour = {}
    for (h, s, b) in job_keys:
        jobs_by_hour.setdefault(h, []).append((s, b))

    df2 = pd.read_excel(xls, sheet_name='BLOCK-WEIGHT CLASS', header=0)
    supply = {}
    for idx, row in df2.iterrows():
        block = str(row.iloc[0]).strip()
        w1 = row.iloc[1] if pd.notna(row.iloc[1]) else 0
        w2 = row.iloc[2] if pd.notna(row.iloc[2]) else 0
        w3 = row.iloc[3] if pd.notna(row.iloc[3]) else 0
        w4 = row.iloc[4] if pd.notna(row.iloc[4]) else 0
        supply[block] = {1: int(w1), 2: int(w2), 3: int(w3), 4: int(w4)}

    blocks = list(supply.keys())
    weight_classes = [1, 2, 3, 4]

    # ============================================================
    # 2. Check demand vs supply
    # ============================================================
    total_demand = {w: 0 for w in weight_classes}
    for job in job_keys:
        for w, qty in demands[job].items():
            total_demand[w] += qty
    total_supply = {w: 0 for w in weight_classes}
    for block in blocks:
        for w in weight_classes:
            total_supply[w] += supply[block][w]

    for w in weight_classes:
        if total_demand[w] != total_supply[w]:
            print(f"ERROR: Mismatch for weight class {w}: demand={total_demand[w]}, supply={total_supply[w]}")
            raise ValueError(f"Tổng cầu và cung không khớp cho weight class {w}")

    # ============================================================
    # 3. Solve optimisation
    # ============================================================
    prob = pulp.LpProblem("Minimize_Clashes", pulp.LpMinimize)
    y_vars = {}
    for (h, s, bay) in job_keys:
        for b in blocks:
            y_vars[(h, s, bay, b)] = pulp.LpVariable(f"y_{h}_{s}_{bay}_{b}", cat='Binary')

    x_vars = {}
    for (h, s, bay) in job_keys:
        for w in demands[(h, s, bay)]:
            for b in blocks:
                x_vars[(h, s, bay, b, w)] = pulp.LpVariable(f"x_{h}_{s}_{bay}_{b}_{w}", lowBound=0, cat='Integer')

    u_vars = {}
    e_vars = {}
    for h in jobs_by_hour:
        for b in blocks:
            u_vars[(h, b)] = pulp.LpVariable(f"u_{h}_{b}", lowBound=0, cat='Integer')
            e_vars[(h, b)] = pulp.LpVariable(f"e_{h}_{b}", lowBound=0, cat='Integer')
            prob += u_vars[(h, b)] == pulp.lpSum(y_vars[(h, s, bay, b)] for (s, bay) in jobs_by_hour[h])
            prob += e_vars[(h, b)] >= u_vars[(h, b)] - 1

    prob += pulp.lpSum(e_vars.values())

    for (h, s, bay) in job_keys:
        for w, d in demands[(h, s, bay)].items():
            prob += pulp.lpSum(x_vars[(h, s, bay, b, w)] for b in blocks) == d

    for b in blocks:
        for w in weight_classes:
            prob += pulp.lpSum(
                x_vars[(h, s, bay, b, w)]
                for (h, s, bay) in job_keys if w in demands[(h, s, bay)]
            ) <= supply[b][w]

    for (h, s, bay) in job_keys:
        for w in demands[(h, s, bay)]:
            d = demands[(h, s, bay)][w]
            for b in blocks:
                prob += x_vars[(h, s, bay, b, w)] <= d * y_vars[(h, s, bay, b)]

    solver = pulp.PULP_CBC_CMD(msg=True, timeLimit=300)
    prob.solve(solver)

    status = prob.status
    print(f"Status: {pulp.LpStatus[status]}")
    if status not in (1,):
        print("No optimal solution found.")
        raise ValueError("Không tìm được lời giải tối ưu.")

    # ============================================================
    # 4. Extract result
    # ============================================================
    result_rows = []
    for (h, s, bay, b) in y_vars:
        if y_vars[(h, s, bay, b)].varValue > 0.5:
            for w in demands[(h, s, bay)]:
                qty = x_vars[(h, s, bay, b, w)].varValue
                if qty > 0.5:
                    result_rows.append({
                        'MOVE HOUR': h, 'STS': s, 'BAY': bay,
                        'ASSIGNED BLOCK': b, 'WEIGHT CLASS': w, 'QUANTITIES': int(qty)
                    })

    df_result = pd.DataFrame(result_rows)
    df_result.sort_values(['MOVE HOUR', 'STS', 'BAY', 'ASSIGNED BLOCK'], inplace=True)

    # ============================================================
    # 5. Prepare MATRIX data
    # ============================================================
    df_matrix_base = df_result.groupby(
        ['MOVE HOUR', 'STS', 'BAY', 'ASSIGNED BLOCK'], as_index=False
    )['QUANTITIES'].sum()

    sts_list = sorted(df_result['STS'].unique(), key=lambda x: int(x.replace('STS', '')))
    sts_bay_blocks = {}

    def _first_hour(sts, bay, df):
        """Return the earliest MOVE HOUR for a given STS+BAY (used for sorting bays)."""
        hours = df[(df['STS'] == sts) & (df['BAY'] == bay)]['MOVE HOUR'].unique()
        return sorted(hours)[0]

    for sts in sts_list:
        bays = df_result[df_result['STS'] == sts]['BAY'].unique()
        # Sort bays by their earliest MOVE HOUR so they appear in chronological order
        bays = sorted(bays, key=lambda bay: _first_hour(sts, bay, df_result))
        sts_bay_blocks[sts] = {}
        for bay in bays:
            blks = df_result[(df_result['STS'] == sts) & (df_result['BAY'] == bay)]['ASSIGNED BLOCK'].unique()
            sts_bay_blocks[sts][bay] = sorted(blks)

    matrix_cols = []
    for sts in sts_list:
        for bay in sts_bay_blocks[sts]:
            for block in sts_bay_blocks[sts][bay]:
                matrix_cols.append((sts, bay, block))

    hour_list = sorted(df_matrix_base['MOVE HOUR'].unique())
    matrix_data = {}
    for h in hour_list:
        matrix_data[h] = {}
        for col in matrix_cols:
            matrix_data[h][col] = 0

    for _, row in df_matrix_base.iterrows():
        key = (row['STS'], row['BAY'], row['ASSIGNED BLOCK'])
        if key in matrix_data.get(row['MOVE HOUR'], {}):
            matrix_data[row['MOVE HOUR']][key] = row['QUANTITIES']

    # ============================================================
    # 6. Prepare DETAIL groups (one table per STS with stacked bays)
    # ============================================================
    # Group: for each (STS, BAY) we have a sub-table
    # Tables are placed side by side per STS
    # Within one STS column, bays are stacked vertically

    detail_groups_by_sts = {}
    for sts in sts_list:
        detail_groups_by_sts[sts] = []
        for bay in sts_bay_blocks[sts]:
            df_sub = df_result[(df_result['STS'] == sts) & (df_result['BAY'] == bay)]
            rows_idx = df_sub[['MOVE HOUR', 'WEIGHT CLASS']].drop_duplicates().sort_values(['MOVE HOUR', 'WEIGHT CLASS'])
            blks = sorted(df_sub['ASSIGNED BLOCK'].unique())
            pivot = df_sub.pivot_table(index=['MOVE HOUR', 'WEIGHT CLASS'], columns='ASSIGNED BLOCK',
                                       values='QUANTITIES', fill_value=0, aggfunc='sum')
            pivot = pivot.reindex(
                index=pd.MultiIndex.from_tuples([(h, w) for h, w in rows_idx.values]),
                columns=blks, fill_value=0
            )
            table_rows = []
            for (hour, wc) in pivot.index:
                row_data = [hour, wc] + [int(pivot.loc[(hour, wc), b]) for b in blks] + [int(pivot.loc[(hour, wc)].sum())]
                table_rows.append(row_data)

            # Column totals (excluding hour and wc cols)
            col_totals = [None, None]
            for b in blks:
                col_totals.append(int(pivot[b].sum()))
            col_totals.append(int(pivot.values.sum()))

            detail_groups_by_sts[sts].append({
                'bay': bay,
                'blocks': blks,
                'rows': table_rows,
                'col_totals': col_totals,
                'num_rows': len(table_rows)
            })

    # Compute table width for each STS (max across its bays: 2 + max_blocks + 1)
    sts_table_widths = {}
    for sts in sts_list:
        max_w = max(2 + len(g['blocks']) + 1 for g in detail_groups_by_sts[sts])
        sts_table_widths[sts] = max_w

    # ============================================================
    # 7. Write Excel with openpyxl
    # ============================================================
    import openpyxl
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # --- SHEET: MOVEHOUR-WEIGHTCLASS ---
    ws_mh = wb.create_sheet('MOVEHOUR-WEIGHTCLASS')
    for r_idx, row in enumerate(df1.values, 1):
        for c_idx, val in enumerate(row, 1):
            ws_mh.cell(row=r_idx, column=c_idx, value=val if pd.notna(val) else None)

    # --- SHEET: BLOCK-WEIGHT CLASS ---
    ws_bw = wb.create_sheet('BLOCK-WEIGHT CLASS')
    headers = list(df2.columns)
    for c_idx, h in enumerate(headers, 1):
        ws_bw.cell(row=1, column=c_idx, value=h)
    for r_idx, row in enumerate(df2.values, 2):
        for c_idx, val in enumerate(row, 1):
            ws_bw.cell(row=r_idx, column=c_idx, value=val if pd.notna(val) else None)

    # ============================================================
    # SHEET: RESULT  (formatted like sample)
    # ============================================================
    ws_result = wb.create_sheet('RESULT')
    col_names = ['MOVE HOUR', 'STS', 'BAY', 'ASSIGNED BLOCK', 'WEIGHT CLASS', 'QUANTITIES']
    for c_idx, cn in enumerate(col_names, 1):
        cell = ws_result.cell(row=1, column=c_idx, value=cn)
        cell.font = _font(bold=True, color=C_WHITE)
        cell.fill = _fill(C_DARK_BLUE)
        cell.alignment = _align()
        cell.border = _thin_border()

    for r_idx, (_, row) in enumerate(df_result.iterrows(), 2):
        fill_color = C_ALT_ROW if (r_idx % 2 == 0) else C_WHITE
        for c_idx, cn in enumerate(col_names, 1):
            cell = ws_result.cell(row=r_idx, column=c_idx, value=row[cn])
            cell.font = _font(color="FF000000")
            cell.fill = _fill(fill_color)
            cell.alignment = _align()
            cell.border = _thin_border()

    for c in range(1, 7):
        ws_result.column_dimensions[get_column_letter(c)].width = 16

    # ============================================================
    # SHEET: MATRIX  (formatted like sample)
    # ============================================================
    ws_matrix = wb.create_sheet('MATRIX')

    # Title row 1
    total_matrix_cols = len(matrix_cols) + 2  # +1 for MOVE HOUR col, +1 for TOTAL col
    title_cell = ws_matrix.cell(row=1, column=1, value='MA TRẬN PHÂN BỔ BLOCK  ▸  MOVE HOUR × STS / BAY / BLOCK')
    title_cell.font = _font(bold=True, color=C_DARK_BLUE, size=11)
    title_cell.fill = _fill(C_TITLE_BG_M)
    title_cell.alignment = _align()
    title_cell.border = _thin_border()
    ws_matrix.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_matrix_cols)
    ws_matrix.row_dimensions[1].height = 16.8

    # Row 2: MOVE HOUR (merged rows 2-4), then STS headers, then TOTAL
    # Merge A2:A4 for "MOVE\nHOUR"
    mh_cell = ws_matrix.cell(row=2, column=1, value='MOVE\nHOUR')
    mh_cell.font = _font(bold=True, color=C_WHITE)
    mh_cell.fill = _fill(C_DARK_BLUE)
    mh_cell.alignment = _align(wrap=True)
    mh_cell.border = _thin_border()
    ws_matrix.merge_cells(start_row=2, start_column=1, end_row=4, end_column=1)

    # TOTAL header (merge rows 2-4)
    total_col = total_matrix_cols
    tc = ws_matrix.cell(row=2, column=total_col, value='TOTAL')
    tc.font = _font(bold=True, color=C_WHITE)
    tc.fill = _fill(C_GREEN)
    tc.alignment = _align()
    tc.border = _thin_border()
    ws_matrix.merge_cells(start_row=2, start_column=total_col, end_row=4, end_column=total_col)

    # STS headers row 2, BAY headers row 3, BLOCK headers row 4
    col_offset = 2
    for sts in sts_list:
        sts_start = col_offset
        for bay in sts_bay_blocks[sts]:
            bay_start = col_offset
            for block in sts_bay_blocks[sts][bay]:
                # Row 4: block
                bc = ws_matrix.cell(row=4, column=col_offset, value=block)
                bc.font = _font(bold=True, color=C_DARK_BLUE)
                bc.fill = _fill(C_LIGHT_BLUE)
                bc.alignment = _align()
                bc.border = _thin_border()
                col_offset += 1
            # Merge bay cells in row 3
            bay_end = col_offset - 1
            bayc = ws_matrix.cell(row=3, column=bay_start, value=bay)
            bayc.font = _font(bold=True, color=C_WHITE)
            bayc.fill = _fill(C_MID_BLUE)
            bayc.alignment = _align()
            bayc.border = _thin_border()
            if bay_start < bay_end:
                ws_matrix.merge_cells(start_row=3, start_column=bay_start, end_row=3, end_column=bay_end)
                for mc in range(bay_start+1, bay_end+1):
                    ws_matrix.cell(row=3, column=mc).border = _thin_border()
            # Fill row 2 STS placeholder for this bay (will merge later)
            for mc in range(bay_start, bay_end+1):
                ws_matrix.cell(row=2, column=mc).border = _thin_border()
        sts_end = col_offset - 1
        stsc = ws_matrix.cell(row=2, column=sts_start, value=sts)
        stsc.font = _font(bold=True, color=C_WHITE)
        stsc.fill = _fill(C_DARK_BLUE)
        stsc.alignment = _align()
        stsc.border = _thin_border()
        if sts_start < sts_end:
            ws_matrix.merge_cells(start_row=2, start_column=sts_start, end_row=2, end_column=sts_end)

    # Data rows
    for r_idx, hour in enumerate(hour_list):
        excel_row = 5 + r_idx
        fill_color = C_ALT_ROW if (r_idx % 2 == 0) else C_WHITE
        # Hour cell
        hc = ws_matrix.cell(row=excel_row, column=1, value=hour)
        hc.font = _font(bold=True, color=C_DARK_BLUE)
        hc.fill = _fill(C_PALE_BLUE)
        hc.alignment = _align()
        hc.border = _thin_border()
        # Data cells
        row_total = 0
        for c_idx, col_key in enumerate(matrix_cols, 2):
            val = matrix_data[hour].get(col_key, 0)
            dc = ws_matrix.cell(row=excel_row, column=c_idx)
            if val == 0:
                dc.value = '—'
                dc.font = _font(color=C_GREY_FONT)
            else:
                dc.value = val
                dc.font = _font(color="FF000000")
                row_total += val
            dc.fill = _fill(fill_color)
            dc.alignment = _align()
            dc.border = _thin_border()
        # Row total
        rtc = ws_matrix.cell(row=excel_row, column=total_col, value=row_total)
        rtc.font = _font(bold=True, color="FF000000")
        rtc.fill = _fill(C_YELLOW)
        rtc.alignment = _align()
        rtc.border = _thin_border()

    # Column total row
    total_row = 5 + len(hour_list)
    trc = ws_matrix.cell(row=total_row, column=1, value='TOTAL')
    trc.font = _font(bold=True, color=C_WHITE)
    trc.fill = _fill(C_GREEN)
    trc.alignment = _align()
    trc.border = _thin_border()

    grand_total = 0
    for c_idx, col_key in enumerate(matrix_cols, 2):
        col_sum = sum(matrix_data[h].get(col_key, 0) for h in hour_list)
        tc2 = ws_matrix.cell(row=total_row, column=c_idx, value=col_sum)
        tc2.font = _font(bold=True, color="FF000000")
        tc2.fill = _fill(C_YELLOW)
        tc2.alignment = _align()
        tc2.border = _thin_border()
        grand_total += col_sum

    gtc = ws_matrix.cell(row=total_row, column=total_col, value=grand_total)
    gtc.font = _font(bold=True, color=C_WHITE)
    gtc.fill = _fill(C_GREEN)
    gtc.alignment = _align()
    gtc.border = _thin_border()

    # Column widths
    ws_matrix.column_dimensions['A'].width = 12
    for c in range(2, total_matrix_cols + 1):
        ws_matrix.column_dimensions[get_column_letter(c)].width = 8

    # ============================================================
    # SHEET: DETAIL (formatted like sample)
    # Each STS = one column group side by side
    # Within each STS, bays are stacked vertically
    # ============================================================
    ws_detail = wb.create_sheet('DETAIL')

    # Compute column start for each STS (gap of 1 col between STS groups)
    sts_col_start = {}
    current_col = 1
    for sts in sts_list:
        sts_col_start[sts] = current_col
        current_col += sts_table_widths[sts] + 1  # +1 for gap

    total_detail_cols = current_col - 2  # last occupied column

    # Row 1: Title spanning all columns
    title_d = ws_detail.cell(row=1, column=1, value='TỔNG HỢP CHI TIẾT  ▸  STS / BAY / MOVE HOUR / WC / BLOCK')
    title_d.font = _font(bold=True, color=C_DARK_BLUE, size=11)
    title_d.fill = _fill(C_HEADER_BG)
    title_d.alignment = _align()
    title_d.border = _thin_border()
    ws_detail.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_detail_cols)

    def write_detail_sts(ws, sts, start_col, table_width, groups, row_start):
        """Write one STS block starting at row_start, starting at start_col.
        Returns the next available row after writing all bays."""
        end_col = start_col + table_width - 1
        current_row = row_start

        # STS header
        sts_cell = ws.cell(row=current_row, column=start_col, value=sts)
        sts_cell.font = _font(bold=True, color=C_WHITE)
        sts_cell.fill = _fill(C_DARK_BLUE)
        sts_cell.alignment = _align()
        sts_cell.border = _thin_border()
        ws.merge_cells(start_row=current_row, start_column=start_col, end_row=current_row, end_column=end_col)
        for mc in range(start_col+1, end_col+1):
            ws.cell(row=current_row, column=mc).border = _thin_border()
        current_row += 1

        for g_idx, group in enumerate(groups):
            bay = group['bay']
            blks = group['blocks']
            tbl_rows = group['rows']
            col_totals = group['col_totals']
            num_data_cols = 2 + len(blks) + 1  # hour + wc + blocks + total

            # BAY header row
            bay_cell = ws.cell(row=current_row, column=start_col, value=bay)
            bay_cell.font = _font(bold=True, color=C_WHITE)
            bay_cell.fill = _fill(C_MID_BLUE)
            bay_cell.alignment = _align()
            bay_cell.border = _thin_border()
            ws.merge_cells(start_row=current_row, start_column=start_col, end_row=current_row, end_column=end_col)
            for mc in range(start_col+1, end_col+1):
                ws.cell(row=current_row, column=mc).border = _thin_border()
            current_row += 1

            # Column headers: MOVE HOUR, WC, blocks..., TOTAL
            col = start_col
            for hdr in ['MOVE HOUR', 'WC'] + blks:
                hc = ws.cell(row=current_row, column=col, value=hdr)
                hc.font = _font(bold=True, color=C_WHITE)
                hc.fill = _fill(C_DARK_BLUE)
                hc.alignment = _align()
                hc.border = _thin_border()
                col += 1
            # Pad remaining cols up to end_col
            while col <= end_col - 1:
                pc = ws.cell(row=current_row, column=col)
                pc.font = _font(bold=True, color=C_WHITE)
                pc.fill = _fill(C_DARK_BLUE)
                pc.alignment = _align()
                pc.border = _thin_border()
                col += 1
            tc_hdr = ws.cell(row=current_row, column=end_col, value='TOTAL')
            tc_hdr.font = _font(bold=True, color=C_WHITE)
            tc_hdr.fill = _fill(C_GREEN)
            tc_hdr.alignment = _align()
            tc_hdr.border = _thin_border()
            current_row += 1

            # Data rows - group by MOVE HOUR (first col), alternate color per hour group
            hour_color_map = {}
            color_toggle = True
            for rd in tbl_rows:
                h_val = rd[0]
                if h_val not in hour_color_map:
                    hour_color_map[h_val] = C_ALT_ROW if color_toggle else C_WHITE
                    color_toggle = not color_toggle

            prev_hour = None
            for rd in tbl_rows:
                h_val = rd[0]
                wc_val = rd[1]
                qty_vals = rd[2:-1]
                total_val = rd[-1]
                row_fill = hour_color_map[h_val]

                col = start_col
                # MOVE HOUR cell (only show on first WC of same hour)
                if h_val != prev_hour:
                    hc2 = ws.cell(row=current_row, column=col, value=h_val)
                    hc2.font = _font(bold=True, color=C_DARK_BLUE)
                    hc2.fill = _fill(C_PALE_BLUE)
                    hc2.alignment = _align()
                    hc2.border = _thin_border()
                else:
                    ec = ws.cell(row=current_row, column=col)
                    ec.fill = _fill(C_PALE_BLUE)
                    ec.border = _thin_border()
                    ec.alignment = _align()
                col += 1
                prev_hour = h_val

                # WC cell
                wcc = ws.cell(row=current_row, column=col, value=wc_val)
                wcc.font = _font(bold=True, color=C_ORANGE_FONT)
                wcc.fill = _fill(C_ORANGE_FILL)
                wcc.alignment = _align()
                wcc.border = _thin_border()
                col += 1

                # Quantity cells
                for q in qty_vals:
                    qc = ws.cell(row=current_row, column=col)
                    if q == 0:
                        qc.value = '—'
                        qc.font = _font(color=C_GREY_FONT)
                    else:
                        qc.value = q
                        qc.font = _font(color="FF000000")
                    qc.fill = _fill(row_fill)
                    qc.alignment = _align()
                    qc.border = _thin_border()
                    col += 1

                # Pad to end_col - 1
                while col <= end_col - 1:
                    pc2 = ws.cell(row=current_row, column=col)
                    pc2.fill = _fill(row_fill)
                    pc2.border = _thin_border()
                    pc2.alignment = _align()
                    col += 1

                # Total cell
                totc = ws.cell(row=current_row, column=end_col, value=total_val)
                totc.font = _font(bold=True, color="FF000000")
                totc.fill = _fill(C_YELLOW)
                totc.alignment = _align()
                totc.border = _thin_border()
                current_row += 1

            # TOTAL row for this bay
            col = start_col
            tr_cell = ws.cell(row=current_row, column=col, value='TOTAL')
            tr_cell.font = _font(bold=True, color=C_WHITE)
            tr_cell.fill = _fill(C_GREEN)
            tr_cell.alignment = _align()
            tr_cell.border = _thin_border()
            col += 1

            # WC total cell (skip)
            wc_total = ws.cell(row=current_row, column=col)
            wc_total.fill = _fill(C_PALE_BLUE)
            wc_total.border = _thin_border()
            col += 1

            # Block totals
            blk_totals = col_totals[2:-1]
            for bt in blk_totals:
                btc = ws.cell(row=current_row, column=col, value=bt)
                btc.font = _font(bold=True, color="FF000000")
                btc.fill = _fill(C_YELLOW)
                btc.alignment = _align()
                btc.border = _thin_border()
                col += 1

            # Pad remaining
            while col <= end_col - 1:
                pc3 = ws.cell(row=current_row, column=col)
                pc3.fill = _fill(C_YELLOW)
                pc3.border = _thin_border()
                pc3.alignment = _align()
                col += 1

            # Grand total for bay
            gt_bay = col_totals[-1]
            gt_c = ws.cell(row=current_row, column=end_col, value=gt_bay)
            gt_c.font = _font(bold=True, color=C_WHITE)
            gt_c.fill = _fill(C_GREEN)
            gt_c.alignment = _align()
            gt_c.border = _thin_border()
            current_row += 1

        return current_row

    # Write each STS block
    # All STS groups start at row 2 (STS header) and grow downward
    # But they are side by side - so we track row per STS independently
    sts_next_rows = {sts: 2 for sts in sts_list}
    sts_max_row = 2

    for sts in sts_list:
        next_row = write_detail_sts(
            ws_detail, sts,
            sts_col_start[sts],
            sts_table_widths[sts],
            detail_groups_by_sts[sts],
            sts_next_rows[sts]
        )
        sts_next_rows[sts] = next_row
        sts_max_row = max(sts_max_row, next_row)

    # Column widths for DETAIL
    for c in range(1, total_detail_cols + 2):
        ws_detail.column_dimensions[get_column_letter(c)].width = 8

    # ============================================================
    # 8. Lưu file vào bộ nhớ đệm (BytesIO) thay vì ghi đĩa
    # ============================================================
    # MODIFIED: ghi workbook vào BytesIO thay vì file
    output_buffer = BytesIO()
    wb.save(output_buffer)
    output_buffer.seek(0)

    # Tính các thông số thống kê để trả về
    total_rows = len(df_result)
    objective_value = pulp.value(prob.objective)   # giá trị hàm mục tiêu (tổng clash)

    # ADDED: trả về buffer và các thông số
    return output_buffer, total_rows, objective_value

# ============================================================
# Giữ khả năng chạy độc lập (khi file được thực thi trực tiếp)
# ============================================================
if __name__ == "__main__":
    # Ví dụ: chạy thử với file TEST2.xlsx (nếu có)
    try:
        buf, rows, obj = run_optimization('TEST2.xlsx')
        with open('optimized_allocation.xlsx', 'wb') as f:
            f.write(buf.read())
        print(f"Đã tạo file optimized_allocation.xlsx")
        print(f"Số dòng phân bổ: {rows}")
        print(f"Giá trị mục tiêu (tổng clash): {obj}")
    except Exception as e:
        print(f"Lỗi khi chạy thử: {e}")