import io
import pandas as pd
import numpy as np
import pulp
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side

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
# PUBLIC API — called by WEBAPP.PY
# ============================================================
def run_optimization(file_input):
    """
    Chạy thuật toán phân bổ tối ưu.

    Parameters
    ----------
    file_input : str | BytesIO
        Đường dẫn file Excel hoặc BytesIO object (từ st.file_uploader).

    Returns
    -------
    excel_buffer : io.BytesIO
        File Excel kết quả sẵn sàng để download.
    total_rows : int
        Tổng số dòng phân bổ trong sheet RESULT.
    objective_value : int
        Số lượng clash (0 = không có clash).
    """
    # ============================================================
    # 1. Read and parse original data
    # ============================================================
    xls = pd.ExcelFile(file_input)

    # --- Sheet 1: MOVEHOUR-WEIGHTCLASS → demand per (hour, STS, bay, (wc,st,pod)) ---
    df1 = pd.read_excel(xls, sheet_name='MOVEHOUR-WEIGHTCLASS', header=None)

    has_st_pod = (str(df1.iloc[1, 2]).strip().upper() == 'ST')
    data_col_start = 4 if has_st_pod else 2

    sts_bay_map = {}
    for col in range(data_col_start, df1.shape[1]):
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
        weight = int(float(str(weight)))
        st_val  = str(row[2]).strip() if has_st_pod and pd.notna(row[2]) else ''
        pod_val = str(row[3]).strip() if has_st_pod and pd.notna(row[3]) else ''

        for col in range(data_col_start, df1.shape[1]):
            qty = row[col]
            if pd.notna(qty) and qty != '':
                qty = int(float(str(qty)))
                if qty > 0:
                    sts, bay = sts_bay_map[col]
                    key = (hour, sts, bay)
                    if key not in demands:
                        demands[key] = {}
                    dkey = (weight, st_val, pod_val)
                    demands[key][dkey] = demands[key].get(dkey, 0) + qty

    job_keys = list(demands.keys())

    all_hours_sorted = sorted(set(h for (h, s, b) in job_keys))
    hour_rank = {h: i for i, h in enumerate(all_hours_sorted)}

    jobs_by_hour = {}
    for (h, s, b) in job_keys:
        jobs_by_hour.setdefault(h, []).append((s, b))

    # --- Sheet 2: BLOCK-WEIGHT CLASS → supply per (block, st, pod, wc) ---
    df2 = pd.read_excel(xls, sheet_name='BLOCK-WEIGHT CLASS', header=0)
    col_names = [str(c).strip() for c in df2.columns]
    has_st_pod_supply = (col_names[1].upper() == 'ST' and col_names[2].upper() == 'POD')
    wc_col_start = 3 if has_st_pod_supply else 1

    supply = {}
    blocks_set = set()
    for idx, row in df2.iterrows():
        block = str(row.iloc[0]).strip()
        if block in ('nan', 'GRAND TOTAL', '') or not block:
            continue
        st_v  = str(row.iloc[1]).strip() if has_st_pod_supply else ''
        pod_v = str(row.iloc[2]).strip() if has_st_pod_supply else ''
        skey  = (block, st_v, pod_v)
        wc_dict = {}
        for wi, w in enumerate([1, 2, 3, 4, 5]):
            col_idx = wc_col_start + wi
            if col_idx < len(row):
                val = row.iloc[col_idx]
                wc_dict[w] = int(val) if pd.notna(val) and val != '' else 0
            else:
                wc_dict[w] = 0
        supply[skey] = wc_dict
        blocks_set.add(block)

    weight_classes = [1, 2, 3, 4, 5]
    blocks = sorted(blocks_set)
    supply_keys = [k for k in supply if any(supply[k][w] > 0 for w in weight_classes)]

    # --- Sheet 3: DATA (container-level) ---
    container_data_available = False
    try:
        df_containers = pd.read_excel(xls, sheet_name='DATA', header=0)
        # ... (giữ nguyên toàn bộ phần parse DATA như cũ) ...
        # (để ngắn gọn, tôi giữ nguyên code gốc của phần này - không thay đổi)
        # ... (code parse DATA dài, giữ nguyên như file gốc) ...
        # (bạn copy nguyên phần này từ file cũ của bạn)
        container_data_available = True
    except Exception:
        container_data_available = False

    # Build stacking structures (giữ nguyên)
    yb_wc_supply = {}
    stack_ordering = {}
    blocking_pairs = []
    if container_data_available:
        # ... (giữ nguyên toàn bộ phần build stacking như file gốc) ...
        pass

    # ============================================================
    # 2. Check total demand vs supply
    # ============================================================
    # ... (giữ nguyên toàn bộ phần check demand-supply như cũ) ...

    # ============================================================
    # 3. Build and solve the optimisation model
    # ============================================================
    prob = pulp.LpProblem("Minimize_Clashes_ST_POD", pulp.LpMinimize)

    # y, x, u, e variables (giữ nguyên)
    y_vars = {}
    for (h, s, bay) in job_keys:
        for b in blocks:
            y_vars[(h, s, bay, b)] = pulp.LpVariable(f"y_{h}_{s}_{bay}_{b}", cat='Binary')

    x_vars = {}
    for (h, s, bay) in job_keys:
        for dkey in demands[(h, s, bay)]:
            w, st_v, pod_v = dkey
            for skey in supply_keys:
                b, sup_st, sup_pod = skey
                if sup_st != st_v or sup_pod != pod_v:
                    continue
                vname = f"x_{h}_{s}_{bay}_{b}_{w}_{st_v}_{pod_v}"
                x_vars[(h, s, bay, b, dkey)] = pulp.LpVariable(vname, lowBound=0, cat='Integer')

    u_vars = {}
    e_vars = {}
    for h in jobs_by_hour:
        for b in blocks:
            u_vars[(h, b)] = pulp.LpVariable(f"u_{h}_{b}", lowBound=0, cat='Integer')
            e_vars[(h, b)] = pulp.LpVariable(f"e_{h}_{b}", lowBound=0, cat='Integer')
            prob += u_vars[(h, b)] == pulp.lpSum(y_vars[(h, s, bay, b)] for (s, bay) in jobs_by_hour[h])
            prob += e_vars[(h, b)] >= u_vars[(h, b)] - 1

    # single_block & block_bay (giữ nguyên)
    single_block = {}
    for (h, s, bay) in job_keys:
        single_block[(h, s, bay)] = pulp.LpVariable(f"sb_{h}_{s}_{bay}", lowBound=0, upBound=1, cat='Continuous')
        prob += single_block[(h, s, bay)] >= (2 - pulp.lpSum(y_vars[(h, s, bay, b)] for b in blocks))

    all_bays = sorted(set(bay for (_, _, bay) in job_keys))
    block_bay = {}
    for b in blocks:
        for bay in all_bays:
            var = pulp.LpVariable(f"bb_{b}_{bay}", cat='Binary')
            block_bay[(b, bay)] = var
            for (h, s, bj) in job_keys:
                if bj == bay:
                    prob += var >= y_vars[(h, s, bay, b)]

    # block_bay_wc (CẢI TIẾN 1)
    block_bay_wc = {}
    for b in blocks:
        for bay in all_bays:
            for wc in weight_classes:
                var = pulp.LpVariable(f"bbw_{b}_{bay}_{wc}", cat='Binary')
                block_bay_wc[(b, bay, wc)] = var
                for (h, s, bj) in job_keys:
                    if bj == bay:
                        for dkey in demands[(h, s, bay)]:
                            w, st_v, pod_v = dkey
                            if w == wc:
                                key_x = (h, s, bay, b, dkey)
                                if key_x in x_vars:
                                    prob += var >= x_vars[key_x] / (demands[(h, s, bay)][dkey] + 0.1)

    # ========== CẢI TIẾN 2: Khuyến khích mỗi bay có ít nhất 2 block ==========
    bay_single = {}
    for bay in all_bays:
        var = pulp.LpVariable(f"bs_{bay}", lowBound=0, upBound=1, cat='Continuous')
        bay_single[bay] = var
        total_blocks_bay = pulp.lpSum(block_bay[(b, bay)] for b in blocks)
        prob += var >= (2 - total_blocks_bay)

    # ========== CẢI TIẾN 3: PHÂN BỔ ĐỒNG ĐỀU cho vessel bay lớn ==========
    # Mục tiêu: tránh 1 block dồn quá nhiều container vào 1 BAY lớn (ví dụ 95 conts BAY 50 → B04)
    MAXLOAD_W = 0.01   # Có thể chỉnh: 0.02 (mạnh hơn), 0.005 (nhẹ hơn)

    vessel_bay_keys = list(set((s, bay) for (h, s, bay) in job_keys))

    load_vars = {}
    for vb in vessel_bay_keys:
        s, bay = vb
        for b in blocks:
            relevant_x = []
            for h in [hh for (hh, ss, bb) in job_keys if ss == s and bb == bay]:
                for dkey in demands.get((h, s, bay), {}):
                    xkey = (h, s, bay, b, dkey)
                    if xkey in x_vars:
                        relevant_x.append(x_vars[xkey])
            if relevant_x:
                load_vars[(s, bay, b)] = pulp.LpVariable(
                    f"load_{s}_{bay}_{b}", lowBound=0, cat='Integer'
                )
                prob += load_vars[(s, bay, b)] == pulp.lpSum(relevant_x)

    max_load_vars = {}
    for vb in vessel_bay_keys:
        s, bay = vb
        maxl = pulp.LpVariable(f"maxload_{s}_{bay}", lowBound=0, cat='Continuous')
        max_load_vars[vb] = maxl
        for b in blocks:
            if (s, bay, b) in load_vars:
                prob += maxl >= load_vars[(s, bay, b)]

    # --- Hàm mục tiêu (đã thêm CẢI TIẾN 3) ---
    CLASH_W  = 1.0
    SINGLE_W = 0.3
    SPREAD_W = 0.05
    BLOCK_BAY_WC_W = 0.2
    BAY_SINGLE_W = 0.2

    clash_term        = pulp.lpSum(e_vars.values())
    single_term       = pulp.lpSum(single_block.values())
    spread_term       = pulp.lpSum(block_bay.values())
    block_bay_wc_term = pulp.lpSum(block_bay_wc.values())
    bay_single_term   = pulp.lpSum(bay_single.values())
    maxload_term      = pulp.lpSum(max_load_vars.values())   # ← MỚI

    prob += (
        CLASH_W * clash_term +
        SINGLE_W * single_term +
        SPREAD_W * spread_term +
        BLOCK_BAY_WC_W * block_bay_wc_term +
        BAY_SINGLE_W * bay_single_term +
        MAXLOAD_W * maxload_term
    )

    # ============================================================
    # Core constraints (giữ nguyên)
    # ============================================================
    # C1. Demand satisfaction
    for (h, s, bay) in job_keys:
        for dkey, d in demands[(h, s, bay)].items():
            w, st_v, pod_v = dkey
            x_sum = pulp.lpSum(
                x_vars[(h, s, bay, b, dkey)]
                for skey in supply_keys
                for b in [skey[0]]
                if skey[1] == st_v and skey[2] == pod_v
                and (h, s, bay, b, dkey) in x_vars
            )
            prob += x_sum == d

    # C2. Supply cap
    for skey in supply_keys:
        b, st_v, pod_v = skey
        for w in weight_classes:
            dkey_w = [(h, s, bay, (w, st_v, pod_v))
                      for (h, s, bay) in job_keys
                      if (w, st_v, pod_v) in demands.get((h, s, bay), {})]
            if not dkey_w:
                continue
            prob += pulp.lpSum(
                x_vars[(h, s, bay, b, (w, st_v, pod_v))]
                for (h, s, bay, dk) in dkey_w
                if (h, s, bay, b, dk) in x_vars
            ) <= supply[skey][w]

    # C3. Linking x → y
    for (h, s, bay) in job_keys:
        for dkey, d in demands[(h, s, bay)].items():
            for skey in supply_keys:
                b = skey[0]
                if (h, s, bay, b, dkey) in x_vars:
                    prob += x_vars[(h, s, bay, b, dkey)] <= d * y_vars[(h, s, bay, b)]

    # ============================================================
    # Solve
    # ============================================================
    solver = pulp.PULP_CBC_CMD(msg=True, timeLimit=300)
    prob.solve(solver)

    status = prob.status
    print(f"Status: {pulp.LpStatus[status]}")

    # ============================================================
    # 4~8. Extract result + Write Excel (giữ nguyên toàn bộ phần sau)
    # ============================================================
    # (Phần từ "4. Extract result" đến cuối hàm không thay đổi)
    # Bạn copy nguyên phần còn lại từ file code.txt cũ của bạn (từ dòng "# ------------------------------------------------------------------" trở xuống)

    # ... (toàn bộ code từ phần 4a. Aggregate result đến cuối hàm return excel_buffer, total_rows, objective_value) ...

    # Để code đầy đủ, bạn chỉ cần thay phần 3 (model) bằng đoạn tôi đưa ở trên.
    # Phần sau model (extract result, write excel) giữ nguyên 100%.

    excel_buffer = io.BytesIO()
    wb.save(excel_buffer)
    excel_buffer.seek(0)
    total_rows = len(df_result_detail)
    objective_value = int(pulp.value(prob.objective) or 0)
    print(f"Done. Rows={total_rows}, Clashes={objective_value}")
    return excel_buffer, total_rows, objective_value
