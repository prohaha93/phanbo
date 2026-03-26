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
    # New format: col 0=MOVE HOUR, col 1=WC, col 2=ST, col 3=POD, col 4+=STS×BAY qty
    df1 = pd.read_excel(xls, sheet_name='MOVEHOUR-WEIGHTCLASS', header=None)

    # Detect if ST/POD columns exist (new format has 4 fixed cols, old has 2)
    # Row 1 (index 1): col 2 = 'ST' or BAY label?
    has_st_pod = (str(df1.iloc[1, 2]).strip().upper() == 'ST')
    data_col_start = 4 if has_st_pod else 2   # data starts at col 4 (new) or col 2 (old)

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
        # Read ST and POD if present (new format)
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
                    # demand key is (wc, st, pod) — backward compat: if no ST/POD, use (wc,'','')
                    dkey = (weight, st_val, pod_val)
                    demands[key][dkey] = demands[key].get(dkey, 0) + qty

    print(f"Demand format: {'WC+ST+POD' if has_st_pod else 'WC only (legacy)'}")

    job_keys = list(demands.keys())

    # Build sorted hour list for ordering constraints
    all_hours_sorted = sorted(set(h for (h, s, b) in job_keys))
    hour_rank = {h: i for i, h in enumerate(all_hours_sorted)}

    jobs_by_hour = {}
    for (h, s, b) in job_keys:
        jobs_by_hour.setdefault(h, []).append((s, b))

    # --- Sheet 2: BLOCK-WEIGHT CLASS → supply per (block, st, pod, wc) ---
    # New format: BLOCK/WEIGHT CLASS | ST | POD | 1 | 2 | 3 | 4 | 5 | TOTALS
    # Old format: BLOCK/WEIGHT CLASS | 1 | 2 | 3 | 4 | 5 | TOTALS
    df2 = pd.read_excel(xls, sheet_name='BLOCK-WEIGHT CLASS', header=0)
    col_names = [str(c).strip() for c in df2.columns]

    # Detect new format: col 1 = 'ST', col 2 = 'POD'
    has_st_pod_supply = (col_names[1].upper() == 'ST' and col_names[2].upper() == 'POD')
    wc_col_start = 3 if has_st_pod_supply else 1  # WC columns start at idx 3 (new) or 1 (old)

    # supply: (block, st, pod) → {wc: qty}
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
    # supply_keys: all (block, st, pod) tuples with non-zero supply
    supply_keys = [k for k in supply if any(supply[k][w] > 0 for w in weight_classes)]
    print(f"Supply format: {'BLOCK+ST+POD' if has_st_pod_supply else 'BLOCK only (legacy)'}")
    print(f"Supply keys: {len(supply_keys)} (block×ST×POD combinations)")

    # --- Sheet 3 (DATA file): container-level layout ---
    # Column mapping (TEST2.xlsx):
    #   A=YARD, B=YC(real WC), C=YP(yard position), D=ID(real cont ID),
    #   E=ST(size type), F=POD(port of discharge)
    #   L=ST_PROJ, M=POD_PROJ (projection — same values in current data)
    #   O=MOVE HOUR, P=BAY, Q=YB, R=YR, S=YT
    container_data_available = False
    try:
        df_containers = pd.read_excel(xls, sheet_name='DATA', header=0)
        cols = list(df_containers.columns)
        # Map column names — support both new (YC/YP/ID/ST/POD) and old (Unnamed:x) formats
        def find_col(candidates):
            for c in candidates:
                if c in cols: return c
            return None
        wc_src   = find_col(['YC', 'Unnamed: 1'])   # col B: real WC
        yp_src   = find_col(['YP', 'Unnamed: 2'])   # col C: yard position
        id_src   = find_col(['ID', 'Unnamed: 3'])   # col D: real container ID
        st_src   = find_col(['ST'])                  # col E: size type
        pod_src  = find_col(['POD'])                 # col F: port of discharge

        required_found = (wc_src and yp_src
                          and 'YB' in cols and 'YR' in cols and 'YT' in cols)
        if required_found:
            df_containers = df_containers.dropna(
                subset=[wc_src, yp_src, 'YB', 'YR', 'YT']
            ).copy()
            df_containers['REAL_WC']      = df_containers[wc_src].astype(float).astype(int)
            df_containers['YARD_POS']     = df_containers[yp_src].astype(str).str.strip()
            df_containers['REAL_CONT_ID'] = (df_containers[id_src].fillna('').astype(str).str.strip()
                                             if id_src else '')
            df_containers['CONT_ST']      = (df_containers[st_src].fillna('').astype(str).str.strip()
                                             if st_src else '')
            df_containers['CONT_POD']     = (df_containers[pod_src].fillna('').astype(str).str.strip()
                                             if pod_src else '')
            df_containers['YARD'] = df_containers['YARD'].astype(str).str.strip()
            df_containers['YB']   = df_containers['YB'].astype(float).astype(int)
            df_containers['YR']   = df_containers['YR'].astype(float).astype(int)
            df_containers['YT']   = df_containers['YT'].astype(float).astype(int)
            container_data_available = True
            print("Container-level DATA sheet found – stacking rules will be applied.")
            print(f"  {len(df_containers)} containers loaded.")
            print(f"  ST values : {sorted(df_containers['CONT_ST'].unique().tolist())}")
            print(f"  POD values: {sorted(df_containers['CONT_POD'].unique().tolist())}")
        else:
            print(f"DATA sheet missing required columns (need WC col + YP/YB/YR/YT). Skipped.")
    except Exception as e:
        print(f"No DATA sheet found – stacking rules skipped. ({e})")
    # 1b. Build container-level stacking structures (if DATA available)
    # ============================================================
    # Physical constraint: within (YARD=block, YB=yard_bay, YR=row),
    #   container at tier T cannot be picked until ALL containers at tier T+1, T+2, ... are picked.
    # Rule priority:
    #   P1. Prefer to exhaust one YB fully before starting another YB of same YARD.
    #   P2. Within a mixed-WC stack (YB+YR), must pick highest tier first (physically forced).
    #   P3. When WC1 sits atop WC2 in same stack: WC1 goes to current/earlier MH,
    #       WC2 can only go to equal-or-later MH than the LAST WC1 in that stack.

    # Data structures built:
    #   yb_wc_supply[block][yb][wc]     = count of containers of that WC in that YB
    #   stack_ordering[block][yb][yr]   = list of (tier, wc) sorted HIGH→LOW tier
    #   yb_order[block]                 = list of YBs sorted by earliest-accessible WC priority
    #   blocking_pairs                  = list of (block, yb, yr, wc_above, wc_below, count_above)
    #     meaning: must pick count_above units of wc_above from this stack before picking wc_below

    yb_wc_supply   = {}   # block → yb → wc → count
    stack_ordering = {}   # block → yb → yr → [(tier, wc), ...] high→low
    blocking_pairs = []   # (block, yb, yr, wc_top, count_top, wc_bottom, count_bottom)

    if container_data_available:
        # Use REAL_WC (col B) — already parsed above
        df_c = df_containers[['YARD','YB','YR','YT','REAL_WC','YARD_POS','REAL_CONT_ID','CONT_ST','CONT_POD']].copy()

        for block in blocks:
            block_df = df_c[df_c['YARD'] == block]
            if block_df.empty:
                continue
            yb_wc_supply[block] = {}
            stack_ordering[block] = {}

            for yb, yb_df in block_df.groupby('YB'):
                yb_wc_supply[block][yb] = {}
                stack_ordering[block][yb] = {}

                # Count REAL_WC per YB
                for wc, cnt in yb_df.groupby('REAL_WC').size().items():
                    yb_wc_supply[block][yb][wc] = int(cnt)

                # Build stack per row, sorted highest tier first
                for yr, yr_df in yb_df.groupby('YR'):
                    ordered = yr_df.sort_values('YT', ascending=False)[['YT','REAL_WC']].values.tolist()
                    stack_ordering[block][yb][yr] = [(int(t), int(w)) for t, w in ordered]

                # Find blocking pairs: where higher-WC tiers sit ABOVE lower-WC tiers in same row
                # (physically: higher tier number = physically on top = must move first)
                for yr, tiers in stack_ordering[block][yb].items():
                    # tiers is sorted high→low (must pick in this order)
                    # Scan: any tier with WC_a above a tier with WC_b where WC_a != WC_b
                    # Count how many containers in this stack sit ABOVE each WC
                    wcs_above = []
                    for tier, wc in tiers:
                        if wcs_above:
                            # All containers in wcs_above must be picked before this one
                            for (prev_wc, prev_tier) in wcs_above:
                                if prev_wc != wc:
                                    # Record: prev_wc at prev_tier blocks wc at this tier in same (block,yb,yr)
                                    blocking_pairs.append((block, yb, yr, prev_tier, prev_wc, tier, wc))
                        wcs_above.append((wc, tier))

        print(f"Stacking structures built: {len(blocking_pairs)} cross-WC blocking pairs found.")

    # ============================================================
    # 2. Check total demand vs supply per (wc, st, pod)
    # ============================================================
    total_demand = {}  # (wc, st, pod) → qty
    for job in job_keys:
        for dkey, qty in demands[job].items():
            total_demand[dkey] = total_demand.get(dkey, 0) + qty

    total_supply = {}  # (wc, st, pod) → qty
    for skey in supply_keys:
        block, st_v, pod_v = skey
        for w in weight_classes:
            tkey = (w, st_v, pod_v)
            total_supply[tkey] = total_supply.get(tkey, 0) + supply[skey][w]

    print("Total demand per (WC, ST, POD):")
    for k in sorted(total_demand):
        print(f"  WC={k[0]} ST={k[1]} POD={k[2]}: {total_demand[k]}")
    print("Total supply per (WC, ST, POD):")
    for k in sorted(total_supply):
        print(f"  WC={k[0]} ST={k[1]} POD={k[2]}: {total_supply[k]}")

    ok = True
    for k in set(list(total_demand.keys()) + list(total_supply.keys())):
        d = total_demand.get(k, 0)
        s = total_supply.get(k, 0)
        if d != s:
            print(f"ERROR Mismatch WC={k[0]} ST={k[1]} POD={k[2]}: demand={d}, supply={s}")
            ok = False
    if not ok:
        raise ValueError("Demand/supply mismatch — kiểm tra lại file input.")
    print("Demand/supply balanced OK.")

    # ============================================================
    # 3. Build and solve the optimisation model
    # ============================================================
    # Decision variables:
    #   y[h,s,bay,b]      = 1 if block b is used for job (h,s,bay)   [Binary]
    #   x[h,s,bay,b,w]    = qty of WC w from block b to job (h,s,bay) [Integer ≥ 0]
    #
    # NEW variables (when container data available):
    #   z[h,s,bay,b,yb]   = 1 if yard-bay yb of block b is used for job (h,s,bay) [Binary]
    #                        (drives the "exhaust one YB first" preference via objective penalty)
    #   xq[h,s,bay,b,yb,w]= qty of WC w from (block b, yard-bay yb) to job [Integer ≥ 0]
    #
    # Stacking constraint (hard):
    #   For each blocking pair (b, yb, yr, tier_top, wc_top, tier_bottom, wc_bottom):
    #   SUM_{h'≤h} xq[h',*,*,b,yb,wc_top across that row] ≥ xq[h,*,*,b,yb,wc_bottom in that row]
    #   i.e. cumulative picks of wc_top up to hour h ≥ picks of wc_bottom at hour h

    prob = pulp.LpProblem("Minimize_Clashes_ST_POD", pulp.LpMinimize)

    # y[h,s,bay,b] = 1 if block b used for job (h,s,bay) — clash counting at block level
    y_vars = {}
    for (h, s, bay) in job_keys:
        for b in blocks:
            y_vars[(h, s, bay, b)] = pulp.LpVariable(f"y_{h}_{s}_{bay}_{b}", cat='Binary')

    # x[h,s,bay,b,(w,st,pod)] = qty of (WC,ST,POD) from block b to job
    x_vars = {}
    for (h, s, bay) in job_keys:
        for dkey in demands[(h, s, bay)]:          # dkey = (wc, st, pod)
            w, st_v, pod_v = dkey
            for skey in supply_keys:
                b, sup_st, sup_pod = skey
                if sup_st != st_v or sup_pod != pod_v:
                    continue                        # ST/POD must match
                vname = f"x_{h}_{s}_{bay}_{b}_{w}_{st_v}_{pod_v}"
                x_vars[(h, s, bay, b, dkey)] = pulp.LpVariable(vname, lowBound=0, cat='Integer')

    # Clash counting: u[h,b] = #jobs at hour h that use block b; e[h,b] = max(0, u-1)
    u_vars = {}
    e_vars = {}
    for h in jobs_by_hour:
        for b in blocks:
            u_vars[(h, b)] = pulp.LpVariable(f"u_{h}_{b}", lowBound=0, cat='Integer')
            e_vars[(h, b)] = pulp.LpVariable(f"e_{h}_{b}", lowBound=0, cat='Integer')
            prob += u_vars[(h, b)] == pulp.lpSum(
                y_vars[(h, s, bay, b)] for (s, bay) in jobs_by_hour[h])
            prob += e_vars[(h, b)] >= u_vars[(h, b)] - 1

    # ============================================================
    # 3a. Objective: minimise clashes + movement penalties
    # ============================================================
    # Weight tuning (cải tiến: tăng mạnh trọng số để khuyến khích phân bổ đều)
    CLASH_W  = 100.0    # ưu tiên cao nhất: giảm clash
    SINGLE_W = 10.0     # phạt nặng nếu job chỉ dùng 1 block
    SPREAD_W = 5.0      # phạt mỗi cặp (block, bay) → hạn chế một block phục vụ nhiều bay
    BLOCK_BAY_WC_W = 2.0
    BAY_SINGLE_W = 10.0   # phạt nặng nếu bay chỉ có 1 block

    # single_block[h,s,bay] = 1 if only 1 block serves this job
    # Constraint: single_block >= 2 - sum_b(y[h,s,bay,b])
    # When sum_b y = 1 → single_block >= 1 (penalised)
    # When sum_b y >= 2 → constraint trivially satisfied → single_block = 0
    single_block = {}
    for (h, s, bay) in job_keys:
        single_block[(h, s, bay)] = pulp.LpVariable(
            f"sb_{h}_{s}_{bay}", lowBound=0, upBound=1, cat='Continuous')
        prob += single_block[(h, s, bay)] >= (
            2 - pulp.lpSum(y_vars[(h, s, bay, b)] for b in blocks)
        )

    # block_bay[b, bay] = 1 if block b serves ANY job at vessel bay 'bay'
    # Constraint: block_bay[b,bay] >= y[h,s,bay,b]  for each h,s
    all_bays = sorted(set(bay for (_, _, bay) in job_keys))
    block_bay = {}
    for b in blocks:
        for bay in all_bays:
            var = pulp.LpVariable(f"bb_{b}_{bay}", cat='Binary')
            block_bay[(b, bay)] = var
            for (h, s, bj) in job_keys:
                if bj == bay:
                    prob += var >= y_vars[(h, s, bay, b)]

    # ========== CẢI TIẾN 1: Hạn chế một block phục vụ quá nhiều (bay, weight class) ==========
    block_bay_wc = {}
    for b in blocks:
        for bay in all_bays:
            for wc in weight_classes:
                var = pulp.LpVariable(f"bbw_{b}_{bay}_{wc}", cat='Binary')
                block_bay_wc[(b, bay, wc)] = var
                # Liên kết var với x_vars: nếu có x_vars nào dùng (b, bay, wc) thì var >= 1
                for (h, s, bj) in job_keys:
                    if bj == bay:
                        for dkey in demands[(h, s, bay)]:
                            w, st_v, pod_v = dkey
                            if w == wc:
                                key_x = (h, s, bay, b, dkey)
                                if key_x in x_vars:
                                    # Dùng demand làm hệ số lớn-M: nếu x>0 thì var phải >= 1
                                    prob += var >= x_vars[key_x] / (demands[(h, s, bay)][dkey] + 0.1)

    # ========== CẢI TIẾN 2: Khuyến khích mỗi bay có ít nhất 2 block ==========
    bay_single = {}
    for bay in all_bays:
        var = pulp.LpVariable(f"bs_{bay}", lowBound=0, upBound=1, cat='Continuous')
        bay_single[bay] = var
        total_blocks_bay = pulp.lpSum(block_bay[(b, bay)] for b in blocks)
        prob += var >= (2 - total_blocks_bay)

    # ========== RÀNG BUỘC CỨNG: MỖI BAY PHẢI CÓ ÍT NHẤT 2 BLOCK ==========
    min_blocks_per_bay = 2
    for bay in all_bays:
        prob += pulp.lpSum(block_bay[(b, bay)] for b in blocks) >= min_blocks_per_bay

    # --- Hàm mục tiêu ---
    clash_term    = pulp.lpSum(e_vars.values())
    single_term   = pulp.lpSum(single_block.values())
    spread_term   = pulp.lpSum(block_bay.values())
    block_bay_wc_term = pulp.lpSum(block_bay_wc.values())
    bay_single_term   = pulp.lpSum(bay_single.values())

    prob += (CLASH_W * clash_term +
             SINGLE_W * single_term +
             SPREAD_W * spread_term +
             BLOCK_BAY_WC_W * block_bay_wc_term +
             BAY_SINGLE_W * bay_single_term)

    # ============================================================
    # 3b. Core constraints
    # ============================================================
    # C1. Demand satisfaction per (h, s, bay, wc, st, pod)
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

    # C2. Supply cap per (block, st, pod, wc)
    for skey in supply_keys:
        b, st_v, pod_v = skey
        for w in weight_classes:
            dkey_w = [(h, s, bay, (w, st_v, pod_v))
                      for (h, s, bay) in job_keys
                      if (w, st_v, pod_v) in demands[(h, s, bay)]]
            if not dkey_w:
                continue
            prob += pulp.lpSum(
                x_vars[(h, s, bay, b, (w, st_v, pod_v))]
                for (h, s, bay, dk) in dkey_w
                if (h, s, bay, b, dk) in x_vars
            ) <= supply[skey][w]

    # C3. Linking x → y: can only use block b if y[h,s,bay,b]=1
    for (h, s, bay) in job_keys:
        for dkey, d in demands[(h, s, bay)].items():
            for skey in supply_keys:
                b = skey[0]
                if (h, s, bay, b, dkey) in x_vars:
                    prob += x_vars[(h, s, bay, b, dkey)] <= d * y_vars[(h, s, bay, b)]

    # NOTE: YB concentration and physical tier-ordering enforced by greedy post-processor.

    # 3d. Solve
    # ============================================================
    solver = pulp.PULP_CBC_CMD(msg=True, timeLimit=300)
    prob.solve(solver)

    status = prob.status
    print(f"Status: {pulp.LpStatus[status]}")
    if status == pulp.LpStatusInfeasible:
        raise RuntimeError("Model infeasible — kiểm tra supply/demand và ràng buộc.")
    elif status not in (1,):
        print("No optimal solution found within time limit – using best solution found.")

    # ============================================================
    # 4. Extract result  +  map individual containers to each assignment
    # ============================================================

    # ------------------------------------------------------------------
    # 4a. Aggregate result (same as before)
    # ------------------------------------------------------------------
    result_rows = []
    for (h, s, bay, b) in y_vars:
        if pulp.value(y_vars[(h, s, bay, b)]) is not None and        pulp.value(y_vars[(h, s, bay, b)]) > 0.5:
            for dkey in demands[(h, s, bay)]:
                w, st_v, pod_v = dkey
                xkey = (h, s, bay, b, dkey)
                if xkey not in x_vars:
                    continue
                qty = pulp.value(x_vars[xkey])
                if qty is not None and qty > 0.5:
                    result_rows.append({
                        'MOVE HOUR': h, 'STS': s, 'BAY': bay,
                        'ASSIGNED BLOCK': b,
                        'WEIGHT CLASS': w, 'ST': st_v, 'POD': pod_v,
                        'QUANTITIES': int(round(qty))
                    })

    df_result = pd.DataFrame(result_rows)
    df_result.sort_values(['MOVE HOUR', 'STS', 'BAY', 'ASSIGNED BLOCK'], inplace=True)

    # ------------------------------------------------------------------
    # 4b. Map individual containers to assignments (when DATA available)
    #
    # Strategy (greedy, respects physical stacking):
    #   For each assignment (MOVE HOUR h, STS s, BAY bay, BLOCK b, WC w, QTY qty):
    #     Pick exactly `qty` containers from block b with weight class w,
    #     selecting in this priority order:
    #       1. YB with most containers of WC w first (concentrate per YB → P1 rule)
    #       2. Within YB, prefer containers at HIGHEST tier first (P2/P3 rule: top tier out first)
    #       3. Respect stacking: only pick a container if ALL containers of higher tier
    #          in the SAME (YB, YR) have already been picked in earlier/current hours
    #
    # All hours are processed in chronological order to maintain stacking state.
    # ------------------------------------------------------------------

    df_result_detail = []   # one row per container

    if container_data_available:
        # Build container pool from DATA sheet (REAL_WC = col B)
        pool = {}
        for _, row in df_containers[['YARD','YB','YR','YT','REAL_WC',
                                       'YARD_POS','REAL_CONT_ID',
                                       'CONT_ST','CONT_POD']].iterrows():
            blk = row['YARD']
            pool.setdefault(blk, []).append({
                'yb': int(row['YB']), 'yr': int(row['YR']), 'yt': int(row['YT']),
                'wc': int(row['REAL_WC']),
                'yard_pos':     row['YARD_POS'],
                'real_cont_id': row['REAL_CONT_ID'],
                'st':           row['CONT_ST'],
                'pod':          row['CONT_POD'],
                'picked': False, 'pick_h': None
            })

        # ── Sticky YB tracker ────────────────────────────────────────────────
        # opened_ybs[(block, yb)] = True  khi đã bắt đầu lấy từ YB này
        # YB đã được "mở" (partially picked) sẽ được ưu tiên cao nhất,
        # để dứt điểm lấy hết trước khi chuyển sang YB mới trong cùng block.
        opened_ybs = set()   # (block, yb) pairs that have been started

        def accessible_at(cont, containers, h_rank_val):
            for c in containers:
                if c is cont:
                    continue
                if c['yb'] == cont['yb'] and c['yr'] == cont['yr'] and c['yt'] > cont['yt']:
                    if not c['picked']:
                        return False
                    if c['pick_h'] is not None and hour_rank[c['pick_h']] > h_rank_val:
                        return False
            return True

        def pick_n(block, wc, st_match, pod_match, qty, h, s_job, bay_job, h_rank_val, result_list):
            """
            Pick qty containers of (wc, st_match) from block, with priority:
              P0. STICKY YB: YB đã được mở (partially picked) → ưu tiên tuyệt đối
                  để dứt điểm 1 YB trước khi sang YB mới (tránh di chuyển rải rác).
              P1. YB mới: chọn YB có nhiều container nhất (concentrate per YB).
              P2. Within YB: lowest YR first (row 1 → 7).
              P3. Within row: highest YT first (top → bottom, no re-handling).
            Incremental: re-evaluate after each pick.
            """
            containers = pool[block]
            remaining = qty
            def matches(c):
                if c['wc'] != wc: return False
                if st_match  and c.get('st','')  != st_match:  return False
                if pod_match and c.get('pod','') != pod_match: return False
                return True
            while remaining > 0:
                cands = [c for c in containers
                         if not c['picked'] and matches(c)
                         and accessible_at(c, containers, h_rank_val)]
                if not cands:
                    break
                # Count accessible matching containers per YB
                yb_cnt = {}
                for c in cands:
                    yb_cnt[c['yb']] = yb_cnt.get(c['yb'], 0) + 1
                # Sort key:
                #   P0: sticky=0 (already opened) beats sticky=1 (fresh YB)
                #   P1: most containers in YB (descending)
                #   P2: lower YB id (tie-break, consistent ordering)
                #   P3: lower YR (row 1 → 7)
                #   P4: highest tier first
                cands.sort(key=lambda c: (
                    0 if (block, c['yb']) in opened_ybs else 1,  # P0: sticky first
                    -yb_cnt[c['yb']],  # P1: most-loaded YB first
                    c['yb'],           # P2: tie-break YB id
                    c['yr'],           # P3: row 1 → 7
                    -c['yt']           # P4: highest tier first
                ))
                best = cands[0]
                best['picked'] = True
                best['pick_h'] = h
                opened_ybs.add((block, best['yb']))   # mark YB as opened
                result_list.append({
                    'MOVE HOUR':      h,
                    'CONTAINER ID':   best['real_cont_id'],
                    'ST':             best.get('st', st_match),   # actual ST from container
                    'POD':            best.get('pod', pod_match),  # actual POD from container
                    'STS': s_job,     'BAY': bay_job,
                    'ASSIGNED BLOCK': block,
                    'WEIGHT CLASS':   wc,
                    'QUANTITIES':     qty,
                    'YB': best['yb'], 'YR': best['yr'], 'YT': best['yt'],
                    'YARD POSITION':  best['yard_pos']
                })
                remaining -= 1
            return remaining

        df_result_sorted = df_result.copy()
        df_result_sorted['_hr'] = df_result_sorted['MOVE HOUR'].map(hour_rank)
        df_result_sorted.sort_values(['_hr','STS','BAY','ASSIGNED BLOCK','WEIGHT CLASS'],
                                      inplace=True)

        deferred = []

        for h in all_hours_sorted:
            h_rank_val = hour_rank[h]
            hour_asgns = df_result_sorted[df_result_sorted['MOVE HOUR'] == h]

            for _, asg in hour_asgns.iterrows():
                s, bay_job, b = asg['STS'], asg['BAY'], asg['ASSIGNED BLOCK']
                w    = int(asg['WEIGHT CLASS'])
                st_v = str(asg.get('ST', '')).strip()
                pod_v= str(asg.get('POD', '')).strip()
                qty  = int(asg['QUANTITIES'])
                if b not in pool:
                    df_result_detail.append({
                        'MOVE HOUR': h, 'STS': s, 'BAY': bay_job,
                        'ASSIGNED BLOCK': b, 'WEIGHT CLASS': w,
                        'CONTAINER ID': '', 'ST': '', 'POD': '', 'QUANTITIES': qty, 'YB': '', 'YR': '', 'YT': '', 'YARD POSITION': ''
                    })
                    continue
                rem = pick_n(b, w, st_v, pod_v, qty, h, s, bay_job, h_rank_val, df_result_detail)
                if rem > 0:
                    deferred.append({'b': b, 'wc': w, 'st': st_v, 'pod': pod_v,
                                      'qty': rem, 'h_orig': h, 's': s, 'bay': bay_job,
                                      'h_rank_min': h_rank_val})

            still_deferred = []
            for d in deferred:
                rem = pick_n(d['b'], d['wc'], d.get('st',''), d.get('pod',''),
                             d['qty'], h, d['s'], d['bay'], h_rank_val, df_result_detail)
                if rem > 0:
                    d2 = d.copy()
                    d2['qty'] = rem
                    still_deferred.append(d2)
            deferred = still_deferred

        rehandle_count = sum(d['qty'] for d in deferred)
        if rehandle_count > 0:
            print(f"  INFO: {rehandle_count} containers require re-handling "
                  f"(cross-WC stacking unavoidable):")
            for d in deferred:
                print(f"    Block {d['b']} WC{d['wc']} x{d['qty']} "
                      f"(originally at {d['h_orig']})")
        else:
            print("  All containers assigned with no re-handling required.")

        df_result_detail = pd.DataFrame(df_result_detail)

    else:
        df_result_detail = df_result.copy()
        df_result_detail.insert(1, 'CONTAINER ID', '')
        df_result_detail.insert(2, 'ST', '')
        df_result_detail.insert(3, 'POD', '')
        df_result_detail['YB'] = ''
        df_result_detail['YR'] = ''
        df_result_detail['YT'] = ''
        df_result_detail['YARD POSITION'] = ''

    df_result_detail.sort_values(
        ['MOVE HOUR', 'STS', 'BAY', 'ASSIGNED BLOCK', 'WEIGHT CLASS', 'YB', 'YR', 'YT'],
        inplace=True
    )

    # ============================================================
    # 4c. Extract clash details for reporting (NEW)
    # ============================================================
    clash_details = []
    for (h, b) in u_vars:
        u_val = pulp.value(u_vars[(h, b)])
        e_val = pulp.value(e_vars[(h, b)])
        if u_val is not None and u_val > 1:
            # Tìm các job (STS, BAY) mà block b phục vụ trong giờ h
            jobs = []
            for (s, bay) in jobs_by_hour.get(h, []):
                y_key = (h, s, bay, b)
                if y_key in y_vars and pulp.value(y_vars[y_key]) > 0.5:
                    jobs.append(f"{s}@{bay}")
            clash_details.append({
                'MOVE HOUR': h,
                'BLOCK': b,
                'SỐ LƯỢNG BAY (u)': int(u_val),
                'CLASH (e = u-1)': int(e_val),
                'DANH SÁCH JOB (STS@BAY)': ', '.join(jobs)
            })
    df_clash = pd.DataFrame(clash_details)
    if not df_clash.empty:
        df_clash.sort_values(['MOVE HOUR', 'BLOCK'], inplace=True)

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
    # SHEET: CLASH (ALWAYS CREATED)
    # ============================================================
    ws_clash = wb.create_sheet('CLASH')
    headers_clash = ['MOVE HOUR', 'BLOCK', 'SỐ LƯỢNG BAY (u)', 'CLASH (e = u-1)', 'DANH SÁCH JOB (STS@BAY)']
    for c_idx, hdr in enumerate(headers_clash, 1):
        cell = ws_clash.cell(row=1, column=c_idx, value=hdr)
        cell.font = _font(bold=True, color=C_WHITE)
        cell.fill = _fill(C_DARK_BLUE)
        cell.alignment = _align()
        cell.border = _thin_border()

    if not df_clash.empty:
        for r_idx, row in enumerate(df_clash.itertuples(index=False), 2):
            for c_idx, val in enumerate(row, 1):
                cell = ws_clash.cell(row=r_idx, column=c_idx, value=val)
                cell.font = _font()
                cell.fill = _fill(C_WHITE if r_idx % 2 == 0 else C_ALT_ROW)
                cell.alignment = _align()
                cell.border = _thin_border()
    else:
        # Ghi thông báo không có clash
        cell = ws_clash.cell(row=2, column=1, value='Không có clash nào xảy ra.')
        cell.font = _font()
        cell.fill = _fill(C_WHITE)
        cell.alignment = _align()
        ws_clash.merge_cells(start_row=2, start_column=1, end_row=2, end_column=5)

    # Độ rộng cột
    ws_clash.column_dimensions['A'].width = 14
    ws_clash.column_dimensions['B'].width = 12
    ws_clash.column_dimensions['C'].width = 18
    ws_clash.column_dimensions['D'].width = 18
    ws_clash.column_dimensions['E'].width = 50

    # ============================================================
    # SHEET: RESULT (split per ST — one sheet per size type)
    # ============================================================

    # ── Column definitions (shared across all RESULT sheets) ─────────────────
    core_cols     = ['MOVE HOUR', 'CONT LIST', 'CONTAINER ID',
                      'ST', 'POD', 'STS', 'BAY',
                      'ASSIGNED BLOCK', 'WEIGHT CLASS', 'QUANTITIES']
    position_cols = ['YB', 'YR', 'YT', 'YARD POSITION']
    if container_data_available:
        all_result_cols = core_cols + position_cols
    else:
        all_result_cols = ['MOVE HOUR', 'STS', 'BAY',
                            'ASSIGNED BLOCK', 'WEIGHT CLASS', 'QUANTITIES']

    CONT_LIST_COLS = {'CONT LIST'}
    CONT_ID_COLS   = {'CONTAINER ID', 'ST', 'POD'}
    POSITION_COLS  = set(position_cols)

    col_widths = {
        'MOVE HOUR': 14, 'CONT LIST': 45, 'CONTAINER ID': 20,
        'ST': 10,        'POD': 10,       'STS': 10,  'BAY': 10,
        'ASSIGNED BLOCK': 16, 'WEIGHT CLASS': 14, 'QUANTITIES': 12,
        'YB': 8, 'YR': 8, 'YT': 8, 'YARD POSITION': 18,
    }

    # ── Lấy danh sách ST duy nhất, sắp xếp để đặt tên sheet nhất quán ────────
    if container_data_available and 'ST' in df_result_detail.columns:
        st_values = sorted(df_result_detail['ST'].dropna().unique().tolist())
        st_values = [s for s in st_values if str(s).strip() not in ('', 'nan')]
    else:
        st_values = ['ALL']   # fallback: no ST column → 1 sheet

    if not st_values:
        st_values = ['ALL']

    # ── Viết từng sheet RESULT per ST ────────────────────────────────────────
    for st_idx, st_val in enumerate(st_values, 1):
        sheet_name = f"RESULT {st_idx} ({st_val})" if st_val != 'ALL' else 'RESULT'
        # Truncate to 31 chars (Excel limit)
        sheet_name = sheet_name[:31]

        ws_result = wb.create_sheet(sheet_name)

        # Filter data for this ST
        if st_val == 'ALL':
            df_rd = df_result_detail.reset_index(drop=True)
        else:
            df_rd = df_result_detail[
                df_result_detail['ST'].astype(str).str.strip() == str(st_val).strip()
            ].reset_index(drop=True)

        n_rows = len(df_rd)

        # ── Pre-compute CONT LIST per (MOVE HOUR, BAY) for this ST ───────────
        cont_list_map = {}
        if container_data_available and 'CONTAINER ID' in df_rd.columns:
            for (mh, bay), grp in df_rd.groupby(['MOVE HOUR', 'BAY']):
                ids = [str(v).strip() for v in grp['CONTAINER ID']
                       if str(v).strip() not in ('', 'nan')]
                cont_list_map[(mh, bay)] = ', '.join(ids) if ids else ''

        # ── Header row ────────────────────────────────────────────────────────
        for c_idx, cn in enumerate(all_result_cols, 1):
            cell = ws_result.cell(row=1, column=c_idx, value=cn)
            if cn in CONT_LIST_COLS:
                cell.fill = _fill(C_PALE_BLUE)
            elif cn in CONT_ID_COLS:
                cell.fill = _fill(C_MID_BLUE)
            elif cn in POSITION_COLS:
                cell.fill = _fill(C_LIGHT_BLUE)
            else:
                cell.fill = _fill(C_DARK_BLUE)
            cell.font      = _font(bold=True, color=C_WHITE)
            cell.alignment = _align(wrap=True)
            cell.border    = _thin_border()

        # ── Build CONT LIST merge groups ──────────────────────────────────────
        merge_groups = []
        cont_list_col_idx = all_result_cols.index('CONT LIST') + 1 if 'CONT LIST' in all_result_cols else None
        if container_data_available and cont_list_col_idx:
            prev_key  = None
            grp_start = 2
            for i, (_, row) in enumerate(df_rd.iterrows()):
                cur_key   = (row.get('MOVE HOUR', ''), row.get('BAY', ''))
                excel_row = i + 2
                if cur_key != prev_key:
                    if prev_key is not None:
                        merge_groups.append((prev_key, grp_start, excel_row - 1,
                                             cont_list_map.get(prev_key, '')))
                    prev_key  = cur_key
                    grp_start = excel_row
            if prev_key is not None:
                merge_groups.append((prev_key, grp_start, n_rows + 1,
                                     cont_list_map.get(prev_key, '')))

        # ── Write data rows ───────────────────────────────────────────────────
        group_key   = None
        group_shade = C_ALT_ROW

        for r_idx, (_, row) in enumerate(df_rd.iterrows(), 2):
            this_key = (row.get('MOVE HOUR'), row.get('STS'), row.get('BAY'),
                        row.get('ASSIGNED BLOCK'), row.get('WEIGHT CLASS'))
            if this_key != group_key:
                group_shade = C_WHITE if group_shade == C_ALT_ROW else C_ALT_ROW
                group_key = this_key

            for c_idx, cn in enumerate(all_result_cols, 1):
                if cn == 'CONT LIST':
                    val = None
                else:
                    val = row.get(cn, '')
                    if cn in ('YB', 'YR', 'YT') and val != '':
                        try:
                            val = int(val)
                        except (ValueError, TypeError):
                            pass
                    if val == '' or (isinstance(val, float) and str(val) == 'nan'):
                        val = None

                cell = ws_result.cell(row=r_idx, column=c_idx, value=val)
                cell.font      = _font(color='FF000000')
                cell.fill      = _fill(group_shade)
                cell.alignment = _align(wrap=(cn == 'CONT LIST'))
                cell.border    = _thin_border()

        # ── Write and merge CONT LIST column ──────────────────────────────────
        if cont_list_col_idx:
            for (mh, bay), r_start, r_end, list_text in merge_groups:
                cell = ws_result.cell(row=r_start, column=cont_list_col_idx,
                                       value=list_text or None)
                cell.font      = _font(color='FF000000', size=9)
                cell.fill      = _fill(C_PALE_BLUE)
                cell.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
                cell.border    = _thin_border()

                if r_end > r_start:
                    ws_result.merge_cells(
                        start_row=r_start, start_column=cont_list_col_idx,
                        end_row=r_end,     end_column=cont_list_col_idx
                    )
                    ws_result.cell(row=r_start, column=cont_list_col_idx).alignment =                         Alignment(horizontal='left', vertical='top', wrap_text=True)

            # Row heights
            for (mh, bay), r_start, r_end, list_text in merge_groups:
                span     = r_end - r_start + 1
                n_ids    = len([x for x in list_text.split(',') if x.strip()]) if list_text else 0
                rows_needed = max(1, -(-n_ids // max(1, span)))
                rh = max(15, min(60, rows_needed * 13))
                for r in range(r_start, r_end + 1):
                    ws_result.row_dimensions[r].height = rh

        # ── Column widths ──────────────────────────────────────────────────────
        for c_idx, cn in enumerate(all_result_cols, 1):
            ws_result.column_dimensions[get_column_letter(c_idx)].width =                 col_widths.get(cn, 14)

        print(f"  Sheet '{sheet_name}': {n_rows} rows written.")

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
    # 8. Save to BytesIO buffer
    # ============================================================
    excel_buffer = io.BytesIO()
    wb.save(excel_buffer)
    excel_buffer.seek(0)
    total_rows = len(df_result_detail)
    objective_value = int(pulp.value(prob.objective) or 0)
    print(f"Done. Rows={total_rows}, Clashes={objective_value}")
    return excel_buffer, total_rows, objective_value
