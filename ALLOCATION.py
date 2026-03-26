# ============================================================
# 3a. Objective: minimise clashes + movement penalties
# ============================================================
CLASH_W  = 100.0    # ưu tiên cao nhất
SINGLE_W = 10.0     # phạt nặng nếu job chỉ có 1 block
SPREAD_W = 5.0      # phạt mỗi cặp (block, bay)
BLOCK_BAY_WC_W = 2.0
BAY_SINGLE_W = 10.0

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

# ========== Ràng buộc cứng: mỗi bay có ít nhất 2 block ==========
min_blocks_per_bay = 2
for bay in all_bays:
    prob += pulp.lpSum(block_bay[(b, bay)] for b in blocks) >= min_blocks_per_bay
