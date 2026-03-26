import streamlit as st
import io
import pandas as pd
import time

# ============================================================
# ADDED: Import hàm xử lý chính từ file ALLOCATION.py
# ============================================================
from ALLOCATION import run_optimization

# ==========================================
# CẤU HÌNH GIAO DIỆN WEB
# ==========================================
st.set_page_config(page_title="DISTRIBUTION CONTAINER", page_icon="🚢", layout="centered")

st.title("🚢 HỆ THỐNG PHÂN BỔ TỐI ƯU")
st.markdown("Tải file dữ liệu đầu vào (Excel) để hệ thống chạy thuật toán và trả về kết quả.")

# ==========================================
# KHU VỰC TẢI FILE LÊN (IMPORT)
# ==========================================
uploaded_file = st.file_uploader("📂 Tải lên file Excel Input của bạn", type=["xlsx", "xls"])

if uploaded_file is not None:
    st.info("Đã nhận file. Bấm nút bên dưới để tiến hành tính toán.")
    
    if st.button("🚀 CHẠY THUẬT TOÁN TỐI ƯU", use_container_width=True):
        start_time = time.time()  # bắt đầu đo
        with st.spinner('Hệ thống đang tính toán và xuất file... Vui lòng đợi!'):
            try:
                # ==========================================================
                # Gọi hàm run_optimization với file upload (đã là BytesIO)
                # Hàm trả về buffer (chứa file Excel kết quả), số dòng, tổng clash, thời gian chạy
                # ==========================================================
                excel_buffer, total_rows, total_clashes, exec_time = run_optimization(uploaded_file)
                # exec_time là thời gian đã được tính bên trong hàm (có thể bỏ qua)
                elapsed = time.time() - start_time  # tính lại để chính xác hơn
                
                st.success(f"✅ Tính toán hoàn tất! Thời gian thực: {elapsed:.2f} giây")
                
                # ==========================================================
                # Hiển thị thống kê ngắn gọn
                # ==========================================================
                col1, col2 = st.columns(2)
                col1.metric("Tổng số dòng phân bổ", f"{total_rows} dòng")
                col2.metric("Số lượng Clash (u-1)", f"{total_clashes}", delta_color="inverse")
                
                # ==========================================================
                # Nút tải file kết quả
                # ==========================================================
                st.download_button(
                    label="📥 TẢI FILE KẾT QUẢ ",
                    data=excel_buffer,
                    file_name="PHAN_BO.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )

            except Exception as e:
                st.error(f"❌ Có lỗi xảy ra trong quá trình tính toán.\n\nChi tiết lỗi: {e}")
