import streamlit as st
import io
import time
import threading
import pandas as pd
from ALLOCATION import run_optimization

# ==========================================
# CẤU HÌNH GIAO DIỆN WEB
# ==========================================
st.set_page_config(page_title="DISTRIBUTION CONTAINER", page_icon="🚢", layout="centered")
st.markdown("""
    <style>
        .main > div {
            max-width: 800px;
            margin: 0 auto;
        }
        .timer-box {
            display: flex;
            align-items: center;
            gap: 12px;
            background: #f0f4ff;
            border: 1px solid #c9d8f5;
            border-radius: 10px;
            padding: 14px 20px;
            margin-top: 8px;
        }
        .timer-label {
            color: #444;
            font-size: 15px;
            font-weight: 500;
        }
        .timer-value {
            font-size: 28px;
            font-weight: 700;
            color: #1a56db;
            font-variant-numeric: tabular-nums;
            letter-spacing: 2px;
        }
        .spinner-msg {
            color: #555;
            font-size: 15px;
            margin-bottom: 4px;
        }
    </style>
""", unsafe_allow_html=True)

st.title("🚢 HỆ THỐNG PHÂN BỔ TỐI ƯU TỰ ĐỘNG")
st.markdown("Tải file dữ liệu đầu vào (Excel) để hệ thống chạy thuật toán và trả về kết quả.")

# ==========================================
# KHU VỰC TẢI FILE LÊN (IMPORT)
# ==========================================
uploaded_file = st.file_uploader("📂 Tải lên file Excel Input của bạn", type=["xlsx", "xls"])

if uploaded_file is not None:
    st.info("Đã nhận file. Bấm nút bên dưới để tiến hành tính toán.")

    if st.button("🚀 CHẠY THUẬT TOÁN TỐI ƯU", use_container_width=True):

        # --- Đọc bytes trước để tránh stream bị đóng khi chạy thread ---
        file_bytes = io.BytesIO(uploaded_file.read())

        # --- Kết quả & lỗi được truyền qua dict dùng chung ---
        result_holder = {"done": False, "result": None, "error": None}

        def run_in_thread():
            try:
                result_holder["result"] = run_optimization(file_bytes)
            except Exception as e:
                result_holder["error"] = e
            finally:
                result_holder["done"] = True

        # --- Khởi động thread tính toán ---
        t = threading.Thread(target=run_in_thread, daemon=True)
        t.start()
        start_time = time.time()

        # --- Khu vực hiển thị thông báo + đồng hồ ---
        st.markdown('<p class="spinner-msg">⏳ Hệ thống đang tính toán và xuất file... Vui lòng đợi!</p>',
                    unsafe_allow_html=True)
        timer_placeholder = st.empty()

        # --- Vòng lặp cập nhật đồng hồ cho đến khi xong ---
        while not result_holder["done"]:
            elapsed = int(time.time() - start_time)
            mm, ss = divmod(elapsed, 60)
            timer_placeholder.markdown(
                f"""
                <div class="timer-box">
                    <span class="timer-label">⏱ Thời gian tính toán:</span>
                    <span class="timer-value">{mm:02d}:{ss:02d}</span>
                </div>
                """,
                unsafe_allow_html=True
            )
            time.sleep(0.5)

        # --- Hiển thị thời gian cuối cùng (đứng yên) ---
        elapsed = int(time.time() - start_time)
        mm, ss = divmod(elapsed, 60)
        timer_placeholder.markdown(
            f"""
            <div class="timer-box" style="background:#f0fff4; border-color:#a3d9a5;">
                <span class="timer-label">✅ Hoàn tất sau:</span>
                <span class="timer-value" style="color:#1a7f37;">{mm:02d}:{ss:02d}</span>
            </div>
            """,
            unsafe_allow_html=True
        )

        # --- Xử lý kết quả ---
        if result_holder["error"]:
            st.error(f"❌ Có lỗi xảy ra trong quá trình tính toán.\n\nChi tiết lỗi: {result_holder['error']}")
        else:
            excel_buffer, total_rows, objective_value = result_holder["result"]

            st.success("🎉 Tính toán hoàn tất!")

            col1, col2 = st.columns(2)
            col1.metric("Tổng số dòng phân bổ", f"{total_rows} dòng")
            col2.metric("Số lượng Clash (giá trị mục tiêu)", f"{objective_value}", delta_color="inverse")

            st.download_button(
                label="📥 TẢI FILE KẾT QUẢ",
                data=excel_buffer,
                file_name="PHAN_BO.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
