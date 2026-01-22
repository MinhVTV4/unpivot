import streamlit as st
import pandas as pd
from io import BytesIO
import json
import os

# Cấu hình trang
st.set_page_config(page_title="Excel Pro Transformer", layout="wide", page_icon="🚀")

CONFIG_FILE = "profiles_config.json"

# --- HÀM LƯU/ĐỌC CẤU HÌNH VÀO FILE ---
def load_profiles():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"Mẫu SDH Gốc": {"h_rows": 3, "id_col": 1, "d_start": 5}}

def save_profiles(profiles):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(profiles, f, ensure_ascii=False, indent=4)

# Khởi tạo danh sách cấu hình
if 'profiles' not in st.session_state:
    st.session_state['profiles'] = load_profiles()

# --- HÀM XỬ LÝ UNPIVOT TỔNG QUÁT ---
def universal_unpivot(df, h_rows, id_col, d_start):
    try:
        headers = df.iloc[0:h_rows, id_col + 1:]
        data_body = df.iloc[d_start - 1:, :]
        
        results = []
        for _, row in data_body.iterrows():
            id_val = str(row[id_col]).strip()
            if not id_val or id_val.lower() in ['nan', 'none']: continue
            
            for col_idx in range(id_col + 1, len(df.columns)):
                val = pd.to_numeric(row[col_idx], errors='coerce')
                if pd.notnull(val) and val > 0:
                    entry = {"Đối tượng/Tên": id_val, "Số tiền": val}
                    for i in range(h_rows):
                        entry[f"Tiêu đề {i+1}"] = headers.iloc[i, col_idx - (id_col + 1)]
                    results.append(entry)
        return pd.DataFrame(results)
    except Exception as e:
        st.error(f"Lỗi xử lý: {e}")
        return None

# --- GIAO DIỆN CHÍNH ---
st.title("🗂️ Trình xử lý Excel Ma trận Vạn năng")
st.markdown("Hỗ trợ xử lý file hàng ngàn dòng, lưu cấu hình và xuất mẫu in tự động.")

# SIDEBAR: QUẢN LÝ CẤU HÌNH
with st.sidebar:
    st.header("⚙️ Thiết lập loại File")
    
    # Chọn Profile
    profile_names = list(st.session_state['profiles'].keys())
    selected_p = st.selectbox("Chọn loại file đã lưu:", profile_names)
    
    # Lấy thông số từ profile đã chọn
    cfg = st.session_state['profiles'][selected_p]
    
    st.markdown("---")
    st.subheader("Tùy chỉnh cấu hình")
    h_rows = st.number_input("Số hàng tiêu đề:", value=cfg['h_rows'])
    id_col = st.number_input("Cột chứa Tên (A=0, B=1...):", value=cfg['id_col'])
    d_start = st.number_input("Dữ liệu bắt đầu từ hàng:", value=cfg['d_start'])
    
    st.markdown("---")
    new_p_name = st.text_input("Lưu cấu hình này với tên mới:", placeholder="Ví dụ: File Kho vận")
    if st.button("💾 Lưu cấu hình"):
        st.session_state['profiles'][new_p_name] = {"h_rows": h_rows, "id_col": id_col, "d_start": d_start}
        save_profiles(st.session_state['profiles'])
        st.success(f"Đã lưu '{new_p_name}' thành công!")
        st.rerun()

# KHU VỰC TẢI FILE
uploaded_file = st.file_uploader("Tải lên file Excel cần xử lý", type=["xlsx", "xls"])

if uploaded_file:
    df_raw = pd.read_excel(uploaded_file, header=None)
    
    st.subheader("1. Kiểm tra cấu trúc File (Preview)")
    st.dataframe(df_raw.head(15), use_container_width=True)
    
    st.write(f"👉 Đang dùng cấu hình: **{selected_p}**")

    if st.button("🚀 Bắt đầu chuyển đổi ngay", type="primary"):
        with st.spinner("Đang 'bẻ' bảng ngang sang dọc..."):
            df_result = universal_unpivot(df_raw, h_rows, id_col, d_start)
            
            if df_result is not None and not df_result.empty:
                st.success(f"Đã xử lý xong {len(df_result)} dòng dữ liệu!")
                
                tab1, tab2 = st.tabs(["📊 Dữ liệu Đích (Dọc)", "🖨️ Xuất Mẫu In Nhanh"])
                
                with tab1:
                    st.dataframe(df_result, use_container_width=True)
                    # Tải file CSV
                    csv = df_result.to_csv(index=False).encode('utf-8-sig')
                    st.download_button("📥 Tải File Đích (.csv)", csv, "ket_qua_doc.csv")
                
                with tab2:
                    st.info("Hệ thống sẽ tạo file Excel có tiêu đề và kẻ bảng tự động dựa trên kết quả dọc.")
                    # Tạo file Excel đẹp
                    out_excel = BytesIO()
                    with pd.ExcelWriter(out_excel, engine='xlsxwriter') as writer:
                        df_result.to_excel(writer, index=False, sheet_name='Mau_In')
                        workbook = writer.book
                        worksheet = writer.sheets['Mau_In']
                        # Định dạng đơn giản
                        fmt_header = workbook.add_format({'bold': True, 'bg_color': '#D9D9D9', 'border': 1})
                        for col_num, value in enumerate(df_result.columns.values):
                            worksheet.write(0, col_num, value, fmt_header)
                            worksheet.set_column(col_num, col_num, 20)
                    
                    st.download_button("📥 Tải Mẫu In Excel", out_excel.getvalue(), "mau_in_nhanh.xlsx")
            else:
                st.warning("Không tìm thấy dữ liệu phát sinh > 0.")
