import streamlit as st
import pandas as pd
from io import BytesIO
import json
import os

# Cấu hình trang
st.set_page_config(page_title="Excel Hub Pro v2", layout="wide", page_icon="📈")

CONFIG_FILE = "app_profiles.json"

# --- HÀM QUẢN LÝ CẤU HÌNH ---
def load_profiles():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except: return {"Mẫu SDH Mặc định": {"h_rows": 3, "id_col": 1, "d_start": 5}}
    return {"Mẫu SDH Mặc định": {"h_rows": 3, "id_col": 1, "d_start": 5}}

def save_profiles(profiles):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(profiles, f, ensure_ascii=False, indent=4)

if 'profiles' not in st.session_state:
    st.session_state['profiles'] = load_profiles()

# --- MODULE 1: LOGIC UNPIVOT NÂNG CẤP ---

def run_unpivot_fast(df, h_rows, id_col, d_start):
    try:
        # Tách tiêu đề và dữ liệu
        header_part = df.iloc[:h_rows, id_col+1:]
        data_part = df.iloc[d_start-1:, :].copy()
        
        # Tạo tên cột gộp từ các hàng tiêu đề
        combined_columns = []
        for col_idx in range(id_col + 1, len(df.columns)):
            col_parts = [str(header_part.iloc[r, col_idx-(id_col+1)]).replace('nan', '').strip() for r in range(h_rows)]
            combined_columns.append(" | ".join([p for p in col_parts if p]))
            
        # Gán lại tên cột cho phần dữ liệu
        id_col_name = "Mã/Đối tượng"
        # Đặt tên tạm cho các cột trước cột ID
        new_cols = [f"ignore_{i}" for i in range(id_col)] + [id_col_name] + combined_columns
        data_part.columns = new_cols
        
        # Unpivot bằng melt
        result = pd.melt(
            data_part, 
            id_vars=[id_col_name], 
            value_vars=combined_columns,
            var_name="Phân loại/Thời gian", 
            value_name="Giá trị"
        )
        
        # Làm sạch dữ liệu
        result['Giá trị'] = pd.to_numeric(result['Giá trị'], errors='coerce')
        result = result.dropna(subset=['Giá trị'])
        result = result[result['Giá trị'] != 0]
        return result.sort_values(by=id_col_name)
    except Exception as e:
        st.error(f"Lỗi Unpivot: {e}")
        return None

# --- GIAO DIỆN SIDEBAR ---
st.sidebar.title("🎮 Menu Chức năng")
app_mode = st.sidebar.selectbox("Chọn nghiệp vụ:", ["🔄 Unpivot Vạn năng", "🔍 Đối soát & So khớp"])

# --- CHỨC NĂNG 1: UNPIVOT ---
if app_mode == "🔄 Unpivot Vạn năng":
    st.title("🔄 Trình Unpivot Excel Ma trận")
    
    with st.sidebar:
        st.header("⚙️ Cấu hình Profile")
        p_names = list(st.session_state['profiles'].keys())
        sel_p = st.selectbox("Chọn Profile:", p_names)
        cfg = st.session_state['profiles'][sel_p]
        
        h_r = st.number_input("Số hàng tiêu đề:", value=cfg['h_rows'])
        i_c = st.number_input("Cột Định danh (A=0, B=1):", value=cfg['id_col'])
        d_s = st.number_input("Dòng bắt đầu dữ liệu:", value=cfg['d_start'])
        
        new_p = st.text_input("Tên Profile mới:")
        if st.button("💾 Lưu cấu hình"):
            st.session_state['profiles'][new_p] = {"h_rows": h_r, "id_col": i_c, "d_start": d_s}
            save_profiles(st.session_state['profiles'])
            st.success("Đã lưu!")

    file_up = st.file_uploader("Tải file ma trận ngang", type=["xlsx", "xls"])
    if file_up:
        xl = pd.ExcelFile(file_up)
        sheet = st.selectbox("Chọn Sheet dữ liệu:", xl.sheet_names)
        df_raw = xl.parse(sheet, header=None)
        
        st.subheader("Xem trước dữ liệu gốc")
        st.dataframe(df_raw.head(10), use_container_width=True)
        
        if st.button("🚀 Thực hiện Unpivot"):
            with st.spinner('Đang xoay trục dữ liệu...'):
                res = run_unpivot_fast(df_raw, h_r, i_c, d_s)
                if res is not None:
                    st.success(f"Xử lý xong! Tìm thấy {len(res)} bản ghi có giá trị.")
                    st.dataframe(res, use_container_width=True)
                    
                    out = BytesIO()
                    res.to_excel(out, index=False)
                    st.download_button("📥 Tải File Dọc (.xlsx)", out.getvalue(), "unpivot_result.xlsx")

# --- CHỨC NĂNG 2: ĐỐI SOÁT ---
elif app_mode == "🔍 Đối soát & So khớp":
    st.title("🔍 Hệ thống Đối soát & Cảnh báo")

    col_a, col_b = st.columns(2)
    with col_a:
        f_master = st.file_uploader("1. File Gốc (Master)", type=["xlsx", "csv"])
    with col_b:
        f_check = st.file_uploader("2. File Cần đối soát", type=["xlsx", "csv"])

    if f_master and f_check:
        # Đọc dữ liệu
        df_m = pd.read_excel(f_master) if f_master.name.endswith('xlsx') else pd.read_csv(f_master)
        df_c = pd.read_excel(f_check) if f_check.name.endswith('xlsx') else pd.read_csv(f_check)
        
        st.divider()
        c1, c2 = st.columns(2)
        with c1:
            key_m = st.selectbox("Cột Khóa (Gốc):", df_m.columns, key="km")
            val_m = st.selectbox("Cột Tiền (Gốc):", df_m.columns, key="vm")
        with c2:
            key_c = st.selectbox("Cột Khóa (Đối soát):", df_c.columns, key="kc")
            val_c = st.selectbox("Cột Tiền (Đối soát):", df_c.columns, key="vc")

        if st.button("🚀 Bắt đầu đối soát"):
            with st.spinner('Đang so khớp dữ liệu...'):
                # Merge dữ liệu
                merged = pd.merge(
                    df_m[[key_m, val_m]], 
                    df_c[[key_c, val_c]], 
                    left_on=key_m, 
                    right_on=key_c, 
                    how='outer', 
                    suffixes=('_Gốc', '_ThựcTế')
                )
                
                # Xử lý giá trị Null
                merged = merged.fillna(0)
                # Đảm bảo cột ID không bị 0 nếu một bên thiếu
                merged['ID_Final'] = merged[key_m].where(merged[key_m] != 0, merged[key_c])
                
                # Tính toán
                merged['Chênh lệch'] = merged[f'{val_m}_Gốc'] - merged[f'{val_c}_ThựcTế']
                
                # Cảnh báo Outliers
                std = merged['Chênh lệch'].std()
                merged['Trạng thái'] = merged['Chênh lệch'].apply(
                    lambda x: '🚩 Sai lệch lớn' if abs(x) > (2 * std) and x != 0 else ('✅ Khớp' if x == 0 else '⚠️ Lệch nhẹ')
                )

                # Hiển thị thống kê
                s1, s2, s3 = st.columns(3)
                s1.metric("Tổng dòng", len(merged))
                s2.metric("Số dòng lệch", len(merged[merged['Chênh lệch'] != 0]))
                s3.metric("Tổng chênh lệch", f"{merged['Chênh lệch'].sum():,.0f}")

                st.subheader("Bảng chi tiết kết quả")
                st.dataframe(
                    merged.style.applymap(
                        lambda x: 'background-color: #ffcccc' if x == '🚩 Sai lệch lớn' else ('background-color: #fff4cc' if x == '⚠️ Lệch nhẹ' else ''),
                        subset=['Trạng thái']
                    ), use_container_width=True
                )
                
                # Xuất file
                out_err = BytesIO()
                merged.to_excel(out_err, index=False)
                st.download_button("📥 Tải Báo cáo Đối soát FULL", out_err.getvalue(), "doi_soat_chi_tiet.xlsx")
