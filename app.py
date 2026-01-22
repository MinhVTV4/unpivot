import streamlit as st
import pandas as pd
from io import BytesIO
import json
import os

# Cấu hình trang
st.set_page_config(page_title="Excel Hub Pro v2.1", layout="wide", page_icon="📈")

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

# --- MODULE 1: LOGIC UNPIVOT CHI TIẾT (FIXED) ---
def run_unpivot_detailed(df, h_rows, id_col, d_start):
    try:
        # 1. Tách tiêu đề và dữ liệu
        header_part = df.iloc[:h_rows, id_col+1:]
        data_part = df.iloc[d_start-1:, :].copy()
        
        # 2. Tạo ID tạm cho các cột bằng cách nối tiêu đề với ký tự đặc biệt "||"
        separator = "||"
        combined_headers = []
        for col_idx in range(id_col + 1, len(df.columns)):
            # Lấy giá trị của từng hàng tiêu đề tại cột này
            h_vals = [str(header_part.iloc[r, col_idx-(id_col+1)]).strip() for r in range(h_rows)]
            combined_headers.append(separator.join(h_vals))
            
        # 3. Gán tên cột cho data_part
        id_col_name = "Đối tượng"
        # Đặt tên cho các cột không dùng đến để tránh trùng lặp
        new_cols = [f"tmp_{i}" for i in range(id_col)] + [id_col_name] + combined_headers
        data_part.columns = new_cols
        
        # 4. Thực hiện Melt (Xoay bảng)
        result = pd.melt(
            data_part, 
            id_vars=[id_col_name], 
            value_vars=combined_headers,
            var_name="Temp_Header", 
            value_name="Giá trị"
        )
        
        # 5. Tách ngược Temp_Header ra lại thành các cột Tiêu đề 1, Tiêu đề 2...
        header_split = result['Temp_Header'].str.split(separator, expand=True)
        for i in range(h_rows):
            result[f"Tiêu đề {i+1}"] = header_split[i].replace('nan', '')

        # 6. Dọn dẹp: Bỏ cột tạm, ép kiểu số, lọc bỏ giá trị trống/bằng 0
        result = result.drop(columns=['Temp_Header'])
        result['Giá trị'] = pd.to_numeric(result['Giá trị'], errors='coerce')
        result = result.dropna(subset=['Giá trị'])
        result = result[result['Giá trị'] != 0]
        
        # Sắp xếp lại thứ tự cột cho đẹp: Đối tượng -> Các tiêu đề -> Giá trị
        cols_order = [id_col_name] + [f"Tiêu đề {i+1}" for i in range(h_rows)] + ["Giá trị"]
        return result[cols_order]

    except Exception as e:
        st.error(f"Lỗi Unpivot chi tiết: {e}")
        return None

# --- GIAO DIỆN ---
st.sidebar.title("🎮 Menu Chức năng")
app_mode = st.sidebar.selectbox("Chọn nghiệp vụ:", ["🔄 Unpivot Vạn năng", "🔍 Đối soát & So khớp"])

if app_mode == "🔄 Unpivot Vạn năng":
    st.title("🔄 Trình Unpivot Excel Ma trận (Chi tiết)")
    
    with st.sidebar:
        st.header("⚙️ Cấu hình Profile")
        p_names = list(st.session_state['profiles'].keys())
        sel_p = st.selectbox("Chọn Profile:", p_names)
        cfg = st.session_state['profiles'][sel_p]
        
        h_r = st.number_input("Số hàng tiêu đề:", value=int(cfg['h_rows']))
        i_c = st.number_input("Cột Định danh (B=1):", value=int(cfg['id_col']))
        d_s = st.number_input("Dòng bắt đầu dữ liệu:", value=int(cfg['d_start']))
        
        if st.button("💾 Lưu cấu hình mới"):
            new_p_name = st.text_input("Tên Profile:", value="Profile mới")
            st.session_state['profiles'][new_p_name] = {"h_rows": h_r, "id_col": i_c, "d_start": d_s}
            save_profiles(st.session_state['profiles'])
            st.success("Đã lưu!")

    file_up = st.file_uploader("Tải file ma trận ngang", type=["xlsx"])
    if file_up:
        xl = pd.ExcelFile(file_up)
        sheet = st.selectbox("Chọn Sheet:", xl.sheet_names)
        df_raw = xl.parse(sheet, header=None)
        
        st.write("---")
        if st.button("🚀 Thực hiện Unpivot Chi tiết"):
            with st.spinner('Đang xử lý...'):
                res = run_unpivot_detailed(df_raw, h_r, i_c, d_s)
                if res is not None:
                    st.success(f"Xong! Đã tách thành {len(res)} dòng chi tiết.")
                    st.dataframe(res, use_container_width=True)
                    
                    # Tải file
                    out = BytesIO()
                    res.to_excel(out, index=False)
                    st.download_button("📥 Tải File Kết Quả", out.getvalue(), "unpivot_detailed.xlsx")

elif app_mode == "🔍 Đối soát & So khớp":
    # (Giữ nguyên phần đối soát ở bản trước vì nó đã tách biệt các cột tiền và khóa)
    st.title("🔍 Hệ thống Đối soát & Cảnh báo")
    f_master = st.file_uploader("Tải File Gốc (Master)", type=["xlsx"])
    f_check = st.file_uploader("Tải File Đối soát", type=["xlsx"])
    
    if f_master and f_check:
        df_m = pd.read_excel(f_master)
        df_c = pd.read_excel(f_check)
        
        c1, c2 = st.columns(2)
        with c1: key_m = st.selectbox("Cột Khóa (Gốc):", df_m.columns)
        with c2: key_c = st.selectbox("Cột Khóa (Thực tế):", df_c.columns)
        
        val_m = st.selectbox("Cột Số tiền cần so sánh:", df_m.columns)

        if st.button("🚀 Chạy Đối soát"):
            merged = pd.merge(df_m, df_c, left_on=key_m, right_on=key_c, how='outer', suffixes=('_Gốc', '_Thực tế'))
            merged = merged.fillna(0)
            # Giả định cột tiền ở file check có tên tương đương hoặc người dùng chọn
            # Để đơn giản, tôi lấy cột có tên giống val_m ở file check
            val_c = val_m if val_m in df_c.columns else df_c.columns[0] 
            
            merged['Chênh lệch'] = merged[f'{val_m}_Gốc'] - merged.get(f'{val_m}_Thực tế', 0)
            st.dataframe(merged)
