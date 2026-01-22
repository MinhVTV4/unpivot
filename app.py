import streamlit as st
import pandas as pd
from io import BytesIO
import json
import os

# Cấu hình trang
st.set_page_config(page_title="Excel Hub Pro v3", layout="wide", page_icon="📑")

CONFIG_FILE = "excel_hub_profiles.json"

# --- HÀM QUẢN LÝ CẤU HÌNH ---
def load_profiles():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except: return {}
    return {"Mẫu SDH Mặc định": {"h_rows": 3, "id_col": 1, "d_start": 5}}

def save_profiles(profiles):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(profiles, f, ensure_ascii=False, indent=4)

if 'profiles' not in st.session_state:
    st.session_state['profiles'] = load_profiles()

# --- MODULE XỬ LÝ UNPIVOT ---
def run_unpivot(df, h_rows, id_col, d_start):
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
                    entry = {"Đối tượng": id_val, "Số tiền": val}
                    for i in range(h_rows):
                        entry[f"Tiêu đề {i+1}"] = headers.iloc[i, col_idx - (id_col + 1)]
                    results.append(entry)
        return pd.DataFrame(results)
    except Exception as e:
        st.error(f"Lỗi logic: {e}")
        return None

# --- GIAO DIỆN CHÍNH ---
st.sidebar.title("🎮 Hệ thống Excel Pro")
app_mode = st.sidebar.radio("Nghiệp vụ cần xử lý:", ["🔄 Unpivot (Ngang -> Dọc)", "🔍 Đối soát dữ liệu"])

# --- 1. MODULE UNPIVOT ---
if app_mode == "🔄 Unpivot (Ngang -> Dọc)":
    st.title("🔄 Unpivot Ma trận Đa Sheet")
    st.markdown("Chọn Sheet và cấu hình bên trái để 'bẻ' bảng ngang.")

    file_up = st.file_uploader("Bước 1: Tải file Excel lên", type=["xlsx", "xls"], key="unpivot_upload")

    if file_up:
        # Lấy danh sách Sheet
        xl = pd.ExcelFile(file_up)
        sheet_names = xl.sheet_names
        selected_sheet = st.selectbox("Bước 2: Chọn Sheet chứa dữ liệu:", sheet_names)

        # ĐỌC DỮ LIỆU VÀ HIỂN THỊ PREVIEW NGAY LẬP TỨC
        df_raw = pd.read_excel(file_up, sheet_name=selected_sheet, header=None)
        
        st.subheader(f"📋 3. Preview dữ liệu (Sheet: {selected_sheet})")
        st.dataframe(df_raw.head(20), use_container_width=True)

        # CẤU HÌNH SIDEBAR
        with st.sidebar:
            st.markdown("---")
            st.header("⚙️ Cấu hình cấu trúc")
            p_names = list(st.session_state['profiles'].keys())
            sel_p = st.selectbox("Sử dụng Profile:", p_names)
            cfg = st.session_state['profiles'][sel_p]
            
            h_r = st.number_input("Số hàng tiêu đề:", value=cfg['h_rows'])
            i_c = st.number_input("Cột Tên (A=0, B=1...):", value=cfg['id_col'])
            d_s = st.number_input("Dòng bắt đầu dữ liệu:", value=cfg['d_start'])
            
            save_name = st.text_input("Lưu cấu hình mới với tên:")
            if st.button("💾 Lưu Profile"):
                st.session_state['profiles'][save_name] = {"h_rows": h_r, "id_col": i_c, "d_start": d_s}
                save_profiles(st.session_state['profiles'])
                st.success("Đã lưu cấu hình!")

        if st.button("🚀 Bắt đầu Unpivot", type="primary"):
            with st.spinner("Đang xử lý hàng ngàn dòng..."):
                res = run_unpivot(df_raw, h_r, i_c, d_s)
                if res is not None and not res.empty:
                    st.success("Hoàn tất!")
                    st.dataframe(res, use_container_width=True)
                    out = BytesIO()
                    res.to_excel(out, index=False)
                    st.download_button("📥 Tải kết quả", out.getvalue(), f"unpivot_{selected_sheet}.xlsx")
                else:
                    st.warning("Không tìm thấy dữ liệu phát sinh > 0.")

# --- 2. MODULE ĐỐI SOÁT ---
elif app_mode == "🔍 Đối soát dữ liệu":
    st.title("🔍 Đối soát dữ liệu Đa Sheet")
    
    col1, col2 = st.columns(2)
    with col1:
        f_m = st.file_uploader("Tải File Master (Gốc)", type=["xlsx"], key="m")
        if f_m:
            xl_m = pd.ExcelFile(f_m)
            s_m = st.selectbox("Chọn Sheet Master:", xl_m.sheet_names)
            df_m = pd.read_excel(f_m, sheet_name=s_m)
            st.dataframe(df_m.head(5))

    with col2:
        f_c = st.file_uploader("Tải File Cần đối soát", type=["xlsx"], key="c")
        if f_c:
            xl_c = pd.ExcelFile(f_c)
            s_c = st.selectbox("Chọn Sheet đối soát:", xl_c.sheet_names)
            df_c = pd.read_excel(f_c, sheet_name=s_c)
            st.dataframe(df_c.head(5))

    if f_m and f_c:
        st.sidebar.markdown("---")
        st.sidebar.header("⚙️ Cài đặt so khớp")
        # Giả định file đối soát đã có header chuẩn
        key_col = st.sidebar.selectbox("Cột Mã khóa (Key):", df_m.columns)
        val_col = st.sidebar.selectbox("Cột Số tiền cần so:", df_m.columns)

        if st.button("🚀 Thực hiện đối soát", type="primary"):
            # Logic Merge & So sánh
            merged = pd.merge(df_m, df_c[[key_col, val_col]], on=key_col, how='outer', suffixes=('_Gốc', '_ThựcTế'))
            merged = merged.fillna(0)
            merged['Chênh lệch'] = merged[f'{val_col}_Gốc'] - merged[f'{val_col}_ThựcTế']
            
            st.subheader("Báo cáo chênh lệch")
            st.dataframe(merged[merged['Chênh lệch'] != 0])
            
            out_err = BytesIO()
            merged.to_excel(out_err, index=False)
            st.download_button("📥 Tải toàn bộ báo cáo đối soát", out_err.getvalue(), "doi_soat_chi_tiet.xlsx")
