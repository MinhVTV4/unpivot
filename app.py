import streamlit as st
import pandas as pd
from io import BytesIO
import json
import os

st.set_page_config(page_title="Excel Hub Pro v2", layout="wide", page_icon="📑")

CONFIG_FILE = "app_profiles_v2.json"

# --- HÀM TRỢ GIÚP ---
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

# --- MODULE 1: UNPIVOT ---
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
        st.error(f"Lỗi: {e}")
        return None

# --- GIAO DIỆN CHÍNH ---
st.sidebar.title("🎮 Hệ thống Xử lý Excel")
app_mode = st.sidebar.selectbox("Nghiệp vụ:", ["🔄 Unpivot (Ngang sang Dọc)", "🔍 Đối soát dữ liệu"])

# --- TAB 1: UNPIVOT ---
if app_mode == "🔄 Unpivot (Ngang sang Dọc)":
    st.title("🔄 Unpivot Ma trận Đa Sheet")
    
    file_up = st.file_uploader("Tải file Excel", type=["xlsx", "xls"])
    
    if file_up:
        # Lấy danh sách Sheet mà không cần load toàn bộ data (tiết kiệm RAM)
        xl = pd.ExcelFile(file_up)
        sheet_names = xl.sheet_names
        selected_sheet = st.selectbox("📂 Chọn Sheet chứa dữ liệu ma trận:", sheet_names)
        
        # Đọc dữ liệu từ sheet đã chọn
        df_raw = pd.read_excel(file_up, sheet_name=selected_sheet, header=None)
        
        with st.sidebar:
            st.header("⚙️ Cấu hình Profile")
            p_names = list(st.session_state['profiles'].keys())
            sel_p = st.selectbox("Chọn Profile:", p_names)
            cfg = st.session_state['profiles'][sel_p]
            
            h_r = st.number_input("Số hàng tiêu đề:", value=cfg['h_rows'])
            i_c = st.number_input("Cột Tên (B=1):", value=cfg['id_col'])
            d_s = st.number_input("Dòng bắt đầu data:", value=cfg['d_start'])
            
            if st.button("🚀 Chạy Unpivot"):
                res = run_unpivot(df_raw, h_r, i_c, d_s)
                if res is not None:
                    st.success(f"Xử lý thành công sheet '{selected_sheet}'")
                    st.dataframe(res)
                    out = BytesIO()
                    res.to_excel(out, index=False)
                    st.download_button("📥 Tải File Đích", out.getvalue(), f"unpivot_{selected_sheet}.xlsx")

# --- TAB 2: ĐỐI SOÁT ---
elif app_mode == "🔍 Đối soát dữ liệu":
    st.title("🔍 Đối soát Đa Sheet")
    
    c1, c2 = st.columns(2)
    with c1:
        f_m = st.file_uploader("Tải File Master", type=["xlsx"])
        if f_m:
            xl_m = pd.ExcelFile(f_m)
            s_m = st.selectbox("Chọn Sheet Master:", xl_m.sheet_names)
            df_m = pd.read_excel(f_m, sheet_name=s_m)
            
    with c2:
        f_c = st.file_uploader("Tải File Cần đối soát", type=["xlsx"])
        if f_c:
            xl_c = pd.ExcelFile(f_c)
            s_c = st.selectbox("Chọn Sheet cần check:", xl_c.sheet_names)
            df_c = pd.read_excel(f_c, sheet_name=s_c)

    if f_m and f_check:
        st.sidebar.header("⚙️ Cài đặt So khớp")
        key = st.sidebar.selectbox("Cột Mã khóa (Key):", df_m.columns)
        val = st.sidebar.selectbox("Cột Số tiền để so:", df_m.columns)

        if st.button("🚀 Bắt đầu đối soát"):
            # Logic Đối soát... (giữ nguyên như bản trước)
            merged = pd.merge(df_m, df_c[[key, val]], on=key, how='outer', suffixes=('_Gốc', '_ThựcTế'))
            merged = merged.fillna(0)
            merged['Chênh lệch'] = merged[f'{val}_Gốc'] - merged[f'{val}_ThựcTế']
            
            st.subheader(f"Kết quả đối soát giữa '{s_m}' và '{s_c}'")
            st.dataframe(merged)
            
            out_err = BytesIO()
            merged[merged['Chênh lệch'] != 0].to_excel(out_err, index=False)
            st.download_button("📥 Tải báo cáo chênh lệch", out_err.getvalue(), "diff.xlsx")
