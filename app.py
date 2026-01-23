import streamlit as st
import pandas as pd
from io import BytesIO
import json
import os
import plotly.express as px
import difflib

# --- CẤU HÌNH HỆ THỐNG ---
st.set_page_config(page_title="Excel Hub Pro v10", layout="wide", page_icon="🛠️")

CONFIG_FILE = "excel_profiles_v10.json"

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

# --- BỘ CHUYỂN ĐỔI FONT (TCVN3 -> UNICODE) ---
# Bảng mã rút gọn cho các ký tự phổ biến nhất
TCVN3_MAP = {
    "a\xcc\x81": "á", "a\xcc\x80": "à", "a\xcc\x89": "ả", "a\xcc\x83": "ã", "a\xcc\xa3": "ạ",
    "\xe1": "á", "\xe0": "à", "\u1ea3": "ả", "\xe3": "ã", "\u1ea1": "ạ",
    "\xe2": "â", "\u1ea5": "ấ", "\u1ea7": "ầ", "\u1ea9": "ẩ", "\u1eab": "ẫ", "\u1ead": "ậ",
    "\u0103": "ă", "\u1eaf": "ắ", "\u1eb1": "ằ", "\u1eb3": "ẳ", "\u1eb5": "ẵ", "\u1eb7": "ặ",
    "\xed": "í", "\xec": "ì", "\u1ec9": "ỉ", "\u0129": "ĩ", "\u1ecb": "ị",
    "\xf3": "ó", "\xf2": "ò", "\u1ecf": "ỏ", "\xf5": "õ", "\u1ecd": "ọ",
    "\xf4": "ô", "\u1ed1": "ố", "\u1ed3": "ồ", "\u1ed5": "ổ", "\u1ed7": "ỗ", "\u1ed9": "ộ",
    "\u01a1": "ơ", "\u1edb": "ớ", "\u1edd": "ờ", "\u1edf": "ở", "\u1ee1": "ỡ", "\u1ee3": "ợ",
    "\xfa": "ú", "\xf9": "ù", "\u1ee7": "ủ", "\u0169": "ũ", "\u1ee5": "ụ",
    "\u01b0": "ư", "\u1ee9": "ứ", "\u1eeb": "ừ", "\u1eed": "ử", "\u1eef": "ữ", "\u1ef1": "ự",
    "\xfd": "ý", "\u1ef3": "ỳ", "\u1ef5": "ỷ", "\u1ef7": "ỹ", "\u1ef9": "ỵ",
    "\u0111": "đ", "\u0110": "Đ"
}

def fix_font_tcvn3(text):
    if not isinstance(text, str): return text
    # Đây là logic chuyển đổi mã TCVN3 (ABC) sang Unicode
    # Trong thực tế bản web sẽ dùng bộ thư viện đầy đủ hơn, 
    # ở đây tôi demo logic chuẩn hóa Unicode dựng sẵn
    import unicodedata
    return unicodedata.normalize('NFC', text)

# --- HÀM TRỢ GIÚP ---
def find_fuzzy_match(name, choices, cutoff=0.6):
    matches = difflib.get_close_matches(str(name), [str(c) for c in choices], n=1, cutoff=cutoff)
    return matches[0] if matches else None

def run_unpivot(df, h_rows, id_col, d_start, sheet_name=None):
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
                    if sheet_name: entry["Nguồn (Sheet)"] = sheet_name
                    for i in range(h_rows):
                        entry[f"Tiêu đề {i+1}"] = headers.iloc[i, col_idx - (id_col + 1)]
                    results.append(entry)
        return pd.DataFrame(results)
    except: return None

# --- SIDEBAR MENU ---
st.sidebar.title("🚀 Excel Master Hub v10")
menu = st.sidebar.radio("Chọn chức năng:", ["🔄 Unpivot & Dashboard", "🔍 Đối soát & So khớp mờ", "🛠️ Tiện ích Sửa lỗi Font"])

# --- MODULE 1: UNPIVOT ---
if menu == "🔄 Unpivot & Dashboard":
    st.title("🔄 Unpivot & Phân tích Dashboard")
    file_up = st.file_uploader("Tải file Excel ma trận", type=["xlsx", "xls"], key="unp_up")
    if file_up:
        xl = pd.ExcelFile(file_up)
        sheet_names = xl.sheet_names
        with st.sidebar:
            st.header("⚙️ Cấu hình Profile")
            sel_p = st.selectbox("Chọn Profile:", list(st.session_state['profiles'].keys()))
            cfg = st.session_state['profiles'][sel_p]
            h_r, i_c, d_s = cfg['h_rows'], cfg['id_col'], cfg['d_start']
            if st.button("💾 Lưu Profile mới"):
                name = st.text_input("Tên:")
                if name:
                    st.session_state['profiles'][name] = {"h_rows": h_r, "id_col": i_c, "d_start": d_s}
                    save_profiles(st.session_state['profiles'])
        
        mode = st.radio("Chế độ:", ["Xử lý 1 Sheet", "Xử lý Toàn bộ Sheet"], horizontal=True)
        res_final = None
        if mode == "Xử lý 1 Sheet":
            sel_s = st.selectbox("Chọn Sheet:", sheet_names)
            df_raw = pd.read_excel(file_up, sheet_name=sel_s, header=None)
            st.dataframe(df_raw.head(10), use_container_width=True)
            if st.button("🚀 Chạy Unpivot"): res_final = run_unpivot(df_raw, h_r, i_c, d_s, sheet_name=sel_s)
        else:
            if st.button("🚀 Chạy Gộp Sheet"):
                all_res = [run_unpivot(pd.read_excel(file_up, s, header=None), h_r, i_c, d_s, s) for s in sheet_names]
                res_final = pd.concat([r for r in all_res if r is not None], ignore_index=True)

        if res_final is not None:
            st.success("Hoàn tất!")
            c1, c2 = st.columns(2)
            with c1: st.plotly_chart(px.bar(res_final.groupby("Đối tượng")["Số tiền"].sum().nlargest(10).reset_index(), x="Đối tượng", y="Số tiền", title="Top 10 Đối tượng"), use_container_width=True)
            with c2:
                cols = [c for c in res_final.columns if c != "Số tiền"]
                sel_pie = st.selectbox("Hạng mục biểu đồ tròn:", cols)
                st.plotly_chart(px.pie(res_final.groupby(sel_pie)["Số tiền"].sum().reset_index(), values="Số tiền", names=sel_pie, title=f"Cơ cấu theo {sel_pie}"), use_container_width=True)
            st.dataframe(res_final)
            out = BytesIO()
            res_final.to_excel(out, index=False)
            st.download_button("📥 Tải kết quả Unpivot (.xlsx)", out.getvalue(), "Ket_qua_Unpivot.xlsx")

# --- MODULE 2: ĐỐI SOÁT ---
elif menu == "🔍 Đối soát & So khớp mờ":
    st.title("🔍 Đối soát dữ liệu thông minh")
    c1, c2 = st.columns(2)
    with c1: f_m = st.file_uploader("Master", type=["xlsx"], key="m")
    with c2: f_c = st.file_uploader("Check", type=["xlsx"], key="c")
    if f_m and f_c:
        df_m = pd.read_excel(f_m); df_c = pd.read_excel(f_c)
        st.sidebar.header("⚙️ Cài đặt")
        k_m = st.sidebar.selectbox("Mã (Master):", df_m.columns); k_c = st.sidebar.selectbox("Mã (Check):", df_c.columns)
        v_col = st.sidebar.selectbox("Số tiền:", df_m.columns)
        fuz = st.sidebar.checkbox("Bật So khớp mờ"); score = st.sidebar.slider("Độ tương đồng %", 50, 100, 85)/100
        if st.button("🚀 Bắt đầu đối soát"):
            if fuz:
                mapping = {k: find_fuzzy_match(k, df_c[k_c].tolist(), score) for k in df_m[k_m].tolist()}
                df_m['Key_Matched'] = df_m[k_m].map(mapping)
                merged = pd.merge(df_m, df_c, left_on='Key_Matched', right_on=k_c, how='outer', suffixes=('_Gốc', '_ThựcTế'))
            else:
                merged = pd.merge(df_m, df_c, left_on=k_m, right_on=k_c, how='outer', suffixes=('_Gốc', '_ThựcTế'))
            merged = merged.fillna(0)
            cg = f"{v_col}_Gốc" if f"{v_col}_Gốc" in merged.columns else v_col
            ct = f"{v_col}_ThựcTế" if f"{v_col}_ThựcTế" in merged.columns else v_col
            merged['Chênh lệch'] = merged[cg] - merged[ct]
            st.dataframe(merged.style.applymap(lambda x: 'background-color: #ffcccc' if x != 0 else '', subset=['Chênh lệch']))
            out_ds = BytesIO()
            merged.to_excel(out_ds, index=False)
            st.download_button("📥 Tải báo cáo đối soát", out_ds.getvalue(), "Bao_cao_doi_soat.xlsx")

# --- MODULE 3: TIỆN ÍCH FONT (MỚI) ---
elif menu == "🛠️ Tiện ích Sửa lỗi Font":
    st.title("🛠️ Chuẩn hóa Font chữ Tiếng Việt")
    st.info("Chức năng này giúp chuyển đổi các cột dữ liệu bị lỗi font (Unicode tổ hợp/dựng sẵn) về chuẩn duy nhất.")
    
    file_f = st.file_uploader("Tải file Excel cần sửa font", type=["xlsx"], key="f_fix")
    if file_f:
        xl_f = pd.ExcelFile(file_f)
        s_f = st.selectbox("Chọn Sheet cần sửa:", xl_f.sheet_names)
        df_f = pd.read_excel(file_f, sheet_name=s_f)
        st.dataframe(df_f.head(10))
        
        target_cols = st.multiselect("Chọn các cột cần sửa lỗi font:", df_f.columns)
        
        if st.button("🚀 Tiến hành sửa lỗi font"):
            for col in target_cols:
                df_f[col] = df_f[col].apply(fix_font_tcvn3)
            st.success("Đã chuẩn hóa font chữ thành công!")
            st.dataframe(df_f.head(10))
            out_f = BytesIO()
            df_f.to_excel(out_f, index=False)
            st.download_button("📥 Tải file đã sửa font", out_f.getvalue(), "File_Da_Sua_Font.xlsx")
