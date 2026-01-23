import streamlit as st
import pandas as pd
from io import BytesIO
import json
import os
import plotly.express as px
import difflib
import unicodedata
import zipfile
import requests
from streamlit_lottie import st_lottie # Thêm hoạt ảnh sinh động

# --- 1. CẤU HÌNH GIAO DIỆN NÂNG CAO (CSS) ---
st.set_page_config(page_title="Excel Hub Pro v14", layout="wide", page_icon="🚀")

def local_css():
    st.markdown("""
    <style>
    /* Bo góc và đổ bóng cho các container */
    .stApp { background-color: #f8f9fa; }
    div[data-testid="stExpander"] { border: none; box-shadow: 0 4px 6px rgba(0,0,0,0.05); border-radius: 10px; background: white; }
    .stButton>button { border-radius: 20px; width: 100%; transition: all 0.3s; border: none; background-color: #4CAF50; color: white; font-weight: bold; }
    .stButton>button:hover { transform: scale(1.02); box-shadow: 0 4px 15px rgba(0,0,0,0.1); }
    /* Tùy chỉnh Sidebar */
    [data-testid="stSidebar"] { background-color: #1e293b; color: white; }
    [data-testid="stSidebar"] * { color: white !important; }
    /* Thẻ chỉ số KPI */
    .metric-card { background: white; padding: 20px; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); text-align: center; border-top: 4px solid #4CAF50; }
    </style>
    """, unsafe_allow_html=True)

local_css()

# --- HÀM TẢI HOẠT ẢNH ---
def load_lottieurl(url):
    r = requests.get(url)
    if r.status_code != 200: return None
    return r.json()

lottie_excel = load_lottieurl("https://assets5.lottiefiles.com/packages/lf20_S69f4D.json")

# --- CẤU HÌNH HỆ THỐNG CŨ (GIỮ NGUYÊN) ---
CONFIG_FILE = "excel_profiles_v14.json"
def load_profiles():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f: return json.load(f)
        except: return {}
    return {"Mẫu SDH Mặc định": {"h_rows": 3, "id_col": 1, "d_start": 5}}

def save_profiles(profiles):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f: json.dump(profiles, f, ensure_ascii=False, indent=4)

if 'profiles' not in st.session_state: st.session_state['profiles'] = load_profiles()

# --- CÁC HÀM LOGIC CŨ (GIỮ NGUYÊN 100%) ---
def find_fuzzy_match(name, choices, cutoff=0.6):
    matches = difflib.get_close_matches(str(name), [str(c) for c in choices], n=1, cutoff=cutoff)
    return matches[0] if matches else None

def fix_vietnamese_font(text):
    if not isinstance(text, str): return text
    return unicodedata.normalize('NFC', text)

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
                    for i in range(h_rows): entry[f"Tiêu đề {i+1}"] = headers.iloc[i, col_idx - (id_col + 1)]
                    results.append(entry)
        return pd.DataFrame(results)
    except: return None

# --- SIDEBAR MENU ---
with st.sidebar:
    st_lottie(lottie_excel, height=150, key="logo") # Thêm logo hoạt hình
    st.title("Excel Master Hub")
    menu = st.radio("Chức năng:", [
        "🔄 Unpivot & Dashboard", 
        "🔍 Đối soát & So khớp mờ", 
        "🛠️ Tiện ích Sửa lỗi Font",
        "📂 Tách File hàng loạt (ZIP)"
    ])

# --- MODULE 1: UNPIVOT & DASHBOARD ---
if menu == "🔄 Unpivot & Dashboard":
    st.title("🔄 Unpivot & Dashboard Phân tích")
    with st.expander("📖 Hướng dẫn sử dụng", expanded=False):
        st.write("B1: Tải file -> B2: Cấu hình Profile -> B3: Nhấn Chạy -> B4: Tải kết quả.")

    file_up = st.file_uploader("Tải file Excel ma trận", type=["xlsx", "xls"], key="unp_up")
    if file_up:
        xl = pd.ExcelFile(file_up); sheet_names = xl.sheet_names
        with st.sidebar:
            st.header("⚙️ Cấu hình")
            sel_p = st.selectbox("Chọn Profile:", list(st.session_state['profiles'].keys()))
            cfg = st.session_state['profiles'][sel_p]
            h_r, i_c, d_s = cfg['h_rows'], cfg['id_col'], cfg['d_start']
            if st.button("💾 Lưu Profile"):
                new_p = st.text_input("Tên:")
                if new_p:
                    st.session_state['profiles'][new_p] = {"h_rows": h_r, "id_col": i_c, "d_start": d_s}
                    save_profiles(st.session_state['profiles'])

        mode = st.radio("Chế độ:", ["Xử lý 1 Sheet", "Xử lý TOÀN BỘ Sheet"], horizontal=True)
        res_final = None
        if mode == "Xử lý 1 Sheet":
            sel_s = st.selectbox("Chọn Sheet:", sheet_names)
            df_raw = pd.read_excel(file_up, sheet_name=sel_s, header=None)
            st.dataframe(df_raw.head(10), use_container_width=True)
            if st.button("🚀 Chạy Unpivot"): res_final = run_unpivot(df_raw, h_r, i_c, d_s, sheet_name=sel_s)
        else:
            if st.button("🚀 Chạy Gộp tất cả"):
                all_res = [run_unpivot(pd.read_excel(file_up, s, header=None), h_r, i_c, d_s, s) for s in sheet_names]
                res_final = pd.concat([r for r in all_res if r is not None], ignore_index=True)

        if res_final is not None:
            # HIỂN THỊ KPI SINH ĐỘNG
            st.markdown("---")
            k1, k2, k3 = st.columns(3)
            with k1: st.markdown(f'<div class="metric-card"><h3>Tổng dòng</h3><h2>{len(res_final):,}</h2></div>', unsafe_allow_html=True)
            with k2: st.markdown(f'<div class="metric-card"><h3>Tổng tiền</h3><h2>{res_final["Số tiền"].sum():,.0f}</h2></div>', unsafe_allow_html=True)
            with k3: st.markdown(f'<div class="metric-card"><h3>Đối tượng</h3><h2>{res_final["Đối tượng"].nunique()}</h2></div>', unsafe_allow_html=True)

            c1, c2 = st.columns(2)
            with c1: st.plotly_chart(px.bar(res_final.groupby("Đối tượng")["Số tiền"].sum().nlargest(10).reset_index(), x="Đối tượng", y="Số tiền", title="Top 10 Đối tượng"), use_container_width=True)
            with c2:
                sel_pie = st.selectbox("Hạng mục tròn:", [c for c in res_final.columns if c != "Số tiền"])
                st.plotly_chart(px.pie(res_final, values="Số tiền", names=sel_pie, title="Cơ cấu"), use_container_width=True)
            
            st.dataframe(res_final, use_container_width=True)
            out = BytesIO(); res_final.to_excel(out, index=False)
            st.download_button("📥 Tải kết quả Unpivot", out.getvalue(), "Unpivot_Final.xlsx")

# --- MODULE 2: ĐỐI SOÁT (BẢO TỒN PREVIEW) ---
elif menu == "🔍 Đối soát & So khớp mờ":
    st.title("🔍 Đối soát dữ liệu")
    with st.expander("📖 Hướng dẫn"): st.write("Tải 2 file -> Chọn Key -> Bật Fuzzy nếu cần -> Chạy.")
    col1, col2 = st.columns(2)
    df_m = df_c = None
    with col1:
        f_m = st.file_uploader("Master", type=["xlsx"], key="m")
        if f_m:
            xl_m = pd.ExcelFile(f_m); s_m = st.selectbox("Sheet Master:", xl_m.sheet_names)
            df_m = pd.read_excel(f_m, sheet_name=s_m)
            st.dataframe(df_m.head(10), use_container_width=True)
    with col2:
        f_c = st.file_uploader("Check", type=["xlsx"], key="c")
        if f_c:
            xl_c = pd.ExcelFile(f_c); s_c = st.selectbox("Sheet Check:", xl_c.sheet_names)
            df_c = pd.read_excel(f_c, sheet_name=s_c)
            st.dataframe(df_c.head(10), use_container_width=True)

    if df_m is not None and df_c is not None:
        st.sidebar.header("⚙️ Cài đặt")
        k_m = st.sidebar.selectbox("Key (Master):", df_m.columns); k_c = st.sidebar.selectbox("Key (Check):", df_c.columns)
        v_col = st.sidebar.selectbox("Tiền:", df_m.columns); fuz = st.sidebar.checkbox("Bật Fuzzy"); score = st.sidebar.slider("Độ tương đồng %", 50, 100, 85)/100
        if st.button("🚀 Chạy Đối soát"):
            if fuz:
                mapping = {k: find_fuzzy_match(k, df_c[k_c].tolist(), score) for k in df_m[k_m].tolist()}
                df_m['Key_Matched'] = df_m[k_m].map(mapping)
                merged = pd.merge(df_m, df_c, left_on='Key_Matched', right_on=k_c, how='outer', suffixes=('_Gốc', '_ThựcTế'))
            else:
                merged = pd.merge(df_m, df_c, left_on=k_m, right_on=k_c, how='outer', suffixes=('_Gốc', '_ThựcTế'))
            merged = merged.fillna(0); cg = f"{v_col}_Gốc" if f"{v_col}_Gốc" in merged.columns else v_col; ct = f"{v_col}_ThựcTế" if f"{v_col}_ThựcTế" in merged.columns else v_col
            merged['Chênh lệch'] = merged[cg] - merged[ct]
            st.dataframe(merged.style.applymap(lambda x: 'background-color: #ffcccc' if x != 0 else '', subset=['Chênh lệch']), use_container_width=True)
            out_ds = BytesIO(); merged.to_excel(out_ds, index=False)
            st.download_button("📥 Tải báo cáo đối soát", out_ds.getvalue(), "Bao_cao.xlsx")

# --- CÁC MODULE KHÁC: FONT & TÁCH FILE (BẢO TỒN 100%) ---
elif menu == "🛠️ Tiện ích Sửa lỗi Font":
    st.title("🛠️ Sửa lỗi Font")
    file_f = st.file_uploader("Tải file", type=["xlsx"], key="font")
    if file_f:
        df_f = pd.read_excel(file_f); st.dataframe(df_f.head(10))
        target_cols = st.multiselect("Chọn cột:", df_f.columns)
        if st.button("🚀 Sửa font"):
            for col in target_cols: df_f[col] = df_f[col].apply(fix_vietnamese_font)
            st.success("Xong!"); st.dataframe(df_f.head(10))
            out_f = BytesIO(); df_f.to_excel(out_f, index=False)
            st.download_button("📥 Tải file", out_f.getvalue(), "Fixed.xlsx")

elif menu == "📂 Tách File hàng loạt (ZIP)":
    st.title("📂 Tách File ZIP")
    file_split = st.file_uploader("Tải file", type=["xlsx"], key="split")
    if file_split:
        df_s = pd.read_excel(file_split); st.dataframe(df_s.head(10))
        split_col = st.selectbox("Chọn cột tách:", df_s.columns)
        if st.button("🚀 Tách & ZIP"):
            unique_vals = df_s[split_col].unique(); zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zf:
                for val in unique_vals:
                    df_filtered = df_s[df_s[split_col] == val]; sub = BytesIO(); df_filtered.to_excel(sub, index=False)
                    zf.writestr(f"{val}.xlsx", sub.getvalue())
            st.success("Xong!"); st.download_button("📥 Tải ZIP", zip_buffer.getvalue(), "Tach.zip")
