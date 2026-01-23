import streamlit as st
import pandas as pd
from io import BytesIO
import json
import os
import plotly.express as px
import difflib
import unicodedata
import zipfile

# --- 1. CẤU HÌNH GIAO DIỆN (CSS CUSTOM) ---
st.set_page_config(page_title="Excel Hub Pro v15", layout="wide", page_icon="🚀")

def apply_custom_css():
    st.markdown("""
    <style>
    /* Tổng thể app */
    .stApp { background-color: #f1f3f6; }
    /* Tùy chỉnh Sidebar */
    [data-testid="stSidebar"] { background-color: #111827; color: white; }
    [data-testid="stSidebar"] * { color: white !important; }
    /* Container & Expander */
    div[data-testid="stExpander"] { border: none; box-shadow: 0 4px 12px rgba(0,0,0,0.05); border-radius: 12px; background: white; margin-bottom: 20px; }
    /* Nút bấm Pro */
    .stButton>button { border-radius: 12px; width: 100%; height: 45px; background-color: #2563eb; color: white; border: none; font-weight: 600; transition: 0.3s; }
    .stButton>button:hover { background-color: #1d4ed8; transform: translateY(-2px); box-shadow: 0 4px 12px rgba(37, 99, 235, 0.3); }
    /* KPI Cards */
    .kpi-container { display: flex; gap: 20px; margin-bottom: 25px; }
    .kpi-card { flex: 1; background: white; padding: 20px; border-radius: 15px; box-shadow: 0 4px 6px rgba(0,0,0,0.02); text-align: center; border-bottom: 4px solid #2563eb; }
    .kpi-card h3 { color: #64748b; font-size: 0.9rem; margin-bottom: 5px; }
    .kpi-card h2 { color: #1e293b; font-size: 1.8rem; margin: 0; }
    </style>
    """, unsafe_allow_html=True)

apply_custom_css()

# --- 2. HỆ THỐNG CỐT LÕI (GIỮ NGUYÊN) ---
CONFIG_FILE = "excel_profiles_v15.json"
def load_profiles():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f: return json.load(f)
        except: return {}
    return {"Mẫu SDH Mặc định": {"h_rows": 3, "id_col": 1, "d_start": 5}}

def save_profiles(profiles):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f: json.dump(profiles, f, ensure_ascii=False, indent=4)

if 'profiles' not in st.session_state: st.session_state['profiles'] = load_profiles()

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

# --- 3. SIDEBAR & MENU ---
with st.sidebar:
    st.title("🚀 Excel Master Hub")
    st.markdown("---")
    menu = st.radio("Chức năng hệ thống:", [
        "🔄 Unpivot & Dashboard", 
        "🔍 Đối soát & So khớp mờ", 
        "🛠️ Tiện ích Sửa lỗi Font",
        "📂 Tách File hàng loạt (ZIP)"
    ])

# --- MODULE 1: UNPIVOT & DASHBOARD ---
if menu == "🔄 Unpivot & Dashboard":
    st.title("🔄 Unpivot & Dashboard Phân tích")
    with st.expander("📖 Hướng dẫn sử dụng", expanded=False):
        st.write("Tải file -> Cấu hình Profile tại Sidebar -> Chọn chế độ xử lý -> Tải file kết quả.")

    file_up = st.file_uploader("Tải file Excel ma trận", type=["xlsx", "xls"], key="unp")
    if file_up:
        xl = pd.ExcelFile(file_up); sheet_names = xl.sheet_names
        with st.sidebar:
            st.header("⚙️ Cấu hình Profile")
            sel_p = st.selectbox("Chọn Profile:", list(st.session_state['profiles'].keys()))
            cfg = st.session_state['profiles'][sel_p]
            h_r, i_c, d_s = cfg['h_rows'], cfg['id_col'], cfg['d_start']
            if st.button("💾 Lưu cấu hình hiện tại"):
                n_p = st.text_input("Tên Profile mới:")
                if n_p:
                    st.session_state['profiles'][n_p] = {"h_rows": h_r, "id_col": i_c, "d_start": d_s}
                    save_profiles(st.session_state['profiles'])

        mode = st.radio("Chế độ:", ["Xử lý 1 Sheet (Preview)", "Xử lý Toàn bộ Sheet (Gộp)"], horizontal=True)
        res_final = None
        if mode == "Xử lý 1 Sheet (Preview)":
            sel_s = st.selectbox("Chọn Sheet:", sheet_names)
            df_raw = pd.read_excel(file_up, sheet_name=sel_s, header=None)
            st.dataframe(df_raw.head(10), use_container_width=True)
            if st.button("🚀 Thực hiện Unpivot"): res_final = run_unpivot(df_raw, h_r, i_c, d_s, sheet_name=sel_s)
        else:
            if st.button("🚀 Bắt đầu gộp tất cả"):
                all_res = [run_unpivot(pd.read_excel(file_up, s, header=None), h_r, i_c, d_s, s) for s in sheet_names]
                res_final = pd.concat([r for r in all_res if r is not None], ignore_index=True)

        if res_final is not None:
            # KPI Cards
            st.markdown(f"""
            <div class="kpi-container">
                <div class="kpi-card"><h3>Tổng dòng</h3><h2>{len(res_final):,}</h2></div>
                <div class="kpi-card"><h3>Tổng tiền</h3><h2>{res_final['Số tiền'].sum():,.0f}</h2></div>
                <div class="kpi-card"><h3>Đối tượng</h3><h2>{res_final['Đối tượng'].nunique()}</h2></div>
            </div>
            """, unsafe_allow_html=True)
            
            c1, c2 = st.columns(2)
            with c1: st.plotly_chart(px.bar(res_final.groupby("Đối tượng")["Số tiền"].sum().nlargest(10).reset_index(), x="Đối tượng", y="Số tiền", title="Top 10 Đối tượng"), use_container_width=True)
            with c2: 
                sel_p = st.selectbox("Vẽ biểu đồ tròn theo:", [c for c in res_final.columns if c != "Số tiền"])
                st.plotly_chart(px.pie(res_final, values="Số tiền", names=sel_p, title=f"Cơ cấu theo {sel_pie}"), use_container_width=True)
            
            st.dataframe(res_final, use_container_width=True)
            out = BytesIO(); res_final.to_excel(out, index=False)
            st.download_button("📥 Tải kết quả Unpivot (.xlsx)", out.getvalue(), "Unpivot_Final.xlsx")

# --- MODULE 2: ĐỐI SOÁT (BẢO TỒN PREVIEW) ---
elif menu == "🔍 Đối soát & So khớp mờ":
    st.title("🔍 Đối soát dữ liệu thông minh")
    with st.expander("📖 Hướng dẫn Đối soát", expanded=False):
        st.write("Tải file Master & Check -> Chọn Key -> Bật Fuzzy nếu tên không khớp 100% -> Chạy.")
    
    col1, col2 = st.columns(2)
    df_m = df_c = None
    with col1:
        f_m = st.file_uploader("File Master", type=["xlsx"], key="m")
        if f_m:
            xl_m = pd.ExcelFile(f_m); s_m = st.selectbox("Chọn Sheet Master:", xl_m.sheet_names)
            df_m = pd.read_excel(f_m, sheet_name=s_m); st.dataframe(df_m.head(10), use_container_width=True)
    with col2:
        f_c = st.file_uploader("File Đối soát", type=["xlsx"], key="c")
        if f_c:
            xl_c = pd.ExcelFile(f_c); s_c = st.selectbox("Chọn Sheet Check:", xl_c.sheet_names)
            df_c = pd.read_excel(f_c, sheet_name=s_c); st.dataframe(df_c.head(10), use_container_width=True)

    if df_m is not None and df_c is not None:
        st.sidebar.header("⚙️ Cài đặt")
        k_m = st.sidebar.selectbox("Key (Master):", df_m.columns); k_c = st.sidebar.selectbox("Key (Check):", df_c.columns)
        v_col = st.sidebar.selectbox("Cột Tiền:", df_m.columns); fuz = st.sidebar.checkbox("Bật So khớp mờ"); score = st.sidebar.slider("Độ tương đồng %", 50, 100, 85)/100
        if st.button("🚀 Thực hiện đối soát"):
            if fuz:
                mapping = {k: find_fuzzy_match(k, df_c[k_c].tolist(), score) for k in df_m[k_m].tolist()}
                df_m['Key_Matched'] = df_m[k_m].map(mapping)
                merged = pd.merge(df_m, df_c, left_on='Key_Matched', right_on=k_c, how='outer', suffixes=('_Gốc', '_ThựcTế'))
            else:
                merged = pd.merge(df_m, df_c, left_on=k_m, right_on=k_c, how='outer', suffixes=('_Gốc', '_ThựcTế'))
            merged = merged.fillna(0); cg, ct = f"{v_col}_Gốc", f"{v_col}_ThựcTế"
            if cg not in merged.columns: cg, ct = v_col, v_col
            merged['Chênh lệch'] = merged[cg] - merged[ct]
            st.dataframe(merged.style.applymap(lambda x: 'background-color: #ffcccc' if x != 0 else '', subset=['Chênh lệch']), use_container_width=True)
            out_ds = BytesIO(); merged.to_excel(out_ds, index=False)
            st.download_button("📥 Tải báo cáo đối soát", out_ds.getvalue(), "Bao_cao_doi_soat.xlsx")

# --- CÁC MODULE KHÁC (BẢO TỒN 100%) ---
elif menu == "🛠️ Tiện ích Sửa lỗi Font":
    st.title("🛠️ Chuẩn hóa Font Tiếng Việt")
    f_f = st.file_uploader("Tải file", type=["xlsx"], key="font")
    if f_f:
        df_f = pd.read_excel(f_f); st.dataframe(df_f.head(10)); target = st.multiselect("Chọn cột:", df_f.columns)
        if st.button("🚀 Chạy chuẩn hóa"):
            for c in target: df_f[c] = df_f[c].apply(fix_vietnamese_font)
            st.success("Xong!"); st.dataframe(df_f.head(10))
            out_f = BytesIO(); df_f.to_excel(out_f, index=False)
            st.download_button("📥 Tải file sạch", out_f.getvalue(), "Fixed_Font.xlsx")

elif menu == "📂 Tách File hàng loạt (ZIP)":
    st.title("📂 Tách File ZIP hàng loạt")
    f_s = st.file_uploader("Tải file", type=["xlsx"], key="split")
    if f_s:
        df_s = pd.read_excel(f_s); st.dataframe(df_s.head(10)); split_col = st.selectbox("Chọn cột tách:", df_s.columns)
        if st.button("🚀 Bắt đầu tách"):
            vals = df_s[split_col].unique(); zip_buf = BytesIO()
            with zipfile.ZipFile(zip_buf, "a", zipfile.ZIP_DEFLATED, False) as zf:
                for v in vals:
                    df_v = df_s[df_s[split_col] == v]; buf = BytesIO(); df_v.to_excel(buf, index=False)
                    zf.writestr(f"{v}.xlsx", buf.getvalue())
            st.success("Hoàn tất!"); st.download_button("📥 Tải ZIP", zip_buf.getvalue(), "Files_Tach.zip")
