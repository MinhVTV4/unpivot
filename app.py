import streamlit as st
import pandas as pd
from io import BytesIO
import json
import os
import plotly.express as px
import difflib
import unicodedata
import zipfile
import re # Thư viện cho Regex

# --- 1. CẤU HÌNH GIAO DIỆN (BẢO TRÌ SIDEBAR XANH NHẠT) ---
st.set_page_config(page_title="Excel Hub Pro v21", layout="wide", page_icon="🚀")

def apply_custom_css():
    st.markdown("""
    <style>
    .stApp { background-color: #f8fafc; }
    [data-testid="stSidebar"] { background-color: #e0f2fe; border-right: 1px solid #bae6fd; }
    [data-testid="stSidebar"] * { color: #0369a1 !important; }
    div[data-testid="stExpander"] { border: none; box-shadow: 0 4px 12px rgba(0,0,0,0.05); border-radius: 12px; background: white; margin-bottom: 20px; }
    .stButton>button { border-radius: 12px; width: 100%; height: 45px; background-color: #0284c7; color: white; border: none; font-weight: 600; transition: 0.3s; }
    .stButton>button:hover { background-color: #0369a1; transform: translateY(-2px); box-shadow: 0 4px 12px rgba(2, 132, 199, 0.3); }
    .kpi-container { display: flex; gap: 20px; margin-bottom: 25px; }
    .kpi-card { flex: 1; background: white; padding: 20px; border-radius: 15px; box-shadow: 0 4px 6px rgba(0,0,0,0.02); text-align: center; border-bottom: 4px solid #0284c7; }
    .kpi-card h3 { color: #64748b; font-size: 0.9rem; margin-bottom: 5px; }
    .kpi-card h2 { color: #0c4a6e; font-size: 1.8rem; margin: 0; }
    </style>
    """, unsafe_allow_html=True)

apply_custom_css()

# --- 2. HỆ THỐNG CỐT LÕI (GIỮ NGUYÊN) ---
CONFIG_FILE = "excel_profiles_v21.json"
def load_profiles():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f: return json.load(f)
        except: return {}
    return {"Mẫu SDH Mặc định": {"h_rows": 3, "id_col": 1, "d_start": 5}}

def save_profiles(profiles):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f: json.dump(profiles, f, ensure_ascii=False, indent=4)

if 'profiles' not in st.session_state: st.session_state['profiles'] = load_profiles()
if 'unpivot_result' not in st.session_state: st.session_state['unpivot_result'] = None

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

# --- 3. SIDEBAR MENU (BẢO TỒN 100%) ---
with st.sidebar:
    st.title("🚀 Excel Master Hub")
    st.markdown("---")
    menu = st.sidebar.radio("Nghiệp vụ:", [
        "🔄 Unpivot & Dashboard", 
        "🔍 Đối soát & So khớp mờ", 
        "🛠️ Tiện ích Sửa lỗi Font",
        "📂 Tách File hàng loạt (ZIP)",
        "🔎 Trích xuất thông minh (Regex)" # MỚI
    ])

# --- MODULE 1: UNPIVOT & DASHBOARD (BẢO TỒN) ---
if menu == "🔄 Unpivot & Dashboard":
    st.title("🔄 Unpivot & Dashboard")
    file_up = st.file_uploader("Tải file Excel", type=["xlsx", "xls"], key="unp")
    if file_up:
        xl = pd.ExcelFile(file_up); sheet_names = xl.sheet_names
        # Gợi ý Profile (v20)
        fname = file_up.name.lower(); p_list = list(st.session_state['profiles'].keys())
        d_idx = 0
        for i, p_n in enumerate(p_list):
            if p_n.lower() in fname: d_idx = i; st.sidebar.info(f"💡 Gợi ý: {p_n}"); break
        with st.sidebar:
            st.header("⚙️ Cấu hình"); sel_p = st.selectbox("Sử dụng Profile:", p_list, index=d_idx)
            cfg = st.session_state['profiles'][sel_p]
            h_r = st.number_input("Hàng tiêu đề:", value=cfg['h_rows'])
            i_c = st.number_input("Cột Tên:", value=cfg['id_col'])
            d_s = st.number_input("Dòng bắt đầu:", value=cfg['d_start'])

        mode = st.radio("Chế độ:", ["Xử lý 1 Sheet", "Xử lý Toàn bộ"], horizontal=True)
        if mode == "Xử lý 1 Sheet":
            s_s = st.selectbox("Chọn Sheet:", sheet_names); df_r = pd.read_excel(file_up, sheet_name=s_s, header=None)
            st.dataframe(df_r.head(5), use_container_width=True)
            if st.button("🚀 Chạy"): st.session_state['unpivot_result'] = run_unpivot(df_r, h_r, i_c, d_s, s_s)
        else:
            if st.button("🚀 Chạy Gộp"):
                all_r = [run_unpivot(pd.read_excel(file_up, s, header=None), h_r, i_c, d_s, s) for s in sheet_names]
                st.session_state['unpivot_result'] = pd.concat([r for r in all_r if r is not None], ignore_index=True)

        if st.session_state['unpivot_result'] is not None:
            res = st.session_state['unpivot_result']
            st.markdown(f'<div class="kpi-container"><div class="kpi-card"><h3>Tổng dòng</h3><h2>{len(res):,}</h2></div><div class="kpi-card"><h3>Tổng tiền</h3><h2>{res["Số tiền"].sum():,.0f}</h2></div></div>', unsafe_allow_html=True)
            sel_pie = st.selectbox("Hạng mục biểu đồ:", [c for c in res.columns if c != "Số tiền"])
            st.plotly_chart(px.pie(res, values="Số tiền", names=sel_pie, title="Cơ cấu"), use_container_width=True)
            st.dataframe(res, use_container_width=True)
            out = BytesIO(); res.to_excel(out, index=False); st.download_button("📥 Tải về", out.getvalue(), "Result.xlsx")

# --- MODULE 2 & 3 & 4 (BẢO TỒN 100%) ---
# (Phần này giữ nguyên logic Preview song song, So khớp mờ, ZIP, Font như v20)
elif menu == "🔍 Đối soát & So khớp mờ":
    st.title("🔍 Đối soát dữ liệu")
    col1, col2 = st.columns(2)
    df_m = df_c = None
    with col1:
        f_m = st.file_uploader("File Master", type=["xlsx"], key="m")
        if f_m:
            xl_m = pd.ExcelFile(f_m); s_m = st.selectbox("Sheet Master:", xl_m.sheet_names)
            df_m = pd.read_excel(f_m, sheet_name=s_m); st.dataframe(df_m.head(10), use_container_width=True)
    with col2:
        f_c = st.file_uploader("File Check", type=["xlsx"], key="c")
        if f_c:
            xl_c = pd.ExcelFile(f_c); s_c = st.selectbox("Sheet Check:", xl_c.sheet_names)
            df_c = pd.read_excel(f_c, sheet_name=s_c); st.dataframe(df_c.head(10), use_container_width=True)
    if df_m is not None and df_c is not None:
        k_m = st.sidebar.selectbox("Key Master:", df_m.columns); k_c = st.sidebar.selectbox("Key Check:", df_c.columns)
        v_col = st.sidebar.selectbox("Số tiền:", df_m.columns); fuz = st.sidebar.checkbox("Bật Fuzzy")
        if st.button("🚀 Đối soát"):
            if fuz:
                mapping = {k: find_fuzzy_match(k, df_c[k_c].tolist(), 0.85) for k in df_m[k_m].tolist()}
                df_m['Key_Matched'] = df_m[k_m].map(mapping)
                merged = pd.merge(df_m, df_c, left_on='Key_Matched', right_on=k_c, how='outer', suffixes=('_Gốc', '_ThựcTế'))
            else:
                merged = pd.merge(df_m, df_c, left_on=k_m, right_on=k_c, how='outer', suffixes=('_Gốc', '_ThựcTế'))
            merged = merged.fillna(0); cg, ct = f"{v_col}_Gốc", f"{v_col}_ThựcTế"
            if cg not in merged.columns: cg, ct = v_col, v_col
            merged['Chênh lệch'] = merged[cg] - merged[ct]
            st.dataframe(merged.style.applymap(lambda x: 'background-color: #ffcccc' if x != 0 else '', subset=['Chênh lệch']), use_container_width=True)

elif menu == "🛠️ Tiện ích Sửa lỗi Font":
    st.title("🛠️ Sửa lỗi Font")
    f_f = st.file_uploader("Tải file", type=["xlsx"], key="font")
    if f_f:
        df_f = pd.read_excel(f_f); target = st.multiselect("Chọn cột:", df_f.columns)
        if st.button("🚀 Chạy"):
            for c in target: df_f[c] = df_f[c].apply(fix_vietnamese_font)
            st.dataframe(df_f.head(10)); out = BytesIO(); df_f.to_excel(out, index=False); st.download_button("📥 Tải", out.getvalue(), "Fixed.xlsx")

elif menu == "📂 Tách File hàng loạt (ZIP)":
    st.title("📂 Tách File ZIP")
    f_s = st.file_uploader("Tải file", type=["xlsx"], key="split")
    if f_s:
        df_s = pd.read_excel(f_s); split_col = st.selectbox("Chọn cột tách:", df_s.columns)
        if st.button("🚀 Tách"):
            vals = df_s[split_col].unique(); zip_buf = BytesIO()
            with zipfile.ZipFile(zip_buf, "a", zipfile.ZIP_DEFLATED, False) as zf:
                for v in vals:
                    df_v = df_s[df_s[split_col] == v]; buf = BytesIO(); df_v.to_excel(buf, index=False); zf.writestr(f"{v}.xlsx", buf.getvalue())
            st.download_button("📥 Tải ZIP", zip_buf.getvalue(), "Files_Tach.zip")

# --- MODULE 5: TRÍCH XUẤT THÔNG MINH (REGEX - MỚI) ---
elif menu == "🔎 Trích xuất thông minh (Regex)":
    st.title("🔎 Trích xuất dữ liệu từ chuỗi văn bản")
    st.info("Ví dụ: Tách số điện thoại, Email hoặc Mã số thuế từ một cột ghi chú hỗn hợp.")
    
    file_reg = st.file_uploader("Tải file Excel cần trích xuất", type=["xlsx"], key="reg_up")
    if file_reg:
        df_reg = pd.read_excel(file_reg)
        st.subheader("📋 Preview dữ liệu")
        st.dataframe(df_reg.head(10), use_container_width=True)
        
        target_col = st.selectbox("Chọn cột chứa văn bản hỗn hợp:", df_reg.columns)
        ext_type = st.radio("Bạn muốn trích xuất gì?", ["Số điện thoại", "Email", "Mã số thuế (10-13 số)", "Tự định nghĩa (Regex)"], horizontal=True)
        
        regex_pattern = ""
        if ext_type == "Số điện thoại": regex_pattern = r"(0[3|5|7|8|9][0-9]{8})"
        elif ext_type == "Email": regex_pattern = r"([a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+)"
        elif ext_type == "Mã số thuế (10-13 số)": regex_pattern = r"([0-9]{10,13})"
        else: regex_pattern = st.text_input("Nhập biểu thức Regex của bạn:", r"")

        if st.button("🚀 Bắt đầu trích xuất", type="primary"):
            if regex_pattern:
                def extract_info(text):
                    matches = re.findall(regex_pattern, str(text))
                    return ", ".join(matches) if matches else ""
                
                df_reg[f"Trích xuất_{ext_type}"] = df_reg[target_col].apply(extract_info)
                st.success("Đã trích xuất xong!")
                st.dataframe(df_reg.head(15), use_container_width=True)
                
                out_reg = BytesIO()
                df_reg.to_excel(out_reg, index=False)
                st.download_button("📥 Tải file kết quả trích xuất", out_reg.getvalue(), "Extracted_Data.xlsx")
