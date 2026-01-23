import streamlit as st
import pandas as pd
from io import BytesIO
import json
import os
import plotly.express as px
import difflib
import unicodedata
import zipfile
import re

# --- 1. CẤU HÌNH GIAO DIỆN (BẢO TRÌ SIDEBAR XANH NHẠT & CSS) ---
st.set_page_config(page_title="Excel Hub Pro v23", layout="wide", page_icon="🚀")

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

# --- 2. HỆ THỐNG CỐT LÕI ---
CONFIG_FILE = "excel_profiles_v23.json"
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

# --- LOGIC UNPIVOT CHUẨN ---
def run_unpivot(df, h_rows, id_col, d_start, sheet_name=None):
    try:
        headers = df.iloc[0:h_rows, id_col + 1:]
        data_body = df.iloc[d_start - 1:, :]
        results = []
        for _, row in data_body.iterrows():
            id_val = str(row[id_col]).strip()
            if not id_val or id_val.lower() in ['nan', 'none', 'tổng cộng']: continue
            for col_idx in range(id_col + 1, len(df.columns)):
                val = pd.to_numeric(row[col_idx], errors='coerce')
                if pd.notnull(val) and val > 0:
                    entry = {"Đối tượng": id_val, "Số tiền": val}
                    if sheet_name: entry["Nguồn (Sheet)"] = sheet_name
                    for i in range(h_rows):
                        entry[f"Tiêu đề {i+1}"] = headers.iloc[i, col_idx - (id_col + 1)]
                    results.append(entry)
        return pd.DataFrame(results)
    except Exception as e:
        st.error(f"Lỗi Unpivot: {e}"); return None

def find_fuzzy_match(name, choices, cutoff=0.6):
    matches = difflib.get_close_matches(str(name), [str(c) for c in choices], n=1, cutoff=cutoff)
    return matches[0] if matches else None

def fix_vietnamese_font(text):
    if not isinstance(text, str): return text
    return unicodedata.normalize('NFC', text)

# --- 3. SIDEBAR MENU ---
with st.sidebar:
    st.title("🚀 Excel Master Hub")
    st.markdown("---")
    menu = st.sidebar.radio("Nghiệp vụ:", [
        "🔄 Unpivot & Dashboard", 
        "🔍 Đối soát & So khớp mờ", 
        "🛠️ Tiện ích Sửa lỗi Font",
        "📂 Tách File hàng loạt (ZIP)",
        "🔎 Trích xuất thông minh (Regex)"
    ])

# --- MODULE 1: UNPIVOT ---
if menu == "🔄 Unpivot & Dashboard":
    st.title("🔄 Unpivot & Dashboard Phân tích")
    with st.expander("📖 Hướng dẫn Unpivot", expanded=False):
        st.markdown("""
        1. Tải file Excel ma trận. Hệ thống sẽ tự gợi ý Profile nếu tên file khớp.
        2. Chỉnh cấu hình tại Sidebar (**Hàng tiêu đề, Cột tên, Dòng bắt đầu**).
        3. Nhấn **Chạy Unpivot** -> Dashboard KPI và biểu đồ sẽ hiện ra.
        4. Tải file kết quả tại nút phía dưới cùng.
        """)

    file_up = st.file_uploader("Tải file Excel ma trận", type=["xlsx", "xls"], key="unp")
    if file_up:
        xl = pd.ExcelFile(file_up); sheet_names = xl.sheet_names
        fname = file_up.name.lower(); p_list = list(st.session_state['profiles'].keys())
        d_idx = 0
        for i, p_n in enumerate(p_list):
            if p_n.lower() in fname: d_idx = i; st.sidebar.info(f"💡 Gợi ý Profile: {p_n}"); break
        
        with st.sidebar:
            st.header("⚙️ Cấu hình Unpivot")
            sel_p_cfg = st.selectbox("Sử dụng Profile:", p_list, index=d_idx)
            cfg = st.session_state['profiles'][sel_p_cfg]
            h_r = st.number_input("Số hàng tiêu đề:", value=cfg['h_rows'], min_value=0)
            i_c = st.number_input("Cột Tên (A=0, B=1...):", value=cfg['id_col'], min_value=0)
            d_s = st.number_input("Dòng bắt đầu data:", value=cfg['d_start'], min_value=1)
            if st.button("💾 Lưu Profile mới"):
                n_p = st.text_input("Tên:"); st.session_state['profiles'][n_p] = {"h_rows": h_r, "id_col": i_c, "d_start": d_s}
                save_profiles(st.session_state['profiles'])

        mode = st.radio("Chế độ:", ["Xử lý 1 Sheet", "Xử lý Toàn bộ Sheet"], horizontal=True)
        if mode == "Xử lý 1 Sheet":
            sel_s = st.selectbox("Chọn Sheet:", sheet_names)
            df_raw = pd.read_excel(file_up, sheet_name=sel_s, header=None)
            st.dataframe(df_raw.head(10), use_container_width=True)
            if st.button("🚀 Chạy Unpivot"): st.session_state['unpivot_result'] = run_unpivot(df_raw, h_r, i_c, d_s, sheet_name=sel_s)
        else:
            if st.button("🚀 Chạy Gộp tất cả"):
                all_res = [run_unpivot(pd.read_excel(file_up, s, header=None), h_r, i_c, d_s, s) for s in sheet_names]
                st.session_state['unpivot_result'] = pd.concat([r for r in all_res if r is not None], ignore_index=True)

        if st.session_state['unpivot_result'] is not None:
            res = st.session_state['unpivot_result']
            st.markdown(f'<div class="kpi-container"><div class="kpi-card"><h3>Tổng dòng</h3><h2>{len(res):,}</h2></div><div class="kpi-card"><h3>Tổng tiền</h3><h2>{res["Số tiền"].sum():,.0f}</h2></div><div class="kpi-card"><h3>Đối tượng</h3><h2>{res["Đối tượng"].nunique()}</h2></div></div>', unsafe_allow_html=True)
            c1, c2 = st.columns(2)
            with c1: st.plotly_chart(px.bar(res.groupby("Đối tượng")["Số tiền"].sum().nlargest(10).reset_index(), x="Đối tượng", y="Số tiền", title="Top 10 Đối tượng"), use_container_width=True)
            with c2: 
                sel_pie = st.selectbox("Vẽ biểu đồ tròn theo:", [c for c in res.columns if c != "Số tiền"])
                st.plotly_chart(px.pie(res, values="Số tiền", names=sel_pie, title=f"Cơ cấu theo {sel_pie}"), use_container_width=True)
            st.dataframe(res, use_container_width=True)
            out = BytesIO(); res.to_excel(out, index=False); st.download_button("📥 Tải kết quả", out.getvalue(), "Unpivot_Final.xlsx")

# --- MODULE 2: ĐỐI SOÁT ---
elif menu == "🔍 Đối soát & So khớp mờ":
    st.title("🔍 Đối soát & So khớp mờ")
    with st.expander("📖 Hướng dẫn Đối soát", expanded=False):
        st.markdown("""
        1. Tải file **Master** và file **Check**. Chọn sheet tương ứng để hiện Preview song song.
        2. Chọn cột **Key (Mã khóa)** chung giữa 2 bên.
        3. Bật **So khớp mờ** nếu tên gọi giữa 2 file không khớp nhau hoàn toàn.
        4. Nhấn **Đối soát** -> Hệ thống tự tính chênh lệch và bôi đỏ các dòng lỗi.
        """)
    col1, col2 = st.columns(2)
    df_m = df_c = None
    with col1:
        f_m = st.file_uploader("Master", type=["xlsx"], key="m")
        if f_m:
            xl_m = pd.ExcelFile(f_m); s_m = st.selectbox("Sheet Master:", xl_m.sheet_names)
            df_m = pd.read_excel(f_m, sheet_name=s_m); st.markdown("**Preview Master:**"); st.dataframe(df_m.head(10), use_container_width=True)
    with col2:
        f_c = st.file_uploader("Check", type=["xlsx"], key="c")
        if f_c:
            xl_c = pd.ExcelFile(f_c); s_c = st.selectbox("Sheet Check:", xl_c.sheet_names)
            df_c = pd.read_excel(f_c, sheet_name=s_c); st.markdown("**Preview Check:**"); st.dataframe(df_c.head(10), use_container_width=True)
    if df_m is not None and df_c is not None:
        k_m = st.sidebar.selectbox("Key Master:", df_m.columns); k_c = st.sidebar.selectbox("Key Check:", df_c.columns)
        v_col = st.sidebar.selectbox("Cột Tiền:", df_m.columns); fuz = st.sidebar.checkbox("Bật Fuzzy"); score = st.sidebar.slider("% Khớp", 50, 100, 85)/100
        if st.button("🚀 Đối soát"):
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
            out_ds = BytesIO(); merged.to_excel(out_ds, index=False); st.download_button("📥 Tải báo cáo", out_ds.getvalue(), "Doi_soat.xlsx")

# --- MODULE 3: FONT ---
elif menu == "🛠️ Tiện ích Sửa lỗi Font":
    st.title("🛠️ Sửa Font Unicode")
    with st.expander("📖 Hướng dẫn sửa Font", expanded=False):
        st.write("Dùng khi file bị lỗi hiển thị tiếng Việt. Tải file -> Chọn cột cần sửa -> Nhấn Sửa -> Tải file sạch.")
    f_f = st.file_uploader("Tải file", type=["xlsx"], key="font")
    if f_f:
        xl_f = pd.ExcelFile(f_f); s_f = st.selectbox("Chọn Sheet:", xl_f.sheet_names)
        df_f = pd.read_excel(f_f, sheet_name=s_f); target = st.multiselect("Chọn cột:", df_f.columns)
        if st.button("🚀 Sửa"):
            for c in target: df_f[c] = df_f[c].apply(fix_vietnamese_font)
            st.dataframe(df_f.head(10)); out = BytesIO(); df_f.to_excel(out, index=False); st.download_button("📥 Tải", out.getvalue(), "Fixed.xlsx")

# --- MODULE 4: ZIP ---
elif menu == "📂 Tách File hàng loạt (ZIP)":
    st.title("📂 Tách File ZIP")
    with st.expander("📖 Hướng dẫn Tách File", expanded=False):
        st.write("Tách 1 file tổng thành nhiều file nhỏ theo Cột chỉ định (ví dụ Tỉnh thành) và nén vào ZIP.")
    f_s = st.file_uploader("Tải file", type=["xlsx"], key="split")
    if f_s:
        xl_s = pd.ExcelFile(f_s); s_s = st.selectbox("Chọn Sheet:", xl_s.sheet_names)
        df_s = pd.read_excel(f_s, sheet_name=s_s); split_col = st.selectbox("Cột tách:", df_s.columns)
        if st.button("🚀 Tách"):
            vals = df_s[split_col].unique(); zip_buf = BytesIO()
            with zipfile.ZipFile(zip_buf, "a", zipfile.ZIP_DEFLATED, False) as zf:
                for v in vals:
                    df_v = df_s[df_s[split_col] == v]; buf = BytesIO(); df_v.to_excel(buf, index=False); zf.writestr(f"{v}.xlsx", buf.getvalue())
            st.download_button("📥 Tải ZIP", zip_buf.getvalue(), "Tach.zip")

# --- MODULE 5: REGEX ---
elif menu == "🔎 Trích xuất thông minh (Regex)":
    st.title("🔎 Trích xuất Regex")
    with st.expander("📖 Hướng dẫn Regex", expanded=False):
        st.write("Tự động 'nhặt' Số điện thoại, Email hoặc Mã số thuế từ các dòng ghi chú hỗn hợp.")
    f_r = st.file_uploader("Tải file", type=["xlsx"], key="reg")
    if f_r:
        df_r = pd.read_excel(f_r); target = st.selectbox("Cột văn bản:", df_r.columns)
        pat = st.text_input("Regex Pattern:", r"([a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+)")
        if st.button("🚀 Trích xuất"):
            df_r["Kết quả"] = df_r[target].apply(lambda x: ", ".join(re.findall(pat, str(x))))
            st.dataframe(df_r.head(10)); out = BytesIO(); df_r.to_excel(out, index=False); st.download_button("📥 Tải", out.getvalue(), "Regex.xlsx")
