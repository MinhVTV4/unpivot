import streamlit as st
import pandas as pd
from io import BytesIO
import json
import os
import plotly.express as px
import difflib
import unicodedata
import zipfile # Thư viện mới để nén file

# --- CẤU HÌNH HỆ THỐNG ---
st.set_page_config(page_title="Excel Hub Pro v12", layout="wide", page_icon="🚀")

CONFIG_FILE = "excel_profiles_v12.json"

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

# --- CÁC HÀM BỔ TRỢ ---
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
                    for i in range(h_rows):
                        entry[f"Tiêu đề {i+1}"] = headers.iloc[i, col_idx - (id_col + 1)]
                    results.append(entry)
        return pd.DataFrame(results)
    except: return None

# --- SIDEBAR MENU (KHÔNG BỎ CÁI CŨ) ---
st.sidebar.title("🎮 Excel Master Hub v12")
menu = st.sidebar.radio("Chọn nghiệp vụ:", [
    "🔄 Unpivot & Dashboard", 
    "🔍 Đối soát & So khớp mờ", 
    "🛠️ Tiện ích Sửa lỗi Font",
    "📂 Tách File hàng loạt (ZIP)" # TÍNH NĂNG MỚI
])

# --- MODULE 1: UNPIVOT & DASHBOARD (GIỮ NGUYÊN 100%) ---
if menu == "🔄 Unpivot & Dashboard":
    st.title("🔄 Unpivot Ma trận & Phân tích Dashboard")
    file_up = st.file_uploader("1. Tải file Excel ma trận", type=["xlsx", "xls"], key="unp_up")
    if file_up:
        xl = pd.ExcelFile(file_up)
        sheet_names = xl.sheet_names
        with st.sidebar:
            st.header("⚙️ Profile cấu hình")
            p_names = list(st.session_state['profiles'].keys())
            sel_p = st.selectbox("Chọn Profile:", p_names)
            cfg = st.session_state['profiles'][sel_p]
            h_r, i_c, d_s = cfg['h_rows'], cfg['id_col'], cfg['d_start']
            if st.button("💾 Lưu cấu hình hiện tại"):
                new_p = st.text_input("Tên Profile:")
                if new_p:
                    st.session_state['profiles'][new_p] = {"h_rows": h_r, "id_col": i_c, "d_start": d_s}
                    save_profiles(st.session_state['profiles'])

        mode = st.radio("Chế độ:", ["Xử lý 1 Sheet", "Xử lý TOÀN BỘ Sheet"], horizontal=True)
        res_final = None
        if mode == "Xử lý 1 Sheet":
            sel_s = st.selectbox("Chọn Sheet:", sheet_names)
            df_raw = pd.read_excel(file_up, sheet_name=sel_s, header=None)
            st.subheader(f"📋 Preview dữ liệu: {sel_s}")
            st.dataframe(df_raw.head(10), use_container_width=True)
            if st.button("🚀 Chạy Unpivot"): res_final = run_unpivot(df_raw, h_r, i_c, d_s, sheet_name=sel_s)
        else:
            if st.button("🚀 Chạy Gộp tất cả Sheet"):
                all_res = [run_unpivot(pd.read_excel(file_up, s, header=None), h_r, i_c, d_s, s) for s in sheet_names]
                res_final = pd.concat([r for r in all_res if r is not None], ignore_index=True)

        if res_final is not None:
            st.success(f"Xử lý xong {len(res_final)} dòng.")
            c1, c2 = st.columns(2)
            with c1: st.plotly_chart(px.bar(res_final.groupby("Đối tượng")["Số tiền"].sum().nlargest(10).reset_index(), x="Đối tượng", y="Số tiền", title="Top 10 Đối tượng"), use_container_width=True)
            with c2:
                sel_pie = st.selectbox("Chọn hạng mục biểu đồ tròn:", [c for c in res_final.columns if c != "Số tiền"])
                st.plotly_chart(px.pie(res_final, values="Số tiền", names=sel_pie, title=f"Cơ cấu theo {sel_pie}"), use_container_width=True)
            st.dataframe(res_final, use_container_width=True)
            out = BytesIO()
            res_final.to_excel(out, index=False)
            st.download_button("📥 Tải kết quả Unpivot (.xlsx)", out.getvalue(), "Unpivot_Final.xlsx")

# --- MODULE 2: ĐỐI SOÁT & SO KHỚP MỜ (GIỮ NGUYÊN 100% PREVIEW) ---
elif menu == "🔍 Đối soát & So khớp mờ":
    st.title("🔍 Đối soát & So khớp mờ Thông minh")
    col1, col2 = st.columns(2)
    df_m = df_c = None
    with col1:
        f_m = st.file_uploader("File Gốc (Master)", type=["xlsx"], key="m")
        if f_m:
            xl_m = pd.ExcelFile(f_m)
            s_m = st.selectbox("Chọn Sheet Master:", xl_m.sheet_names)
            df_m = pd.read_excel(f_m, sheet_name=s_m)
            st.markdown(f"**Preview Master ({s_m}):**")
            st.dataframe(df_m.head(10), use_container_width=True)
    with col2:
        f_c = st.file_uploader("File Đối soát", type=["xlsx"], key="c")
        if f_c:
            xl_c = pd.ExcelFile(f_c)
            s_c = st.selectbox("Chọn Sheet Check:", xl_c.sheet_names)
            df_c = pd.read_excel(f_c, sheet_name=s_c)
            st.markdown(f"**Preview Check ({s_c}):**")
            st.dataframe(df_c.head(10), use_container_width=True)

    if df_m is not None and df_c is not None:
        st.sidebar.header("⚙️ Cài đặt Đối soát")
        k_m = st.sidebar.selectbox("Cột Key (Master):", df_m.columns); k_c = st.sidebar.selectbox("Cột Key (Check):", df_c.columns)
        v_col = st.sidebar.selectbox("Cột Tiền so khớp:", df_m.columns)
        fuz = st.sidebar.checkbox("Bật So khớp mờ"); score = st.sidebar.slider("Độ tương đồng %", 50, 100, 85) / 100
        if st.button("🚀 Bắt đầu đối soát", type="primary"):
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
            st.dataframe(merged.style.applymap(lambda x: 'background-color: #ffcccc' if x != 0 else '', subset=['Chênh lệch']), use_container_width=True)
            out_ds = BytesIO()
            merged.to_excel(out_ds, index=False)
            st.download_button("📥 Tải báo cáo đối soát (.xlsx)", out_ds.getvalue(), "Bao_cao_doi_soat.xlsx")

# --- MODULE 3: SỬA LỖI FONT (GIỮ NGUYÊN 100%) ---
elif menu == "🛠️ Tiện ích Sửa lỗi Font":
    st.title("🛠️ Chuẩn hóa Font chữ Tiếng Việt")
    file_f = st.file_uploader("Tải file cần sửa font", type=["xlsx"], key="font")
    if file_f:
        xl_f = pd.ExcelFile(file_f); s_f = st.selectbox("Chọn Sheet:", xl_f.sheet_names)
        df_f = pd.read_excel(file_f, sheet_name=s_f)
        st.dataframe(df_f.head(10)); target_cols = st.multiselect("Chọn các cột cần sửa:", df_f.columns)
        if st.button("🚀 Thực hiện sửa font"):
            for col in target_cols: df_f[col] = df_f[col].apply(fix_vietnamese_font)
            st.success("Đã chuẩn hóa!"); st.dataframe(df_f.head(10))
            out_f = BytesIO(); df_f.to_excel(out_f, index=False)
            st.download_button("📥 Tải file đã sửa (.xlsx)", out_f.getvalue(), "File_Unicode.xlsx")

# --- MODULE 4: TÁCH FILE HÀNG LOẠT (TÍNH NĂNG MỚI) ---
elif menu == "📂 Tách File hàng loạt (ZIP)":
    st.title("📂 Chia tách File lớn thành nhiều File nhỏ")
    st.info("Tính năng này giúp tách một file tổng thành nhiều file Excel riêng biệt dựa trên một cột bạn chọn (ví dụ: tách theo Tỉnh, theo Phòng ban).")
    
    file_split = st.file_uploader("Tải file Excel cần tách", type=["xlsx"], key="split_up")
    if file_split:
        xl_s = pd.ExcelFile(file_split); s_s = st.selectbox("Chọn Sheet dữ liệu:", xl_s.sheet_names)
        df_s = pd.read_excel(file_split, sheet_name=s_s)
        st.subheader("📋 Preview dữ liệu")
        st.dataframe(df_s.head(10), use_container_width=True)
        
        split_col = st.selectbox("Chọn cột dùng để tách file (Ví dụ: Tỉnh thành):", df_s.columns)
        
        if st.button("🚀 Bắt đầu tách và nén ZIP", type="primary"):
            unique_vals = df_s[split_col].unique()
            zip_buffer = BytesIO()
            
            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                for val in unique_vals:
                    # Lọc dữ liệu theo từng giá trị duy nhất
                    df_filtered = df_s[df_s[split_col] == val]
                    # Lưu file con vào bộ nhớ
                    sub_buffer = BytesIO()
                    df_filtered.to_excel(sub_buffer, index=False)
                    # Đưa vào file ZIP
                    zip_file.writestr(f"{val}.xlsx", sub_buffer.getvalue())
            
            st.success(f"Đã tách thành công thành {len(unique_vals)} file nhỏ!")
            st.download_button(
                label="📥 Tải toàn bộ File Tách (.ZIP)",
                data=zip_buffer.getvalue(),
                file_name="File_Tach_Tong_Hop.zip",
                mime="application/zip"
            )
