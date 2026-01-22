import streamlit as st
import pandas as pd
from io import BytesIO
import json
import os
import plotly.express as px
import difflib

# --- CẤU HÌNH HỆ THỐNG ---
st.set_page_config(page_title="Excel Hub Pro v8", layout="wide", page_icon="🚀")

CONFIG_FILE = "excel_profiles_v8.json"

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
st.sidebar.title("🛠️ Excel Master Hub")
menu = st.sidebar.radio("Nghiệp vụ:", ["🔄 Unpivot & Dashboard", "🔍 Đối soát & So khớp mờ"])

# --- MODULE 1: UNPIVOT & DASHBOARD ---
if menu == "🔄 Unpivot & Dashboard":
    st.title("🔄 Unpivot Ma trận & Phân tích Dashboard")
    file_up = st.file_uploader("1. Tải file Excel ma trận", type=["xlsx", "xls"], key="unp_up")
    
    if file_up:
        xl = pd.ExcelFile(file_up)
        sheet_names = xl.sheet_names
        
        with st.sidebar:
            st.header("⚙️ Profile cấu hình")
            sel_p = st.selectbox("Chọn Profile:", list(st.session_state['profiles'].keys()))
            cfg = st.session_state['profiles'][sel_p]
            h_r = st.number_input("Số hàng tiêu đề:", value=cfg['h_rows'])
            i_c = st.number_input("Cột Tên (A=0, B=1...):", value=cfg['id_col'])
            d_s = st.number_input("Dòng bắt đầu data:", value=cfg['d_start'])
            if st.button("💾 Lưu Profile"):
                name = st.text_input("Tên:")
                if name:
                    st.session_state['profiles'][name] = {"h_rows": h_r, "id_col": i_c, "d_start": d_s}
                    save_profiles(st.session_state['profiles'])
        
        mode = st.radio("Chế độ:", ["Xử lý 1 Sheet", "Xử lý Toàn bộ Sheet"], horizontal=True)
        res_final = None

        if mode == "Xử lý 1 Sheet":
            sel_s = st.selectbox("Chọn Sheet:", sheet_names)
            df_raw = pd.read_excel(file_up, sheet_name=sel_s, header=None)
            st.subheader(f"📋 Preview: {sel_s}")
            st.dataframe(df_raw.head(10), use_container_width=True)
            if st.button("🚀 Chạy Unpivot"):
                res_final = run_unpivot(df_raw, h_r, i_c, d_s, sheet_name=sel_s)
        else:
            if st.button("🚀 Chạy Gộp Toàn bộ"):
                all_res = [run_unpivot(pd.read_excel(file_up, s, header=None), h_r, i_c, d_s, s) for s in sheet_names]
                res_final = pd.concat([r for r in all_res if r is not None], ignore_index=True)

        if res_final is not None:
            st.success(f"✅ Đã xử lý {len(res_final)} dòng.")
            # Dashboard
            c1, c2 = st.columns(2)
            with c1: st.plotly_chart(px.bar(res_final.groupby("Đối tượng")["Số tiền"].sum().nlargest(10).reset_index(), x="Đối tượng", y="Số tiền", title="Top 10 Đối tượng"), use_container_width=True)
            with c2: st.plotly_chart(px.pie(res_final.groupby(res_final.columns[-1])["Số tiền"].sum().reset_index(), values="Số tiền", names=res_final.columns[-1], title="Cơ cấu"), use_container_width=True)
            
            st.dataframe(res_final, use_container_width=True)
            # NÚT TẢI FILE - KHÔNG ĐƯỢC THIẾU
            out = BytesIO()
            res_final.to_excel(out, index=False)
            st.download_button("📥 Tải kết quả Unpivot (.xlsx)", out.getvalue(), "Ket_qua_Unpivot.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# --- MODULE 2: ĐỐI SOÁT & SO KHỚP MỜ ---
elif menu == "🔍 Đối soát & So khớp mờ":
    st.title("🔍 Đối soát dữ liệu thông minh")
    
    col_a, col_b = st.columns(2)
    with col_a:
        f_m = st.file_uploader("File Gốc (Master)", type=["xlsx"], key="m")
        if f_m:
            xl_m = pd.ExcelFile(f_m)
            s_m = st.selectbox("Sheet Master:", xl_m.sheet_names)
            df_m = pd.read_excel(f_m, sheet_name=s_m)
            st.dataframe(df_m.head(5))

    with col_b:
        f_c = st.file_uploader("File Thực tế (Check)", type=["xlsx"], key="c")
        if f_c:
            xl_c = pd.ExcelFile(f_c)
            s_c = st.selectbox("Sheet Check:", xl_c.sheet_names)
            df_c = pd.read_excel(f_c, sheet_name=s_c)
            st.dataframe(df_c.head(5))

    if f_m and f_c:
        st.sidebar.header("⚙️ Cài đặt Đối soát")
        key_m = st.sidebar.selectbox("Cột Mã/Tên (Master):", df_m.columns)
        key_c = st.sidebar.selectbox("Cột Mã/Tên (Check):", df_c.columns)
        val_col = st.sidebar.selectbox("Cột Số tiền để so sánh:", df_m.columns)
        
        fuzzy_on = st.sidebar.checkbox("Bật So khớp mờ (Fuzzy Match)")
        score = st.sidebar.slider("Độ tương đồng (%)", 50, 100, 85) / 100

        if st.button("🚀 Thực hiện đối soát", type="primary"):
            try:
                with st.spinner("Đang khớp dữ liệu..."):
                    if fuzzy_on:
                        m_list = df_m[key_m].astype(str).tolist()
                        c_list = df_c[key_c].astype(str).tolist()
                        mapping = {k: find_fuzzy_match(k, c_list, score) for k in m_list}
                        df_m['Key_Matched'] = df_m[key_m].map(mapping)
                        merged = pd.merge(df_m, df_c, left_on='Key_Matched', right_on=key_c, how='outer', suffixes=('_Gốc', '_ThựcTế'))
                    else:
                        merged = pd.merge(df_m, df_c, left_on=key_m, right_on=key_c, how='outer', suffixes=('_Gốc', '_ThựcTế'))
                    
                    merged = merged.fillna(0)
                    # Xác định cột tiền sau khi merge
                    c_g = f"{val_col}_Gốc" if f"{val_col}_Gốc" in merged.columns else val_col
                    c_t = f"{val_col}_ThựcTế" if f"{val_col}_ThựcTế" in merged.columns else val_col
                    merged['Chênh lệch'] = merged[c_g] - merged[c_t]
                    
                    st.subheader("Báo cáo chênh lệch")
                    st.dataframe(merged, use_container_width=True)
                    
                    # NÚT TẢI BÁO CÁO - KHÔNG ĐƯỢC THIẾU
                    out_ds = BytesIO()
                    merged.to_excel(out_ds, index=False)
                    st.download_button("📥 Tải báo cáo đối soát (.xlsx)", out_ds.getvalue(), "Bao_cao_doi_soat.xlsx")
            except Exception as e:
                st.error(f"Lỗi đối soát: {e}. Vui lòng kiểm tra lại tên cột giữa 2 sheet.")
