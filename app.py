import streamlit as st
import pandas as pd
from io import BytesIO
import json
import os
import plotly.express as px
import difflib

# --- CẤU HÌNH HỆ THỐNG ---
st.set_page_config(page_title="Excel Hub Pro v7", layout="wide", page_icon="🚀")

CONFIG_FILE = "excel_profiles_v7.json"

# Hàm quản lý cấu hình Profile
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

# --- HÀM SO KHỚP MỜ (FUZZY MATCHING) ---
def find_fuzzy_match(name, choices, cutoff=0.6):
    matches = difflib.get_close_matches(str(name), [str(c) for c in choices], n=1, cutoff=cutoff)
    return matches[0] if matches else None

# --- MODULE XỬ LÝ UNPIVOT ---
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

# --- GIAO DIỆN SIDEBAR ---
st.sidebar.title("🎮 Trung tâm Excel Pro")
menu = st.sidebar.radio("Nghiệp vụ:", ["🔄 Unpivot & Dashboard", "🔍 Đối soát & So khớp mờ"])

# --- MODULE 1: UNPIVOT & DASHBOARD ---
if menu == "🔄 Unpivot & Dashboard":
    st.title("🔄 Chuyển đổi Ma trận & Phân tích Dashboard")
    
    file_up = st.file_uploader("1. Tải file Excel ma trận", type=["xlsx", "xls"], key="unp_up")
    
    if file_up:
        xl = pd.ExcelFile(file_up)
        sheet_names = xl.sheet_names
        
        with st.sidebar:
            st.header("⚙️ Cấu hình Profile")
            p_names = list(st.session_state['profiles'].keys())
            sel_p = st.selectbox("Sử dụng Profile:", p_names)
            cfg = st.session_state['profiles'][sel_p]
            
            h_r = st.number_input("Số hàng tiêu đề:", value=cfg['h_rows'])
            i_c = st.number_input("Cột Tên (A=0, B=1...):", value=cfg['id_col'])
            d_s = st.number_input("Dòng bắt đầu data:", value=cfg['d_start'])
            
            new_p_name = st.text_input("Tên profile mới:")
            if st.button("💾 Lưu Cấu hình"):
                st.session_state['profiles'][new_p_name] = {"h_rows": h_r, "id_col": i_c, "d_start": d_s}
                save_profiles(st.session_state['profiles'])
                st.success("Đã lưu vĩnh viễn!")

        mode = st.radio("Chế độ xử lý:", ["Xử lý 1 Sheet (Có Preview)", "Xử lý TẤT CẢ Sheet (Gộp dữ liệu)"], horizontal=True)

        res_final = None
        if mode == "Xử lý 1 Sheet (Có Preview)":
            sel_s = st.selectbox("Chọn Sheet:", sheet_names)
            df_raw = pd.read_excel(file_up, sheet_name=sel_s, header=None)
            st.subheader(f"📋 Preview dữ liệu: {sel_s}")
            st.dataframe(df_raw.head(15), use_container_width=True)
            if st.button("🚀 Chạy Unpivot Sheet này"):
                res_final = run_unpivot(df_raw, h_r, i_c, d_s, sheet_name=sel_s)
        else:
            st.info(f"Hệ thống sẽ gộp {len(sheet_names)} sheet.")
            if st.button("🚀 Chạy gộp tất cả Sheet"):
                all_res = []
                for s in sheet_names:
                    df_s = pd.read_excel(file_up, sheet_name=s, header=None)
                    all_res.append(run_unpivot(df_s, h_r, i_c, d_s, s))
                res_final = pd.concat([r for r in all_res if r is not None], ignore_index=True)

        if res_final is not None and not res_final.empty:
            st.success(f"Xử lý thành công {len(res_final)} dòng dữ liệu!")
            
            # DASHBOARD
            st.markdown("---")
            st.subheader("📊 Dashboard Phân tích")
            col_d1, col_d2 = st.columns(2)
            with col_d1:
                top_data = res_final.groupby("Đối tượng")["Số tiền"].sum().nlargest(10).reset_index()
                st.plotly_chart(px.bar(top_data, x="Đối tượng", y="Số tiền", title="Top 10 người nhận tiền nhiều nhất"), use_container_width=True)
            with col_d2:
                pie_col = "Tiêu đề 1" if "Tiêu đề 1" in res_final.columns else "Đối tượng"
                pie_data = res_final.groupby(pie_col)["Số tiền"].sum().reset_index()
                st.plotly_chart(px.pie(pie_data, values="Số tiền", names=pie_col, title="Cơ cấu theo danh mục"), use_container_width=True)
            
            # HIỂN THỊ DATA & NÚT TẢI
            st.dataframe(res_final, use_container_width=True)
            out = BytesIO()
            res_final.to_excel(out, index=False)
            st.download_button(label="📥 Tải kết quả xử lý (.xlsx)", data=out.getvalue(), file_name="Ket_qua_Unpivot.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# --- MODULE 2: ĐỐI SOÁT & SO KHỚP MỜ ---
elif menu == "🔍 Đối soát dữ liệu":
    st.title("🔍 Đối soát & So khớp mờ Thông minh")
    
    col_a, col_b = st.columns(2)
    with col_a:
        f_m = st.file_uploader("File Master (Gốc)", type=["xlsx"], key="m_up")
        if f_m:
            xl_m = pd.ExcelFile(f_m)
            s_m = st.selectbox("Chọn Sheet Master:", xl_m.sheet_names)
            df_m = pd.read_excel(f_m, sheet_name=s_m)
            st.dataframe(df_m.head(5))

    with col_b:
        f_c = st.file_uploader("File Đối soát", type=["xlsx"], key="c_up")
        if f_c:
            xl_c = pd.ExcelFile(f_c)
            s_c = st.selectbox("Chọn Sheet đối soát:", xl_c.sheet_names)
            df_c = pd.read_excel(f_c, sheet_name=s_c)
            st.dataframe(df_c.head(5))

    if f_m and f_c:
        st.sidebar.header("⚙️ Cài đặt So khớp")
        key_m = st.sidebar.selectbox("Cột Key (Master):", df_m.columns)
        key_c = st.sidebar.selectbox("Cột Key (Check):", df_c.columns)
        val_col = st.sidebar.selectbox("Cột Tiền so sánh:", df_m.columns)
        
        is_fuzzy = st.sidebar.checkbox("Bật So khớp mờ (Fuzzy)")
        cutoff = st.sidebar.slider("Độ tương đồng (%)", 50, 100, 85) / 100

        if st.button("🚀 Thực hiện đối soát", type="primary"):
            with st.spinner("Đang tính toán chênh lệch..."):
                if is_fuzzy:
                    m_keys = df_m[key_m].astype(str).tolist()
                    c_keys = df_c[key_c].astype(str).tolist()
                    mapping = {k: find_fuzzy_match(k, c_keys, cutoff) for k in m_keys}
                    df_m['Key_Matched'] = df_m[key_m].map(mapping)
                    merged = pd.merge(df_m, df_c, left_on='Key_Matched', right_on=key_c, how='outer', suffixes=('_Gốc', '_ThựcTế'))
                else:
                    merged = pd.merge(df_m, df_c, left_on=key_m, right_on=key_c, how='outer', suffixes=('_Gốc', '_ThựcTế'))
                
                merged = merged.fillna(0)
                col_g = f"{val_col}_Gốc" if f"{val_col}_Gốc" in merged.columns else val_col
                col_t = f"{val_col}_ThựcTế" if f"{val_col}_ThựcTế" in merged.columns else val_col
                merged['Chênh lệch'] = merged[col_g] - merged[col_t]
                
                st.subheader("Báo cáo chênh lệch")
                st.dataframe(merged.style.applymap(lambda x: 'background-color: #ffcccc' if x != 0 else '', subset=['Chênh lệch']))
                
                # NÚT TẢI BÁO CÁO ĐỐI SOÁT
                out_ds = BytesIO()
                merged.to_excel(out_ds, index=False)
                st.download_button(label="📥 Tải báo cáo đối soát (.xlsx)", data=out_ds.getvalue(), file_name="Bao_cao_doi_soat.xlsx")
