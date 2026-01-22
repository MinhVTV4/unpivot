import streamlit as st
import pandas as pd
from io import BytesIO
import json
import os
import plotly.express as px
import difflib

# --- CẤU HÌNH HỆ THỐNG ---
st.set_page_config(page_title="Excel Hub Pro v9", layout="wide", page_icon="📊")

CONFIG_FILE = "excel_profiles_v9.json"

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
                        # Lấy tiêu đề tương ứng từ các hàng đầu
                        entry[f"Tiêu đề {i+1}"] = headers.iloc[i, col_idx - (id_col + 1)]
                    results.append(entry)
        return pd.DataFrame(results)
    except: return None

# --- SIDEBAR MENU ---
st.sidebar.title("🛠️ Siêu công cụ Excel")
menu = st.sidebar.radio("Nghiệp vụ cần làm:", ["🔄 Unpivot & Dashboard", "🔍 Đối soát & So khớp mờ"])

# --- MODULE 1: UNPIVOT & DASHBOARD ---
if menu == "🔄 Unpivot & Dashboard":
    st.title("🔄 Chuyển đổi Ma trận & Phân tích Dashboard")
    file_up = st.file_uploader("1. Tải file Excel ma trận", type=["xlsx", "xls"], key="unp_up")
    
    if file_up:
        xl = pd.ExcelFile(file_up)
        sheet_names = xl.sheet_names
        
        with st.sidebar:
            st.header("⚙️ Cấu hình Profile")
            sel_p = st.selectbox("Chọn Profile:", list(st.session_state['profiles'].keys()))
            cfg = st.session_state['profiles'][sel_p]
            h_r = st.number_input("Số hàng tiêu đề:", value=cfg['h_rows'])
            i_c = st.number_input("Cột Tên (A=0, B=1...):", value=cfg['id_col'])
            d_s = st.number_input("Dòng bắt đầu data:", value=cfg['d_start'])
            if st.button("💾 Lưu cấu hình này"):
                name = st.text_input("Đặt tên Profile mới:")
                if name:
                    st.session_state['profiles'][name] = {"h_rows": h_r, "id_col": i_c, "d_start": d_s}
                    save_profiles(st.session_state['profiles'])
        
        mode = st.radio("Chế độ xử lý:", ["Xử lý 1 Sheet (Có Preview)", "Xử lý Toàn bộ Sheet (Gộp dữ liệu)"], horizontal=True)
        res_final = None

        if mode == "Xử lý 1 Sheet (Có Preview)":
            sel_s = st.selectbox("Chọn Sheet:", sheet_names)
            df_raw = pd.read_excel(file_up, sheet_name=sel_s, header=None)
            st.subheader(f"📋 Preview dữ liệu: {sel_s}")
            st.dataframe(df_raw.head(10), use_container_width=True)
            if st.button("🚀 Chạy Unpivot Sheet này"):
                res_final = run_unpivot(df_raw, h_r, i_c, d_s, sheet_name=sel_s)
        else:
            if st.button("🚀 Chạy Gộp TẤT CẢ Sheet"):
                with st.spinner("Đang gộp dữ liệu các sheet..."):
                    all_res = [run_unpivot(pd.read_excel(file_up, s, header=None), h_r, i_c, d_s, s) for s in sheet_names]
                    res_final = pd.concat([r for r in all_res if r is not None], ignore_index=True)

        if res_final is not None and not res_final.empty:
            st.success(f"✅ Đã xử lý thành công {len(res_final)} dòng.")
            
            # --- DASHBOARD (SỬA LỖI BIỂU ĐỒ TRÒN) ---
            st.markdown("---")
            st.subheader("📊 Dashboard Phân tích")
            c1, c2 = st.columns(2)
            
            with c1:
                top_data = res_final.groupby("Đối tượng")["Số tiền"].sum().nlargest(10).reset_index()
                st.plotly_chart(px.bar(top_data, x="Đối tượng", y="Số tiền", title="Top 10 Đối tượng cao nhất", color="Số tiền"), use_container_width=True)
            
            with c2:
                # Tính năng mới: Cho phép chọn cột để vẽ biểu đồ tròn để không bao giờ bị lỗi
                available_cols = [c for c in res_final.columns if c != "Số tiền"]
                # Ưu tiên chọn 'Nguồn (Sheet)' hoặc 'Tiêu đề 1'
                default_idx = 0
                if "Tiêu đề 1" in available_cols: default_idx = available_cols.index("Tiêu đề 1")
                elif "Nguồn (Sheet)" in available_cols: default_idx = available_cols.index("Nguồn (Sheet)")
                
                sel_pie = st.selectbox("Chọn hạng mục vẽ biểu đồ tròn:", available_cols, index=default_idx)
                pie_data = res_final.groupby(sel_pie)["Số tiền"].sum().reset_index()
                st.plotly_chart(px.pie(pie_data, values="Số tiền", names=sel_pie, title=f"Cơ cấu tiền theo {sel_pie}"), use_container_width=True)
            
            st.dataframe(res_final, use_container_width=True)
            # NÚT TẢI FILE UNPIVOT
            out = BytesIO()
            res_final.to_excel(out, index=False)
            st.download_button("📥 Tải kết quả Unpivot (.xlsx)", out.getvalue(), "Ket_qua_Unpivot.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# --- MODULE 2: ĐỐI SOÁT & SO KHỚP MỜ ---
elif menu == "🔍 Đối soát & So khớp mờ":
    st.title("🔍 Đối soát dữ liệu thông minh")
    
    col_a, col_b = st.columns(2)
    with col_a:
        f_m = st.file_uploader("File Gốc (Master)", type=["xlsx"], key="f_m")
        if f_m:
            xl_m = pd.ExcelFile(f_m)
            s_m = st.selectbox("Chọn Sheet Master:", xl_m.sheet_names)
            df_m = pd.read_excel(f_m, sheet_name=s_m)
            st.dataframe(df_m.head(5))

    with col_b:
        f_c = st.file_uploader("File Đối soát", type=["xlsx"], key="f_c")
        if f_c:
            xl_c = pd.ExcelFile(f_c)
            s_c = st.selectbox("Chọn Sheet Đối soát:", xl_c.sheet_names)
            df_c = pd.read_excel(f_c, sheet_name=s_c)
            st.dataframe(df_c.head(5))

    if f_m and f_c:
        st.sidebar.header("⚙️ Cài đặt Đối soát")
        k_m = st.sidebar.selectbox("Cột Key (Master):", df_m.columns)
        k_c = st.sidebar.selectbox("Cột Key (Check):", df_c.columns)
        v_col = st.sidebar.selectbox("Cột Tiền để so khớp:", df_m.columns)
        
        fuz = st.sidebar.checkbox("Bật So khớp mờ (Fuzzy)")
        score = st.sidebar.slider("Độ tương đồng (%)", 50, 100, 85) / 100

        if st.button("🚀 Bắt đầu đối soát ngay", type="primary"):
            try:
                with st.spinner("Đang tính toán..."):
                    if fuz:
                        m_keys = df_m[k_m].astype(str).tolist()
                        c_keys = df_c[k_c].astype(str).tolist()
                        mapping = {k: find_fuzzy_match(k, c_keys, score) for k in m_keys}
                        df_m['Key_Matched'] = df_m[k_m].map(mapping)
                        merged = pd.merge(df_m, df_c, left_on='Key_Matched', right_on=k_c, how='outer', suffixes=('_Gốc', '_ThựcTế'))
                    else:
                        merged = pd.merge(df_m, df_c, left_on=k_m, right_on=k_c, how='outer', suffixes=('_Gốc', '_ThựcTế'))
                    
                    merged = merged.fillna(0)
                    col_g = f"{v_col}_Gốc" if f"{v_col}_Gốc" in merged.columns else v_col
                    col_t = f"{v_col}_ThựcTế" if f"{v_col}_ThựcTế" in merged.columns else v_col
                    merged['Chênh lệch'] = merged[col_g] - merged[col_t]
                    
                    st.subheader("Báo cáo kết quả đối soát")
                    st.dataframe(merged.style.applymap(lambda x: 'background-color: #ffcccc' if x != 0 else '', subset=['Chênh lệch']))
                    
                    # NÚT TẢI FILE ĐỐI SOÁT
                    out_ds = BytesIO()
                    merged.to_excel(out_ds, index=False)
                    st.download_button("📥 Tải báo cáo đối soát (.xlsx)", out_ds.getvalue(), "Bao_cao_doi_soat.xlsx")
            except Exception as e:
                st.error(f"Lỗi khi đối soát: {e}. Hãy kiểm tra xem tên cột có bị trùng lặp không.")
