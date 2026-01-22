import streamlit as st
import pandas as pd
from io import BytesIO
import json
import os
import plotly.express as px
import difflib # Thư viện dùng để so khớp mờ

# --- CẤU HÌNH HỆ THỐNG ---
st.set_page_config(page_title="Excel Hub Pro v6", layout="wide", page_icon="🚀")

CONFIG_FILE = "excel_profiles_v6.json"

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
    """Tìm tên gần giống nhất trong danh sách choices"""
    matches = difflib.get_close_matches(name, choices, n=1, cutoff=cutoff)
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

# --- SIDEBAR MENU ---
st.sidebar.title("🎮 Siêu công cụ Excel")
menu = st.sidebar.radio("Chọn nghiệp vụ:", ["🔄 Unpivot & Dashboard", "🔍 Đối soát & So khớp mờ"])

# --- MODULE 1: UNPIVOT & DASHBOARD ---
if menu == "🔄 Unpivot & Dashboard":
    st.title("🔄 Unpivot & Phân tích Dashboard")
    file_up = st.file_uploader("Tải file Excel ma trận", type=["xlsx", "xls"])
    
    if file_up:
        xl = pd.ExcelFile(file_up)
        sheet_names = xl.sheet_names
        with st.sidebar:
            st.header("⚙️ Cấu hình Profile")
            p_names = list(st.session_state['profiles'].keys())
            sel_p = st.selectbox("Sử dụng Profile:", p_names)
            cfg = st.session_state['profiles'][sel_p]
            h_r, i_c, d_s = cfg['h_rows'], cfg['id_col'], cfg['d_start']
            
        mode = st.radio("Chế độ:", ["Xử lý 1 Sheet", "Xử lý Toàn bộ Sheet"], horizontal=True)
        res_final = None
        if mode == "Xử lý 1 Sheet":
            sel_s = st.selectbox("Chọn Sheet:", sheet_names)
            df_raw = pd.read_excel(file_up, sheet_name=sel_s, header=None)
            if st.button("🚀 Chạy Unpivot"):
                res_final = run_unpivot(df_raw, h_r, i_c, d_s, sheet_name=sel_s)
        else:
            if st.button("🚀 Chạy Tất cả Sheet & Gộp"):
                all_res = [run_unpivot(pd.read_excel(file_up, s, header=None), h_r, i_c, d_s, s) for s in sheet_names]
                res_final = pd.concat([r for r in all_res if r is not None], ignore_index=True)

        if res_final is not None:
            st.success("Xử lý thành công!")
            # Dashboard
            c1, c2 = st.columns(2)
            with c1:
                top_data = res_final.groupby("Đối tượng")["Số tiền"].sum().nlargest(10).reset_index()
                st.plotly_chart(px.bar(top_data, x="Đối tượng", y="Số tiền", title="Top 10 Đối tượng"), use_container_width=True)
            with c2:
                pie_col = "Tiêu đề 1" if "Tiêu đề 1" in res_final.columns else "Đối tượng"
                pie_data = res_final.groupby(pie_col)["Số tiền"].sum().reset_index()
                st.plotly_chart(px.pie(pie_data, values="Số tiền", names=pie_col, title="Cơ cấu tiền"), use_container_width=True)
            st.dataframe(res_final)

# --- MODULE 2: ĐỐI SOÁT & SO KHỚP MỜ ---
elif menu == "🔍 Đối soát & So khớp mờ":
    st.title("🔍 Đối soát dữ liệu thông minh")
    st.markdown("Hỗ trợ tìm kiếm các dòng dữ liệu gần giống nhau khi tên gọi không khớp 100%.")

    c1, c2 = st.columns(2)
    with c1:
        f_m = st.file_uploader("File Master (Gốc)", type=["xlsx"], key="m")
    with c2:
        f_c = st.file_uploader("File Cần đối soát", type=["xlsx"], key="c")

    if f_m and f_c:
        df_m = pd.read_excel(f_m)
        df_c = pd.read_excel(f_c)
        
        st.sidebar.header("⚙️ Cấu hình So khớp")
        key_m = st.sidebar.selectbox("Cột Mã/Tên (Master):", df_m.columns)
        key_c = st.sidebar.selectbox("Cột Mã/Tên (Check):", df_c.columns)
        val_col = st.sidebar.selectbox("Cột Số tiền để so:", df_m.columns)
        
        is_fuzzy = st.sidebar.checkbox("Bật So khớp mờ (Fuzzy Matching)")
        threshold = st.sidebar.slider("Độ tương đồng (%)", 50, 100, 80) / 100

        if st.button("🚀 Bắt đầu đối soát", type="primary"):
            with st.spinner("Đang thực hiện so khớp..."):
                if is_fuzzy:
                    # Logic So khớp mờ
                    master_keys = df_m[key_m].astype(str).tolist()
                    check_keys = df_c[key_c].astype(str).tolist()
                    
                    # Tạo bảng ánh xạ
                    mapping = {}
                    for k in master_keys:
                        match = find_fuzzy_match(k, check_keys, cutoff=threshold)
                        mapping[k] = match
                    
                    df_m['Key_Matched'] = df_m[key_m].map(mapping)
                    merged = pd.merge(df_m, df_c, left_on='Key_Matched', right_on=key_c, how='left', suffixes=('_Gốc', '_ThựcTế'))
                else:
                    # Logic So khớp chính xác
                    merged = pd.merge(df_m, df_c, left_on=key_m, right_on=key_c, how='left', suffixes=('_Gốc', '_ThựcTế'))
                
                merged = merged.fillna(0)
                # Đảm bảo lấy đúng cột tiền sau merge
                col_goc = f"{val_col}_Gốc" if f"{val_col}_Gốc" in merged.columns else val_col
                col_tt = f"{val_col}_ThựcTế" if f"{val_col}_ThựcTế" in merged.columns else val_col
                
                merged['Chênh lệch'] = merged[col_goc] - merged[col_tt]
                
                st.subheader("Kết quả đối soát")
                st.dataframe(merged.style.applymap(lambda x: 'background-color: #ffcccc' if x != 0 else '', subset=['Chênh lệch']))
                
                out = BytesIO()
                merged.to_excel(out, index=False)
                st.download_button("📥 Tải báo cáo đối soát", out.getvalue(), "Doi_soat_Fuzzy.xlsx")
