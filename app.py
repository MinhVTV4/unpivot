import streamlit as st
import pandas as pd
from io import BytesIO
import json
import os
import plotly.express as px

# --- CẤU HÌNH HỆ THỐNG ---
st.set_page_config(page_title="Excel Hub Pro v5", layout="wide", page_icon="🚀")

CONFIG_FILE = "excel_profiles_v5.json"

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

# --- HÀM XỬ LÝ UNPIVOT ---
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
menu = st.sidebar.radio("Chọn nghiệp vụ:", ["🔄 Unpivot & Dashboard", "🔍 Đối soát dữ liệu"])

# --- MODULE 1: UNPIVOT & DASHBOARD ---
if menu == "🔄 Unpivot & Dashboard":
    st.title("🔄 Unpivot Ma trận & Phân tích Dashboard")
    
    file_up = st.file_uploader("1. Tải file Excel ma trận", type=["xlsx", "xls"], key="up_main")
    
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
            
            new_p = st.text_input("Lưu thành Profile mới:")
            if st.button("💾 Lưu Cấu hình"):
                st.session_state['profiles'][new_p] = {"h_rows": h_r, "id_col": i_c, "d_start": d_s}
                save_profiles(st.session_state['profiles'])
                st.success("Đã lưu!")

        mode = st.radio("Chế độ:", ["Xử lý 1 Sheet (Có Preview)", "Xử lý TẤT CẢ Sheet (Gộp)"], horizontal=True)

        res_final = None
        if mode == "Xử lý 1 Sheet (Có Preview)":
            sel_s = st.selectbox("Chọn Sheet:", sheet_names)
            df_raw = pd.read_excel(file_up, sheet_name=sel_s, header=None)
            st.dataframe(df_raw.head(15), use_container_width=True)
            if st.button("🚀 Chạy Unpivot"):
                res_final = run_unpivot(df_raw, h_r, i_c, d_s, sheet_name=sel_s)
        else:
            if st.button("🚀 Chạy Tất cả Sheet & Gộp"):
                all_res = []
                for s in sheet_names:
                    df_s = pd.read_excel(file_up, sheet_name=s, header=None)
                    all_res.append(run_unpivot(df_s, h_r, i_c, d_s, s))
                res_final = pd.concat([r for r in all_res if r is not None], ignore_index=True)

        if res_final is not None:
            st.success(f"Đã xử lý xong {len(res_final)} dòng!")
            
            # --- DASHBOARD ---
            st.markdown("---")
            st.subheader("📊 Dashboard Phân tích Nhanh")
            c1, c2 = st.columns(2)
            with c1:
                top_data = res_final.groupby("Đối tượng")["Số tiền"].sum().nlargest(10).reset_index()
                st.plotly_chart(px.bar(top_data, x="Đối tượng", y="Số tiền", title="Top 10 Đối tượng cao nhất"), use_container_width=True)
            with c2:
                pie_col = "Tiêu đề 1" if "Tiêu đề 1" in res_final.columns else "Đối tượng"
                pie_data = res_final.groupby(pie_col)["Số tiền"].sum().reset_index()
                st.plotly_chart(px.pie(pie_data, values="Số tiền", names=pie_col, title="Cơ cấu tiền"), use_container_width=True)
            
            st.dataframe(res_final, use_container_width=True)
            out = BytesIO()
            res_final.to_excel(out, index=False)
            st.download_button("📥 Tải kết quả (.xlsx)", out.getvalue(), "Unpivot_Result.xlsx")

# --- MODULE 2: ĐỐI SOÁT DỮ LIỆU ---
elif menu == "🔍 Đối soát dữ liệu":
    st.title("🔍 Đối soát & So khớp Đa Sheet")
    st.markdown("So sánh chênh lệch giữa 2 file bất kỳ.")

    col1, col2 = st.columns(2)
    with col1:
        f_m = st.file_uploader("Tải File Master (Gốc)", type=["xlsx"], key="m")
        if f_m:
            xl_m = pd.ExcelFile(f_m)
            s_m = st.selectbox("Chọn Sheet Master:", xl_m.sheet_names)
            df_m = pd.read_excel(f_m, sheet_name=s_m)
            st.dataframe(df_m.head(5))

    with col2:
        f_c = st.file_uploader("Tải File Đối Soát", type=["xlsx"], key="c")
        if f_c:
            xl_c = pd.ExcelFile(f_c)
            s_c = st.selectbox("Chọn Sheet Đối soát:", xl_c.sheet_names)
            df_c = pd.read_excel(f_c, sheet_name=s_c)
            st.dataframe(df_c.head(5))

    if f_m and f_c:
        st.sidebar.header("⚙️ Cài đặt Đối soát")
        key = st.sidebar.selectbox("Cột Mã khóa (Key):", df_m.columns)
        val = st.sidebar.selectbox("Cột Số tiền để so khớp:", df_m.columns)

        if st.button("🚀 Thực hiện đối soát", type="primary"):
            # Logic Merge & So khớp
            merged = pd.merge(df_m, df_c[[key, val]], on=key, how='outer', suffixes=('_Gốc', '_ThựcTế'))
            merged = merged.fillna(0)
            merged['Chênh lệch'] = merged[f'{val}_Gốc'] - merged[f'{val}_ThựcTế']
            
            # Cảnh báo rủi ro (Outliers) dùng công thức thống kê
            # Lệch > mean + 2*std
            m_val = merged['Chênh lệch'].mean()
            s_val = merged['Chênh lệch'].std()
            merged['Cảnh báo'] = merged['Chênh lệch'].apply(lambda x: '🚩 Sai lệch lớn' if abs(x) > (m_val + 2*s_val) and x != 0 else 'Bình thường')
            
            st.subheader("Báo cáo chênh lệch")
            st.dataframe(merged.style.applymap(lambda x: 'background-color: #ffcccc' if x != 0 else '', subset=['Chênh lệch']))
            
            out_err = BytesIO()
            merged.to_excel(out_err, index=False)
            st.download_button("📥 Tải Báo cáo Đối soát", out_err.getvalue(), "Bao_cao_doi_soat.xlsx")
