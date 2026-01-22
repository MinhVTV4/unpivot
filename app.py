import streamlit as st
import pandas as pd
from io import BytesIO
import json
import os

# Cấu hình trang
st.set_page_config(page_title="Excel Hub Pro", layout="wide", page_icon="📈")

CONFIG_FILE = "app_profiles.json"

# --- HÀM QUẢN LÝ CẤU HÌNH ---
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

# --- MODULE 1: LOGIC UNPIVOT ---
def run_unpivot(df, h_rows, id_col, d_start):
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
                    entry = {"Đối tượng": id_val, "Giá trị": val}
                    for i in range(h_rows):
                        entry[f"Tiêu đề {i+1}"] = headers.iloc[i, col_idx - (id_col + 1)]
                    results.append(entry)
        return pd.DataFrame(results)
    except Exception as e:
        st.error(f"Lỗi Unpivot: {e}")
        return None

# --- GIAO DIỆN SIDEBAR ---
st.sidebar.title("🎮 Menu Chức năng")
app_mode = st.sidebar.selectbox("Chọn nghiệp vụ cần làm:", ["🔄 Unpivot Vạn năng", "🔍 Đối soát & So khớp"])

# --- CHỨC NĂNG 1: UNPIVOT ---
if app_mode == "🔄 Unpivot Vạn năng":
    st.title("🔄 Trình Unpivot Excel Ma trận")
    st.markdown("Biến mọi bảng ngang phức tạp thành danh sách dọc để đối soát.")
    
    with st.sidebar:
        st.header("⚙️ Cấu hình Profile")
        p_names = list(st.session_state['profiles'].keys())
        sel_p = st.selectbox("Chọn Profile:", p_names)
        cfg = st.session_state['profiles'][sel_p]
        
        h_r = st.number_input("Số hàng tiêu đề:", value=cfg['h_rows'])
        i_c = st.number_input("Cột Định danh (B=1):", value=cfg['id_col'])
        d_s = st.number_input("Dòng bắt đầu dữ liệu:", value=cfg['d_start'])
        
        new_p = st.text_input("Lưu thành Profile mới:")
        if st.button("💾 Lưu cấu hình"):
            st.session_state['profiles'][new_p] = {"h_rows": h_r, "id_col": i_c, "d_start": d_s}
            save_profiles(st.session_state['profiles'])
            st.success("Đã lưu!")

    file_up = st.file_uploader("Tải file ma trận ngang", type=["xlsx"])
    if file_up:
        df_raw = pd.read_excel(file_up, header=None)
        st.subheader("Xem trước dữ liệu")
        st.dataframe(df_raw.head(10))
        
        if st.button("🚀 Thực hiện Unpivot"):
            res = run_unpivot(df_raw, h_r, i_c, d_s)
            if res is not None:
                st.success(f"Xong! {len(res)} dòng.")
                st.dataframe(res)
                out = BytesIO()
                res.to_excel(out, index=False)
                st.download_button("📥 Tải File Đọc (.xlsx)", out.getvalue(), "unpivot_result.xlsx")

# --- CHỨC NĂNG 2: ĐỐI SOÁT ---
elif app_mode == "🔍 Đối soát & So khớp":
    st.title("🔍 Hệ thống Đối soát & Cảnh báo")
    st.markdown("So sánh 2 file (Ví dụ: File Gốc vs File Thực tế) để tìm chênh lệch.")

    c1, c2 = st.columns(2)
    with c1:
        f_master = st.file_uploader("Tải File Master (Gốc)", type=["xlsx"])
    with c2:
        f_check = st.file_uploader("Tải File Cần đối soát", type=["xlsx"])

    if f_master and f_check:
        df_m = pd.read_excel(f_master)
        df_c = pd.read_excel(f_check)
        
        st.sidebar.header("⚙️ Cài đặt Đối soát")
        key = st.sidebar.selectbox("Cột Mã khóa (để khớp nhau):", df_m.columns)
        val = st.sidebar.selectbox("Cột Số tiền để so sánh:", df_m.columns)

        if st.button("🚀 Bắt đầu đối soát"):
            # Logic Đối soát
            merged = pd.merge(df_m, df_c[[key, val]], on=key, how='outer', suffixes=('_Gốc', '_ThựcTế'))
            merged = merged.fillna(0)
            merged['Chênh lệch'] = merged[f'{val}_Gốc'] - merged[f'{val}_ThựcTế']
            
            # Cảnh báo rủi ro (Outliers)
            mean_diff = merged['Chênh lệch'].mean()
            std_diff = merged['Chênh lệch'].std()
            merged['Cảnh báo'] = merged['Chênh lệch'].apply(lambda x: '🚩 Sai lệch lớn' if abs(x) > (mean_diff + 2*std_diff) else 'Bình thường')

            st.subheader("Kết quả đối soát")
            st.dataframe(merged.style.applymap(lambda x: 'background-color: #ffcccc' if x != 0 else '', subset=['Chênh lệch']))
            
            # Xuất báo cáo lỗi
            errors = merged[merged['Chênh lệch'] != 0]
            out_err = BytesIO()
            errors.to_excel(out_err, index=False)
            st.download_button("📥 Tải Báo cáo Chênh lệch", out_err.getvalue(), "bao_cao_chenh_lech.xlsx")
