import streamlit as st
import pandas as pd
from io import BytesIO
import json
import os

# Cấu hình trang
st.set_page_config(page_title="Excel Hub Pro v4", layout="wide", page_icon="📑")

CONFIG_FILE = "excel_hub_profiles.json"

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
                    # Gắn tên sheet nếu có
                    if sheet_name:
                        entry["Nguồn (Sheet)"] = sheet_name
                    # Gắn tiêu đề động
                    for i in range(h_rows):
                        entry[f"Tiêu đề {i+1}"] = headers.iloc[i, col_idx - (id_col + 1)]
                    results.append(entry)
        return pd.DataFrame(results)
    except Exception as e:
        return None

# --- GIAO DIỆN CHÍNH ---
st.sidebar.title("🎮 Hệ thống Excel Pro")
app_mode = st.sidebar.radio("Nghiệp vụ cần xử lý:", ["🔄 Unpivot (Ngang -> Dọc)", "🔍 Đối soát dữ liệu"])

# --- 1. MODULE UNPIVOT ---
if app_mode == "🔄 Unpivot (Ngang -> Dọc)":
    st.title("🔄 Unpivot Ma trận Đa năng")
    st.markdown("Hỗ trợ xử lý đơn lẻ từng sheet hoặc gộp toàn bộ các sheet trong file.")

    file_up = st.file_uploader("Bước 1: Tải file Excel lên", type=["xlsx", "xls"], key="unpivot_upload")

    if file_up:
        xl = pd.ExcelFile(file_up)
        sheet_names = xl.sheet_names
        
        # --- CẤU HÌNH SIDEBAR ---
        with st.sidebar:
            st.markdown("---")
            st.header("⚙️ Cấu hình Profile")
            p_names = list(st.session_state['profiles'].keys())
            sel_p = st.selectbox("Sử dụng Profile:", p_names)
            cfg = st.session_state['profiles'][sel_p]
            
            h_r = st.number_input("Số hàng tiêu đề:", value=cfg['h_rows'])
            i_c = st.number_input("Cột Tên (A=0, B=1...):", value=cfg['id_col'])
            d_s = st.number_input("Dòng bắt đầu dữ liệu:", value=cfg['d_start'])
            
            if st.button("💾 Lưu Profile mới"):
                save_name = st.text_input("Tên profile:")
                if save_name:
                    st.session_state['profiles'][save_name] = {"h_rows": h_r, "id_col": i_c, "d_start": d_s}
                    save_profiles(st.session_state['profiles'])
                    st.success("Đã lưu!")

        # --- CHỌN CHẾ ĐỘ XỬ LÝ ---
        st.subheader("📋 Bước 2: Chọn chế độ xử lý")
        mode = st.radio("Chế độ:", ["Xử lý 1 Sheet (Có Preview)", "Xử lý tất cả Sheet (Gộp dữ liệu)"], horizontal=True)

        if mode == "Xử lý 1 Sheet (Có Preview)":
            selected_sheet = st.selectbox("Chọn Sheet hiển thị:", sheet_names)
            df_raw = pd.read_excel(file_up, sheet_name=selected_sheet, header=None)
            st.dataframe(df_raw.head(15), use_container_width=True)
            
            if st.button("🚀 Chạy Unpivot Sheet này", type="primary"):
                res = run_unpivot(df_raw, h_r, i_c, d_s, sheet_name=selected_sheet)
                if res is not None and not res.empty:
                    st.success(f"Hoàn tất sheet {selected_sheet}!")
                    st.dataframe(res)
                    out = BytesIO()
                    res.to_excel(out, index=False)
                    st.download_button("📥 Tải kết quả", out.getvalue(), f"unpivot_{selected_sheet}.xlsx")

        else: # CHẾ ĐỘ XỬ LÝ TẤT CẢ SHEET
            st.warning("⚠️ Chế độ này sẽ áp dụng cấu hình trên cho TẤT CẢ các sheet trong file.")
            st.write(f"Danh sách sheet sẽ xử lý: {', '.join(sheet_names)}")
            
            if st.button("🚀 Chạy Unpivot TOÀN BỘ Sheet", type="primary"):
                all_results = []
                progress_bar = st.progress(0)
                
                for idx, s_name in enumerate(sheet_names):
                    df_s = pd.read_excel(file_up, sheet_name=s_name, header=None)
                    res_s = run_unpivot(df_s, h_r, i_c, d_s, sheet_name=s_name)
                    if res_s is not None:
                        all_results.append(res_s)
                    progress_bar.progress((idx + 1) / len(sheet_names))
                
                if all_results:
                    final_df = pd.concat(all_results, ignore_index=True)
                    st.success(f"Đã gộp thành công {len(sheet_names)} sheet. Tổng cộng {len(final_df)} dòng.")
                    st.dataframe(final_df)
                    
                    out_all = BytesIO()
                    final_df.to_excel(out_all, index=False)
                    st.download_button("📥 Tải file Gộp tất cả Sheet", out_all.getvalue(), "Unpivot_All_Sheets.xlsx")
                else:
                    st.error("Không có dữ liệu nào được tìm thấy trong các sheet.")

# --- 2. MODULE ĐỐI SOÁT (Giữ nguyên) ---
elif app_mode == "🔍 Đối soát dữ liệu":
    st.title("🔍 Đối soát dữ liệu")
    # ... (Giữ nguyên code module đối soát từ bản v3) ...
