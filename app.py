import streamlit as st
import pandas as pd
from io import BytesIO

# Cấu hình trang web
st.set_page_config(page_title="Chuyển đổi Excel Ngang sang Dọc", layout="wide", page_icon="📊")

# --- CSS ĐỂ GIAO DIỆN ĐẸP HƠN ---
st.markdown("""
<style>
    .stDataFrame { border: 1px solid #e6e9ef; border-radius: 5px; }
    .main { background-color: #f8f9fa; }
</style>
""", unsafe_allow_html=True)

def transform_horizontal_to_vertical(df):
    """
    Hàm xử lý logic: Xoay bảng từ ngang sang dọc.
    - Hàng 1 (index 0): Ngày Giao dịch
    - Hàng 2 (index 1): Dòng mã
    - Hàng 3 (index 2): Nội dung
    - Cột 1 (index 0): Tên khoản mục
    """
    try:
        # 1. Tách header (3 hàng đầu, bỏ cột đầu tiên)
        headers = df.iloc[0:3, 1:]
        
        # 2. Tách dữ liệu chính (Từ hàng 4 trở đi)
        data_rows = df.iloc[3:, :]
        
        results = []
        
        # Duyệt qua từng hàng (Khoản mục)
        for _, row in data_rows.iterrows():
            item_name = str(row[0]).strip() # Lấy tên khoản mục ở cột A
            
            # Nếu tên khoản mục trống thì bỏ qua
            if not item_name or item_name == 'nan':
                continue
                
            # Duyệt qua từng cột (tương ứng với các cột Ngày/Mã/Nội dung)
            for col_idx in range(1, len(df.columns)):
                amount_raw = row[col_idx]
                
                # --- KHẮC PHỤC LỖI: Ép kiểu dữ liệu an toàn ---
                # Chuyển đổi về dạng số, nếu là chữ hoặc ký tự lạ sẽ biến thành NaN
                amount = pd.to_numeric(amount_raw, errors='coerce')
                
                # Chỉ lấy những ô có số tiền hợp lệ và lớn hơn 0
                if pd.notnull(amount) and amount > 0:
                    results.append({
                        "Ngày Giao dịch": headers.iloc[0, col_idx-1],
                        "Dòng mã": headers.iloc[1, col_idx-1],
                        "Nội dung": headers.iloc[2, col_idx-1],
                        "Khoản mục": item_name,
                        "Số tiền": amount
                    })
        
        # Chuyển danh sách kết quả thành DataFrame
        if not results:
            return pd.DataFrame()
            
        final_df = pd.DataFrame(results)
        
        # Định dạng lại cột Ngày nếu có (tùy chọn)
        # final_df['Ngày Giao dịch'] = pd.to_datetime(final_df['Ngày Giao dịch']).dt.strftime('%d/%m/%Y')
        
        return final_df
        
    except Exception as e:
        st.error(f"⚠️ Lỗi trong quá trình xử lý logic: {e}")
        return None

# --- GIAO DIỆN NGƯỜI DÙNG (UI) ---
st.title("🔄 Công cụ Unpivot Excel Chuyên nghiệp")
st.markdown("Chuyển đổi các bảng kê ngang (Ma trận) thành dạng danh sách dọc để dễ dàng quản lý và lọc dữ liệu.")

# 1. Khu vực Upload File
with st.container():
    uploaded_file = st.file_uploader("Tải lên file Excel (.xlsx)", type=["xlsx"])

if uploaded_file:
    # Đọc file thô không lấy header
    df_raw = pd.read_excel(uploaded_file, header=None)
    
    st.subheader("📋 1. Xem trước dữ liệu gốc")
    st.dataframe(df_raw.head(10), use_container_width=True)
    
    # 2. Nút bấm xử lý
    if st.button("🚀 Bắt đầu chuyển đổi ngay", type="primary"):
        with st.spinner("Đang xử lý và lọc dữ liệu..."):
            df_result = transform_horizontal_to_vertical(df_raw)
            
            if df_result is not None and not df_result.empty:
                st.subheader("✅ 2. Kết quả sau khi chuyển dọc")
                st.success(f"Đã tìm thấy {len(df_result)} dòng có phát sinh số tiền.")
                
                # Hiển thị bảng kết quả
                st.dataframe(df_result, use_container_width=True)
                
                # 3. Nút tải file Excel
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_result.to_excel(writer, index=False, sheet_name='Du_lieu_doc')
                    
                    # Tự động căn chỉnh độ rộng cột cho file Excel tải về
                    worksheet = writer.sheets['Du_lieu_doc']
                    for i, col in enumerate(df_result.columns):
                        column_len = max(df_result[col].astype(str).str.len().max(), len(col)) + 2
                        worksheet.set_column(i, i, column_len)

                st.download_button(
                    label="📥 Tải file kết quả Excel về máy",
                    data=output.getvalue(),
                    file_name="ket_qua_chuyen_doi.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            elif df_result is not None and df_result.empty:
                st.warning("⚠️ Không tìm thấy dữ liệu nào có số tiền lớn hơn 0.")
else:
    # Hướng dẫn khi chưa có file
    st.info("💡 Vui lòng tải lên file Excel có cấu trúc 3 hàng đầu là tiêu đề (Ngày, Mã, Nội dung) để bắt đầu.")

# Chân trang
st.markdown("---")
st.caption("Ứng dụng được xây dựng dựa trên cấu trúc xử lý của hang3.html")
