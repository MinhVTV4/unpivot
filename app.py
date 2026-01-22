import streamlit as st
import pandas as pd
from io import BytesIO

# Cấu hình trang
st.set_page_config(page_title="Xử lý Excel Ngang sang Dọc", layout="wide")

def transform_horizontal_to_vertical(df):
    """
    Logic cốt lõi: Biến các cột Ngày/Chứng từ thành hàng dọc.
    Dựa theo file hang3.html: 
    - 3 hàng đầu chứa thông tin header (Ngày, Mã, Nội dung)
    - Cột đầu tiên chứa Tên khoản mục
    """
    try:
        # Lấy thông tin header từ 3 hàng đầu
        headers = df.iloc[0:3, 1:] # Bỏ cột đầu tiên
        data_rows = df.iloc[3:, :] # Dữ liệu bắt đầu từ hàng 4
        
        results = []
        
        # Duyệt qua từng hàng dữ liệu (Khoản mục)
        for _, row in data_rows.iterrows():
            item_name = row[0] # Tên khoản mục ở cột A
            
            # Duyệt qua từng cột (tương ứng với từng ngày/chứng từ)
            for col_idx in range(1, len(df.columns)):
                amount = row[col_idx]
                
                # Chỉ lấy các dòng có phát sinh tiền > 0
                if pd.notnull(amount) and amount > 0:
                    results.append({
                        "Ngày Giao dịch": headers.iloc[0, col_idx-1],
                        "Dòng mã": headers.iloc[1, col_idx-1],
                        "Nội dung": headers.iloc[2, col_idx-1],
                        "Khoản mục": item_name,
                        "Số tiền": amount
                    })
        
        return pd.DataFrame(results)
    except Exception as e:
        st.error(f"Lỗi cấu trúc file: {e}")
        return None

# --- GIAO DIỆN ---
st.title("🔄 Chuyển đổi Excel Ngang sang Dọc (Unpivot)")
st.info("Hệ thống sẽ tự động nhận diện 3 hàng đầu là Ngày, Mã, Nội dung và chuyển thành bảng dọc.")

uploaded_file = st.file_uploader("Tải file Excel cần xử lý", type=["xlsx"])

if uploaded_file:
    # Đọc file (giữ nguyên định dạng thô không lấy header tự động)
    df_raw = pd.read_excel(uploaded_file, header=None)
    
    st.subheader("1. Dữ liệu gốc (Bảng ngang)")
    st.dataframe(df_raw.head(10), use_container_width=True)

    if st.button("🚀 Bắt đầu chuyển đổi"):
        with st.spinner("Đang tính toán..."):
            df_vertical = transform_horizontal_to_vertical(df_raw)
            
            if df_vertical is not None:
                st.subheader("2. Kết quả sau khi chuyển đổi (Bảng dọc)")
                st.success(f"Đã xử lý xong {len(df_vertical)} dòng dữ liệu.")
                
                # Hiển thị kết quả
                st.dataframe(df_vertical, use_container_width=True)
                
                # Nút tải file
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_vertical.to_excel(writer, index=False, sheet_name='Ket_qua_doc')
                
                st.download_button(
                    label="📥 Tải file kết quả Excel",
                    data=output.getvalue(),
                    file_name="excel_vertical_result.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

# --- PHẦN MỞ RỘNG: AI HỖ TRỢ PHÂN TÍCH (Tùy chọn giống file HTML) ---
st.sidebar.header("AI Assistant")
api_key = st.sidebar.text_input("Nhập Gemini API Key (nếu muốn dùng AI)", type="password")
if api_key and uploaded_file:
    import google.generativeai as genai
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel('gemini-1.5-flash')
    
    user_q = st.sidebar.text_area("Hỏi AI về dữ liệu này:")
    if st.sidebar.button("Hỏi AI"):
        prompt = f"Dưới đây là dữ liệu Excel: {df_raw.iloc[:10, :10].to_string()}... \nCâu hỏi: {user_q}"
        response = model.generate_content(prompt)
        st.sidebar.write(response.text)
