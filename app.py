import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Tool Chuyển Đổi Mẫu Excel", layout="wide", page_icon="📊")

st.title("📊 Công Cụ Đổi Mẫu File Excel Tự Động")
st.markdown("Tải file Excel nguồn của bạn lên -> Hệ thống tự động chuyển đổi sang mẫu tiêu chuẩn.")

uploaded_file = st.file_uploader("Chọn file Excel cần chuyển đổi (.xlsx, .xls)", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        df_raw = pd.read_excel(uploaded_file)
        
        with st.expander("👁️ Xem dữ liệu file gốc tải lên"):
            st.dataframe(df_raw.head())

        # Tạo DataFrame mẫu chuẩn
        df_target = pd.DataFrame()
        
        # Ánh xạ từ khóa tên cột từ file gốc
        col_map = {
            'STT': ['stt', 'số thứ tự', 'no', 'tt'],
            'Tên thiết bị': ['tên thiết bị', 'thiết bị', 'tên hàng', 'mô tả', 'sản phẩm'],
            'Mã hàng': ['mã hàng', 'mã sản phẩm', 'mã', 'part number', 'code'],
            'Số lượng': ['số lượng', 'sl', 'qty', 'count'],
            'Đơn giá': ['giá cost', 'đơn giá', 'cost', 'giá', 'đơn giá cost', 'đơn giá mua']
        }
        
        # Tự động gán dữ liệu
        for std_col, keywords in col_map.items():
            found = False
            for c in df_raw.columns:
                if str(c).strip().lower() in keywords:
                    df_target[std_col] = df_raw[c]
                    found = True
                    break
            if not found:
                df_target[std_col] = ""

        # Ép kiểu số & Tính Thành tiền = Số lượng * Đơn giá
        df_target['Số lượng'] = pd.to_numeric(df_target['Số lượng'], errors='coerce').fillna(0)
        df_target['Đơn giá'] = pd.to_numeric(df_target['Đơn giá'], errors='coerce').fillna(0)
        df_target['Thành tiền'] = df_target['Số lượng'] * df_target['Đơn giá']

        # Thứ tự cột đầu ra chuẩn
        final_cols = ['STT', 'Tên thiết bị', 'Mã hàng', 'Số lượng', 'Đơn giá', 'Thành tiền']
        df_target = df_target.reindex(columns=final_cols)

        st.success("✅ Đã chuyển đổi dữ liệu thành công theo form chuẩn!")
        st.subheader("📋 Kết quả chuyển đổi:")
        st.dataframe(df_target)

        # Xuất file Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_target.to_excel(writer, index=False, sheet_name='Form_Chuan')
            
        st.download_button(
            label="📥 Tải file Excel mẫu chuẩn (.xlsx)",
            data=output.getvalue(),
            file_name="File_Excel_Da_Chuyen_Doi.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
        
    except Exception as e:
        st.error(f"Lỗi xử lý file: {e}")
