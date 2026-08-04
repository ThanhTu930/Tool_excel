import io
import math
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Tool Xuất Báo Giá Theo Mẫu", layout="wide", page_icon="📊"
)

st.title("📊 Công Cụ Đổ Dữ Liệu Vào Mẫu Báo Giá Chi Tiết")
st.markdown(
    "Upload **File Dữ Liệu Nguồn (Hình 2)** -> Tool tự động map dữ liệu, tính"
    " toán công thức và xuất ra **Bảng Giá Chi Tiết (Hình 1)**."
)

# Upload File Nguồn (Hình 2)
uploaded_file = st.file_uploader(
    "Tải lên File Nguồn dữ liệu nhập vào (.xlsx, .xls)", type=["xlsx", "xls"]
)


def roundup_thousand(val):
  """Làm tròn lên đến hàng nghìn =ROUNDUP(val, -3)"""
  if pd.isna(val) or val <= 0:
    return 0
  return math.ceil(val / 1000) * 1000


if uploaded_file is not None:
  try:
    # 1. Đọc dữ liệu từ File Nguồn (Hình 2)
    df_source = pd.read_excel(uploaded_file)

    with st.expander("👁️ Xem trước dữ liệu file nguồn vừa upload"):
      st.dataframe(df_source.head())

    # 2. Tạo Khung Dữ Liệu Theo Đúng Cấu Trúc Form Mẫu Đích (Hình 1)
    df_final = pd.DataFrame()

    # Hàm hỗ trợ tìm tên cột linh hoạt từ File Nguồn
    def get_source_col(possible_names):
      for col in df_source.columns:
        if str(col).strip().lower() in [
            name.lower() for name in possible_names
        ]:
          return df_source[col]
      return ""

    # --- LẤY DỮ LIỆU TỪ HÌNH 2 ĐỔ SANG HÌNH 1 ---
    df_final["STT"] = get_source_col(["STT", "No"])
    df_final["Thiết bị"] = get_source_col(["Thiết bị", "Tên thiết bị"])
    df_final["Mã hàng"] = get_source_col(["Mã hàng", "Part Number"])
    df_final["Hình ảnh"] = ""  # Cột Hình ảnh để trống để chèn sau
    df_final["Hãng / Xuất xứ"] = get_source_col(["Hãng / Xuất xứ", "Xuất xứ"])
    df_final["ĐVT"] = get_source_col(["ĐVT", "Đơn vị tính"])

    # Ép kiểu dữ liệu số để tính toán
    df_final["Số lượng"] = pd.to_numeric(
      get_source_col(["Số lượng", "SL"]), errors="coerce"
    ).fillna(0)
    df_final["Ghi chú"] = get_source_col(["Ghi chú"])

    df_final["Margin Thiết bị"] = pd.to_numeric(
      get_source_col(["Margin Thiết bị", "Margin"]), errors="coerce"
    ).fillna(0)
    df_final["ĐG COST Thiết bị"] = pd.to_numeric(
      get_source_col(["ĐG COST Thiết bị", "DG COST Thiet bi"]), errors="coerce"
    ).fillna(0)
    df_final["ĐG COST Lắp đặt"] = pd.to_numeric(
      get_source_col(["ĐG COST Lắp đặt", "DG COST Lap dat"]), errors="coerce"
    ).fillna(0)

    df_final["NCC"] = get_source_col(["NCC", "Nhà cung cấp"])
    df_final["NOTE"] = get_source_col(["NOTE", "Note"])

    # --- TÍNH TOÁN CÁC CỘT TỰ ĐỘNG CHO FORM HÌNH 1 ---

    # Tính Đơn giá (VNĐ) = ROUNDUP( ĐG COST Thiết bị / (1 - Margin Thiết bị) ; -3 )
    def calc_don_gia(row):
      cost = row["ĐG COST Thiết bị"]
      margin = row["Margin Thiết bị"]
      # Nếu margin nhập dạng 20 thì tự chia 100 thành 0.2
      if margin >= 1:
        margin = margin / 100
      if (1 - margin) <= 0:
        return 0
      return roundup_thousand(cost / (1 - margin))

    df_final["Đơn giá (VNĐ)"] = df_final.apply(calc_don_gia, axis=1)

    # Tính Thành tiền (VNĐ) = Số lượng * Đơn giá (VNĐ)
    df_final["Thành tiền (VNĐ)"] = (
        df_final["Số lượng"] * df_final["Đơn giá (VNĐ)"]
    )

    # Tính TT COST Thiết bị = Số lượng * ĐG COST Thiết bị
    df_final["TT COST Thiết bị"] = (
        df_final["Số lượng"] * df_final["ĐG COST Thiết bị"]
    )

    # Tính TT COST Lắp đặt = Số lượng * ĐG COST Lắp đặt
    df_final["TT COST Lắp đặt"] = (
        df_final["Số lượng"] * df_final["ĐG COST Lắp đặt"]
    )

    # --- SẮP XẾP CHUẨN THỨ TỰ CÁC CỘT THEO FORM HÌNH 1 ---
    form_1_columns = [
        "STT",
        "Thiết bị",
        "Mã hàng",
        "Hình ảnh",
        "Hãng / Xuất xứ",
        "ĐVT",
        "Số lượng",
        "Đơn giá (VNĐ)",
        "Thành tiền (VNĐ)",
        "Ghi chú",
        "Margin Thiết bị",
        "ĐG COST Thiết bị",
        "TT COST Thiết bị",
        "ĐG COST Lắp đặt",
        "TT COST Lắp đặt",
        "NCC",
        "NOTE",
    ]
    df_final = df_final.reindex(columns=form_1_columns)

    # Display Preview
    st.success(
        "✅ Đã xử lý và tính toán thành công dữ liệu sang Form BẢNG GIÁ CHI"
        " TIẾT!"
    )
    st.subheader("📋 Kết quả File Cuối Cùng (Mẫu Hình 1):")
    st.dataframe(df_final)

    # --- XUẤT FILE EXCEL CUỐI CÙNG ---
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
      df_final.to_excel(writer, index=False, sheet_name="BANG_GIA_CHI_TIET")

    st.download_button(
        label="📥 Tải File BẢNG GIÁ CHI TIẾT Hoàn Chỉnh (.xlsx)",
        data=output.getvalue(),
        file_name="Bang_Gia_Chi_Tiet_Final.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

  except Exception as e:
    st.error(f"Lỗi khi xử lý dữ liệu: {e}")
