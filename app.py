import io
import math
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Tool Xuất Báo Giá Theo Mẫu", layout="wide", page_icon="📊"
)

st.title("📊 Công Cụ Đổ Dữ Liệu Vào Mẫu BẢNG GIÁ CHI TIẾT")
st.markdown(
    "Upload **File Dữ Liệu Nguồn (Hình 2)** -> Tool tự động map dữ liệu, tính"
    " toán công thức và xuất ra **Bảng Giá Chi Tiết (Hình 1)**."
)

# Upload File Nguồn
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
    # --- ĐỌC FILE VỚI NHIỀU PHƯƠNG ÁN TỰ ĐỘNG ---
    df_source = None
    try:
      df_source = pd.read_excel(uploaded_file, engine="openpyxl")
    except Exception:
      try:
        uploaded_file.seek(0)
        df_source = pd.read_excel(uploaded_file, engine="xlrd")
      except Exception:
        try:
          uploaded_file.seek(0)
          df_source = pd.read_html(uploaded_file)[0]
        except Exception:
          uploaded_file.seek(0)
          df_source = pd.read_csv(uploaded_file)

    with st.expander("👁️ Xem trước dữ liệu file nguồn vừa upload"):
      st.dataframe(df_source.head())

    # --- KHỞI TẠO DATAFRAME THEO MẪU ĐÍCH ---
    df_final = pd.DataFrame()

    def get_source_col(possible_names):
      for col in df_source.columns:
        if str(col).strip().lower() in [
            name.lower() for name in possible_names
        ]:
          return df_source[col]
      return ""

    # --- LẤY DỮ LIỆU TỪ FILE NGUỒN ---
    df_final["STT"] = get_source_col(["STT", "No"])
    df_final["Thiết bị"] = get_source_col(["Thiết bị", "Tên thiết bị"])
    df_final["Mã hàng"] = get_source_col(["Mã hàng", "Part Number"])
    df_final["Hình ảnh"] = ""
    df_final["Hãng / Xuất xứ"] = get_source_col(["Hãng / Xuất xứ", "Xuất xứ"])
    df_final["ĐVT"] = get_source_col(["ĐVT", "Đơn vị tính"])

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

    # --- CÔNG THỨC TÍNH TOÁN ---
    def calc_don_gia(row):
      cost = row["ĐG COST Thiết bị"]
      margin = row["Margin Thiết bị"]
      if margin >= 1:
        margin = margin / 100
      if (1 - margin) <= 0:
        return 0
      return roundup_thousand(cost / (1 - margin))

    df_final["Đơn giá (VNĐ)"] = df_final.apply(calc_don_gia, axis=1)
    df_final["Thành tiền (VNĐ)"] = (
        df_final["Số lượng"] * df_final["Đơn giá (VNĐ)"]
    )
    df_final["TT COST Thiết bị"] = (
        df_final["Số lượng"] * df_final["ĐG COST Thiết bị"]
    )
    df_final["TT COST Lắp đặt"] = (
        df_final["Số lượng"] * df_final["ĐG COST Lắp đặt"]
    )

    # --- SẮP XẾP CỘT THEO CẤU TRÚC HÌNH 1 ---
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

    st.success(
        "✅ Đã xử lý và tính toán thành công dữ liệu sang Form BẢNG GIÁ CHI"
        " TIẾT!"
    )
    st.subheader("📋 Kết quả File Cuối Cùng (Mẫu Hình 1):")
    st.dataframe(df_final)

    # --- XUẤT FILE EXCEL ---
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
