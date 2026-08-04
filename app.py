import io
import math
import pandas as pd
import streamlit as st

# Cấu hình giao diện Web
st.set_page_config(
    page_title="Tool Xuất Báo Giá Chi Tiết", layout="wide", page_icon="📊"
)

st.title("📊 Công Cụ Đổ Dữ Liệu Vào Mẫu BẢNG GIÁ CHI TIẾT")
st.markdown(
    "Upload **File Dữ Liệu Nguồn (Hình 2)** -> Tool tự động map dữ liệu, tính"
    " toán công thức và xuất ra **Bảng Giá Chi Tiết (Hình 1)**."
)

# Khung tải file
uploaded_file = st.file_uploader(
    "Tải lên File Nguồn dữ liệu nhập vào (.xlsx, .xls)", type=["xlsx", "xls"]
)


def roundup_thousand(val):
  """Hàm làm tròn lên hàng nghìn tương đương =ROUNDUP(val, -3) trong Excel"""
  if pd.isna(val) or val <= 0:
    return 0
  return math.ceil(val / 1000) * 1000


if uploaded_file is not None:
  try:
    # --- 1. ĐỌC FILE VỚI CÁC PHƯƠNG ÁN TỰ ĐỘNG THÍCH ỨNG ---
    df_source = None

    # Phương án 1: Đọc .xlsx chuẩn
    try:
      df_source = pd.read_excel(uploaded_file, engine="openpyxl")
    except Exception:
      pass

    # Phương án 2: Đọc .xls chuẩn (Excel 97-2003)
    if df_source is None:
      try:
        uploaded_file.seek(0)
        df_source = pd.read_excel(uploaded_file, engine="xlrd")
      except Exception:
        pass

    # Phương án 3: Đọc file XML/HTML đổi đuôi (Xuất từ phần mềm Kế toán/ERP)
    if df_source is None:
      try:
        uploaded_file.seek(0)
        dfs = pd.read_html(uploaded_file)
        df_source = dfs[0]
      except Exception:
        pass

    # Phương án 4: Đọc file CSV/Text (nếu có)
    if df_source is None:
      for enc in ["utf-8", "cp1252", "latin1", "utf-16", "utf-8-sig"]:
        try:
          uploaded_file.seek(0)
          df_source = pd.read_csv(uploaded_file, encoding=enc)
          break
        except Exception:
          continue

    if df_source is None:
      st.error(
          "⚠️ Không thể đọc định dạng file này. Vui lòng mở file bằng Excel,"
          " bấm 'Save As' và chọn định dạng 'Excel Workbook (*.xlsx)'."
      )
    else:
      with st.expander("👁️ Xem trước dữ liệu file nguồn vừa upload"):
        st.dataframe(df_source.head())

      # --- 2. ÁNH XẠ DỮ LIỆU SANG FORM MẪU CHI TIẾT (HÌNH 1) ---
      df_final = pd.DataFrame()

      def get_source_col(possible_names):
        """Hàm tìm kiếm tên cột linh hoạt từ File Nguồn"""
        for col in df_source.columns:
          if str(col).strip().lower() in [
              name.lower() for name in possible_names
          ]:
            return df_source[col]
        return ""

      # Lấy dữ liệu cơ bản
      df_final["STT"] = get_source_col(["STT", "No"])
      df_final["Thiết bị"] = get_source_col(["Thiết bị", "Tên thiết bị"])
      df_final["Mã hàng"] = get_source_col(["Mã hàng", "Part Number"])
      df_final["Hình ảnh"] = ""  # Để trống cột hình ảnh
      df_final["Hãng / Xuất xứ"] = get_source_col(
          ["Hãng / Xuất xứ", "Xuất xứ", "Hãng"]
      )
      df_final["ĐVT"] = get_source_col(["ĐVT", "Đơn vị tính"])

      # Chuyển đổi dữ liệu số
      df_final["Số lượng"] = pd.to_numeric(
          get_source_col(["Số lượng", "SL"]), errors="coerce"
      ).fillna(0)
      df_final["Ghi chú"] = get_source_col(["Ghi chú"])

      df_final["Margin Thiết bị"] = pd.to_numeric(
          get_source_col(["Margin Thiết bị", "Margin"]), errors="coerce"
      ).fillna(0)
      df_final["ĐG COST Thiết bị"] = pd.to_numeric(
          get_source_col(["ĐG COST Thiết bị", "DG COST Thiet bi"]),
          errors="coerce",
      ).fillna(0)
      df_final["ĐG COST Lắp đặt"] = pd.to_numeric(
          get_source_col(["ĐG COST Lắp đặt", "DG COST Lap dat"]),
          errors="coerce",
      ).fillna(0)

      df_final["NCC"] = get_source_col(["NCC", "Nhà cung cấp"])
      df_final["NOTE"] = get_source_col(["NOTE", "Note"])

      # --- 3. CÔNG THỨC TÍNH TOÁN ---

      # Tính Đơn giá (VNĐ) = ROUNDUP( ĐG COST Thiết bị / (1 - Margin Thiết bị) ; -3 )
      def calc_don_gia(row):
        cost = row["ĐG COST Thiết bị"]
        margin = row["Margin Thiết bị"]
        if margin >= 1:
          margin = margin / 100  # Đổi dạng 20 sang 0.2
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

      # --- 4. SẮP XẾP ĐÚNG THỨ TỰ CÁC CỘT THEO HÌNH 1 ---
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

      # Hiển thị kết quả trên Web
      st.success(
          "✅ Đã xử lý và tính toán thành công dữ liệu sang Form BẢNG GIÁ CHI"
          " TIẾT!"
      )
      st.subheader("📋 Kết quả File Cuối Cùng (Mẫu Hình 1):")
      st.dataframe(df_final)

      # --- 5. TẠO NÚT TẢI FILE EXCEL HOÀN CHỈNH ---
      output = io.BytesIO()
      with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_final.to_excel(writer, index=False, sheet_name="BANG_GIA_CHI_TIET")

      st.download_button(
          label="📥 Tải File BẢNG GIÁ CHI TIẾT Hoàn Chỉnh (.xlsx)",
          data=output.getvalue(),
          file_name="Bang_Gia_Chi_Tiet_Final.xlsx",
          mime=(
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          ),
          type="primary",
      )

  except Exception as e:
    st.error(f"Lỗi khi xử lý dữ liệu: {e}")
