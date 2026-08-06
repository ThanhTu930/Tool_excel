import io
import os
import re
import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
# --- 1. CSS CHỈNH CHỮ TO CHO NÚT BẤM VÀ KHUNG UPLOAD ---
st.markdown(
    """
    <style>
    /* Tăng cỡ chữ & độ đậm cho nút Download */
    div.stDownloadButton > button {
        min-height: 50px !important; /* Tăng chiều cao nút cho cân đối */
    }
    div.stDownloadButton > button p {
        font-size: 18px !important;  /* Cỡ chữ nút bấm (tùy chỉnh 18px, 20px,...) */
        font-weight: bold !important; /* Chữ đậm */
    }
    </style>
""",
    unsafe_allow_html=True,
)
st.set_page_config(
    page_title="Tool nhập liệu DVCTECH", layout="wide"
)

st.title("TOOL NHẬP LIỆU BẢNG BÁO GIÁ")


# --- 1. HÀM TẠO FILE FORM MẪU ĐỂ TẢI VỀ (ĐÃ CẬP NHẬT CHUẨN THEO ẢNH) ---
def generate_sample_template():
  sample_df = pd.DataFrame({
      "Stt": [],
      "Thiết bị": [],
      "Mã hàng": [],
      "Mô tả": [],
      "Hãng/Xuất xứ": [],
      "ĐVT": [],
      "Số lượng": [],
      "Thời gian bảo hành": [],
      "Margin Thiết bị": [],
      "ĐG COST Thiết bị": [],
      "ĐG COST Lắp đặt": [],
      "NCC": [],
      "NOTE": [],
  })

  output = io.BytesIO()
  with pd.ExcelWriter(output, engine="openpyxl") as writer:
    sample_df.to_excel(writer, index=False, sheet_name="FORM_MAU")

    worksheet = writer.sheets["FORM_MAU"]
    gray_fill = PatternFill(
        start_color="D9D9D9", end_color="D9D9D9", fill_type="solid"
    )
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )

    # Format 13 cột của form mẫu
    for col in range(1, 14):
      cell = worksheet.cell(row=1, column=col)
      cell.fill = gray_fill
      cell.font = Font(bold=True)
      cell.alignment = Alignment(horizontal="center", vertical="center")
      cell.border = thin_border

    for row in range(2, 4):
      for col in range(1, 14):
        worksheet.cell(row=row, column=col).border = thin_border

  return output.getvalue()


# --- 2. HÀM LÀM SẠCH DỮ LIỆU SỐ ---
def clean_currency(val):
  if pd.isna(val) or val == "" or val is None:
    return 0.0
  if isinstance(val, (int, float)):
    return float(val)

  val_str = str(val).strip()
  if "%" in val_str:
    val_str = val_str.replace("%", "").strip()
    try:
      return float(val_str) / 100.0
    except ValueError:
      return 0.0

  clean_str = re.sub(r"[^\d.]", "", val_str.replace(",", "."))
  try:
    return float(clean_str) if clean_str else 0.0
  except ValueError:
    return 0.0


# --- 3. GIAO DIỆN HƯỚNG DẪN & NÚT TẢI FORM MẪU ---
st.markdown(
    " Tải file BG mẫu để nhập dữ liệu theo form của hệ thống:"
)
sample_file_data = generate_sample_template()
st.download_button(
    label="FORM BÁO GIÁ MẪU (.xlsx)",
    data=sample_file_data,
    file_name="BG Mẫu - DVC.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True
)
st.divider()

# --- 4. UPLOAD FILE ĐẦU VÀO ---
st.markdown("UPLOAD FILE EXCEL TẠI ĐÂY (.xlsx, .xls):")
uploaded_file = st.file_uploader(
    type=["xlsx", "xls"],
)

if uploaded_file is not None:
  try:
    input_df = pd.read_excel(uploaded_file)

    st.subheader("Xem trước dữ liệu vừa tải lên:")
    st.dataframe(input_df, use_container_width=True, hide_index=True)

    cols = [str(c).strip() for c in input_df.columns]
    input_df.columns = cols

    def get_col_val(df, possible_names, default=""):
      for name in possible_names:
        for c in df.columns:
          if name.lower() in c.lower():
            return df[c]
      return pd.Series([default] * len(df))

    if st.button("XUẤT FILE BẢNG GIÁ CHI TIẾT", type="primary"):
      df_final = pd.DataFrame()

      df_final["STT"] = get_col_val(input_df, ["stt"], 1)
      df_final["Thiết bị"] = get_col_val(
          input_df,
          ["thiết bị", "tên thiết bị", "tên hàng hóa/dịch vụ", "tên hàng hóa"],
          "",
      )
      df_final["Mã hàng"] = get_col_val(input_df, ["mã hàng"], "")
      df_final["Mô tả"] = get_col_val(input_df, ["mô tả"], "")
      df_final["Hình ảnh"] = ""
      df_final["Hãng/Xuất xứ"] = get_col_val(
          input_df, ["hãng/xuất xứ", "nhãn hiệu/xuất xứ", "xuất xứ", "hãng"], ""
      )
      df_final["ĐVT"] = get_col_val(input_df, ["đvt"], "Cái")

      raw_sl = get_col_val(input_df, ["số lượng"], 1)
      df_final["Số lượng"] = raw_sl.apply(clean_currency)

      df_final["Đơn giá (VNĐ)"] = None
      df_final["Thành tiền (VNĐ)"] = None
      df_final["Thời gian bảo hành"] = get_col_val(input_df, ["Thời gian bảo hành"], "")

      raw_margin = get_col_val(
          input_df, ["margin thiết bị", "margin tb", "margin"], 0
      )
      margin_clean = raw_margin.apply(clean_currency)
      df_final["Margin Thiết bị"] = margin_clean.apply(
          lambda x: x / 100.0 if x >= 1.0 else x
      )

      raw_cost = get_col_val(
          input_df, ["đg cost thiết bị", "cost thiết bị", "giá cost"], 0
      )
      df_final["ĐG COST Thiết bị"] = raw_cost.apply(clean_currency)
      df_final["TT COST Thiết bị"] = None

      raw_lapdat = get_col_val(
          input_df, ["đg cost lắp đặt", "cost lắp đặt", "giá lắp đặt"], 0
      )
      df_final["ĐG COST Lắp đặt"] = raw_lapdat.apply(clean_currency)
      df_final["TT COST Lắp đặt"] = None

      df_final["NCC"] = get_col_val(input_df, ["ncc"], "")
      df_final["NOTE"] = get_col_val(input_df, ["note"], "")

      form_columns = [
          "STT",
          "Thiết bị",
          "Mã hàng",
          "Mô tả",
          "Hãng/Xuất xứ",
          "ĐVT",
          "Số lượng",
          "Đơn giá (VNĐ)",
          "Thành tiền (VNĐ)",
          "Thời gian bảo hành",
          "Margin Thiết bị",
          "ĐG COST Thiết bị",
          "TT COST Thiết bị",
          "ĐG COST Lắp đặt",
          "TT COST Lắp đặt",
          "NCC",
          "NOTE",
      ]
      df_final = df_final.reindex(columns=form_columns)

      output = io.BytesIO()
      with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_final.to_excel(
            writer, index=False, sheet_name="BANG_GIA_CHI_TIET", startrow=3
        )

        workbook = writer.book
        worksheet = writer.sheets["BANG_GIA_CHI_TIET"]

        try:
          if not hasattr(worksheet, "sheet_views") or not worksheet.sheet_views:
            worksheet.views.sheetView[0].sheetViewType = "pageBreakPreview"
          else:
            worksheet.sheet_views[0].sheetViewType = "pageBreakPreview"
        except Exception:
          pass

        worksheet.merge_cells("B2:J2")
        title_cell = worksheet["B2"]
        title_cell.value = "BẢNG GIÁ CHI TIẾT"
        title_cell.font = Font(name="Times New Roman", size=26, bold=True)
        title_cell.alignment = Alignment(
            horizontal="center", vertical="center"
        )

        num_format_vnd = "#,##0"

        for i in range(len(df_final)):
          r = 5 + i

          # L (Cột 12) = Margin
          worksheet.cell(row=r, column=12).number_format = "0%"

          # M (Cột 13) = ĐG COST Thiết bị
          worksheet.cell(row=r, column=13).number_format = num_format_vnd

          # I (Cột 9) = Đơn giá bán = ROUNDUP(Cost / (1 - Margin), -3)
          cell_don_gia = worksheet.cell(row=r, column=9)
          cell_don_gia.value = (
              f"=ROUNDUP(M{r}/(1-L{r}),-3)"
          )
          cell_don_gia.number_format = num_format_vnd

          # J (Cột 10) = Thành tiền = Số lượng * Đơn giá
          cell_thanh_tien = worksheet.cell(row=r, column=10)
          cell_thanh_tien.value = f"=H{r}*I{r}"
          cell_thanh_tien.number_format = num_format_vnd

          # N (Cột 14) = TT COST Thiết bị = Số lượng * ĐG COST Thiết bị
          cell_tt_cost_tb = worksheet.cell(row=r, column=14)
          cell_tt_cost_tb.value = f"=H{r}*M{r}"
          cell_tt_cost_tb.number_format = num_format_vnd

          # O (Cột 15) = ĐG COST Lắp đặt
          worksheet.cell(row=r, column=15).number_format = num_format_vnd

          # P (Cột 16) = TT COST Lắp đặt = Số lượng * ĐG COST Lắp đặt
          cell_tt_cost_ld = worksheet.cell(row=r, column=16)
          cell_tt_cost_ld.value = f"=H{r}*O{r}"
          cell_tt_cost_ld.number_format = num_format_vnd

        # Format Header & Viền
        gray_fill = PatternFill(
            start_color="D9D9D9", end_color="D9D9D9", fill_type="solid"
        )
        thin_border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin"),
        )
        blue_thick_side = Side(style="medium", color="0000FF")

        for col in range(1, 12):
          cell = worksheet.cell(row=4, column=col)
          cell.fill = gray_fill
          cell.font = Font(
              name="Times New Roman", size=10, bold=True, color="000000"
          )
          cell.alignment = Alignment(
              horizontal="center", vertical="center", wrap_text=True
          )
          cell.border = thin_border

        for col in range(12, 19):
          cell = worksheet.cell(row=4, column=col)
          cell.fill = gray_fill
          cell.font = Font(
              name="Times New Roman", size=10, bold=True, color="FF0000"
          )
          cell.alignment = Alignment(
              horizontal="center", vertical="center", wrap_text=True
          )
          cell.border = thin_border

        for row in range(5, len(df_final) + 5):
          for col in range(1, 19):
            c = worksheet.cell(row=row, column=col)
            c.font = Font(name="Times New Roman", size=10)
            c.border = thin_border

          cell_k = worksheet.cell(row=row, column=11)
          cell_k.border = Border(
              left=cell_k.border.left,
              top=cell_k.border.top,
              right=blue_thick_side,
              bottom=cell_k.border.bottom,
          )

      name_without_ext, ext = os.path.splitext(uploaded_file.name)
      output_filename = f"{name_without_ext} - R1{ext}"

      st.success("Đã xử lý xong!")
      st.download_button(
          label="DOWNLOAD FILE TẠI ĐÂY (.xlsx)",
          data=output.getvalue(),
          file_name=output_filename,
          mime=(
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          ),
          type="primary",
      )

  except Exception as e:
    st.error(f"⚠️ Có lỗi khi xử lý file: {e}")
