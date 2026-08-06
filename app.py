import io
import re
import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

st.set_page_config(
    page_title="TOOL CHUYỂN ĐỔI BÁO GIÁ", layout="wide", page_icon="📊"
)

st.title("TẢI FILE MẪU ĐỂ NHẬP DỮ LIỆU")


# --- 1. HÀM TẠO FILE FORM MẪU ĐỂ TẢI VỀ ---
def generate_sample_template():
  sample_df = pd.DataFrame({
      "Stt": [],
      "Tên thiết bị": [],
      "Mã hàng": [],
      "Mô tả": [],
      "Nhãn hiệu/Xuất xứ": [],
      "ĐVT": [],
      "Số lượng": [],
      "Ghi chú": [],
      "Margin": [],
      "Giá Cost": [],
      "Giá lắp đặt": [],
      "NCC": [],
  })

  output = io.BytesIO()
  with pd.ExcelWriter(output, engine="openpyxl") as writer:
    sample_df.to_excel(writer, index=False, sheet_name="FORM_MAU")

    # Format sơ bộ cho file mẫu đẹp mắt
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

    for col in range(1, 13):
      cell = worksheet.cell(row=1, column=col)
      cell.fill = gray_fill
      cell.font = Font(bold=True)
      cell.alignment = Alignment(horizontal="center", vertical="center")
      cell.border = thin_border

    for row in range(2, 4):
      for col in range(1, 13):
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

  clean_str = re.sub(r"[^\d]", "", val_str)
  try:
    return float(clean_str) if clean_str else 0.0
  except ValueError:
    return 0.0


# --- 3. GIAO DIỆN HƯỚNG DẪN & NÚT TẢI FORM MẪU ---
col1, col2 = st.columns([3, 1])

with col2:
  # NÚT TẢI FILE FORM MẪU ĐẦU VÀO
  sample_file_data = generate_sample_template()
  st.download_button(
      label="📥 Tải Form Mẫu Đầu Vào (.xlsx)",
      data=sample_file_data,
      file_name="Form_Mau_Nhap_Lieu_Dau_Vao.xlsx",
      mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      use_container_width=True,
  )

st.divider()

# --- 4. UPLOAD FILE ĐẦU VÀO ---
uploaded_file = st.file_uploader(
    "📥 **Tải file Excel dữ liệu đầu vào của bạn tại đây (.xlsx, .xls):**",
    type=["xlsx", "xls"],
)

if uploaded_file is not None:
  try:
    input_df = pd.read_excel(uploaded_file)

    st.subheader("👀 Xem trước dữ liệu vừa tải lên:")
    st.dataframe(input_df, use_container_width=True)

    cols = [str(c).strip() for c in input_df.columns]
    input_df.columns = cols

    def get_col_val(df, possible_names, default=""):
      for name in possible_names:
        for c in df.columns:
          if name.lower() in c.lower():
            return df[c]
      return pd.Series([default] * len(df))

    if st.button("🚀 Xử Lý & Xuất File BẢNG GIÁ CHI TIẾT Chuẩn", type="primary"):
      df_final = pd.DataFrame()

      df_final["STT"] = get_col_val(input_df, ["stt"], 1)
      df_final["Thiết bị"] = get_col_val(
          input_df,
          ["tên thiết bị", "tên hàng hóa/dịch vụ", "tên hàng hóa", "thiết bị"],
          "",
      )
      df_final["Mã hàng"] = get_col_val(input_df, ["mã hàng"], "")
      df_final["Mô tả"] = get_col_val(input_df, ["mô tả"], "")
      df_final["Hình ảnh"] = ""
      df_final["Hãng / Xuất xứ"] = get_col_val(
          input_df, ["nhãn hiệu/xuất xứ", "xuất xứ", "hãng"], ""
      )
      df_final["ĐVT"] = get_col_val(input_df, ["đvt"], "Cái")

      raw_sl = get_col_val(input_df, ["số lượng"], 1)
      df_final["Số lượng"] = raw_sl.apply(clean_currency)

      df_final["Đơn giá (VNĐ)"] = 0
      df_final["Thành tiền (VNĐ)"] = 0
      df_final["Ghi chú"] = get_col_val(input_df, ["ghi chú"], "")

      raw_margin = get_col_val(input_df, ["margin"], 0)
      margin_clean = raw_margin.apply(clean_currency)
      df_final["Margin Thiết bị"] = margin_clean.apply(
          lambda x: x / 100.0 if x >= 1.0 else x
      )

      raw_cost = get_col_val(input_df, ["giá cost", "cost thiết bị"], 0)
      df_final["ĐG COST Thiết bị"] = raw_cost.apply(clean_currency)
      df_final["TT COST Thiết bị"] = 0

      raw_lapdat = get_col_val(input_df, ["giá lắp đặt", "cost lắp đặt"], 0)
      df_final["ĐG COST Lắp đặt"] = raw_lapdat.apply(clean_currency)
      df_final["TT COST Lắp đặt"] = 0

      df_final["NCC"] = get_col_val(input_df, ["ncc"], "")
      df_final["NOTE"] = ""

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
          "Ghi chú",
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
        title_cell.font = Font(name="Times New Roman", size=18, bold=True)
        title_cell.alignment = Alignment(
            horizontal="center", vertical="center"
        )

        num_format_vnd = "#,##0"

        for i in range(len(df_final)):
          r = 5 + i

          # Công thức Đơn giá ROUNDUP
          cell_don_gia = worksheet.cell(row=r, column=9)
          cell_don_gia.value = (
              f"=ROUNDUP(M{r}/(1 -L{r}),-3)"
          )
          cell_don_gia.number_format = num_format_vnd

          # Thành tiền = Số lượng * Đơn giá
          cell_thanh_tien = worksheet.cell(row=r, column=10)
          cell_thanh_tien.value = f"=H{r}*I{r}"
          cell_thanh_tien.number_format = num_format_vnd

          worksheet.cell(row=r, column=13).number_format = num_format_vnd

          cell_tt_cost_tb = worksheet.cell(row=r, column=14)
          cell_tt_cost_tb.value = f"=H{r}*M{r}"
          cell_tt_cost_tb.number_format = num_format_vnd

          worksheet.cell(row=r, column=15).number_format = num_format_vnd

          cell_tt_cost_ld = worksheet.cell(row=r, column=16)
          cell_tt_cost_ld.value = f"=H{r}*O{r}"
          cell_tt_cost_ld.number_format = num_format_vnd

          # Margin hiển thị dạng % số nguyên (ví dụ 30%)
          worksheet.cell(row=r, column=12).number_format = "0%"

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

      st.success("🎉 Đã xử lý xong!")
      st.download_button(
          label="📥 Tải File BẢNG GIÁ CHI TIẾT (.xlsx)",
          data=output.getvalue(),
          file_name="BG - DVC.xlsx",
          mime=(
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          ),
          type="primary",
      )

  except Exception as e:
    st.error(f"⚠️ Có lỗi khi xử lý file: {e}")
