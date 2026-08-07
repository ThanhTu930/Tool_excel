import io
import os
import re
import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

# --- 1. CẤU HÌNH STREAMLIT (ĐẶT LÊN ĐẦU TIÊN) ---
st.set_page_config(page_title="Tool nhập liệu DVCTECH", layout="wide")

# --- 2. CSS CHỈNH CHỮ TO CHO NÚT BẤM VÀ KHUNG UPLOAD ---
st.markdown(
    """
    <style>
    /* 1. Đổi màu nút Tải Form Mẫu ở trên */
    div.stDownloadButton > button {
        background-color: Green !important; /* Màu xanh lá chuẩn */
        color: #FFFFFF !important;              /* Mặc định chữ màu trắng */
        font-weight: bold !important;          /* Chữ in đậm */
        border-radius: 6px !important;          /* Bo góc nút */
        border: none !important;
        padding: 0.5rem 1.2rem !important;
    }
    div.stDownloadButton > button:hover {
        background-color: DarkGreen !important; /* Xanh đậm khi rê chuột */
        color: #FFFFFF !important;
    }
    div.stDownloadButton > button p {
        font-size: 18px !important;  /* Cỡ chữ nút bấm */
        font-weight: bold !important; /* Chữ đậm */
    }
    div.stDownloadButton > button[kind="primary"] {
        background-color: Red !important; /* Màu đỏ chuẩn */
        color: #FFFFFF !important;              /* Mặc định chữ màu trắng */
        font-weight: bold !important;          /* Chữ in đậm */
        border-radius: 6px !important;          /* Bo góc nút */
        border: none !important;
        padding: 0.5rem 1.2rem !important;
    }
    div.stDownloadButton > button[kind="primary"]:hover {
        background-color: DarkRed !important; /* Đỏ đậm khi rê chuột */
        color: #FFFFFF !important;
    }
    div.stDownloadButton > button[kind="primary"] p {
        font-size: 18px !important;  /* Cỡ chữ nút bấm */
        font-weight: bold !important; /* Chữ đậm */
    }
    /* 2. Đổi màu nút Upload bên trong khung Tải File ở dưới */
    div[data-testid="stFileUploader"] button {
        background-color: Blue !important; /* Màu xanh lá đồng bộ */
        color: #FFFFFF !important;              /* Chữ màu trắng */
        font-weight: bold !important;          /* Chữ in đậm */
        border-radius: 6px !important;          /* Bo góc nút */
        border: none !important;
    }
    div[data-testid="stFileUploader"] button:hover {
        background-color: darkblue !important; /* Xanh đậm khi rê chuột */
        color: #FFFFFF !important;
    }
    div[data-testid="stFileUploader"] p {
        font-size: 18px !important;  /* Cỡ chữ nút bấm */
        font-weight: bold !important; /* Chữ đậm */
    }
    div[data-testid="stMarkdownContainer"] p {
    font-size: 36px !important;
    font-weight: bold !important;
    }
    </style>
""",
    unsafe_allow_html=True,
)

st.title("TOOL NHẬP LIỆU BẢNG BÁO GIÁ")


# --- 3. HÀM TẠO FILE FORM MẪU ĐỂ TẢI VỀ ---
def generate_sample_template():
  sample_df = pd.DataFrame({
      "Stt": [],
      "Thiết bị": [],
      "Mã hàng": [],
      "Hãng/Xuất xứ": [],  # Đổi vị trí lên trước Mô tả
      "Mô tả": [],
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


# --- 4. HÀM LÀM SẠCH DỮ LIỆU SỐ ---
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


# --- 5. PHẦN TẢI FORM MẪU (Ở TRÊN) ---
st.write("TẢI FILE BÁO GIÁ MẪU ĐỂ NHẬP THEO FORM CỦA HỆ THỐNG")

sample_file_data = generate_sample_template()
st.download_button(
    label="FORM BÁO GIÁ MẪU (.xlsx)",
    data=sample_file_data,
    file_name="BG Mẫu - DVC.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=False,
)

st.divider()

# --- 6. PHẦN UPLOAD FILE (Ở DƯỚI) ---
st.write("UPLOAD FILE EXCEL TẠI ĐÂY (.xlsx, .xls):")

uploaded_file = st.file_uploader(
    "Upload", type=["xlsx", "xls"], label_visibility="collapsed"
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

    # --- TỰ ĐỘNG XỬ LÝ DỮ LIỆU NGAY KHI UPLOAD ---
    df_final = pd.DataFrame()

    df_final["STT"] = get_col_val(input_df, ["stt"], 1)
    df_final["Thiết bị"] = get_col_val(
        input_df,
        ["thiết bị", "tên thiết bị", "tên hàng hóa/dịch vụ", "tên hàng hóa"],
        "",
    )
    df_final["Mã hàng"] = get_col_val(input_df, ["mã hàng"], "")

    # ĐƯA HÃNG/XUẤT XỨ LÊN TRƯỚC MÔ TẢ
    df_final["Hãng/Xuất xứ"] = get_col_val(
        input_df, ["hãng/xuất xứ", "nhãn hiệu/xuất xứ", "xuất xứ", "hãng"], ""
    )
    df_final["Mô tả"] = get_col_val(input_df, ["mô tả"], "")
    df_final["Hình ảnh"] = ""

    df_final["ĐVT"] = get_col_val(input_df, ["đvt"], "Cái")

    raw_sl = get_col_val(input_df, ["số lượng"], 1)
    df_final["Số lượng"] = raw_sl.apply(clean_currency)

    df_final["Đơn giá (VNĐ)"] = None
    df_final["Thành tiền (VNĐ)"] = None
    df_final["Thời gian bảo hành"] = get_col_val(
        input_df, ["Thời gian bảo hành"], ""
    )

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

    # THỨ TỰ CỘT TRONG FILE EXCEL XUẤT RA
    form_columns = [
        "STT",  # Col 1 (A)
        "Thiết bị",  # Col 2 (B)
        "Mã hàng",  # Col 3 (C)
        "Hãng/Xuất xứ",  # Col 4 (D) -> Đã chuyển lên trước Mô tả
        "Mô tả",  # Col 5 (E)
        "ĐVT",  # Col 6 (F)
        "Số lượng",  # Col 7 (G)
        "Đơn giá (VNĐ)",  # Col 8 (H)
        "Thành tiền (VNĐ)",  # Col 9 (I)
        "Thời gian bảo hành",  # Col 10 (J)
        "Margin Thiết bị",  # Col 11 (K)
        "ĐG COST Thiết bị",  # Col 12 (L)
        "TT COST Thiết bị",  # Col 13 (M)
        "ĐG COST Lắp đặt",  # Col 14 (N)
        "TT COST Lắp đặt",  # Col 15 (O)
        "NCC",  # Col 16 (P)
        "NOTE",  # Col 17 (Q)
    ]
    df_final = df_final.reindex(columns=form_columns)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
      df_final.to_excel(
          writer, index=False, sheet_name="CHI TIẾT", startrow=3
      )

      workbook = writer.book
      worksheet = writer.sheets["CHI TIẾT"]

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
      title_cell.alignment = Alignment(horizontal="center", vertical="center")

      num_format_vnd = "#,##0"

      for i in range(len(df_final)):
        r = 5 + i

        # K (Cột 11) = Margin Thiết bị (Format %)
        worksheet.cell(row=r, column=11).number_format = "0%"

        # L (Cột 12) = ĐG COST Thiết bị (Format VNĐ)
        worksheet.cell(row=r, column=12).number_format = num_format_vnd

        # H (Cột 8) = Đơn giá bán = ROUNDUP(L / (1 - K), -3)
        cell_don_gia = worksheet.cell(row=r, column=8)
        cell_don_gia.value = f"=ROUNDUP(L{r}/(1-K{r}), -3)"
        cell_don_gia.number_format = num_format_vnd

        # I (Cột 9) = Thành tiền = Số lượng (G) * Đơn giá (H)
        cell_thanh_tien = worksheet.cell(row=r, column=9)
        cell_thanh_tien.value = f"=G{r}*H{r}"
        cell_thanh_tien.number_format = num_format_vnd

        # M (Cột 13) = TT COST Thiết bị = Số lượng (G) * ĐG COST TB (L)
        cell_tt_cost_tb = worksheet.cell(row=r, column=13)
        cell_tt_cost_tb.value = f"=G{r}*L{r}"
        cell_tt_cost_tb.number_format = num_format_vnd

        # N (Cột 14) = ĐG COST Lắp đặt (Format VNĐ)
        worksheet.cell(row=r, column=14).number_format = num_format_vnd

        # O (Cột 15) = TT COST Lắp đặt = Số lượng (G) * ĐG COST Lắp đặt (N)
        cell_tt_cost_ld = worksheet.cell(row=r, column=15)
        cell_tt_cost_ld.value = f"=G{r}*N{r}"
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

      # Cột 1 -> 10: Chữ đen (Cho khách hàng)
      for col in range(1, 11):
        cell = worksheet.cell(row=4, column=col)
        cell.fill = gray_fill
        cell.font = Font(
            name="Times New Roman", size=12, bold=True, color="000000"
        )
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )
        cell.border = thin_border

      # Cột 11 -> 17: Chữ đỏ (Nội bộ / Cost)
      for col in range(11, 18):
        cell = worksheet.cell(row=4, column=col)
        cell.fill = gray_fill
        cell.font = Font(
            name="Times New Roman", size=12, bold=True, color="FF0000"
        )
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )
        cell.border = thin_border

      # Kẻ viền từ cột 1 đến 17
      for row in range(5, len(df_final) + 5):
        for col in range(1, 18):
          c = worksheet.cell(row=row, column=col)
          c.font = Font(name="Times New Roman", size=10)
          c.border = thin_border

        # Viền xanh đứng phân cách ở sau cột 10 (J - Thời gian bảo hành)
        cell_j = worksheet.cell(row=row, column=10)
        cell_j.border = Border(
            left=cell_j.border.left,
            top=cell_j.border.top,
            right=blue_thick_side,
            bottom=cell_j.border.bottom,
        )

    name_without_ext, ext = os.path.splitext(uploaded_file.name)
    output_filename = f"{name_without_ext} - R1{ext}"

    # --- NÚT XUẤT FILE DUY NHẤT: BẤM LÀ TẢI FILE TRỰC TIẾP VỀ MÁY ---
    st.download_button(
        label="XUẤT FILE BẢNG GIÁ CHI TIẾT",
        data=output.getvalue(),
        file_name=output_filename,
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
        type="primary",
    )

  except Exception as e:
    st.error(f"⚠️ Có lỗi khi xử lý file: {e}")
