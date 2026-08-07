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
    font-size: 18px !important;
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

    form_columns = [
        "STT",
        "Thiết bị",
        "Mã hàng",
        "Hãng/Xuất xứ",
        "Mô tả",
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
      # 1. TẠO SHEET 1: BÁO GIÁ (Tổng thể)
      ws_bg = writer.book.create_sheet(title="BÁO GIÁ", index=0)

      # 2. TẠO SHEET 2: CHI TIẾT
      df_final.to_excel(writer, index=False, sheet_name="CHI TIẾT", startrow=3)
      ws_ct = writer.sheets["CHI TIẾT"]

      # --- CẤU HÌNH DISPLAY & PAGE BREAK ---
      for ws in [ws_bg, ws_ct]:
        try:
          ws.views.sheetView[0].sheetViewType = "pageBreakPreview"
        except Exception:
          pass

      # =========================================================
      # A. XỬ LÝ FORMAT VÀ CÔNG THỨC SHEET CHI TIẾT
      # =========================================================
      ws_ct.merge_cells("B2:J2")
      title_ct = ws_ct["B2"]
      title_ct.value = "BẢNG GIÁ CHI TIẾT"
      title_ct.font = Font(name="Times New Roman", size=26, bold=True)
      title_ct.alignment = Alignment(horizontal="center", vertical="center")

      num_format_vnd = "#,##0"
      last_row_ct = 4 + len(df_final)

      for i in range(len(df_final)):
        r = 5 + i
        ws_ct.cell(row=r, column=11).number_format = "0%"  # Margin
        ws_ct.cell(row=r, column=12).number_format = (
            num_format_vnd  # ĐG COST TB
        )

        # H = ROUNDUP(L / (1 - K), -3)
        cell_dg = ws_ct.cell(row=r, column=8)
        cell_dg.value = f"=ROUNDUP(L{r}/(1-K{r}), -3)"
        cell_dg.number_format = num_format_vnd

        # I = G * H (Thành tiền)
        cell_tt = ws_ct.cell(row=r, column=9)
        cell_tt.value = f"=G{r}*H{r}"
        cell_tt.number_format = num_format_vnd

        # M = G * L
        cell_tt_cost_tb = ws_ct.cell(row=r, column=13)
        cell_tt_cost_tb.value = f"=G{r}*L{r}"
        cell_tt_cost_tb.number_format = num_format_vnd

        ws_ct.cell(row=r, column=14).number_format = (
            num_format_vnd  # ĐG COST Lắp đặt
        )

        # O = G * N
        cell_tt_cost_ld = ws_ct.cell(row=r, column=15)
        cell_tt_cost_ld.value = f"=G{r}*N{r}"
        cell_tt_cost_ld.number_format = num_format_vnd

      # Dòng Tổng Cộng ở Sheet Chi Tiết
      tot_row_ct = last_row_ct + 1
      ws_ct.cell(row=tot_row_ct, column=2, value="TỔNG CỘNG").font = Font(
          name="Times New Roman", size=11, bold=True
      )
      ws_ct.cell(
          row=tot_row_ct, column=9, value=f"=SUM(I5:I{last_row_ct})"
      ).number_format = num_format_vnd
      ws_ct.cell(
          row=tot_row_ct, column=13, value=f"=SUM(M5:M{last_row_ct})"
      ).number_format = num_format_vnd
      ws_ct.cell(
          row=tot_row_ct, column=15, value=f"=SUM(O5:O{last_row_ct})"
      ).number_format = num_format_vnd

      # Style Header & Viền cho Sheet Chi Tiết
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

      for col in range(1, 11):
        cell = ws_ct.cell(row=4, column=col)
        cell.fill, cell.border = gray_fill, thin_border
        cell.font = Font(
            name="Times New Roman", size=12, bold=True, color="000000"
        )
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )

      for col in range(11, 18):
        cell = ws_ct.cell(row=4, column=col)
        cell.fill, cell.border = gray_fill, thin_border
        cell.font = Font(
            name="Times New Roman", size=12, bold=True, color="FF0000"
        )
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )

      for r in range(5, tot_row_ct + 1):
        for c in range(1, 18):
          cell = ws_ct.cell(row=r, column=c)
          cell.font = Font(
              name="Times New Roman",
              size=10,
              bold=(True if r == tot_row_ct else False),
          )
          cell.border = thin_border

        cell_j = ws_ct.cell(row=r, column=10)
        cell_j.border = Border(
            left=cell_j.border.left,
            top=cell_j.border.top,
            right=blue_thick_side,
            bottom=cell_j.border.bottom,
        )

      # =========================================================
      # B. XỬ LÝ DỮ LIỆU & FORMAT SHEET BÁO GIÁ (SHEET TỔNG THỂ)
      # =========================================================
      # Header Công Ty
      ws_bg["C1"] = "CÔNG TY TNHH CÔNG NGHỆ DVC"
      ws_bg["C1"].font = Font(name="Times New Roman", size=12, bold=True)
      ws_bg["C1"].alignment = Alignment(horizontal="center")

      ws_bg["C2"] = "**********"
      ws_bg["C2"].alignment = Alignment(horizontal="center")

      ws_bg["C3"] = "Hotline: 0909 661 579"
      ws_bg["C3"].font = Font(name="Times New Roman", size=10)
      ws_bg["C3"].alignment = Alignment(horizontal="center")

      ws_bg["C4"] = "Email: dvc@dvctech.vn - Website: dvctech.vn"
      ws_bg["C4"].font = Font(name="Times New Roman", size=10, underline="single")
      ws_bg["C4"].alignment = Alignment(horizontal="center")

      # Tiêu đề
      ws_bg.merge_cells("A6:G6")
      ws_bg["A6"] = "BẢNG BÁO GIÁ"
      ws_bg["A6"].font = Font(name="Times New Roman", size=20, bold=True)
      ws_bg["A6"].alignment = Alignment(horizontal="center")

      # Thông tin khách hàng & người gửi
      ws_bg["A8"] = "Kính gửi:"
      ws_bg["F8"] = "Người gửi:"
      ws_bg["A9"] = "Người nhận:"
      ws_bg["F9"] = "Điện thoại:"
      ws_bg["A10"] = "Email/Sdt:"

      for cell_id in ["A8", "F8", "A9", "F9", "A10"]:
        ws_bg[cell_id].font = Font(name="Times New Roman", size=11, bold=True)

      ws_bg["A12"] = "Nội dung:"
      ws_bg["A12"].font = Font(name="Times New Roman", size=11, bold=True)

      ws_bg.merge_cells("A14:G14")
      ws_bg["A14"] = (
          "Cảm ơn Quý khách hàng đã quan tâm và tin tưởng sản phẩm và dịch vụ"
          " của công ty chúng tôi. Chúng tôi hân hạnh gửi đến Quý khách hàng"
          " bảng chào giá như sau:"
      )
      ws_bg["A14"].font = Font(name="Times New Roman", size=11, italic=True)

      # Header Bảng Giá
      headers_bg = [
          ("A16", "Stt"),
          ("B16", "Nội dung báo giá"),
          ("D16", "ĐVT"),
          ("E16", "Số lượng"),
          ("F16", "Đơn giá\n(VNĐ)"),
          ("G16", "Thành tiền\n(VNĐ)"),
      ]

      ws_bg.merge_cells("B16:C16")
      for cell_id, text in headers_bg:
        c = ws_bg[cell_id]
        c.value = text
        c.font = Font(name="Times New Roman", size=11, bold=True)
        c.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )

      # Dòng dữ liệu hệ thống (Hàng 18)
      ws_bg["A18"] = 1
      ws_bg["A18"].alignment = Alignment(horizontal="center")

      ws_bg.merge_cells("B18:C18")
      ws_bg["B18"] = "Hệ thống"
      ws_bg["D18"] = "Hệ thống"
      ws_bg["D18"].alignment = Alignment(horizontal="center")
      ws_bg["E18"] = 1
      ws_bg["E18"].alignment = Alignment(horizontal="center")

      # CÔNG THỨC LẤY ĐƠN GIÁ TỪ SHEET CHI TIẾT (Lấy từ cell Tổng cộng)
      ws_bg["F18"] = f"='CHI TIẾT'!I{tot_row_ct}"
      ws_bg["F18"].number_format = num_format_vnd
      ws_bg["F18"].alignment = Alignment(horizontal="right")

      ws_bg["G18"] = "=E18*F18"
      ws_bg["G18"].number_format = num_format_vnd
      ws_bg["G18"].alignment = Alignment(horizontal="right")

      # Đóng khung bảng dữ liệu báo giá
      for r in range(16, 19):
        for col_idx in range(1, 8):
          ws_bg.cell(row=r, column=col_idx).border = thin_border

      # Ghi chú & Điều kiện thương mại
      ws_bg["A20"] = "Ghi chú:"
      ws_bg["A20"].font = Font(name="Times New Roman", size=11, bold=True)
      ws_bg["B20"] = (
          "Thuế GTGT tạm tính, được điều chỉnh theo quy định tại thời điểm xuất"
          " hóa đơn."
      )
      ws_bg["B20"].font = Font(name="Times New Roman", size=11)

      ws_bg["A22"] = "Điều kiện thương mại:"
      ws_bg["A22"].font = Font(
          name="Times New Roman", size=11, bold=True, underline="single"
      )

      terms = [
          ("A23", "1. Địa điểm thực hiện:"),
          ("A25", "2. Giá đã bao gồm:"),
          ("B26", "- Chi phí vận chuyển, lắp đặt hệ thống do bên Bán chịu."),
          ("A27", "3. Thanh toán:"),
          (
              "B28",
              (
                  "- Thanh toán 100% giá trị hợp đồng trong vòng 07 ngày làm"
                  " việc sau khi hoàn thành lắp đặt, nghiệm thu."
              ),
          ),
          ("A29", "4. Thời gian thực hiện hợp đồng:"),
          ("B30", "- Thời gian thực hiện: trong vòng 07 ngày kể từ ngày ký hợp đồng."),
          ("A31", "5. Thời gian bảo hành:"),
          ("B32", "- BH lắp đặt hệ thống: 12 tháng kể từ ngày nghiệm thu, bàn giao."),
          (
              "B33",
              (
                  "- BH thiết bị theo chính sách của hãng sản xuất (xem bảng giá"
                  " chi tiết)."
              ),
          ),
          ("A34", "6. Thời hạn chào giá: 30 ngày."),
          ("A36", "Chúng tôi rất mong nhận được sự hợp tác với Quý khách hàng!"),
      ]

      for cell_id, text in terms:
        c = ws_bg[cell_id]
        c.value = text
        is_bold = True if cell_id[0] == "A" and cell_id != "A36" else False
        is_italic = True if cell_id == "A36" else False
        c.font = Font(
            name="Times New Roman", size=11, bold=is_bold, italic=is_italic
        )

      # Chữ ký đại diện công ty
      ws_bg["F38"] = "Công ty TNHH Công Nghệ DVC"
      ws_bg["F38"].font = Font(name="Times New Roman", size=12, bold=True)
      ws_bg["F38"].alignment = Alignment(horizontal="center")

      # Chỉnh độ rộng các cột cho Sheet Báo Giá
      col_widths_bg = {
          "A": 6,
          "B": 25,
          "C": 25,
          "D": 12,
          "E": 10,
          "F": 18,
          "G": 20,
      }
      for col_letter, width in col_widths_bg.items():
        ws_bg.column_dimensions[col_letter].width = width

    # --- TẠO NÚT TẢI FILE HOÀN CHỈNH ---
    name_without_ext, ext = os.path.splitext(uploaded_file.name)
    output_filename = f"{name_without_ext} - R1{ext}"

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
