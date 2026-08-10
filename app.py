import io
import os
import re
import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

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
      ws_bg = writer.book.create_sheet(title="BÁO GIÁ", index=0)

      df_final.to_excel(writer, index=False, sheet_name="CHI TIẾT", startrow=3)
      ws_ct = writer.sheets["CHI TIẾT"]

      for ws in [ws_bg, ws_ct]:
        try:
          ws.views.sheetView[0].sheetViewType = "pageBreakPreview"
          ws.views.sheetView[0].showGridLines = True
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
        ws_ct.cell(row=r, column=11).number_format = "0%"
        ws_ct.cell(row=r, column=12).number_format = num_format_vnd

        cell_dg = ws_ct.cell(row=r, column=8)
        cell_dg.value = f"=ROUNDUP(L{r}/(1-K{r}), -3)"
        cell_dg.number_format = num_format_vnd

        cell_tt = ws_ct.cell(row=r, column=9)
        cell_tt.value = f"=G{r}*H{r}"
        cell_tt.number_format = num_format_vnd

        cell_tt_cost_tb = ws_ct.cell(row=r, column=13)
        cell_tt_cost_tb.value = f"=G{r}*L{r}"
        cell_tt_cost_tb.number_format = num_format_vnd

        ws_ct.cell(row=r, column=14).number_format = num_format_vnd

        cell_tt_cost_ld = ws_ct.cell(row=r, column=15)
        cell_tt_cost_ld.value = f"=G{r}*N{r}"
        cell_tt_cost_ld.number_format = num_format_vnd

      # Dòng Tổng Cộng ở Sheet Chi Tiết
      tot_row_ct = last_row_ct + 1

      # Gộp ô từ cột STT (1) đến cột Đơn Giá (8)
      ws_ct.merge_cells(
          start_row=tot_row_ct, start_column=1, end_row=tot_row_ct, end_column=8
      )
      cell_total_label = ws_ct.cell(row=tot_row_ct, column=1, value="TỔNG CỘNG")
      cell_total_label.font = Font(name="Times New Roman", size=11, bold=True)
      cell_total_label.alignment = Alignment(
          horizontal="center", vertical="center"
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

          # Tô màu xám (gray_fill) cho toàn bộ dòng Tổng cộng
          if r == tot_row_ct:
            cell.fill = gray_fill

        cell_j = ws_ct.cell(row=r, column=10)
        cell_j.border = Border(
            left=cell_j.border.left,
            top=cell_j.border.top,
            right=blue_thick_side,
            bottom=cell_j.border.bottom,
        )
    
      # =========================================================
      # B. XỬ LÝ DỮ LIỆU & FORMAT SHEET BÁO GIÁ
      # =========================================================
      ws_bg.page_setup.orientation = ws_bg.ORIENTATION_PORTRAIT
      ws_bg.page_setup.paperSize = ws_bg.PAPERSIZE_A4
      ws_bg.sheet_properties.pageSetUpPr.fitToPage = True
      ws_bg.page_setup.fitToWidth = 1
      ws_bg.page_setup.fitToHeight = 0

      ws_bg.merge_cells("A1:H1")
      ws_bg["A1"] = "CÔNG TY TNHH CÔNG NGHỆ DVC"
      ws_bg["A1"].font = Font(name="Times New Roman", size=12, bold=True)
      ws_bg["A1"].alignment = Alignment(horizontal="center")

      ws_bg.merge_cells("A2:H2")
      ws_bg["A2"] = "**********"
      ws_bg["A2"].alignment = Alignment(horizontal="center")

      ws_bg.merge_cells("A3:H3")
      ws_bg["A3"] = "Hotline: 0909 661 579"
      ws_bg["A3"].font = Font(name="Times New Roman", size=10)
      ws_bg["A3"].alignment = Alignment(horizontal="center")

      ws_bg.merge_cells("A4:H4")
      ws_bg["A4"] = "Email: dvc@dvctech.vn - Website: dvctech.vn"
      ws_bg["A4"].font = Font(
          name="Times New Roman", size=10, underline="single"
      )
      ws_bg["A4"].alignment = Alignment(horizontal="center")

      ws_bg.merge_cells("A5:H5")
      ws_bg["A5"] = "BẢNG BÁO GIÁ"
      ws_bg["A5"].font = Font(name="Times New Roman", size=20, bold=True)
      ws_bg["A5"].alignment = Alignment(horizontal="center")

      ws_bg["A6"] = "Kính gửi:"
      ws_bg["G6"] = "Người gửi:"
      ws_bg["A7"] = "Người nhận:"
      ws_bg["G7"] = "Điện thoại:"
      ws_bg["A8"] = "Email/Sdt:"

      for cell_id in ["A6", "G6", "A7", "G7", "A8"]:
        ws_bg[cell_id].font = Font(name="Times New Roman", size=11, bold=True)

      ws_bg.merge_cells("A9:H9")
      ws_bg["A9"] = "Nội dung:"
      ws_bg["A9"].font = Font(name="Times New Roman", size=11, bold=True)
      # --- TẠO VIỀN NGOÀI (OUTLINE) CHO KHỐI A6:H9 ---
      thin_side = Side(style="thin", color="000000")
      
      for r in range(6, 10):  # Dòng từ 6 đến 9
        for c in range(1, 9):  # Cột từ A (1) đến H (8)
          cell = ws_bg.cell(row=r, column=c)
      
          # Chỉ gán viền ở 4 mép ngoài cùng của vùng A6:H9
          cell.border = Border(
              top=thin_side if r == 6 else cell.border.top,
              bottom=thin_side if r == 9 else cell.border.bottom,
              left=thin_side if c == 1 else cell.border.left,
              right=thin_side if c == 8 else cell.border.right,
          )
      ws_bg.merge_cells("A10:H10")
      ws_bg["A10"] = (
          "Cảm ơn Quý khách hàng đã quan tâm và tin tưởng sản phẩm và dịch vụ"
          " của công ty chúng tôi. Chúng tôi hân hạnh gửi đến Quý khách hàng"
          " bảng chào giá như sau:"
      )
      ws_bg["A10"].font = Font(name="Times New Roman", size=11, italic=True)

      headers_bg = [
          ("A11", "Stt"),
          ("B11", "Nội dung báo giá"),
          ("D11", "ĐVT"),
          ("E11", "Số lượng"),
          ("F11", "Đơn giá\n(VNĐ)"),
          ("G11", "Thuế GTGT"),
          ("H11", "Thành tiền\n(VNĐ)"),
      ]

      ws_bg.merge_cells("B11:C11")
      for cell_id, text in headers_bg:
        c = ws_bg[cell_id]
        c.value = text
        c.font = Font(name="Times New Roman", size=11, bold=True)
        c.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )

      ws_bg["A12"] = 1
      ws_bg["A12"].alignment = Alignment(horizontal="center")

      ws_bg.merge_cells("B12:C12")
      ws_bg["B12"] = "Hệ thống"
      ws_bg["D12"] = "Hệ thống"
      ws_bg["D12"].alignment = Alignment(horizontal="center")
      ws_bg["E12"] = 1
      ws_bg["E12"].alignment = Alignment(horizontal="center")

      ws_bg["F12"] = f"='CHI TIẾT'!I{tot_row_ct}"
      ws_bg["F12"].number_format = num_format_vnd
      ws_bg["F12"].alignment = Alignment(horizontal="right")

      ws_bg["G12"] = "=F12*8%"
      ws_bg["G12"].number_format = num_format_vnd
      ws_bg["G12"].alignment = Alignment(horizontal="right")

      ws_bg["H12"] = "=F12+G12"
      ws_bg["H12"].number_format = num_format_vnd
      ws_bg["H12"].alignment = Alignment(horizontal="right")

      for r in range(11, 13):
        for col_idx in range(1, 9):
          ws_bg.cell(row=r, column=col_idx).border = thin_border

      ws_bg.merge_cells("A13:H13")
      ws_bg["A13"] = (
          "Ghi chú: Thuế GTGT tạm tính, được điều chỉnh theo quy định tại thời"
          " điểm xuất hóa đơn."
      )
      ws_bg["A13"].font = Font(name="Times New Roman", size=11)

      ws_bg.merge_cells("A14:H14")
      ws_bg["A14"] = "Điều kiện thương mại:"
      ws_bg["A14"].font = Font(
          name="Times New Roman", size=11, bold=True, underline="single"
      )

      ws_bg.merge_cells("A15:H15")
      ws_bg["A15"] = "1. Địa điểm thực hiện:"
      ws_bg["A15"].font = Font(name="Times New Roman", size=11, bold=True)
      
      # --- Chừa 1 dòng trống bên dưới cho mục 1 để điền địa điểm ---
      ws_bg.merge_cells("A16:H16")
      ws_bg["A16"] = "   - "  # Hoặc để trống "" để nhập liệu sau
      ws_bg["A16"].font = Font(name="Times New Roman", size=11, bold=False)
      
      ws_bg.merge_cells("A17:H17")
      ws_bg["A17"] = "2. Giá đã bao gồm:"
      ws_bg["A17"].font = Font(name="Times New Roman", size=11, bold=True)
      
      ws_bg.merge_cells("A18:H18")
      ws_bg["A18"] = "   - Chi phí vận chuyển, lắp đặt hệ thống do bên Bán chịu."
      ws_bg["A18"].font = Font(name="Times New Roman", size=11, bold=False)
      
      ws_bg.merge_cells("A19:H19")
      ws_bg["A19"] = "3. Thanh toán:"
      ws_bg["A19"].font = Font(name="Times New Roman", size=11, bold=True)
      
      ws_bg.merge_cells("A20:H20")
      ws_bg["A20"] = (
          "   - Thanh toán 100% giá trị hợp đồng trong vòng 07 ngày làm việc sau khi"
          " hoàn thành lắp đặt, nghiệm thu."
      )
      ws_bg["A20"].font = Font(name="Times New Roman", size=11, bold=False)
      
      ws_bg.merge_cells("A21:H21")
      ws_bg["A21"] = "4. Thời gian thực hiện hợp đồng:"
      ws_bg["A21"].font = Font(name="Times New Roman", size=11, bold=True)
      
      ws_bg.merge_cells("A22:H22")
      ws_bg["A22"] = (
          "   - Thời gian thực hiện: trong vòng 07 ngày kể từ ngày ký hợp đồng."
      )
      ws_bg["A22"].font = Font(name="Times New Roman", size=11, bold=False)
      
      ws_bg.merge_cells("A23:H23")
      ws_bg["A23"] = "5. Thời gian bảo hành:"
      ws_bg["A23"].font = Font(name="Times New Roman", size=11, bold=True)
      
      ws_bg.merge_cells("A24:H24")
      ws_bg["A24"] = (
          "   - BH lắp đặt hệ thống: 12 tháng kể từ ngày nghiệm thu, bàn giao."
      )
      ws_bg["A24"].font = Font(name="Times New Roman", size=11, bold=False)
      
      ws_bg.merge_cells("A25:H25")
      ws_bg["A25"] = (
          "   - BH thiết bị theo chính sách của hãng sản xuất (xem bảng giá chi"
          " tiết)."
      )
      ws_bg["A25"].font = Font(name="Times New Roman", size=11, bold=False)
      
      ws_bg.merge_cells("A26:H26")
      ws_bg["A26"] = "6. Thời hạn chào giá:"
      ws_bg["A26"].font = Font(name="Times New Roman", size=11, bold=True)
        
      ws_bg["A27"] = (
          "   - 30 ngày"
      )
      ws_bg["A27"].font = Font(name="Times New Roman", size=11, bold=False)
        
      ws_bg.merge_cells("A28:H28")
      ws_bg["A28"] = "Chúng tôi rất mong nhận được sự hợp tác với Quý khách hàng!"
      ws_bg["A28"].font = Font(name="Times New Roman", size=11, italic=True)
      
      # Dòng tên công ty đẩy từ G29 xuống G30 do chèn thêm 1 dòng ở trên
      ws_bg["G30"] = "Công ty TNHH Công Nghệ DVC"
      ws_bg["G30"].font = Font(name="Times New Roman", size=11, bold=True)
      ws_bg["G30"].alignment = Alignment(horizontal="center")
      col_widths_bg = {
          "A": 6,
          "B": 16,
          "C": 16,
          "D": 10,
          "E": 10,
          "F": 15,
          "G": 14,
          "H": 16,
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
