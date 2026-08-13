import io
import os
import re
import openpyxl
import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image

# --- 1. CẤU HÌNH STREAMLIT ---
st.set_page_config(page_title="Tool nhập liệu DVCTECH", layout="wide")

# --- 2. CSS TÙY CHỈNH GIAO DIỆN ---
st.markdown(
    """
    <style>
    div.stDownloadButton > button {
        background-color: DarkGreen !important;
        color: #FFFFFF !important;
        font-weight: bold !important;
        border-radius: 6px !important;
        border: none !important;
        padding: 0.5rem 1.2rem !important;
    }
    div.stDownloadButton > button:hover {
        background-color: Green !important;
        color: #FFFFFF !important;
    }
    div.stDownloadButton > button p {
        font-size: 18px !important;
        font-weight: bold !important;
    }
    div.stDownloadButton > button[kind="primary"] {
        background-color: DarkRed !important;
        color: #FFFFFF !important;
        font-weight: bold !important;
        border-radius: 6px !important;
        border: none !important;
        padding: 0.5rem 1.2rem !important;
    }
    div.stDownloadButton > button[kind="primary"]:hover {
        background-color: Red !important;
        color: #FFFFFF !important;
    }
    div.stDownloadButton > button[kind="primary"] p {
        font-size: 18px !important;
        font-weight: bold !important;
    }
    div[data-testid="stFileUploader"] button {
        background-color: darkblue !important;
        color: #FFFFFF !important;
        font-weight: bold !important;
        border-radius: 6px !important;
        border: none !important;
    }
    div[data-testid="stFileUploader"] button:hover {
        background-color: #0066CC !important;
        color: #FFFFFF !important;
    }
    div[data-testid="stFileUploader"] p {
        font-size: 18px !important;
        font-weight: bold !important;
    }
    div[data-testid="stMarkdownContainer"] p {
        font-size: 18px !important;
        font-weight: bold !important;
    }
    div.stButton > button[kind="primary"] {
        background-color: DarkBlue  !important;
        color: #FFFFFF !important;
        font-weight: bold !important;
        border-radius: 6px !important;
        border: none !important;
        padding: 0.5rem 1.2rem !important;
    }
    div.stButton > button[kind="primary"]:hover {
        background-color: #0066CC !important;
        color: #FFFFFF !important;
    }
    div.stButton > button[kind="primary"] p {
        font-size: 18px !important;
        font-weight: bold !important;
    }
    div.stButton > button[kind="secondary"] {
        background-color: Gray !important;
        color: #FFFFFF !important;
        font-weight: bold !important;
        border-radius: 6px !important;
        border: none !important;
        padding: 0.5rem 1.2rem !important;
    }
    div.stButton > button[kind="secondary"]:hover {
        background-color: DarkGray !important;
        color: #FFFFFF !important;
    }
    div.stButton > button[kind="secondary"] p {
        font-size: 18px !important;
        font-weight: bold !important;
    }
    </style>
""",
    unsafe_allow_html=True,
)

st.title("TOOL NHẬP LIỆU BẢNG BÁO GIÁ")


# --- 3. HÀM TẠO FILE FORM MẪU ---
def generate_sample_template():
    sample_df = pd.DataFrame({
        "Stt": [],
        "Thiết bị": [],
        "Mã hàng": [],
        "Hãng/\nXuất xứ": [],
        "Mô tả": [],
        "ĐVT": [],
        "Số lượng": [],
        "Thời gian bảo hành": [],
        "Ghi chú": [],
        "Margin Thiết bị": [],
        "ĐG COST Thiết bị": [],
        "Margin Lắp đặt": [],
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

        for col in range(1, 16):
            cell = worksheet.cell(row=1, column=col)
            cell.fill = gray_fill
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border

        for row in range(2, 4):
            for col in range(1, 16):
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


# --- 5. HÀM XỬ LÝ DỮ LIỆU VÀ TẠO FILE EXCEL HOÀN CHỈNH ---
def process_dataframe_and_generate_excel(input_df):
    cols = [str(c).strip() for c in input_df.columns]
    input_df.columns = cols

    def get_col_val(df, possible_names, default=""):
        for name in possible_names:
            for c in df.columns:
                if name.lower() in c.lower():
                    return df[c]
        return pd.Series([default] * len(df))

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
        input_df, ["thời gian bảo hành", "bảo hành"], ""
    )
    df_final["Ghi chú"] = get_col_val(input_df, ["ghi chú"], "")

    # Margin & Cost Thiết bị
    raw_margin = get_col_val(input_df, ["margin thiết bị", "margin tb"], 0)
    margin_clean = raw_margin.apply(clean_currency)
    df_final["Margin Thiết bị"] = margin_clean.apply(
        lambda x: x / 100.0 if x >= 1.0 else x
    )

    raw_cost = get_col_val(
        input_df, ["đg cost thiết bị", "cost thiết bị", "giá cost"], 0
    )
    df_final["ĐG COST Thiết bị"] = raw_cost.apply(clean_currency)
    df_final["TT COST Thiết bị"] = None

    # Margin & Cost Lắp đặt
    raw_margin_ld = get_col_val(input_df, ["margin lắp đặt", "margin ld"], 0)
    margin_ld_clean = raw_margin_ld.apply(clean_currency)
    df_final["Margin Lắp đặt"] = margin_ld_clean.apply(
        lambda x: x / 100.0 if x >= 1.0 else x
    )

    raw_lapdat = get_col_val(
        input_df, ["đg cost lắp đặt", "cost lắp đặt", "giá lắp đặt"], 0
    )
    df_final["ĐG COST Lắp đặt"] = raw_lapdat.apply(clean_currency)
    df_final["TT COST Lắp đặt"] = None

    df_final["NCC"] = get_col_val(input_df, ["ncc"], "")
    df_final["NOTE"] = get_col_val(input_df, ["note"], "")

    form_columns = [
        "STT", "Thiết bị", "Mã hàng", "Hãng/\nXuất xứ", "Mô tả", "ĐVT",
        "Số lượng", "Đơn giá (VNĐ)", "Thành tiền (VNĐ)", "Thời gian bảo hành",
        "Ghi chú", "Margin Thiết bị", "ĐG COST Thiết bị", "TT COST Thiết bị",
        "Margin Lắp đặt", "ĐG COST Lắp đặt", "TT COST Lắp đặt", "NCC", "NOTE"
    ]
    df_final = df_final.reindex(columns=form_columns)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        ws_bg = writer.book.create_sheet(title="BÁO GIÁ", index=0)
        df_final.to_excel(
            writer, index=False, header=False, sheet_name="CHI TIẾT", startrow=5
        )
        ws_ct = writer.sheets["CHI TIẾT"]

        for col_idx, col_name in enumerate(form_columns, 1):
            ws_ct.cell(row=4, column=col_idx, value=col_name)

        # =========================================================
        # A. XỬ LÝ FORMAT VÀ CÔNG THỨC SHEET CHI TIẾT
        # =========================================================
        ws_ct.merge_cells("B2:J2")
        title_ct = ws_ct["B2"]
        title_ct.value = "BẢNG GIÁ CHI TIẾT"
        title_ct.font = Font(name="Times New Roman", size=26, bold=True)
        title_ct.alignment = Alignment(horizontal="center", vertical="center")

        max_len = max([len(str(cell.value or '')) for cell in ws_ct['B']])
        ws_ct.column_dimensions['B'].width = max(max_len + 3, 25)

        num_format_vnd = "#,##0"
        num_items = len(df_final)
        last_item_row = 5 + num_items if num_items > 0 else 6

        row_II = last_item_row + 1
        tot_row_ct = row_II + 1

        # --- 1. MỤC I: HÀNG HÓA/THIẾT BỊ CHÍNH ---
        ws_ct.cell(row=5, column=1, value="I").alignment = Alignment(
            horizontal="center", vertical="center"
        )
        cell_i_tb = ws_ct.cell(row=5, column=2, value="Hàng hóa/Thiết bị chính")
        cell_i_tb.font = Font(name="Times New Roman", size=11, bold=True)

        for col in range(3, 9):
            ws_ct.cell(row=5, column=col).value = None

        cell_i_tt = ws_ct.cell(
            row=5, column=9, value=f"=SUM(I6:I{last_item_row})"
        )
        cell_i_tt.font = Font(name="Times New Roman", size=11, bold=True)
        cell_i_tt.number_format = num_format_vnd
        cell_i_tt.alignment = Alignment(horizontal="right", vertical="center")

        for col in range(10, 20):
            ws_ct.cell(row=5, column=col).value = None

        # --- 2. DÒNG DỮ LIỆU THIẾT BỊ CHÍNH ---
        for i in range(num_items):
            r = 6 + i
            ws_ct.cell(row=r, column=1).alignment = Alignment(horizontal="center", vertical="center")
            ws_ct.cell(row=r, column=6).alignment = Alignment(horizontal="center", vertical="center")
            ws_ct.cell(row=r, column=7).alignment = Alignment(horizontal="center", vertical="center")

            ws_ct.cell(row=r, column=12).number_format = "0%"
            ws_ct.cell(row=r, column=13).number_format = num_format_vnd

            cell_dg = ws_ct.cell(row=r, column=8)
            cell_dg.value = f"=ROUNDUP(M{r}/(1-L{r}), -3)"
            cell_dg.number_format = num_format_vnd

            cell_tt = ws_ct.cell(row=r, column=9)
            cell_tt.value = f"=G{r}*H{r}"
            cell_tt.number_format = num_format_vnd

            cell_tt_cost_tb = ws_ct.cell(row=r, column=14)
            cell_tt_cost_tb.value = f"=G{r}*M{r}"
            cell_tt_cost_tb.number_format = num_format_vnd

            ws_ct.cell(row=r, column=15).number_format = "0%"
            ws_ct.cell(row=r, column=16).number_format = num_format_vnd

            cell_tt_cost_ld = ws_ct.cell(row=r, column=17)
            cell_tt_cost_ld.value = f"=G{r}*P{r}"
            cell_tt_cost_ld.number_format = num_format_vnd

        # --- 3. MỤC II: CHI PHÍ TRIỂN KHAI ---
        ws_ct.cell(row=row_II, column=1, value="II").alignment = Alignment(horizontal="center", vertical="center")
        cell_ii_tb = ws_ct.cell(
            row=row_II,
            column=2,
            value="Chi phí triển khai",
        )
        cell_ii_tb.font = Font(name="Times New Roman", size=11, bold=True)
        cell_ii_tb.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

        for col in range(3, 6):
            ws_ct.cell(row=row_II, column=col).value = None

        cell_ii_dvt = ws_ct.cell(row=row_II, column=6, value="Gói")
        cell_ii_dvt.alignment = Alignment(horizontal="center", vertical="center")
        cell_ii_sl = ws_ct.cell(row=row_II, column=7, value=1)
        cell_ii_sl.alignment = Alignment(horizontal="center", vertical="center")

        ws_ct.cell(row=row_II, column=8, value=0).number_format = num_format_vnd
        ws_ct.cell(row=row_II, column=9, value=f"=G{row_II}*H{row_II}").number_format = num_format_vnd

        for col in range(10, 20):
            ws_ct.cell(row=row_II, column=col).value = None

        # --- 4. DÒNG TỔNG CỘNG SHEET CHI TIẾT ---
        ws_ct.merge_cells(start_row=tot_row_ct, start_column=1, end_row=tot_row_ct, end_column=8)
        cell_total_label = ws_ct.cell(row=tot_row_ct, column=1, value="TỔNG CỘNG")
        cell_total_label.font = Font(name="Times New Roman", size=11, bold=True)
        cell_total_label.alignment = Alignment(horizontal="center", vertical="center")

        ws_ct.cell(row=tot_row_ct, column=9, value=f"=I5+I{row_II}").number_format = num_format_vnd
        
        # Giữ tổng TT COST Thiết bị (Cột N - Col 14) và TT COST Lắp đặt (Cột Q - Col 17) tại dòng TỔNG CỘNG
        for col in range(10, 20):
            if col == 14:
                cell_sum_n = ws_ct.cell(row=tot_row_ct, column=col, value=f"=SUM(N6:N{last_item_row})")
                cell_sum_n.font = Font(name="Times New Roman", size=11, bold=True)
                cell_sum_n.number_format = num_format_vnd
                cell_sum_n.alignment = Alignment(horizontal="right", vertical="center")
            elif col == 17:
                cell_sum_q = ws_ct.cell(row=tot_row_ct, column=col, value=f"=SUM(Q6:Q{last_item_row})")
                cell_sum_q.font = Font(name="Times New Roman", size=11, bold=True)
                cell_sum_q.number_format = num_format_vnd
                cell_sum_q.alignment = Alignment(horizontal="right", vertical="center")
            else:
                ws_ct.cell(row=tot_row_ct, column=col).value = None

        # --- 5. ĐỊNH DẠNG VÀ KẺ BẢNG SHEET CHI TIẾT ---
        gray_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
        thin_border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
        blue_thick_side = Side(style="medium", color="0000FF")

        for col in range(1, 12):
            cell = ws_ct.cell(row=4, column=col)
            cell.fill, cell.border = gray_fill, thin_border
            cell.font = Font(name="Times New Roman", size=12, bold=True, color="000000")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        for col in range(12, 20):
            cell = ws_ct.cell(row=4, column=col)
            cell.fill, cell.border = gray_fill, thin_border
            cell.font = Font(name="Times New Roman", size=12, bold=True, color="FF0000")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
        align_left_wrap = Alignment(horizontal="left", vertical="center", wrap_text=True)

        cols_center = [1, 3, 4, 6, 7, 10, 12, 15]
        cols_left = [2, 5, 11, 18, 19]

        for r in range(5, tot_row_ct + 1):
            is_section_or_total = (r in (5, row_II, tot_row_ct))
            for c in range(1, 20):
                cell = ws_ct.cell(row=r, column=c)
                
                if r == row_II and c >= 12:
                    cell.font = Font(name="Times New Roman", size=10, bold=False)
                else:
                    cell.font = Font(name="Times New Roman", size=10, bold=is_section_or_total)
                
                cell.border = thin_border

                if c in cols_center:
                    cell.alignment = align_center
                elif c in cols_left:
                    cell.alignment = align_left_wrap
                else:
                    cell.alignment = align_right

                if r == tot_row_ct:
                    cell.fill = gray_fill

            cell_j = ws_ct.cell(row=r, column=11)
            cell_j.border = Border(
                left=cell_j.border.left, top=cell_j.border.top,
                right=blue_thick_side, bottom=cell_j.border.bottom
            )

        for row in ws_ct.iter_rows():
            for cell in row:
                if cell.row == 2:
                    continue
                if cell.font:
                    cell.font = Font(
                        name=cell.font.name or "Times New Roman",
                        size=10,
                        bold=cell.font.bold,
                        italic=cell.font.italic,
                        color=cell.font.color,
                    )
                else:
                    cell.font = Font(name="Times New Roman", size=10)

        # =========================================================
        # B. TẠO VÀ XỬ LÝ SHEET GUI_KH (BẢN GỬI KHÁCH HÀNG)
        # =========================================================
        ws_kh = writer.book.create_sheet(title="GUI_KH", index=1)

        ws_kh.merge_cells("A2:K2")
        ws_kh["A2"] = "BẢNG GIÁ CHI TIẾT"
        ws_kh["A2"].font = Font(name="Times New Roman", size=18, bold=True)
        ws_kh["A2"].alignment = Alignment(horizontal="center", vertical="center")

        headers_kh = [
            ("A4", "STT"), ("B4", "Thiết bị"), ("C4", "Mã hàng"), 
            ("D4", "Hãng/\nXuất xứ"), ("E4", "Mô tả"), ("F4", "ĐVT"), 
            ("G4", "Số lượng"), ("H4", "Đơn giá\n(VNĐ)"), 
            ("I4", "Thành tiền\n(VNĐ)"), ("J4", "Thời gian\nbảo hành"), ("K4", "Ghi chú")
        ]

        ws_kh.row_dimensions[4].height = 28
        for cell_id, text in headers_kh:
            cell = ws_kh[cell_id]
            cell.value = text
            cell.font = Font(name="Times New Roman", size=10, bold=True)
            cell.fill = gray_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border

        for r in range(5, tot_row_ct):
            ws_kh.row_dimensions[r].height = ws_ct.row_dimensions[r].height or 20
            
            for col_letter in ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K"]:
                if col_letter == "I":
                    # CỘT THÀNH TIỀN DÙNG CÔNG THỨC ĐƠN GIÁ * SỐ LƯỢNG
                    ws_kh[f"I{r}"] = f"=G{r}*H{r}"
                else:
                    ws_kh[f"{col_letter}{r}"] = f"=IF('CHI TIẾT'!{col_letter}{r}=\"\",\"\",'CHI TIẾT'!{col_letter}{r})"

            is_bold = (r in [5, row_II])
            for col_letter in ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K"]:
                c = ws_kh[f"{col_letter}{r}"]
                c.font = Font(name="Times New Roman", size=10, bold=is_bold)
                c.border = thin_border
                
                if col_letter in ["A", "F", "G", "J"]:
                    c.alignment = Alignment(horizontal="center", vertical="center")
                elif col_letter in ["H", "I"]:
                    c.alignment = Alignment(horizontal="right", vertical="center")
                    c.number_format = num_format_vnd
                else:
                    c.alignment = Alignment(horizontal="left", vertical="center")

        ws_kh.merge_cells(f"A{tot_row_ct}:H{tot_row_ct}")
        ws_kh[f"A{tot_row_ct}"] = "TỔNG CỘNG"
        ws_kh[f"A{tot_row_ct}"].font = Font(name="Times New Roman", size=10, bold=True)
        ws_kh[f"A{tot_row_ct}"].alignment = Alignment(horizontal="center", vertical="center")

        # TỔNG CỘNG TRÊN SHEET GUI_KH TÍNH TỔNG TỪ CỘT I
        ws_kh[f"I{tot_row_ct}"] = f"=SUM(I5:I{tot_row_ct-1})"
        ws_kh[f"I{tot_row_ct}"].font = Font(name="Times New Roman", size=10, bold=True)
        ws_kh[f"I{tot_row_ct}"].alignment = Alignment(horizontal="right", vertical="center")
        ws_kh[f"I{tot_row_ct}"].number_format = num_format_vnd

        for col_idx in range(1, 12):
            col_letter = get_column_letter(col_idx)
            cell = ws_kh[f"{col_letter}{tot_row_ct}"]
            cell.border = thin_border
            cell.fill = gray_fill

        col_widths_kh = {
            "A": 6, "B": 28, "C": 15, "D": 14, "E": 18, 
            "F": 8, "G": 10, "H": 14, "I": 16, "J": 12, "K": 12
        }
        for col_letter, width in col_widths_kh.items():
            ws_kh.column_dimensions[col_letter].width = width

        sheet_order = ["BÁO GIÁ", "GUI_KH", "CHI TIẾT"]
        writer.book._sheets = [writer.book[s] for s in sheet_order if s in writer.book.sheetnames]

        # =========================================================
        # C. XỬ LÝ SHEET BÁO GIÁ (LẤY DỮ LIỆU TỪ SHEET GUI_KH)
        # =========================================================
        ws_bg.page_setup.orientation = ws_bg.ORIENTATION_PORTRAIT
        ws_bg.page_setup.paperSize = ws_bg.PAPERSIZE_A4
        ws_bg.sheet_properties.pageSetUpPr.fitToPage = True
        ws_bg.page_setup.fitToWidth = 1
        ws_bg.page_setup.fitToHeight = 0

        try:
            ws_bg.views.sheetView[0].sheetViewType = "pageBreakPreview"
            ws_bg.views.sheetView[0].showGridLines = True
            ws_ct.views.sheetView[0].sheetViewType = "pageBreakPreview"
            ws_ct.views.sheetView[0].showGridLines = True
            ws_kh.views.sheetView[0].sheetViewType = "pageBreakPreview"
            ws_kh.views.sheetView[0].showGridLines = True
        except Exception:
            pass

        try:
            img = Image("logo_dvc.png")
            img.width = 90
            img.height = 90
            ws_bg.add_image(img, "A1")
        except Exception:
            pass

        ws_bg.merge_cells("A1:G1")
        ws_bg["A1"] = "CÔNG TY TNHH CÔNG NGHỆ DVC"
        ws_bg["A1"].font = Font(name="Times New Roman", size=12, bold=True)
        ws_bg["A1"].alignment = Alignment(horizontal="center", vertical="center")

        ws_bg.merge_cells("A2:G2")
        ws_bg["A2"] = "**********"
        ws_bg["A2"].alignment = Alignment(horizontal="center", vertical="center")

        ws_bg.merge_cells("A3:G3")
        ws_bg["A3"] = "Hotline: 0909 661 579"
        ws_bg["A3"].font = Font(name="Times New Roman", size=10)
        ws_bg["A3"].alignment = Alignment(horizontal="center", vertical="center")

        ws_bg.merge_cells("A4:G4")
        ws_bg["A4"] = "Email: dvc@dvctech.vn - Website: dvctech.vn"
        ws_bg["A4"].font = Font(name="Times New Roman", size=10, underline="single")
        ws_bg["A4"].alignment = Alignment(horizontal="center", vertical="center")

        ws_bg.merge_cells("A5:G5")
        ws_bg["A5"] = "BẢNG BÁO GIÁ"
        ws_bg["A5"].font = Font(name="Times New Roman", size=20, bold=True)
        ws_bg["A5"].alignment = Alignment(horizontal="center", vertical="center")

        ws_bg["A6"] = "Kính gửi:"
        ws_bg["A7"] = "Người nhận:"
        ws_bg["A8"] = "Email/Sdt:"
        for cell_id in ["A6", "A7", "A8"]:
            ws_bg[cell_id].font = Font(name="Times New Roman", size=10, bold=True)
            ws_bg[cell_id].alignment = Alignment(horizontal="left", vertical="center")

        ws_bg["G6"] = "Người gửi:"
        ws_bg["G7"] = "Điện thoại:"
        ws_bg["G8"] = "TPHCM, ngày tháng năm 2026"
        for cell_id2 in ["G6", "G7", "G8"]:
            ws_bg[cell_id2].font = Font(name="Times New Roman", size=10, bold=True)
            ws_bg[cell_id2].alignment = Alignment(horizontal="right", vertical="center")

        ws_bg.merge_cells("A9:G9")
        ws_bg["A9"] = "Nội dung:"
        ws_bg["A9"].font = Font(name="Times New Roman", size=10, bold=True)

        thin_side = Side(style="thin", color="000000")
        for r in range(6, 9):
            for c in range(1, 8):
                cell = ws_bg.cell(row=r, column=c)
                cell.border = Border(
                    top=thin_side if r == 6 else cell.border.top,
                    bottom=thin_side if r == 8 else cell.border.bottom,
                    left=thin_side if c == 1 else cell.border.left,
                    right=thin_side if c == 7 else cell.border.right,
                )
        for c in range(1, 8):
            ws_bg.cell(row=9, column=c).border = Border(bottom=thin_side)

        ws_bg.merge_cells("A10:G10")
        ws_bg["A10"] = (
            "Cảm ơn Quý khách hàng đã quan tâm và tin tưởng sản phẩm và dịch vụ"
            " của công ty chúng tôi. Chúng tôi hân hạnh gửi đến Quý khách hàng"
            " bảng chào giá như sau:"
        )    
        ws_bg["A10"].font = Font(name="Times New Roman", size=10, italic=True)
        ws_bg["A10"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        ws_bg.row_dimensions[10].height = 35

        headers_bg = [
            ("A11", "Stt"),
            ("B11", "Nội dung báo giá"),
            ("C11", "ĐVT"),
            ("D11", "Số lượng"),
            ("E11", "Đơn giá\n(VNĐ)"),
            ("F11", "Thuế GTGT"),
            ("G11", "Thành tiền\n(VNĐ)"),
        ]

        for cell_id, text in headers_bg:
            c = ws_bg[cell_id]
            c.value = text
            c.font = Font(name="Times New Roman", size=10, bold=True)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        ws_bg["A12"] = 1
        ws_bg["A12"].alignment = Alignment(horizontal="center", vertical="center")

        ws_bg["B12"] = "Hệ thống\n(Xem bảng giá chi tiết đính kèm)"
        ws_bg["B12"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

        ws_bg["C12"] = "Hệ thống"
        ws_bg["C12"].alignment = Alignment(horizontal="center", vertical="center")

        ws_bg["D12"] = 1
        ws_bg["D12"].alignment = Alignment(horizontal="center", vertical="center")

        ws_bg["E12"] = f"='GUI_KH'!I{tot_row_ct}"
        ws_bg["E12"].number_format = num_format_vnd
        ws_bg["E12"].alignment = Alignment(horizontal="right", vertical="center")

        ws_bg["F12"] = "=E12*8%"
        ws_bg["F12"].number_format = num_format_vnd
        ws_bg["F12"].alignment = Alignment(horizontal="right", vertical="center")

        ws_bg["G12"] = "=E12+F12"
        ws_bg["G12"].number_format = num_format_vnd
        ws_bg["G12"].alignment = Alignment(horizontal="right", vertical="center")

        ws_bg.row_dimensions[12].height = 35

        for r in range(11, 13):
            for col_idx in range(1, 8):
                cell = ws_bg.cell(row=r, column=col_idx)
                cell.border = thin_border
                if cell.font:
                    cell.font = Font(
                        name="Times New Roman",
                        size=10,
                        bold=cell.font.bold,
                        italic=cell.font.italic
                    )

        terms_bg = [
            ("A13:G13", "Ghi chú: Thuế GTGT tạm tính, được điều chỉnh theo quy định tại thời điểm xuất hóa đơn.", False, True),
            ("A14:G14", "Điều kiện thương mại:", True, False),
            ("A15:G15", "1. Địa điểm thực hiện:", True, False),
            ("A16:G16", "   - Phạm vi Thành phố Hồ Chí Minh", False, False),
            ("A17:G17", "2. Giá đã bao gồm:", True, False),
            ("A18:G18", "   - Chi phí vận chuyển, lắp đặt do bên Bán chịu.", False, False),
            ("A19:G19", "3. Thanh toán:", True, False),
            ("A20:G20", "   - Đợt 1: Tạm ứng 50% giá trị hợp đồng sau khi hợp đồng được ký kết.", False, False),
            ("A21:G21", "   - Đợt 2: Thanh toán 50% giá trị còn lại của hợp đồng trong vòng 15 ngày làm việc sau khi hai Bên ký Biên bản nghiệm thu hoàn thành và đưa vào sử dụng, đồng thời bên Mua đã nhận đầy đủ hồ sơ thanh toán hợp lệ của bên Bán.", False, False),
            ("A22:G22", "4. Thời gian thực hiện hợp đồng:", True, False),
            ("A23:G23", "   - Trong vòng 15 ngày kể từ ngày kí hợp đồng và bên Bán nhận được tạm ứng của bên Mua.", False, False),
            ("A24:G24", "5. Thời gian bảo hành:", True, False),
            ("A25:G25", "   - BH lắp đặt hệ thống: 12 tháng kể từ ngày nghiệm thu, bàn giao.", False, False),
            ("A26:G26", "   - BH thiết bị theo chính sách của hãng sản xuất (xem bảng giá chi tiết).", False, False),
            ("A27:G27", "6. Thời hạn chào giá:", True, False),
            ("A29:G29", "Chúng tôi rất mong nhận được sự hợp tác với Quý khách hàng!", False, True),
        ]

        for range_str, text, is_bold, is_italic in terms_bg:
            ws_bg.merge_cells(range_str)
            first_cell = ws_bg[range_str.split(":")[0]]
            first_cell.value = text
            first_cell.font = Font(
                name="Times New Roman", size=10, bold=is_bold, italic=is_italic,
                underline="single" if text == "Điều kiện thương mại:" else None
            )
            first_cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        
        ws_bg.row_dimensions[21].height = 35

        ws_bg["A28"] = "   - 30 ngày"
        ws_bg["A28"].font = Font(name="Times New Roman", size=10, bold=False)
        ws_bg["A28"].alignment = Alignment(horizontal="left", vertical="center")

        ws_bg["F31"] = "Công ty TNHH Công Nghệ DVC"
        ws_bg["F31"].font = Font(name="Times New Roman", size=10, bold=True)
        ws_bg["F31"].alignment = Alignment(horizontal="center", vertical="center")

        col_widths_bg = {"A": 6, "B": 32, "C": 10, "D": 10, "E": 15, "F": 14, "G": 16}
        for col_letter, width in col_widths_bg.items():
            ws_bg.column_dimensions[col_letter].width = width

    return output.getvalue()


# --- 6. GIAO DIỆN TẢI FORM MẪU & NHẬP TRỰC TIẾP ---
if "show_manual_input" not in st.session_state:
    st.session_state.show_manual_input = False

col_template, col_manual_btn = st.columns([1, 1])

with col_template:
    st.write("TẢI FILE BÁO GIÁ MẪU ĐỂ NHẬP THEO FORM HỆ THỐNG:")
    sample_file_data = generate_sample_template()
    st.download_button(
        label="Form Báo Giá Mẫu (.xlsx)",
        data=sample_file_data,
        file_name="BG Mẫu - DVC.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

with col_manual_btn:
    st.write("HOẶC NHẬP DỮ LIỆU BÁO GIÁ TRỰC TIẾP TRÊN WEB:")
    if st.button("Nhập Báo Giá Trực Tiếp", use_container_width=True, type="secondary"):
        st.session_state.show_manual_input = not st.session_state.show_manual_input

# --- 7. KHUNG NHẬP DỮ LIỆU BÁO GIÁ TRỰC TIẾP ---
if st.session_state.show_manual_input:
    st.markdown("---")
    st.subheader("Bảng nhập thông tin báo giá trực tiếp")
    st.caption("Gõ thông tin trực tiếp vào bảng dưới đây. Bấm nút (+) ở dưới cùng bảng để thêm dòng mới:")

    sample_manual_df = pd.DataFrame([
        {
            "Stt": 1,
            "Thiết bị": "",
            "Mã hàng": "",
            "Hãng/Xuất xứ": "",
            "Mô tả": "",
            "ĐVT": "",
            "Số lượng": "",
            "Thời gian bảo hành": "",
            "Ghi chú": "",
            "Margin Thiết bị": "",
            "ĐG COST Thiết bị": 0,
            "Margin Lắp đặt": "",
            "ĐG COST Lắp đặt": 0,
            "NCC": "",
            "NOTE": ""
        }
    ])

    column_configs = {
        "Stt": st.column_config.NumberColumn("Stt", width="small", min_value=1, step=1),
        "Thiết bị": st.column_config.TextColumn("Thiết bị", width="large", required=True),
        "Mã hàng": st.column_config.TextColumn("Mã hàng", width="medium"),
        "Hãng/Xuất xứ": st.column_config.TextColumn("Hãng/Xuất xứ", width="medium"),
        "Mô tả": st.column_config.TextColumn("Mô tả", width="large"),
        "ĐVT": st.column_config.SelectboxColumn("ĐVT", width="small", options=["Cái", "Bộ", "Mét", "Cuộn", "Lô", "Hệ thống", "Gói"], required=True),
        "Số lượng": st.column_config.NumberColumn("Số lượng", width="small", min_value=1, step=1, default=1),
        "Thời gian bảo hành": st.column_config.TextColumn("Thời gian bảo hành", width="medium"),
        "Ghi chú": st.column_config.TextColumn("Ghi chú", width="medium"),
        "Margin Thiết bị": st.column_config.TextColumn("Margin Thiết bị", width="small"),
        "ĐG COST Thiết bị": st.column_config.NumberColumn("ĐG COST Thiết bị", width="medium", format="%d ₫"),
        "Margin Lắp đặt": st.column_config.TextColumn("Margin Lắp đặt", width="small"),
        "ĐG COST Lắp đặt": st.column_config.NumberColumn("ĐG COST Lắp đặt", width="medium", format="%d ₫"),
        "NCC": st.column_config.TextColumn("NCC", width="medium"),
        "NOTE": st.column_config.TextColumn("NOTE", width="medium"),
    }

    edited_manual_df = st.data_editor(
        sample_manual_df,
        column_config=column_configs,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        key="manual_entry_editor"
    )

    if st.button("Upload Dữ Liệu", type="primary"):
        df_valid = edited_manual_df[edited_manual_df["Thiết bị"].astype(str).str.strip() != ""].copy()

        if df_valid.empty:
            st.error("Vui lòng điền thông tin tên 'Thiết bị' ít nhất 1 dòng trước khi Upload!")
        else:
            try:
                excel_bytes = process_dataframe_and_generate_excel(df_valid)

                st.subheader("Xem trước dữ liệu vừa nhập:")
                st.dataframe(df_valid, use_container_width=True, hide_index=True)

                st.download_button(
                    label="Xuất FILE Bảng Giá Chi Tiết",
                    data=excel_bytes,
                    file_name="BG - DVC - R1.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                )
            except Exception as e:
                st.error(f"Có lỗi khi xử lý dữ liệu nhập tay: {e}")

st.divider()

# --- 8. GIAO DIỆN UPLOAD FILE EXCEL CÓ SẴN ---
st.write("HOẶC UPLOAD FILE EXCEL CÓ SẴN TẠI ĐÂY (.xlsx, .xls):")

uploaded_file = st.file_uploader(
    "Upload", type=["xlsx", "xls"], label_visibility="collapsed"
)

if uploaded_file is not None:
    try:
        file_name = uploaded_file.name.lower()
        if file_name.endswith(".xlsx"):
            input_df = pd.read_excel(uploaded_file, engine="openpyxl")
        elif file_name.endswith(".xls"):
            try:
                input_df = pd.read_excel(uploaded_file, engine="xlrd")
            except Exception:
                uploaded_file.seek(0)
                input_df = pd.read_excel(uploaded_file, engine="openpyxl")
        else:
            input_df = pd.read_excel(uploaded_file)
            
        st.subheader("Xem trước dữ liệu vừa tải lên:")
        st.dataframe(input_df, use_container_width=True, hide_index=True)

        excel_bytes = process_dataframe_and_generate_excel(input_df)

        name_without_ext, ext = os.path.splitext(uploaded_file.name)
        output_filename = f"{name_without_ext} - R1{ext}"

        st.download_button(
            label="Xuất FILE Bảng Giá Chi Tiết",
            data=excel_bytes,
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

    except Exception as e:
        st.error(f"Có lỗi khi xử lý file: {e}")
