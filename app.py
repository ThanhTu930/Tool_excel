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


# --- 5. HÀM Chuẩn hóa 1 Dataframe đơn lẻ ---
def standardize_df(input_df):
    def get_col_val(df, possible_names, default=""):
        for name in possible_names:
            for c in df.columns:
                c_clean = c.lower().replace("/", "").replace(" ", "")
                target_clean = name.lower().replace("/", "").replace(" ", "")
                if target_clean in c_clean:
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
    df_final["Hãng/\nXuất xứ"] = get_col_val(
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
    return df_final.reindex(columns=form_columns)


# --- 6. HÀM XỬ LÝ DỮ LIỆU VÀ TẠO FILE EXCEL HOÀN CHỈNH ---
def process_dataframe_and_generate_excel(raw_input_df):
    cols = [str(c).replace("\n", " ").strip() for c in raw_input_df.columns]
    raw_input_df.columns = cols

    # 1. Tìm vị trí dòng trống đầu tiên làm ranh giới tách Mục I và Mục II
    split_idx = None
    for idx, row in raw_input_df.iterrows():
        device_val = str(row.get("Thiết bị", "")).strip()
        if device_val == "" or pd.isna(row.get("Thiết bị")):
            remaining = raw_input_df.iloc[idx + 1 :]
            if not remaining.empty and remaining["Thiết bị"].dropna().astype(str).str.strip().ne("").any():
                split_idx = idx
                break

    # 2. Tách và lọc sạch các dòng trống/rác
    if split_idx is not None:
        raw_sec1 = raw_input_df.iloc[:split_idx].copy()
        raw_sec2 = raw_input_df.iloc[split_idx + 1 :].copy()
    else:
        raw_sec1 = raw_input_df.copy()
        raw_sec2 = pd.DataFrame()

    if "Thiết bị" in raw_sec1.columns:
        raw_sec1 = raw_sec1[raw_sec1["Thiết bị"].dropna().astype(str).str.strip().ne("")].reset_index(drop=True)
    if "Thiết bị" in raw_sec2.columns and not raw_sec2.empty:
        raw_sec2 = raw_sec2[raw_sec2["Thiết bị"].dropna().astype(str).str.strip().ne("")].reset_index(drop=True)

    df_sec1 = standardize_df(raw_sec1)
    df_sec2 = standardize_df(raw_sec2)

    # Đánh lại STT liên tục cho từng phần
    if not df_sec1.empty:
        df_sec1["STT"] = range(1, len(df_sec1) + 1)
    if not df_sec2.empty:
        df_sec2["STT"] = range(1, len(df_sec2) + 1)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        ws_bg = writer.book.create_sheet(title="BÁO GIÁ", index=0)
        ws_ct = writer.book.create_sheet(title="CHI TIẾT", index=1)

        form_columns = [
            "STT", "Thiết bị", "Mã hàng", "Hãng/\nXuất xứ", "Mô tả", "ĐVT",
            "Số lượng", "Đơn giá (VNĐ)", "Thành tiền (VNĐ)", "Thời gian bảo hành",
            "Ghi chú", "Margin Thiết bị", "ĐG COST Thiết bị", "TT COST Thiết bị",
            "Margin Lắp đặt", "ĐG COST Lắp đặt", "TT COST Lắp đặt", "NCC", "NOTE"
        ]

        for col_idx, col_name in enumerate(form_columns, 1):
            ws_ct.cell(row=4, column=col_idx, value=col_name)

        ws_ct.merge_cells("B2:J2")
        title_ct = ws_ct["B2"]
        title_ct.value = "BẢNG GIÁ CHI TIẾT"
        title_ct.font = Font(name="Times New Roman", size=18, bold=True)
        title_ct.alignment = Alignment(horizontal="center", vertical="center")

        num_format_vnd = "#,##0"

        # --- GHI MỤC I ---
        ws_ct.cell(row=5, column=1, value="I").alignment = Alignment(horizontal="center", vertical="center")
        ws_ct.cell(row=5, column=2, value="Hàng hóa/Thiết bị chính").font = Font(name="Times New Roman", size=10, bold=True)

        n_sec1 = len(df_sec1)
        start_r_sec1 = 6
        end_r_sec1 = start_r_sec1 + n_sec1 - 1 if n_sec1 > 0 else start_r_sec1

        if n_sec1 > 0:
            for i, (_, row_data) in enumerate(df_sec1.iterrows()):
                r = start_r_sec1 + i
                for c_idx, val in enumerate(row_data, 1):
                    ws_ct.cell(row=r, column=c_idx, value=val)
            ws_ct.cell(row=5, column=9, value=f"=SUM(I{start_r_sec1}:I{end_r_sec1})").number_format = num_format_vnd
        else:
            ws_ct.cell(row=5, column=9, value=0).number_format = num_format_vnd

        ws_ct.cell(row=5, column=9).font = Font(name="Times New Roman", size=10, bold=True)

        # Công thức cho các dòng thuộc Mục I
        for r in range(start_r_sec1, start_r_sec1 + n_sec1):
            ws_ct.cell(row=r, column=8, value=f"=ROUNDUP(M{r}/(1-L{r}), -3)").number_format = num_format_vnd
            ws_ct.cell(row=r, column=9, value=f"=G{r}*H{r}").number_format = num_format_vnd
            ws_ct.cell(row=r, column=14, value=f"=G{r}*M{r}").number_format = num_format_vnd
            ws_ct.cell(row=r, column=17, value=f"=G{r}*P{r}").number_format = num_format_vnd

        # --- GHI MỤC II ---
        row_II = end_r_sec1 + 1 if n_sec1 > 0 else 6
        ws_ct.cell(row=row_II, column=1, value="II").alignment = Alignment(horizontal="center", vertical="center")
        ws_ct.cell(row=row_II, column=2, value="Chi phí triển khai").font = Font(name="Times New Roman", size=10, bold=True)

        n_sec2 = len(df_sec2)
        start_r_sec2 = row_II + 1
        end_r_sec2 = start_r_sec2 + n_sec2 - 1 if n_sec2 > 0 else start_r_sec2

        if n_sec2 > 0:
            for i, (_, row_data) in enumerate(df_sec2.iterrows()):
                r = start_r_sec2 + i
                for c_idx, val in enumerate(row_data, 1):
                    ws_ct.cell(row=r, column=c_idx, value=val)
            ws_ct.cell(row=row_II, column=9, value=f"=SUM(I{start_r_sec2}:I{end_r_sec2})").number_format = num_format_vnd

            # Công thức cho các dòng phụ thuộc Mục II
            for r in range(start_r_sec2, start_r_sec2 + n_sec2):
                ws_ct.cell(row=r, column=8, value=f"=ROUNDUP(M{r}/(1-L{r}), -3)").number_format = num_format_vnd
                ws_ct.cell(row=r, column=9, value=f"=G{r}*H{r}").number_format = num_format_vnd
                ws_ct.cell(row=r, column=14, value=f"=G{r}*M{r}").number_format = num_format_vnd
                ws_ct.cell(row=r, column=17, value=f"=G{r}*P{r}").number_format = num_format_vnd
        else:
            ws_ct.cell(row=row_II, column=6, value="Gói").alignment = Alignment(horizontal="center", vertical="center")
            ws_ct.cell(row=row_II, column=7, value=1).alignment = Alignment(horizontal="center", vertical="center")
            ws_ct.cell(row=row_II, column=8, value=0).number_format = num_format_vnd
            ws_ct.cell(row=row_II, column=9, value=f"=G{row_II}*H{row_II}").number_format = num_format_vnd

        ws_ct.cell(row=row_II, column=9).font = Font(name="Times New Roman", size=10, bold=True)

        # --- DÒNG TỔNG CỘNG ---
        tot_row_ct = end_r_sec2 + 1 if n_sec2 > 0 else row_II + 1
        ws_ct.merge_cells(start_row=tot_row_ct, start_column=1, end_row=tot_row_ct, end_column=8)
        cell_tot = ws_ct.cell(row=tot_row_ct, column=1, value="TỔNG CỘNG")
        cell_tot.font = Font(name="Times New Roman", size=10, bold=True)
        cell_tot.alignment = Alignment(horizontal="center", vertical="center")

        ws_ct.cell(row=tot_row_ct, column=9, value=f"=I5+I{row_II}").number_format = num_format_vnd
        ws_ct.cell(row=tot_row_ct, column=9).font = Font(name="Times New Roman", size=10, bold=True)

        # Tăng độ rộng cột I sheet CHI TIẾT để tránh lỗi #####
       
        # Áp dụng độ rộng cột cố định cho sheet CHI TIẾT
        col_widths_ct = {
            "A": 5, "B": 28, "C": 10, "D": 10, "E": 18,
            "F": 5, "G": 9, "H": 10, "I": 12, "J": 12, "K": 12,
            "L": 12, "M": 15, "N": 15, "O": 12, "P": 15, "Q": 15, "R": 15, "S": 15
        }
        for col_letter, width in col_widths_ct.items():
            ws_ct.column_dimensions[col_letter].width = width
        # Định dạng kẻ bảng cho Sheet CHI TIẾT
        gray_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
        thin_border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))

        for col in range(1, 12):
            cell = ws_ct.cell(row=4, column=col)
            cell.fill, cell.border = gray_fill, thin_border
            cell.font = Font(name="Times New Roman", size=10, bold=True, color="000000")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        for col in range(12, 20):
            cell = ws_ct.cell(row=4, column=col)
            cell.fill, cell.border = gray_fill, thin_border
            cell.font = Font(name="Times New Roman", size=10, bold=True, color="FF0000")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        align_center = Alignment(horizontal="center", vertical="center")
        align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
        align_right = Alignment(horizontal="right", vertical="center")

        for r in range(5, tot_row_ct + 1):
            is_header = r in (5, row_II, tot_row_ct)
            for c in range(1, 20):
                cell = ws_ct.cell(row=r, column=c)
                cell.border = thin_border
                cell.font = Font(name="Times New Roman", size=10, bold=is_header)

                if c in [1, 3, 4, 6, 7, 10, 12, 15]:
                    cell.alignment = align_center
                elif c in [2, 5, 11, 18, 19]:
                    cell.alignment = align_left
                else:
                    cell.alignment = align_right

                if r == tot_row_ct:
                    cell.fill = gray_fill

        # =========================================================
        # B. TẠO VÀ XỬ LÝ SHEET GUI_KH
        # =========================================================
        ws_kh = writer.book.create_sheet(title="GUI_KH", index=2)
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
            c = ws_kh[cell_id]
            c.value = text
            c.font = Font(name="Times New Roman", size=10, bold=True)
            c.fill = gray_fill
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = thin_border

        for r in range(5, tot_row_ct):
            for col_letter in ["A", "B", "C", "D", "E", "F", "G", "H", "J", "K"]:
                ws_kh[f"{col_letter}{r}"] = f"=IF('CHI TIẾT'!{col_letter}{r}=\"\",\"\",'CHI TIẾT'!{col_letter}{r})"

            # TỰ ĐỘNG GÁN CÔNG THỨC THÀNH TIỀN THEO TỪNG CẤP DÒNG
            if r == 5:
                ws_kh[f"I{r}"] = f"=SUM(I{start_r_sec1}:I{end_r_sec1})" if n_sec1 > 0 else 0
            elif r == row_II:
                ws_kh[f"I{r}"] = f"=SUM(I{start_r_sec2}:I{end_r_sec2})" if n_sec2 > 0 else f"=G{r}*H{r}"
            else:
                ws_kh[f"I{r}"] = f"=G{r}*H{r}"

            is_bold = r in (5, row_II)
            for col_letter in ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K"]:
                c = ws_kh[f"{col_letter}{r}"]
                c.font = Font(name="Times New Roman", size=10, bold=is_bold)
                c.border = thin_border
                if col_letter in ["A", "C", "D", "F", "G", "J"]:
                    c.alignment = align_center
                elif col_letter in ["H", "I"]:
                    c.alignment = align_right
                    c.number_format = num_format_vnd
                else:
                    c.alignment = align_left

        # Dòng TỔNG CỘNG GUI_KH
        ws_kh.merge_cells(f"A{tot_row_ct}:H{tot_row_ct}")
        ws_kh[f"A{tot_row_ct}"] = "TỔNG CỘNG"
        ws_kh[f"A{tot_row_ct}"].font = Font(name="Times New Roman", size=10, bold=True)
        ws_kh[f"A{tot_row_ct}"].alignment = align_center

        ws_kh[f"I{tot_row_ct}"] = f"=I5+I{row_II}"
        ws_kh[f"I{tot_row_ct}"].font = Font(name="Times New Roman", size=10, bold=True)
        ws_kh[f"I{tot_row_ct}"].alignment = align_right
        ws_kh[f"I{tot_row_ct}"].number_format = num_format_vnd

        for col_idx in range(1, 12):
            col_letter = get_column_letter(col_idx)
            cell = ws_kh[f"{col_letter}{tot_row_ct}"]
            cell.border = thin_border
            cell.fill = gray_fill

        col_widths_kh = {
            "A": 5, "B": 28, "C": 10, "D": 10, "E": 18, 
            "F": 5, "G": 9, "H": 10, "I": 12, "J": 12, "K": 12
        }
        for col_letter, width in col_widths_kh.items():
            ws_kh.column_dimensions[col_letter].width = width

        # =========================================================
        # C. TẠO VÀ XỬ LÝ SHEET BÁO GIÁ
        # =========================================================
        ws_bg.page_setup.orientation = ws_bg.ORIENTATION_PORTRAIT
        ws_bg.page_setup.paperSize = ws_bg.PAPERSIZE_A4
        ws_bg.sheet_properties.pageSetUpPr.fitToPage = True
        ws_bg.page_setup.fitToWidth = 1
        ws_bg.page_setup.fitToHeight = 0

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

        col_widths_bg = {"A": 6, "B": 32, "C": 10, "D": 10, "E": 20, "F": 14, "G": 20}
        for col_letter, width in col_widths_bg.items():
            ws_bg.column_dimensions[col_letter].width = width

        sheet_order = ["BÁO GIÁ", "GUI_KH", "CHI TIẾT"]
        writer.book._sheets = [writer.book[s] for s in sheet_order if s in writer.book.sheetnames]

    return output.getvalue()


# --- 7. GIAO DIỆN TẢI FORM MẪU & NHẬP TRỰC TIẾP ---
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

# --- 8. KHUNG NHẬP DỮ LIỆU BÁO GIÁ TRỰC TIẾP ---
if st.session_state.show_manual_input:
    st.markdown("---")
    st.subheader("Bảng nhập thông tin báo giá trực tiếp")
    st.caption("Gõ thông tin trực tiếp vào bảng dưới đây. Để phân tách giữa **Mục I (Thiết bị)** và **Mục II (Chi phí triển khai)**, hãy để trống 1 dòng ở giữa!")

    sample_manual_df = pd.DataFrame([
        {
            "Stt": 1,
            "Thiết bị": "",
            "Mã hàng": "",
            "Hãng/Xuất xứ": "",
            "Mô tả": "",
            "ĐVT": "Cái",
            "Số lượng": 1,
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
        "Thiết bị": st.column_config.TextColumn("Thiết bị", width="large"),
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
        if edited_manual_df.empty:
            st.error("Vui lòng nhập dữ liệu trước khi Upload!")
        else:
            try:
                excel_bytes = process_dataframe_and_generate_excel(edited_manual_df)

                st.subheader("Xem trước dữ liệu vừa nhập:")
                st.dataframe(edited_manual_df, use_container_width=True, hide_index=True)

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

# --- 9. GIAO DIỆN UPLOAD FILE EXCEL CÓ SẴN ---
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
