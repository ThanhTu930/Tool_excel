import io
import os
import re
import openpyxl
import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# --- 1. CẤU HÌNH STREAMLIT ---
st.set_page_config(page_title="Tool nhập liệu DVCTECH", layout="wide")

# --- 2. CSS TÙY CHỈNH GIAO DIỆN ---
st.markdown(
    """
    <style>
    div.stDownloadButton > button {
        background-color: Green !important;
        color: #FFFFFF !important;
        font-weight: bold !important;
        border-radius: 6px !important;
        border: none !important;
        padding: 0.5rem 1.2rem !important;
    }
    div.stDownloadButton > button:hover {
        background-color: DarkGreen !important;
        color: #FFFFFF !important;
    }
    div.stDownloadButton > button p {
        font-size: 18px !important;
        font-weight: bold !important;
    }
    div.stDownloadButton > button[kind="primary"] {
        background-color: Red !important;
        color: #FFFFFF !important;
        font-weight: bold !important;
        border-radius: 6px !important;
        border: none !important;
        padding: 0.5rem 1.2rem !important;
    }
    div.stDownloadButton > button[kind="primary"]:hover {
        background-color: DarkRed !important;
        color: #FFFFFF !important;
    }
    div.stDownloadButton > button[kind="primary"] p {
        font-size: 18px !important;
        font-weight: bold !important;
    }
    div[data-testid="stFileUploader"] button {
        background-color: Blue !important;
        color: #FFFFFF !important;
        font-weight: bold !important;
        border-radius: 6px !important;
        border: none !important;
    }
    div[data-testid="stFileUploader"] button:hover {
        background-color: darkblue !important;
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
        "Hãng/Xuất xứ": [],
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


# --- 5. GIAO DIỆN TẢI FORM MẪU ---
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

# --- 6. GIAO DIỆN UPLOAD FILE VÀ XỬ LÝ ---
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

        # --- ĐỌC VÀ CHUẨN HÓA DỮ LIỆU ---
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
        df_final["Ghi chú"] = get_col_val(input_df, ["ghi chú"], "")

        # Margin & Cost Thiết bị
        raw_margin = get_col_val(
            input_df, ["margin thiết bị", "margin tb"], 0
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

        # Margin & Cost Lắp đặt
        raw_margin_ld = get_col_val(
            input_df, ["margin lắp đặt", "margin ld"], 0
        )
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

        # Cấu trúc 19 cột cho Sheet CHI TIẾT
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
            "Ghi chú",
            "Margin Thiết bị",    # Col 12 (L)
            "ĐG COST Thiết bị",   # Col 13 (M)
            "TT COST Thiết bị",   # Col 14 (N)
            "Margin Lắp đặt",    # Col 15 (O)
            "ĐG COST Lắp đặt",   # Col 16 (P)
            "TT COST Lắp đặt",   # Col 17 (Q)
            "NCC",              # Col 18 (R)
            "NOTE",             # Col 19 (S)
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

            # Tự động tính độ rộng theo chuỗi dài nhất trong cột B
            max_len = max([len(str(cell.value or '')) for cell in ws_ct['B']])
            ws_ct.column_dimensions['B'].width = max(max_len + 3, 25) # Đảm bảo độ rộng tối thiểu là 25

            num_format_vnd = "#,##0"
            num_items = len(df_final)
            last_item_row = 5 + num_items if num_items > 0 else 6

            row_II = last_item_row + 1
            row_III = last_item_row + 2
            tot_row_ct = row_III + 1

            # Tọa độ bảng Phân tích Chi phí bên phải
            note_col = 19
            start_col = note_col + 3  # Cột V
            r0 = tot_row_ct + 3        # Dòng TỔNG TRƯỚC THUẾ
            c_val = get_column_letter(start_col + 1)  # Cột W (Giá trị chi phí)

            # --- 1. MỤC I: THIẾT BỊ CHÍNH (DÒNG 5) ---
            ws_ct.cell(row=5, column=1, value="I").alignment = Alignment(
                horizontal="center", vertical="center"
            )
            cell_i_tb = ws_ct.cell(row=5, column=2, value="Thiết bị chính")
            cell_i_tb.font = Font(name="Times New Roman", size=11, bold=True)

            cell_i_tt = ws_ct.cell(
                row=5, column=9, value=f"=SUM(I6:I{last_item_row})"
            )
            cell_i_tt.font = Font(name="Times New Roman", size=11, bold=True)
            cell_i_tt.number_format = num_format_vnd
            cell_i_tt.alignment = Alignment(horizontal="right", vertical="center")

            # --- 2. DÒNG DỮ LIỆU THIẾT BỊ CHÍNH (DÒNG 6 ĐẾN LAST_ITEM_ROW) ---
            for i in range(num_items):
                r = 6 + i
                ws_ct.cell(row=r, column=1).alignment = Alignment(
                    horizontal="center", vertical="center"
                )
                ws_ct.cell(row=r, column=6).alignment = Alignment(
                    horizontal="center", vertical="center"
                )
                ws_ct.cell(row=r, column=7).alignment = Alignment(
                    horizontal="center", vertical="center"
                )

                # Cost & Margin Thiết bị (L, M, N)
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

                # Cost & Margin Lắp đặt (O, P, Q)
                ws_ct.cell(row=r, column=15).number_format = "0%"
                ws_ct.cell(row=r, column=16).number_format = num_format_vnd

                cell_tt_cost_ld = ws_ct.cell(row=r, column=17)
                cell_tt_cost_ld.value = f"=G{r}*P{r}"
                cell_tt_cost_ld.number_format = num_format_vnd

            # --- 3. MỤC II: VẬT TƯ THI CÔNG ---
            ws_ct.cell(row=row_II, column=1, value="II").alignment = Alignment(
                horizontal="center", vertical="center"
            )

            cell_ii_tb = ws_ct.cell(
                row=row_II,
                column=2,
                value=(
                    "Vật tư thi công\n(Bao gồm các vật tư phụ, dây cáp, gen,"
                    " phụ kiện...)"
                ),
            )
            cell_ii_tb.font = Font(name="Times New Roman", size=11, bold=True)
            cell_ii_tb.alignment = Alignment(
                horizontal="left", vertical="center", wrap_text=True
            )

            cell_ii_dvt = ws_ct.cell(row=row_II, column=6, value="Gói")
            cell_ii_dvt.alignment = Alignment(horizontal="center", vertical="center")
            cell_ii_sl = ws_ct.cell(row=row_II, column=7, value=1)
            cell_ii_sl.alignment = Alignment(horizontal="center", vertical="center")

            # Margin Thiết bị = 20%
            ws_ct.cell(row=row_II, column=12, value=0.2).number_format = "0%"
            ws_ct.cell(row=row_II, column=13).number_format = num_format_vnd

            # Đơn giá bán = ROUNDUP(Cost/(1-Margin), -3)
            ws_ct.cell(
                row=row_II, column=8, value=f"=ROUNDUP(M{row_II}/(1-L{row_II}), -3)"
            ).number_format = num_format_vnd

            # Thành tiền bán = SL * ĐG
            ws_ct.cell(
                row=row_II, column=9, value=f"=G{row_II}*H{row_II}"
            ).number_format = num_format_vnd

            # Thành tiền COST = SL * ĐG COST
            ws_ct.cell(
                row=row_II, column=14, value=f"=G{row_II}*M{row_II}"
            ).number_format = num_format_vnd

            # --- 4. MỤC III: NHÂN CÔNG LẮP ĐẶT ---
            ws_ct.cell(row=row_III, column=1, value="III").alignment = Alignment(
                horizontal="center", vertical="center"
            )

            cell_iii_tb = ws_ct.cell(
                row=row_III,
                column=2,
                value="Nhân công lắp đặt, cấu hình, bàn giao, hướng dẫn sử dụng",
            )
            cell_iii_tb.font = Font(name="Times New Roman", size=11, bold=True)
            cell_iii_tb.alignment = Alignment(
                horizontal="left", vertical="center", wrap_text=True
            )

            cell_iii_dvt = ws_ct.cell(row=row_III, column=6, value="Gói")
            cell_iii_dvt.alignment = Alignment(horizontal="center", vertical="center")
            cell_iii_sl = ws_ct.cell(row=row_III, column=7, value=1)
            cell_iii_sl.alignment = Alignment(horizontal="center", vertical="center")

            # Margin Lắp đặt = 20%
            ws_ct.cell(row=row_III, column=15, value=0.2).number_format = "0%"

            # ĐG COST Lắp đặt = Tổng cộng chi phí triển khai
            ws_ct.cell(
                row=row_III, column=16, value=f"={c_val}{r0}"
            ).number_format = num_format_vnd

            # Thành tiền COST Lắp đặt = SL * ĐG COST Lắp đặt
            ws_ct.cell(
                row=row_III, column=17, value=f"=G{row_III}*P{row_III}"
            ).number_format = num_format_vnd

            # Đơn giá bán = ROUNDUP(Cost_LĐ/(1-Margin_LĐ), -3)
            ws_ct.cell(
                row=row_III, column=8, value=f"=ROUNDUP(P{row_III}/(1-O{row_III}), -3)"
            ).number_format = num_format_vnd

            # Thành tiền bán = SL * ĐG
            ws_ct.cell(
                row=row_III, column=9, value=f"=G{row_III}*H{row_III}"
            ).number_format = num_format_vnd

            # --- 5. DÒNG TỔNG CỘNG ---
            ws_ct.merge_cells(
                start_row=tot_row_ct, start_column=1, end_row=tot_row_ct, end_column=8
            )
            cell_total_label = ws_ct.cell(row=tot_row_ct, column=1, value="TỔNG CỘNG")
            cell_total_label.font = Font(name="Times New Roman", size=11, bold=True)
            cell_total_label.alignment = Alignment(
                horizontal="center", vertical="center"
            )

            ws_ct.cell(
                row=tot_row_ct, column=9, value=f"=I5+I{row_II}+I{row_III}"
            ).number_format = num_format_vnd

            ws_ct.cell(
                row=tot_row_ct, column=14, value=f"=SUM(N6:N{row_II})"
            ).number_format = num_format_vnd

            ws_ct.cell(
                row=tot_row_ct, column=17, value=f"=SUM(Q6:Q{last_item_row})"
            ).number_format = num_format_vnd

            # --- 6. TÍNH TOÁN BẢNG PHÂN TÍCH MARGIN / COST (BÊN PHẢI) ---
            c_cost = get_column_letter(start_col + 3)
            c_sale = get_column_letter(start_col + 4)

            font_bold = Font(name="Times New Roman", size=11, bold=True)
            font_regular = Font(name="Times New Roman", size=11, bold=False)
            align_left = Alignment(horizontal="left", vertical="center")
            align_right = Alignment(horizontal="right", vertical="center")

            num_format_number = "#,##0"
            num_format_percent = "0%"

            # Header Bảng phân tích
            row_header = r0 - 1
            ws_ct.cell(row=row_header, column=start_col + 3, value="COST").font = font_bold
            ws_ct.cell(row=row_header, column=start_col + 3).alignment = align_right

            ws_ct.cell(row=row_header, column=start_col + 4, value="GIÁ BÁN").font = font_bold
            ws_ct.cell(row=row_header, column=start_col + 4).alignment = align_right

            ws_ct.cell(row=row_header, column=start_col + 5, value="MARGIN").font = font_bold
            ws_ct.cell(row=row_header, column=start_col + 5).alignment = align_right

            # Dòng TỔNG TRƯỚC THUẾ
            ws_ct.cell(row=r0, column=start_col, value="Chi phí triển khai").font = font_bold
            ws_ct.cell(row=r0, column=start_col).alignment = align_right

            cell_sum_cp = ws_ct.cell(
                row=r0, column=start_col + 1, value=f"=SUM({c_val}{r0+1}:{c_val}{r0+8})"
            )
            cell_sum_cp.font = font_bold
            cell_sum_cp.alignment = align_right
            cell_sum_cp.number_format = num_format_number

            cell_cost_tot = ws_ct.cell(
                row=r0,
                column=start_col + 3,
                value=f"=SUM({c_cost}{r0+1}:{c_cost}{r0+2})",
            )
            cell_cost_tot.font = font_bold
            cell_cost_tot.alignment = align_right
            cell_cost_tot.number_format = num_format_number

            cell_sale_tot = ws_ct.cell(
                row=r0,
                column=start_col + 4,
                value=f"=SUM({c_sale}{r0+1}:{c_sale}{r0+2})",
            )
            cell_sale_tot.font = font_bold
            cell_sale_tot.alignment = align_right
            cell_sale_tot.number_format = num_format_number

            cell_margin_tot = ws_ct.cell(
                row=r0,
                column=start_col + 5,
                value=f"=IF({c_sale}{r0}=0, 0, ({c_sale}{r0}-{c_cost}{r0})/{c_sale}{r0})",
            )
            cell_margin_tot.font = font_bold
            cell_margin_tot.alignment = align_right
            cell_margin_tot.number_format = num_format_percent

            ws_ct.cell(row=r0, column=start_col + 6, value="TỔNG TRƯỚC THUẾ").font = font_bold
            ws_ct.cell(row=r0, column=start_col + 6).alignment = align_left

            # Dòng Chi tiết Chi phí
            items_cp = [
                "Nhân công lắp đặt",
                "Vận chuyển",
                "Di chuyển (vé xe, xe cty, xăng ...)",
                "Thuê chỗ ở",
                "Nghiệm thu, hướng dẫn sử dụng",
                "Mua, thuê dụng cụ thi công",
                "Chi phí khác (ATLĐ, Bảo hiểm, ...)",
                "Chi phí nhân sự quản lý dự án",
            ]

            for idx, label in enumerate(items_cp, start=1):
                curr_row = r0 + idx
                cell_lbl = ws_ct.cell(row=curr_row, column=start_col, value=label)
                cell_lbl.font = font_regular
                cell_lbl.alignment = align_right

                cell_val_item = ws_ct.cell(row=curr_row, column=start_col + 1)
                cell_val_item.font = font_regular
                cell_val_item.alignment = align_right
                cell_val_item.number_format = num_format_number
            # Dòng đầu tiên (Nhân công lắp đặt) lấy bằng Tổng cộng TT COST Lắp đặt
                if idx == 1:
                    cell_val_item.value = f"=Q{tot_row_ct}"

            # Dòng Thiết bị (Dòng r0 + 1)
            row_thiet_bi = r0 + 1
            cell_cost_tb = ws_ct.cell(
                row=row_thiet_bi, column=start_col + 3, value=f"=N{tot_row_ct}"
            )
            cell_cost_tb.font = font_regular
            cell_cost_tb.alignment = align_right
            cell_cost_tb.number_format = num_format_number

            cell_sale_tb = ws_ct.cell(row=row_thiet_bi, column=start_col + 4, value="=I5")
            cell_sale_tb.font = font_regular
            cell_sale_tb.alignment = align_right
            cell_sale_tb.number_format = num_format_number

            cell_mg_tb = ws_ct.cell(
                row=row_thiet_bi,
                column=start_col + 5,
                value=f"=IF({c_sale}{row_thiet_bi}=0, 0, ({c_sale}{row_thiet_bi}-{c_cost}{row_thiet_bi})/{c_sale}{row_thiet_bi})",
            )
            cell_mg_tb.font = font_regular
            cell_mg_tb.alignment = align_right
            cell_mg_tb.number_format = num_format_percent

            ws_ct.cell(row=row_thiet_bi, column=start_col + 6, value="Thiết bị").font = font_regular

            # Dòng Chi phí triển khai (Dòng r0 + 2)
            row_cptk = r0 + 2
            cell_cost_cptk = ws_ct.cell(
                row=row_cptk, column=start_col + 3, value=f"={c_val}{r0}"
            )
            cell_cost_cptk.font = font_regular
            cell_cost_cptk.alignment = align_right
            cell_cost_cptk.number_format = num_format_number

            cell_sale_cptk = ws_ct.cell(
                row=row_cptk, column=start_col + 4, value=f"=I{row_III}"
            )
            cell_sale_cptk.font = font_regular
            cell_sale_cptk.alignment = align_right
            cell_sale_cptk.number_format = num_format_number

            cell_mg_cptk = ws_ct.cell(
                row=row_cptk,
                column=start_col + 5,
                value=f"=IF({c_sale}{row_cptk}=0, 0, ({c_sale}{row_cptk}-{c_cost}{row_cptk})/{c_sale}{row_cptk})",
            )
            cell_mg_cptk.font = font_regular
            cell_mg_cptk.alignment = align_right
            cell_mg_cptk.number_format = num_format_percent

            ws_ct.cell(
                row=row_cptk, column=start_col + 6, value="Chi phí triển khai"
            ).font = font_regular

            # Định dạng Header & Đường viền Sheet CHI TIẾT
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
                cell = ws_ct.cell(row=4, column=col)
                cell.fill, cell.border = gray_fill, thin_border
                cell.font = Font(
                    name="Times New Roman", size=12, bold=True, color="000000"
                )
                cell.alignment = Alignment(
                    horizontal="center", vertical="center", wrap_text=True
                )

            for col in range(12, 20):
                cell = ws_ct.cell(row=4, column=col)
                cell.fill, cell.border = gray_fill, thin_border
                cell.font = Font(
                    name="Times New Roman", size=12, bold=True, color="FF0000"
                )
                cell.alignment = Alignment(
                    horizontal="center", vertical="center", wrap_text=True
                )

            # --- ĐOẠN CODE MỚI DÁN VÀO ĐÂY ---
            align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
            align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
            align_right = Alignment(horizontal="right", vertical="center")
            
            cols_center = [1, 3, 4, 6, 7, 10, 12, 15]  # STT, Mã hàng, Hãng, ĐVT, SL, Bảo hành, Margin...
            cols_left = [2, 5, 11, 18, 19]             # Tên thiết bị, Mô tả, Ghi chú, NCC, NOTE
            
            for r in range(5, tot_row_ct + 1):
                is_section_or_total = (r in (5, row_II, row_III, tot_row_ct))
                for c in range(1, 20):
                    cell = ws_ct.cell(row=r, column=c)
                    
                    # Gán Font & Border
                    cell.font = Font(
                        name="Times New Roman",
                        size=10,
                        bold=is_section_or_total
                    )
                    cell.border = thin_border
                    
                    # Gán Căn lề (Alignment) chuẩn theo loại cột
                    if c in cols_center:
                        cell.alignment = align_center
                    elif c in cols_left:
                        cell.alignment = align_left
                    else:
                        cell.alignment = align_right
            
                    if r == tot_row_ct:
                        cell.fill = gray_fill
            
                cell_j = ws_ct.cell(row=r, column=11)
                cell_j.border = Border(
                    left=cell_j.border.left,
                    top=cell_j.border.top,
                    right=blue_thick_side,
                    bottom=cell_j.border.bottom
                )
            
            # =========================================================
            # B. XỬ LÝ SHEET BÁO GIÁ  <-- Dán ngay phía trên dòng này
            # =========================================================

            # =========================================================
            # B. XỬ LÝ SHEET BÁO GIÁ
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
            ws_bg["G8"] = "TPHCM, ngày tháng năm 2026"

            for cell_id in ["A6", "G6", "A7", "G7", "A8", "G8"]:
                ws_bg[cell_id].font = Font(name="Times New Roman", size=11, bold=True)

            ws_bg.merge_cells("A9:H9")
            ws_bg["A9"] = "Nội dung:"
            ws_bg["A9"].font = Font(name="Times New Roman", size=11, bold=True)

            thin_side = Side(style="thin", color="000000")
            for r in range(6, 9):
                for c in range(1, 9):
                    cell = ws_bg.cell(row=r, column=c)
                    cell.border = Border(
                        top=thin_side if r == 6 else cell.border.top,
                        bottom=thin_side if r == 8 else cell.border.bottom,
                        left=thin_side if c == 1 else cell.border.left,
                        right=thin_side if c == 8 else cell.border.right,
                    )
            for c in range(1, 9):
                ws_bg.cell(row=9, column=c).border = Border(bottom=thin_side)

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

            ws_bg.merge_cells("A16:H16")
            ws_bg["A16"] = "   - "
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

            ws_bg["A27"] = "   - 30 ngày"
            ws_bg["A27"].font = Font(name="Times New Roman", size=11, bold=False)

            ws_bg.merge_cells("A28:H28")
            ws_bg["A28"] = "Chúng tôi rất mong nhận được sự hợp tác với Quý khách hàng!"
            ws_bg["A28"].font = Font(name="Times New Roman", size=11, italic=True)

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

        # NÚT XUẤT FILE HOÀN CHỈNH
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
