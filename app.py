import io
import math
import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

st.set_page_config(
    page_title="Tool Nhập Liệu & Báo Giá Chi Tiết", layout="wide", page_icon="📝"
)

st.title("📝 Công Cụ Nhập Liệu & Xuất BẢNG GIÁ CHI TIẾT (Định Dạng Chuẩn)")


def roundup_thousand(val):
  if pd.isna(val) or val <= 0:
    return 0
  return math.ceil(val / 1000) * 1000


# 1. Dữ liệu mẫu ban đầu
initial_data = pd.DataFrame({
    "STT": [1, 2],
    "Thiết bị": ["Camera IP Dome 2MP", "Switch PoE 24 Port Gigabit"],
    "Mã hàng": ["DS-2CD1123G0-I", "CBS220-24P-4G"],
    "Mô tả chi tiết": [
        "Camera quan sát trong nhà",
        "Switch chia mạng cấp nguồn PoE",
    ],
    "Hãng / Xuất xứ": ["Hikvision / China", "Cisco / China"],
    "ĐVT": ["Cái", "Cái"],
    "Số lượng": [4, 1],
    "Ghi chú": ["Kèm chân đế", "Tủ rack trung tâm"],
    "Margin Thiết bị": [0.20, 0.15],
    "ĐG COST Thiết bị": [850000, 9800000],
    "ĐG COST Lắp đặt": [150000, 500000],
    "NCC": ["Phúc Bình", "FPT"],
    "NOTE": ["Hàng có sẵn", "Đặt hàng 2 tuần"],
})

st.subheader("📋 Bảng Nhập Dữ Liệu Đầu Vào:")
edited_df = st.data_editor(
    initial_data,
    num_rows="dynamic",
    use_container_width=True,
    column_config={
        "STT": st.column_config.NumberColumn("STT", width="small"),
        "Margin Thiết bị": st.column_config.NumberColumn(
            "Margin Thiết bị", format="%.2f"
        ),
        "ĐG COST Thiết bị": st.column_config.NumberColumn(
            "ĐG COST Thiết bị", format="%d VNĐ"
        ),
        "ĐG COST Lắp đặt": st.column_config.NumberColumn(
            "ĐG COST Lắp đặt", format="%d VNĐ"
        ),
    },
)

st.divider()

if st.button("🚀 Xử Lý & Xuất File BẢNG GIÁ CHI TIẾT Chuẩn", type="primary"):
  try:
    # 2. Xử lý tính toán Dataframe
    df_final = pd.DataFrame()
    df_final["STT"] = edited_df["STT"]
    df_final["Thiết bị"] = edited_df["Thiết bị"]
    df_final["Mã hàng"] = edited_df["Mã hàng"]
    df_final["Mô tả chi tiết"] = edited_df["Mô tả chi tiết"]
    df_final["Hình ảnh"] = ""
    df_final["Hãng / Xuất xứ"] = edited_df["Hãng / Xuất xứ"]
    df_final["ĐVT"] = edited_df["ĐVT"]
    df_final["Số lượng"] = pd.to_numeric(
        edited_df["Số lượng"], errors="coerce"
    ).fillna(0)

    # Tính Đơn giá & Thành tiền
    def calc_don_gia(row):
      cost = row["ĐG COST Thiết bị"]
      margin = row["Margin Thiết bị"]
      if margin >= 1:
        margin = margin / 100
      if (1 - margin) <= 0:
        return 0
      return roundup_thousand(cost / (1 - margin))

    df_final["ĐG COST Thiết bị"] = pd.to_numeric(
        edited_df["ĐG COST Thiết bị"], errors="coerce"
    ).fillna(0)
    df_final["Margin Thiết bị"] = pd.to_numeric(
        edited_df["Margin Thiết bị"], errors="coerce"
    ).fillna(0)
    df_final["ĐG COST Lắp đặt"] = pd.to_numeric(
        edited_df["ĐG COST Lắp đặt"], errors="coerce"
    ).fillna(0)

    df_final["Đơn giá (VNĐ)"] = df_final.apply(calc_don_gia, axis=1)
    df_final["Thành tiền (VNĐ)"] = (
        df_final["Số lượng"] * df_final["Đơn giá (VNĐ)"]
    )
    df_final["Ghi chú"] = edited_df["Ghi chú"]

    df_final["TT COST Thiết bị"] = (
        df_final["Số lượng"] * df_final["ĐG COST Thiết bị"]
    )
    df_final["TT COST Lắp đặt"] = (
        df_final["Số lượng"] * df_final["ĐG COST Lắp đặt"]
    )
    df_final["NCC"] = edited_df["NCC"]
    df_final["NOTE"] = edited_df["NOTE"]

    # Đặt thứ tự cột chuẩn theo hình
    form_columns = [
        "STT",
        "Thiết bị",
        "Mã hàng",
        "Mô tả chi tiết",
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
    df_final = df_final.reindex(columns=form_columns)

    # 3. Xuất ra Excel và áp dụng định dạng nâng cao bằng OPENPYXL
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
      # Ghi dữ liệu từ dòng thứ 4 (chừa dòng 1, 2, 3 làm Tiêu đề)
      df_final.to_excel(
          writer, index=False, sheet_name="BANG_GIA_CHI_TIET", startrow=3
      )

      workbook = writer.book
      worksheet = writer.sheets["BANG_GIA_CHI_TIET"]

      # --- A. BẬT CHẾ ĐỘ PAGE BREAK PREVIEW (Cú pháp chuẩn sheet_views) ---
      worksheet.sheet_views[0].sheetViewType = "pageBreakPreview"

      # --- B. MERGE CỘT VÀ TẠO TIÊU ĐỀ LỚN "BẢNG GIÁ CHI TIẾT" ---
      worksheet.merge_cells("B2:J2")
      title_cell = worksheet["B2"]
      title_cell.value = "BẢNG GIÁ CHI TIẾT"
      title_cell.font = Font(name="Times New Roman", size=18, bold=True)
      title_cell.alignment = Alignment(horizontal="center", vertical="center")

      # --- C. TÔ MÀU & ĐỊNH DẠNG CÁC DÒNG TIÊU ĐỀ CỘT (HEADER - Dòng 4) ---
      gray_fill = PatternFill(
          start_color="D9D9D9", end_color="D9D9D9", fill_type="solid"
      )
      thin_border = Border(
          left=Side(style="thin"),
          right=Side(style="thin"),
          top=Side(style="thin"),
          bottom=Side(style="thin"),
      )

      # Định dạng Cột A -> K (Chữ Đen, Nền Xám, Căn Giữa)
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

      # Định dạng Cột L -> R (COST/Margin: Chữ Đỏ, Nền Xám, Căn Giữa)
      for col in range(12, 19):
        cell = worksheet.cell(row=4, column=col)
        cell.fill = gray_fill
        cell.font = Font(
            name="Times New Roman", size=10, bold=True, color="FF0000"
        )  # Màu đỏ
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )
        cell.border = thin_border

      # --- D. KẺ ĐƯỜNG PHÂN CÁCH DẦY MÀU XANH GIỮA VÙNG IN VÀ VÙNG COST ---
      blue_thick_side = Side(style="medium", color="0000FF")
      for row in range(4, len(df_final) + 5):
        cell_k = worksheet.cell(row=row, column=11)
        cell_k.border = Border(
            left=cell_k.border.left,
            top=cell_k.border.top,
            right=blue_thick_side,
            bottom=cell_k.border.bottom,
        )

      # --- E. ĐỊNH DẠNG DỮ LIỆU CÁC DÒNG BÊN DƯỚI ---
      for row in range(5, len(df_final) + 5):
        for col in range(1, 19):
          c = worksheet.cell(row=row, column=col)
          c.font = Font(name="Times New Roman", size=10)
          c.border = thin_border

    st.success("✅ Đã tạo xong file Excel chuẩn giao diện!")
    st.download_button(
        label="📥 Tải File BẢNG GIÁ CHI TIẾT Đã Định Dạng (.xlsx)",
        data=output.getvalue(),
        file_name="Bang_Gia_Chi_Tiet_Formatted.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
        type="primary",
    )

  except Exception as e:
    st.error(f"⚠️ Có lỗi xảy ra: {e}")
