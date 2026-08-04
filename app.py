import io
import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

st.set_page_config(
    page_title="Tool Nhập Liệu & Báo Giá Chi Tiết", layout="wide", page_icon="📝"
)

st.title("📝 Công Cụ Nhập Liệu & Xuất BẢNG GIÁ CHI TIẾT (Giữ Công Thức Excel)")

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

if st.button("🚀 Xử Lý & Xuất File Giữ Nguyên Công Thức Excel", type="primary"):
  try:
    # 2. Xử lý khung Dataframe cơ bản
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

    # Đặt giá trị mặc định cho cột tính toán (sẽ chèn công thức Excel vào sau)
    df_final["Đơn giá (VNĐ)"] = 0
    df_final["Thành tiền (VNĐ)"] = 0
    df_final["Ghi chú"] = edited_df["Ghi chú"]

    df_final["Margin Thiết bị"] = pd.to_numeric(
        edited_df["Margin Thiết bị"], errors="coerce"
    ).fillna(0)
    df_final["ĐG COST Thiết bị"] = pd.to_numeric(
        edited_df["ĐG COST Thiết bị"], errors="coerce"
    ).fillna(0)
    df_final["TT COST Thiết bị"] = 0
    df_final["ĐG COST Lắp đặt"] = pd.to_numeric(
        edited_df["ĐG COST Lắp đặt"], errors="coerce"
    ).fillna(0)
    df_final["TT COST Lắp đặt"] = 0

    df_final["NCC"] = edited_df["NCC"]
    df_final["NOTE"] = edited_df["NOTE"]

    # Đặt đúng vị trí cột theo thứ tự
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

    # 3. Xuất ra Excel và chèn công thức bằng OPENPYXL
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
      # Ghi dữ liệu từ dòng 4 (Tiêu đề cột ở dòng 4, dữ liệu bắt đầu từ dòng 5)
      df_final.to_excel(
          writer, index=False, sheet_name="BANG_GIA_CHI_TIET", startrow=3
      )

      workbook = writer.book
      worksheet = writer.sheets["BANG_GIA_CHI_TIET"]

      # --- A. BẬT CHẾ ĐỘ PAGE BREAK PREVIEW (An toàn) ---
      try:
        if not hasattr(worksheet, "sheet_views") or not worksheet.sheet_views:
          worksheet.views.sheetView[0].sheetViewType = "pageBreakPreview"
        else:
          worksheet.sheet_views[0].sheetViewType = "pageBreakPreview"
      except Exception:
        pass

      # --- B. MERGE CỘT TIÊU ĐỀ LỚN ---
      worksheet.merge_cells("B2:J2")
      title_cell = worksheet["B2"]
      title_cell.value = "BẢNG GIÁ CHI TIẾT"
      title_cell.font = Font(name="Times New Roman", size=18, bold=True)
      title_cell.alignment = Alignment(horizontal="center", vertical="center")

      # --- C. CHÈN CÔNG THỨC EXCEL VÀ ĐỊNH DẠNG DỮ LIỆU TỪNG DÒNG ---
      num_format_vnd = (
          "#,##0"  # Định dạng hiển thị số tiền có dấu phân cách hàng nghìn
      )

      for i in range(len(df_final)):
        r = 5 + i  # Dòng thực tế trong Excel (Bắt đầu từ 5)

        # 1. Công thức Đơn giá (Cột I - Cột thứ 9)
        # =ROUNDUP( ĐG_COST / (1 - IF(Margin>=1, Margin/100, Margin)) ; -3 )
        formula_don_gia = f"=IF(M{r}<=0, 0, ROUNDUP(M{r} / (1 - IF(L{r}>=1, L{r}/100, L{r})), -3))"
        cell_don_gia = worksheet.cell(row=r, column=9)
        cell_don_gia.value = formula_don_gia
        cell_don_gia.number_format = num_format_vnd

        # 2. Công thức Thành tiền (Cột J - Cột thứ 10) = Số lượng (H) * Đơn giá (I)
        cell_thanh_tien = worksheet.cell(row=r, column=10)
        cell_thanh_tien.value = f"=H{r}*I{r}"
        cell_thanh_tien.number_format = num_format_vnd

        # 3. Định dạng ĐG COST Thiết bị (Cột M - Cột thứ 13)
        worksheet.cell(row=r, column=13).number_format = num_format_vnd

        # 4. Công thức TT COST Thiết bị (Cột N - Cột thứ 14) = Số lượng (H) * ĐG COST Thiết bị (M)
        cell_tt_cost_tb = worksheet.cell(row=r, column=14)
        cell_tt_cost_tb.value = f"=H{r}*M{r}"
        cell_tt_cost_tb.number_format = num_format_vnd

        # 5. Định dạng ĐG COST Lắp đặt (Cột O - Cột thứ 15)
        worksheet.cell(row=r, column=15).number_format = num_format_vnd

        # 6. Công thức TT COST Lắp đặt (Cột P - Cột thứ 16) = Số lượng (H) * ĐG COST Lắp đặt (O)
        cell_tt_cost_ld = worksheet.cell(row=r, column=16)
        cell_tt_cost_ld.value = f"=H{r}*O{r}"
        cell_tt_cost_ld.number_format = num_format_vnd

        # Định dạng Margin (Cột L - Cột 12) theo %
        worksheet.cell(row=r, column=12).number_format = "0.00%"

      # --- D. TÔ MÀU & KẺ VIỀN CHO HEADER VÀ DỮ LIỆU ---
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

      # Format Header (Dòng 4)
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

      # Format Border & Font cho dữ liệu
      for row in range(5, len(df_final) + 5):
        for col in range(1, 19):
          c = worksheet.cell(row=row, column=col)
          c.font = Font(name="Times New Roman", size=10)
          c.border = thin_border

        # Đường viền phân cách xanh lục ở cột K
        cell_k = worksheet.cell(row=row, column=11)
        cell_k.border = Border(
            left=cell_k.border.left,
            top=cell_k.border.top,
            right=blue_thick_side,
            bottom=cell_k.border.bottom,
        )

    st.success("✅ Đã xuất file thành công! Tất cả ô tính toán đều giữ công thức Excel sống.")
    st.download_button(
        label="📥 Tải File BẢNG GIÁ CHI TIẾT (Có Công Thức) (.xlsx)",
        data=output.getvalue(),
        file_name="Bang_Gia_Chi_Tiet_With_Formulas.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
        type="primary",
    )

  except Exception as e:
    st.error(f"⚠️ Có lỗi xảy ra: {e}")
