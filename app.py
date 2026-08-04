import io
import math
import pandas as pd
import streamlit as st

# Cấu hình trang Web
st.set_page_config(
    page_title="Tool Nhập Liệu & Báo Giá Chi Tiết", layout="wide", page_icon="📝"
)

st.title("📝 Công Cụ Nhập Liệu & Tự Động Xuất BẢNG GIÁ CHI TIẾT")
st.markdown(
    "Nhập trực tiếp dữ liệu thiết bị vào bảng bên dưới (giống mẫu **Hình 2**) ->"
    " Bấm nút **Xuất File Báo Giá** để tải về kết quả theo **Hình 1**."
)


def roundup_thousand(val):
  """Làm tròn lên đến hàng nghìn =ROUNDUP(val, -3)"""
  if pd.isna(val) or val <= 0:
    return 0
  return math.ceil(val / 1000) * 1000


# 1. Dữ liệu mẫu ban đầu theo form Hình 2
initial_data = pd.DataFrame({
    "STT": [1, 2],
    "Thiết bị": [
        "Camera IP Dome 2MP",
        "Switch PoE 24 Port Gigabit",
    ],
    "Mã hàng": ["DS-2CD1123G0-I", "CBS220-24P-4G"],
    "Mô tả chi tiết": [
        "Camera quan sát trong nhà",
        "Switch chia mạng cấp nguồn PoE",
    ],
    "Hãng / Xuất xứ": ["Hikvision / China", "Cisco / China"],
    "ĐVT": ["Cái", "Cái"],
    "Số lượng": [4, 1],
    "Ghi chú": ["Kèm chân đế", "Tủ rack trung tâm"],
    "Margin Thiết bị": [0.20, 0.15],  # 20% và 15%
    "ĐG COST Thiết bị": [850000, 9800000],
    "ĐG COST Lắp đặt": [150000, 500000],
    "NCC": ["Phúc Bình", "FPT"],
    "NOTE": ["Hàng có sẵn", "Đặt hàng 2 tuần"],
})

st.subheader("📋 Bảng Nhập Dữ Liệu Đầu Vào (Nhập trực tiếp vào đây):")
st.info(
    "💡 **Mẹo:** Bạn có thể copy nhiều dòng/cột từ Excel rồi **Paste (Ctrl+V)**"
    " trực tiếp vào bảng dưới đây, hoặc nhấn nút **'+'** ở cuối bảng để thêm dòng"
    " mới!"
)

# 2. Hiển thị bảng tương tác cho người dùng sửa/nhập dữ liệu trực tiếp (Bảng theo Hình 2)
edited_df = st.data_editor(
    initial_data,
    num_rows="dynamic",  # Cho phép thêm/xóa dòng linh hoạt
    use_container_width=True,
    column_config={
        "STT": st.column_config.NumberColumn("STT", width="small"),
        "Margin Thiết bị": st.column_config.NumberColumn(
            "Margin Thiết bị",
            help="Nhập dạng số thập phân (Ví dụ: 0.2 là 20%) hoặc số 20",
            format="%.2f",
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

# 3. Nút bấm tính toán và xuất file theo Form Hình 1
if st.button("🚀 Xử Lý & Xuất File BẢNG GIÁ CHI TIẾT", type="primary"):
  try:
    df_final = pd.DataFrame()

    # Chép dữ liệu cơ bản từ Bảng Nhập Liệu
    df_final["STT"] = edited_df["STT"]
    df_final["Thiết bị"] = edited_df["Thiết bị"]
    df_final["Mã hàng"] = edited_df["Mã hàng"]
    df_final["Hình ảnh"] = ""  # Để trống cột hình ảnh
    df_final["Hãng / Xuất xứ"] = edited_df["Hãng / Xuất xứ"]
    df_final["ĐVT"] = edited_df["ĐVT"]

    df_final["Số lượng"] = pd.to_numeric(
        edited_df["Số lượng"], errors="coerce"
    ).fillna(0)
    df_final["Ghi chú"] = edited_df["Ghi chú"]

    df_final["Margin Thiết bị"] = pd.to_numeric(
        edited_df["Margin Thiết bị"], errors="coerce"
    ).fillna(0)
    df_final["ĐG COST Thiết bị"] = pd.to_numeric(
        edited_df["ĐG COST Thiết bị"], errors="coerce"
    ).fillna(0)
    df_final["ĐG COST Lắp đặt"] = pd.to_numeric(
        edited_df["ĐG COST Lắp đặt"], errors="coerce"
    ).fillna(0)

    df_final["NCC"] = edited_df["NCC"]
    df_final["NOTE"] = edited_df["NOTE"]

    # Công thức tính Đơn giá = ROUNDUP( ĐG COST / (1 - Margin) ; -3 )
    def calc_don_gia(row):
      cost = row["ĐG COST Thiết bị"]
      margin = row["Margin Thiết bị"]
      if margin >= 1:
        margin = margin / 100
      if (1 - margin) <= 0:
        return 0
      return roundup_thousand(cost / (1 - margin))

    df_final["Đơn giá (VNĐ)"] = df_final.apply(calc_don_gia, axis=1)

    # Các công thức nhân Thành tiền & TT COST
    df_final["Thành tiền (VNĐ)"] = (
        df_final["Số lượng"] * df_final["Đơn giá (VNĐ)"]
    )
    df_final["TT COST Thiết bị"] = (
        df_final["Số lượng"] * df_final["ĐG COST Thiết bị"]
    )
    df_final["TT COST Lắp đặt"] = (
        df_final["Số lượng"] * df_final["ĐG COST Lắp đặt"]
    )

    # Sắp xếp đúng 100% thứ tự các cột của Form Hình 1
    form_1_columns = [
        "STT",
        "Thiết bị",
        "Mã hàng",
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
    df_final = df_final.reindex(columns=form_1_columns)

    # Hiển thị kết quả tính toán
    st.success("✅ Đã tính toán xong! Xem trước BẢNG GIÁ CHI TIẾT chuẩn:")
    st.dataframe(df_final)

    # Tạo file Excel để người dùng tải về
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
      df_final.to_excel(writer, index=False, sheet_name="BANG_GIA_CHI_TIET")

    st.download_button(
        label="📥 Tải File BẢNG GIÁ CHI TIẾT (.xlsx)",
        data=output.getvalue(),
        file_name="Bang_Gia_Chi_Tiet_Final.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
        type="primary",
    )

  except Exception as e:
    st.error(f"⚠️ Có lỗi xảy ra trong quá trình xử lý: {e}")
