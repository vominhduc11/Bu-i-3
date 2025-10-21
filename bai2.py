import pandas as pd
import numpy as np

# --- I. Đọc và Xử lý Dữ liệu Đầu vào ---
print(">>> Bước 1: Đang đọc và xử lý dữ liệu từ file Excel...")

try:
    df = pd.read_excel('data_logistics.xlsx')
    # SỬA LỖI: Tự động xóa khoảng trắng thừa ở đầu/cuối tên cột để tránh lỗi KeyError
    df.columns = df.columns.str.strip()
    print("Đọc file 'data_logistics.xlsx' và làm sạch tên cột thành công.")
except FileNotFoundError:
    print("Lỗi: Không tìm thấy file 'data_logistics.xlsx'. Đang tạo file mẫu...")
    # Dữ liệu mẫu không thay đổi
    data_mau = {
        'Mã đơn': ['DH001', 'DH002', 'DH003', 'DH004', 'DH005', 'DH006'],
        'Khách hàng': ['Công ty A', 'Cửa hàng B', 'Công ty A', 'Khách lẻ C', 'Đại lý D', 'Công ty E'],
        'Loại hàng': ['Điện tử', 'Thời trang', 'Gia dụng', 'Mỹ phẩm', 'Điện tử', 'Nội thất'],
        'Nơi gửi': ['Hà Nội', 'TP.HCM', 'Hà Nội', 'Đà Nẵng', 'Cần Thơ', 'Hải Phòng'],
        'Nơi nhận': ['Hải Phòng', 'Cần Thơ', 'Đà Nẵng', 'Huế', 'TP.HCM', 'TP.HCM'],
        'Khối lượng (kg)': [25, 4, 15, 8, 1, 50],
        'Loại vận chuyển': ['Tiêu chuẩn', 'Siêu tốc', 'Nhanh', 'Tiêu chuẩn', 'Siêu tốc', 'Tiêu chuẩn'],
        'Cước cơ bản(VND/kg)': [10000, 15000, 12000, 11000, 18000, 8000],
        'Phí vận chuyển (VND)': [0, 0, 0, 0, 0, 0]
    }
    df = pd.DataFrame(data_mau)
    df.to_excel('data_logistics.xlsx', index=False)
    print("Đã tạo file 'data_logistics.xlsx' với dữ liệu mẫu.")

# --- II. Tính Toán Phí Vận Chuyển (Phiên bản tối ưu) ---
print("\n>>> Bước 2: Bắt đầu tính toán phí vận chuyển (đã tối ưu)...")

# 1. Tính phí ban đầu
phi_ban_dau = df['Cước cơ bản (đ/kg)'] * df['Khối lượng (kg)']

# 2. Áp dụng giảm giá theo khối lượng
conditions_kg = [
    df['Khối lượng (kg)'] > 20,
    df['Khối lượng (kg)'] > 5
]
choices_kg_discount = [0.90, 0.95] # Mức giảm 10%, 5%
phi_sau_giam_gia_kg = phi_ban_dau * np.select(conditions_kg, choices_kg_discount, default=1.0)

# 3. Cộng thêm phụ phí vận chuyển
# np.where(điều kiện, giá trị nếu đúng, giá trị nếu sai)
phu_phi_sieu_toc = np.where(df['Loại vận chuyển'] == 'Siêu tốc', 20000, 0)
phi_sau_phu_phi = phi_sau_giam_gia_kg + phu_phi_sieu_toc

# 4. Áp dụng giảm giá theo khu vực
def phan_loai_mien(tinh_series):
    mien_bac = ['Hà Nội', 'Hải Phòng', 'Nam Định', 'Quảng Ninh', 'Hải Dương', 'Hưng Yên', 'Bắc Ninh']
    mien_trung = ['Đà Nẵng', 'Huế', 'Quảng Nam', 'Quảng Ngãi', 'Bình Định', 'Phú Yên', 'Thanh Hóa', 'Nghệ An']
    mien_nam = ['TP.HCM', 'Cần Thơ', 'Vũng Tàu', 'Bình Dương', 'Đồng Nai', 'Long An']
    
    conditions = [
        tinh_series.isin(mien_bac),
        tinh_series.isin(mien_trung),
        tinh_series.isin(mien_nam)
    ]
    choices = ['Bắc', 'Trung', 'Nam']
    return np.select(conditions, choices, default='Không xác định')

mien_gui = phan_loai_mien(df['Nơi gửi'])
mien_nhan = phan_loai_mien(df['Nơi nhận'])

# Điều kiện giảm giá: cùng miền VÀ miền đó phải xác định được
dieu_kien_giam_mien = (mien_gui == mien_nhan) & (mien_gui != 'Không xác định')
giam_gia_mien = np.where(dieu_kien_giam_mien, 0.95, 1.0)

# Phí cuối cùng
df['Phí vận chuyển (VND)'] = phi_sau_phu_phi * giam_gia_mien

print("Tính toán phí vận chuyển hoàn tất.")

# --- III. Phân Loại Phí Vận Chuyển (Phiên bản tối ưu) ---
print("\n>>> Bước 3: Phân loại cước vận chuyển...")
conditions_phan_loai = [
    df['Phí vận chuyển (VND)'] > 500000,
    df['Phí vận chuyển (VND)'] > 200000
]
choices_phan_loai = ['Cao', 'Trung bình']
df['Phân loại cước'] = np.select(conditions_phan_loai, choices_phan_loai, default='Thấp')
print("Phân loại cước hoàn tất.")

# --- IV. Cập Nhật và Lưu Trữ (Không thay đổi) ---
print("\n>>> Bước 4: Đang lưu kết quả ra file Excel mới...")
output_filename = 'BaiTap_VanChuyen_IFELSE_Solved.xlsx'
df.to_excel(output_filename, index=False)
print(f"Đã lưu thành công kết quả vào file: '{output_filename}'")
print("\n--- XEM TRƯỚC 5 DÒNG ĐẦU CỦA KẾT QUẢ ---")
print(df.head())

# --- V. Phân tích Chiến lược (Thêm cột Miền để phân tích) ---
print("\n\n--- PHÂN TÍCH CHIẾN LƯỢC (MỞ RỘNG) ---")
# Thêm cột miền vào df để groupby cho dễ
df['Miền gửi'] = mien_gui
df['Miền nhận'] = mien_nhan

# 1. Tính tổng doanh thu theo loại vận chuyển
doanh_thu_theo_loai_vc = df.groupby('Loại vận chuyển')['Phí vận chuyển (VND)'].sum().round(0)
print("\n1. Tổng doanh thu theo từng loại vận chuyển:")
print(doanh_thu_theo_loai_vc.to_string(float_format='{:,.0f} VND'.format))

# 2. Thống kê số đơn theo miền gửi/nhận
so_don_theo_mien_gui = df['Miền gửi'].value_counts()
so_don_theo_mien_nhan = df['Miền nhận'].value_counts()
print("\n2. Thống kê số lượng đơn hàng:")
print("- Theo miền gửi:")
print(so_don_theo_mien_gui)
print("\n- Theo miền nhận:")
print(so_don_theo_mien_nhan)

# 3. Lọc top 5 khách hàng có tổng cước cao nhất
top_5_khach_hang = df.groupby('Khách hàng')['Phí vận chuyển (VND)'].sum().nlargest(5)
print("\n3. Top 5 khách hàng có tổng cước phí cao nhất:")
print(top_5_khach_hang.to_string(float_format='{:,.0f} VND'.format))