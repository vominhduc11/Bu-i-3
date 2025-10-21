import pandas as pd
import numpy as np

# --- Bước 0: Tạo dữ liệu mẫu (Bạn có thể bỏ qua bước này nếu đã có file Excel) ---
# Đoạn code này tạo một file Excel mẫu để chương trình có thể chạy ngay lập tức.
try:
    data = {
        'Mã sản phẩm': ['SP001', 'SP002', 'SP001', 'SP003', 'SP002', 'SP001', 'SP003'],
        'Tên sản phẩm': ['Laptop A', 'Mouse B', 'Laptop A', 'Keyboard C', 'Mouse B', 'Laptop A', 'Keyboard C'],
        'Ngày nhập/xuất': pd.to_datetime(['2023-01-10', '2023-01-11', '2023-01-15', '2023-01-20', '2023-02-01', '2023-02-05', '2023-02-10']),
        'Số lượng': [50, 200, 30, 150, 100, 20, 80],
        'Đơn giá': [180000, 85000, 210000, 120000, 95000, 220000, 135000],
        'Loại giao dịch': ['Nhập', 'Nhập', 'Xuất', 'Nhập', 'Xuất', 'Xuất', 'Xuất']
    }
    df_sample = pd.DataFrame(data)
    df_sample.to_excel("data_nhap_xuat_kho.xlsx", index=False)
    print("Đã tạo file 'data_nhap_xuat_kho.xlsx' mẫu thành công.")
except Exception as e:
    print(f"Lỗi khi tạo file mẫu: {e}")

# --- Chương trình chính ---

def quan_ly_kho(file_path):
    """
    Hàm đọc dữ liệu từ file Excel, tính toán lợi nhuận, phân loại
    và lưu kết quả vào một file Excel mới.
    """
    try:
        # I. Đọc và Xử lý Dữ liệu Đầu vào
        df = pd.read_excel(file_path)

        # Đảm bảo cột ngày tháng được sắp xếp đúng thứ tự để tìm giá nhập gần nhất
        df = df.sort_values(by='Ngày nhập/xuất').reset_index(drop=True)

        # Tạo các cột mới và khởi tạo giá trị
        df['Lợi nhuận'] = 0.0
        df['Phân loại Lợi nhuận'] = ''

        # Dùng dictionary để lưu giá nhập gần nhất của mỗi sản phẩm
        gia_nhap_gan_nhat = {}

        # II. Tính toán Lợi nhuận
        for index, row in df.iterrows():
            ma_sp = row['Mã sản phẩm']
            loai_gd = row['Loại giao dịch']
            don_gia = row['Đơn giá']

            # Nếu là giao dịch 'Nhập', cập nhật giá nhập gần nhất
            if loai_gd == 'Nhập':
                gia_nhap_gan_nhat[ma_sp] = don_gia
            
            # Nếu là giao dịch 'Xuất', tính toán lợi nhuận
            elif loai_gd == 'Xuất':
                gia_ban = don_gia
                so_luong_xuat = row['Số lượng']
                
                # Lấy giá nhập tương ứng từ dictionary
                gia_nhap = gia_nhap_gan_nhat.get(ma_sp, 0) # Mặc định là 0 nếu không tìm thấy

                if gia_nhap > 0:
                    loi_nhuan = (gia_ban - gia_nhap) * so_luong_xuat
                    df.at[index, 'Lợi nhuận'] = loi_nhuan
                    
                    # III. Phân Loại Lợi Nhuận
                    if gia_ban >= 200000:
                        df.at[index, 'Phân loại Lợi nhuận'] = 'Cao'
                    elif 100000 <= gia_ban < 200000:
                        df.at[index, 'Phân loại Lợi nhuận'] = 'Trung bình'
                    else:
                        df.at[index, 'Phân loại Lợi nhuận'] = 'Thấp'

        # IV. Cập nhật và Lưu Trữ
        output_file = "ket_qua_quan_ly_kho.xlsx"
        df.to_excel(output_file, index=False)
        
        print("\nHoàn tất xử lý!")
        print(f"Kết quả đã được lưu vào file: '{output_file}'")
        print("\nXem trước 5 dòng kết quả:")
        print(df.head())

    except FileNotFoundError:
        print(f"Lỗi: Không tìm thấy file '{file_path}'. Vui lòng kiểm tra lại đường dẫn.")
    except Exception as e:
        print(f"Đã xảy ra lỗi không mong muốn: {e}")

# --- Thực thi chương trình ---
if __name__ == "__main__":
    file_excel = "data_nhap_xuat_kho.xlsx"
    quan_ly_kho(file_excel)