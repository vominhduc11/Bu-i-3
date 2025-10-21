import pandas as pd
import numpy as np

def tinh_toan_kho(file_path="data_nhap_xuat_kho.xlsx", output_file="ket_qua_quan_ly_kho.xlsx"):
    """
    Hàm đọc dữ liệu kho, tính toán lợi nhuận và phân loại một cách hiệu quả
    sử dụng các hàm của pandas.
    """
    try:
        # I. Đọc và Xử lý Dữ liệu Đầu vào
        df = pd.read_excel(file_path)

        # Đảm bảo dữ liệu được sắp xếp theo ngày để xử lý logic chính xác
        df = df.sort_values(by='Ngày nhập/xuất').reset_index(drop=True)

        # --- II. Tính toán Lợi nhuận (Cách làm không dùng vòng lặp) ---

        # 1. Tạo một bảng chỉ chứa các giao dịch 'Nhập' và giá nhập gần nhất
        df_nhap = df[df['Loại giao dịch'] == 'Nhập'].copy()
        # Giữ lại lần nhập cuối cùng (giá gần nhất) cho mỗi sản phẩm
        gia_nhap_gan_nhat = df_nhap.drop_duplicates(subset='Mã sản phẩm', keep='last')
        
        # Tạo một cột 'Giá nhập' mới bằng cách ánh xạ giá từ bảng trên
        df['Giá nhập'] = df['Mã sản phẩm'].map(gia_nhap_gan_nhat.set_index('Mã sản phẩm')['Đơn giá'])
        
        # Điền các giá trị NaN bằng giá trị liền trước để đảm bảo các giao dịch 'Xuất' có giá nhập
        df['Giá nhập'] = df['Giá nhập'].fillna(method='ffill')

        # 2. Tính lợi nhuận chỉ cho các giao dịch 'Xuất'
        # Mặc định lợi nhuận là 0
        df['Lợi nhuận'] = 0.0
        # Dùng np.where để tính toán có điều kiện một cách nhanh chóng
        df['Lợi nhuận'] = np.where(
            df['Loại giao dịch'] == 'Xuất',
            (df['Đơn giá'] - df['Giá nhập']) * df['Số lượng'],
            0
        )
        
        # --- III. Phân Loại Lợi Nhuận (Cách làm không dùng vòng lặp) ---

        # 1. Định nghĩa các điều kiện và kết quả tương ứng
        conditions = [
            (df['Loại giao dịch'] == 'Xuất') & (df['Đơn giá'] >= 200000),
            (df['Loại giao dịch'] == 'Xuất') & (df['Đơn giá'] >= 100000),
            (df['Loại giao dịch'] == 'Xuất')
        ]
        choices = ['Cao', 'Trung bình', 'Thấp']
        
        # 2. Dùng np.select để tạo cột phân loại dựa trên các điều kiện trên
        df['Phân loại Lợi nhuận'] = np.select(conditions, choices, default='')

        # --- IV. Cập nhật và Lưu Trữ ---
        
        # Xóa cột 'Giá nhập' trung gian nếu không muốn hiển thị trong file kết quả
        df = df.drop(columns=['Giá nhập'])

        df.to_excel(output_file, index=False)
        
        print(f"✅ Hoàn tất xử lý! Kết quả đã được lưu vào file: '{output_file}'")
        print("\nXem trước 5 dòng kết quả:")
        print(df.head())

    except FileNotFoundError:
        print(f"❌ Lỗi: Không tìm thấy file '{file_path}'. Vui lòng kiểm tra lại.")
    except Exception as e:
        print(f"❌ Đã xảy ra lỗi không mong muốn: {e}")

# --- Thực thi chương trình ---
if __name__ == "__main__":
    tinh_toan_kho()