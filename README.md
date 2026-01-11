# TOOL CHẤM CÔNG TÁC

## File thực thi
`dist\ChamCongCongTac.exe`

---

## Tính năng

- Tự động xử lý dữ liệu công tác từ file **Lệnh điều động**
- Chấm lịch công tác cho toàn bộ công nhân tài xế trong 1 tháng
- Tự động lấy địa điểm và fill theo lịch công tác
- Xuất dữ liệu ra sheet tháng/năm
- Tự động điền sheet **"Chủ Nhật"** cho người làm ngày Chủ Nhật

---

## Cách sử dụng

### 1. Chuẩn bị file

#### File "Lệnh điều động" (ví dụ: `LenhDieuDongT11-2025.xlsx`)
- Có sheet **"Nhap Lieu"** chứa dữ liệu
- **Cột R**: Tài xế
- **Cột S**: Công nhân
- **Cột U**: Địa điểm công tác
- **Cột W**: Thời gian công tác

#### File template (ví dụ: `Cham cong 11 2025.xlsx` hoặc `template.xlsx`)
- Có sheet tháng/năm (ví dụ: **"11 2025"**)
- Có sheet **"chủ nhật"**

### 2. Chạy file exe
- Double-click vào `ChamCongCongTac.exe`
- Giao diện sẽ hiện ra

### 3. Các bước thực hiện
1. Click **"Thêm file excel"** → chọn file Lệnh điều động
2. Tháng/năm sẽ tự động được nhận diện
3. Click **"Xuất file excel"**

### 4. Kết quả
- **File kết quả**: `Cham cong MM-YYYY_ket_qua.xlsx`
- Nằm cùng thư mục với file Lệnh điều động
- **Sheet tháng/năm**: có dữ liệu công tác
- **Sheet "chủ nhật"**: có danh sách người làm ngày CN

---

## Yêu cầu hệ thống

- Windows 10 SP1 trở lên (Win 10, 11)
- **KHÔNG CẦN** cài đặt Python
- **KHÔNG CẦN** cài đặt thư viện

---

## Cấu trúc file input

#### Format thời gian

- **Ngày đơn**: `01/11/2025`
- **Liên tục**: `03-06/11/2025` → ngày 3, 4, 5, 6
- **Rời rạc**: `10-11-13/12/2025` → ngày 10, 11, 13

### File template

#### Sheet tháng/năm (ví dụ: "11 2025")
- Header row chứa **"Họ và tên"**
- Cột ngày bắt đầu từ **cột D**
- Có cột **"Ghi chú công tác"**

#### Sheet "chủ nhật"
- **Row 5**: Header (Họ và tên, Chức vụ, số ngày, các ngày CN, ghi chú)
- **Row 6 trở đi**: Dữ liệu

---

## Xử lý sự cố

### Lỗi: "File mẫu không tìm thấy"
> Đảm bảo có file `Cham cong MM YYYY.xlsx` hoặc `template.xlsx` trong cùng thư mục với file Lệnh điều động

### Lỗi: "Không có dữ liệu"
- Kiểm tra sheet **"Nhap Lieu"** có đúng format không
- Kiểm tra tháng/năm đã chọn đúng chưa

### Windows Defender chặn
- Nhấn **"More info"** → **"Run anyway"**
- Hoặc thêm exception trong Windows Defender
