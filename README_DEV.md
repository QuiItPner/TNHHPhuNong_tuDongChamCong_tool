# Tool Chấm Công Công Tác

## Tổng quan

Tool tự động xử lý dữ liệu công tác từ file "Lệnh điều động" và xuất kết quả ra file template chấm công, bao gồm:
- Sheet tháng/năm với dữ liệu công tác
- Sheet "Chủ Nhật" với danh sách người làm việc ngày Chủ Nhật

## Cấu trúc Project

```
excel-timekeeping-automation-tool/
├── excel_tool.py              # Main source code
├── app_icon.ico              # Icon cho exe
├── logo.png                  # Logo trong GUI
├── logo.ico                  # Logo icon
├── requirements.txt          # Python dependencies
├── ChamCongCongTac.spec      # PyInstaller spec file
├── build_exe.py              # Build script
├── BUILD_INSTRUCTIONS.txt    # Hướng dẫn build
├── HUONG_DAN_SU_DUNG.txt    # Hướng dẫn user
├── dist/                     # Output folder
│   └── ChamCongCongTac.exe  # File exe (30MB)
└── build/                    # PyInstaller build temp
```

## Dependencies

- **openpyxl** (>=3.0.0): Đọc/ghi Excel files
- **Pillow** (>=9.0.0): Xử lý images (logo)
- **tkinter**: GUI (built-in Python)

## Tính năng chính

### 1. Parse dữ liệu công tác
- Đọc sheet "Nhap Lieu" từ file Lệnh điều động
- Parse thời gian công tác (hỗ trợ nhiều format)
- Parse địa điểm công tác (giữ nguyên format gốc)
- Lưu trữ dữ liệu cho cả tài xế và công nhân

### 2. Xuất sheet tháng/năm
- Tự động điền weekday (T2-CN)
- Fill màu Chủ Nhật
- Điền "CT" (màu đỏ) cho ngày công tác
- Điền "1" cho ngày làm việc bình thường
- Ghi chú công tác (giữ nguyên địa điểm)

### 3. Xuất sheet Chủ Nhật
- Tự động điền header với các ngày Chủ Nhật trong tháng
- Lọc người làm việc ngày Chủ Nhật
- Điền dấu "X" vào ngày làm việc
- Ghi chú địa điểm công tác

### 4. Xử lý tên
- Loại bỏ số điện thoại khỏi tên: `Lê Tuấn Anh (0947 635 029)` → `Lê Tuấn Anh`
- Giữ nguyên dấu tiếng Việt
- Normalize để so khớp (bỏ dấu, uppercase)

## Code Structure

### Main Functions

#### `parse_date_range(date_str, target_month, target_year)`
Parse chuỗi ngày thành list các ngày:
- `"03-06/11/2025"` → `[3, 4, 5, 6]` (liên tục)
- `"10-11-13/12/2025"` → `[10, 11, 13]` (rời rạc)
- `"01/11/2025"` → `[1]`

#### `collect_work_data(input_path, target_month, target_year)`
Đọc dữ liệu từ sheet "Nhap Lieu", trả về dict:
```python
{
    "NGUYEN VAN A": {
        "days": {1, 3, 5, 7},
        "notes": ["1-7(TP HCM)", ...],
        "original_name": "Nguyễn Văn A"
    }
}
```

#### `fill_template(template_path, work_data, target_month, target_year, output_path)`
Fill dữ liệu vào sheet tháng/năm

#### `fill_sunday_sheet(wb, work_data, target_month, target_year)`
Fill dữ liệu vào sheet "Chủ Nhật"

#### `get_sundays_in_month(year, month)`
Tính các ngày Chủ Nhật trong tháng

## Build Executable

### Quick Build
```bash
python build_exe.py
```

### Manual Build
```bash
pip install -r requirements.txt
pip install pyinstaller
pyinstaller --clean ChamCongCongTac.spec
```

File exe nằm trong `dist/ChamCongCongTac.exe`

## Compatibility

- **Python**: 3.8 - 3.11 (để hỗ trợ Windows 7)
- **Windows**: 7 SP1, 8, 10, 11
- **Excel**: .xlsx format (OpenXML)

## Testing

Đã test với:
- Windows 10 ✓
- Python 3.10 ✓
- openpyxl 3.1.5 ✓
- Pillow 10.1.0 ✓

## Notes

- File exe chứa toàn bộ dependencies, không cần Python trên máy user
- Lần đầu chạy exe có thể mất 5-10 giây (giải nén libraries)
- Windows Defender có thể cảnh báo với exe mới build
- Kích thước exe: ~30MB (đã nén với UPX)

## Future Improvements

- [ ] Thêm validation cho input data
- [ ] Export PDF report
- [ ] Multi-sheet support
- [ ] Custom template builder
- [ ] Error logging to file

## License

Internal use only.

## Author

Created: December 2025
