# Excel2Markdown

Công cụ đơn giản để chuyển đổi file Excel sang định dạng Markdown.

## Tính năng

- Chuyển đổi file Excel sang bảng Markdown
- Hỗ trợ nhiều sheet trong một file Excel
- Tùy chọn bao gồm/loại trừ cột index
- Đầu ra có thể là file hoặc hiển thị trực tiếp trên console
- Giao diện dòng lệnh và GUI để dễ dàng sử dụng
- Hỗ trợ các tính năng nâng cao:
  - Xử lý các ô đã được merge
  - Giữ định dạng (bold, italic)
  - Căn chỉnh các cột (trái, giữa, phải)
  - Hiển thị công thức (tùy chọn)
- Chuyển đổi hàng loạt (batch convert) nhiều file Excel cùng lúc
- Tự động mở thư mục chứa file sau khi chuyển đổi thành công
- Tự động xử lý lỗi và cung cấp thông báo chi tiết

## Yêu cầu

- Python 3.6 trở lên
- Thư viện pandas, tabulate, openpyxl, numpy

## Cài đặt

```bash
# Clone repository
git clone https://github.com/username/Excel2Markdown.git
cd Excel2Markdown

# Cài đặt các thư viện cần thiết
pip install -r requirements.txt
```

## Sử dụng

### Dòng lệnh cơ bản

```bash
python excel2markdown.py path/to/file.xlsx
```

Lệnh này sẽ chuyển đổi tất cả các sheet trong file Excel và lưu kết quả vào file có cùng tên nhưng với phần mở rộng `.md`.

### Tùy chọn dòng lệnh

```bash
# Chỉ định file đầu ra
python excel2markdown.py path/to/file.xlsx -o output.md

# Chỉ chuyển đổi một sheet cụ thể
python excel2markdown.py path/to/file.xlsx -s "Sheet1"

# Bao gồm cột index trong kết quả
python excel2markdown.py path/to/file.xlsx -i
```

### Giao diện đồ họa (GUI)

```bash
python excel2markdown_gui.py
```

Giao diện GUI cung cấp các tính năng:
- Chọn file Excel đầu vào
- Chọn file Markdown đầu ra
- Xem trước kết quả Markdown
- Chọn sheet cụ thể hoặc tất cả các sheet
- Tùy chọn bao gồm/loại trừ cột index

### Chuyển đổi hàng loạt (batch convert)

```bash
python batch_convert.py
```

Hoặc sử dụng file thực thi:
- Windows: Chạy file `Convert_Excel_To_Markdown.bat`
- macOS: Chạy file `Convert_Excel_To_Markdown.command`

Tính năng chuyển đổi hàng loạt giúp:
- Tự động chuyển đổi tất cả các file Excel (.xlsx, .xls) trong thư mục `input/`
- Lưu các file Markdown trong thư mục `output/`
- Tự động tạo thư mục `input/` và `output/` nếu chưa tồn tại
- Hiển thị thông tin về tiến trình và kết quả chuyển đổi
- Tự động mở thư mục chứa file sau khi chuyển đổi thành công

### Sử dụng bộ chuyển đổi nâng cao

```bash
python advanced_converter.py path/to/file.xlsx output.md [sheet_name]
```

Chuyển đổi nâng cao hỗ trợ:
- Xử lý các ô đã được merge
- Giữ định dạng (bold, italic)
- Căn chỉnh các cột (trái, giữa, phải)

## Cách cài đặt trên Windows và macOS

### Windows
1. Tải xuống và cài đặt [Python](https://www.python.org/downloads/) (đảm bảo chọn "Add Python to PATH" trong quá trình cài đặt)
2. Tải xuống và giải nén Excel2Markdown
3. Chạy file `Install_Dependencies.bat` để cài đặt các thư viện cần thiết
4. Chạy file `Convert_Excel_To_Markdown.bat` để sử dụng chế độ chuyển đổi hàng loạt

### macOS
1. macOS thường đã có sẵn Python. Nếu chưa có, tải xuống và cài đặt [Python](https://www.python.org/downloads/)
2. Tải xuống và giải nén Excel2Markdown
3. Mở Terminal trong thư mục Excel2Markdown
4. Chạy `chmod +x *.command` để cấp quyền thực thi cho các file command
5. Chạy file `Install_Dependencies.command` để cài đặt các thư viện cần thiết
6. Chạy file `Convert_Excel_To_Markdown.command` để sử dụng chế độ chuyển đổi hàng loạt

## Xử lý sự cố

### Các vấn đề thường gặp

1. **Thiếu thư viện**
   - Chạy `Install_Dependencies.bat` (Windows) hoặc `Install_Dependencies.command` (macOS)
   - Hoặc chạy lệnh `pip install -r requirements.txt` từ terminal/command prompt

2. **File không chuyển đổi được**
   - Đảm bảo file Excel không bị lỗi hoặc bảo vệ bằng mật khẩu
   - Kiểm tra quyền truy cập vào thư mục đầu ra

3. **Bảng Markdown hiển thị không đúng**
   - Đảm bảo file Excel tuân thủ định dạng bảng chuẩn
   - Thử sử dụng bộ chuyển đổi nâng cao thay vì bộ chuyển đổi cơ bản

## Đóng góp

Các đóng góp luôn được hoan nghênh! Nếu bạn muốn cải thiện dự án:

1. Fork dự án
2. Tạo nhánh tính năng (`git checkout -b feature/amazing-feature`)
3. Commit các thay đổi (`git commit -m 'Add some amazing feature'`)
4. Push lên nhánh của bạn (`git push origin feature/amazing-feature`)
5. Mở Pull Request

## Ví dụ

### File Excel

| Tên | Tuổi | Thành phố |
|-----|------|-----------|
| An  | 25   | Hà Nội    |
| Bình| 30   | TP HCM    |
| Cường | 22 | Đà Nẵng   |

### Kết quả Markdown

```markdown
| Tên    | Tuổi   | Thành phố   |
|:-------|:-------|:------------|
| An     | 25     | Hà Nội      |
| Bình   | 30     | TP HCM      |
| Cường  | 22     | Đà Nẵng     |
```

## Hướng dẫn cho nhà phát triển

Dự án Excel2Markdown bao gồm các module chính:

1. `excel2markdown.py` - Module cơ bản cho chuyển đổi Excel sang Markdown
2. `excel2markdown_gui.py` - Giao diện đồ họa người dùng
3. `advanced_converter.py` - Bộ chuyển đổi nâng cao xử lý các trường hợp phức tạp

### Mở rộng ứng dụng

Để thêm tính năng mới, bạn có thể mở rộng các lớp hiện có hoặc tạo plugin mới.

## Giấy phép

MIT 