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

## Yêu cầu

- Python 3.6 trở lên
- Thư viện pandas, tabulate, openpyxl

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

### Sử dụng bộ chuyển đổi nâng cao

```bash
python advanced_converter.py path/to/file.xlsx output.md [sheet_name]
```

Chuyển đổi nâng cao hỗ trợ:
- Xử lý các ô đã được merge
- Giữ định dạng (bold, italic)
- Căn chỉnh các cột (trái, giữa, phải)

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