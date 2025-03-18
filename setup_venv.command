#!/bin/bash

# Hiển thị tiêu đề
echo "====================================================="
echo "EXCEL2MARKDOWN - THIẾT LẬP MÔI TRƯỜNG ẢO"
echo "====================================================="

# Chuyển đến thư mục chứa script này
cd "$(dirname "$0")"
CURRENT_DIR=$(pwd)

# Kiểm tra Python
if ! command -v python3 &> /dev/null; then
    echo "Lỗi: Không tìm thấy Python3. Vui lòng cài đặt Python trước!"
    echo "Truy cập: https://www.python.org/downloads/"
    read -p "Nhấn Enter để thoát..."
    exit 1
fi

echo "Đang sử dụng Python: $(python3 --version)"

# Tạo thư mục môi trường ảo nếu chưa tồn tại
VENV_DIR="$CURRENT_DIR/venv"
if [ ! -d "$VENV_DIR" ]; then
    echo -e "\nĐang tạo môi trường ảo..."
    python3 -m venv "$VENV_DIR"
    if [ ! -d "$VENV_DIR" ]; then
        echo "Lỗi: Không thể tạo môi trường ảo!"
        read -p "Nhấn Enter để thoát..."
        exit 1
    fi
    echo "Đã tạo môi trường ảo."
else
    echo -e "\nĐã tìm thấy môi trường ảo."
fi

# Kích hoạt môi trường ảo
echo -e "\nĐang kích hoạt môi trường ảo..."
source "$VENV_DIR/bin/activate"

# Nâng cấp pip
echo -e "\nĐang nâng cấp pip..."
pip install --upgrade pip

# Cài đặt các thư viện cần thiết
echo -e "\nĐang cài đặt các thư viện cần thiết..."
pip install pandas openpyxl tabulate numpy

# Kiểm tra cài đặt
echo -e "\nKiểm tra cài đặt thư viện:"
pip list | grep -E "pandas|openpyxl|tabulate|numpy"

# Tạo file wrapper để chạy ứng dụng với môi trường ảo
cat > "$CURRENT_DIR/Run_Excel2Markdown.command" << 'EOF'
#!/bin/bash

# Chuyển đến thư mục chứa script này
cd "$(dirname "$0")"
CURRENT_DIR=$(pwd)

# Kích hoạt môi trường ảo
source "$CURRENT_DIR/venv/bin/activate"

# Chạy chương trình chuyển đổi
python3 batch_convert.py

# Thoát môi trường ảo khi hoàn tất
deactivate
EOF

# Cấp quyền thực thi cho file wrapper
chmod +x "$CURRENT_DIR/Run_Excel2Markdown.command"

echo -e "\n====================================================="
echo "HOÀN TẤT THIẾT LẬP!"
echo "====================================================="
echo -e "\nBạn có thể chạy ứng dụng bằng cách double-click vào file:"
echo "Run_Excel2Markdown.command"
echo -e "\nChú ý: Hãy sử dụng file này thay vì Convert_Excel_To_Markdown.command"

# Thoát môi trường ảo khi hoàn tất
deactivate

read -p "Nhấn Enter để thoát..."
exit 0 