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
