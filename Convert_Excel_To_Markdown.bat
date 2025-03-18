@echo off
title Excel To Markdown Converter

:: Chuyển đến thư mục chứa script này
cd /d "%~dp0"

:: Kiểm tra Python đã được cài đặt chưa
python --version >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo Python không được tìm thấy. Vui lòng cài đặt Python trước khi chạy script này.
    echo Bạn có thể tải Python từ https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)

:: Chạy script Python
python batch_convert.py

:: Nếu xuất hiện lỗi
if %ERRORLEVEL% neq 0 (
    echo.
    echo Đã xảy ra lỗi khi chạy script.
    pause
    exit /b 1
)

exit /b 0 