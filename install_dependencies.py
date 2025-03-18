#!/usr/bin/env python3
import sys
import subprocess
import importlib.util
import os

def check_pip():
    """Kiểm tra xem pip đã được cài đặt chưa"""
    try:
        import pip
        return True
    except ImportError:
        return False

def install_package(package_name):
    """Cài đặt thư viện với pip"""
    print(f"Đang cài đặt {package_name}...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", package_name])
    print(f"Đã cài đặt {package_name} thành công!")

def is_package_installed(package_name):
    """Kiểm tra xem thư viện đã được cài đặt chưa sử dụng importlib"""
    try:
        spec = importlib.util.find_spec(package_name)
        return spec is not None
    except (ImportError, AttributeError):
        return False

def main():
    """Kiểm tra và cài đặt các thư viện cần thiết"""
    print("=" * 60)
    print("EXCEL2MARKDOWN - CÀI ĐẶT THƯ VIỆN")
    print("=" * 60)
    
    # Kiểm tra pip
    if not check_pip():
        print("Lỗi: Không tìm thấy pip. Vui lòng cài đặt Python và pip trước!")
        print("Truy cập: https://www.python.org/downloads/")
        input("\nNhấn Enter để thoát...")
        return 1
    
    # Danh sách thư viện cần thiết
    required_packages = [
        "pandas",
        "openpyxl",
        "tabulate",
        "numpy"
    ]
    
    # Kiểm tra và cài đặt các thư viện
    need_install = []
    for package in required_packages:
        if not is_package_installed(package):
            need_install.append(package)
    
    if not need_install:
        print("\nTất cả thư viện cần thiết đã được cài đặt!")
    else:
        print(f"\nCần cài đặt {len(need_install)} thư viện:")
        for package in need_install:
            print(f"- {package}")
        
        print("\nĐang bắt đầu cài đặt...\n")
        
        # Cài đặt từng thư viện
        for package in need_install:
            try:
                install_package(package)
            except Exception as e:
                print(f"Lỗi khi cài đặt {package}: {str(e)}")
                print("Vui lòng thử cài đặt thủ công với lệnh:")
                print(f"pip install {package}")
        
        # Kiểm tra lại sau khi cài đặt
        all_installed = True
        for package in need_install:
            if not is_package_installed(package):
                all_installed = False
                print(f"Thư viện {package} chưa được cài đặt thành công!")
        
        if all_installed:
            print("\nTất cả thư viện đã được cài đặt thành công!")
    
    print("\nBạn đã sẵn sàng sử dụng Excel2Markdown!")
    print("Có thể chạy Convert_Excel_To_Markdown.command (Mac) hoặc Convert_Excel_To_Markdown.bat (Windows)")
    input("\nNhấn Enter để thoát...")
    return 0

if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as e:
        print(f"\nĐã xảy ra lỗi: {str(e)}")
        input("\nNhấn Enter để thoát...")