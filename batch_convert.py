#!/usr/bin/env python3
import os
import sys
from pathlib import Path
import time

# Thêm thư mục hiện tại vào PYTHONPATH để import các module
current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.append(current_dir)

# Import converter từ advanced_converter.py (sẽ sử dụng bộ converter nâng cao)
from advanced_converter import convert_excel_advanced

def main():
    """
    Tự động chuyển đổi tất cả file Excel trong thư mục input sang Markdown trong thư mục output
    """
    print("=" * 60)
    print("EXCEL TO MARKDOWN CONVERTER")
    print("=" * 60)
    
    # Kiểm tra các thư viện cần thiết
    try:
        import pandas
        import numpy
        import openpyxl
        import tabulate
        print(f"✓ Pandas version: {pandas.__version__}")
        print(f"✓ Openpyxl version: {openpyxl.__version__}")
        print(f"✓ Tabulate version: {tabulate.__version__}")
        print(f"✓ Numpy version: {numpy.__version__}")
    except ImportError as e:
        print(f"Lỗi: Thiếu thư viện - {str(e)}")
        print("Vui lòng chạy Install_Dependencies.command (Mac) hoặc Install_Dependencies.bat (Windows)")
        input("\nNhấn Enter để thoát...")
        return
    
    # Đường dẫn tuyệt đối đến thư mục input và output
    input_dir = os.path.join(current_dir, "input")
    output_dir = os.path.join(current_dir, "output")
    
    # Kiểm tra thư mục input và output
    if not os.path.exists(input_dir):
        try:
            os.makedirs(input_dir)
            print(f"Đã tạo thư mục input: {input_dir}")
        except OSError as e:
            print(f"Lỗi khi tạo thư mục input: {str(e)}")
            input("\nNhấn Enter để thoát...")
            return
    
    if not os.path.exists(output_dir):
        try:
            os.makedirs(output_dir)
            print(f"Đã tạo thư mục output: {output_dir}")
        except OSError as e:
            print(f"Lỗi khi tạo thư mục output: {str(e)}")
            input("\nNhấn Enter để thoát...")
            return
    
    # Tìm tất cả file Excel trong thư mục input
    excel_files = []
    for file in os.listdir(input_dir):
        if file.endswith(".xlsx") or file.endswith(".xls"):
            excel_files.append(file)
    
    # Kiểm tra nếu không có file Excel nào
    if not excel_files:
        print("\nKhông tìm thấy file Excel nào trong thư mục input!")
        print(f"Vui lòng đặt file Excel vào thư mục: {input_dir}")
        
        # Mở thư mục input để người dùng có thể đặt file vào đó
        try:
            if sys.platform == 'win32':
                os.startfile(input_dir)
            elif sys.platform == 'darwin':  # macOS
                import subprocess
                subprocess.call(['open', input_dir])
            else:  # Linux
                import subprocess
                subprocess.call(['xdg-open', input_dir])
        except Exception:
            pass  # Nếu không mở được thư mục thì bỏ qua
            
        input("\nNhấn Enter để thoát...")
        return
    
    # Convert từng file Excel sang Markdown
    print(f"\nĐã tìm thấy {len(excel_files)} file Excel để chuyển đổi:")
    
    success_count = 0
    error_count = 0
    
    for i, excel_file in enumerate(excel_files, 1):
        input_path = os.path.join(input_dir, excel_file)
        
        # Tạo tên file output (thay đổi phần mở rộng từ .xlsx/.xls sang .md)
        output_file = Path(excel_file).stem + ".md"
        output_path = os.path.join(output_dir, output_file)
        
        print(f"\n{i}. Đang chuyển đổi: {excel_file} -> {output_file}")
        
        try:
            # Sử dụng advanced converter để chuyển đổi
            result = convert_excel_advanced(
                excel_file=input_path,
                output_file=output_path,
                include_formulas=True  # Bao gồm công thức trong file Markdown
            )
            
            if result:
                print(f"   ✓ Chuyển đổi thành công: {output_path}")
                success_count += 1
            else:
                print(f"   ✗ Chuyển đổi thất bại: {excel_file}")
                error_count += 1
                
        except Exception as e:
            print(f"   ✗ Lỗi khi chuyển đổi {excel_file}: {str(e)}")
            import traceback
            traceback.print_exc()
            error_count += 1
    
    # Hiển thị kết quả
    print("\n" + "=" * 60)
    print(f"KẾT QUẢ: Thành công: {success_count}, Thất bại: {error_count}")
    print("=" * 60)
    
    if success_count > 0:
        print(f"\nCác file Markdown đã được lưu trong thư mục: {output_dir}")
        
        # Mở thư mục output nếu có file thành công
        try:
            if sys.platform == 'win32':
                os.startfile(output_dir)
            elif sys.platform == 'darwin':  # macOS
                import subprocess
                subprocess.call(['open', output_dir])
            else:  # Linux
                import subprocess
                subprocess.call(['xdg-open', output_dir])
        except Exception:
            pass  # Nếu không mở được thư mục thì bỏ qua
    
    print("\nCảm ơn bạn đã sử dụng Excel2Markdown!")
    input("\nNhấn Enter để thoát...")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\nĐã xảy ra lỗi: {str(e)}")
        input("\nNhấn Enter để thoát...") 