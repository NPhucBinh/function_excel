# function_excel
Các hàm function cho excel

Các bước thực hiện/chuẩn bị:
1. Cài đặt Python cho laptop/PC: https://www.python.org/downloads/

2. Cài lib vào Python: pip install [tên thư viện]
pip install RStockvn
pip install xlwings


3. Tải file xlsm và file python
https://github.com/NPhucBinh/function_excel/tree/main/function_excel


4. Thiết lập Excel
chạy cmd: xlwings addin install
Mở Excel vào File vào chọn More chọn Option chọn thẻ Trust Center chọn Trust Center Settings chọn Macro Settings, 
Mục Macro Settings Chọn Enable VBA marcos
Mục Developer Macros Seting chọn "Trust access to the VBA..."
Lưu lại đóng tab

Chọn thẻ xlwings đổi đường path PYTHONPATH thành địa chỉ thư mục lưu file, tiếp click "Import Funtions"


Đối với trường hợp dữ liệu không trả về dữ liệu dạng bảng thêm dòng sau @xw.ret(expand='table') sau mỗi @xw.func()\

### Liên hệ hỗ trợ nếu cần [Facebook](https://www.facebook.com/phuc.binh.3839/)