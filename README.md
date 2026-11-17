# Chuyển đổi tài liệu

Ứng dụng Windows giúp chuyển đổi Word, PDF và ảnh ngay trên máy. Dự án phục vụ các thao tác thường gặp như xuất báo cáo Word sang PDF, gộp ảnh chứng từ thành một tài liệu, hoặc lấy từng trang PDF để chèn vào bản trình bày. Ứng dụng có giao diện tiếng Việt và xử lý tệp cục bộ.

## Chuyển đổi ngay từ menu chuột phải

Sau khi cài tích hợp trên Windows 11 x64, bấm chuột phải vào một tệp được hỗ trợ ngay trong File Explorer, chọn **Chuyển đổi tài liệu**, rồi chọn định dạng đích. Bạn có thể bắt đầu ngay tại thư mục đang xem mà không cần mở cửa sổ chính của ứng dụng và tìm lại tệp.

Ví dụ, để gửi một báo cáo Word dưới dạng PDF, bấm chuột phải vào **BaoCao.docx** → **Chuyển đổi tài liệu** → **PDF**. Hộp thoại lưu mở ngay ở thư mục của tệp gốc và điền sẵn tên **BaoCao.pdf**. Bạn chỉ cần bấm **Lưu**. Cửa sổ tiến trình tự đóng khi hoàn tất.

Menu xử lý từng tệp một. Khi xuất PDF thành ảnh từ menu, ứng dụng dùng 300 DPI.

## Tính năng

| Chế độ | Cách hoạt động |
| --- | --- |
| Word → PDF | Xuất DOC hoặc DOCX bằng Microsoft Word để giữ cách dàn trang của tài liệu. |
| PDF → Word | Tạo DOCX từ PDF. Với PDF có metadata nhận diện Microsoft PowerPoint, ứng dụng dùng ảnh nền cho từng trang và đặt chữ ngang có sẵn vào các hộp văn bản có thể chỉnh sửa. |
| Ảnh → PDF | Gộp ảnh JPG, PNG, WebP, BMP, TIFF, GIF, JPEG 2000 hoặc AVIF thành một PDF. Có thể sắp xếp lại ảnh trước khi xuất. |
| PDF → Ảnh | Xuất từng trang thành PNG, JPEG, TIFF, BMP hoặc WebP ở 72, 150, 200, 300 hoặc 600 DPI. |

Bố cục PDF → Word có thể thay đổi, nhất là với bảng phức tạp hoặc phông chữ thiếu trên máy. PDF scan cần được OCR trước nếu muốn sửa chữ trong Word.

## Chạy từ mã nguồn

Cần Windows và Python 3.13 có Tkinter. Riêng chế độ Word → PDF cần Microsoft Word bản desktop đã được mở và thiết lập ít nhất một lần.

Mở PowerShell tại thư mục repo, rồi chạy các lệnh sau.

~~~powershell
py -3.13 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe app.py
~~~

Chọn chế độ, bấm **Chọn tệp…**, rồi bấm **Chuyển đổi…** và chọn nơi lưu. Khi gộp ảnh, dùng **Lên** và **Xuống** để đổi thứ tự trang. Sau khi hoàn tất, bấm **Mở kết quả**.

## Đóng gói ứng dụng và menu Explorer

Có thể tạo tệp chạy độc lập bằng PyInstaller trên Windows x64.

~~~powershell
.\.venv\Scripts\python.exe -m pip install pyinstaller==6.22.3
.\.venv\Scripts\python.exe -m PyInstaller --onefile --windowed --name ChuyenDoiTaiLieu --icon assets/app.ico --add-data "assets;assets" --collect-all pypdfium2_raw --collect-all pypdfium2 app.py
~~~

Tệp chạy được tạo tại **dist/ChuyenDoiTaiLieu.exe**.

Để tạo gói menu chuột phải cho Windows 11 x64, cần LLVM MinGW x64 và MakeAppx từ Windows SDK. Điền đường dẫn thực tế trên máy vào hai biến dưới đây.

~~~powershell
$cxx = 'C:\path\to\x86_64-w64-mingw32-clang++.exe'
$makeappx = 'C:\path\to\MakeAppx.exe'

.\.venv\Scripts\python.exe -m PyInstaller --onedir --windowed --name ChuyenDoiTaiLieu --icon assets/app.ico --add-data "assets;assets" --collect-all pypdfium2_raw --collect-all pypdfium2 app.py
.\.venv\Scripts\python.exe explorer\build.py --cxx $cxx --makeappx $makeappx
.\.venv\Scripts\python.exe explorer\build_setup.py --cxx $cxx
~~~

Chạy **dist/CaiDat-ChuyenDoiTaiLieu.exe** và xác nhận quyền quản trị để cài gói MSIX chưa ký do bạn tự build. Để gỡ, vào **Settings → Apps → Installed apps → Chuyển đổi tài liệu**.

## Cấu trúc mã nguồn

- **app.py** dựng giao diện Tkinter và cập nhật tiến trình từ luồng chuyển đổi.
- **converter.py** xử lý chuyển đổi bằng Word COM, pdf2docx, img2pdf và PDFium. Kết quả được ghi vào vị trí tạm trước khi chuyển đến đích.
- **slide_docx.py** dựng trang Word từ PDF xuất từ PowerPoint bằng PyMuPDF và python-docx.
- **explorer_menu.py** xử lý chọn nơi lưu và trạng thái chuyển đổi khi mở từ Explorer.
- **explorer/** chứa mã C++ cho lệnh Explorer và bộ cài, cùng manifest và script đóng gói.
