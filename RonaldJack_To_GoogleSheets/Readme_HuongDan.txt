HUONG DAN SU DUNG GOI KET NOI MAY CHAM CONG
=====================================================
Gói mã nguồn này đã được cấu hình sẵn theo các thông số bạn cung cấp:
- Gửi dữ liệu về Google Sheet ID: 1LshZZc_Pqf2hTq6CCC-6y69RjOPZoylYbGroaAqUrNM
- Qua Web App URL: https://script.google.com/macros/s/AKfycbz_qDgGg-iEfVDmyzf-YzNrr8720YsYDXorpD7xP_C8ptGB-GQchpzI4tXelr86VMhaZA/exec
- Quét đồng thời 8 máy chấm công (X1, X2, X3, X4, X5, VP, VP2, HCM).

CAC FILE TRONG GOI ZIP:
1. Program.cs      : Mã nguồn ứng dụng C# console chạy đa luồng để kết nối 8 máy chấm công cùng lúc và đẩy dữ liệu lên Google Sheets.
2. GoogleScript.js : Mã nguồn Apps Script của bạn (đã được cập nhật hàm định danh openById để ghi chính xác vào Sheet của bạn).
3. DangKyThuvien.bat: File chạy đăng ký nhanh zkemkeeper.dll vào Windows.

HUONG DAN CAI DAT & CHAY:
Bước 1: Giải nén toàn bộ thư mục này.
Bước 2: Click chuột phải vào file `DangKyThuvien.bat` -> Chọn "Run as Administrator" để đăng ký thư viện kết nối phần cứng máy chấm công.
Bước 3: Tạo một Project C# Console (.NET Framework 4.7.2 hoặc cao hơn) trong Visual Studio, sao chép nội dung file `Program.cs` vào dự án của bạn và thực hiện Build thành file .exe để sử dụng.
