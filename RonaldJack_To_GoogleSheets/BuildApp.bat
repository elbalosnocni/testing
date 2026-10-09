@echo off
cd /d "%~dp0"

:: Định nghĩa lại đường dẫn chính xác tới trình biên dịch của Windows
set "csc=C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe"

:: Đường dẫn tới các thư viện mạng cốt lõi của hệ thống Windows
set "net_web=C:\Windows\Microsoft.NET\Framework\v4.0.30319\System.Net.Http.WebRequest.dll"

echo Dang tu dong bien dich file EXE, vui long cho...

:: Thực hiện lệnh biên dịch kèm theo đầy đủ tham chiếu hệ thống
"%csc%" /target:exe /out:RonaldJackToSheets.exe /platform:x86 /r:System.dll /r:System.Net.Http.dll /r:System.Net.dll /r:"%net_web%" /r:Interop.zkemkeeper.dll Program.cs

echo.
echo === HOÀN THÀNH ===
echo Kiem tra xem file RonaldJackToSheets.exe da xuat hien chua!
pause
