@echo off
:: Tìm đường dẫn trình biên dịch C# có sẵn của Windows
set "csc=C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe"

echo Dang tu dong bien dich file EXE, vui long cho...
%csc% /target:exe /out:RonaldJackToSheets.exe /platform:x86 Program.cs

echo.
echo === HOÀN THÀNH ===
echo Neu khong co thong bao loi mau do, file RonaldJackToSheets.exe da xuat hien!
pause
