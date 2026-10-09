@echo off
echo ========================================================
echo DANG KY THU VIEN ZKEMKEEPER CHO MAY CHAM CONG
echo ========================================================
cd /d "%~dp0"
regsvr32 /s zkemkeeper.dll
if %errorlevel% equ 0 (
    echo [OK] Dang ky file zkemkeeper.dll thanh cong!
) else (
    echo [LOI] Vui long chay file nay voi quyen Administrator (Run as Administrator)!
)
pause
