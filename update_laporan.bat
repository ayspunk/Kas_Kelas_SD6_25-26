@echo off
title Update Laporan Kas SD6
cd /d "%~dp0"
echo ================================================
echo   UPDATE LAPORAN KAS SD6 - SAI Meruyung
echo ================================================
echo.
echo Membaca Excel dan membuat ulang data.js...
python generate_report.py
if errorlevel 1 (
    echo.
    echo [ERROR] Gagal menjalankan generate_report.py
    echo Pastikan Python sudah terinstall dan file Excel tidak sedang dibuka.
    echo.
    pause
    exit /b 1
)
echo.
echo Membuka laporan di browser...
start "" "Laporan_Pertanggungjawaban_Kas_SD6_SAI_Meruyung.html"
echo.
echo Selesai! Laporan sudah diperbarui.
pause
