@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo.
echo  ==== Doc lai TOAN BO file bao cao (BaoCao + BaoCao\DaNap) ====
echo  Chi dung khi so lieu bi sai hoac doi quy tac tinh.
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\build-data.ps1" -Rebuild
echo.
echo  Xong. Bam update.bat de day len GitHub.
pause
