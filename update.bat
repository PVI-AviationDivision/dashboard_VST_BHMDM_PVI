@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
echo.
echo  ==== Cap nhat dashboard BH Mot doi mot - Viettel Store ====
echo.
echo  [1/2] Doc file bao cao bhkh-*.xlsx trong thu muc BaoCao ...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\build-data.ps1"
if errorlevel 1 (
  echo.
  echo  LOI: khong doc duoc file bao cao. Kiem tra lai file xlsx.
  pause
  exit /b 1
)
echo.
echo  [2/2] Day len GitHub ...
where git >nul 2>nul
if errorlevel 1 (
  echo  Chua cai Git - bo qua buoc day len. Mo index.html de xem.
  goto done
)
if not exist ".git" (
  echo  Thu muc chua ket noi GitHub - hay chay setup-github.bat truoc.
  goto done
)
git add index.html data.js assets scripts update.bat setup-github.bat .gitignore README.md BaoCao/DOC-TRUOC.txt
git commit -m "Cap nhat so lieu %date% %time:~0,5%" >nul
if errorlevel 1 echo  Khong co thay doi moi de commit.
git push
if errorlevel 1 (
  echo  LOI khi push len GitHub. Kiem tra dang nhap / mang.
  pause
  exit /b 1
)
echo  Da day len GitHub. Trang se cap nhat sau khoang 1 phut.
:done
echo.
start "" "%~dp0index.html"
timeout /t 5 >nul
