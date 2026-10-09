@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
echo.
echo  ==== Ket noi thu muc voi GitHub (chay 1 lan) ====
echo.
where git >nul 2>nul
if errorlevel 1 (
  echo  Chua cai Git. Tai va cai tai: https://git-scm.com/download/win
  echo  Cai xong, chay lai file nay.
  pause
  exit /b 1
)
if not exist ".git" git init
for /f "delims=" %%a in ('git config user.name') do set GN=%%a
if "%GN%"=="" (
  set /p GN=Nhap ten hien thi tren GitHub: 
  set /p GE=Nhap email GitHub: 
)
if not "%GE%"=="" (
  git config user.name "%GN%"
  git config user.email "%GE%"
)
git remote remove origin >nul 2>nul
git remote add origin https://github.com/PVI-AviationDivision/dashboard_VST_BHMDM_PVI.git
git add index.html data.js assets scripts update.bat setup-github.bat .gitignore README.md
git commit -m "Khoi tao dashboard BH Mot doi mot - Viettel Store"
git branch -M main
echo.
echo  Dang day len GitHub (lan dau co the hien cua so dang nhap)...
git push -u origin main
if errorlevel 1 (
  echo.
  echo  LOI khi push. Kiem tra tai khoan co quyen ghi vao repo PVI-AviationDivision.
  pause
  exit /b 1
)
echo.
echo  XONG. Vao repo ^> Settings ^> Pages ^> Branch: main / (root) ^> Save
echo  Link dashboard: https://pvi-aviationdivision.github.io/dashboard_VST_BHMDM_PVI/
pause
