@echo off
chcp 65001 >nul
if "%~1"=="" (
  echo 請把 Excel 檔（.xlsx）拖曳到本檔案上執行
  pause & exit /b
)
where python >nul 2>nul || (echo 找不到 Python，請先安裝 https://www.python.org/downloads/ 並勾選 Add to PATH & pause & exit /b)
python "%~dp0mask_names.py" %1
echo.
pause
