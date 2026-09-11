@echo off
chcp 65001 >nul
if "%~1"=="" (
  echo 請把 Excel 檔拖曳到本檔案上執行
  pause & exit /b
)
python "%~dp0mask_names.py" "%~1"
pause
