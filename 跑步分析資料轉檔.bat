@echo off
chcp 65001 >nul
setlocal

cd /d "%~dp0"

set "PYTHON_CMD="

where py >nul 2>nul
if not errorlevel 1 (
  set "PYTHON_CMD=py -3"
) else (
  where python >nul 2>nul
  if not errorlevel 1 (
    set "PYTHON_CMD=python"
  )
)

if "%PYTHON_CMD%"=="" (
  echo.
  echo [Running Analytics Converter]
  echo 找不到 Python，無法啟動跑步分析資料轉檔工具。
  echo.
  echo 請先安裝 Python 3.11 以上：
  echo https://www.python.org/downloads/windows/
  echo.
  echo 安裝時請務必勾選：
  echo Add python.exe to PATH
  echo.
  echo 安裝完成後，請關閉這個視窗，再重新雙擊本檔案。
  echo.
  pause
  exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
  echo 建立 Python 虛擬環境...
  %PYTHON_CMD% -m venv .venv
  if errorlevel 1 (
    echo.
    echo 建立虛擬環境失敗。
    echo 請確認已安裝 Python 3.11 以上，並且安裝時有勾選 Add python.exe to PATH。
    echo.
    pause
    exit /b 1
  )
)

echo 安裝或更新必要套件...
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 (
  echo.
  echo 套件安裝失敗。請確認網路連線正常後再試一次。
  echo.
  pause
  exit /b 1
)

echo 啟動 Running Analytics Converter...
".venv\Scripts\python.exe" app.py

endlocal
