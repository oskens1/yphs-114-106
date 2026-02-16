@echo off
:: 設定編碼為 UTF-8 以支援中文
chcp 65001 >nul

echo ===========================================
echo   學生積分管理系統 - 一鍵啟動
echo ===========================================

:: 取得目前 bat 檔案所在的目錄
cd /d "%~dp0"

:: 偵測 Python 路徑 (嘗試使用您的 Anaconda 環境)
set PYTHON_EXE=python
if exist "C:\Users\osken\anaconda3\envs\omr_mini\python.exe" (
    set PYTHON_EXE="C:\Users\osken\anaconda3\envs\omr_mini\python.exe"
) else if exist "C:\Users\osken\anaconda3\python.exe" (
    set PYTHON_EXE="C:\Users\osken\anaconda3\python.exe"
)

echo 🚀 正在啟動系統...
%PYTHON_EXE% main.py

if %ERRORLEVEL% neq 0 (
    echo.
    echo ❌ 程式執行發生錯誤！
    pause
) else (
    echo.
    echo ✅ 任務正常結束。
    pause
)
