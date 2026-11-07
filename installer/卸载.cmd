@echo off
chcp 936 >nul
setlocal EnableExtensions
title WordNote 卸载

set "ROOT=%~dp0"
set "TOOLS=%ROOT%tools"
if defined WORDNOTE_DEST (set "DEST=%WORDNOTE_DEST%") else (set "DEST=%LOCALAPPDATA%\WordNote")

echo.
echo   WordNote  卸载
echo   -------------------------------------------------------
echo   安装位置：%DEST%
echo.

echo [1/4] 关闭 WordNote ...
taskkill /IM WordNote.exe /F >nul 2>&1
if errorlevel 1 (echo        没有正在运行的实例) else (echo        已结束进程)

echo [2/4] 删除快捷方式 ...
if exist "%TOOLS%\unshortcut.vbs" (
    cscript //nologo "%TOOLS%\unshortcut.vbs" "WordNote"
) else (
    echo        跳过（缺少 unshortcut.vbs）
)

echo [3/4] 处理配置 ...
set "KEEP=1"
if defined WORDNOTE_KEEP_CONFIG set "KEEP=1"
set /p "ANS=是否保留配置 WordNote.ini？(Y/n): "
if /i "%ANS%"=="n" set "KEEP=0"

echo [4/4] 删除文件 ...
if exist "%DEST%\WordNote.exe" del /f /q "%DEST%\WordNote.exe" >nul 2>&1
if exist "%DEST%\wordnote.ico" del /f /q "%DEST%\wordnote.ico" >nul 2>&1
if exist "%DEST%\README.md" del /f /q "%DEST%\README.md" >nul 2>&1
if "%KEEP%"=="0" if exist "%DEST%\WordNote.ini" del /f /q "%DEST%\WordNote.ini" >nul 2>&1

rem 目录空了就删掉
if "%KEEP%"=="0" (
    if exist "%DEST%" rd /s /q "%DEST%" >nul 2>&1
) else (
    dir /b "%DEST%" 2>nul | find /c /v "" >nul
)

echo.
if "%KEEP%"=="1" (
    echo   已卸载。配置保留在：%DEST%\WordNote.ini
) else (
    echo   已卸载，配置已一并删除。
)
echo.
pause
exit /b 0