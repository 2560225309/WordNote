@echo off
chcp 936 >nul
setlocal
set "ROOT=%~dp0"
if defined WORDNOTE_DEST (set "DEST=%WORDNOTE_DEST%") else (set "DEST=%LOCALAPPDATA%\WordNote")
if not exist "%DEST%\WordNote.exe" (
    echo 尚未安装 WordNote，请先运行 一键安装.cmd
    pause
    exit /b 1
)
cscript //nologo "%ROOT%tools\startup.vbs" add "%DEST%" "%DEST%\WordNote.exe" "%DEST%\wordnote.ico"
echo 已设置为开机自启。按任意键关闭。
pause >nul