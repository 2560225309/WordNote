@echo off
chcp 936 >nul
setlocal
set "ROOT=%~dp0"
if defined WORDNOTE_DEST (set "DEST=%WORDNOTE_DEST%") else (set "DEST=%LOCALAPPDATA%\WordNote")
cscript //nologo "%ROOT%tools\startup.vbs" remove
echo 已取消开机自启。按任意键关闭。
pause >nul