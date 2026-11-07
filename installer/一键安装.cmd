@echo off
chcp 936 >nul
setlocal EnableExtensions
title WordNote 一键安装

set "ROOT=%~dp0"
set "PAYLOAD=%ROOT%payload"
set "TOOLS=%ROOT%tools"

if not exist "%PAYLOAD%\WordNote.exe" goto nopayload

if defined WORDNOTE_DEST (set "DEST=%WORDNOTE_DEST%") else (set "DEST=%LOCALAPPDATA%\WordNote")

echo.
echo   WordNote  一键安装
echo   -------------------------------------------------------
echo   安装到：%DEST%
echo.

echo [1/6] 检查运行环境 ...
reg query "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" /v Release >nul 2>&1
if errorlevel 1 goto nonet
echo        .NET Framework 已就绪

echo [2/6] 关闭正在运行的 WordNote ...
taskkill /IM WordNote.exe /F >nul 2>&1
if errorlevel 1 (echo        没有正在运行的实例) else (echo        已关闭旧实例)

echo [3/6] 创建安装目录 ...
if not exist "%DEST%" md "%DEST%"
if not exist "%DEST%" goto nodir

echo [4/6] 复制程序文件 ...
copy /y "%PAYLOAD%\WordNote.exe" "%DEST%\" >nul
copy /y "%PAYLOAD%\wordnote.ico" "%DEST%\" >nul
if exist "%ROOT%README.md" copy /y "%ROOT%README.md" "%DEST%\" >nul
if exist "%DEST%\WordNote.ini" (
    echo        检测到已有配置，保留原 WordNote.ini
) else (
    copy /y "%PAYLOAD%\WordNote.ini" "%DEST%\" >nul
    echo        已写入默认配置
)

echo [5/6] 创建快捷方式 ...
if not defined WORDNOTE_NO_SHORTCUT (
    if exist "%TOOLS%\shortcut.vbs" (
        cscript //nologo "%TOOLS%\shortcut.vbs" "%DEST%" "%DEST%\WordNote.exe" "%DEST%\wordnote.ico" "WordNote" >nul
        echo        桌面 + 开始菜单
    ) else (
        echo        跳过（缺少 shortcut.vbs）
    )
) else (
    echo        已按要求跳过
)

echo [6/6] 启动 ...
if not defined WORDNOTE_NO_LAUNCH (
    start "" "%DEST%\WordNote.exe"
    echo        已启动，托盘会出现蓝色小笔记图标
) else (
    echo        已按要求跳过
)

echo.
echo   安装完成。
echo.
echo   下一步：打开 Word 文档 -^> 按 Ctrl+Alt+S 框选视频区域（只需一次）
echo   之后 Ctrl+Alt+S 秒截进 Word，Ctrl+Alt+N 弹速记框。
echo   热键/外观都在：%DEST%\WordNote.ini
echo.
pause
exit /b 0

:nonet
echo.
echo   [错误] 未检测到 .NET Framework 4.x
echo   请先安装 .NET Framework 4.8 后重试：
echo   https://dotnet.microsoft.com/download/dotnet-framework/net48
echo.
pause
exit /b 1

:nodir
echo.
echo   [错误] 无法创建安装目录：%DEST%
echo   可能被安全软件拦截，或该路径无写入权限。
echo.
pause
exit /b 1

:nopayload
echo.
echo   [错误] 安装包不完整：缺少 payload\WordNote.exe
echo   请重新解压完整的安装包再运行。
echo.
pause
exit /b 1