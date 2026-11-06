@echo off
setlocal
cd /d "%~dp0"

set CSC=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe
if not exist "%CSC%" set CSC=%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe
if not exist "%CSC%" goto nocsc

set ICONOPT=/win32icon:wordnote.ico
if not exist wordnote.ico set ICONOPT=

"%CSC%" /nologo /target:winexe /platform:anycpu /codepage:65001 %ICONOPT% /out:WordNote.exe /r:System.dll /r:System.Windows.Forms.dll /r:System.Drawing.dll /r:Microsoft.CSharp.dll WordNote.cs
if errorlevel 1 goto buildfail

echo.
echo Build OK: WordNote.exe
goto end

:nocsc
echo csc.exe not found - .NET Framework is required.
pause
goto end

:buildfail
echo.
echo Build FAILED.
pause

:end
endlocal
