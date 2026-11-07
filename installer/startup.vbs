' WordNote startup toggle (ASCII only)
' usage: cscript //nologo startup.vbs add <installDir> <exePath> <iconPath>
'        cscript //nologo startup.vbs remove
Option Explicit
Dim sh, fso, mode, lnkPath
Set sh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
If WScript.Arguments.Count < 1 Then WScript.Quit 2
mode = LCase(WScript.Arguments(0))
lnkPath = sh.SpecialFolders("Startup") & "\WordNote.lnk"
If mode = "remove" Then
    If fso.FileExists(lnkPath) Then
        fso.DeleteFile lnkPath, True
        WScript.Echo "OK: removed " & lnkPath
    Else
        WScript.Echo "INFO: not present"
    End If
    WScript.Quit 0
End If
If WScript.Arguments.Count < 4 Then WScript.Quit 2
Dim installDir, exePath, iconPath, sc
installDir = WScript.Arguments(1)
exePath = WScript.Arguments(2)
iconPath = WScript.Arguments(3)
Set sc = sh.CreateShortcut(lnkPath)
sc.TargetPath = exePath
sc.WorkingDirectory = installDir
sc.IconLocation = iconPath & ",0"
sc.Description = "WordNote"
sc.WindowStyle = 1
sc.Save
WScript.Echo "OK: " & lnkPath