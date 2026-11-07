' WordNote shortcut helper (ASCII only)
' usage: cscript //nologo shortcut.vbs <installDir> <exePath> <iconPath> [linkName]
Option Explicit

Dim sh, fso, installDir, exePath, iconPath, linkName
Dim desktopLink, programLink

If WScript.Arguments.Count < 3 Then
    WScript.Echo "usage: shortcut.vbs <installDir> <exePath> <iconPath> [linkName]"
    WScript.Quit 2
End If

installDir = WScript.Arguments(0)
exePath = WScript.Arguments(1)
iconPath = WScript.Arguments(2)
linkName = "WordNote"
If WScript.Arguments.Count >= 4 Then linkName = WScript.Arguments(3)

Set sh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

If Not fso.FileExists(exePath) Then
    WScript.Echo "ERROR: target not found: " & exePath
    WScript.Quit 3
End If

desktopLink = sh.SpecialFolders("Desktop") & "\" & linkName & ".lnk"
programLink = sh.SpecialFolders("Programs") & "\" & linkName & ".lnk"

MakeLink desktopLink
MakeLink programLink

WScript.Echo "OK: " & desktopLink
WScript.Echo "OK: " & programLink

Sub MakeLink(p)
    On Error Resume Next
    Dim sc
    Set sc = sh.CreateShortcut(p)
    sc.TargetPath = exePath
    sc.WorkingDirectory = installDir
    sc.IconLocation = iconPath & ",0"
    sc.Description = "WordNote"
    sc.WindowStyle = 1
    sc.Save
    If Err.Number <> 0 Then
        WScript.Echo "WARN: cannot create " & p & " (" & Err.Description & ")"
        Err.Clear
    End If
End Sub
