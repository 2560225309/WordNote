' WordNote uninstall helper (ASCII only)
' usage: cscript //nologo unshortcut.vbs [linkName]
Option Explicit

Dim sh, fso, linkName, p, n
linkName = "WordNote"
If WScript.Arguments.Count >= 1 Then linkName = WScript.Arguments(0)

Set sh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

n = 0
RemoveLink sh.SpecialFolders("Desktop") & "\" & linkName & ".lnk"
RemoveLink sh.SpecialFolders("Programs") & "\" & linkName & ".lnk"
RemoveLink sh.SpecialFolders("Startup") & "\" & linkName & ".lnk"

WScript.Echo "removed " & n & " shortcut(s)"

Sub RemoveLink(p)
    On Error Resume Next
    If fso.FileExists(p) Then
        fso.DeleteFile p, True
        If Err.Number = 0 Then
            n = n + 1
            WScript.Echo "OK: removed " & p
        Else
            WScript.Echo "WARN: " & p & " (" & Err.Description & ")"
            Err.Clear
        End If
    End If
End Sub
