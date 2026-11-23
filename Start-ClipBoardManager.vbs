' ---------------------------------------------------------------------------
'  Start-ClipBoardManager.vbs
'
'  DOUBLE-CLICK THIS FILE to start ClipBoardManager.
'
'  It starts the ClipBoardManager-CBM-x.y.z.ps1 file that sits in the same
'  folder, with NO console window and no visible flash. If the folder holds
'  more than one version, the highest version number is started, so this
'  file never needs editing when a newer .ps1 is dropped in.
'
'  Why this exists:
'     Double-clicking a .ps1 opens it in Notepad instead of running it, and
'     powershell.exe -WindowStyle Hidden still creates a console host and shows
'     it for a fraction of a second before hiding it. WScript.Shell.Run with a
'     window style of 0 never shows one at all.
'
'  Usage:
'     Double-click this file, or put a shortcut to it in
'     shell:startup  to launch at login.
'
'  The .ps1 must sit in the same folder as this script.
' ---------------------------------------------------------------------------

Option Explicit

Dim fso, shell, scriptDir, target, cmd
Dim pattern, file, parts, version, best

Set fso   = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)

' Look for ClipBoardManager-CBM-<major>.<minor>.<patch>.ps1 and keep the highest.
Set pattern = New RegExp
pattern.Pattern = "^ClipBoardManager-CBM-(\d+)\.(\d+)\.(\d+)\.ps1$"
pattern.IgnoreCase = True

target = ""
best = -1
For Each file In fso.GetFolder(scriptDir).Files
    If pattern.Test(file.Name) Then
        Set parts = pattern.Execute(file.Name)(0).SubMatches
        version = CDbl(parts(0)) * 1000000 + CDbl(parts(1)) * 1000 + CDbl(parts(2))
        If version > best Then
            best = version
            target = file.Path
        End If
    End If
Next

If target = "" Then
    MsgBox "Could not find a ClipBoardManager-CBM-x.y.z.ps1 file in:" & vbCrLf & vbCrLf & _
           scriptDir & vbCrLf & vbCrLf & _
           "Keep this launcher in the same folder as the .ps1 file.", _
           vbExclamation, "ClipBoardManager"
    WScript.Quit 1
End If

cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File """ & target & """"

' 0 = hidden window, False = do not wait for it to finish
shell.Run cmd, 0, False
