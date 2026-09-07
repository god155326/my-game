' start-hidden.vbs
' Runs start.bat with NO visible black console window.
' Double-click this file instead of start.bat to launch the game server silently.
'
' To stop the server later, double-click stop-server.bat
' (there is no window to close, since this runs hidden).

Set fso = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

strFolder = fso.GetParentFolderName(WScript.ScriptFullName)
objShell.CurrentDirectory = strFolder

' 0 = hidden window, False = don't wait for it to finish
objShell.Run """" & strFolder & "\start.bat""", 0, False
