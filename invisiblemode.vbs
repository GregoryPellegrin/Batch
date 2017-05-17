Set WshShell = CreateObject("WScript.Shell") 
WshShell.Run chr(34) & "f:\deleteall.exe" & Chr(34), 0
Set WshShell = Nothing
