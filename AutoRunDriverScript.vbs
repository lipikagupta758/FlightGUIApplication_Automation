Dim objShell
' Create a Shell object
Set objShell = Wscript.CreateObject("WScript.Shell")
objShell.Run "DriverScript.vbs" 
Set objShell = Nothing