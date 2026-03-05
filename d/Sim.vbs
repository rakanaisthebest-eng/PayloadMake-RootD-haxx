WScript.Echo "Starting loop-based delay..."

	

' Add 5 seconds to the current time
dteWait = DateAdd("s", 5, Now())
	

Dim dteWait
Do Until (Now() > dteWait)
Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "sim.bat"
Loop

WScript.Echo "Loop finished after approximately 5 seconds."
Set WshShell = Nothing