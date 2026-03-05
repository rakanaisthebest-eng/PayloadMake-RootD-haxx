
'here

	

' Add 5 seconds to the current time
dteWait = DateAdd("s", 5, Now())
	

Dim dteWait
Do Until (Now() > dteWait)
objShell.SendKeys "p"
objShell.SendKeys "{ENTER}"
objShell.SendKeys "fuh you"
objShell.SendKeys "{ENTER}"
Loop

WScript.Echo "Loop finished after approximately 5 seconds."
Set WshShell = CreateObject("WScript.Shell")
WshShell.Run ".bat"
Set WshShell = Nothing