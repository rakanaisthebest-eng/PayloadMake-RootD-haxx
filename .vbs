
dteWait = DateAdd("s", 5, Now())
	

Dim dteWait
Do Until (Now() > dteWait)
objShell.SendKeys "p"
objShell.SendKeys "{ENTER}"
objShell.SendKeys "fuh you"
objShell.SendKeys "{ENTER}"
Loop

Set WshShell = CreateObject("WScript.Shell")
WshShell.Run ".bat"

Set WshShell = Nothing
