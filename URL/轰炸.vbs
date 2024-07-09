Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.AppActivate"ѩ"
for i=1 to 200
Wscript.Sleep 99
WshShell.SendKeys "^v"
WshShell.SendKeys "%s"
Next