Set objshell = CreateObject("Wscript.shell") 
'objshell.Run "C:\Users\SVATHALU\Desktop\vbscript.txt"
objshell.Run "calc.exe"
WScript.Sleep(2000) 
'objshell.SendKeys("Good morning")
objshell.SendKeys(10)
WScript.Sleep(100) 
objshell.SendKeys("{+}")
WScript.Sleep(100) 
objshell.SendKeys(30)
WScript.Sleep(100) 
objshell.SendKeys("=")
WScript.Sleep(100) 
objshell.SendKeys("^c")
WScript.Sleep(100)
objshell.Run "C:\Users\SVATHALU\Desktop\vbscript.txt"
WScript.Sleep(2000)
objshell.SendKeys("^v")
MsgBox "done"