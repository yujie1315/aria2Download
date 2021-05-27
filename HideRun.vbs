Set wshell=CreateObject("WScript.Shell")
wshell.Run """E:/toolBox/aria2/aria2c.exe"""&" --conf-path=aria2.conf",0
Set wshell = Nothing