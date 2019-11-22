Set shell=CreateObject("WScript.Shell")
Dim A(10)
Dim ab,bb
ab=inputbox("输入一个数字")
ab=int(ab)

'为什么后面不能使用括号?
MsgBox ab+1,65,"ab值加1"
'Set link=shell.CreateShortcut()
'answer=MsgBox("哈哈哈",65,"Example")