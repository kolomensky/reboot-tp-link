Function PcOnline (strComputer)
'Check if the remote machine is online.
Dim objPing,objStatus
Dim TextStream, TimeVar
Dim fso, tf
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}")._
ExecQuery("select Replysize from Win32_PingStatus where address = '" & strComputer & "'")
For Each objStatus in objPing
If IsNull(objStatus.ReplySize) Then
PcOnline=False
'Create log file.
Set fso = CreateObject("Scripting.FileSystemObject")
If (fso.FileExists("d:\routerlog.txt")) Then
Set tf = fso.OpenTextFile("d:\routerlog.txt",ForAppending, True)
tf.WriteLine(Now() & " " & strComputer & " is down ")
tf.Close()
Set fso = Nothing
Set tf = Nothing
Else
Set tf = fso.CreateTextFile("d:\routerlog.txt",ForAppending, True)
tf.WriteLine(Now() & " " & strComputer & " is down ")
tf.Close()
Set fso = Nothing
Set tf = Nothing
End If
'Check access internet
Else
PcOnline = True
'Wscript.Echo strComputer & " is responding to a ping "
End If
Next
Set objPing=Nothing
Set objStatus=Nothing
End Function
Dim fsot, tft
Const ForReading = 1, ForWriting = 2, ForAppending = 8
If PcOnline("www.ya.ru")_
OR PcOnline("www.google.com")_
Then
'Write log, run reeboot
Wscript.Echo "all ok"
Set fsot = CreateObject("Scripting.FileSystemObject")
Set tft = fsot.OpenTextFile("d:\routerlog.txt",ForAppending, True)
tft.WriteLine("-----------------")
tft.Close()
Set fsot = Nothing
Set tft= Nothing
WScript.Quit 0
Else
'WScript.Echo "Reboot"
'Set oShell = WScript.CreateObject("WScript.Shell")
'oShell.Run "telnet.exe 1.1.0.1"
'WScript.Sleep 2000
'oShell.SendKeys "ADMIN" & chr(13)
'WScript.Sleep 2000
'oShell.SendKeys "ADMIN" & chr(13)
'WScript.Sleep 2000
'oShell.SendKeys "reboot" & chr(13)
'WScript.Sleep 2000
'oShell.SendKeys "^({]})q" & chr(13)
'WScript.Quit 255
End If
Set fsot = CreateObject("Scripting.FileSystemObject")
Set tft = fsot.OpenTextFile("d:\routerlog.txt",ForAppending, True)
tft.WriteLine("-----------------")
tft.Close()
Set fsot = Nothing
Set tft= Nothing
