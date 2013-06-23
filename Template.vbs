' Spencer Allen
' 6/22/13

Option Explicit 
 
Dim StartTime,EndTime: StartTime = Now  
Dim oShell
Dim oNet 
Dim oFSO 
Set oShell = WScript.Createobject("WScript.Shell") 
Set oNet = Wscript.Createobject("WScript.Network")  
Set oFSO = Createobject("Scripting.FileSystemobject") 

Wscript.Echo "StartTime = " & StartTime 
' ***************************************************************** '


' VBScript Reference (strings, loops, arrays, subs, functions, ect...)
' ***************************************************************** '
' http://technet.microsoft.com/en-us/library/ee198844.aspx

' Wscript usage object Properties and Methods (arguments, interactive, ect..)
' ***************************************************************** '
' http://msdn.microsoft.com/en-us/library/2795740w(v=vs.84).aspx

' oShell usage object Methods (exec, registry, remove connections, ect...)
' ***************************************************************** '
' http://msdn.microsoft.com/en-us/library/2x3w20xf(v=vs.84).aspx

' oNet usage object Properties and Methods (names, add connections, ect...)
' ***************************************************************** '
' http://msdn.microsoft.com/en-us/library/907chf30(v=vs.84).aspx

' oFSO object Methods (folder and filesystem manipulation)
' ***************************************************************** '
' http://msdn.microsoft.com/en-us/library/6tkce7xa(v=vs.84).aspx


' ***************************************************************** ' 
EndTime = Now 
Wscript.Echo vbCrLf & "EndTime = " & EndTime 
Wscript.Echo "Seconds Elapsed: " & DateDiff("s", StartTime, EndTime) 
Wscript.Echo "Script Complete" 
Wscript.Echo "To Quit press Enter..."
Wscript.Stdin.Readline 
Wscript.Quit(0)