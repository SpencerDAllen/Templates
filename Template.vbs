' Spencer Allen
' 6/22/13

Option Explicit 
 
Dim StartTime,EndTime: StartTime = Now  
Dim objShell 
Dim objNet 
Dim objFSO 
Set objShell = WScript.CreateObject("WScript.Shell") 
Set objNet = Wscript.CreateObject("WScript.Network")  
Set objFSO = CreateObject("Scripting.FileSystemObject") 
 
Wscript.Echo "StartTime = " & StartTime 
' ***************************************************************** '


' VBScript Reference (strings, loops, arrays, subs, functions, ect...)
' ***************************************************************** '
' http://technet.microsoft.com/en-us/library/ee198844.aspx

' Wscript usage Object Properties and Methods (arguments, interactive, ect..)
' ***************************************************************** '
' http://msdn.microsoft.com/en-us/library/2795740w(v=vs.84).aspx

' objShell usage Object Methods (exec, registry, remove connections, ect...)
' ***************************************************************** '
' http://msdn.microsoft.com/en-us/library/2x3w20xf(v=vs.84).aspx

' objNet usage Object Properties and Methods (names, add connections, ect...)
' ***************************************************************** '
' http://msdn.microsoft.com/en-us/library/907chf30(v=vs.84).aspx

' objFSO Object Methods (folder and filesystem manipulation)
' ***************************************************************** '
' http://msdn.microsoft.com/en-us/library/6tkce7xa(v=vs.84).aspx


' ***************************************************************** ' 
EndTime = Now 
Wscript.Echo vbCrLf & "EndTime = " & EndTime 
Wscript.Echo "Seconds Elapsed: " & DateDiff("s", StartTime, EndTime) 
Wscript.Echo "Script Complete" 
Wscript.Quit(0) 