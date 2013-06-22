#Region "Comments"

'**********************************************************************************************
'PURPOSE:
'   ?
'
'PREREQUISITES:
'   ?
'
'USAGE:
'   ?
'
'AUTHOR:
'   AuthorName
'   DTS Applications
'   Date
'
' VERSIONS:
'   1.00 - Base Version (AuthorName)
'**********************************************************************************************

#End Region
Module Main
#Region "Declorations"

Public Product As String = (Left(System.Reflection.Assembly.GetExecutingAssembly().ToString, _
							InStr(System.Reflection.Assembly.GetExecutingAssembly().ToString, _
							",") - 1)).Trim & ".exe"
Public LogFilePath As String
Public Logger As New clsLogFile
Public DaysToKeepLogFiles As Integer = 30
Const NewLogFileFrequency = "Always"  'Options = Always, Daily, Weekly, Monthly
Dim e As Integer = 0 'E is for error

#End Region
#Region "Main"

Sub Main()
	'Start log
	If Not Logger.start(LogFilePath, NewLogFileFrequency = "Always") Then
		e = -1
		ExitCode(e)
	End If
'*************************************************************************
'Your Code here.





'*************************************************************************
	ExitCode(e)
End Sub

#End Region
#Region "ExitCode"

'*************************************************************************
Private Sub ExitCode(ByRef ErrorCode As Integer)
'PURPOSE:
'   To terminate the application
'
'OUTPUT:
'   Deletes old logs and adds a log entry that the application has terminated.
'	Returns the ErrorCode for this Application.
'*************************************************************************
	If Not ErrorCode = -1 Then
		Logger.clearOld(LogFilePath, DaysToKeepLogFiles)
		Logger.finish(LogFilePath)
		Environment.Exit(ErrorCode)
	End If

	Environment.Exit(ErrorCode)
End Sub

#End Region
End Module
