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
'   Date
'
' VERSIONS:
'   1.00 - Base Version (AuthorName)
'**********************************************************************************************

#End Region
Module main
#Region "Declorations"
    Dim LogFolder As String
    Dim NewLogFrequency As String
    Dim DaysToKeepLogs As Integer
    Dim e As Integer = 0 'E is for error
    Public Cls As New Cls
    Public L As Boolean ' L is for logging
#End Region
#Region "Main"

'*************************************************************************
Public Sub Main()
'PURPOSE:
'   To start the program.
'
'OUTPUT:
'   Unknown at this time.
'*************************************************************************
 'Start log
If L Then
    If Not Cls.Logs.start() Then
        MsgBox(Cls.Logs.MyException, MsgBoxStyle.Critical, "Fatal Error!")
        e = -1
        ExitCode(e)
    End If
End If
'*************************************************************************
'Your Code here.






'*************************************************************************
ExitCode(e)
End Sub

#End Region
#Region "ExitCode"

'*************************************************************************
Public Sub ExitCode(ByRef ErrorCode As Integer)
'PURPOSE:
'   To terminate the application
'
'OUTPUT:
'   Deletes old logs and adds a log entry that the application has terminated.
'   Collects all the garbage from memory.
'	Returns the ErrorCode for this application.
'*************************************************************************

'Logs
If L Then
    If Not ErrorCode = -1 Then
        Cls.Logs.clearOld(DaysToKeepLogs)
        Cls.Logs.finish()
'cleanup
        Cls = Nothing
        L = Nothing
        'DaysToKeepLogFiles = Nothing
'Forms.dispose
        GC.Collect()
'Exit
        Environment.Exit(ErrorCode)
    End If
End If

'No logs
'Cleanup
Cls = Nothing
L = Nothing
'DaysToKeepLogFiles = Nothing
'Forms.dispose
GC.Collect()
'Exit
Environment.Exit(ErrorCode)

End Sub

#End Region
End Module