#Region "Comments"

'**********************************************************************************************
'PURPOSE:
'   clsLogFile is a class used to produce useful log files.
'
'PREREQUISITES:
'   "Public LogFilePath As String" must be in the main program's global declarations.
'
'USAGE:
'   boolean = Logfile.start(FilePath As String, [NewLogFrequency As String], [LogFolder As String])     Options = Always, Daily, Weekly, Monthly
'   boolean = Logfile.addEvent(FilePath As String, EventDescription As String)
'   boolean = LogFile.finish(FilePath As String)       
'   Logfile.clearOld(FilePath, daysToKeep As Integer)                                                   Number of days to keep unused logs
'
'AUTHOR:
'   Spencer Allen
'   04-20-2007
'
' VERSIONS:
'   1.00 - Base Version (Spencer Allen)
'**********************************************************************************************

#End Region
Imports System
Imports System.IO
Public Class PwrClsLogs
#Region "Declorations"
Implements IDisposable
    Private LogFile As String

#End Region
#Region "Destructor"

'*************************************************************************
Public Sub Dispose() Implements System.IDisposable.Dispose
'PURPOSE:
'   To Remove all variables from memory.
'
'OUTPUT:
'   Memory is able to be reclamed.
'*************************************************************************

'Cleanup
    LogFile = Nothing
    Exec = Nothing

'Exit
End Sub

#End Region
#Region "Start"

'*************************************************************************
 Public Function start( _
                        Optional ByVal NewLogFrequency As String = "Always", _
                        Optional ByVal LogFolder As String = "Logs" _
                    ) As Boolean
'PURPOSE:
'   To start logging by creating a new log file or re-opening an existing one.
'
'OUTPUT:
'   Passes the log FilePath parameter out, adds log initiated event, and returns FALSE if an error occured.
'*************************************************************************
Dim Reflection As String = System.Reflection.Assembly.GetExecutingAssembly().ToString
Dim ProductName As String = Reflection.Substring(0, Reflection.IndexOf(","))
Dim logEvent As String = UCase(ProductName) & " STARTED"
Dim FilePath As String
Dim noErrorFlag = True 'return value for the function, returns false if an error occurs
Dim DT As DateTime = DateTime.Now  'Date and Time now
Dim strMMin As String = Right("0" & CStr(DT.Day), 2)
Dim strHH As String = Right("0" & CStr(DT.Hour), 2)
Dim strDD As String = Right("0" & CStr(DT.Day), 2)
Dim strMM As String = Right("0" & CStr(DT.Month), 2)
Dim strYYYY As String = Right("0" & CStr(DT.Year), 4)
Dim TrimChars() As Char = {"\", " "}

'If no value was passed set defalults
If LogFolder = "" Then
    LogFolder = "Logs"
End If

If NewLogFrequency = "" Then
    NewLogFrequency = "Always"
End If

'Set our LogFolder and noErrorFlag values
    LogFolder = Trim(LogFolder, TrimChars)
    noErrorFlag = CheckForDirectory(LogFolder)
'Get our File Name and Start the logs
    If noErrorFlag Then
        FilePath = GetFilePath(NewLogFrequency, LogFolder, ProductName, strYYYY, strMM, strDD, strHH, strMMin)
        LogFile = FilePath
        noErrorFlag = addEvent(logEvent)
    End If

'Cleanup
    FilePath = Nothing
    LogFolder = Nothing
    TrimChars = Nothing
    strYYYY = Nothing
    strMM = Nothing
    strDD = Nothing
    strHH = Nothing
    strMMin = Nothing
    DT = Nothing
    logEvent = Nothing
    ProductName = Nothing
    Reflection = Nothing
    NewLogFrequency = Nothing

'Exit
    Return noErrorFlag

 End Function

#End Region
#Region "Trim"

'*************************************************************************
Private Function Trim(ByVal LogFolder As String, _
                ByVal TrimChar() As Char) As String
'PURPOSE:
'   To trim charters off of the log folder string.
'
'OUTPUT:
'   Function returns a clean log folder string.
'*************************************************************************

  'Remove leading / trailing whitespace and trailing "\" if present
        LogFolder = LogFolder.TrimEnd(TrimChar)
        LogFolder = LogFolder.Trim

'Cleanup
    TrimChar = Nothing

'Exit
    Return LogFolder

End Function
#End Region
#Region "CheckForDirectory"

'*************************************************************************
Private Function CheckForDirectory(ByVal LogFolder As String) As Boolean
'PURPOSE:
'   To find or create a useable directory for the log folder.
'
'OUTPUT:
'   Function creates a useable directory for the log folder or returns an error message.
'*************************************************************************
Dim Returns As Boolean = True

'Ensure the log directory exists or can be created
    Pnm.Directory.LogOnOff = False
    If Not Pnm.Directory.Exists(LogFolder) Then
        If Not Pnm.Directory.Create(LogFolder) Then
            MsgBox("Unable to create Log Directory." & vbCrLf & Pnm.Directory.MyException, MsgBoxStyle.Critical, "Fatal Error")
            Returns = False
        End If
    End If
    Pnm.Directory.LogOnOff = True

'Cleanup
    LogFolder = Nothing

'Exit
    Return Returns

End Function

#End Region
#Region "GetFilePath"

'*************************************************************************
Private Function GetFilePath(ByVal newlogs As String, _
                    ByVal LogFolder As String, _
                    ByVal product As String, _
                    ByVal strYYYY As String, _
                    ByVal strMM As String, _
                    ByVal strDD As String, _
                    ByVal strHH As String, _
                    ByVal strMMin As String _
                    ) As String
'PURPOSE:
'   To Create the "FilePath" String.
'
'OUTPUT:
'   Function returns The string "FilePath"
'*************************************************************************
Dim Returns As String

'Generate Log name based on new log file frequency
    Returns = LogFolder & "\" & product
    Select Case newlogs
        Case "Daily"
            Returns = Returns & "_" & strYYYY & "-" & strMM & "-" & "Day-" & strDD
        Case "Weekly"
            Returns = Returns & "_" & strYYYY & "-" & strMM
            Returns = Returns & "_" & "Week-" & Left(CStr((strDD - 1) / 7) + 1, 1)
        Case "Monthly"
            Returns = Returns & "_" & strYYYY & "-" & "Month-" & strMM
        Case Else   'Always
            Returns = Returns & "_" & strYYYY & "-" & strMM & "-" & strDD
            Returns = Returns & "_" & strHH & "." & strMMin
    End Select
    Returns = Returns & ".txt"

'Cleanup
    newlogs = Nothing
    LogFolder = Nothing
    product = Nothing
    strYYYY = Nothing
    strMM = Nothing
    strDD = Nothing
    strHH = Nothing
    strHH = Nothing
    strMMin = Nothing

'Exit
    Return Returns

End Function

#End Region
#Region "AddEvent"

'*************************************************************************
 Public Function addEvent( _
          ByVal EventDescription As String _
           ) As Boolean
'PURPOSE:
'   Add the passed EventDescription to the log.
'
'OUTPUT:
'   Adds the passed EventDescription to the log, returns FALSE if an error occured.
'*************************************************************************
Dim logEvent = String.Format("{0:G}", Now()) & " : " & EventDescription
Dim noErrorFlag = True 'return value for the function, returns false if an error occurs
Dim fWriter As New StreamWriter(LogFile, True)

'Append date/time to Event Description and write to log file
    Try
        fWriter.WriteLine(logEvent)
        fWriter.Close()
        noErrorFlag = True
    Catch ex As Exception
        Exec = "Log File cannot be opened." & vbCrLf & ex.Message
        noErrorFlag = False
    End Try

'Cleanup
    fWriter = Nothing
    logEvent = Nothing
    EventDescription = Nothing

'Exit
  Return noErrorFlag

 End Function

#End Region
#Region "Finish"

'*************************************************************************
 Public Function finish( _
                        ) As Boolean
'PURPOSE:
'   To end logging.
'
'OUTPUT:
'   Adds log terminated event, returns FALSE if an error occured.
'*************************************************************************
Dim Reflection As String = System.Reflection.Assembly.GetExecutingAssembly().ToString
Dim ProductName As String = Reflection.Substring(0, Reflection.IndexOf(","))
Dim logEvent As String = UCase(ProductName) & " FINISHED" & vbCrLf
Dim noErrorFlag = True 'return value for the function, returns false if an error occurs

'Append date/time to Event Description and write to log file
noErrorFlag = addEvent(logEvent)

'Cleanup
    logEvent = Nothing
    ProductName = Nothing
    Reflection = Nothing

'Exit
    Return noErrorFlag

 End Function

#End Region
#Region "ClearOld"

'*************************************************************************
 Public Function clearOld( _
                        Optional ByVal daysToKeep As Integer = 365 _
                        ) As Boolean
'PURPOSE:
'   Remove log files that have exceeded the daysToKeep retention period.
'
'OUTPUT:
'   Adds an Event Description with deletion summary, returns FALSE if an error occured.
'*************************************************************************
Dim logEvent As String = "Log.clearOld: Deleted "
Dim noErrorFlag = True 'return value for the function, returns false if an error occurs
Dim LogFolderPath, LogFileName As String
Dim objFiles() As String
Dim enumerator As System.Collections.IEnumerator
Dim DT As DateTime
Dim Counts As String()
Dim str As String
Dim i As Integer = 0

'Get the log folder path and loop through all the files to remove the old ones.
    If Not LogFile = "" Then
        Pnm.File.LogOnOff = False
        Pnm.Directory.LogOnOff = False
        LogFolderPath = GetLogFolderPath(LogFile)
        objFiles = Pnm.Directory.GetFileNames(LogFolderPath)
        enumerator = objFiles.GetEnumerator
        Counts = DoTheWork(enumerator, daysToKeep, DT, LogFolderPath, logEvent)
        logEvent = "Success: " & logEvent & Counts(2) & " log files of " & _
        Counts(1) & " old logs. Total logs = " & Counts(3) - Counts(2)
        addEvent(logEvent)
        Pnm.Directory.LogOnOff = True
        Pnm.File.LogOnOff = True
    Else
        MsgBox("Unable to write to the log because its not open. Start logging before clearing logs so the action can be documented.", MsgBoxStyle.Critical, "Cannot Clear Logs")
    End If

'Cleanup
    i = Nothing
    str = Nothing
    Counts = Nothing
    DT = Nothing
    enumerator = Nothing
    objFiles = Nothing
    LogFolderPath = Nothing
    LogFileName = Nothing
    logEvent = Nothing
    daysToKeep = Nothing

'Exit
    Return noErrorFlag

 End Function

#End Region
#Region "GetLogFolderPath"

'*************************************************************************
Private Function GetLogFolderPath(ByVal FilePath As String) As String
'PURPOSE:
'   Get the log folder Path
'
'OUTPUT:
'   The Log Folder Path.
'*************************************************************************
Dim Returns As String
Dim Reflection As String = System.Reflection.Assembly.GetExecutingAssembly().ToString

'Get the Log Folder Path
    If Not FilePath.StartsWith("Logs") Then
        Returns = Pnm.Directory.GetParent(FilePath)
    Else
        Returns = Pnm.Directory.GetCurrentDirectory() & "\logs"
    End If

'Cleanup
    FilePath = Nothing
    Reflection = Nothing

'Exit
    Return Returns

End Function

#End Region
#Region "DoTheWork"

'*************************************************************************
Private Function DoTheWork(ByVal enumerator As System.Collections.IEnumerator, _
                    ByVal DaysToKeep As Integer, _
                    ByVal DT As DateTime, _
                    ByVal Filepath As String, _
                    ByVal LogEvent As String _
                    ) As String()
'PURPOSE:
'   Get the log folder Path
'
'OUTPUT:
'   The Log Folder Path.
'*************************************************************************
Dim countOld As Integer
Dim countDeleted As Integer
Dim countTotal As Integer

'for each file Get the last write date If the file is older than the # days to keep then delete it
    While enumerator.MoveNext
        countTotal = countTotal + 1
        DT = Pnm.File.GetLastWriteTime(Filepath & enumerator.Current)
        If DateTime.Now.Subtract(DT).TotalDays > DaysToKeep Then
            Try
                countOld = countOld + 1
                Pnm.File.Delete(enumerator.Current)
                countDeleted = countDeleted + 1
            Catch ex As Exception
                addEvent(LogEvent & "Failed to delete " & enumerator.Current)
            End Try
        End If
    End While

'Cleanup
    LogEvent = Nothing
    Filepath = Nothing
    DT = Nothing
    DaysToKeep = Nothing
    enumerator = Nothing

'Exit
    Return AddCounts(countOld, countDeleted, countTotal)

End Function

#End Region
#Region "AddCounts"

'*************************************************************************
Private Function AddCounts(ByVal countold As Integer, _
                    ByVal countDeleted As Integer, _
                    ByVal CountTotal As Integer _
                    ) As String()
'PURPOSE:
'   Add All of out counts to a String Array for later use.
'
'OUTPUT:
'   The string array.
'*************************************************************************
Dim Returns As String()

'Add all of the counts to the string array.
ReDim Returns(3)
    Returns(1) = countold.ToString
    Returns(2) = countDeleted.ToString
    Returns(3) = CountTotal.ToString

'Cleanup
    countold = Nothing
    countDeleted = Nothing
    CountTotal = Nothing

'Exit
    Return Returns

End Function

#End Region
#Region "MyException"

Private Exec As String
ReadOnly Property MyException() As String
Get
    Return Exec
End Get
End Property

#End Region
End Class