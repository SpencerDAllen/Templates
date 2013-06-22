#Region "Comments"

'**********************************************************************************************
'PURPOSE:
'   clsLogFile is a class used to produce useful log files.
'
'PREREQUISITES:
'   A form must be added to the project, an empty one if no forms are needed for the project.
'   clsLogFile must be added to the project.
'   "Public LogFilePath As String" must be in the main program's global declarations.
'   "Public Log As New clsLogFile" must be in the main program's global declarations.
'
'USAGE:
'   boolean = Logfile.start(FilePath As String, [NewLogFrequency As String], [LogFolder As String])     Options = Always, Daily, Weekly, Monthly
'   boolean = Logfile.addEvent(FilePath As String, EventDescription As String)
'   boolean = LogFile.finish(FilePath As String)       
'   Logfile.clearOld(FilePath, daysToKeep As Integer)                                                   Number of days to keep unused logs
'
'AUTHOR:
'   Shayne Marriage
'   BTS Applications
'   04-20-2007
'
' VERSIONS:
'   1.00 - Base Version (Shayne Marriage)
'   1.01 - Changed log entry for clearOld member to be like other classes
'   1.02 - Adds LogFolder as an optional parameter to change log folder location. Changes clearOld member to accomodate the change
'**********************************************************************************************

#End Region
#Region "Declorations"

Imports System
Imports System.IO

#End Region
Public Class clsLogFile
#Region "Start"

'*************************************************************************
	Public Function start( _
						   ByRef FilePath As String, _
						   Optional ByVal NewLogFrequency As String = "Always", _
						   Optional ByVal LogFolder As String = "Logs" _
						 ) As Boolean
'PURPOSE:
'   To start logging by creating a new log file or re-opening an existing one.
'
'OUTPUT:
'   Passes the log FilePath parameter out, adds log initiated event, and returns FALSE if an error occured.
'*************************************************************************
		Dim logEvent As String = UCase(Main.Product) & " STARTED"
		Dim noErrorFlag = True 'return value for the function, returns false if an error occurs
		Dim DT As DateTime = DateTime.Now  'Date and Time now
		Dim strMMin As String = Right("0" & CStr(DT.Day), 2)
		Dim strHH As String = Right("0" & CStr(DT.Hour), 2)
		Dim strDD As String = Right("0" & CStr(DT.Day), 2)
		Dim strMM As String = Right("0" & CStr(DT.Month), 2)
		Dim strYYYY As String = Right("0" & CStr(DT.Year), 4)
		Dim TrimChars() As Char = {"\", " "}

		'Remove leading / trailing whitespace and trailing "\" if present
		LogFolder = LogFolder.TrimEnd(TrimChars)
		LogFolder = LogFolder.Trim

		'Ensure the log directory exists or can be created
		If Not Directory.Exists(LogFolder) Then
			Try
				Directory.CreateDirectory(LogFolder)
			Catch ex As Exception
				MsgBox("Unable to create Log Directory." & vbCrLf & ex.ToString, MsgBoxStyle.Critical, "Fatal Error")
				noErrorFlag = False
			End Try
		End If

		If noErrorFlag Then
			'Generate Log name based on new log file frequency
			FilePath = LogFolder & "\" & Main.Product
			Select Case NewLogFrequency
				Case "Daily"
					FilePath = FilePath & "_" & strYYYY & "-" & strMM & "-" & "Day-" & strDD
				Case "Weekly"
					FilePath = FilePath & "_" & strYYYY & "-" & strMM
					FilePath = FilePath & "_" & "Week-" & Left(CStr((strDD - 1) / 7) + 1, 1)
				Case "Monthly"
					FilePath = FilePath & "_" & strYYYY & "-" & "Month-" & strMM
				Case Else   'Always
					FilePath = FilePath & "_" & strYYYY & "-" & strMM & "-" & strDD
					FilePath = FilePath & "_" & strHH & "." & strMMin
			End Select
			FilePath = FilePath & ".txt"
			noErrorFlag = clsLogFile.addEvent(FilePath, logEvent)
		End If

		Return noErrorFlag

	End Function

#End Region
#Region "AddEvent"

'*************************************************************************
	Public Shared Function addEvent( _
									 ByVal FilePath As String, _
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

		Try
			Dim fWriter As New StreamWriter(FilePath, True)

			'Append date/time to Event Description and write to log file
			fWriter.WriteLine(logEvent)
			fWriter.Close()
			noErrorFlag = True
		Catch ex As Exception
			MsgBox("Log File cannot be opened." & vbCrLf & ex.ToString, MsgBoxStyle.Critical, "Fatal Error")
			noErrorFlag = False
		End Try

		Return noErrorFlag

	End Function

#End Region
#Region "Finish"

'*************************************************************************
	Public Function finish( _
							ByVal FilePath As String _
						  ) As Boolean
'PURPOSE:
'   To end logging.
'
'OUTPUT:
'   Adds log terminated event, returns FALSE if an error occured.
'*************************************************************************

		Dim logEvent As String = UCase(Main.Product) & " FINISHED" & vbCrLf
		Dim noErrorFlag = True 'return value for the function, returns false if an error occurs

		'Append date/time to Event Description and write to log file
		noErrorFlag = clsLogFile.addEvent(FilePath, logEvent)

		Return noErrorFlag

	End Function

#End Region
#Region "ClearOld"

'*************************************************************************
	Public Function clearOld( _
							  ByVal FilePath As String, _
							  ByVal daysToKeep As Integer _
							) As Boolean
'PURPOSE:
'   Remove log files that have exceeded the daysToKeep retention period.
'
'OUTPUT:
'   Adds an Event Description with deletion summary, returns FALSE if an error occured.
'*************************************************************************

		Dim logEvent As String = "Log.clearOld: Deleted Log Files: "
		Dim noErrorFlag = True 'return value for the function, returns false if an error occurs
		Dim arrFilePath() As String
		Dim LogFolderPath, LogFileName As String
		Dim objFiles() As String
		Dim enumerator As System.Collections.IEnumerator
		Dim newestDate As Date
		Dim newestFile As String
		Dim DT As DateTime
		Dim countOld, countDeleted, countTotal As Integer

		If Not FilePath = "" Then
			countOld = 0
			countDeleted = 0
			countTotal = 0

			'Get LogFolderPath
			arrFilePath = FilePath.Split("\")
			LogFileName = arrFilePath(arrFilePath.Length - 1)
			LogFolderPath = FilePath.Replace("\" & LogFileName, "")

			'Enumerate log folder
			objFiles = System.IO.Directory.GetFiles(LogFolderPath)
			enumerator = objFiles.GetEnumerator

			'for each file
			While enumerator.MoveNext
				countTotal = countTotal + 1
				'Get the last write date
				DT = File.GetLastWriteTime(enumerator.Current)
				'If the file is older than the # days to keep then delete it
				If DateTime.Now.Subtract(DT).TotalDays > daysToKeep Then
					Try
						countOld = countOld + 1
						File.Delete(enumerator.Current)
						countDeleted = countDeleted + 1
					Catch ex As Exception
						clsLogFile.addEvent(FilePath, logEvent & "Failed to delete " & enumerator.Current)
					End Try
				End If
			End While

			logEvent = "Success: " & logEvent & countDeleted & " logs of " & _
						countOld & " old logs. Total logs = " & countTotal - countDeleted
			clsLogFile.addEvent(FilePath, logEvent)
		Else
			MsgBox("Unable to write to the log because its not open. Start logging before clearing logs so the action can be documented.", MsgBoxStyle.Critical, "Cannot Clear Logs")
		End If

		Return noErrorFlag
	End Function

#End Region
End Class
