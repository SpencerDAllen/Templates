#Region "Comments"

'*****************************************************************************
'PURPOSE:
'   A Process component provides access to a process that is running on a 
'   computer. A process, in the simplest terms, is a running application. 
'   A thread is the basic unit to which the operating system allocates processor 
'   time. A thread can execute any part of the code of the process, including 
'   parts currently being executed by another thread.
'
'   The Process component is a useful tool for starting, stopping, controlling, 
'   and monitoring applications. Using the Process component, you can obtain a 
'   list of the processes that are running or start a new process. A Process 
'   component is used to access system processes. After a Process component has 
'   been initialized, it can be used to obtain information about the running process. 
'   Such information includes the set of threads, the loaded modules 
'   (.dll and .exe files), and performance information such as the amount of 
'   memory the process is using.
'
'   If you have a path variable declared in your system using quotes, you must 
'   fully qualify that path when starting any process found in that location. 
'   Otherwise, the system will not find the path. For example, if c:\mypath is 
'   not in your path, and you add it using quotation marks: 
'   path = %path%;"c:\mypath", you must fully qualify any process in c:\mypath 
'   when starting it.
'
'   The process component obtains information about a group of properties all 
'   at once. After the Process component has obtained information about one 
'   member of any group, it will cache the values for the other properties in 
'   that group and not obtain new information about the other members of the 
'   group until you call the Refresh method. Therefore, a property value is not 
'   guaranteed to be any newer than the last call to the Refresh method. The group 
'   breakdowns are operating-system dependent.
'
'   A system process is uniquely identified on the system by its process identifier. 
'   Like many Windows resources, a process is also identified by its handle, which 
'   might not be unique on the computer. A handle is the generic term for an 
'   identifier of a resource. The operating system persists the process handle, 
'   which is accessed through the Handle property of the Process component, even 
'   when the process has exited. Thus, you can get the process's administrative 
'   information, such as the ExitCode (usually either zero for success or a nonzero 
'   error code) and the ExitTime. Handles are an extremely valuable resource, so 
'   leaking handles is more virulent than leaking memory.
'
'USAGE:
'   Integer = StartAndWait(Command as string, WindowStyle as string)
'
'AUTHOR:
'   Spencer Allen
'   11-07-2007
'
'USAGE:
'   
'
'VERSIONS:
'   1.00 - Base Version (Spencer Allen)
'*****************************************************************************

#End Region
Public Class PwrClsProcess
#Region "Declorations"
Implements IDisposable
    Enum Style
        HIDDEN
        MAXIMIZED
        MINIMIZED
        NORMAL
    End Enum

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
    Except = Nothing
    LogsOnOff = Nothing

'Exit
End Sub

#End Region
#Region "StartAndWait"

'*************************************************************************
    Public Function StartandWait( _
                            ByVal Command As String, _
                            ByVal WindowStyle As Style _
                           ) As Integer
'PURPOSE:
'   To seperate .exe calls and arguments and makes the process run. Sets
'   off TaskEnded Event if ended prematurly.
'
'RETURNS:
'   A started process or a failure event.
'*************************************************************************
Dim myProcess As New System.Diagnostics.Process
Dim logEvent As String = "Process.StartandWait: "
Dim ProcessExitCode As Integer
Dim Exe As String
Dim Args As String
Except = ""

'Check for correct string
    If Not CheckForExe(Command, logEvent) Then
        Except = "The command conatained no valid executable."
        If LogsOnOff Then
            Pnm.Logs.addEvent("Failure: " & logEvent & Exe & " " & Args & vbCrLf & Except)
        End If
        Exit Function
    End If

' Seperate Executable from Arguments
    Exe = Command.Substring(0, Command.IndexOf(".exe") + 4)
    If Command.Length > Command.IndexOf(".exe") + 4 Then
        Args = Command.Substring(Exe.Length)
    End If

'Initialize the process object
'Start the process and wait for it to exit
    Try
        myProcess = SetStartInfoPropertys(myProcess, Exe, Args, WindowStyle)
        myProcess.Start()
        ProcessExitCode = Sleep(myProcess)
    Catch ex As Exception
        Except = ex.Message
        ProcessExitCode = -2
    End Try

'Log if logging
    If LogsOnOff Then
        If Not Except = "" Then
            Pnm.Logs.addEvent("Failure: " & logEvent & Exe & " " & Args & vbCrLf & Except)
        Else
            Pnm.Logs.addEvent("Success: " & logEvent & Exe & " " & Args)
        End If
    End If

'Cleanup
    Args = Nothing
    Exe = Nothing
    logEvent = Nothing
    myProcess = Nothing
    WindowStyle = Nothing
    Command = Nothing

'Exit
    Return ProcessExitCode

End Function

#End Region
#Region "CheckForExe"

'*************************************************************************
Private Function CheckForExe(ByVal Command As String, _
                            ByVal Logevent As String _
                            ) As Boolean
'PURPOSE:
'   To check for a valid command
'
'RETURNS:
'   ture if command is valid.
'*************************************************************************
Dim Returned As Boolean = True

    'Check if the command is to an exe
    If Command.IndexOf(".exe") < 1 Then
        Returned = False
        If LogsOnOff Then
            Pnm.Logs.addEvent(Logevent & "Command does not call an executeable: " & Command)
        End If
    End If

'Cleanup
    Logevent = Nothing
    Command = Nothing

'Exit
    Return Returned

End Function

#End Region
#Region "SetStartInfoPropertys"

'*************************************************************************
Private Function SetStartInfoPropertys(ByVal MyProcess As System.Diagnostics.Process, _
                                ByVal Exe As String, _
                                ByVal Args As String, _
                                ByVal WindowStyle As Style _
                                ) As System.Diagnostics.Process
'PURPOSE:
'   To set the propertys of the myprocess object.
'
'RETURNS:
'   A myprocess object with it's propertys set.
'*************************************************************************

    With MyProcess
        .StartInfo.FileName = Exe
        .StartInfo.Arguments = Args
        Select Case WindowStyle.ToString
            Case "HIDDEN"
                .StartInfo.WindowStyle = Diagnostics.ProcessWindowStyle.Hidden
                .StartInfo.CreateNoWindow = True
            Case "MAXIMIZED"
                .StartInfo.WindowStyle = Diagnostics.ProcessWindowStyle.Maximized
            Case "MINIMIZED"
                .StartInfo.WindowStyle = Diagnostics.ProcessWindowStyle.Minimized
            Case "NORMAL"
                .StartInfo.WindowStyle = Diagnostics.ProcessWindowStyle.Normal
        End Select
    End With

'Cleanup
    WindowStyle = Nothing
    Args = Nothing
    Exe = Nothing

'Exit
    Return MyProcess

End Function

#End Region
#Region "Sleep"

'*************************************************************************
Private Function Sleep(ByVal myProcess As System.Diagnostics.Process _
                        ) As Integer
'PURPOSE:
'   To allow the application run without freezing up or taking to many system
'   resources.
'
'RETURNS:
'   The application waits for the process to finish.
'*************************************************************************
Dim Returned As Integer

    Do
        System.Windows.Forms.Application.DoEvents()
        System.Threading.Thread.Sleep(250)
    Loop While Not myProcess.HasExited()
    Returned = myProcess.ExitCode
    myProcess.Close()

'Cleanup
    myProcess = Nothing

'Exit
    Return Returned

End Function

#End Region
#Region "MyException Property"

Private Except As String

'*************************************************************************
ReadOnly Property MyException() As String
'PURPOSE:
'   To make a Property to return the exceptions generated by this class.
'
'OUTPUT:
'   Function returns The exception string or a null string if no exception
'   was raised.
'*************************************************************************
Get
    Return Except
End Get
End Property

#End Region
#Region "LogsOnOFF Property"

Private LogsOnOff As Boolean

'*************************************************************************
WriteOnly Property LogOnOff() As Boolean
'PURPOSE:
'   To Turn logging on or off for the class.
'
'OUTPUT:
'   Logging is turned on or off.
'*************************************************************************
Set(ByVal Value As Boolean)
    LogsOnOff = Value
End Set
End Property
#End Region
End Class