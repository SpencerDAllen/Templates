#Region "Comments"

'*****************************************************************************'
'PURPOSE:
'   Use the Directory class for typical operations such as copying, moving, renaming,
'   creating, and deleting directories. You can also use the Directory class to get and 
'   set DateTime information related to the creation, access, and writing of a directory.
'
'   Because all Directory methods are static, it might be more efficient to use a File 
'   method rather than a corresponding DirectoryInfo instance method if you want to perform 
'   only one action. Most Directory methods require the path to the directory that you are manipulating.
'
'   The static methods of the Directory class perform security checks on all methods. 
'   If you are going to reuse an object several times, consider using the corresponding 
'   instance method of DirectoryInfo instead, because the security check will not always be necessary.
'
'   In members that accept a path as an input string, that path must be well-formed or an 
'   exception is raised. For example, if a path is fully qualified but begins with a space, 
'   the path is not trimmed in methods of the class. Therefore, the path is malformed and an 
'   exception is raised. Similarly, a path or a combination of paths cannot be fully qualified 
'   twice. For example, "c:\temp c:\windows" also raises an exception in most cases. Ensure 
'   that your paths are well-formed when using methods that accept a path string.
'
'   In members that accept a path, the path can refer to a file or just a directory. The 
'   specified path can also refer to a relative path or a Universal Naming Convention (UNC) 
'   path for a server and share name. For example, all the following are acceptable paths: 
'
'   "c:\MyDir"
'   "MyDir\MySubDir" 
'   "\\MyServer\MyShare"
'
'   By default, full read/write access to new directories is granted to all users.
'
'PREREQUISITES:
'   .Net 1.1
'   BaseClsDirectory.vb must be added to the project
'
'USAGE:
'   boolean = Exists(Path as string)
'   boolean = Move(SourcePath as string, DestinationPath as string)
'   boolean = Delete(Path as string)
'   boolean = Create(Path as string)
'   String() = GetSubDirectorys(Path as string)
'   String() = GetFileNames(Path as string)
'   String() = GetLogicalDrives()
'   String = GetParent(Path As String)
'   String = GetDirectoryRoot(Path As String)
'   String = GetCurrentDirectory()
'   String() = GetFileSystemEntries(Path As String)
'   String = GetLastWriteTime(Path As String)
'   String = GetLastWriteTimeUtc(Path As String)
'   String = GetCreationTime(Path As String)
'   String = GetCreationTimeUtc(Path As String)
'   String = GetLastAccessTime(Path As String)
'   String = GetLastAccessTimeUtc(Path As String)
'   boolean = SetLastWriteTime(Path As String, LastAccessTime As String)
'   boolean = SetLastWriteTimeUtc(Path As String, LastAccessTime As String)
'   boolean = SetCreationTime(Path As String, CreationTime As String)
'   boolean = SetCreationTimeUtc(Path As String, CreationTime As String)
'   boolean = SetLastAccessTime(Path As String, LastAccessTime As String)
'   boolean = SetLastAccessTimeUtc(Path As String, LastAccessTime As String)
'
'
'MY CLASS MODS:
'
'AUTHOR:
'   Spencer Allen
'   5/15/08
'
'VERSIONS:
'   1.00 - Base Version (Spencer Allen)
'*****************************************************************************'

#End Region
Public Class MyClsDirectory
#Region "Inherits"

Inherits BaseClsDirectory
Implements IDisposable

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
    LogsOnOff = Nothing

'Exit
End Sub

#End Region
#Region "Exists"

'*************************************************************************
Public Overloads Function Exists(ByVal Path As String _
                                ) As Boolean
'PURPOSE:
'   Find out if a Directory exists.
'
'OUTPUT:
'   Function returns boolean.
'*************************************************************************
Dim LogEvent As String = "Directory.Exists : " & Path
Dim Returned As Boolean = MyBase.Exists(Path)

'Checks to see if the file exists.
    If LogsOnOff Then
        If Returned Then
            LogEvent = "Success: " & LogEvent & " Exists."
            Logs(LogEvent)
        Else
            LogEvent = "Failure: " & LogEvent & " Did not Exist."
            Logs(LogEvent)
        End If
    End If

'Cleanup
    LogEvent = Nothing
    Path = Nothing

'Exit
    Return Returned

End Function

#End Region
#Region "Move"

'*************************************************************************
Public Overloads Function Move(ByVal SourcePath As String, _
                            ByVal DestinationPath As String _
                            ) As Boolean
'PURPOSE:
'   Move the Directory from one location to another.
'
'OUTPUT:
'   Function returns The string "True" if suscessful or ": exception message" 
'   if an exception was thrown.
'*************************************************************************
Dim LogEvent As String = "Directory.Move " & SourcePath & " " & DestinationPath
Dim DestParent As String = GetParent(DestinationPath)

'Checking to see if the parent of the destination directory exists before moving
'a directory into it, If it dosn't exist we are creating it before moving.
    If Not Exists(DestParent) Then
        Create(DestParent)
    End If
    Dim Returns As Boolean = MyBase.Move(SourcePath, DestinationPath)

    If LogsOnOff Then
        If Not MyException = "" Then
            LogEvent = "Failure: " & LogEvent & vbCrLf & MyException
            Logs(LogEvent)
        Else
            LogEvent = "Success: " & LogEvent
            Logs(LogEvent)
        End If
    End If

'Cleanup
    DestParent = Nothing
    DestinationPath = Nothing
    SourcePath = Nothing

'Exit
    Return Returns

End Function

#End Region
#Region "Delete"

'*************************************************************************
Public Overloads Function Delete(ByVal Path As String _
                                ) As Boolean
'PURPOSE:
'   To delete a Directory.
'
'OUTPUT:
'   Function returns The string "True" if suscessful or ": exception message" 
'   if an exception was thrown.
'*************************************************************************
Dim LogEvent As String = "Directory.delete" & Path
Dim Returned As Boolean = MyBase.Delete(Path)

'Check for execption before logging
    If LogsOnOff Then
        If Not MyException = "" Then
            LogEvent = "Failure: " & LogEvent & vbCrLf & MyException
            Logs(LogEvent)
        Else
            LogEvent = "Success: " & LogEvent
            Logs(LogEvent)
        End If
    End If

'Cleanup
    LogEvent = Nothing
    Path = Nothing

'Exit
    Return Returned

End Function

#End Region
#Region "Create"

'*************************************************************************
Public Overloads Function Create(ByVal Path As String _
                                ) As Boolean
'PURPOSE:
'   To Create a Directory or Directory Structure.
'
'OUTPUT:
'   Function returns The string "True" if suscessful or ": exception message" 
'   if an exception was thrown.
'*************************************************************************
Dim LogEvent As String = "Directory.Create" & Path
Dim Returned As Boolean = MyBase.Create(Path)

'Check for execption before logging
    If LogsOnOff Then
        If Not MyException = "" Then
            LogEvent = "Failure: " & LogEvent & vbCrLf & MyException
            Logs(LogEvent)
        Else
            LogEvent = "Success: " & LogEvent
            Logs(LogEvent)
        End If
    End If

'Cleanup
    LogEvent = Nothing
    Path = Nothing

'Exit
    Return Returned

End Function

#End Region
#Region "GetSubDirectorys"

'*************************************************************************
Public Overloads Function GetSubDirectorys(ByVal Path As String _
                                            ) As String()
'PURPOSE:
'   To get a list of all sub directorys.
'
'OUTPUT:
'   Function returns an array of Sub Directory Names or a ": " exception message.
'*************************************************************************
Dim LogEvent As String = "Directory.GetSubDirectorys: " & Path
Dim FileNames As String()
Dim Names As String
Dim i As Integer = 0

'Clean the results from getfilenames to return only file names.
'Not full paths.
    If Not MyBase.GetSubDirectorys(Path) Is Nothing Then
        ReDim FileNames(UBound(MyBase.GetSubDirectorys(Path)))
        For Each Names In MyBase.GetSubDirectorys(Path)
            FileNames(i) = Names.Substring(Path.Length)
            i = i + 1
        Next
    End If

'Log the event if logging is true
    If LogsOnOff Then
        If Not MyException = "" Then
            LogEvent = "Failure: " & LogEvent & vbCrLf & MyException
            Logs(LogEvent)
        Else
            LogEvent = "Success: " & LogEvent
            Logs(LogEvent)
        End If
    End If

'Cleanup
    Names = Nothing
    Path = Nothing
    i = Nothing
    LogEvent = Nothing

'Exit
    Return FileNames

End Function

#End Region
#Region "GetFileNames"

'*************************************************************************
Public Overloads Function GetFileNames(ByVal Path As String _
                                        ) As String()
'PURPOSE:
'   To get a list of all the files in the directory.
'
'OUTPUT:
'   Function returns an array of file Names or a ": " exception message.
'*************************************************************************
Dim LogEvent As String = "Directory.GetFileNames: " & Path
Dim FileNames As String()
Dim Names As String
Dim i As Integer = 0

'Clean the results from getfilenames to return only file names.
'Not full paths.
    If Not MyBase.GetFileNames(Path) Is Nothing Then
        ReDim FileNames(UBound(MyBase.GetFileNames(Path)))
        For Each Names In MyBase.GetFileNames(Path)
            FileNames(i) = Names.Substring(Path.Length)
            i = i + 1
        Next
    End If

'Log the event if logging is true
    If LogsOnOff Then
        If Not MyException = "" Then
            LogEvent = "Failure: " & LogEvent & vbCrLf & MyException
            Logs(LogEvent)
        Else
            For Each Names In FileNames
                LogEvent = "Success: " & LogEvent & " Returned " & Names
                Logs(LogEvent)
            Next
        End If
    End If

'Cleanup
    Names = Nothing
    Path = Nothing
    i = Nothing
    LogEvent = Nothing

'Exit
    Return FileNames

End Function

#End Region
#Region "GetLogicalDrives"

'*************************************************************************
Public Overloads Function GetLogicalDrives( _
                                            ) As String()
'PURPOSE:
'   To get names of all logical drives on the current computer.
'
'OUTPUT:
'   Function returns names of all logical drives on the current computer or a ": " exception message.
'*************************************************************************
Dim LogEvent As String = "Directory.GetLogicalDrives: "
Dim FileNames As String()
Dim Names As String
Dim i As Integer = 0

'Clean the results from getfilenames to return only file names.
'Not full paths.
    If Not MyBase.GetLogicalDrives() Is Nothing Then
        ReDim FileNames(UBound(MyBase.GetLogicalDrives()))
        For Each Names In MyBase.GetLogicalDrives()
            FileNames(i) = Names
        Next
    End If

'Log the event if logging is true
    If LogsOnOff Then
        If Not MyException = "" Then
            LogEvent = "Failure: " & LogEvent & vbCrLf & MyException
            Logs(LogEvent)
        Else
            LogEvent = "Success: " & LogEvent
            Logs(LogEvent)
        End If
    End If

'Cleanup
    Names = Nothing
    i = Nothing
    LogEvent = Nothing

'Exit
    Return FileNames

End Function

#End Region
#Region "GetParent"

'*************************************************************************
Public Overloads Function GetParent(ByVal Path As String _
                                    ) As String
'PURPOSE:
'   To get names of all logical drives on the current computer.
'
'OUTPUT:
'   Function returns names of all logical drives on the current computer or a ": " exception message.
'*************************************************************************
Dim LogEvent As String = "Directory.GetParent" & Path
Dim Returned As String = MyBase.GetParent(Path)

'Check for execption before logging
    If LogsOnOff Then
        If Not MyException = "" Then
            LogEvent = "Failure: " & LogEvent & vbCrLf & MyException
            Logs(LogEvent)
        Else
            LogEvent = "Success: " & LogEvent
            Logs(LogEvent)
        End If
    End If

'Cleanup
    LogEvent = Nothing
    Path = Nothing

'Exit
    Return Returned

End Function

#End Region
#Region "GetDirectoryRoot"

'*************************************************************************
Public Overloads Function GetDirectoryRoot(ByVal Path As String _
                                            ) As String
'PURPOSE:
'   To get volume and root information for a specified path.
'
'OUTPUT:
'   Function returns volume and root information for a specified path or a ": " exception message.
'*************************************************************************
Dim LogEvent As String = "Directory.GetDirectoryRoot" & Path
Dim Returned As String = MyBase.GetDirectoryRoot(Path)

'Check for execption before logging
    If LogsOnOff Then
        If Not MyException = "" Then
            LogEvent = "Failure: " & LogEvent & vbCrLf & MyException
            Logs(LogEvent)
        Else
            LogEvent = "Success: " & LogEvent
            Logs(LogEvent)
        End If
    End If

'Cleanup
    LogEvent = Nothing
    Path = Nothing

'Exit
    Return Returned

End Function

#End Region
#Region "GetCurrentDirectory"

'*************************************************************************
Public Overloads Function GetCurrentDirectory( _
                                            ) As String
'PURPOSE:
'   To get the current working directory of the application.
'
'OUTPUT:
'   Function returns the current working directory of the application as a 
'   string or the standard ": " exception message.
'*************************************************************************
Dim LogEvent As String = "Directory.GetCurrentDirectory"
Dim Returned As String = MyBase.GetCurrentDirectory()

'Check for execption before logging
    If LogsOnOff Then
        If Not MyException = "" Then
            LogEvent = "Failure: " & LogEvent & vbCrLf & MyException
            Logs(LogEvent)
        Else
            LogEvent = "Success: " & LogEvent
            Logs(LogEvent)
        End If
    End If

'Cleanup
    LogEvent = Nothing

'Exit
    Return Returned

End Function

#End Region
#Region "GetFileSystemEntries"

'*************************************************************************
Public Overloads Function GetFileSystemEntries(ByVal Path As String _
                                                ) As String()
'PURPOSE:
'   To get names of all files and subdirectories in the specified directory.
'
'OUTPUT:
'   Function returns names of files and subdirectories in the specified directory or a ": " exception message.
'*************************************************************************
Dim LogEvent As String = "Directory.GetFileSystemEntries: " & Path
Dim FileNames As String()
Dim Names As String
Dim i As Integer = 0

'Clean the results from getfilenames to return only file names.
'Not full paths.
    If Not MyBase.GetFileSystemEntries(Path) Is Nothing Then
        ReDim FileNames(UBound(MyBase.GetFileSystemEntries(Path)))
        For Each Names In MyBase.GetFileSystemEntries(Path)
            FileNames(i) = Names.Substring(Path.Length)
            i = i + 1
        Next
    End If

'Log the event if logging is true
    If LogsOnOff Then
        If Not MyException = "" Then
            LogEvent = "Failure: " & LogEvent & vbCrLf & MyException
            Logs(LogEvent)
        Else
            LogEvent = "Success: " & LogEvent
            Logs(LogEvent)
        End If
    End If

'Cleanup
    Names = Nothing
    Path = Nothing
    i = Nothing
    LogEvent = Nothing

'Exit
    Return FileNames

End Function

#End Region
#Region "GetAttributes"
#Region "GetLastWriteTime"

'*************************************************************************
Public Overloads Function GetLastWriteTime(ByVal Path As String _
                                            ) As String
'PURPOSE:
'   To return the date and time of the specified file or directory.
'
'
'OUTPUT:
'   Function returns the date and time the specified file or directory was last
'   written to or ": exception message" if an exception was thrown.
'*************************************************************************
Dim LogEvent As String = "Directory.GetLastWriteTime" & Path
Dim Returned As String = MyBase.GetLastWriteTime(Path)

'Check for execption before logging
    If LogsOnOff Then
        If Not MyException = "" Then
            LogEvent = "Failure: " & LogEvent & vbCrLf & MyException
            Logs(LogEvent)
        Else
            LogEvent = "Success: " & LogEvent
            Logs(LogEvent)
        End If
    End If

'Cleanup
    LogEvent = Nothing
    Path = Nothing

'Exit
    Return Returned

End Function

#End Region
#Region "GetLastWriteTimeUtc"

'*************************************************************************
Public Overloads Function GetLastWriteTimeUtc(ByVal Path As String _
                                                ) As String
'PURPOSE:
'   To get the last time the specified file was written to. This value is expressed in UTC time.

'
'OUTPUT:
'   Function returns a datetime object or a object with the standard ": " exception message.
'*************************************************************************
Dim LogEvent As String = "Directory.GetLastWriteTimeUtc" & Path
Dim Returned As String = MyBase.GetLastWriteTimeUtc(Path)

'Check for execption before logging
    If LogsOnOff Then
        If Not MyException = "" Then
            LogEvent = "Failure: " & LogEvent & vbCrLf & MyException
            Logs(LogEvent)
        Else
            LogEvent = "Success: " & LogEvent
            Logs(LogEvent)
        End If
    End If

'Cleanup
    LogEvent = Nothing
    Path = Nothing

'Exit
    Return Returned

End Function

#End Region
#Region "GetCreationTime"

'*************************************************************************
Public Overloads Function GetCreationTime(ByVal Path As String _
                                            ) As String
'PURPOSE:
'   To get the creation date and time of a directory.
'
'OUTPUT:
'   Function returns a DateTime object or an object with the standard ": " 
'   exception message.
'*************************************************************************
Dim LogEvent As String = "Directory.GetCreationTime" & Path
Dim Returned As String = MyBase.GetCreationTime(Path)

'Check for execption before logging
    If LogsOnOff Then
        If Not MyException = "" Then
            LogEvent = "Failure: " & LogEvent & vbCrLf & MyException
            Logs(LogEvent)
        Else
            LogEvent = "Success: " & LogEvent
            Logs(LogEvent)
        End If
    End If

'Cleanup
    LogEvent = Nothing
    Path = Nothing

'Exit
    Return Returned

End Function

#End Region
#Region "GetCreationTimeUtc"

'*************************************************************************
Public Overloads Function GetCreationTimeUtc(ByVal Path As String _
                                            ) As String
'PURPOSE:
'   To get the creation date and time of a directory in coordinated universal time (UTC) format.
'
'OUTPUT:
'   Function returns a DateTime object expressed in UTC time or an object with the standard ": " exception message.
'*************************************************************************
Dim LogEvent As String = "Directory.GetCreationTimeUtc" & Path
Dim Returned As String = MyBase.GetCreationTimeUtc(Path)

'Check for execption before logging
    If LogsOnOff Then
        If Not MyException = "" Then
            LogEvent = "Failure: " & LogEvent & vbCrLf & MyException
            Logs(LogEvent)
        Else
            LogEvent = "Success: " & LogEvent
            Logs(LogEvent)
        End If
    End If

'Cleanup
    LogEvent = Nothing
    Path = Nothing

'Exit
    Return Returned

End Function

#End Region
#Region "GetLastAccessTime"

'*************************************************************************
Public Overloads Function GetLastAccessTime(ByVal Path As String _
                                            ) As String
'PURPOSE:
'   To get the date and time the specified file or directory was last accessed.
'
'OUTPUT:
'   Function returns date and time the specified file or directory was last accessed or an object with the standard ": " exception message.
'*************************************************************************
Dim LogEvent As String = "Directory.GetLastAccessTime" & Path
Dim Returned As String = MyBase.GetLastAccessTime(Path)

'Check for execption before logging
    If LogsOnOff Then
        If Not MyException = "" Then
            LogEvent = "Failure: " & LogEvent & vbCrLf & MyException
            Logs(LogEvent)
        Else
            LogEvent = "Success: " & LogEvent
            Logs(LogEvent)
        End If
    End If

'Cleanup
    LogEvent = Nothing
    Path = Nothing

'Exit
    Return Returned

End Function

#End Region
#Region "GetLastAccessTimeUtc"

'*************************************************************************
Public Overloads Function GetLastAccessTimeUtc(ByVal Path As String _
                                                ) As String
'PURPOSE:
'   To get the last time the specified file was written to. This value is expressed in UTC time.

'
'OUTPUT:
'   Function returns a datetime object or a object with the standard ": " exception message.
'*************************************************************************
Dim LogEvent As String = "Directory.GetLastAccessTimeUtc" & Path
Dim Returned As String = MyBase.GetLastAccessTimeUtc(Path)

'Check for execption before logging
    If LogsOnOff Then
        If Not MyException = "" Then
            LogEvent = "Failure: " & LogEvent & vbCrLf & MyException
            Logs(LogEvent)
        Else
            LogEvent = "Success: " & LogEvent
            Logs(LogEvent)
        End If
    End If

'Cleanup
    LogEvent = Nothing
    Path = Nothing

'Exit
    Return Returned

End Function

#End Region
#End Region
#Region "SetAttributes"
#Region "SetLastWriteTime"

'*************************************************************************
Public Overloads Function SetLastWriteTime(ByVal Path As String, _
                                            ByVal LastAccessTime As String _
                                            ) As Boolean
'PURPOSE:
'   Sets the date and time a directory was last written to.
'
'OUTPUT:
'   Function returns The string "True" if suscessful or ": exception message" 
'   if an exception was thrown.
'*************************************************************************
Dim LogEvent As String = "Directory.SetLastWriteTime" & Path & ", " & LastAccessTime
Dim Returned As Boolean = MyBase.SetLastWriteTime(Path, LastAccessTime)

'Check for execption before logging
    If LogsOnOff Then
        If Not MyException = "" Then
            LogEvent = "Failure: " & LogEvent & vbCrLf & MyException
            Logs(LogEvent)
        Else
            LogEvent = "Success: " & LogEvent
            Logs(LogEvent)
        End If
    End If

'Cleanup
    LogEvent = Nothing
    Path = Nothing
    LastAccessTime = Nothing

'Exit
    Return Returned

End Function

#End Region
#Region "SetLastWriteTimeUtc"

'*************************************************************************
Public Overloads Function SetLastWriteTimeUtc(ByVal Path As String, _
                                            ByVal LastAccessTime As String _
                                            ) As Boolean
'PURPOSE:
'   Sets the date and time, in universal coordinated time (UTC) format, that
'   a directory was last written to.
'
'OUTPUT:
'   Function returns The string "True" if suscessful or ": exception message" 
'   if an exception was thrown.
'*************************************************************************
Dim LogEvent As String = "Directory.SetLastWriteTimeUtc" & Path & ", " & LastAccessTime
Dim Returned As Boolean = MyBase.SetLastWriteTimeUtc(Path, LastAccessTime)

'Check for execption before logging
    If LogsOnOff Then
        If Not MyException = "" Then
            LogEvent = "Failure: " & LogEvent & vbCrLf & MyException
            Logs(LogEvent)
        Else
            LogEvent = "Success: " & LogEvent
            Logs(LogEvent)
        End If
    End If

'Cleanup
    LogEvent = Nothing
    Path = Nothing
    LastAccessTime = Nothing

'Exit
    Return Returned

End Function

#End Region
#Region "SetCreationTime"

'*************************************************************************
Public Overloads Function SetCreationTime(ByVal Path As String, _
                                        ByVal CreationTime As String _
                                        ) As Boolean
'PURPOSE:
'   To set the time the file was created to the current date and time.

'
'OUTPUT:
'   Function returns an "True" string or a ": " exception message.
'*************************************************************************
Dim LogEvent As String = "Directory.SetCreationTime" & Path & ", " & CreationTime
Dim Returned As Boolean = MyBase.SetCreationTime(Path, CreationTime)

'Check for execption before logging
    If LogsOnOff Then
        If Not MyException = "" Then
            LogEvent = "Failure: " & LogEvent & vbCrLf & MyException
            Logs(LogEvent)
        Else
            LogEvent = "Success: " & LogEvent
            Logs(LogEvent)
        End If
    End If

'Cleanup
    LogEvent = Nothing
    Path = Nothing
    CreationTime = Nothing

'Exit
    Return Returned

End Function

#End Region
#Region "SetCreationTimeUtc"

'*************************************************************************
Public Overloads Function SetCreationTimeUtc(ByVal Path As String, _
                                            ByVal CreationTime As String _
                                            ) As Boolean
'PURPOSE:
'   To set the time the file was created to the current date and time. This value is expressed in UTC time.

'
'OUTPUT:
'   Function returns an "True" string or a ": " exception message.
'*************************************************************************
Dim LogEvent As String = "Directory.SetCreationTimeUtc" & Path & ", " & CreationTime
Dim Returned As Boolean = MyBase.SetCreationTimeUtc(Path, CreationTime)

'Check for execption before logging
    If LogsOnOff Then
        If Not MyException = "" Then
            LogEvent = "Failure: " & LogEvent & vbCrLf & MyException
            Logs(LogEvent)
        Else
            LogEvent = "Success: " & LogEvent
            Logs(LogEvent)
        End If
    End If

'Cleanup
    LogEvent = Nothing
    Path = Nothing
    CreationTime = Nothing

'Exit
    Return Returned

End Function

#End Region
#Region "SetLastAccessTime"

'*************************************************************************
Public Overloads Function SetLastAccessTime(ByVal Path As String, _
                                            ByVal LastAccessTime As String _
                                            ) As Boolean
'PURPOSE:
'   Sets the date and time the specified file or directory was last accessed.
'
'OUTPUT:
'   Function returns The string "True" if suscessful or ": exception message" 
'   if an exception was thrown.
'*************************************************************************
Dim LogEvent As String = "Directory.SetLastAccessTime" & Path & ", " & LastAccessTime
Dim Returned As Boolean = MyBase.SetLastAccessTime(Path, LastAccessTime)

'Check for execption before logging
    If LogsOnOff Then
        If Not MyException = "" Then
            LogEvent = "Failure: " & LogEvent & vbCrLf & MyException
            Logs(LogEvent)
        Else
            LogEvent = "Success: " & LogEvent
            Logs(LogEvent)
        End If
    End If

'Cleanup
    LogEvent = Nothing
    Path = Nothing
    LastAccessTime = Nothing

'Exit
    Return Returned

End Function

#End Region
#Region "SetLastAccessTimeUtc"

'*************************************************************************
Public Overloads Function SetLastAccessTimeUtc(ByVal Path As String, _
                                                ByVal LastAccessTime As String _
                                                ) As Boolean
'PURPOSE:
'   Sets the date and time, in universal coordinated time (UTC) format, that 
'   the specified file or directory was last accessed.
'
'OUTPUT:
'   Function returns The string "True" if suscessful or ": exception message" 
'   if an exception was thrown.
'*************************************************************************
Dim LogEvent As String = "Directory.SetLastAccessTimeUtc" & Path & ", " & LastAccessTime
Dim Returned As Boolean = MyBase.SetLastAccessTimeUtc(Path, LastAccessTime)

'Check for execption before logging
    If LogsOnOff Then
        If Not MyException = "" Then
            LogEvent = "Failure: " & LogEvent & vbCrLf & MyException
            Logs(LogEvent)
        Else
            LogEvent = "Success: " & LogEvent
            Logs(LogEvent)
        End If
    End If

'Cleanup
    LogEvent = Nothing
    Path = Nothing
    LastAccessTime = Nothing

'Exit
    Return Returned

End Function

#End Region
#End Region
#Region "Logs"

'*************************************************************************
Private Sub Logs(ByVal LogEvent As String)
'PURPOSE:
'   To Write a line to the Log file.
'
'OUTPUT:
'   Another line written to the log file.
'*************************************************************************

'Write the string to the text file.
    Pnm.Logs.addEvent(LogEvent)

'Cleanup
    LogEvent = Nothing

'Exit
End Sub

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