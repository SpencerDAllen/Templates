#Region "Comments"

'*****************************************************************************'
'PURPOSE:
'   Use the File class for typical operations such as copying, moving, renaming, 
'   creating, opening, deleting, and appending to files. You can also use the File 
'   class to get and set file attributes or DateTime information related to the 
'   creation, access, and writing of a file.
'
'   Many of the File methods return other I/O types when you create or open files. 
'   You can use these other types to futher manipulate a file. For more information, 
'   see specific File members such as OpenText, CreateText, or Create.
'
'   Because all File methods are static, it might be more efficient to use a File 
'   method rather than a corresponding FileInfo instance method if you want to perform 
'   only one action. All File methods require the path to the file that you are manipulating.
'
'   The static methods of the File class perform security checks on all methods. If 
'   you are going to reuse an object several times, consider using the corresponding 
'   instance method of FileInfo instead, because the security check will not always 
'   be necessary.
'
'   By default, full read/write access to new files is granted to all users.
'
'   Some Functions return FileStreams. To account for this they have corresponding "CheckForException"
'   Functions. THESE TWO FUNCTIONS NEED TO BE USED AS A PAIR!
'
'PREREQUISITES:
'   .Net 1.1
'   BaseClsFile must be added to the project.
'
'USAGE:
'   Boolean = Exists(Path as string)
'   Boolean = Move(SourcePath as string, DestinationPath as string)
'   Boolean = Copy(SourcePath as sting, DestinationPath as string)
'   Boolean = Delete(Path as string)
'   FileStream = Create(Path As String)
'   FileStream = OpenWrite(path as string)
'   FileStream = OpenRead(path as string)
'   DateTime = GetCreationTime(path as string)
'   DateTime = GetCreationTimeUTC(path as string)
'   DateTime = GetLastAccessTime(path as string)
'   DateTime = GetLastAccessTimeUTC(path as string)
'   DateTime = GetLastWriteTime(path as string)
'   DateTime = GetLastWriteTimeUTC(path as string)
'   Boolean = GetAttribute_Archive(path as string)
'   Boolean = GetAttribute_Compressed(path as string)
'   Boolean = GetAttribute_Encrypted(path as string)
'   Boolean = GetAttribute_Hidded(path as string)
'   Boolean = GetAttribute_NotContentIndexed(path as string)
'   Boolean = GetAttribute_Offline(path as string)
'   Boolean = GetAttribute_ReadOnly(path as string)
'   Boolean = GetAttribute_System(path as string)
'   Boolean = GetAttribute_Temporary(path as string)
'   Boolean = SetCreationTime(path as string, CreationTime as datetime)
'   Boolean = SetCreationTimeUTC(path as string, CreationTimeUTC as datetime)
'   Boolean = SetLastAccessTime(path as string, AccessTime as datetime)
'   Boolean = SetLastAccessTimeUTC(path as string, AccessTimeUTC as datetime)
'   Boolean = SetLastWriteTime(path as string, WriteTime as datetime)
'   Boolean = SetLastWriteTimeUTC(path as string, WriteTimeUTC as datetime)
'   Boolean = SetAttribute_Archive(path as string)
'   Boolean = SetAttribute_Compressed(path as string)
'   Boolean = SetAttribute_Encrypted(path as string)
'   Boolean = SetAttribute_Hidded(path as string)
'   Boolean = SetAttribute_NotContentIndexed(path as string)
'   Boolean = SetAttribute_Offline(path as string)
'   Boolean = SetAttribute_ReadOnly(path as string)
'   Boolean = SetAttribute_System(path as string)
'   Boolean = SetAttribute_Temporary(path as string)
'
'AUTHOR:
'   Spencer Allen
'   3/21/08
'
'VERSIONS:
'   1.00 - Base Version (Spencer Allen)
'*****************************************************************************'

#End Region
Public Class MyClsFile
#Region "Inherits"
Inherits BaseClsFile
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
'   Find out if a file exists.
'
'OUTPUT:
'   Function returns boolean.
'*************************************************************************
Dim LogEvent As String = "File.Exists : " & Path
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
'   To Move a file from one location to another.
'
'OUTPUT:
'   Function returns boolean.
'*************************************************************************
Dim LogEvent As String = "File.Move " & SourcePath & ", " & DestinationPath
Dim Returned As Boolean = MyBase.Move(SourcePath, DestinationPath)

'Check for exception before logging
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
    SourcePath = Nothing
    DestinationPath = Nothing
    LogEvent = Nothing

'Exit
    Return Returned

End Function

#End Region
#Region "Copy"

'*************************************************************************
Public Overloads Function Copy(ByVal SourcePath As String, _
                                ByVal DestinationPath As String, _
                                ByVal OverWrite As Boolean _
                                ) As Boolean
'PURPOSE:
'   To make a copy of a file in a different location.
'
'OUTPUT:
'   Function returns boolean.
'*************************************************************************
Dim LogEvent As String = "File.Copy " & SourcePath & ", " & DestinationPath
Dim Returned As Boolean = MyBase.Copy(SourcePath, DestinationPath, OverWrite)

'Check for exception before logging
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
    SourcePath = Nothing
    DestinationPath = Nothing
    LogEvent = Nothing

'Exit
    Return Returned

End Function

#End Region
#Region "Delete"

'*************************************************************************
Public Overloads Function Delete(ByVal Path As String _
                                ) As Boolean
'PURPOSE:
'   To delete a file.
'
'OUTPUT:
'   Function returns boolean.
'*************************************************************************
Dim LogEvent As String = "File.Delete " & Path
Dim Returned As Boolean = MyBase.Delete(Path)

'Check for exception before logging
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
    Path = Nothing
    LogEvent = Nothing

'Exit
    Return Returned

End Function

#End Region
#Region "Create"

'*************************************************************************
Public Overloads Function Create(ByVal Path As String _
                                ) As System.IO.FileStream
'PURPOSE:
'   To Create a file.
'
'OUTPUT:
'   Function returns a FileStream.
'*************************************************************************
Dim LogEvent As String = "File.Create " & Path
Dim Returned As System.IO.FileStream = MyBase.Create(Path)

'Check for exception before logging
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
#Region "OpenWrite"

'*************************************************************************
Public Overloads Function OpenWrite(ByVal Path As String _
                                    ) As System.IO.FileStream
'PURPOSE:
'   To open a file for editing.
'
'OUTPUT:
'   Function returns A filestream for writing.
'*************************************************************************
Dim LogEvent As String = "File.OpenWrite " & Path
Dim Returned As System.IO.FileStream = MyBase.OpenWrite(Path)

'Check for exception before logging
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
#Region "OpenRead"

'*************************************************************************
Public Overloads Function OpenRead(ByVal Path As String _
                                    ) As System.IO.FileStream
'PURPOSE:
'   To open a file for reading.
'
'OUTPUT:
'   Function returns A filestream.
'*************************************************************************
Dim LogEvent As String = "File.OpenRead " & Path
Dim Returned As System.IO.FileStream = MyBase.OpenRead(Path)

'Check for exception before logging
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
    Path = Nothing
    LogEvent = Nothing

'Exit
    Return Returned

End Function

#End Region
#Region "GetAttribute"
#Region "GetCreationTime"

'*************************************************************************
Public Overloads Function GetCreationTime(ByVal Path As String _
                                                ) As DateTime
'PURPOSE:
'   To get the Creation time attribute of a file.
'
'OUTPUT:
'   Function returns A Date Time object.
'*************************************************************************
Dim LogEvent As String = "File.GetCreationTime " & Path
Dim Returned As DateTime = MyBase.GetCreationTime(Path)

'Check for exception before logging
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
    Path = Nothing
    LogEvent = Nothing

'Exit
    Return Returned

End Function

#End Region
#Region "GetCreationTimeUTC"

'*************************************************************************
Public Overloads Function GetCreationTimeUTC(ByVal Path As String _
                                                ) As DateTime
'PURPOSE:
'   To get the Creation time UTC attribute of a file.
'
'OUTPUT:
'   Function returns A Date Time object.
'*************************************************************************
Dim LogEvent As String = "File.GetCreationTimeUTC " & Path
Dim Returned As DateTime = MyBase.GetCreationTimeUtc(Path)

'Check for exception before logging
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
    Path = Nothing
    LogEvent = Nothing

'Exit
    Return Returned

End Function

#End Region
#Region "GetLastAccessTime"

'*************************************************************************
Public Overloads Function GetLastAccessTime(ByVal Path As String _
                                                ) As DateTime
'PURPOSE:
'   To get the Last Access time attribute of a file.
'
'OUTPUT:
'   Function returns A Date Time object.
'*************************************************************************
Dim LogEvent As String = "File.GetLastAccessTime " & Path
Dim Returned As DateTime = MyBase.GetLastAccessTime(Path)

'Check for exception before logging
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
    Path = Nothing
    LogEvent = Nothing

'Exit
    Return Returned

End Function

#End Region
#Region "GetLastAccessTimeUtc"

'*************************************************************************
Public Overloads Function GetLastAccessTimeUtc(ByVal Path As String _
                                                ) As DateTime
'PURPOSE:
'   To get the Last Access Time UTC attribute of a file.
'
'OUTPUT:
'   Function returns A Date Time object.
'*************************************************************************
Dim LogEvent As String = "File.GetLastAccessTimeUtc " & Path
Dim Returned As DateTime = MyBase.GetLastAccessTimeUtc(Path)

'Check for exception before logging
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
    Path = Nothing
    LogEvent = Nothing

'Exit
    Return Returned

End Function

#End Region
#Region "GetLastWriteTime"

'*************************************************************************
Public Overloads Function GetLastWriteTime(ByVal Path As String _
                                                ) As DateTime
'PURPOSE:
'   To get the Last Write Time attribute of a file.
'
'OUTPUT:
'   Function returns A Date Time object.
'*************************************************************************
Dim LogEvent As String = "File.GetLastWriteTime " & Path
Dim Returned As DateTime = MyBase.GetLastWriteTime(Path)

'Check for exception before logging
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
    Path = Nothing
    LogEvent = Nothing

'Exit
    Return Returned

End Function

#End Region
#Region "GetLastWriteTimeUTC"

'*************************************************************************
Public Overloads Function GetLastWriteTimeUTC(ByVal Path As String _
                                                ) As DateTime
'PURPOSE:
'   To get the Last Write Time UTC attribute of a file.
'
'OUTPUT:
'   Function returns A Date Time object.
'*************************************************************************
Dim LogEvent As String = "File.GetLastWriteTimeUTC " & Path
Dim Returned As DateTime = MyBase.GetLastWriteTimeUtc(Path)

'Check for exception before logging
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
    Path = Nothing
    LogEvent = Nothing

'Exit
    Return Returned

End Function

#End Region
#Region "GetAttribute_Archive"

'*************************************************************************
Public Overloads Function GetAttribute_Archive(ByVal Path As String _
                                                ) As Boolean
'PURPOSE:
'   To get the Archive attribute of a file.
'
'OUTPUT:
'   Function returns boolean.
'*************************************************************************
Dim LogEvent As String = "File.GetAttribute_Archive " & Path
Dim Returned As Boolean = MyBase.GetAttribute_Archive(Path)

'Check for exception before logging
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
    Path = Nothing
    LogEvent = Nothing

'Exit
    Return Returned

End Function

#End Region
#Region "GetAttribute_Compressed"

'*************************************************************************
Public Overloads Function GetAttribute_Compressed(ByVal Path As String _
                                                ) As Boolean
'PURPOSE:
'   To get the Compressed attribute of a file.
'
'OUTPUT:
'   Function returns boolean.
'*************************************************************************
Dim LogEvent As String = "File.GetAttribute_Compressed " & Path
Dim Returned As Boolean = MyBase.GetAttribute_Compressed(Path)

'Check for exception before logging
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
    Path = Nothing
    LogEvent = Nothing

'Exit
    Return Returned

End Function

#End Region
#Region "GetAttribute_Encrypted"

'*************************************************************************
Public Overloads Function GetAttribute_Encrypted(ByVal Path As String _
                                                ) As Boolean
'PURPOSE:
'   To get the Encrypted attribute of a file.
'
'OUTPUT:
'   Function returns boolean.
'*************************************************************************
Dim LogEvent As String = "File.GetAttribute_Encrypted " & Path
Dim Returned As Boolean = MyBase.GetAttribute_Encrypted(Path)

'Check for exception before logging
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
    Path = Nothing
    LogEvent = Nothing

'Exit
    Return Returned

End Function

#End Region
#Region "GetAttribute_Hidden"

'*************************************************************************
Public Overloads Function GetAttribute_Hidden(ByVal Path As String _
                                            ) As Boolean
'PURPOSE:
'   To get the Hidden attribute of a file.
'
'OUTPUT:
'   Function returns boolean.
'*************************************************************************
Dim LogEvent As String = "File.GetAttribute_Hidden " & Path
Dim Returned As Boolean = MyBase.GetAttribute_Hidden(Path)

'Check for exception before logging
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
    Path = Nothing
    LogEvent = Nothing

'Exit
    Return Returned

End Function

#End Region
#Region "GetAttribute_NotContentIndexed"

'*************************************************************************
Public Overloads Function GetAttribute_NotContentIndexed(ByVal Path As String _
                                                        ) As Boolean
'PURPOSE:
'   To get the NotContentIndexed attribute of a file.
'
'OUTPUT:
'   Function returns boolean.
'*************************************************************************
Dim LogEvent As String = "File.GetAttribute_NotContentIndexed " & Path
Dim Returned As Boolean = MyBase.GetAttribute_NotContentIndexed(Path)

'Check for exception before logging
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
    Path = Nothing
    LogEvent = Nothing

'Exit
    Return Returned

End Function

#End Region
#Region "GetAttribute_Offline"

'*************************************************************************
Public Overloads Function GetAttribute_Offline(ByVal Path As String _
                                                ) As Boolean
'PURPOSE:
'   To get the Offline attribute of a file.
'
'OUTPUT:
'   Function returns boolean.
'*************************************************************************
Dim LogEvent As String = "File.GetAttribute_Offline " & Path
Dim Returned As Boolean = MyBase.GetAttribute_Offline(Path)

'Check for exception before logging
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
    Path = Nothing
    LogEvent = Nothing

'Exit
    Return Returned

End Function

#End Region
#Region "GetAttribute_ReadOnly"

'*************************************************************************
Public Overloads Function GetAttribute_ReadOnly(ByVal Path As String _
                                        ) As Boolean
'PURPOSE:
'   To get the ReadOnly attribute of a file.
'
'OUTPUT:
'   Function returns boolean.
'*************************************************************************
Dim LogEvent As String = "File.GetAttribute_ReadOnly " & Path
Dim Returned As Boolean = MyBase.GetAttribute_ReadOnly(Path)

'Check for exception before logging
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
    Path = Nothing
    LogEvent = Nothing

'Exit
    Return Returned

End Function

#End Region
#Region "GetAttribute_System"

'*************************************************************************
Public Overloads Function GetAttribute_System(ByVal Path As String _
                                            ) As Boolean
'PURPOSE:
'   To get the System attribute of a file.
'
'OUTPUT:
'   Function returns boolean.
'*************************************************************************
Dim LogEvent As String = "File.GetAttribute_System " & Path
Dim Returned As Boolean = MyBase.GetAttribute_System(Path)

'Check for exception before logging
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
    Path = Nothing
    LogEvent = Nothing

'Exit
    Return Returned

End Function

#End Region
#Region "GetAttribute_Temporary"

'*************************************************************************
Public Overloads Function GetAttribute_Temporary(ByVal Path As String _
                                                ) As Boolean
'PURPOSE:
'   To get the Temporary attribute of a file.
'
'OUTPUT:
'   Function returns boolean.
'*************************************************************************
Dim LogEvent As String = "File.GetAttribute_Temporary " & Path
Dim Returned As Boolean = MyBase.GetAttribute_Temporary(Path)

'Check for exception before logging
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
    Path = Nothing
    LogEvent = Nothing

'Exit
    Return Returned

End Function

#End Region
#End Region
#Region "SetAttribute"
#Region "SetCreationTime"

'*************************************************************************
Public Overloads Function SetCreationTime(ByVal Path As String, _
                                        ByVal CreationTime As DateTime _
                                        ) As Boolean
'PURPOSE:
'   To Set the Archive attribute of a file.
'
'OUTPUT:
'   Function returns boolean.
'*************************************************************************
Dim LogEvent As String = "File.SetCreationTime " & Path & " " & CreationTime
Dim Returned As Boolean = MyBase.SetCreationTime(Path, CreationTime)

'Check for exception before logging
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
    Path = Nothing
    LogEvent = Nothing

'Exit
    Return Returned

End Function

#End Region
#Region "SetCreationTimeUtc"

'*************************************************************************
Public Overloads Function SetCreationTimeUtc(ByVal Path As String, _
                                        ByVal CreationTimeUtc As DateTime _
                                        ) As Boolean
'PURPOSE:
'   To Set the Archive attribute of a file.
'
'OUTPUT:
'   Function returns boolean.
'*************************************************************************
Dim LogEvent As String = "File.SetCreationTimeUtc " & Path & " " & CreationTimeUtc
Dim Returned As Boolean = MyBase.SetCreationTimeUtc(Path, CreationTimeUtc)

'Check for exception before logging
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
    Path = Nothing
    LogEvent = Nothing

'Exit
    Return Returned

End Function

#End Region
#Region "SetLastAccessTime"

'*************************************************************************
Public Overloads Function SetLastAccessTime(ByVal Path As String, _
                                        ByVal AccessTime As DateTime _
                                        ) As Boolean
'PURPOSE:
'   To Set the Archive attribute of a file.
'
'OUTPUT:
'   Function returns boolean.
'*************************************************************************
Dim LogEvent As String = "File.SetLastAccessTime " & Path & " " & AccessTime
Dim Returned As Boolean = MyBase.SetLastAccessTime(Path, AccessTime)

'Check for exception before logging
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
    Path = Nothing
    LogEvent = Nothing

'Exit
    Return Returned

End Function

#End Region
#Region "SetLastAccessTimeUtc"

'*************************************************************************
Public Overloads Function SetLastAccessTimeUtc(ByVal Path As String, _
                                        ByVal AccessTimeUtc As DateTime _
                                        ) As Boolean
'PURPOSE:
'   To Set the Archive attribute of a file.
'
'OUTPUT:
'   Function returns boolean.
'*************************************************************************
Dim LogEvent As String = "File.SetLastAccessTimeUtc " & Path & " " & AccessTimeUtc
Dim Returned As Boolean = MyBase.SetLastAccessTimeUtc(Path, AccessTimeUtc)

'Check for exception before logging
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
    Path = Nothing
    LogEvent = Nothing

'Exit
    Return Returned

End Function

#End Region
#Region "SetLastWriteTime"

'*************************************************************************
Public Overloads Function SetLastWriteTime(ByVal Path As String, _
                                        ByVal WriteTime As DateTime _
                                        ) As Boolean
'PURPOSE:
'   To Set the Archive attribute of a file.
'
'OUTPUT:
'   Function returns boolean.
'*************************************************************************
Dim LogEvent As String = "File.SetLastWriteTime " & Path & " " & WriteTime
Dim Returned As Boolean = MyBase.SetLastWriteTime(Path, WriteTime)

'Check for exception before logging
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
    Path = Nothing
    LogEvent = Nothing

'Exit
    Return Returned

End Function

#End Region
#Region "SetLastWriteTimeUtc"

'*************************************************************************
Public Overloads Function SetLastWriteTimeUtc(ByVal Path As String, _
                                        ByVal WriteTimeUtc As DateTime _
                                        ) As Boolean
'PURPOSE:
'   To Set the Archive attribute of a file.
'
'OUTPUT:
'   Function returns boolean.
'*************************************************************************
Dim LogEvent As String = "File.SetLastWriteTimeUtc " & Path & " " & WriteTimeUtc
Dim Returned As Boolean = MyBase.SetLastWriteTimeUtc(Path, WriteTimeUtc)

'Check for exception before logging
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
    Path = Nothing
    LogEvent = Nothing

'Exit
    Return Returned

End Function

#End Region
#Region "SetAttribute_Archive"

'*************************************************************************
Public Overloads Function SetAttribute_Archive(ByVal Path As String _
                                                ) As Boolean
'PURPOSE:
'   To Set the Archive attribute of a file.
'
'OUTPUT:
'   Function returns boolean.
'*************************************************************************
Dim LogEvent As String = "File.SetAttribute_Archive " & Path
Dim Returned As Boolean = MyBase.SetAttribute_Archive(Path)

'Check for exception before logging
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
    Path = Nothing
    LogEvent = Nothing

'Exit
    Return Returned

End Function

#End Region
#Region "SetAttribute_Compressed"

'*************************************************************************
Public Overloads Function SetAttribute_Compressed(ByVal Path As String _
                                        ) As Boolean
'PURPOSE:
'   To Set the Compressed attribute of a file.
'
'OUTPUT:
'   Function returns boolean.
'*************************************************************************
Dim LogEvent As String = "File.SetAttribute_Compressed " & Path
Dim Returned As Boolean = MyBase.SetAttribute_Compressed(Path)

'Check for exception before logging
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
    Path = Nothing
    LogEvent = Nothing

'Exit
    Return Returned

End Function

#End Region
#Region "SetAttribute_Encrypted"

'*************************************************************************
Public Overloads Function SetAttribute_Encrypted(ByVal Path As String _
                                                ) As Boolean
'PURPOSE:
'   To Set the Encrypted attribute of a file.
'
'OUTPUT:
'   Function returns boolean.
'*************************************************************************
Dim LogEvent As String = "File.SetAttribute_Encrypted " & Path
Dim Returned As Boolean = MyBase.SetAttribute_Encrypted(Path)

'Check for exception before logging
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
    Path = Nothing
    LogEvent = Nothing

'Exit
    Return Returned

End Function

#End Region
#Region "SetAttribute_Hidden"

'*************************************************************************
Public Overloads Function SetAttribute_Hidden(ByVal Path As String _
                                            ) As Boolean
'PURPOSE:
'   To Set the Hidden attribute of a file.
'
'OUTPUT:
'   Function returns boolean.
'*************************************************************************
Dim LogEvent As String = "File.SetAttribute_Hidden " & Path
Dim Returned As Boolean = MyBase.SetAttribute_Hidden(Path)

'Check for exception before logging
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
    Path = Nothing
    LogEvent = Nothing

'Exit
    Return Returned

End Function

#End Region
#Region "SetAttribute_NotContentIndexed"

'*************************************************************************
Public Overloads Function SetAttribute_NotContentIndexed(ByVal Path As String _
                                                        ) As Boolean
'PURPOSE:
'   To Set the NotContentIndexed attribute of a file.
'
'OUTPUT:
'   Function returns boolean.
'*************************************************************************
Dim LogEvent As String = "File.SetAttribute_NotContentIndexed " & Path
Dim Returned As Boolean = MyBase.SetAttribute_NotContentIndexed(Path)

'Check for exception before logging
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
    Path = Nothing
    LogEvent = Nothing

'Exit
    Return Returned

End Function

#End Region
#Region "SetAttribute_Offline"

'*************************************************************************
Public Overloads Function SetAttribute_Offline(ByVal Path As String _
                                            ) As Boolean
'PURPOSE:
'   To Set the Offline attribute of a file.
'
'OUTPUT:
'   Function returns boolean.
'*************************************************************************
Dim LogEvent As String = "File.SetAttribute_Offline " & Path
Dim Returned As Boolean = MyBase.SetAttribute_Offline(Path)

'Check for exception before logging
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
    Path = Nothing
    LogEvent = Nothing

'Exit
    Return Returned

End Function

#End Region
#Region "SetAttribute_ReadOnly"

'*************************************************************************
Public Overloads Function SetAttribute_ReadOnly(ByVal Path As String _
                                        ) As Boolean
'PURPOSE:
'   To Set the ReadOnly attribute of a file.
'
'OUTPUT:
'   Function returns boolean.
'*************************************************************************
Dim LogEvent As String = "File.SetAttribute_ReadOnly " & Path
Dim Returned As Boolean = MyBase.SetAttribute_ReadOnly(Path)

'Check for exception before logging
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
    Path = Nothing
    LogEvent = Nothing

'Exit
    Return Returned

End Function

#End Region
#Region "SetAttribute_System"

'*************************************************************************
Public Overloads Function SetAttribute_System(ByVal Path As String _
                                            ) As Boolean
'PURPOSE:
'   To Set the System attribute of a file.
'
'OUTPUT:
'   Function returns boolean.
'*************************************************************************
Dim LogEvent As String = "File.SetAttribute_System " & Path
Dim Returned As Boolean = MyBase.SetAttribute_System(Path)

'Check for exception before logging
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
    Path = Nothing
    LogEvent = Nothing

'Exit
    Return Returned

End Function

#End Region
#Region "SetAttribute_Temporary"

'*************************************************************************
Public Overloads Function SetAttribute_Temporary(ByVal Path As String _
                                                ) As Boolean
'PURPOSE:
'   To Set the Temporary attribute of a file.
'
'OUTPUT:
'   Function returns boolean.
'*************************************************************************
Dim LogEvent As String = "File.SetAttribute_Temporary " & Path
Dim Returned As Boolean = MyBase.SetAttribute_Temporary(Path)

'Check for exception before logging
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
    Path = Nothing
    LogEvent = Nothing

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
#Region "OnOff Property"

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