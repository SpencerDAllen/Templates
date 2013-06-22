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
'   Boolean = SetCreationTime(path as string, CreationTime as string)
'   Boolean = SetCreationTimeUTC(path as string, CreationTimeUTC as string)
'   Boolean = SetLastAccessTime(path as string, LastAccessTime as string)
'   Boolean = SetLastAccessTimeUTC(path as string, LastAccessTimeUTC as string)
'   Boolean = SetLastWriteTime(path as string, WriteTime as string)
'   Boolean = SetLastWriteTimeUTC(path as string, WriteTimeUTC as string)
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
#Region "Imports"

    Imports System.IO

#End Region
Public MustInherit Class BaseClsFile
#Region "Exists"

'*************************************************************************
Protected Function Exists(ByVal Path As String _
                        ) As Boolean
'PURPOSE:
'   Find out if a file exists.
'
'OUTPUT:
'   Function returns boolean and sets the MyException Property.
'*************************************************************************
Dim ErrorFlag As Boolean

'Find out if the file exists.
    Try
        ErrorFlag = File.Exists(Path)
        Except = ""
    Catch ex As Exception
        Except = ex.Message
        ErrorFlag = False
    End Try

'Cleanup
    Path = Nothing

'Exit
    Return ErrorFlag

End Function

#End Region
#Region "Move"

'*************************************************************************
Protected Function Move(ByVal SourcePath As String, _
                    ByVal DestinationPath As String _
                    ) As Boolean
'PURPOSE:
'   To Move a file from one location to another.
'
'OUTPUT:
'   Function returns boolean and sets the MyException Property.
'*************************************************************************
Dim Returns As Boolean

'Move the file
    Try
        File.Move(SourcePath, DestinationPath)
        Returns = True
        Except = ""
    Catch ex As Exception
        Returns = False
        Except = ex.Message
    End Try

'Cleanup
    SourcePath = Nothing
    DestinationPath = Nothing

'Exit
    Return Returns

End Function

#End Region
#Region "Copy"

'*************************************************************************
Protected Function Copy(ByVal SourcePath As String, _
                    ByVal DestinationPath As String, _
                    ByVal OverWrite As Boolean _
                    ) As Boolean
'PURPOSE:
'   To make a copy of a file in a different location.
'
'OUTPUT:
'   Function returns boolean and sets the MyException Property.
'*************************************************************************
Dim Returns As Boolean

'Copy the file
    Try
        File.Copy(SourcePath, DestinationPath, OverWrite)
        Returns = True
        Except = ""
    Catch ex As Exception
        Returns = False
        Except = ex.Message
    End Try

'Cleanup
    SourcePath = Nothing
    DestinationPath = Nothing
    OverWrite = Nothing

'Exit
    Return Returns

End Function

#End Region
#Region "Delete"

'*************************************************************************
Protected Function Delete(ByVal Path As String _
                        ) As Boolean
'PURPOSE:
'   To delete a file.
'
'OUTPUT:
'   Function returns boolean and sets the MyException Property.
'*************************************************************************
Dim Returns As Boolean

'Delete the file
    Try
        File.Delete(Path)
        Returns = True
        Except = ""
    Catch ex As Exception
        Returns = False
        Except = ex.Message
    End Try

'Cleanup
    Path = Nothing

'Exit
    Return Returns

End Function

#End Region
#Region "Create"

'*************************************************************************
Protected Function Create(ByVal Path As String _
                        ) As FileStream
'PURPOSE:
'   To Create a file.
'
'OUTPUT:
'   Function returns a filestream and sets the MyException Property.
'*************************************************************************
Dim Returns As FileStream

'Create the file
Try
    Returns = File.Create(Path)
    Except = ""
Catch ex As Exception
    Except = ex.Message
End Try

'Cleanup
    Path = Nothing

'Exit
    Return Returns

End Function

#End Region
#Region "OpenWrite"

'*************************************************************************
Protected Function OpenWrite(ByVal Path As String _
                        ) As FileStream
'PURPOSE:
'   To open a file for editing.
'
'OUTPUT:
'   Function returns a filestream and sets the MyException Property.
'*************************************************************************
Dim Returns As FileStream

'Opens the file for writing
Try
    Returns = File.OpenWrite(Path)
    Except = ""
Catch ex As Exception
    Except = ex.Message
End Try

'Cleanup
    Path = Nothing

'Exit
    Return Returns

End Function

#End Region
#Region "OpenRead"

'*************************************************************************
Protected Function OpenRead(ByVal Path As String _
                        ) As FileStream
'PURPOSE:
'   To open a file for reading.
'
'OUTPUT:
'   Function returns a filestream and sets the MyException Property.
'*************************************************************************
Dim Returns As FileStream

'Opens the file for reading
Try
    Returns = File.OpenRead(Path)
    Except = ""
Catch ex As Exception
    Except = ex.Message
End Try

'Cleanup
    Path = Nothing

'Exit
    Return Returns

End Function

#End Region
#Region "GetAttribute"
#Region "GetCreationTime"

'*************************************************************************
Protected Function GetCreationTime(ByVal Path As String _
                                    ) As DateTime
'PURPOSE:
'   To get the creation date and time of the specified file or directory.
'
'OUTPUT:
'   Function returns a date time object and sets the MyException Property.
'*************************************************************************
Dim Returns As DateTime
Except = ""

'Get the attribute
    Try
        Returns = File.GetCreationTime(Path)
    Catch ex As Exception
        Except = ex.Message
    End Try

'Cleanup
    Path = Nothing

'Exit
    Return Returns

End Function

#End Region
#Region "GetCreationTimeUtc"

'*************************************************************************
Protected Function GetCreationTimeUtc(ByVal Path As String _
                                    ) As DateTime
'PURPOSE:
'   To get the creation date and time, in coordinated universal time (UTC)
'   of the specified file or directory.
'
'OUTPUT:
'   Function returns a date time object and sets the MyException Property.
'*************************************************************************
Dim Returns As DateTime
Except = ""

'Get the attribute
    Try
        Returns = File.GetCreationTimeUtc(Path)
    Catch ex As Exception
        Except = ex.Message
    End Try

'Cleanup
    Path = Nothing

'Exit
    Return Returns

End Function

#End Region
#Region "GetLastAccessTime"

'*************************************************************************
Protected Function GetLastAccessTime(ByVal Path As String _
                                    ) As DateTime
'PURPOSE:
'   To get the date and time the specified file or directory was last accessed.
'
'OUTPUT:
'   Function returns a date time object and sets the MyException Property.
'*************************************************************************
Dim Returns As DateTime
Except = ""

'Get the attribute
    Try
        Returns = File.GetLastAccessTime(Path)
    Catch ex As Exception
        Except = ex.Message
    End Try

'Cleanup
    Path = Nothing

'Exit
    Return Returns

End Function

#End Region
#Region "GetLastAccessTimeUtc"

'*************************************************************************
Protected Function GetLastAccessTimeUtc(ByVal Path As String _
                                    ) As DateTime
'PURPOSE:
'   To get the date and time, in coordinated universal time (UTC), that the
'   specified file or directory was last accessed.
'
'OUTPUT:
'   Function returns a date time object and sets the MyException Property.
'*************************************************************************
Dim Returns As DateTime
Except = ""

'Get the attribute
    Try
        Returns = File.GetLastAccessTimeUtc(Path)
    Catch ex As Exception
        Except = ex.Message
    End Try

'Cleanup
    Path = Nothing

'Exit
    Return Returns

End Function

#End Region
#Region "GetLastWriteTime"

'*************************************************************************
Protected Function GetLastWriteTime(ByVal Path As String _
                                    ) As DateTime
'PURPOSE:
'   To get the date and time the specified file or directory was last written to.
'
'OUTPUT:
'   Function returns a date time object and sets the MyException Property.
'*************************************************************************
Dim Returns As DateTime
Except = ""

'Get the attribute
    Try
        Returns = File.GetLastWriteTime(Path)
    Catch ex As Exception
        Except = ex.Message
    End Try

'Cleanup
    Path = Nothing

'Exit
    Return Returns

End Function

#End Region
#Region "GetLastWriteTimeUtc"

'*************************************************************************
Protected Function GetLastWriteTimeUtc(ByVal Path As String _
                                    ) As DateTime
'PURPOSE:
'   To get the date and time, in coordinated universal time (UTC), that the
'   specified file or directory was last written to.
'
'OUTPUT:
'   Function returns a date time object and sets the MyException Property.
'*************************************************************************
Dim Returns As DateTime
Except = ""

'Get the attribute
    Try
        Returns = File.GetLastWriteTimeUtc(Path)
    Catch ex As Exception
        Except = ex.Message
    End Try

'Cleanup
    Path = Nothing

'Exit
    Return Returns

End Function

#End Region
#Region "GetAttribute_Archive"

'*************************************************************************
Protected Function GetAttribute_Archive(ByVal Path As String _
                                    ) As Boolean
'PURPOSE:
'   To get the Archive attribute of a file.
'
'OUTPUT:
'   Function returns boolean and sets the MyException Property.
'*************************************************************************
Dim Returns As Boolean

'Get the attribute
    Try
        If (File.GetAttributes(Path) And FileAttributes.Archive) = FileAttributes.Archive Then
            Returns = True
            Except = ""
        Else
            Returns = False
            Except = ""
        End If
    Catch ex As Exception
        Except = ex.Message
        Returns = False
    End Try

'Cleanup
    Path = Nothing

'Exit
    Return Returns

End Function

#End Region
#Region "GetAttribute_Compressed"

'*************************************************************************
Protected Function GetAttribute_Compressed(ByVal Path As String _
                                        ) As Boolean
'PURPOSE:
'   To get the Compressed attribute of a file.
'
'OUTPUT:
'   Function returns boolean and sets the MyException Property.
'*************************************************************************
Dim Returns As Boolean

'Get the attribute
    Try
        If (File.GetAttributes(Path) And FileAttributes.Compressed) = FileAttributes.Compressed Then
            Returns = True
            Except = ""
        Else
            Returns = False
            Except = ""
        End If
    Catch ex As Exception
        Returns = False
        Except = ex.Message
    End Try

'Cleanup
    Path = Nothing

'Exit
    Return Returns

End Function

#End Region
#Region "GetAttribute_Encrypted"

'*************************************************************************
Protected Function GetAttribute_Encrypted(ByVal Path As String _
                                        ) As Boolean
'PURPOSE:
'   To get the Encrypted attribute of a file.
'
'OUTPUT:
'   Function returns boolean and sets the MyException Property.
'*************************************************************************
Dim Returns As Boolean

'Get the attribute
    Try
        If (File.GetAttributes(Path) And FileAttributes.Encrypted) = FileAttributes.Encrypted Then
            Returns = True
            Except = ""
        Else
            Returns = False
            Except = ""
        End If
    Catch ex As Exception
        Except = ex.Message
        Returns = False
    End Try

'Cleanup
    Path = Nothing

'Exit
    Return Returns

End Function

#End Region
#Region "GetAttribute_Hidden"

'*************************************************************************
Protected Function GetAttribute_Hidden(ByVal Path As String _
                                    ) As Boolean
'PURPOSE:
'   To get the Hidden attribute of a file.
'
'OUTPUT:
'   Function returns boolean and sets the MyException Property.
'*************************************************************************
Dim Returns As Boolean

'Get the attribute
    Try
        If (File.GetAttributes(Path) And FileAttributes.Hidden) = FileAttributes.Hidden Then
            Returns = True
            Except = ""
        Else
            Returns = False
            Except = ""
        End If
    Catch ex As Exception
        Except = ex.Message
        Returns = False
    End Try

'Cleanup
    Path = Nothing

'Exit
    Return Returns

End Function

#End Region
#Region "GetAttribute_NotContentIndexed"

'*************************************************************************
Protected Function GetAttribute_NotContentIndexed(ByVal Path As String _
                                                ) As Boolean
'PURPOSE:
'   To get the NotContentIndexed attribute of a file.
'
'OUTPUT:
'   Function returns boolean and sets the MyException Property.
'*************************************************************************
Dim Returns As Boolean

'Get the attribute
    Try
        If (File.GetAttributes(Path) And FileAttributes.NotContentIndexed) = FileAttributes.NotContentIndexed Then
            Returns = True
            Except = ""
        Else
            Returns = False
            Except = ""
        End If
    Catch ex As Exception
        Except = ex.Message
        Returns = False
    End Try

'Cleanup
    Path = Nothing

'Exit
    Return Returns

End Function

#End Region
#Region "GetAttribute_Offline"

'*************************************************************************
Protected Function GetAttribute_Offline(ByVal Path As String _
                                    ) As Boolean
'PURPOSE:
'   To get the Offline attribute of a file.
'
'OUTPUT:
'   Function returns boolean and sets the MyException Property.
'*************************************************************************
Dim Returns As Boolean

'Get the attribute
    Try
        If (File.GetAttributes(Path) And FileAttributes.Offline) = FileAttributes.Offline Then
            Returns = True
            Except = ""
        Else
            Returns = False
            Except = ""
        End If
    Catch ex As Exception
        Except = ex.Message
        Returns = False
    End Try

'Cleanup
    Path = Nothing

'Exit
    Return Returns

End Function

#End Region
#Region "GetAttribute_ReadOnly"

'*************************************************************************
Protected Function GetAttribute_ReadOnly(ByVal Path As String _
                                        ) As Boolean
'PURPOSE:
'   To get the ReadOnly attribute of a file.
'
'OUTPUT:
'   Function returns boolean and sets the MyException Property.
'*************************************************************************
Dim Returns As Boolean

'Get the attribute
    Try
        If (File.GetAttributes(Path) And FileAttributes.ReadOnly) = FileAttributes.ReadOnly Then
            Returns = True
            Except = ""
        Else
            Returns = False
            Except = ""
        End If
    Catch ex As Exception
        Except = ex.Message
        Returns = False
    End Try

'Cleanup
    Path = Nothing

'Exit
    Return Returns

End Function

#End Region
#Region "GetAttribute_System"

'*************************************************************************
Protected Function GetAttribute_System(ByVal Path As String _
                                    ) As Boolean
'PURPOSE:
'   To get the System attribute of a file.
'
'OUTPUT:
'   Function returns boolean and sets the MyException Property.
'*************************************************************************
Dim Returns As Boolean

'Get the attribute
    Try
        If (File.GetAttributes(Path) And FileAttributes.System) = FileAttributes.System Then
            Returns = True
            Except = ""
        Else
            Returns = "False"
            Except = ""
        End If
    Catch ex As Exception
        Except = ex.Message
        Returns = False
    End Try

'Cleanup
    Path = Nothing

'Exit
    Return Returns

End Function

#End Region
#Region "GetAttribute_Temporary"

'*************************************************************************
Protected Function GetAttribute_Temporary(ByVal Path As String _
                                    ) As Boolean
'PURPOSE:
'   To get the Temporary attribute of a file.
'
'OUTPUT:
'   Function returns boolean and sets the MyException Property.
'*************************************************************************
Dim Returns As Boolean

'Get the attribute
    Try
        If (File.GetAttributes(Path) And FileAttributes.Temporary) = FileAttributes.Temporary Then
            Returns = True
            Except = ""
        Else
            Returns = False
            Except = ""
        End If
    Catch ex As Exception
        Except = ex.Message
        Returns = False
    End Try

'Cleanup
    Path = Nothing

'Exit
    Return Returns

End Function

#End Region
#End Region
#Region "SetAttribute"
#Region "SetCreationTime"

'*************************************************************************
Protected Function SetCreationTime(ByVal Path As String, _
                                    ByVal CreationTime As DateTime _
                                    ) As Boolean
'PURPOSE:
'   To Set the date and time the file was created.
'
'OUTPUT:
'   Function returns boolean and sets the MyException Property.
'*************************************************************************
Dim Returns As Boolean
Except = ""

'Set the attribute
    Try
        File.SetCreationTime(Path, CreationTime)
        Returns = True
    Catch ex As Exception
        Except = ex.Message
        Returns = False
    End Try

'Cleanup
    Path = Nothing

'Exit
    Return Returns

End Function

#End Region
#Region "SetCreationTimeUtc"

'*************************************************************************
Protected Function SetCreationTimeUtc(ByVal Path As String, _
                                    ByVal CreationTimeUTC As DateTime _
                                    ) As Boolean
'PURPOSE:
'   Sets the date and time, in coordinated universal time (UTC), that the file
'   was created.
'
'OUTPUT:
'   Function returns boolean and sets the MyException Property.
'*************************************************************************
Dim Returns As Boolean
Except = ""

'Set the attribute
    Try
        File.SetCreationTime(Path, CreationTimeUTC)
        Returns = True
    Catch ex As Exception
        Except = ex.Message
        Returns = False
    End Try

'Cleanup
    Path = Nothing

'Exit
    Return Returns

End Function

#End Region
#Region "SetLastAccessTime"

'*************************************************************************
Protected Function SetLastAccessTime(ByVal Path As String, _
                                    ByVal AccessTime As DateTime _
                                    ) As Boolean
'PURPOSE:
'   Sets the date and time the specified file was last accessed.
'
'OUTPUT:
'   Function returns boolean and sets the MyException Property.
'*************************************************************************
Dim Returns As Boolean
Except = ""

'Set the attribute
    Try
        File.SetLastAccessTime(Path, AccessTime)
        Returns = True
    Catch ex As Exception
        Except = ex.Message
        Returns = False
    End Try

'Cleanup
    Path = Nothing

'Exit
    Return Returns

End Function

#End Region
#Region "SetLastAccessTimeUtc"

'*************************************************************************
Protected Function SetLastAccessTimeUtc(ByVal Path As String, _
                                    ByVal AccessTimeUtc As DateTime _
                                    ) As Boolean
'PURPOSE:
'   Sets the date and time, in coordinated universal time (UTC), that the 
'   specified file was last accessed.
'
'OUTPUT:
'   Function returns boolean and sets the MyException Property.
'*************************************************************************
Dim Returns As Boolean
Except = ""

'Set the attribute
    Try
        File.SetLastAccessTimeUtc(Path, AccessTimeUtc)
        Returns = True
    Catch ex As Exception
        Except = ex.Message
        Returns = False
    End Try

'Cleanup
    Path = Nothing

'Exit
    Return Returns

End Function

#End Region
#Region "SetLastWriteTime"

'*************************************************************************
Protected Function SetLastWriteTime(ByVal Path As String, _
                                    ByVal WriteTime As DateTime _
                                    ) As Boolean
'PURPOSE:
'   Sets the date and time the specified file was last accessed.
'
'OUTPUT:
'   Function returns boolean and sets the MyException Property.
'*************************************************************************
Dim Returns As Boolean
Except = ""

'Set the attribute
    Try
        File.SetLastWriteTime(Path, WriteTime)
        Returns = True
    Catch ex As Exception
        Except = ex.Message
        Returns = False
    End Try

'Cleanup
    Path = Nothing

'Exit
    Return Returns

End Function

#End Region
#Region "SetLastWriteTimeUtc"

'*************************************************************************
Protected Function SetLastWriteTimeUtc(ByVal Path As String, _
                                    ByVal WriteTimeUtc As DateTime _
                                    ) As Boolean
'PURPOSE:
'   Sets the date and time, in coordinated universal time (UTC), that the 
'   specified file was last written to.
'
'OUTPUT:
'   Function returns boolean and sets the MyException Property.
'*************************************************************************
Dim Returns As Boolean
Except = ""

'Set the attribute
    Try
        File.SetLastWriteTimeUtc(Path, WriteTimeUtc)
        Returns = True
    Catch ex As Exception
        Except = ex.Message
        Returns = False
    End Try

'Cleanup
    Path = Nothing

'Exit
    Return Returns

End Function

#End Region
#Region "SetAttribute_Archive"

'*************************************************************************
Protected Function SetAttribute_Archive(ByVal Path As String _
                                    ) As Boolean
'PURPOSE:
'   To Set the Archive attribute of a file.
'
'OUTPUT:
'   Function returns boolean and sets the MyException Property.
'*************************************************************************
Dim Returns As Boolean

'Set the attribute
    Try
        File.SetAttributes(Path, FileAttributes.Archive)
        Returns = True
        Except = ""
    Catch ex As Exception
        Except = ex.Message
        Returns = False
    End Try

'Cleanup
    Path = Nothing

'Exit
    Return Returns

End Function

#End Region
#Region "SetAttribute_Compressed"

'*************************************************************************
Protected Function SetAttribute_Compressed(ByVal Path As String _
                                        ) As Boolean
'PURPOSE:
'   To Set the Compressed attribute of a file.
'
'OUTPUT:
'   Function returns boolean and sets the MyException Property.
'*************************************************************************
Dim Returns As Boolean

'Set the attribute
    Try
        File.SetAttributes(Path, FileAttributes.Compressed)
        Returns = True
        Except = ""
    Catch ex As Exception
        Except = ex.Message
        Returns = False
    End Try

'Cleanup
    Path = Nothing

'Exit
    Return Returns

End Function

#End Region
#Region "SetAttribute_Encrypted"

'*************************************************************************
Protected Function SetAttribute_Encrypted(ByVal Path As String _
                                        ) As Boolean
'PURPOSE:
'   To Set the Encrypted attribute of a file.
'
'OUTPUT:
'   Function returns boolean and sets the MyException Property.
'*************************************************************************
Dim Returns As Boolean

'Set the attribute
    Try
        File.SetAttributes(Path, FileAttributes.Encrypted)
        Returns = True
        Except = ""
    Catch ex As Exception
        Except = ex.Message
        Returns = False
    End Try

'Cleanup
    Path = Nothing

'Exit
    Return Returns

End Function

#End Region
#Region "SetAttribute_Hidden"

'*************************************************************************
Protected Function SetAttribute_Hidden(ByVal Path As String _
                                    ) As Boolean
'PURPOSE:
'   To Set the Hidden attribute of a file.
'
'OUTPUT:
'   Function returns boolean and sets the MyException Property.
'*************************************************************************
Dim Returns As Boolean

'Set the attribute
    Try
        File.SetAttributes(Path, FileAttributes.Hidden)
        Returns = True
        Except = ""
    Catch ex As Exception
        Except = ex.Message
        Returns = False
    End Try

'Cleanup
    Path = Nothing

'Exit
    Return Returns

End Function

#End Region
#Region "SetAttribute_NotContentIndexed"

'*************************************************************************
Protected Function SetAttribute_NotContentIndexed(ByVal Path As String _
                                                ) As Boolean
'PURPOSE:
'   To Set the NotContentIndexed attribute of a file.
'
'OUTPUT:
'   Function returns boolean and sets the MyException Property.
'*************************************************************************
Dim Returns As Boolean

'Set the attribute
    Try
        File.SetAttributes(Path, FileAttributes.NotContentIndexed)
        Returns = True
        Except = ""
    Catch ex As Exception
        Except = ex.Message
        Returns = False
    End Try

'Cleanup
    Path = Nothing

'Exit
    Return Returns

End Function

#End Region
#Region "SetAttribute_Offline"

'*************************************************************************
Protected Function SetAttribute_Offline(ByVal Path As String _
                                    ) As Boolean
'PURPOSE:
'   To Set the Offline attribute of a file.
'
'OUTPUT:
'   Function returns boolean and sets the MyException Property.
'*************************************************************************
Dim Returns As Boolean

'Set the attribute
    Try
        File.SetAttributes(Path, FileAttributes.Offline)
        Returns = True
        Except = ""
    Catch ex As Exception
        Except = ex.Message
        Returns = False
    End Try

'Cleanup
    Path = Nothing

'Exit
    Return Returns

End Function

#End Region
#Region "SetAttribute_ReadOnly"

'*************************************************************************
Protected Function SetAttribute_ReadOnly(ByVal Path As String _
                                    ) As Boolean
'PURPOSE:
'   To Set the ReadOnly attribute of a file.
'
'OUTPUT:
'   Function returns boolean and sets the MyException Property.
'*************************************************************************
Dim Returns As Boolean

'Set the attribute
    Try
        File.SetAttributes(Path, FileAttributes.ReadOnly)
        Returns = True
        Except = ""
    Catch ex As Exception
        Except = ex.Message
        Returns = False
    End Try

'Cleanup
    Path = Nothing

'Exit
    Return Returns

End Function

#End Region
#Region "SetAttribute_System"

'*************************************************************************
Protected Function SetAttribute_System(ByVal Path As String _
                                    ) As Boolean
'PURPOSE:
'   To Set the System attribute of a file.
'
'OUTPUT:
'   Function returns boolean and sets the MyException Property.
'*************************************************************************
Dim Returns As Boolean

'Set the attribute
    Try
        File.SetAttributes(Path, FileAttributes.System)
        Returns = True
        Except = ""
    Catch ex As Exception
        Except = ex.Message
        Returns = False
    End Try

'Cleanup
    Path = Nothing

'Exit
    Return Returns

End Function

#End Region
#Region "SetAttribute_Temporary"

'*************************************************************************
Protected Function SetAttribute_Temporary(ByVal Path As String _
                                        ) As Boolean
'PURPOSE:
'   To Set the Temporary attribute of a file.
'
'OUTPUT:
'   Function returns boolean and sets the MyException Property.
'*************************************************************************
Dim Returns As Boolean

'Set the attribute
    Try
        File.SetAttributes(Path, FileAttributes.Temporary)
        Returns = True
        Except = ""
    Catch ex As Exception
        Except = ex.Message
        Returns = False
    End Try

'Cleanup
    Path = Nothing

'Exit
    Return Returns

End Function

#End Region
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
End Class