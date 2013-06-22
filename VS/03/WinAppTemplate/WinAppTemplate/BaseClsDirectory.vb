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
'
'USAGE:
'   boolean = Exists(Path as string)
'   boolean = Move(SourcePath as string, DestinationPath as string)
'   boolean = Delete(Path as string)
'   boolean = Create(Path as string)
'   String() = GetSubDirectorys(Path as string)
'   String() = GetFileNames(Path as string)
'   String() = GetLogicalDrives
'   String = GetParent(Path As String)
'   String = GetDirectoryRoot(Path As String)
'   String = GetCurrentDirectory
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
'AUTHOR:
'   Author
'   5/15/08
'
'VERSIONS:
'   1.00 - Base Version (Author)
'*****************************************************************************'

#End Region
#Region "Imports"

    Imports System.IO

#End Region
Public MustInherit Class BaseClsDirectory
#Region "Exists"

'*************************************************************************
Protected Function Exists(ByVal Path As String _
                ) As Boolean
'PURPOSE:
'   Find out if a Directory exists.
'
'OUTPUT:
'   Function returns boolean.
'*************************************************************************
Dim Returns As Boolean = True

'Find out if the Directory exists.
    If Not Directory.Exists(Path) Then
        Except = "The Path dose not exist."
        Returns = False
    Else
        Except = ""
    End If

'Cleanup
    Path = Nothing

'Exit
    Return Returns

End Function

#End Region
#Region "Move"

'*************************************************************************
Protected Function Move(ByVal SourcePath As String, _
                ByVal DestinationPath As String _
                ) As Boolean
'PURPOSE:
'   Move the Directory from one location to another.
'
'OUTPUT:
'   Function returns boolean.
'*************************************************************************
Dim Returns As Boolean = True

'Move the Directory
    Try
        Directory.Move(SourcePath, DestinationPath)
        Except = ""
    Catch ex As Exception
        Except = ex.Message
    End Try

'Cleanup
    SourcePath = Nothing
    DestinationPath = Nothing

'Exit
    Return Returns

End Function

#End Region
#Region "Delete"

'*************************************************************************
Protected Function Delete(ByVal Path As String _
                        ) As Boolean
'PURPOSE:
'   To delete a Directory.
'
'OUTPUT:
'   Function returns boolean.
'*************************************************************************
Dim Returns As Boolean = True

'Delete the file
    Try
        Directory.Delete(Path, True)
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
#Region "Create"

'*************************************************************************
Protected Function Create(ByVal Path As String _
                        ) As Boolean
'PURPOSE:
'   To Create a Directory or Directory Structure.
'
'OUTPUT:
'   Function returns boolean.
'*************************************************************************
Dim Returns As Boolean = True

'Create the Directory
    Try
        Directory.CreateDirectory(Path)
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
#Region "GetSubDirectorys"

'*************************************************************************
Protected Function GetSubDirectorys(ByVal Path As String _
                                ) As String()
'PURPOSE:
'   To get a list of all sub directorys.
'
'OUTPUT:
'   Function returns an array of Sub Directory Names.
'*************************************************************************
Dim Returns() As String
Dim i As Integer = 0
Dim Str As String

'Opens the specified Directory and returns it's sub directory names.
    Try
        ReDim Returns(UBound(Directory.GetDirectories(Path)))
        For Each Str In Directory.GetDirectories(Path)
            Returns(i) = Str
            i = i + 1
        Next
        Except = ""
    Catch ex As Exception
        Except = ex.Message
    End Try

'Cleanup
    i = Nothing
    Path = Nothing
    Str = Nothing

'Exit
    Return Returns

End Function

#End Region
#Region "GetFileNames"

'*************************************************************************
Protected Function GetFileNames(ByVal Path As String _
                            ) As String()
'PURPOSE:
'   To get a list of all the files in the directory.
'
'OUTPUT:
'   Function returns an array of file Names.
'*************************************************************************
Dim Returns() As String
Dim i As Integer = 0
Dim str As String

'Opens the specified Directory and returns names of all the files in it.
    Try
        ReDim Returns(UBound(Directory.GetFiles(Path)))
        For Each str In Directory.GetFiles(Path)
            Returns(i) = str
            i = i + 1
        Next
        Except = ""
    Catch ex As Exception
        Except = ex.Message
    End Try

'Cleanup
    i = Nothing
    str = Nothing
    Path = Nothing

'Exit
    Return Returns

End Function

#End Region
#Region "GetLogicalDrives"

'*************************************************************************
Protected Function GetLogicalDrives() As String()
'PURPOSE:
'   To get names of all logical drives on the current computer.
'
'OUTPUT:
'   Function returns names of all logical drives on the current computer.
'*************************************************************************
Dim Returns As String()
Dim i As Integer = 0
Dim str As String

'To get names of all logical drives on the current computer.
    Try
        ReDim Returns(UBound(Directory.GetLogicalDrives()))
        For Each str In Directory.GetLogicalDrives()
            Returns(i) = str
            i = i + 1
        Next
        Except = ""
    Catch ex As Exception
        Except = ex.Message
    End Try

'Cleanup
    i = Nothing
    str = Nothing

'Exit
    Return Returns

End Function

#End Region
#Region "GetParent"

'*************************************************************************
Protected Function GetParent(ByVal Path As String _
                            ) As String
'PURPOSE:
'   Retrieves the parent directory of the specified path, including both 
'   absolute and relative paths.
'
'OUTPUT:
'   Function returns absolute and relative paths.
'*************************************************************************
Dim Returns As String

'   To get names of all logical drives on the current computer.
    Try
        Returns = Directory.GetParent(Path).ToString
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
#Region "GetDirectoryRoot"

'*************************************************************************
Protected Function GetDirectoryRoot(ByVal Path As String _
                                ) As String
'PURPOSE:
'   Returns the volume information, root information, or both for the 
'   specified path.
'
'OUTPUT:
'   Function returns volume information, root information, or both for the 
'   specified path.
'*************************************************************************
Dim Returns As String

'Opens the specified Directory and returns names of all the files in it.
    Try
        Returns = Directory.GetDirectoryRoot(Path)
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
#Region "GetCurrentDirectory"

'*************************************************************************
Protected Function GetCurrentDirectory() As String
'PURPOSE:
'   Gets the current working directory of the application.
'
'OUTPUT:
'   Function returns the current working directory of the application.
'*************************************************************************
Dim Returns As String

'Get the current working directory of the application
    Try
        Returns = Directory.GetCurrentDirectory()
        Except = ""
    Catch ex As Exception
        Except = ex.Message
    End Try

'Cleanup


'Exit
    Return Returns

End Function

#End Region
#Region "GetFileSystemEntries"

'*************************************************************************
Protected Function GetFileSystemEntries(ByVal Path As String _
                                    ) As String()
'PURPOSE:
'   Returns the names of all files and subdirectories in the specified directory.
'
'OUTPUT:
'   Function returns the names of all files and subdirectories in the 
'   specified directory.
'*************************************************************************
Dim Returns As String()
Dim i As Integer = 0
Dim str As String

'To get names of all files and subdirectories in the specified directory.
    Try
        ReDim Returns(UBound(Directory.GetFileSystemEntries(Path)))
        For Each str In Directory.GetFileSystemEntries(Path)
            Returns(i) = str
            i = i + 1
        Next
        Except = ""
    Catch ex As Exception
        Except = ex.Message
    End Try

'Cleanup
    Path = Nothing
    i = Nothing
    str = Nothing

'Exit
    Return Returns

End Function

#End Region
#Region "GetAttributes"
#Region "GetLastWriteTime"

'*************************************************************************
Protected Function GetLastWriteTime(ByVal Path As String _
                                ) As String
'PURPOSE:
'   To get the last time the specified file was written to.
'
'OUTPUT:
'   Function returns a datetime object or a object with the standard ": " exception message.
'*************************************************************************
Dim Returns As Object

'get the last time the specified file was written to
    Try
        Returns = Directory.GetLastWriteTime(Path).ToString
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
#Region "GetLastWriteTimeUtc"

'*************************************************************************
Protected Function GetLastWriteTimeUtc(ByVal Path As String _
                                    ) As String
'PURPOSE:
'   To get the last time the specified file was written to. This value is expressed in UTC time.

'
'OUTPUT:
'   Function returns a datetime object or a object with the standard ": " exception message.
'*************************************************************************
Dim Returns As Object

'get the last time the specified file was written to
    Try
        Returns = Directory.GetLastWriteTimeUtc(Path).ToString
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
#Region "GetCreationTime"

'*************************************************************************
Protected Function GetCreationTime(ByVal Path As String _
                                ) As String
'PURPOSE:
'   To get the creation date and time of a directory.
'
'OUTPUT:
'   Function returns a DateTime object or an object with the standard ": " 
'   exception message.
'*************************************************************************
Dim Returns As String

'Get the directory creation time
    Try
        Returns = Directory.GetCreationTime(Path).ToString
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
#Region "GetCreationTimeUtc"

'*************************************************************************
Protected Function GetCreationTimeUtc(ByVal Path As String _
                                    ) As String
'PURPOSE:
'   To get the creation date and time of a directory in coordinated universal time (UTC) format.
'
'OUTPUT:
'   Function returns a DateTime object expressed in UTC time or an object with the standard ": " exception message.
'*************************************************************************
Dim Returns As String

'Get the directory creation time in UTC format
    Try
        Returns = Directory.GetCreationTimeUtc(Path).ToString
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
#Region "GetLastAccessTime"

'*************************************************************************
Protected Function GetLastAccessTime(ByVal Path As String _
                                ) As String
'PURPOSE:
'   To get the date and time the specified file or directory was last accessed.
'
'OUTPUT:
'   Function returns date and time the specified file or directory was last accessed or an object with the standard ": " exception message.
'*************************************************************************
Dim Returns As String

'get date and time the specified file or directory was last accessed
    Try
        Returns = Directory.GetLastAccessTime(Path).ToString
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
#Region "GetLastAccessTimeUtc"

'*************************************************************************
Protected Function GetLastAccessTimeUtc(ByVal Path As String _
                                    ) As String
'PURPOSE:
'   To get the last time the specified file was written to. This value is expressed in UTC time.

'
'OUTPUT:
'   Function returns a datetime object or a object with the standard ": " exception message.
'*************************************************************************
Dim Returns As String

'get the last time the specified file was written to
    Try
        Returns = Directory.GetLastAccessTimeUtc(Path).ToString
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
#End Region
#Region "SetAttributes"
#Region "SetLastWriteTime"

'*************************************************************************
Protected Function SetLastWriteTime(ByVal Path As String, _
                                ByVal LastAccessTime As String _
                                ) As Boolean
'PURPOSE:
'   Sets the date and time a directory was last written to.
'
'OUTPUT:
'   Function returns The string "True" if suscessful or ": exception message" 
'   if an exception was thrown.
'*************************************************************************
Dim Returns As Boolean

'Sets current directory
    Try
        Directory.SetLastWriteTime(Path, LastAccessTime)
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
#Region "SetLastWriteTimeUtc"

'*************************************************************************
Protected Function SetLastWriteTimeUtc(ByVal Path As String, _
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
Dim Returns As Boolean

'Sets current directory
    Try
        Directory.SetLastWriteTimeUtc(Path, LastAccessTime)
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
#Region "SetCreationTime"

'*************************************************************************
Protected Function SetCreationTime(ByVal Path As String, _
                                ByVal CreationTime As String _
                                ) As Boolean
'PURPOSE:
'   To set the time the file was created to the current date and time.

'
'OUTPUT:
'   Function returns an "True" string or a ": " exception message.
'*************************************************************************
Dim Returns As Boolean


'set the time the file was created to the current date and time
    Try
        Directory.SetCreationTime(Path, CreationTime)
        Except = ""
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
                                    ByVal CreationTime As String _
                                    ) As Boolean
'PURPOSE:
'   To set the time the file was created to the current date and time. This value is expressed in UTC time.

'
'OUTPUT:
'   Function returns an "True" string or a ": " exception message.
'*************************************************************************
Dim Returns As Boolean


'set the time the file was created to the current date and time
    Try
        Directory.SetCreationTimeUtc(Path, CreationTime)
        Returns = True
        Except = ""
    Catch ex As Exception
        Except = ex.Message
        Returns = False
    End Try

'Cleanup
    Path = Nothing
    CreationTime = Nothing

'Exit
    Return Returns

End Function

#End Region
#Region "SetLastAccessTime"

'*************************************************************************
Protected Function SetLastAccessTime(ByVal Path As String, _
                                ByVal LastAccessTime As String _
                                ) As Boolean
'PURPOSE:
'   Sets the date and time the specified file or directory was last accessed.
'
'OUTPUT:
'   Function returns The string "True" if suscessful or ": exception message" 
'   if an exception was thrown.
'*************************************************************************
Dim Returns As Boolean

'Sets current directory
    Try
        Directory.SetLastAccessTime(Path, LastAccessTime)
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
#Region "SetLastAccessTimeUtc"

'*************************************************************************
Protected Function SetLastAccessTimeUtc(ByVal Path As String, _
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
Dim Returns As Boolean

'Sets current directory
    Try
        Directory.SetLastAccessTimeUtc(Path, LastAccessTime)
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