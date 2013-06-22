#Region "Comments"

'*****************************************************************************'
' PURPOSE:
'   The Win32 Registry class is the basis of this class, whose purpose is to
'   read and write to the Windows registry. Our class uses the object data
'   type to read and write the registry. The Win32 Registry class uses dynamic
'   data typing, REG_BINARY registry values would be read in as an array of
'   bytes. Writes to the registry work the same way. An array of strings would
'   be written to the registry as REG_MULTI_SZ.
'
'PREREQUISITES:
'   .Net 1.1
'
'USAGE:
'   String() = GetSubKeyNames(Hive As Hives, Key As String)
'   String() = GetSubValueNames(Hive As Hives, Key As String)
'   boolean = DeleteSubKeyTree(Hive As Hives, Key As String)
'   RegistryKey = CreateSubKey(Hive As Hives, Key As String)
'   object = GetValue(Hive As Hives, Key As String, ValueName As String)
'   boolean = DeleteValue(Hive As Hives, Key As String, ValueName As String)
'   boolean = SetValue(Hive As Hives, Key As String, ValueName As String, Value as Object)
'   string = GetRegistryType(Hive As Hives, Key As String, ValueName As String)
'
'REGISTRY DATA TYPE CONVERSION TABLE
'********************************************
'REG_BINARY     = Value() As Byte
'REG_DWORD      = Value As Integer
'REG_QWORD      = Value As UInt64
'REG_SZ         = Value As String
'REG_MULTI_SZ   = Value() As String
'REG_EXPAND_SZ  = ?
'
'AUTHOR:
'   Spencer Allen
'   2/11/08
'
' VERSIONS:
'   1.00 - Base Version (Spencer Allen)
'*****************************************************************************'

#End Region
#Region "Imports"

Imports Microsoft.Win32

#End Region
Public MustInherit Class BaseClsRegistry
#Region "Dimensions"

    Dim ObjHive As New Object
    Dim ObjKey As New Object

#Region "Enum Hives"

Enum Hives
    HKEY_LOCAL_MACHINE
    HKEY_CLASSES_ROOT
    HKEY_CURRENT_CONFIG
    HKEY_USERS
    HKEY_CURRENT_USER
    HKEY_DYN_DATA
    HKEY_PERFORMANCE_DATA
End Enum

#End Region

#End Region
#Region "GetSubKeyNames"

'*************************************************************************
Protected Function GetSubKeyNames( _
                                ByVal Hive As Hives, _
                                ByVal Key As String _
                                ) As String()
'PURPOSE:
'   Retrieves an array of strings that contains all the subkey names.
'
'OUTPUT:
'   Function returns an array of SubKeyNames or nothing if there are none
'*************************************************************************
Except = ""
Dim strSubKey() As String
Dim i As Integer = 0

'Opens the specified key and returns it's sub key names
    Try
        ObjKey = OpenSubKey(Hive, Key)

        If Not ObjKey Is Nothing Then
            'Insert SubKeyNames into an array
            ReDim strSubKey(ObjKey.SubKeyCount - 1)
            For i = 0 To ObjKey.SubKeyCount - 1
                strSubKey(i) = ObjKey.GetSubKeyNames(i)
            Next
        Else
            Except = "The key did not Exist."
        End If
    Catch ex As Exception
        Except = ": " & ex.Message
    End Try

'Cleanup
    i = Nothing
    Hive = Nothing
    Key = Nothing

'Exit
    Return strSubKey

End Function

#End Region
#Region "GetSubKeyValueNames"

'*************************************************************************
Protected Function GetSubKeyValueNames( _
                                    ByVal Hive As Hives, _
                                    ByVal Key As String _
                                    ) As String()
'PURPOSE:
'   Retrieves an array of strings that contains all the sub Value names.
'
'OUTPUT:
'   Function returns an array of SubValueNames or nothing if there are none
'*************************************************************************
Except = ""
Dim LogEvent As String = Hive & "\" & Key & " "
Dim strSubKey() As String
Dim i As Integer

'Opens the specified key and returns all of the value names it contains
    Try
        ObjKey = OpenSubKey(Hive, Key)

        If Not ObjKey Is Nothing Then
            'Insert SubKeyNames into an array
            ReDim strSubKey(ObjKey.ValueCount - 1)
            For i = 0 To ObjKey.ValueCount - 1
                strSubKey(i) = ObjKey.GetValueNames(i)
            Next
        Else
            Except = "The key did not Exist."
        End If
    Catch ex As Exception
        Except = ex.Message
    End Try

'Cleanup
    LogEvent = Nothing
    i = Nothing
    Hive = Nothing
    Key = Nothing

'Exit
    Return strSubKey

End Function

#End Region
#Region "DeleteSubKeyTree"

'*************************************************************************
Protected Function DeleteSubKeyTree( _
                                ByVal Hive As Hives, _
                                ByVal Key As String _
                                ) As Boolean
'PURPOSE:
'   Deletes the specified subkey using the DeleteSubKeyTree method
'OUTPUT:
'   Returns TRUE if SubKey was deleted or did not previously exist
'*************************************************************************
Except = ""
ObjHive = GetHive(Hive)
Dim returns As Boolean

'Delets the specified sub key and everything contained in it
    If Not ObjHive Is Nothing Then
        Try
            ObjKey = ObjHive.DeleteSubKeyTree(Key)
            returns = True
            Except = ""
        Catch ex As Exception
            Except = ex.Message
            returns = False
        End Try
    Else
        returns = False
        Except = "The key did not Exist."
    End If

'Cleanup
    ObjHive = Nothing
    Hive = Nothing
    Key = Nothing

'Exit
    Return returns

End Function

#End Region
#Region "CreateSubKey"

'*************************************************************************
Protected Function CreateSubKey( _
                            ByVal Hive As Hives, _
                            ByVal Key As String _
                            ) As RegistryKey
'PURPOSE:
'   Creates a new subkey or opens an existing subkey.
'OUTPUT:
'   Returns TRUE if SubKey was created or did previously exist
'*************************************************************************
Except = ""
Dim objNewKey As RegistryKey
Dim strNewKey As String
Dim arrKey() As String
Dim i As Integer

'Opens the specified key and creates a new sub key inside it.
    Try
        ObjKey = OpenSubKey(Hive, Key)
        If ObjKey Is Nothing Then
            ' Get the highest level subkey
            arrKey = Key.Split("\")
            objNewKey = OpenSubKey(Hive, arrKey(0))
            strNewKey = Key.Remove(0, arrKey(0).Length + 1)
            objNewKey = objNewKey.CreateSubKey(strNewKey)
        End If
    Catch ex As Exception
        Except = ex.Message
    End Try

'Cleanup
    objNewKey = Nothing
    strNewKey = Nothing
    arrKey = Nothing
    i = Nothing
    Hive = Nothing
    Key = Nothing

'Exit
    Return objNewKey

End Function

#End Region
#Region "GetValue"

'*************************************************************************
Protected Function GetValue( _
                        ByVal Hive As Hives, _
                        ByVal Key As String, _
                        ByVal ValueName As String _
                        ) As Object
'PURPOSE:
'   Retrieves the specified value.
'OUTPUT:
'   Returns the value of ValueName or a message stating why the value was not retrived.
'*************************************************************************
Except = ""
Dim objKey As RegistryKey
Dim Returns As Object

'Opens the specified the value and returns it's value
    Try
        objKey = OpenSubKey(Hive, Key)
        If Not objKey Is Nothing Then
            Returns = objKey.GetValue(ValueName, Nothing)
        Else
            Except = "The key did not Exist."
        End If
    Catch ex As Exception
        Except = ex.Message
    End Try

'Cleanup
    objKey = Nothing
    Hive = Nothing
    Key = Nothing
    ValueName = Nothing

'Exit
    Return Returns

End Function

#End Region
#Region "DeleteValue"

'*************************************************************************
Protected Function DeleteValue( _
                            ByVal Hive As Hives, _
                            ByVal Key As String, _
                            ByVal ValueName As String _
                            ) As Boolean
'PURPOSE:
'   Deletes the specified value from this key.
'
'OUTPUT:
'   Returns TRUE if value was deleted or does not exist
'*************************************************************************
Except = ""
Dim Value As Object
Dim TheString As String
Dim Returns As Boolean = True

'Opens the specified value and deletes it
    Try
        ObjKey = OpenSubKey(Hive, Key)
        Value = GetValue(Hive, Key, ValueName)
        If Not ObjKey Is Nothing Then
            Try
                TheString = Value.ToString

                If MyException = "" Then
                    ObjKey.DeleteValue(ValueName, True)
                Else
                    Returns = False
                End If
            Catch ex As Exception
                Except = ex.Message
                Returns = False
            End Try
        Else
            Returns = False
            Except = "The Key did not exist."
        End If
    Catch ex As Exception
        Except = ex.Message
    End Try

'Clean up
    TheString = Nothing
    ObjKey = Nothing
    Value = Nothing
    ValueName = Nothing
    Key = Nothing
    Hive = Nothing

'Exit
    Return Returns

End Function

#End Region
#Region "SetValue"

'*************************************************************************
Protected Function SetValue( _
                        ByVal Hive As Hives, _
                        ByVal Key As String, _
                        ByVal ValueName As String, _
                        ByVal Value As Object _
                        ) As Boolean
'PURPOSE:
'   Sets the specified value. 
'
'OUTPUT:
'   Returns the type of value created
'*************************************************************************
Except = ""
Dim objKey As Object
Dim MYvalue As Object
Dim returns As Boolean

'Opens the specified value if it exists if not it is created so the it's value can be set.
'Returns the type of value created
    Try
        objKey = OpenSubKey(Hive, Key)
        MYvalue = GetValue(Hive, Key, ValueName)
        If Not objKey Is Nothing Then
            If MYvalue Is Nothing Then
                objKey.SetValue(ValueName, Value)
                returns = True
            Else
                CreateSubKey(Hive, Key)
                objKey = OpenSubKey(Hive, Key)
                objKey.SetValue(ValueName, Value)
                returns = True
            End If
        Else
            CreateSubKey(Hive, Key)
            objKey = OpenSubKey(Hive, Key)
            objKey.SetValue(ValueName, Value)
            returns = True
        End If
    Catch ex As Exception
        returns = False
        Except = ex.Message
    End Try

'Cleanup
    objKey = Nothing
    MYvalue = Nothing
    Hive = Nothing
    Key = Nothing
    ValueName = Nothing
    Value = Nothing

'Exit
    Return returns

End Function

#End Region
#Region "GetRegistryType"

'*************************************************************************
Protected Function GetRegistryType( _
                                ByVal Hive As Hives, _
                                ByVal Key As String, _
                                ByVal ValueName As String _
                                ) As String
'PURPOSE:
'   To determine the registry data type of a registry value.
'
'OUTPUT:
'   Returns a registry string that describes registry data type.
'*************************************************************************
Except = ""
Dim LogEvent As String = Hive & "\" & Key & " "
Dim Value As Object
Dim Returns As String

'Returns the Common name for registry types
    Value = GetValue(Hive, Key, ValueName)
    If Not Value Is Nothing Then
        Try
            Select Case Value.GetType.FullName
                Case "System.Byte[]"
                    Returns = "REG_BINARY"
                Case "System.Int32"
                    Returns = "REG_DWORD"
                Case "System.UInt64"
                    Returns = "REG_QWORD"
                Case "System.String"
                    Returns = "REG_SZ"
                Case "System.String[]"
                    Returns = "REG_MULTI_SZ"
                Case Else
                    Returns = ": " & LogEvent & "Did not exist."
                    Except = "The key did not Exist."
            End Select
        Catch ex As Exception
            Except = ex.Message
            Returns = ": " & LogEvent & "Did not exist."
        End Try
    Else
        Returns = ": " & LogEvent & "Did not exist."
        Except = "The key did not Exist."
    End If

'Cleanup
    Value = Nothing
    Hive = Nothing
    Key = Nothing
    ValueName = Nothing

'Exit
    Return Returns

End Function

#End Region
#Region "OpenSubKey"

'*************************************************************************
Private Function OpenSubKey( _
                            ByVal Hive As Hives, _
                            ByVal Key As String _
                            ) As RegistryKey
'PURPOSE:
'   To create a reigistry object to the desired key passed in the parameter
'   list. As this is a private function, only critical errors will be logged.
'
'OUTPUT:
'   Returns a registry key object
'*************************************************************************
Except = ""
ObjHive = GetHive(Hive)

'Opens the specified key and returns it as an object. Private use only.
    If Not ObjHive Is Nothing Then
        Try
            ObjKey = ObjHive.OpenSubKey(Key, True)
        Catch ex As Exception
            ObjKey = Nothing
            Except = ex.Message
        End Try
    Else
        ObjKey = Nothing
    End If

'Cleanup
    ObjHive = Nothing
    Hive = Nothing
    Key = Nothing

'Exit
    Return ObjKey

End Function

#End Region
#Region "GetHive"

'*************************************************************************
Private Function GetHive( _
                        ByVal Hive As Hives _
                        ) As RegistryKey
'PURPOSE:
'   To create a reigistry object to the desired Hive.
'
'OUTPUT:
'   Returns a registry key object
'*************************************************************************

'Returns a registry object
    Select Case Hive
        Case Hives.HKEY_LOCAL_MACHINE
            ObjHive = Registry.LocalMachine
        Case Hives.HKEY_CLASSES_ROOT
            ObjHive = Registry.ClassesRoot
        Case Hives.HKEY_CURRENT_CONFIG
            ObjHive = Registry.CurrentConfig
        Case Hives.HKEY_USERS
            ObjHive = Registry.Users
        Case Hives.HKEY_CURRENT_USER
            ObjHive = Registry.CurrentUser
        Case Hives.HKEY_DYN_DATA
            ObjHive = Registry.DynData
        Case Hives.HKEY_PERFORMANCE_DATA
            ObjHive = Registry.PerformanceData
    End Select

'Cleanup
    Hive = Nothing

'Exit
    Return ObjHive

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
End Class