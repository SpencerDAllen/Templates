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
'   BaseRegistry must be added to the project.
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
Public Class MyClsRegistry
#Region "Inherits"

Inherits BaseClsRegistry
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
#Region "GetSubKeyNames"

'*************************************************************************
Public Overloads Function GetSubKeyNames( _
                                        ByVal Hive As Hives, _
                                        ByVal Key As String _
                                        ) As String()
'PURPOSE:
'   Retrieves an array of strings that contains all the subkey names.
'
'OUTPUT:
'   Function returns an array of SubKeyNames or failure message if there are none
'*************************************************************************
Dim LogEvent As String = "Registry.GetSubKeyNames: " & Hive.ToString & "\" & Key & " "
Dim i As Integer = 0
Dim str As String
Dim SubKeyNames As String() = MyBase.GetSubKeyNames(Hive, Key)

'If program calls for logging then log if not return Key Names
    If LogsOnOff Then
        If MyException = "" Then
            LogEvent = "Success: " & LogEvent & "returned " & SubKeyNames.Length & " SubKeyNames."
            Logs(LogEvent)
        Else
            LogEvent = "Failure: " & LogEvent & "returned " & vbCrLf & MyException
            Logs(LogEvent)
        End If
    End If

'Cleanup
    Hive = Nothing
    Key = Nothing
    LogEvent = Nothing
    i = Nothing

'Exit
    Return SubKeyNames

End Function

#End Region
#Region "GetSubKeyValueNames"

'*************************************************************************
Public Overloads Function GetSubKeyValueNames( _
                                                ByVal Hive As Hives, _
                                                ByVal Key As String _
                                                ) As String()
'PURPOSE:
'   Retrieves an array of strings that contains all the sub Value names.
'
'OUTPUT:
'   Function returns an array of SubValueNames or failure message if there are none
'*************************************************************************
Dim LogEvent As String = "Registry.GetSubKeyValueNames: " & Hive.ToString & Key & " "
Dim i As Integer = 0
Dim SubKeyValueNames As String() = MyBase.GetSubKeyValueNames(Hive, Key)

'If program calls for logging then log if not return Key Names
    If LogsOnOff Then
        If MyException = "" Then
            LogEvent = "Success: " & LogEvent & "returned " & SubKeyValueNames.Length & " SubKeyValueNames."
            Logs(LogEvent)
        Else
            'It failed.
            LogEvent = "Failure: " & LogEvent & "returned " & SubKeyValueNames.Length & " SubKeyValueNames." _
            & vbCrLf & SubKeyValueNames.IndexOf(SubKeyValueNames, (i)).ToString
            Logs(LogEvent)
        End If
    End If

'Cleanup
    Hive = Nothing
    Key = Nothing
    LogEvent = Nothing
    i = Nothing

'Exit
    Return SubKeyValueNames

End Function

#End Region
#Region "DeleteSubKeyTree"

'*************************************************************************
Public Overloads Function DeleteSubKeyTree( _
                                            ByVal Hive As Hives, _
                                            ByVal Key As String _
                                            ) As Boolean
'PURPOSE:
'   Deletes the specified subkey using the DeleteSubKeyTree method
'OUTPUT:
'   Returns "TRUE" if SubKey was deleted or did not previously exist
'*************************************************************************
Dim LogEvent As String = "Registry.DeleteSubKeyTree: " & Hive.ToString & Key & " "
Dim Returned As Boolean = MyBase.DeleteSubKeyTree(Hive, Key)

'If program calls for logging then log.
    If LogsOnOff Then
        If MyException = "" Then
            LogEvent = "Success: " & LogEvent & "Deleted " & Hive & "\" & Key
            Logs(LogEvent)
        Else
            'It Failed
            LogEvent = "Failure: " & LogEvent & vbCrLf & MyException
            Logs(LogEvent)
        End If
    End If

'Cleanup
    Hive = Nothing
    Key = Nothing
    LogEvent = Nothing

'Exit
    Return Returned

End Function

#End Region
#Region "CreateSubKey"

'*************************************************************************
Public Overloads Function CreateSubKey( _
                                        ByVal Hive As Hives, _
                                        ByVal Key As String _
                                        ) As Microsoft.Win32.RegistryKey
'PURPOSE:
'   Creates a new subkey or opens an existing subkey.
'OUTPUT:
'   Returns TRUE if SubKey was created or did previously exist
'*************************************************************************
Dim LogEvent As String = "Registry.CreateSubKey: " & Hive.ToString & Key & " "
Dim Returned As Microsoft.Win32.RegistryKey = MyBase.CreateSubKey(Hive, Key)

'If program calls for logging then log.
    If LogsOnOff Then
        If MyException = "" Then
            LogEvent = "Success: " & LogEvent & "Created " & Hive & "\" & Key
            Logs(LogEvent)
        Else
            'It Failed
            LogEvent = "Failure: " & LogEvent & vbCrLf & MyException
            Logs(LogEvent)
        End If
    End If

'Cleanup
    LogEvent = Nothing
    Hive = Nothing
    Key = Nothing

'Exit
    Return Returned

End Function

#End Region
#Region "GetValue"

'*************************************************************************
Public Overloads Function GetValue( _
                                    ByVal Hive As Hives, _
                                    ByVal Key As String, _
                                    ByVal ValueName As String _
                                    ) As Object
'PURPOSE:
'   Retrieves the specified value.
'OUTPUT:
'   Returns the value of ValueName or nothing if the value not present
'*************************************************************************
Dim LogEvent As String = "Registry.GetValue: " & Hive.ToString & Key & "\" & ValueName
Dim Returned As Object = MyBase.GetValue(Hive, Key, ValueName)
Dim TheString As String
Dim TheType As String

'If program calls for logging then log.
    If LogsOnOff Then
        If MyException = "" Then
            LogEvent = "Success: " & LogEvent
            Logs(LogEvent)
        Else
            'It Failed
            LogEvent = "Failure: " & LogEvent & vbCrLf & MyException
            Logs(LogEvent)
        End If
    End If

'Cleanup
    LogEvent = Nothing
    TheString = Nothing
    TheType = Nothing
    Hive = Nothing
    Key = Nothing
    ValueName = Nothing

'exit
    Return Returned

End Function

#End Region
#Region "DeleteValue"

'*************************************************************************
Public Overloads Function DeleteValue( _
                                    ByVal Hive As Hives, _
                                    ByVal Key As String, _
                                    ByVal ValueName As String _
                                    ) As Boolean
'PURPOSE:
'   Deletes the specified value from this key.
'
'OUTPUT:
'   Returns TRUE if value was deleted or a message stating why
'*************************************************************************
Dim LogEvent As String = "Registry.DeleteValue: " & Hive.ToString & Key & "\" & ValueName & " "
Dim Returned As Boolean = MyBase.DeleteValue(Hive, Key, ValueName)

'If program calls for logging then log.
    If LogsOnOff Then
        If MyException = "" Then
            LogEvent = "Success: " & LogEvent & "Was Deleted."
            Logs(LogEvent)
        Else
            'It Failed
            LogEvent = "Failure: " & LogEvent & vbCrLf & MyException
            Logs(LogEvent)
        End If
    End If

'Cleanup
    LogEvent = Nothing
    Hive = Nothing
    Key = Nothing
    ValueName = Nothing

'Exit
    Return Returned

End Function

#End Region
#Region "SetValue"

'*************************************************************************
Public Overloads Function SetValue( _
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
Dim LogEvent As String = "Registry.SetValue: " & Hive.ToString & Key & "\" & ValueName & " "
Dim Returned As Boolean = MyBase.SetValue(Hive, Key, ValueName, Value)

'If program calls for logging then log.
    If LogsOnOff Then
        If MyException = "" Then
            LogEvent = "Success: " & LogEvent
            Logs(LogEvent)
        Else
            'It Failed
            LogEvent = "Failure: " & LogEvent & vbCrLf & MyException
            Logs(LogEvent)
        End If
    End If

'Cleanup
    LogEvent = Nothing
    Hive = Nothing
    Key = Nothing
    Value = Nothing
    ValueName = Nothing

'exit
    Return Returned

End Function

#End Region
#Region "GetRegistryType"

'*************************************************************************
Public Overloads Function GetRegistryType( _
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
Dim LogEvent As String = "Registry.GetRegistryType: " & Hive.ToString & Key & "\" & ValueName & " "
Dim Returned As String = MyBase.GetRegistryType(Hive, Key, ValueName)

'If program calls for logging then log.
    If LogsOnOff Then
        If MyException = "" Then
            LogEvent = "Success: " & LogEvent & " returned " & Returned
            Logs(LogEvent)
        Else
            'It Failed
            LogEvent = "Failure: " & LogEvent & vbCrLf & MyException
            Logs(LogEvent)
        End If
    End If

'Cleanup
    LogEvent = Nothing
    Hive = Nothing
    Key = Nothing
    ValueName = Nothing

'exit
    Return Returned

End Function

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