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
'   clsLogFile must be added to the project.
'   BaseRegistry must be added to the project.
'   "Public LogFilePath As String" must be in the main program's global declarations.
'   "Public Log As New clsLogFile" must be in the main program's global declarations.
'
'USAGE:
'   String() = GetSubKeyNames(Hive As String, Key As String, Logging as boolean)
'   String() = GetSubValueNames(Hive As String, Key As String)
'   boolean = DeleteSubKeyTree(Hive As String, Key As String)
'   boolean = CreateSubKey(Hive As String, Key As String)
'   object = GetValue(Hive As String, Key As String, ValueName As String)
'   boolean = DeleteValue(Hive As String, Key As String, ValueName As String)
'   boolean = SetValue(Hive As String, Key As String, ValueName As String, Value as Object)
'   string = GetRegistryType(Hive As String, Key As String, ValueName As String)
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
'   Dts Applications
'   2/11/08
'
' VERSIONS:
'   1.00 - Base Version (Spencer Allen)
'*****************************************************************************'

#End Region
Public Class MyRegistry
#Region "Inherits"

Inherits BaseRegistry

#End Region
#Region "GetSubKeyNames"

'*************************************************************************
Public Overloads Function GetSubKeyNames( _
                                        ByVal Hive As String, _
                                        ByVal Key As String, _
                                        ByVal Logging As Boolean _
                                        ) As String()
'PURPOSE:
'   Retrieves an array of strings that contains all the subkey names.
'
'OUTPUT:
'   Function returns an array of SubKeyNames or failure message if there are none
'*************************************************************************
Dim LogEvent As String = "Registry.GetSubKeyNames: " & Hive & "\" & Key & " "
Dim i As Integer = 0
Dim SubKeyNames As String() = GetSubKeyNames(Hive, Key)

'If program calls for logging then log if not return Key Names
    If Logging Then
        If Not SubKeyNames.IndexOf(SubKeyNames, (i)).ToString.StartsWith(":") Then
            LogEvent = "Success: " & LogEvent & "returned " & SubKeyNames.Length & " SubKeyNames."
            main.Logger.addEvent(main.LogFilePath, LogEvent)
        Else
            LogEvent = "Failure: " & LogEvent & "returned " & SubKeyNames.Length & " SubKeyNames." _
            & vbCrLf & SubKeyNames.IndexOf(SubKeyNames, (i)).ToString
            main.Logger.addEvent(main.LogFilePath, LogEvent)
        End If
    End If

'Cleanup
    Hive = Nothing
    Key = Nothing
    Logging = Nothing
    LogEvent = Nothing
    i = Nothing

'Exit
    Return SubKeyNames

End Function

#End Region
#Region "GetSubKeyValueNames"

'*************************************************************************
Public Overloads Function GetSubKeyValueNames( _
                                                ByVal Hive As String, _
                                                ByVal Key As String, _
                                                ByVal Logging As Boolean _
                                                ) As String()
'PURPOSE:
'   Retrieves an array of strings that contains all the sub Value names.
'
'OUTPUT:
'   Function returns an array of SubValueNames or failure message if there are none
'*************************************************************************
Dim LogEvent As String = "Registry.GetSubKeyValueNames: " & Hive & "\" & Key & " "
Dim i As Integer = 0
Dim SubKeyValueNames As String() = GetSubKeyValueNames(Hive, Key)

'If program calls for logging then log if not return Key Names
    If Logging Then
        If Not SubKeyValueNames.IndexOf(SubKeyValueNames, (i)).ToString.StartsWith(":") Then
            LogEvent = "Success: " & LogEvent & "returned " & SubKeyValueNames.Length & " SubKeyValueNames."
            main.Logger.addEvent(main.LogFilePath, LogEvent)
        Else
            'It failed.
            LogEvent = "Failure: " & LogEvent & "returned " & SubKeyValueNames.Length & " SubKeyValueNames." _
            & vbCrLf & SubKeyValueNames.IndexOf(SubKeyValueNames, (i)).ToString
            main.Logger.addEvent(main.LogFilePath, LogEvent)
        End If
    End If

'Cleanup
    Hive = Nothing
    Key = Nothing
    Logging = Nothing
    LogEvent = Nothing
    i = Nothing

'Exit
    Return SubKeyValueNames

End Function

#End Region
End Class
