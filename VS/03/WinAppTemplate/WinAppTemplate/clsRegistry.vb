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
'   "Public LogFilePath As String" must be in the main program's global declarations.
'   "Public Log As New clsLogFile" must be in the main program's global declarations.
'
'USAGE:
'   String() = GetSubKeyNames(Hive As String, Key As String)
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
'   Shayne Marriage
'   BTS Applications
'   10-20-2007
'
' VERSIONS:
'   1.00 - Base Version (Shayne Marriage)
'   1.01 - Added error trapping for all members (Spencer Allen)
'   1.02 - Rewrites GetSubKeyNames to return an array of subkey names (Shayne Marriage)
'   1.03 - Renames GetRegistryHiveObject to OpenSubKey and deletes the original OpenSubKey member (Shayne Marriage)
'   1.04 - Renames DeleteSubKey to DeleteSubKeyTree and rewrites the member completely (Shayne Marriage)
'   1.05 - Rewrites CreateSubKey (Shayne Marriage)
'   1.06 - Rewrites GetValue (Spencer Allen)
'   1.07 - Added Value type handeling in logs for most of member SetValue (SpencerAllen)
'   1.08 - adds Value parameter and logs KEY not present (Shayne Marriage)
'   1.09 - adds GetRegistryType (Spencer Allen)
'   1.10 - completed SetRegistryValue (Spencer Allen)
'   1.11 - completes DeleteValue (Shayne Marriage)
'   1.12 - traps exception in GetValueType if GetValue returns nothing (Shayne Marriage)
'   1.13 - Adds GetSubKeyValueNames(Spencer Allen)
'*****************************************************************************'

Imports Microsoft.Win32
'Imports Microsoft.Win32.RegistryHive
'Imports Microsoft.Win32.RegistryKey

Public Class clsRegistry
    Dim Logfile As New clsLogFile    'Log file class for using log files

'*************************************************************************
    Public Function GetSubKeyNames( _
                             ByVal Hive As String, _
                             ByVal Key As String _
                            ) As String()
'PURPOSE:
'   Retrieves an array of strings that contains all the subkey names.
'
'OUTPUT:
'   Function returns an array of SubKeyNames or nothing if there are none
'*************************************************************************
        Dim objKey As RegistryKey
        Dim strSubKey() As String
        Dim logEvent As String = "Registry.GetSubKeyNames: " & Hive & "\" & Key & " "
        Dim noErrorFlag = True
        Dim i As Integer

        Try
            objKey = OpenSubKey(Hive, Key)

            If Not objKey Is Nothing Then
                'Insert SubKeyNames into an array
                ReDim strSubKey(objKey.SubKeyCount - 1)
                For i = 0 To objKey.SubKeyCount - 1
                    strSubKey(i) = objKey.GetSubKeyNames(i)
                Next
            Else
                strSubKey = Nothing
            End If
        Catch ex As Exception
            noErrorFlag = False
            logEvent = "Failure:" & logEvent & ex.ToString & ")"
        End Try

        If noErrorFlag Then
            If objKey Is Nothing Then
                logEvent = "Information: " & logEvent & " does not exist"
            Else
                logEvent = "Success: " & logEvent & " returned " & strSubKey.Length & " SubKeyNames"
            End If
        End If

        Logger.addEvent(LogFilePath, logEvent)
        Return strSubKey
    End Function


    '*************************************************************************
    Public Function GetSubKeyValueNames( _
                             ByVal Hive As String, _
                             ByVal Key As String _
                            ) As String()
        'PURPOSE:
        '   Retrieves an array of strings that contains all the sub Value names.
        '
        'OUTPUT:
        '   Function returns an array of SubValueNames or nothing if there are none
        '*************************************************************************
        Dim objKey As RegistryKey
        Dim strSubKey() As String
        Dim logEvent As String = "Registry.GetSubValueNames: " & Hive & "\" & Key & " "
        Dim noErrorFlag = True
        Dim i As Integer

        Try
            objKey = OpenSubKey(Hive, Key)

            If Not objKey Is Nothing Then
                'Insert SubKeyNames into an array
                ReDim strSubKey(objKey.ValueCount - 1)
                For i = 0 To objKey.ValueCount - 1
                    strSubKey(i) = objKey.GetValueNames(i)
                Next
            Else
                strSubKey = Nothing
            End If
        Catch ex As Exception
            noErrorFlag = False
            logEvent = "Failure:" & logEvent & ex.ToString & ")"
        End Try

        If noErrorFlag Then
            If objKey Is Nothing Then
                logEvent = "Information: " & logEvent & " does not exist"
            Else
                logEvent = "Success: " & logEvent & " returned " & strSubKey.Length & " SubValueNames"
            End If
        End If

        Logger.addEvent(LogFilePath, logEvent)
        Return strSubKey
    End Function

    '*************************************************************************
    Public Function DeleteSubKeyTree( _
                             ByVal Hive As String, _
                             ByVal Key As String _
                            ) As Object
        'PURPOSE:
        '   Deletes the specified subkey using the DeleteSubKeyTree method
        'OUTPUT:
        '   Returns TRUE if SubKey was deleted or did not previously exist
        '*************************************************************************
        Dim objKey As RegistryKey
        Dim logEvent As String = "Registry.DeleteSubKeyTree: " & Hive & "\" & Key
        Dim noErrorFlag = True

        objKey = OpenSubKey(Hive, Key)
        If Not objKey Is Nothing Then
            Try
                Select Case Hive
                    Case "HKLM"
                        Registry.LocalMachine.DeleteSubKeyTree(Key)
                    Case "HKCR"
                        Registry.ClassesRoot.DeleteSubKeyTree(Key)
                    Case "HKCC"
                        Registry.CurrentConfig.DeleteSubKeyTree(Key)
                    Case "HKU"
                        Registry.Users.DeleteSubKeyTree(Key)
                    Case "HKCU"
                        Registry.CurrentUser.DeleteSubKeyTree(Key)
                    Case Else
                        noErrorFlag = False
                End Select
            Catch ex As Exception
                noErrorFlag = False
                logEvent = "Failure: " & logEvent & ex.ToString
            End Try
        End If

        If noErrorFlag Then
            logEvent = "Success: " & logEvent
        Else
            logEvent = "Information: " & logEvent & " does not exist"
        End If

        Logger.addEvent(LogFilePath, logEvent)
        Return noErrorFlag
    End Function

    '*************************************************************************
    Public Function CreateSubKey( _
                             ByVal Hive As String, _
                             ByVal Key As String _
                            ) As Boolean
        'PURPOSE:
        '   Creates a new subkey or opens an existing subkey.
        'OUTPUT:
        '   Returns TRUE if SubKey was created or did previously exist
        '*************************************************************************
        Dim objKey As RegistryKey
        Dim objNewKey As RegistryKey
        Dim strNewKey As String
        Dim arrKey() As String
        Dim i As Integer
        Dim logEvent As String = "Registry.CreateSubKey: " & Hive & "\" & Key
        Dim noErrorFlag = True

        Try
            objKey = OpenSubKey(Hive, Key)
            If objKey Is Nothing Then
                ' Get the highest level subkey
                arrKey = Key.Split("\")
                objNewKey = OpenSubKey(Hive, arrKey(0))
                strNewKey = Key.Remove(0, arrKey(0).Length + 1)
                objNewKey = objNewKey.CreateSubKey(strNewKey)
            End If
        Catch ex As Exception
            noErrorFlag = False
            logEvent = logEvent & "(" & ex.ToString & ")"
        End Try

        If noErrorFlag Then
            If objKey Is Nothing Then
                If Not objNewKey Is Nothing Then
                    logEvent = "Success: " & logEvent
                End If
            Else
                logEvent = "Information: " & logEvent & " already exists"
            End If
        End If

        Logger.addEvent(LogFilePath, logEvent)
        Return noErrorFlag
    End Function


    '*************************************************************************
    Public Function GetValue( _
                             ByVal Hive As String, _
                             ByVal Key As String, _
                             ByVal ValueName As String _
                            ) As Object
        'PURPOSE:
        '   Retrieves the specified value.
        'OUTPUT:
        '   Returns the value of ValueName or nothing if the value not present
        '*************************************************************************
        Dim objKey As RegistryKey
        Dim success As Boolean = True
        Dim Value As Object
        Dim logEvent As String = "Registry.GetValue: " & Hive & "\" & Key
        Dim noErrorFlag = True

        Try
            objKey = OpenSubKey(Hive, Key)
            If Not objKey Is Nothing Then
                Value = objKey.GetValue(ValueName, Nothing)
            Else
                logEvent = "Information: " & logEvent & " does not exist."
            End If
        Catch ex As Exception
            noErrorFlag = False
            logEvent = "Failure: Registry - GetValue: (" & ex.ToString & ")"
        End Try

        If noErrorFlag And Not objKey Is Nothing Then
            logEvent = "Success: " & logEvent
        End If

        Logger.addEvent(LogFilePath, logEvent)
        Return Value
    End Function

    '*************************************************************************
    Public Function DeleteValue( _
                                ByVal Hive As String, _
                                ByVal Key As String, _
                                ByVal ValueName As String _
                                ) As Boolean
        'PURPOSE:
        '   Deletes the specified value from this key.
        '
        'OUTPUT:
        '   Returns TRUE if value was deleted or does not exist
        '*************************************************************************
        Dim objKey As Object
        Dim Value As Object
        Dim logEvent As String = "Registry.DeleteValue: " & Hive & "\" & Key & "\" & ValueName
        Dim noErrorFlag = True

        Try
            objKey = OpenSubKey(Hive, Key)
            Value = GetValue(Hive, Key, ValueName)
            If Not objKey Is Nothing Then
                If Not Value Is Nothing Then
                    objKey.DeleteValue(ValueName, True)
                End If
            End If
        Catch ex As Exception
            noErrorFlag = False
            logEvent = "Failure: " & logEvent & " (" & ex.ToString & ")"
        End Try

        If noErrorFlag Then
            If objKey Is Nothing Or Value Is Nothing Then
                logEvent = "Information: " & logEvent & " does not exist"
            Else
                logEvent = "Success: " & logEvent
            End If
        End If

        Logger.addEvent(LogFilePath, logEvent)
        Return noErrorFlag
    End Function

    '*************************************************************************
    Public Function SetValue( _
                             ByVal Hive As String, _
                             ByVal Key As String, _
                             ByVal ValueName As String, _
                             ByVal Value As Object _
                            ) As Boolean
        'PURPOSE:
        '   Sets the specified value. 
        '
        'OUTPUT:
        '   Returns TRUE if value was set
        '*************************************************************************
        Dim objKey As Object
        Dim ValType As String
        Dim logEvent As String = "Registry.SetValue: " & Hive & "\" & Key & "\" & ValueName
        Dim noErrorFlag = True

        Try
            objKey = OpenSubKey(Hive, Key)
            If Not objKey Is Nothing Then
            Else
                CreateSubKey(Hive, Key)
                objKey = OpenSubKey(Hive, Key)
            End If
            If Not Value Is Nothing Then
                objKey.SetValue(ValueName, Value)
            Else
                noErrorFlag = False
                logEvent = "Skipped: " & logEvent & " value cannot be null"
            End If

        Catch ex As Exception
            noErrorFlag = False
            logEvent = "Failure: " & logEvent & " (" & ex.ToString & ")"
        End Try

        If noErrorFlag Then
            logEvent = "Success: " & logEvent & " written as " & GetRegistryType(Hive, Key, ValueName)
        End If

        Logger.addEvent(LogFilePath, logEvent)
        Return noErrorFlag
    End Function

    '*************************************************************************
    Private Function OpenSubKey( _
                                        ByVal Hive As String, _
                                        ByVal Key As String _
                                          ) As RegistryKey
        'PURPOSE:
        '   To create a reigistry object to the desired key passed in the parameter
        '   list. As this is a private function, only critical errors will be logged.
        '
        'OUTPUT:
        '   Returns a registry key object
        '*************************************************************************
        Dim logEvent As String = "Registry.OpenSubKey: " & Hive & "\" & Key
        Dim noErrorFlag = True                  'return value for the function, returns false if an error occurs

        Dim objKey As Object
        Try
            Select Case Hive
                Case "HKLM"
                    objKey = Registry.LocalMachine.OpenSubKey(Key, True)
                Case "HKCR"
                    objKey = Registry.ClassesRoot.OpenSubKey(Key, True)
                Case "HKCC"
                    objKey = Registry.CurrentConfig.OpenSubKey(Key, True)
                Case "HKU"
                    objKey = Registry.Users.OpenSubKey(Key, True)
                Case "HKCU"
                    objKey = Registry.CurrentUser.OpenSubKey(Key, True)
                Case Else
                    noErrorFlag = False
                    Logger.addEvent(LogFilePath, logEvent = "Failure: " & logEvent & ", " & Hive & " is not a valid Hive Type")
            End Select
        Catch ex As Exception
            noErrorFlag = False
            Logger.addEvent(LogFilePath, "Failure: " & logEvent & "(" & ex.ToString & ")")
        End Try

        Return objKey
    End Function


    '*************************************************************************
    Public Function GetRegistryType( _
                             ByVal Hive As String, _
                             ByVal Key As String, _
                             ByVal ValueName As String _
                                    ) As String
        'PURPOSE:
        '   To determine the registry data type of a registry value.
        '
        'OUTPUT:
        '   Returns a registry string that describes registry data type.
        '*************************************************************************
        Dim objKey As Object
        Dim Value As Object
        Dim ValType As String
        Dim logEvent As String = "Registry.GetRegistryType: " & Hive & "\" & Key & "\" & ValueName
        Dim noErrorFlag = True

        Value = GetValue(Hive, Key, ValueName)
        If Not Value Is Nothing Then
            Try
                Select Case Value.GetType.FullName
                    Case "System.Byte[]"
                        ValType = "REG_BINARY"
                    Case "System.Int32"
                        ValType = "REG_DWORD"
                    Case "System.UInt64"
                        ValType = "REG_QWORD"
                    Case "System.String"
                        ValType = "REG_SZ"
                    Case "System.String[]"
                        ValType = "REG_MULTI_SZ"
                    Case Else
                        noErrorFlag = False
                        logEvent = "Failure: " & logEvent & " unable to determine registry data type"
                End Select
            Catch ex As Exception
                noErrorFlag = False
                logEvent = "Failure: " & logEvent & " (" & ex.ToString & ")"
            End Try
        Else
            ValType = Nothing
        End If

        Return ValType
    End Function
End Class

