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
'   Spencer Allen
'   Dts Applications
'   2/11/08
'
' VERSIONS:
'   1.00 - Base Version (Spencer Allen)
'*****************************************************************************'

#End Region
#Region "Imports"

Imports Microsoft.Win32

#End Region
Public Class BaseRegistry
#Region "Dimensions"

Dim ObjHive As New Object
Dim ObjKey As New Object

#End Region
#Region "GetSubKeyNames"

'*************************************************************************
Friend Function GetSubKeyNames( _
                                ByVal Hive As String, _
                                ByVal Key As String _
                                ) As String()
'PURPOSE:
'   Retrieves an array of strings that contains all the subkey names.
'
'OUTPUT:
'   Function returns an array of SubKeyNames or nothing if there are none
'*************************************************************************
Dim LogEvent As String = Hive & "\" & Key & " "
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
            strSubKey(i) = ": " & LogEvent & "Did not exist."
        End If
    Catch ex As Exception
        i = 0
        strSubKey(i) = ": " & ex.Message
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
#Region "GetSubKeyValueNames"

'*************************************************************************
Friend Function GetSubKeyValueNames( _
                                    ByVal Hive As String, _
                                    ByVal Key As String _
                                    ) As String()
'PURPOSE:
'   Retrieves an array of strings that contains all the sub Value names.
'
'OUTPUT:
'   Function returns an array of SubValueNames or nothing if there are none
'*************************************************************************
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
            i = 0
            strSubKey(i) = ": " & LogEvent & "Did not exist."
        End If
    Catch ex As Exception
        i = 0
        strSubKey(i) = ": " & ex.Message
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
Public Function DeleteSubKeyTree( _
								ByVal Hive As String, _
								ByVal Key As String _
								) As string
'PURPOSE:
'   Deletes the specified subkey using the DeleteSubKeyTree method
'OUTPUT:
'   Returns TRUE if SubKey was deleted or did not previously exist
'*************************************************************************
ObjHive = GetHive(Hive)
Dim returns As String = "True"

'Delets the specified sub key and everything contained in it
	If Not ObjHive Is Nothing Then
		Try
            ObjKey = ObjHive.DeleteSubKeyTree(Key)
            returns = "True"
		Catch ex As Exception
			returns = "Failure: " & ex.Message
		End Try
	Else
		returns = "Failure: The key did not exist."
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
Public Function CreateSubKey( _
							ByVal Hive As String, _
							ByVal Key As String _
							) As string
'PURPOSE:
'   Creates a new subkey or opens an existing subkey.
'OUTPUT:
'   Returns TRUE if SubKey was created or did previously exist
'*************************************************************************
Dim objNewKey As RegistryKey
Dim strNewKey As String
Dim arrKey() As String
Dim i As Integer
Dim returns As String = "True"

'Opens the specified key and creates a new sub key inside it.
	Try
		ObjKey = OpenSubKey(Hive, Key)
		If ObjKey Is Nothing Then
			' Get the highest level subkey
			arrKey = Key.Split("\")
			objNewKey = OpenSubKey(Hive, arrKey(0))
			strNewKey = Key.Remove(0, arrKey(0).Length + 1)
            objNewKey = objNewKey.CreateSubKey(strNewKey)
        Else
            returns = "Failure: The key did not exist."
        End If
    Catch ex As Exception
        returns = "Failure: " & ex.Message
    End Try

'Cleanup
    objNewKey = Nothing
    strNewKey = Nothing
    arrKey = Nothing
    i = Nothing
    Hive = Nothing
    Key = Nothing

'Exit
    Return returns

End Function

#End Region
#Region "GetValue"

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
Dim Value As Object

'Opens the specified the value and returns it's value
	Try
		objKey = OpenSubKey(Hive, Key)
		If Not objKey Is Nothing Then
			Value = objKey.GetValue(ValueName, Nothing)
		Else
			Value = "Failure: The Value did not exist."
		End If
	Catch ex As Exception
		Value = "Failure: " & ex.Message
	End Try

'Cleanup
    objKey = Nothing
    Hive = Nothing
    Key = Nothing
    ValueName = Nothing

'Exit
	Return Value

End Function

#End Region
#Region "DeleteValue"

'*************************************************************************
Public Function DeleteValue( _
							ByVal Hive As String, _
							ByVal Key As String, _
							ByVal ValueName As String _
							) As string
'PURPOSE:
'   Deletes the specified value from this key.
'
'OUTPUT:
'   Returns TRUE if value was deleted or does not exist
'*************************************************************************
Dim objKey As Object
Dim Value As Object
Dim returns As String = "True"

'Opens the specified value and deletes it
	Try
		objKey = OpenSubKey(Hive, Key)
		Value = GetValue(Hive, Key, ValueName)
		If Not objKey Is Nothing Then
			If Not Value Is Nothing Then
                objKey.DeleteValue(ValueName, True)
            Else
                returns = "Failure: The value did not exist."
            End If
        Else
            returns = "Failure: The Key did not exist."
        End If
    Catch ex As Exception
        returns = "Failure: " & ex.Message
    End Try

'Clean up
    objKey = Nothing
    Value = Nothing
    ValueName = Nothing
    Key = Nothing
    Hive = Nothing

'Exit
    Return returns

End Function

#End Region
#Region "SetValue"

'*************************************************************************
Public Function SetValue( _
						ByVal Hive As String, _
						ByVal Key As String, _
						ByVal ValueName As String, _
						ByVal Value As Object _
						) As string
'PURPOSE:
'   Sets the specified value. 
'
'OUTPUT:
'   Returns the type of value created
'*************************************************************************
Dim objKey As Object
Dim MYvalue As Object
Dim returns As String

'Opens the specified value if it exists if not it is created so the it's value can be set.
'Returns the type of value created
	Try
    objKey = OpenSubKey(Hive, Key)
    MYvalue = GetValue(Hive, Key, ValueName)
        If Not objKey Is Nothing Then
            If Not MYvalue Is Nothing Then
                objKey.SetValue(ValueName, Value)
            Else
                CreateSubKey(Hive, Key)
                objKey = OpenSubKey(Hive, Key)
                objKey.SetValue(ValueName, Value)
                returns = GetRegistryType(Hive, Key, ValueName)
            End If
        Else
            CreateSubKey(Hive, Key)
            objKey = OpenSubKey(Hive, Key)
            objKey.SetValue(ValueName, Value)
            returns = GetRegistryType(Hive, Key, ValueName)
        End If
    Catch ex As Exception
		returns = "Failure: " & ex.Message
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
#Region "OpenSubKey"

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
ObjHive = GetHive(Hive)

'Opens the specified key and returns it as an object. Private use only.
	If Not ObjHive Is Nothing Then
		Try
			ObjKey = ObjHive.OpenSubKey(Key, True)
		Catch ex As Exception
			ObjKey = Nothing
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
						ByVal Hive As String _
						) As RegistryKey
'PURPOSE:
'   To create a reigistry object to the desired Hive.
'
'OUTPUT:
'   Returns a registry key object
'*************************************************************************

'Returns a registry object
	Try
		Select Case Hive.ToUpper
			Case "HKLM"
				ObjHive = Registry.LocalMachine
			Case "HKEY_LOCAL_MACHINE"
				ObjHive = Registry.LocalMachine

			Case "HKCR"
				ObjHive = Registry.ClassesRoot
			Case "HKEY_CLASSES_ROOT"
				ObjHive = Registry.ClassesRoot

			Case "HKCC"
				ObjHive = Registry.CurrentConfig
			Case "HKEY_CURRENT_CONFIG"
				ObjHive = Registry.CurrentConfig

			Case "HKU"
				ObjHive = Registry.Users
			Case "HKEY_USERS"
				ObjHive = Registry.Users

			Case "HKCU"
				ObjHive = Registry.CurrentUser
			Case "HKEY_CURRENT_USER"
				ObjHive = Registry.CurrentUser

			Case "HKDD"
				ObjHive = Registry.DynData
			Case "HKEY_DYN_DATA"
				ObjHive = Registry.DynData

			Case "HKPD"
				ObjHive = Registry.PerformanceData
			Case "HKEY_PERFORMANCE_DATA"
				ObjHive = Registry.PerformanceData

			Case Else
				ObjHive = Nothing
		End Select
	Catch ex As Exception
		ObjHive = Nothing
	 End Try

'Cleanup
    Hive = Nothing

'Exit
	Return ObjHive

End Function

#End Region
#Region "GetRegistryType"

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
Dim Value As Object
Dim returns As String

'Returns the Common name for registry types
    Value = GetValue(Hive, Key, ValueName)
	If Not Value Is Nothing Then
		Try
			Select Case Value.GetType.FullName
				Case "System.Byte[]"
					returns = "REG_BINARY"
				Case "System.Int32"
					returns = "REG_DWORD"
				Case "System.UInt64"
					returns = "REG_QWORD"
				Case "System.String"
					returns = "REG_SZ"
				Case "System.String[]"
					returns = "REG_MULTI_SZ"
				Case Else
					returns = "Failure: The value type could not be found."
			End Select
		Catch ex As Exception
			returns = "Failure: " & ex.Message
		End Try
	Else
		returns = "Failure:  The value could not be found."
	End If

'Cleanup
    Value = Nothing
    Hive = Nothing
    Key = Nothing
    ValueName = Nothing

'Exit
	Return returns

End Function

#End Region
End Class