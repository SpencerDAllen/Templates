#Region "Comments"

'*****************************************************************************'
' PURPOSE: To retrieve configuration parameters for applications.
'   This allows software to be reconfigured without change management or
'   recompiling source code.
'
'PREREQUISITES:
'   A properly configured xml document must be located in the same directory as
'   the executable. The xml document must be named the same as the executable. with 
'   the extention of .config. If you need to pass a value with a quote(") substute
'   the xml equivelint (&quot;) It will be converted automaticly.
'
'File Name Example: WinAppTemplate.exe.config
'Configuration example:
'
'   <?xml version="1.0" encoding="utf-8" ?>
'   <configuration>
'       <!-- Task sequencer settings are the settings specific for this application -->
'       <TaskSequencerSettings Properties="Title" Values="Task Sequencer" />
'       <TaskSequencerSettings Properties="SilentRun" Values="false" />
'       <TaskSequencerSettings Properties="SecondsToDelayStart" Values="0" />
'       <TaskSequencerSettings Properties="ShowStartButton" Values="false" />
'       <TaskSequencerSettings Properties="ShowCancelButton" Values="false" />
'       <TaskSequencerSettings Properties="StopOnErrors" Values="true" />
'       <TaskSequencerSettings Properties="UseRegistryLogging" Values="true" />
'       <TaskSequencerSettings Properties="RegistryPath" Values="HKLM\SOFTWARE\PNM\Unattend\ImageRevision" />
'       <TaskSequencerSettings Properties="ResumeOnLogon" Values="false" />
'       <!-- Another section of settings -->
'       <Tasks Properties="Install .Net 2.0" Values="\\albvecg01\orbit$\dotNET_Framework\2.0\install.exe /q" />
'       <Tasks Properties="Install .Net 2.0 security permissions" Values="msiexec.exe /i &quot;\\albvecg01\orbit$\dotNET_Framework\2.0\IntranetSecurity2.msi&quot; /qn" />
'       <Tasks Properties="Install the Recovery console" Values="C:\i386\winnt32.exe /cmdcons /unattend" />
'       <Tasks Properties="Install maintenance tasks" Values="\\albvecg01\SourceCode\Administration\ScheduledTasks-Maintenance\ScheduleMaintenance.exe" />
'       <Tasks Properties="Enable audit policies" Values="c:\postprep\auditpol.exe /enable /system:all /logon:all /object:failure /privilege:failure /policy:all /sam:all" />
'       <Tasks Properties="Allow users to change system time" Values="c:\postprep\ntrights.exe +r SeSystemtimePrivilege -u &quot;Authenticated Users&quot;" />
'       <Tasks Properties="Install Internet Explorer 7" Values="cmd.exe /c &quot;\\albvecg01\orbit$\Microsoft\Internet_Explorer\7.0\Install\IE7-WindowsXP-x86-enu.exe&quot; /quiet /passive /norestart" />
'   </configuration>
'
'USAGE:
'
'   Intger = GetCfgTypeCount(Additional as Boolean)
'   String() = GetCfgTypes(Additional as Boolean)
'   Intger = GetPropCount(Type As String, Additional as Boolean)
'   String = GetPropValuePair(Type as integer, Prop as integer, Additional as Boolean)
'   String = GetPropValuePair(Type as integer, Prop as string, Additional as Boolean)
'   String = GetPropValuePair(Type as string, Prop as integer, Additional as Boolean)
'   String = GetPropValuePair(Type as string, Prop as string, Additional as Boolean)
'   Boolean = AddPropValuePair(Type as integer, NewProp as string, NewValue as string, Additional as Boolean)
'   Boolean = AddPropValuePair(Type as string, NewProp as string, NewValue as string, Additional as Boolean)
'   Boolean = ChangeValue(Type as string, Prop as string, NewValue as string, Additional as Boolean)
'   Boolean = SaveConfig(, Additional as Boolean)
'
'AUTHOR:
'   Spencer Allen
'   06-20-2008
'
' VERSIONS:
'   1.00 - Base Version (Spencer Allen)
'   1.01 - Adds the ability to use an additional configuration file (Spencer Allen)
'   1.02 - Adds Delete Data Row (Spencer Allen)
'*****************************************************************************'

#End Region
Public Class PwrClsConfigFile
#Region "Declorations"
Implements IDisposable

Private myConfig As New Xml.XmlDataDocument
Private ExtraConfig As New Xml.XmlDataDocument

#End Region
#Region "Constructor"

'*************************************************************************
Public Sub New()
'PURPOSE:
'   Build the data set that contains the config file information.
'
'OUTPUT:
'   Unknown at this time.
'*************************************************************************
Dim Reflection As String = System.Reflection.Assembly.GetExecutingAssembly().ToString
Reflection = Reflection.Substring(0, Reflection.IndexOf(","))
Dim AppPath As String = System.Windows.Forms.Application.StartupPath & "\" & Reflection & ".exe.config"

'Create the DataSet to use
    Try
        myConfig.DataSet.ReadXml(AppPath)
        LogOnOff = False
    Catch ex As Exception
        MsgBox("The configuration file for this application is improperly configured." & vbCrLf _
        & "See your Application developer for assistance." & vbCrLf & ex.Message, MsgBoxStyle.Critical, _
        "Fatal Error!")
        Environment.Exit(-3)
    End Try

'Cleanup

'Exit
End Sub

#End Region
#Region "ImportConfig"

'*************************************************************************
Public Sub ImportConfig(ByVal ConfigPath As String)
'PURPOSE:
'   Build a data set from a different seccondary configuration file.
'
'OUTPUT:
'   Unknown at this time.
'*************************************************************************

'Create the DataSet to use
    Try
        ExtraConfig.DataSet.ReadXml(ConfigPath)
        LogOnOff = False
    Catch ex As Exception
        MsgBox("The configuration file for this application is improperly configured." & vbCrLf _
        & "See your Application developer for assistance." & vbCrLf & ex.Message, MsgBoxStyle.Critical, _
        "Fatal Error!")
        Environment.Exit(-3)
    End Try

'Cleanup

'Exit
End Sub

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
    myConfig = Nothing
    ExtraConfig = Nothing
    Except = Nothing
    LogsOnOff = Nothing

'Exit
End Sub

#End Region
#Region "GetCfgTypeCount"

'*************************************************************************
Function GetCfgTypeCount(Optional ByVal AdditionalConfig As Boolean = False) As Integer
'PURPOSE:
'   Retrive the number of types in the config file.
'
'OUTPUT:
'   integer.
'*************************************************************************
Dim Logevent As String = "Cfg.GetCfgTypeCount: "
Dim Returned As Integer

    If Not AdditionalConfig Then
        Try
            Returned = myConfig.DataSet.Tables.Count
        Catch ex As Exception
            Except = ex.Message
            Returned = 0
        End Try
    Else
        Try
            Returned = ExtraConfig.DataSet.Tables.Count
        Catch ex As Exception
            Except = ex.Message
            Returned = 0
        End Try
    End If

    If LogsOnOff Then
        If Not Except = "" Then
            Pnm.Logs.addEvent("Failure: " & Logevent & vbCrLf & Except)
        Else
            Pnm.Logs.addEvent("Success: " & Logevent)
        End If
    End If

'Cleanup
    Logevent = Nothing

'Exit
    Return Returned

End Function

#End Region
#Region "GetCfgTypes"

'*************************************************************************
Function GetCfgTypes(Optional ByVal AdditionalConfig As Boolean = False) As String()
'PURPOSE:
'   Retrive a string array of all the tables\types in the dataset\config file.
'
'OUTPUT:
'   string array.
'*************************************************************************
Dim LogEvent As String = "Cfg.GetCfgTypes: "
Dim Tables As String()
Dim i As Integer
Except = ""


    If Not AdditionalConfig Then
        Try
            ReDim Tables(myConfig.DataSet.Tables.Count - 1)
            For i = 0 To myConfig.DataSet.Tables.Count - 1
                Tables(i) = myConfig.DataSet.Tables(i).ToString
            Next
        Catch ex As Exception
            Except = ex.Message
        End Try
    Else
        Try
            ReDim Tables(ExtraConfig.DataSet.Tables.Count - 1)
            For i = 0 To ExtraConfig.DataSet.Tables.Count - 1
                Tables(i) = ExtraConfig.DataSet.Tables(i).ToString
            Next
        Catch ex As Exception
            Except = ex.Message
        End Try
    End If

    If LogsOnOff Then
        If Not Except = "" Then
            Pnm.Logs.addEvent("Failure: " & LogEvent & vbCrLf & Except)
        Else
            Pnm.Logs.addEvent("Success: " & LogEvent)
        End If
    End If

'Cleanup
    i = Nothing
    LogEvent = Nothing

'Cleanup
    Return Tables

End Function

#End Region
#Region "GetPropCount"

'*************************************************************************
Function GetPropCount(ByVal Type As String, _
                    Optional ByVal AdditionalConfig As Boolean = False) As Integer
'PURPOSE:
'   Retrive the total number of property value pairs.
'
'OUTPUT:
'   integer.
'*************************************************************************
Dim LogEvent As String = "CFG.GetPropCount: " & Type
Dim Returned As Integer
Except = ""

    If Not AdditionalConfig Then
        Try
            Returned = myConfig.DataSet.Tables(Type).Rows.Count
        Catch ex As Exception
            Except = ex.Message
            Returned = 0
        End Try
    Else
        Try
            Returned = ExtraConfig.DataSet.Tables(Type).Rows.Count
        Catch ex As Exception
            Except = ex.Message
            Returned = 0
        End Try
    End If

    If LogsOnOff Then
        If Not Except = "" Then
            Pnm.Logs.addEvent("Failure: " & LogEvent & vbCrLf & Except)
        Else
            Pnm.Logs.addEvent("Success: " & LogEvent)
        End If
    End If

'Cleanup
    Type = Nothing
    LogEvent = Nothing

'Exit
    Return Returned

End Function

#End Region
#Region "GetPropValuePairFunctions"
#Region "GetPropValuePair,integer,integer"

'*************************************************************************
Overloads Function GetPropValuePair(ByVal Type As Integer, _
                                    ByVal Prop As Integer, _
                                    Optional ByVal AdditionalConfig As Boolean = False _
                                    ) As String
'PURPOSE:
'   Retrive a string of a single property and value pair from the specified
'   type.
'
'OUTPUT:
'   string.
'*************************************************************************
Dim LogEvent As String = "CFG.GetPropValuePair: " & Type.ToString & " " & Prop.ToString
Dim Proper As String
Dim Value As String
Dim Returned As String
Except = ""

'Get the property and value and return concantinated coma seperated

    If Not AdditionalConfig Then
        Try
            Proper = myConfig.DataSet.Tables(Type).Rows(Prop).Item(0)
            Value = myConfig.DataSet.Tables(Type).Rows(Prop).Item(1)
            Returned = Proper & "," & Value
        Catch ex As Exception
            Except = ex.Message
        End Try
    Else
        Try
            Proper = ExtraConfig.DataSet.Tables(Type).Rows(Prop).Item(0)
            Value = ExtraConfig.DataSet.Tables(Type).Rows(Prop).Item(1)
            Returned = Proper & "," & Value
        Catch ex As Exception
            Except = ex.Message
        End Try
    End If

    If LogsOnOff Then
        If Not Except = "" Then
            Pnm.Logs.addEvent("Failure: " & LogEvent & vbCrLf & Except)
        Else
            Pnm.Logs.addEvent("Success: " & LogEvent)
        End If
    End If

'Cleanup
    Value = Nothing
    Prop = Nothing
    LogEvent = Nothing
    Prop = Nothing
    Type = Nothing

'Exit
    Return Returned

End Function

#End Region
#Region "GetPropValuePair,integer,string"

'*************************************************************************
Overloads Function GetPropValuePair(ByVal Type As Integer, _
                                    ByVal Prop As String, _
                                    Optional ByVal AdditionalConfig As Boolean = False _
                                    ) As String
'PURPOSE:
'   Retrive a string of a single property and value pair from the specified
'   type.
'
'OUTPUT:
'   string.
'*************************************************************************
Dim LogEvent As String = "CFG.GetPropValuePair: " & Type.ToString & " " & Prop
Dim Value As String
Dim Returned As String
Dim i As Integer
Except = ""

'Get the property and value and concantinated coma seperated

    If Not AdditionalConfig Then
        Try
            For i = 0 To myConfig.DataSet.Tables(Type).Rows.Count - 1
                If myConfig.DataSet.Tables(Type).Rows(i).Item(0) = Prop Then
                    Value = myConfig.DataSet.Tables(Type).Rows(i).Item(1)
                    Returned = Prop & "," & Value
                    Except = ""
                    Exit For
                Else
                    Returned = ""
                    Except = "The property was not found in this type."
                End If
            Next
        Catch ex As Exception
            Except = ex.Message
        End Try
    Else
        Try
            For i = 0 To ExtraConfig.DataSet.Tables(Type).Rows.Count - 1
                If ExtraConfig.DataSet.Tables(Type).Rows(i).Item(0) = Prop Then
                    Value = ExtraConfig.DataSet.Tables(Type).Rows(i).Item(1)
                    Returned = Prop & "," & Value
                    Except = ""
                    Exit For
                Else
                    Returned = ""
                    Except = "The property was not found in this type."
                End If
            Next
        Catch ex As Exception
            Except = ex.Message
        End Try
    End If

    If LogsOnOff Then
        If Not Except = "" Then
            Pnm.Logs.addEvent("Failure: " & LogEvent & vbCrLf & Except)
        Else
            Pnm.Logs.addEvent("Success: " & LogEvent)
        End If
    End If

'Cleanup
    Value = Nothing
    Prop = Nothing
    LogEvent = Nothing
    Prop = Nothing
    Type = Nothing
    i = Nothing

'Exit
    Return Returned

End Function

#End Region
#Region "GetPropValuePair,string,integer"

'*************************************************************************
Overloads Function GetPropValuePair(ByVal Type As String, _
                                    ByVal Prop As Integer, _
                                    Optional ByVal AdditionalConfig As Boolean = False _
                                    ) As String
'PURPOSE:
'   Retrive a string of a single property and value pair from the specified
'   type.
'
'OUTPUT:
'   string.
'*************************************************************************
Dim LogEvent As String = "CFG.GetPropValuePair: " & Type & " " & Prop.ToString
Dim Proper As String
Dim Value As String
Dim Returned As String
Except = ""

'Get the property and value and concantinated coma seperated

    If Not AdditionalConfig Then
        Try
            Proper = myConfig.DataSet.Tables(Type).Rows(Prop).Item(0)
            Value = myConfig.DataSet.Tables(Type).Rows(Prop).Item(1)
            Returned = Proper & "," & Value
        Catch ex As Exception
            Except = ex.Message
        End Try
    Else
        Try
            Proper = ExtraConfig.DataSet.Tables(Type).Rows(Prop).Item(0)
            Value = ExtraConfig.DataSet.Tables(Type).Rows(Prop).Item(1)
            Returned = Proper & "," & Value
        Catch ex As Exception
            Except = ex.Message
        End Try
    End If

    If LogsOnOff Then
        If Not Except = "" Then
            Pnm.Logs.addEvent("Failure: " & LogEvent & vbCrLf & Except)
        Else
            Pnm.Logs.addEvent("Success: " & LogEvent)
        End If
    End If

'Cleanup
    Value = Nothing
    Proper = Nothing
    LogEvent = Nothing
    Prop = Nothing
    Type = Nothing

'Exit
    Return Returned

End Function

#End Region
#Region "GetPropValuePair,string,string"

'*************************************************************************
Overloads Function GetPropValuePair(ByVal Type As String, _
                                    ByVal Prop As String, _
                                    Optional ByVal AdditionalConfig As Boolean = False _
                                    ) As String
'PURPOSE:
'   Retrive a string of a single property and value pair from the specified
'   type.
'
'OUTPUT:
'   string.
'*************************************************************************
Dim LogEvent As String = "CFG.GetPropValuePair: " & Type & " " & Prop
Dim Value As String
Dim Returned As String
Dim i As Integer
Except = ""

'Get the property and value and concantinated coma seperated

    If Not AdditionalConfig Then
        Try
            For i = 0 To myConfig.DataSet.Tables(Type).Rows.Count - 1
                If myConfig.DataSet.Tables(Type).Rows(i).Item(0) = Prop Then
                    Value = myConfig.DataSet.Tables(Type).Rows(i).Item(1)
                    Returned = Prop & "," & Value
                    Except = ""
                    Exit For
                Else
                    Returned = ""
                    Except = "The property was not found in this type."
                End If
            Next
        Catch ex As Exception
            Except = ex.Message
        End Try
    Else
        Try
            For i = 0 To ExtraConfig.DataSet.Tables(Type).Rows.Count - 1
                If ExtraConfig.DataSet.Tables(Type).Rows(i).Item(0) = Prop Then
                    Value = ExtraConfig.DataSet.Tables(Type).Rows(i).Item(1)
                    Returned = Prop & "," & Value
                    Except = ""
                    Exit For
                Else
                    Returned = ""
                    Except = "The property was not found in this type."
                End If
            Next
        Catch ex As Exception
            Except = ex.Message
        End Try
    End If

    If LogsOnOff Then
        If Not Except = "" Then
            Pnm.Logs.addEvent("Failure: " & LogEvent & vbCrLf & Except)
        Else
            Pnm.Logs.addEvent("Success: " & LogEvent)
        End If
    End If

'Cleanup
    Value = Nothing
    Prop = Nothing
    LogEvent = Nothing
    Prop = Nothing
    Type = Nothing
    i = Nothing

'Exit
    Return Returned

End Function

#End Region
#End Region
#Region "AddPropValuePairFunctions"
#Region "AddPropValuepair, integer, string, string"

'*************************************************************************
Overloads Function AddPropValuePair(ByVal Type As Integer, _
                                    ByVal NewProp As String, _
                                    ByVal NewValue As String, _
                                    Optional ByVal AdditionalConfig As Boolean = False _
                                    ) As Boolean
'PURPOSE:
'   add a property and value pair to the specified type.
'
'OUTPUT:
'   boolean.
'*************************************************************************
Dim LogEvent As String = "CFG.AddPropValuePair: " & Type.ToString & " " & NewProp & " " & NewValue
Dim Returned As Boolean = True
Dim MyRow As DataRow
Except = ""

'Add the property and value pair to the DataSet.

    If Not AdditionalConfig Then
        Try
            MyRow = myConfig.DataSet.Tables(Type).NewRow
            MyRow(0) = NewProp
            MyRow(1) = NewValue
            myConfig.DataSet.Tables(Type).Rows.Add(MyRow)
            myConfig.DataSet.AcceptChanges()
        Catch ex As Exception
            Except = ex.Message
            Returned = False
        End Try
    Else
        Try
            MyRow = ExtraConfig.DataSet.Tables(Type).NewRow
            MyRow(0) = NewProp
            MyRow(1) = NewValue
            ExtraConfig.DataSet.Tables(Type).Rows.Add(MyRow)
            ExtraConfig.DataSet.AcceptChanges()
        Catch ex As Exception
            Except = ex.Message
            Returned = False
        End Try
    End If

    If LogsOnOff Then
        If Not Except = "" Then
            Pnm.Logs.addEvent("Failure: " & LogEvent & vbCrLf & Except)
        Else
            Pnm.Logs.addEvent("Success: " & LogEvent)
        End If
    End If

'Cleanup
    LogEvent = Nothing
    NewValue = Nothing
    NewProp = Nothing
    Type = Nothing

'Exit
    Return Returned

End Function

#End Region
#Region "AddPropValuepair, string, string, string"

'*************************************************************************
Overloads Function AddPropValuePair(ByVal Type As String, _
                                    ByVal NewProp As String, _
                                    ByVal NewValue As String, _
                                    Optional ByVal AdditionalConfig As Boolean = False _
                                    ) As Boolean
'PURPOSE:
'   add a property and value pair to the specified type.
'
'OUTPUT:
'   boolean.
'*************************************************************************
Dim LogEvent As String = "CFG.AddPropValuePair: " & Type & " " & NewProp & " " & NewValue
Dim Returned As Boolean = True
Dim MyRow As DataRow
Except = ""

'Add the property and value pair to the DataSet.

    If Not AdditionalConfig Then
        Try
            MyRow = myConfig.DataSet.Tables(Type).NewRow()
            MyRow(0) = NewProp
            MyRow(1) = NewValue
            myConfig.DataSet.Tables(Type).Rows.Add(MyRow)
            myConfig.DataSet.AcceptChanges()
        Catch ex As Exception
            Except = ex.Message
            Returned = False
        End Try
    Else
        Try
            MyRow = ExtraConfig.DataSet.Tables(Type).NewRow()
            MyRow(0) = NewProp
            MyRow(1) = NewValue
            ExtraConfig.DataSet.Tables(Type).Rows.Add(MyRow)
            ExtraConfig.DataSet.AcceptChanges()
        Catch ex As Exception
            Except = ex.Message
            Returned = False
        End Try
    End If

    If LogsOnOff Then
        If Not Except = "" Then
            Pnm.Logs.addEvent("Failure: " & LogEvent & vbCrLf & Except)
        Else
            Pnm.Logs.addEvent("Success: " & LogEvent)
        End If
    End If

'Cleanup
    LogEvent = Nothing
    NewValue = Nothing
    NewProp = Nothing
    Type = Nothing
    MyRow = Nothing

'Exit
    Return Returned

End Function

#End Region
#End Region
#Region "DeleteDataRow"

'*************************************************************************
Function DeleleteDataRow(ByVal Type As String, _
                        ByVal Prop As String, _
                        Optional ByVal AdditionalConfig As Boolean = False _
                        ) As Boolean
'PURPOSE:
'   Remove unwanted data Rows.
'
'OUTPUT:
'   boolean.
'*************************************************************************
Dim LogEvent As String = "CFG.DeleleteDataRow: " & Type & " " & Prop
Dim Returned As Boolean = True
Dim ConfigProp As String
Dim i As Integer
Except = ""

        Try
            For i = 0 To Me.GetPropCount(Type) - 1
                ConfigProp = Me.GetPropValuePair(Type, i)
                ConfigProp = ConfigProp.Substring(0, ConfigProp.IndexOf(","))
                If ConfigProp = Prop Then
                    myConfig.DataSet.Tables(Type).Rows(i).Delete()
                    Exit For
                End If
            Next
        Catch ex As Exception
            Except = ex.Message
            Returned = False
        End Try
End Function

#End Region
#Region "ChangeValue"

'*************************************************************************
Overloads Function ChangeValue(ByVal Type As String, _
                                ByVal Prop As String, _
                                ByVal NewValue As String, _
                                Optional ByVal AdditionalConfig As Boolean = False _
                                ) As Boolean
'PURPOSE:
'   add a property and value pair to the specified type.
'
'OUTPUT:
'   boolean.
'*************************************************************************
Dim LogEvent As String = "CFG.ChangeValue: " & Type & " " & Prop & " " & NewValue
Dim Returned As Boolean = True
Dim i As Integer
Except = ""

    If Not AdditionalConfig Then
        Try
            For i = 0 To myConfig.DataSet.Tables(Type).Rows.Count - 1
                If myConfig.DataSet.Tables(Type).Rows(i).Item(0) = Prop Then
                    myConfig.DataSet.Tables(Type).Rows(i).Item(1) = NewValue
                Else
                    Returned = False
                    Except = "The specified property was not found in this type."
                End If
            Next
        Catch ex As Exception
            Returned = False
            Except = ex.Message
        End Try
    Else
        Try
            For i = 0 To ExtraConfig.DataSet.Tables(Type).Rows.Count - 1
                If ExtraConfig.DataSet.Tables(Type).Rows(i).Item(0) = Prop Then
                    ExtraConfig.DataSet.Tables(Type).Rows(i).Item(1) = NewValue
                Else
                    Returned = False
                    Except = "The specified property was not found in this type."
                End If
            Next
        Catch ex As Exception
            Returned = False
            Except = ex.Message
        End Try
    End If

    If LogsOnOff Then
        If Not Except = "" Then
            Pnm.Logs.addEvent("Failure: " & LogEvent & vbCrLf & Except)
        Else
            Pnm.Logs.addEvent("Success: " & LogEvent)
        End If
    End If

'Cleanup
    i = Nothing
    LogEvent = Nothing
    NewValue = Nothing
    Prop = Nothing
    Type = Nothing

'Exit
    Return Returned

End Function

#End Region
#Region "SaveConfig"

'*************************************************************************
Function SaveConfig(Optional ByVal AdditionalConfig As Boolean = False, _
                    Optional ByVal AdditionalConfigPath As String = "" _
                    ) As Boolean
'PURPOSE:
'   Creates a new configuration file.
'
'OUTPUT:
'   boolean.
'*************************************************************************
Dim LogEvent As String = "CFG.SaveConfig: "
Dim Returned As Boolean = True
Except = ""
Dim Reflection As String = System.Reflection.Assembly.GetExecutingAssembly().ToString
Reflection = Reflection.Substring(0, Reflection.IndexOf(","))
Dim AppPath As String = System.Windows.Forms.Application.StartupPath & "\" & Reflection

'Check to see if the programmer is attempting to overwrite the current config file.

    If Not AdditionalConfig Then
        Try
            myConfig.DataSet.WriteXml(AppPath & ".exe.config")
        Catch ex As Exception
            Except = ex.Message
            Returned = False
        End Try
    Else
        Try
            ExtraConfig.DataSet.WriteXml(AdditionalConfigPath)
        Catch ex As Exception
            Except = ex.Message
            Returned = False
        End Try
    End If

    If LogsOnOff Then
        If Not Except = "" Then
            Pnm.Logs.addEvent("Failure: " & LogEvent & vbCrLf & Except)
        Else
            Pnm.Logs.addEvent("Success: " & LogEvent)
        End If
    End If

'Cleanup
    LogEvent = Nothing

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