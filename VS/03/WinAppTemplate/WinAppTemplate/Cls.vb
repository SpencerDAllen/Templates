#Region "Comments"

'*****************************************************************************'
' PURPOSE:
'   To provide a simple uniform method of accessing all class members and any 
'   exceptions they may throw. This Object also provides a simple All On or All Off
'   Method of suppling verbos logging.
'
'PREREQUISITES:
'   BaseRegistry must be added to the project.
'   MyClsRegistry must be added to the Project.
'   BaseClsFile must be added to the project.
'   MyClsFile must be added to the project.
'   BaseClsDirectory must be added to the project.
'   MyClsDirectory must be added to the project.
'   PwrClsLogs must be added to the project.
'   PwrClsProcess must be added to the project.
'   PwrClsConfig must be added to the project.
'   A properly configured xml document must be located in the same directory as
'   the executable. 
'   The xml document must be named the same as the executable. with the extention 
'   of .config.
'   See Config comments for further details and examples.
'   Microsoft .Net Framework 1.1
'
#Region "Registry"
'
' PURPOSE:
'   The Win32 Registry class is the basis of this class, whose purpose is to
'   read and write to the Windows registry. Our class uses the object data
'   type to read and write the registry. The Win32 Registry class uses dynamic
'   data typing, REG_BINARY registry values would be read in as an array of
'   bytes. Writes to the registry work the same way. An array of strings would
'   be written to the registry as REG_MULTI_SZ.
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
#End Region
#Region "File"
'
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
#End Region
#Region "Directory"
'
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
'   None of the string arrays return with the full path.
'
#End Region
#Region "Logs"
'
'PURPOSE:
'   clsLogFile is a class used to produce useful log files.
'
'PREREQUISITES:
'   see config.
'
'USAGE:
'   boolean = Logfile.start(FilePath As String, [NewLogFrequency As String], [LogFolder As String])     Options = Always, Daily, Weekly, Monthly
'   boolean = Logfile.addEvent(FilePath As String, EventDescription As String)
'   boolean = LogFile.finish(FilePath As String)       
'   Logfile.clearOld(FilePath, daysToKeep As Integer)                                                   Number of days to keep unused logs
'
#End Region
#Region "Process"
'PURPOSE:
'   A Process component provides access to a process that is running on a 
'   computer. A process, in the simplest terms, is a running application. 
'   A thread is the basic unit to which the operating system allocates processor 
'   time. A thread can execute any part of the code of the process, including 
'   parts currently being executed by another thread.
'
'   The Process component is a useful tool for starting, stopping, controlling, 
'   and monitoring applications. Using the Process component, you can obtain a 
'   list of the processes that are running or start a new process. A Process 
'   component is used to access system processes. After a Process component has 
'   been initialized, it can be used to obtain information about the running process. 
'   Such information includes the set of threads, the loaded modules 
'   (.dll and .exe files), and performance information such as the amount of 
'   memory the process is using.
'
'   If you have a path variable declared in your system using quotes, you must 
'   fully qualify that path when starting any process found in that location. 
'   Otherwise, the system will not find the path. For example, if c:\mypath is 
'   not in your path, and you add it using quotation marks: 
'   path = %path%;"c:\mypath", you must fully qualify any process in c:\mypath 
'   when starting it.
'
'   The process component obtains information about a group of properties all 
'   at once. After the Process component has obtained information about one 
'   member of any group, it will cache the values for the other properties in 
'   that group and not obtain new information about the other members of the 
'   group until you call the Refresh method. Therefore, a property value is not 
'   guaranteed to be any newer than the last call to the Refresh method. The group 
'   breakdowns are operating-system dependent.
'
'   A system process is uniquely identified on the system by its process identifier. 
'   Like many Windows resources, a process is also identified by its handle, which 
'   might not be unique on the computer. A handle is the generic term for an 
'   identifier of a resource. The operating system persists the process handle, 
'   which is accessed through the Handle property of the Process component, even 
'   when the process has exited. Thus, you can get the process's administrative 
'   information, such as the ExitCode (usually either zero for success or a nonzero 
'   error code) and the ExitTime. Handles are an extremely valuable resource, so 
'   leaking handles is more virulent than leaking memory.
'
'USAGE:
'   Integer = StartAndWait(Command as string, WindowStyle as Style)
'
#End Region
#Region "PwrCls3DesEncryption"
'
'   PwrCls3DesEncryption is a class used to encrypt a text string with triple DES encryption.
'
'DESCRIPTION:
'   The original Data Encryption Standard (DES) encrytion standard encrypts data using 8 byte
'   blocks. The same key used to encrypt data is also used to decrypt data. When encrypting
'   data with DES, you must provide an 8-byte key which is reduced to a 7-byte key 
'   because the algorithm removed the 8th bit of every key byte for parity purposes. DES
'   employs 16 rounds of encryption to every block of data. The key is then slightly modified
'   and the block of data is encrypted again. This continues on until the block of data has
'   been encrypted 16 times. 
'   DES encryption uses standard mathematical and logical operators for encryption - it was
'   implemented very easily in the late 1970s with the computer hardware available at that
'   time. DES encryption was officially broken in 1997 during a challenge sponsored by RSA
'   Security. As the name implies, TripleDES performs three times as much encryption as
'   standard DES. TripleDES requires a 24-byte key, which is divided into three 8-byte keys
'   for encrypting each block 3 times. When you take the rounds of encryption into
'   consideration, each block is actually encrypted 48 times. TripleDES is a very secure
'   encryption algorithm and will be the basis of this class. You will want to change the
'   key for use with your applications.
'
'Usage:
'   boolean = TDES.Encrypt(mystring)
'   boolean = TDES.Decrypt(myencryptedtext)
'   boolean = TDES.saveEncryptedText(myencryptedtext, storedfilepath)
'   boolean = TDES.loadEncryptedText(storedfilepath)
'
#End Region
#Region "Config"
'
' PURPOSE: To retrieve configuration parameters for applications.
'   This allows software to be reconfigured without change management or
'   recompiling source code.
'
'Details: If you need to pass a value with a quote(") substute
'   the xml equivelint (&quot;) It will be converted automaticly.
'
'Fine Name Example: WinAppTemplate.exe.config
'Configuration example:
'
'   <?xml version="1.0" encoding="utf-8" ?>
'   <configuration>
'       <!-- This section of settings is common to all programs built upon the DTS WinAppTemplate -->
'       <AllProgramSettings property="LongsOnOff" Value="True" />
'       <!-- LongsOnOff is the main logging setting if it is set to false then all of 
'       the "AllProgramSettings" will be ignored. This value can only be set to "True" or
'       "False" -->
'       <AllProgramSettings property="DaysToKeepLogs" Value="365" />
'       <!-- The default is 365 days, any positive whole numeric value will be accepted. -->
'       <AllProgramSettings property="NewLogFrequency" Value="Always" />
'       <!-- Options = Always, Daily, Weekly, Monthly -->
'       <AllProgramSettings property="LogFolder" Value="" />
'       <!-- The default is a logs folder relitive to the executable -->
'
'
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
'   Intger = GetCfgTypeCount()
'   String() = GetCfgTypes()
'   Intger = GetPropCount(Type As String)
'   String = GetPropValuePair(Type as integer, Prop as integer)
'   String = GetPropValuePair(Type as integer, Prop as string)
'   String = GetPropValuePair(Type as string, Prop as integer)
'   String = GetPropValuePair(Type as string, Prop as string)
'   Boolean = AddPropValuePair(Type as integer, NewProp as string, NewValue as string)
'   Boolean = AddPropValuePair(Type as string, NewProp as string, NewValue as string)
'   Boolean = ChangeValue(Type as string, Prop as string, NewValue as string)
'   Boolean = SaveConfig()
'
#End Region
'USAGE:
'   All classes will return exceptions as string from the individual classes 
'   MyException property.
'
'   All Classes except PwrClsLogs will turn logging on or off by way of the 
'   individual classes LogOnOff property.
'
'AUTHOR:
'   Spencer Allen
'   2/11/08
'
' VERSIONS:
'   1.00 - Base Version (Spencer Allen)
'*****************************************************************************'

#End Region
Public Class Cls
#Region "Declorations"
Implements IDisposable

Public Cfg As New PwrClsConfigFile
Public File As New MyClsFile
Public Logs As New PwrClsLogs
Public Directory As New MyClsDirectory
Public Registry As New MyClsRegistry
Public Process As New PwrClsProcess

#End Region
#Region "Constructor"

'*************************************************************************
Public Sub New()
'PURPOSE:
'   To turn Logging on or off for the entire program.
'
'OUTPUT:
'   Logging is turned on or off.
'*************************************************************************
Dim L As String = Cfg.GetPropValuePair("Logging", "LogsOnOff")
L = L.Substring(L.IndexOf(",") + 1)
Dim logging As Boolean = L

File.LogOnOff = logging
Directory.LogOnOff = logging
Registry.LogOnOff = logging
Process.LogOnOff = logging
Cfg.LogOnOff = logging

'Cleanup
L = Nothing
logging = Nothing

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

Cfg.Dispose()
File.Dispose()
Logs.Dispose()
Directory.Dispose()
Registry.Dispose()
Process.Dispose()

'Cleanup
Cfg = Nothing
File = Nothing
Logs = Nothing
Directory = Nothing
Registry = Nothing
Process = Nothing

'exit
End Sub

#End Region
End Class