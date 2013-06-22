#Region "Comments"

'**********************************************************************************************
'PURPOSE:
'   clsDESstringEncrytion is a class used to encrypt a text string with triple DES encryption.
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
'AUTHOR:
'   Shayne Marriage
'   BTS Applications
'   04-20-2007
'
'Usage:
'myencryptedtext = TDES.Encrypt(mystring)
'mydecryptedtext = TDES.Decrypt(myencryptedtext)
'savesuccessfull = TDES.saveEncryptedText(myencryptedtext, storedfilepath)
'mystoredtext = TDES.loadEncryptedText(storedfilepath)
'
' VERSIONS:
'   1.00 - Base Version (Shayne Marriage)
'**********************************************************************************************

#End Region
#Region "Imports"

Imports System
Imports System.IO
Imports System.Text
Imports System.Security.Cryptography

#End Region
Public Class PwrCls3DesEncryption
#Region "Declorations"
Implements IDisposable
    '24 byte DES Private key
    Private key() As Byte = {23, 20, 6, 10, 9, 7, 18, 22, 19, 14, 3, 5, 13, 4, 15, 12, 17, 8, 1, 11, 24, 2, 21, 16}
    '8 byte DES Initialization Vector
    Private iv() As Byte = {65, 110, 68, 26, 69, 178, 200, 219}

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
    key = Nothing
    iv = Nothing
    Except = Nothing
    LogsOnOff = Nothing

'Exit
End Sub

#End Region
#Region "Encrypt"

'*************************************************************************
Public Function Encrypt(ByVal plainText As String _
                        ) As Byte()
'PURPOSE:
'   Declare a UTF8Encoding object so we may use the GetByte method to transform 
'   the plainText into a Byte array.
'
'   Create a new TripleDES service provider.
'
'   The ICryptTransform interface uses the TripleDES crypt provider along 
'   with encryption key and init vector information.
'
'   All cryptographic functions need a stream to output the encrypted information. 
'   Here we declare a memory stream for this purpose.
'
'   Write the encrypted information to the stream. Flush the information when
'   done to ensure everything is out of the buffer. 
'
'   Read the stream back into a Byte array and return it to the calling 
'   method.
'
'RETURNS:
'   True if sucessful.
'*************************************************************************
Dim utf8encoder As New UTF8Encoding
Dim logEvent As String = "Encryption.Encrypt: " & plainText & " "
Dim noErrorFlag As Boolean = True
Dim TDesProvider As New TripleDESCryptoServiceProvider
Dim EncryptedStream As New MemoryStream
Dim result(EncryptedStream.Length - 1) As Byte

    Try
        Dim inputInBytes() As Byte = utf8encoder.GetBytes(plainText)
        Dim cryptoTransform As ICryptoTransform = TDesProvider.CreateEncryptor(Me.key, Me.iv)
        Dim cryptStream As CryptoStream = New CryptoStream(EncryptedStream, cryptoTransform, CryptoStreamMode.Write)
        cryptStream.Write(inputInBytes, 0, inputInBytes.Length)
        cryptStream.FlushFinalBlock()
        EncryptedStream.Position = 0
        ReDim result(EncryptedStream.Length - 1)
        EncryptedStream.Read(result, 0, EncryptedStream.Length)
        cryptStream.Close()
    Catch ex As Exception
        noErrorFlag = False
        Except = ex.Message
    End Try

    If LogsOnOff Then
        If Not Except = "" Then
            logEvent = logEvent & "Failure: " & logEvent & vbCrLf & Except
            Pnm.Logs.addEvent(logEvent)
        Else
            logEvent = "Success: " & logEvent
            Pnm.Logs.addEvent(logEvent)
        End If
    End If

'Cleanup
    logEvent = Nothing
    EncryptedStream = Nothing
    TDesProvider = Nothing
    utf8encoder = Nothing
    plainText = Nothing

'Exit
    Return result

End Function

#End Region
#Region "Decrypt"

'*************************************************************************
Private Function Decrypting(ByVal inputInBytes() As Byte _
                        ) As String
'PURPOSE:
'   UTFEncoding is used to transform the decrypted Byte Array information back 
'   into a string.
'
'   As before we must provide the encryption/decryption key along with the init vector. 
'
'   Provide a memory stream to decrypt information into 
'
'   Read the memory stream and convert it back into a string 
'
'RETURNS:
'   True if sucessful.
'*************************************************************************
Dim logEvent As String = "Encryption.Decrypt: "
Dim noErrorFlag As Boolean = True
Dim utf8encoder As UTF8Encoding = New UTF8Encoding
Dim tdesProvider As TripleDESCryptoServiceProvider = New TripleDESCryptoServiceProvider
Dim decryptedStream As MemoryStream = New MemoryStream
'Dim myutf As New UTF8Encoding

    Try
        Dim cryptoTransform As ICryptoTransform = tdesProvider.CreateDecryptor(Me.key, Me.iv)
        Dim cryptStream As CryptoStream = New CryptoStream(decryptedStream, cryptoTransform, CryptoStreamMode.Write)
        cryptStream.Write(inputInBytes, 0, inputInBytes.Length)
        cryptStream.FlushFinalBlock()
        decryptedStream.Position = 0
        Dim result(decryptedStream.Length - 1) As Byte
        decryptedStream.Read(result, 0, decryptedStream.Length)
        cryptStream.Close()
        Dim myutf As UTF8Encoding = New UTF8Encoding
        Return myutf.GetString(result)
    Catch ex As Exception
        noErrorFlag = False
        Except = ex.Message
    End Try

    If LogsOnOff Then
        If Not Except = "" Then
            logEvent = logEvent & "Failure: " & logEvent & vbCrLf & Except
            Pnm.Logs.addEvent(logEvent)
        Else
            logEvent = "Success: " & logEvent
            Pnm.Logs.addEvent(logEvent)
        End If
    End If

'Cleanup
    logEvent = Nothing
    decryptedStream = Nothing
    tdesProvider = Nothing
    utf8encoder = Nothing
    inputInBytes = Nothing

'Exit
End Function

#End Region
#Region "SaveEncryptedText"

'*************************************************************************
Public Function SaveEncryptedText(ByVal BytEncryptedText() As Byte, _
                                    ByVal StoredFilePath As String _
                                    ) As Boolean
'PURPOSE:
'   Write the encrypted text to a text file
'
'RETURNS:
'   True if sucessful.
'*************************************************************************
Dim logEvent As String = "Encryption.SaveEncryptedText: " & StoredFilePath & " "
Dim noErrorFlag As Boolean = True

    Try
        If Pnm.File.Exists(StoredFilePath) Then
            Pnm.File.Delete(StoredFilePath)
        End If
        Dim fStream As New FileStream(StoredFilePath, FileMode.Create, FileAccess.Write)
        Dim bWriter As New BinaryWriter(fStream)
        bWriter.Write(BytEncryptedText)
        bWriter.Close()
        fStream.Close()
        noErrorFlag = Pnm.File.Exists(StoredFilePath)
    Catch ex As Exception
        noErrorFlag = False
        Except = ex.Message
    End Try

    If LogsOnOff Then
        If Not Except = "" Then
            logEvent = logEvent & "Failure: " & logEvent & vbCrLf & Except
            Pnm.Logs.addEvent(logEvent)
        Else
            logEvent = "Success: " & logEvent
            Pnm.Logs.addEvent(logEvent)
        End If
    End If

'Cleanup
    BytEncryptedText = Nothing
    StoredFilePath = Nothing
    logEvent = Nothing

'Exit
    Return noErrorFlag

End Function

#End Region
#Region "LoadEncryptedText"

'*************************************************************************
Public Function Decrypt(ByVal storedFilePath As String _
                                    ) As String
'PURPOSE:
'   Write the encrypted text to a text file
'
'RETURNS:
'   True if sucessful.
'*************************************************************************
Dim logEvent As String = "Loaded Encrypted Text: " & storedFilePath
Dim returnValue As String
Dim bytEncryptedText As Byte()

    Try
        Dim fInfo As New FileInfo(storedFilePath)
        Dim numBytes As Long = fInfo.Length
        Dim fStream As New FileStream(storedFilePath, FileMode.Open, FileAccess.Read)
        Dim bReader As New BinaryReader(fStream)
        bytEncryptedText = bReader.ReadBytes(CInt(numBytes))
        returnValue = Decrypting(bytEncryptedText)
    Catch ex As Exception
        returnValue = False
        Except = ex.Message
    End Try

    If LogsOnOff Then
        If Not Except = "" Then
            logEvent = logEvent & "Failure: " & logEvent & vbCrLf & Except
            Pnm.Logs.addEvent(logEvent)
        Else
            logEvent = "Success: " & logEvent
            Pnm.Logs.addEvent(logEvent)
        End If
    End If

'Cleanup
    storedFilePath = Nothing
    bytEncryptedText = Nothing
    logEvent = Nothing

'Exit
    Return returnValue

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