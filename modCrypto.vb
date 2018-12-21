Imports System.Security.Cryptography

' modCrypto cannot be disabled and doesn't need to be loaded or unloaded

Module modCrypto
    'Initial basis coming from Ussiane Lepsiq here: https://stackoverflow.com/a/37413712

    ''' <summary>
    ''' Decrypts a string
    ''' </summary>
    ''' <param name="strCiphertext">Text to be decrypted</param>
    ''' <param name="strKey">Encryption key</param>
    ''' <param name="intVer">Crypto version for futureproofing</param>
    ''' <returns></returns>
    Function Decrypt(ByVal strCiphertext As String, ByVal strKey As String, Optional ByVal intVer As Integer = 0) As String
        Dim AES As New RijndaelManaged
        Dim SHA256 As New SHA256Cng
        Dim strPlaintext As String = ""
        Dim strIV As String = ""
        Try
            Dim ivct = strCiphertext.Split({"=="}, StringSplitOptions.None)
            strIV = ivct(0) & "=="
            strCiphertext = If(ivct.Length = 3, ivct(1) & "==", ivct(1))
            AES.Key = SHA256.ComputeHash(Text.Encoding.ASCII.GetBytes(strKey))
            AES.IV = Convert.FromBase64String(strIV)
            AES.Mode = CipherMode.CBC
            Dim DESDecryptor As ICryptoTransform = AES.CreateDecryptor
            Dim Buffer As Byte() = Convert.FromBase64String(strCiphertext)
            strPlaintext = Text.Encoding.ASCII.GetString(DESDecryptor.TransformFinalBlock(Buffer, 0, Buffer.Length))
            Return strPlaintext
        Catch CryptoExcep As CryptographicException
            My.Application.Log.WriteException(CryptoExcep)
            Return "CRYPTO_FAILED"
        End Try
    End Function

    ''' <summary>
    ''' Encrypts a string
    ''' </summary>
    ''' <param name="strPlaintext">Text to be encrypted</param>
    ''' <param name="strKey">Encryption key</param>
    ''' <param name="intVer">Crypto version for futureproofing</param>
    ''' <returns></returns>
    Function Encrypt(ByVal strPlaintext As String, ByVal strKey As String, Optional ByVal intVer As Integer = 0) As String
        Dim AES As New RijndaelManaged
        Dim SHA256 As New SHA256Cng
        Dim strCiphertext As String = ""
        Try
            AES.GenerateIV()
            AES.Key = SHA256.ComputeHash(Text.Encoding.ASCII.GetBytes(strKey))
            AES.Mode = CipherMode.CBC
            Dim DESEncryptor As ICryptoTransform = AES.CreateEncryptor
            Dim Buffer As Byte() = Text.Encoding.ASCII.GetBytes(strPlaintext)
            strCiphertext = Convert.ToBase64String(DESEncryptor.TransformFinalBlock(Buffer, 0, Buffer.Length))
            Return Convert.ToBase64String(AES.IV) & Convert.ToBase64String(DESEncryptor.TransformFinalBlock(Buffer, 0, Buffer.Length))
        Catch CryptoExcep As CryptographicException
            My.Application.Log.WriteException(CryptoExcep)
            Return "CRYPTO_FAILED"
        End Try
    End Function

    ''' <summary>
    ''' Encrypts and then decrypts a string to ensure cryptographic functions are available and working.
    ''' </summary>
    ''' <param name="intVer">Crypto version for futureproofing</param>
    ''' <returns></returns>
    Function TestCrypto(Optional ByVal intVer As Integer = 0) As String
        Dim strTestText As String = "This is a test string for the cryptographic module."
        Dim strTestKey As String = "Encrypt-00-Test-0"
        My.Application.Log.WriteEntry("Running cryptographic test", TraceEventType.Information)
        Dim strTestOutput As String = Encrypt(strTestText, strTestKey, intVer)
        My.Application.Log.WriteEntry("Ciphertext output: " & strTestOutput, TraceEventType.Verbose)
        Dim strTestDecrypt As String = Decrypt(strTestOutput, strTestKey, intVer)
        My.Application.Log.WriteEntry("Decrypted output: " & strTestDecrypt, TraceEventType.Verbose)
        Return "Cryptographic module test complete"
    End Function
End Module
