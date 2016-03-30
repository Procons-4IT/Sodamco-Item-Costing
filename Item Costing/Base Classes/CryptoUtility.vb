Imports System.Security.Cryptography
Imports System.IO

Public Class CryptoUtility

    Dim KEY_64 = New Byte() {42, 16, 93, 156, 78, 4, 218, 32}
    Dim IV_64 = New Byte() {55, 103, 246, 79, 36, 99, 167, 3}

    ' <summary>
    '      Encrypt the parameter provided to this method
    '  </summary>
    '   <param name="value"></param>
    '     <returns></returns>
    Public Function Encrypt(value As String) As String
        If Not value.Equals(String.Empty) Then
            Dim desCryptoServiceProvider As DESCryptoServiceProvider = New DESCryptoServiceProvider()
            Dim memoryStream As MemoryStream = New MemoryStream()
            Dim cryptoStream As CryptoStream = New CryptoStream(memoryStream, desCryptoServiceProvider.CreateEncryptor(KEY_64, IV_64), CryptoStreamMode.Write)
            Dim streamWritter As StreamWriter = New StreamWriter(cryptoStream)
            streamWritter.Write(value)
            streamWritter.Flush()
            cryptoStream.FlushFinalBlock()
            memoryStream.Flush()
            Return Convert.ToBase64String(memoryStream.GetBuffer(), 0, Convert.ToInt32(memoryStream.Length))

        End If

        Return String.Empty
    End Function

    '    /// <summary>
    '/// Decrypt the parameter provided to this method
    '/// </summary>
    '/// <param name="value"></param>
    '/// <returns></returns>

    Public Function Decrypt(value As String) As String
        If Not value.Equals(String.Empty) Then
            Dim desCryptoServiceProvider As DESCryptoServiceProvider = New DESCryptoServiceProvider()
            Dim buffer() As Byte = Convert.FromBase64String(value)
            Dim memoryStream As MemoryStream = New MemoryStream(buffer)
            Dim cryptoStream As CryptoStream = New CryptoStream(memoryStream, desCryptoServiceProvider.CreateDecryptor(KEY_64, IV_64), CryptoStreamMode.Read)
            Dim streamReader As StreamReader = New StreamReader(cryptoStream)
            Return streamReader.ReadToEnd()

        End If

        Return String.Empty
    End Function


End Class
