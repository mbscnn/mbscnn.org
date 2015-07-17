Option Explicit On
Option Strict On

Imports System.IO
Imports System.Text
Imports System.Security     '保持機密的文字
Imports System.Security.Cryptography '密碼編譯



Public Class EncryptUtility

    'sIV 8碼, sKey 24碼
    Public Shared Function Encrypto(ByVal pToEncrypt As String, Optional ByVal sIV As String = "enteloan", Optional ByVal sKey As String = "ibank.entiebank.com.tw@@") As String
        'Dim input_value As String
        Dim data As Byte() = Encoding.Default.GetBytes(pToEncrypt)
        Dim tdes As TripleDES = TripleDES.Create()
        tdes.IV = Encoding.ASCII.GetBytes(sIV)
        tdes.Key = Encoding.ASCII.GetBytes(sKey)
        tdes.Mode = CipherMode.ECB
        tdes.Padding = PaddingMode.Zeros
        Dim ict As ICryptoTransform = tdes.CreateEncryptor()
        Dim enc As Byte() = ict.TransformFinalBlock(data, 0, data.Length)
        Return Convert.ToBase64String(enc).Trim
    End Function


    Public Shared Function Decrypto(ByVal pToDecrypt As String, Optional ByVal sIV As String = "enteloan", Optional ByVal sKey As String = "ibank.entiebank.com.tw@@") As String
        Dim data As Byte() = Convert.FromBase64String(pToDecrypt.Trim)
        Dim tdes As TripleDES = TripleDES.Create()
        tdes.IV = Encoding.ASCII.GetBytes(sIV)
        tdes.Key = Encoding.ASCII.GetBytes(sKey)
        tdes.Mode = CipherMode.ECB
        tdes.Padding = PaddingMode.Zeros
        Dim ict As ICryptoTransform = tdes.CreateDecryptor()
        Dim enc As Byte() = ict.TransformFinalBlock(data, 0, data.Length)
        Dim s As String = Encoding.ASCII.GetString(enc).Trim
        Return s
    End Function


    ''' <summary>
    ''' 計算字串的MD5值
    ''' </summary>
    ''' <param name="input">要計算MD5的原始字串</param>
    ''' <returns>傳回MD5值，以字串的形式傳回</returns>
    ''' <remarks></remarks>
    Public Shared Function MD5(input As String) As String
        ' Use input string to calculate MD5 hash
        Dim md5value As System.Security.Cryptography.MD5 = System.Security.Cryptography.MD5.Create()
        Dim inputBytes As Byte() = System.Text.Encoding.ASCII.GetBytes(input)
        Dim hashBytes As Byte() = md5value.ComputeHash(inputBytes)

        ' Convert the byte array to hexadecimal string
        Dim sb As New StringBuilder()
        For i As Integer = 0 To hashBytes.Length - 1
            ' To force the hex string to lower-case letters instead of
            ' upper-case, use he following line instead:
            ' sb.Append(hashBytes[i].ToString("x2")); 
            sb.Append(hashBytes(i).ToString("X2"))
        Next
        Return sb.ToString()
    End Function

End Class


