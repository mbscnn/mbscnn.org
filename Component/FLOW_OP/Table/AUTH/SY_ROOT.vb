Option Explicit On
Option Strict On

Imports com.Azion.NET.VB
Imports System.Text

Namespace TABLE

    Public Class SY_ROOT
        Inherits BosBase

        ''' <summary>
        ''' Constructor
        ''' </summary>
        ''' <param name="dbManager">DatabaseManager</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal dbManager As com.Azion.NET.VB.DatabaseManager)
            MyBase.new("SY_ROOT", dbManager)
        End Sub


        ''' <summary>
        ''' 設定新密碼
        ''' </summary>
        ''' <param name="Pwd1"></param>
        ''' <param name="pwd2"></param>
        ''' <remarks></remarks>
        Public Function SetPassword(ByVal pwd1 As String, ByVal pwd2 As String) As Integer
            Try
                Dbg.Assert(Not IsNothing(pwd1))
                Dbg.Assert(Not IsNothing(pwd2))

                pwd1 = SY_ROOT.MD5(pwd1)
                pwd2 = SY_ROOT.MD5(pwd2)

                Dim nRecord As Integer

                nRecord = InsertUpdate(
                            "ROOTNAME", "ROOT", _
                            "PASSWORD1", pwd1, _
                            "PASSWORD2", pwd2)

                Dbg.Assert(nRecord = 1)

                Return nRecord

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Function



        ''' <summary>
        ''' 檢查密碼是否正確
        ''' </summary>
        ''' <param name="pwd1"></param>
        ''' <param name="pwd2"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ValidPasswords(ByRef pwd1 As String, pwd2 As String) As Boolean

            Try
                Dim record As DataRow

                pwd1 = SY_ROOT.MD5(pwd1)
                pwd2 = SY_ROOT.MD5(pwd2)

                record = GetDataRow( _
                    "select PASSWORD1, PASSWORD2 from SY_ROOT where ROOTNAME = @ROOTNAME@", _
                    "ROOTNAME", "ROOT")

                Dbg.Assert(Not IsNothing(record))

                If pwd1.CompareTo(record("PASSWORD1")) = 0 AndAlso _
                    pwd2.CompareTo(record("PASSWORD2")) = 0 Then
                    Return True
                End If

                Return False

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

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

End Namespace



