Option Explicit On
Option Strict On

Imports com.Azion.NET.VB
Imports com.Azion.EloanUtility

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
        'Public Function SetPassword(ByVal pwd1 As String, ByVal pwd2 As String) As Integer
        '    Try
        '        Dbg.Assert(Not IsNothing(pwd1))
        '        Dbg.Assert(Not IsNothing(pwd2))

        '        Pwd1 = EncryptUtility.MD5(Pwd1)
        '        Pwd2 = EncryptUtility.MD5(Pwd2)

        '        Dim nRecord As Integer

        '        nRecord = InsertUpdate(New BosParameter() _
        '                    {PARAMETER("ROOTNAME", "ROOT"), _
        '                     PARAMETER("PASSWORD1", pwd1), _
        '                     PARAMETER("PASSWORD2", pwd2)} _
        '            )

        '        Dbg.Assert(nRecord = 1)

        '        Return nRecord

        '    Catch ex As Exception
        '        Throw
        '    End Try
        'End Function



        ''' <summary>
        ''' 檢查密碼是否正確
        ''' </summary>
        ''' <param name="pwd1"></param>
        ''' <param name="pwd2"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        'Public Function ValidPasswords(ByRef pwd1 As String, pwd2 As String) As Boolean

        '    Try
        '        Dim record As DataRow

        '        pwd1 = EncryptUtility.MD5(pwd1)
        '        pwd2 = EncryptUtility.MD5(pwd2)

        '        record = GetDataRow( _
        '            "SELECT PASSWORD1, PASSWORD2 FROM SY_ROOT WHERE ROOTNAME = @ROOTNAME@", _
        '            New BosParameter() _
        '                    {PARAMETER("ROOTNAME", "ROOT")} _
        '                    )

        '        Dbg.Assert(Not IsNothing(record))

        '        If pwd1.CompareTo(record("PASSWORD1")) = 0 AndAlso _
        '            pwd2.CompareTo(record("PASSWORD2")) = 0 Then
        '            Return True
        '        End If

        '        Return False

        '    Catch ex As Exception
        '        Throw
        '    End Try

        'End Function

    End Class

End Namespace



