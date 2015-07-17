Imports com.Azion.NET.VB

Public Class MB_MEMBER_TEMPList
    Inherits BosList

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.New("MB_MEMBER_TEMP", dbManager)
    End Sub

    Overrides Function newBos() As BosBase
        Dim bos As MB_MEMBER_TEMP = New MB_MEMBER_TEMP(MyBase.getDatabaseManager)
        Return bos
    End Function

    Function LoadNULL(ByVal sMB_AREA As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            If Utility.isValidateData(sMB_AREA) Then
                sqlStr = "SELECT * FROM MB_MEMBER_TEMP WHERE MB_DELFLG IS NULL AND MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA "

                Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_AREA", sMB_AREA)

                Return Me.loadBySQLOnlyDs(sqlStr, para)
            Else
                sqlStr = "SELECT * FROM MB_MEMBER_TEMP WHERE MB_DELFLG IS NULL"

                Return Me.loadBySQLOnlyDs(sqlStr)
            End If
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Function Load_D(ByVal sMB_AREA As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            If Utility.isValidateData(sMB_AREA) Then
                sqlStr = "SELECT * FROM MB_MEMBER_TEMP WHERE IFNULL(MB_DELFLG, ' ') = 'D' AND MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA "

                Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_AREA", sMB_AREA)

                Return Me.loadBySQLOnlyDs(sqlStr, para)
            Else
                sqlStr = "SELECT * FROM MB_MEMBER_TEMP WHERE IFNULL(MB_DELFLG, ' ') = 'D'"

                Return Me.loadBySQLOnlyDs(sqlStr)
            End If
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function
End Class
