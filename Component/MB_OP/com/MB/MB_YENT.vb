Imports com.Azion.NET.VB

Public Class MB_YENT
    Inherits BosBase

    Sub New()
        MyBase.New()
        Me.setPrimaryKeys()
    End Sub

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.New("MB_YENT", dbManager)
    End Sub

#Region "Ted Function"
    Public Function getMAX_SDATE() As Decimal
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT MAX(SDATE) FROM MB_YENT"

            Dim D_SDATE As Object = Convert.DBNull
            If Me.getDatabaseManager.isTransaction Then
                D_SDATE = DBObject.ExecuteScalar(Me.getDatabaseManager.getTransaction.Connection, CommandType.Text, sqlStr)
            Else
                D_SDATE = DBObject.ExecuteScalar(Me.getDatabaseManager.getConnection, CommandType.Text, sqlStr)
            End If

            If IsDate(D_SDATE.ToString) Then
                Return CDate(D_SDATE.ToString).Year
            Else
                Return 0
            End If
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function
#End Region

End Class
