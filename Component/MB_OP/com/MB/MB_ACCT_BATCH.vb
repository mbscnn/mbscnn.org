Imports com.Azion.NET.VB

Public Class MB_ACCT_BATCH
    Inherits BosBase

    Sub New()
        MyBase.New()
        Me.setPrimaryKeys()
    End Sub

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.New("MB_ACCT_BATCH", dbManager)
    End Sub

    Public Function LoadByPK(ByVal sMB_ACCT As String) As Boolean
        Try
            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_ACCT", sMB_ACCT)
            Return Me.loadBySQL(para)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function
End Class
