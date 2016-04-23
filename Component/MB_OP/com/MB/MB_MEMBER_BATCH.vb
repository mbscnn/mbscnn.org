Imports com.Azion.NET.VB

Public Class MB_MEMBER_BATCH
    Inherits BosBase

    Sub New()
        MyBase.New()
        Me.setPrimaryKeys()
    End Sub

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.new("MB_MEMBER_BATCH", dbManager)
    End Sub

    Function loadByPK(ByVal iMB_MEMSEQ As Decimal) As Boolean
        Try
            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_MEMSEQ", iMB_MEMSEQ)

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
