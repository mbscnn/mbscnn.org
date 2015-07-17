Imports com.Azion.NET.VB

Public Class MB_MBIMP
    Inherits BosBase

    Sub New()
        MyBase.New()
        Me.setPrimaryKeys()
    End Sub

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.new("MB_MBIMP", dbManager)
    End Sub

#Region "Ted Function"
    ''' ----------------------------------------------------------------------------
    ''' <summary>
    ''' 根據[會員編號]取得資料
    ''' </summary>
    ''' <param name="iMB_MEMSEQ">MB_MEMSEQ會員編號</param>
    ''' <returns>Boolean 是否有取得</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Function LoadByPK(ByVal iMB_MEMSEQ As Decimal) As Boolean
        Try
            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_MEMSEQ", iMB_MEMSEQ)

            Return Me.loadBySQL(para)
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

End Class
