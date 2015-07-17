Imports com.Azion.NET.VB

Public Class MB_MEMFEE
    Inherits BosBase

    Sub New()
        MyBase.New()
        Me.setPrimaryKeys()
    End Sub

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.new("MB_MEMFEE", dbManager)
    End Sub

#Region "Ted Function"
    ''' ----------------------------------------------------------------------------
    ''' <summary>
    ''' 根據[會員編號][繳會費年度]取得資料
    ''' </summary>
    ''' <param name="iMB_MEMSEQ">MB_MEMSEQ會員編號</param>
    ''' <param name="iMB_YYYY">MB_YYYY繳會費年度</param>
    ''' <returns>Boolean 是否有取得</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Function LoadByPK(ByVal iMB_MEMSEQ As Decimal, ByVal iMB_YYYY As Decimal) As Boolean
        Try
            Dim paras(1) As IDbDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("MB_MEMSEQ", iMB_MEMSEQ)

            paras(1) = ProviderFactory.CreateDataParameter("MB_YYYY", iMB_YYYY)

            Return Me.loadBySQL(paras)
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
