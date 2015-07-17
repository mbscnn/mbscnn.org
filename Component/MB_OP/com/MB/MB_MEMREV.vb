Imports com.Azion.NET.VB

Public Class MB_MEMREV
    Inherits BosBase

    Sub New()
        MyBase.New()
        Me.setPrimaryKeys()
    End Sub

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.new("MB_MEMREV", dbManager)
    End Sub

#Region "Ted Function"
    ''' ----------------------------------------------------------------------------
    ''' <summary>
    ''' 根據[會員編號][繳費序號]取得資料
    ''' </summary>
    ''' <param name="iMB_MEMSEQ">MB_MEMSEQ會員編號</param>
    ''' <param name="iMB_SEQNO">MB_SEQNO會員編號</param>
    ''' <returns>Boolean 是否有取得</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Function loadByPK(ByVal iMB_MEMSEQ As Decimal, ByVal iMB_SEQNO As Decimal) As Boolean
        Try
            Dim paras(1) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("MB_MEMSEQ", iMB_MEMSEQ)
            paras(1) = ProviderFactory.CreateDataParameter("MB_SEQNO", iMB_SEQNO)

            Return Me.loadBySQL(paras)
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' ----------------------------------------------------------------------------
    ''' <summary>
    ''' 根據[會員編號]取得最大繳費序號+1
    ''' </summary>
    ''' <param name="iMB_MEMSEQ">MB_MEMSEQ會員編號</param>
    ''' <returns>decimal 最大繳費序號+1</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Function getMB_SEQNOPlus1(ByVal iMB_MEMSEQ As Decimal) As Decimal
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "select ifnull(max(MB_SEQNO),0)+1 from mb_memrev where MB_MEMSEQ=" & ProviderFactory.PositionPara & "MB_MEMSEQ"

            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_MEMSEQ", iMB_MEMSEQ)

            Return Me.loadBySQLScalar(sqlStr, para)
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
