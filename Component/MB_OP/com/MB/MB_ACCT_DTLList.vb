Imports com.Azion.NET.VB

Public Class MB_ACCT_DTLList
    Inherits BosList

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.New("MB_ACCT_DTL", dbManager)
    End Sub

    Overrides Function newBos() As BosBase
        Dim bos As MB_ACCT_DTL = New MB_ACCT_DTL(MyBase.getDatabaseManager)
        Return bos
    End Function

#Region "Ted Function"
    ''' ----------------------------------------------------------------------------
    ''' <summary>
    ''' 根據[傳票號碼-西元年][傳票號碼]取得資料
    ''' </summary>
    ''' <param name="iMB_VOUCHER_Y">MB_VOUCHER_Y傳票號碼-西元年</param>
    ''' <param name="iMB_VOUCHER_N">MB_VOUCHER_N傳票號碼</param>
    ''' <returns>Boolean 是否有取得</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Function LoadByVOUCHER(ByVal iMB_VOUCHER_Y As Decimal, ByVal iMB_VOUCHER_N As Decimal) As Integer
        Try
            Dim paras(1) As IDbDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("MB_VOUCHER_Y", iMB_VOUCHER_Y)

            paras(1) = ProviderFactory.CreateDataParameter("MB_VOUCHER_N", iMB_VOUCHER_N)

            Return Me.loadBySQL(paras)
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
