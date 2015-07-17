Imports com.Azion.NET.VB

Public Class MB_ACCT_DTL
    Inherits BosBase

    Sub New()
        MyBase.New()
        Me.setPrimaryKeys()
    End Sub

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.new("MB_ACCT_DTL", dbManager)
    End Sub

#Region "Ted Function"
    ''' ----------------------------------------------------------------------------
    ''' <summary>
    ''' 根據[傳票號碼-西元年][傳票號碼][序號]取得資料
    ''' </summary>
    ''' <param name="iMB_VOUCHER_Y">MB_VOUCHER_Y傳票號碼-西元年</param>
    ''' <param name="iMB_VOUCHER_N">MB_VOUCHER_N傳票號碼</param>
    ''' <param name="iMB_SEQ">序號</param>
    ''' <returns>Boolean 是否有取得</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Function LoadByPK(ByVal iMB_VOUCHER_Y As Decimal, ByVal iMB_VOUCHER_N As Decimal, ByVal iMB_SEQ As Decimal) As Boolean
        Try
            Dim paras(2) As IDbDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("MB_VOUCHER_Y", iMB_VOUCHER_Y)

            paras(1) = ProviderFactory.CreateDataParameter("MB_VOUCHER_N", iMB_VOUCHER_N)

            paras(2) = ProviderFactory.CreateDataParameter("MB_SEQ", iMB_SEQ)

            Return Me.loadBySQL(paras)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function getMAX_MB_SEQ(ByVal iMB_VOUCHER_Y As Decimal, ByVal iMB_VOUCHER_N As Decimal) As Decimal
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT IFNULL(MAX(MB_SEQ), 0) AS MB_SEQ " & _
                     "  FROM MB_ACCT_DTL " & _
                     " WHERE MB_VOUCHER_Y = " & ProviderFactory.PositionPara & "MB_VOUCHER_Y AND MB_VOUCHER_N = " & ProviderFactory.PositionPara & "MB_VOUCHER_N "

            Dim paras(1) As IDbDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("MB_VOUCHER_Y", iMB_VOUCHER_Y)

            paras(1) = ProviderFactory.CreateDataParameter("MB_VOUCHER_N", iMB_VOUCHER_N)

            If Me.getDatabaseManager.isTransaction Then
                Return DBObject.ExecuteScalar(Me.getDatabaseManager.getTransaction, CommandType.Text, sqlStr, paras)
            Else
                Return DBObject.ExecuteScalar(Me.getDatabaseManager.getConnection, CommandType.Text, sqlStr, paras)
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
