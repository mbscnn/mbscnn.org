Imports com.Azion.NET.VB

Public Class MB_FNCAUTH
    Inherits BosBase

    Sub New()
        MyBase.New()
        Me.setPrimaryKeys()
    End Sub

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.New("MB_FNCAUTH", dbManager)
    End Sub

#Region "Ted Function"
    ''' <summary>
    ''' 根據"帳號","父類別","功能代號"取得資料
    ''' </summary>
    ''' <param name="sMB_ACCT">帳號</param>
    ''' <param name="iMB_UPCODE">父類別</param>
    ''' <param name="sMB_FUCCODE">功能代號</param>
    ''' <returns>Boolean 是否有取得</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Function LoadByPK(ByVal sMB_ACCT As String, ByVal iMB_UPCODE As Integer, ByVal sMB_FUCCODE As String) As Boolean
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT * FROM MB_FNCAUTH WHERE MB_ACCT = " & ProviderFactory.PositionPara & "MB_ACCT " & _
                     "  AND MB_UPCODE = " & ProviderFactory.PositionPara & "MB_UPCODE " & _
                     "  AND MB_FUCCODE = " & ProviderFactory.PositionPara & "MB_FUCCODE "

            Dim paras(2) As IDbDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("MB_ACCT", sMB_ACCT)

            paras(1) = ProviderFactory.CreateDataParameter("MB_UPCODE", iMB_UPCODE)

            paras(2) = ProviderFactory.CreateDataParameter("MB_FUCCODE", sMB_FUCCODE)

            Return Me.loadBySQL(sqlStr, paras)
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
