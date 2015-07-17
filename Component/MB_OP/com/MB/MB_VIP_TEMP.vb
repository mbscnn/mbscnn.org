Imports com.Azion.NET.VB

Public Class MB_VIP_TEMP
    Inherits BosBase

    Sub New()
        MyBase.New()
        Me.setPrimaryKeys()
    End Sub

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.New("MB_VIP_TEMP", dbManager)
    End Sub

#Region "Ted Function"
    ''' <summary>
    ''' 根據"會員編號","護持會員;1長期護持會員;2種子護法"取得資料
    ''' </summary>
    ''' <param name="iMB_DATE">會員編號</param>
    ''' <param name="sVIP">護持會員;1長期護持會員;2種子護法</param>
    ''' <returns>Boolean 是否有取得</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Function LoadByPK(ByVal iMB_DATE As Long, ByVal sVIP As String) As Boolean
        Try
            Dim paras(1) As IDbDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("MB_DATE", iMB_DATE)

            paras(1) = ProviderFactory.CreateDataParameter("VIP", sVIP)

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
