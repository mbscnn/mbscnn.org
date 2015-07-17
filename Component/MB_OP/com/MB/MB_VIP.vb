Imports com.Azion.NET.VB

Public Class MB_VIP
    Inherits BosBase

    Sub New()
        MyBase.New()
        Me.setPrimaryKeys()
    End Sub

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.New("MB_VIP", dbManager)
    End Sub

#Region "Ted Function"
    ''' <summary>
    ''' 根據"會員編號","護持會員"取得資料
    ''' </summary>
    ''' <param name="iMB_MEMSEQ">會員編號</param>
    ''' <param name="sVIP">護持會員</param>
    ''' <returns>Boolean 是否有取得</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Function LoadByPK(ByVal iMB_MEMSEQ As Decimal, ByVal sVIP As String) As Boolean
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT * FROM MB_VIP WHERE MB_MEMSEQ = " & ProviderFactory.PositionPara & "MB_MEMSEQ " & _
                     "  AND VIP = " & ProviderFactory.PositionPara & "VIP "

            Dim paras(1) As IDbDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("MB_MEMSEQ", iMB_MEMSEQ)

            paras(1) = ProviderFactory.CreateDataParameter("VIP", sVIP)

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
