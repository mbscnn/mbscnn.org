Imports com.Azion.NET.VB

Public Class MB_MEMBATCH
    Inherits BosBase

    Sub New()
        MyBase.New()
        Me.setPrimaryKeys()
    End Sub

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.New("MB_MEMBATCH", dbManager)
    End Sub

#Region "Ted Function"
    ''' <summary>
    ''' 根據"學員編號","課程序號","梯次"取得資料
    ''' </summary>
    ''' <param name="iMB_MEMSEQ">學員編號</param>
    ''' <param name="iMB_SEQ">課程序號</param>
    ''' <param name="iMB_BATCH">課程序號</param>
    ''' <returns>Boolean</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Function LoadByPK(ByVal iMB_MEMSEQ As Decimal, ByVal iMB_SEQ As Decimal, ByVal iMB_BATCH As Decimal) As Boolean
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT * FROM MB_MEMBATCH " &
                     " WHERE MB_MEMSEQ = " & ProviderFactory.PositionPara & "MB_MEMSEQ " &
                     "  AND MB_SEQ = " & ProviderFactory.PositionPara & "MB_SEQ " &
                     "  AND MB_BATCH = " & ProviderFactory.PositionPara & "MB_BATCH "

            Dim paras(2) As IDbDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("MB_MEMSEQ", iMB_MEMSEQ)
            paras(1) = ProviderFactory.CreateDataParameter("MB_SEQ", iMB_SEQ)
            paras(2) = ProviderFactory.CreateDataParameter("MB_BATCH", iMB_BATCH)

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
