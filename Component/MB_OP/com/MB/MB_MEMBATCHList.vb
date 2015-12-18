Imports com.Azion.NET.VB

Public Class MB_MEMBATCHList
    Inherits BosList

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.New("MB_MEMBATCH", dbManager)
    End Sub

    Overrides Function newBos() As BosBase
        Dim bos As MB_MEMBATCH = New MB_MEMBATCH(MyBase.getDatabaseManager)
        Return bos
    End Function

#Region "Ted Function"
    ''' <summary>
    ''' 根據"學員編號","課程序號"取得資料
    ''' </summary>
    ''' <param name="iMB_MEMSEQ">學員編號</param>
    ''' <param name="iMB_SEQ">課程序號</param>
    ''' <returns>Boolean</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Function LoadBySEQ(ByVal iMB_MEMSEQ As Decimal, ByVal iMB_SEQ As Decimal) As Boolean
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT * FROM MB_MEMBATCH " &
                     " WHERE MB_MEMSEQ = " & ProviderFactory.PositionPara & "MB_MEMSEQ " &
                     "  AND MB_SEQ = " & ProviderFactory.PositionPara & "MB_SEQ "

            Dim paras(1) As IDbDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("MB_MEMSEQ", iMB_MEMSEQ)
            paras(1) = ProviderFactory.CreateDataParameter("MB_SEQ", iMB_SEQ)

            Return Me.loadBySQL(sqlStr, paras)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' 根據"學員編號","課程序號"刪除資料
    ''' </summary>
    ''' <param name="iMB_MEMSEQ">學員編號</param>
    ''' <param name="iMB_SEQ">課程序號</param>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Sub DeleteBySEQ(ByVal iMB_MEMSEQ As Decimal, ByVal iMB_SEQ As Decimal)
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "DELETE FROM MB_MEMBATCH " &
                     " WHERE MB_MEMSEQ = " & ProviderFactory.PositionPara & "MB_MEMSEQ " &
                     "  AND MB_SEQ = " & ProviderFactory.PositionPara & "MB_SEQ "

            Dim paras(1) As IDbDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("MB_MEMSEQ", iMB_MEMSEQ)
            paras(1) = ProviderFactory.CreateDataParameter("MB_SEQ", iMB_SEQ)

            DBObject.ExecuteNonQuery(Me.getDatabaseManager, CommandType.Text, sqlStr, paras)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Sub
#End Region

End Class
