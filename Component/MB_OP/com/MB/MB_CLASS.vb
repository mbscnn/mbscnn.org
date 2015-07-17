Imports com.Azion.NET.VB

Public Class MB_CLASS
    Inherits BosBase

    Sub New()
        MyBase.New()
        Me.setPrimaryKeys()
    End Sub

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.new("MB_CLASS", dbManager)
    End Sub

#Region "Ted Function"
    ''' <summary>
    ''' 根據"課程序號","梯次"取得資料
    ''' </summary>
    ''' <param name="iMB_SEQ">課程序號</param>
    ''' <param name="iMB_BATCH">梯次</param>
    ''' <returns>Boolean 是否有取得</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Function loadByPK(ByVal iMB_SEQ As Decimal, ByVal iMB_BATCH As Decimal) As Boolean
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT * FROM MB_CLASS WHERE MB_SEQ = " & ProviderFactory.PositionPara & "MB_SEQ " & _
                     "  AND MB_BATCH = " & ProviderFactory.PositionPara & "MB_BATCH "

            Dim paras(1) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("MB_SEQ", iMB_SEQ)
            paras(1) = ProviderFactory.CreateDataParameter("MB_BATCH", iMB_BATCH)

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
