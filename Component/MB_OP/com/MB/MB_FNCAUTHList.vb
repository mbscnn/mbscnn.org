Imports com.Azion.NET.VB

Public Class MB_FNCAUTHList
    Inherits BosList

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.New("MB_FNCAUTH", dbManager)
    End Sub

    Overrides Function newBos() As BosBase
        Dim bos As MB_FNCAUTH = New MB_FNCAUTH(MyBase.getDatabaseManager)
        Return bos
    End Function

#Region "Ted Function"
    ''' <summary>
    ''' 根據"帳號"取得資料
    ''' </summary>
    ''' <param name="sMB_ACCT">帳號</param>
    ''' <returns>Boolean 是否有取得</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Function LoadByACCT_CODE(ByVal sMB_ACCT As String, ByVal iMB_UPCODE As Integer) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT * FROM MB_FNCAUTH WHERE MB_ACCT = " & ProviderFactory.PositionPara & "MB_ACCT " & _
                     "  AND MB_UPCODE = " & ProviderFactory.PositionPara & "MB_UPCODE "

            Dim paras(1) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("MB_ACCT", sMB_ACCT)
            paras(1) = ProviderFactory.CreateDataParameter("MB_UPCODE", iMB_UPCODE)

            Return Me.loadBySQLOnlyDs(sqlStr, paras)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' 根據"帳號"刪除資料
    ''' </summary>
    ''' <param name="sMB_ACCT">帳號</param>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Sub DelByACCT(ByVal sMB_ACCT As String)
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "DELETE FROM MB_FNCAUTH WHERE MB_ACCT = " & ProviderFactory.PositionPara & "MB_ACCT "

            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_ACCT", sMB_ACCT)

            DBObject.ExecuteNonQuery(Me.getDatabaseManager, CommandType.Text, sqlStr, para)
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
