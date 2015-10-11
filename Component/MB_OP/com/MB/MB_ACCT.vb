Imports com.Azion.NET.VB

Public Class MB_ACCT
    Inherits BosBase

    Sub New()
        MyBase.New()
        Me.setPrimaryKeys()
    End Sub

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.New("MB_ACCT", dbManager)
    End Sub

#Region "Ted Function"
    ''' ----------------------------------------------------------------------------
    ''' <summary>
    ''' 根據"帳號(Mail)"取得資料
    ''' </summary>
    ''' <param name="sMB_ACCT">帳號(Mail)</param>
    ''' <returns>Boolean 是否有取得</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Function LoadByPK(ByVal sMB_ACCT As String) As Boolean
        Try
            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_ACCT", sMB_ACCT)
            Return Me.loadBySQL(para)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' 根據"驗證用ID"取得資料
    ''' </summary>
    ''' <param name="sMB_APVID">驗證用ID</param>
    ''' <returns>Boolean 是否有取得</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Function LoadByMB_APVID(ByVal sMB_APVID As String) As Boolean
        Try
            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_APVID", sMB_APVID)
            Return Me.loadBySQL(para)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' 根據"驗證用ID"取得資料
    ''' </summary>
    ''' <param name="sMB_PASSVID">密碼重設驗證用ID</param>
    ''' <returns>Boolean 是否有取得</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Function LoadByMB_PASSVID(ByVal sMB_PASSVID As String) As Boolean
        Try
            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_PASSVID", sMB_PASSVID)
            Return Me.loadBySQL(para)
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
