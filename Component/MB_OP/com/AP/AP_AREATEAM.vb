Imports com.Azion.NET.VB

Public Class AP_AREATEAM
    Inherits BosBase

    Sub New()
        MyBase.New()
        Me.setPrimaryKeys()
    End Sub

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.new("AP_AREATEAM", dbManager)
    End Sub

#Region "Ted Function"
    ''' ----------------------------------------------------------------------------
    ''' <summary>
    ''' 根據[縣市代碼]取得資料
    ''' </summary>
    ''' <param name="sA_CITY_ID">A_CITY_ID縣市代碼</param>
    ''' <returns>Boolean 是否有取得</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Function loadByPK(ByVal sA_CITY_ID As String) As Boolean
        Try
            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("A_CITY_ID", sA_CITY_ID)

            Return Me.loadBySQL(para)
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

#End Region

End Class
