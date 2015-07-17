Imports com.Azion.NET.VB

Public Class AP_AREATEAMList
    Inherits BosList

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.New("AP_AREATEAM", dbManager)
    End Sub

    Overrides Function newBos() As BosBase
        Dim bos As AP_AREATEAM = New AP_AREATEAM(MyBase.getDatabaseManager)
        Return bos
    End Function

#Region "Ted Function"
    ''' ----------------------------------------------------------------------------
    ''' <summary>
    ''' 根據[區域組別代碼]取得資料
    ''' </summary>
    ''' <param name="sGA_AREA">GA_AREA區域組別代碼</param>
    ''' <returns>Integer 取得筆數</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Function loadByGA_AREA(ByVal sGA_AREA As String) As Integer
        Try
            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("GA_AREA", sGA_AREA)

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
