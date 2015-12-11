Option Explicit On
Option Strict On

Imports com.Azion.NET.VB
Public Class GA_CASEAPV
    Inherits BosBase

    Sub New()
        MyBase.New()
        Me.setPrimaryKeys()
    End Sub

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.new("GA_CASEAPV", dbManager)
    End Sub

#Region "nick Function"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 根據傳入的"案件編號"取得GA_CASEAPV Record
    ''' </summary>
    ''' <param name="sAPV_CAS_ID">安泰案件編號(10碼)</param>
    ''' <returns>Boolean 是否有取得</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Nick]	2011/08/13	Created
    ''' </history>
    Public Function LoadByPK(ByVal sAPV_CAS_ID As String) As Boolean
        Try
            Dim paras As IDbDataParameter

            paras = ProviderFactory.CreateDataParameter("APV_CAS_ID", sAPV_CAS_ID)

            Return Me.loadBySQL(paras)
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 根據傳入的"案件編號"取得GA_CASEAPV Record
    ''' </summary>
    ''' <param name="sCASEID">案件編號(15碼)</param>
    ''' <returns>Boolean 是否有取得</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Nick]	2011/08/13	Created
    ''' </history>
    Public Function LoadByUNKey(ByVal sCASEID As String) As Boolean
        Try
            Dim paras As IDbDataParameter

            paras = ProviderFactory.CreateDataParameter("APV_CAS_ID_04", sCASEID)

            Return Me.loadBySQL(paras)
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

#Region "nick Function"
  ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 根據CaseID取得資料 For PageTab 
    ''' </summary>
    ''' <param name="context">頁面參數</param>
    ''' <returns>Boolean 是否取得</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2011/08/23	Created
    ''' </history>
    Function loadFunctionByTabPage(ByRef context As System.Web.HttpContext) As Boolean
        Dim sCaseId As String = context.Request.QueryString("CASEID")
        Try
            Return LoadByUNKey(sCaseId)
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
