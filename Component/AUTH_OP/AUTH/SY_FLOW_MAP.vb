Imports com.Azion.NET.VB

Public Class SY_FLOW_MAP
    Inherits BosBase

    Public Sub New(ByVal dbManager As com.Azion.NET.VB.DatabaseManager)
        MyBase.new("SY_FLOW_MAP", dbManager)
    End Sub

    Sub New()
        MyBase.New()
        Me.setPrimaryKeys()
    End Sub


#Region "濟南昱勝添加"
#Region "Lake Function"
    ''' <summary>
    ''' 根據案件編號及流程步驟取得資料
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>[Lake] 2012/05/12 Created</remarks>
    Public Function loadByPK(ByVal sCaseId As String, ByVal sStepNo As String) As Boolean
        Try
            Dim paras(1) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("CASEID", sCaseId)
            paras(1) = ProviderFactory.CreateDataParameter("STEP_NO", sStepNo)

            Return Me.loadData(paras)
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region
#End Region
End Class
