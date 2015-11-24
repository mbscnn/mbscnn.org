Imports com.Azion.NET.VB

Public Class SY_REL_BRANCH_USER_HIS
    Inherits BosBase

    Public Sub New(ByVal dbManager As com.Azion.NET.VB.DatabaseManager)
        MyBase.new("SY_REL_BRANCH_USER_HIS", dbManager)
    End Sub

    Sub New()
        MyBase.New()
        Me.setPrimaryKeys()
    End Sub

#Region "濟南昱勝添加"
#Region "Lake Function"
    ''' <summary>
    ''' 根據主鍵查詢
    ''' </summary>
    ''' <param name="sStaffId">員工編號</param>
    ''' <param name="sBraDepNo">部門代碼</param>
    ''' <param name="sCaseId">案件編號</param>
    ''' <param name="sStepNo">流程步驟</param>
    ''' <param name="sSubFlowSeq"></param>
    ''' <param name="sSubFlowCount"></param>
    ''' <returns></returns>
    ''' <remarks>[Lake] 2012/05/24 Created</remarks>
    Public Function loadByPk(ByVal sStaffId As String, ByVal sBraDepNo As String, _
                             ByVal sCaseId As String, ByVal sStepNo As String, _
                             ByVal sSubFlowSeq As String, ByVal sSubFlowCount As String) As Boolean
        Try
            Dim paras(5) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("STAFFID", sStaffId)
            paras(1) = ProviderFactory.CreateDataParameter("BRA_DEPNO", sBraDepNo)
            paras(2) = ProviderFactory.CreateDataParameter("CASEID", sCaseId)
            paras(3) = ProviderFactory.CreateDataParameter("STEP_NO", sStepNo)
            paras(4) = ProviderFactory.CreateDataParameter("SUBFLOW_SEQ", sSubFlowSeq)
            paras(5) = ProviderFactory.CreateDataParameter("SUBFLOW_COUNT", sSubFlowCount)

            Return Me.loadData(paras)
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 根據參數取得最大SUBFLOW_SEQ
    ''' </summary>
    ''' <param name="sStaffId">員工編號</param>
    ''' <param name="sBraDepNo">部門代碼</param>
    ''' <param name="sCaseId">案件編號</param>
    ''' <param name="sStepNo">流程步驟</param>
    ''' <param name="sSubFlowCount"></param>
    ''' <returns></returns>
    ''' <remarks>[Lake] 2012/07/02 Created</remarks>
    Public Function getMaxSubFlowSeq(ByVal sStaffId As String, ByVal sBraDepNo As String, ByVal sCaseId As String, _
                                     ByVal sStepNo As String, ByVal sSubFlowCount As String) As String
        Try
            Dim sSql As String = "SELECT ISNULL(MAX(SUBFLOW_SEQ),0) AS SUBFLOW_SEQ FROM SY_REL_BRANCH_USER_HIS " & _
                                 "WHERE STAFFID = " & ProviderFactory.PositionPara & "STAFFID " & _
                                 "AND BRA_DEPNO = " & ProviderFactory.PositionPara & "BRADEPNO " & _
                                 "AND CASEID = " & ProviderFactory.PositionPara & "CASEID " & _
                                 "AND STEP_NO = " & ProviderFactory.PositionPara & "STEPNO " & _
                                 "AND SUBFLOW_COUNT = " & ProviderFactory.PositionPara & "SUBFLOWCOUNT"

            Dim paras(4) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("STAFFID", sStaffId)
            paras(1) = ProviderFactory.CreateDataParameter("BRADEPNO", sBraDepNo)
            paras(2) = ProviderFactory.CreateDataParameter("CASEID", sCaseId)
            paras(3) = ProviderFactory.CreateDataParameter("STEPNO", sStepNo)
            paras(4) = ProviderFactory.CreateDataParameter("SUBFLOWCOUNT", sSubFlowCount)

            loadBySQL(sSql, paras)

            Return Me.getAttribute("SUBFLOW_SEQ").ToString()
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
