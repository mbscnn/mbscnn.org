Imports com.Azion.NET.VB

Public Class SY_BRANCH_HIS
    Inherits BosBase

    Public Sub New(ByVal dbManager As com.Azion.NET.VB.DatabaseManager)
        MyBase.new("SY_BRANCH_HIS", dbManager)
    End Sub

    Sub New()
        MyBase.New()
        Me.setPrimaryKeys()
    End Sub

#Region "濟南昱勝添加"
#Region "Lake Function"
    ''' <summary>
    ''' 根據主鍵查詢部門歷史表
    ''' </summary>
    ''' <param name="sBraDepNO">部門代碼</param>
    ''' <param name="sCaseId">案件編號</param>
    ''' <param name="sStepNo">流程工作步驟編號</param>
    ''' <param name="sSubFlowSeq">子流程序號</param>
    ''' <param name="sSubFlowCount">已執行的子流程數量</param>
    ''' <returns></returns>
    ''' <remarks>[Lake] 2012/05/17 Created</remarks>
    Public Function loadByPK(ByVal sBraDepNO As String, ByVal sCaseId As String, ByVal sStepNo As String, ByVal sSubFlowSeq As String, ByVal sSubFlowCount As String) As Boolean
        Try
            Dim paras(4) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("BRA_DEPNO", sBraDepNO)
            paras(1) = ProviderFactory.CreateDataParameter("CASEID", sCaseId)
            paras(2) = ProviderFactory.CreateDataParameter("STEP_NO", sStepNo)
            paras(3) = ProviderFactory.CreateDataParameter("SUBFLOW_SEQ", sSubFlowSeq)
            paras(4) = ProviderFactory.CreateDataParameter("SUBFLOW_COUNT", sSubFlowCount)

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
    ''' <param name="sBraDepNo">部門代碼</param>
    ''' <param name="sCaseId">案件編號</param>
    ''' <param name="sStepNo">流程步驟</param>
    ''' <param name="sSubFlowCount"></param>
    ''' <returns></returns>
    ''' <remarks>[Lake] 2012/07/02 Created</remarks>
    Public Function getMaxSubFlowSeq(ByVal sBraDepNo As String, ByVal sCaseId As String, _
                                     ByVal sStepNo As String, ByVal sSubFlowCount As String) As String
        Try
            Dim sSql As String = "SELECT ISNULL(MAX(SUBFLOW_SEQ),0) AS SUBFLOW_SEQ FROM SY_BRANCH_HIS " & _
                                 "WHERE BRA_DEPNO = " & ProviderFactory.PositionPara & "BRADEPNO " & _
                                 "AND CASEID = " & ProviderFactory.PositionPara & "CASEID " & _
                                 "AND STEP_NO = " & ProviderFactory.PositionPara & "STEPNO " & _
                                 "AND SUBFLOW_COUNT = " & ProviderFactory.PositionPara & "SUBFLOWCOUNT"

            Dim paras(3) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("BRADEPNO", sBraDepNo)
            paras(1) = ProviderFactory.CreateDataParameter("CASEID", sCaseId)
            paras(2) = ProviderFactory.CreateDataParameter("STEPNO", sStepNo)
            paras(3) = ProviderFactory.CreateDataParameter("SUBFLOWCOUNT", sSubFlowCount)

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

    ''' <summary>
    ''' 批量新增歷史檔
    ''' </summary>
    ''' <param name="sBraDepNo"></param>
    ''' <param name="sDisabled"></param>
    ''' <param name="sMgrBrid"></param>
    ''' <param name="sCaseId"></param>
    ''' <param name="sStepNo"></param>
    ''' <param name="sSubFlowSeq"></param>
    ''' <param name="sSubFlowCount"></param>
    ''' <returns></returns>
    ''' <remarks>[Lake] 2012/07/14 Created</remarks>
    Public Function insertByParas(ByVal sBraDepNo As String, ByVal sDisabled As String, ByVal sMgrBrid As String _
                                , Optional ByVal sCaseId As String = "", Optional ByVal sStepNo As String = "" _
                                , Optional ByVal sSubFlowSeq As String = "", Optional ByVal sSubFlowCount As String = "") As Boolean
        Try
            Dim sSql As String = String.Empty
            Dim paras(7) As IDbDataParameter

            sSql = "insert into sy_branch_his select " & ProviderFactory.PositionPara & "BRADEPNO " & _
                   "," & ProviderFactory.PositionPara & "CASEID " & _
                   "," & ProviderFactory.PositionPara & "STEPNO " & _
                   "," & ProviderFactory.PositionPara & "SUBFLOWSEQ " & _
                   "," & ProviderFactory.PositionPara & "SUBFLOWCOUNT " & _
                   ",BRID,PARENT,BRCNAME,BRCADDR,BRATEL,BRTEL,NULL,NULL,NULL " & _
                   "," & ProviderFactory.PositionPara & "DISABLED " & _
                   ",BRCCITY,BRCAREA,HOFLAG,EPSDEP,'Y',NULL,'U' " & _
                   "," & ProviderFactory.PositionPara & "MGRBRID " & _
                   "from sy_branch where bra_depno = " & ProviderFactory.PositionPara & "BRADEPNO "

            paras(0) = ProviderFactory.CreateDataParameter("CASEID", sCaseId)
            paras(1) = ProviderFactory.CreateDataParameter("STEPNO", sStepNo)
            paras(2) = ProviderFactory.CreateDataParameter("SUBFLOWSEQ", sSubFlowSeq)
            paras(3) = ProviderFactory.CreateDataParameter("SUBFLOWCOUNT", sSubFlowCount)
            paras(4) = ProviderFactory.CreateDataParameter("BRADEPNO", sBraDepNo)
            paras(5) = ProviderFactory.CreateDataParameter("DISABLED", sDisabled)
            paras(6) = ProviderFactory.CreateDataParameter("MGRBRID", sMgrBrid)

            If Me.getDatabaseManager.isTransaction Then
                Return DBObject.ExecuteNonQuery(Me.getDatabaseManager.getTransaction, CommandType.Text, sSql, paras) > 0
            Else
                Return DBObject.ExecuteNonQuery(Me.getDatabaseManager.getConnection, CommandType.Text, sSql, paras) > 0
            End If
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
