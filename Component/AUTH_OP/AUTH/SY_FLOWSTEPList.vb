Imports com.Azion.NET.VB

Public Class SY_FLOWSTEPList
    Inherits BosList

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.New("SY_FLOWSTEP", dbManager)
    End Sub

    Overrides Function newBos() As BosBase
        Dim bos As SY_FLOWSTEP = New SY_FLOWSTEP(MyBase.getDatabaseManager)
        Return bos
    End Function

#Region "濟南昱勝添加"
#Region "Lake Function"
    ''' <summary>
    ''' 根據流程案件編號取得流程案件狀態
    ''' </summary>
    ''' <param name="sCaseId">流程案件編號</param>
    ''' <returns></returns>
    ''' <remarks>[Lake] 2012/05/10 Created</remarks>
    Public Function getFlowCaseStatus(ByVal sCaseId As String) As Boolean
        Try
            Dim sSql As String = _
                "select (ssn.STEP_NAME + case " & vbCrLf & _
                "         when syf.REVISION_SEQNO is null then " & vbCrLf & _
                "          '[送出]' " & vbCrLf & _
                "         else " & vbCrLf & _
                "          '[退關]' " & vbCrLf & _
                "       end) AS STEPNAME," & vbCrLf & _
                "       syf.CLIENT," & vbCrLf & _
                "       syf.SENDER," & vbCrLf & _
                "       (case when (syf.CLIENT<>syf.SENDER) then t1.USERNAME + '-[' + syf.SENDER + ']' else '' end) AS USERNAME_AGENT, " & vbCrLf & _
                "       syu.USERNAME + '-[' + syf.CLIENT + ']' AS USERNAME, " & vbCrLf & _
                "       syu.OFFICETEL, " & vbCrLf & _
                "       syf.STARTTIME, " & vbCrLf & _
                "       case " & vbCrLf & _
                "         when syf.STATUS = '3' then " & vbCrLf & _
                "          '已完成' " & vbCrLf & _
                "         else " & vbCrLf & _
                "          '處理中' " & vbCrLf & _
                "       end STATUS " & vbCrLf & _
                "  from SY_FLOWSTEP syf " & vbCrLf & _
                "  left join SY_USER syu " & vbCrLf & _
                "    on syf.CLIENT = syu.STAFFID " & vbCrLf & _
                "  join SY_STEP_NO ssn " & vbCrLf & _
                "    on ssn.STEP_NO = syf.STEP_NO " & vbCrLf & _
                "  join SY_CASEID cas " & vbCrLf & _
                "    on cas.CASEID = syf.CASEID " & vbCrLf & _
                "  join SY_FLOW_MAP fmp " & vbCrLf & _
                "    on fmp.FLOW_ID = cas.FLOW_ID " & vbCrLf & _
                "   and fmp.STEP_NO = ssn.STEP_NO " & vbCrLf & _
                "  left join SY_USER t1 " & vbCrLf & _
                "    on syf.SENDER = t1.STAFFID " & vbCrLf & _
                " where syf.CASEID = " & ProviderFactory.PositionPara & "CASEID " & vbCrLf & _
                "   and (fmp.TYPE is NULL or fmp.STEP_NO in ('START', 'END', 'TAKEBACK', 'CANCEL'))" & _
                " order by syf.STARTTIME " & vbCrLf

            ' and fmp.TYPE is NULL " & vbCrLf & _

            Dim paras(1) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("CASEID", sCaseId)

            Return Me.loadBySQL(sSql, paras) > 0
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

#Region "Zack Function"

    ''' <summary>
    ''' 根據狀態值查詢數據
    ''' </summary>
    ''' <param name="sStatus"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Zack 2012-06-04 Create
    ''' </remarks>
    Public Function loadBySatus(ByVal sStatus As String) As Integer
        Try
            Dim sSql As String = "SELECT DISTINCT(CASEID) FROM SY_FLOWSTEP  WHERE status!=" & ProviderFactory.PositionPara & "status"

            Dim paras(1) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("status", sStatus)

            Return Me.loadBySQL(sSql, paras) > 0
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
