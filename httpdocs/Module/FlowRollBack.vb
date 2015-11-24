Public Class FlowRollBack
    Private m_databaseManager As com.Azion.NET.VB.DatabaseManager
    Private m_sCaseId As String
  
    Public Sub New(ByVal databaseManager As com.Azion.NET.VB.DatabaseManager, ByVal sCaseId As String)
        m_databaseManager = databaseManager
        m_sCaseId = sCaseId
    End Sub

    Function rollBack_06410300() As String
        Dim sPreUser As String = String.Empty
        Try
            Dim gaCASEAPV As New MBSC.MB_OP.GA_CASEAPV(m_databaseManager)
            If gaCASEAPV.LoadByPK(Me.m_sCaseId) Then
                sPreUser = gaCASEAPV.getString("CPL_CHG_UID")
            End If
        Catch ex As Exception
            Throw
        End Try
        Return sPreUser
    End Function

    Function rollBack_06415500() As String
        Dim sPreUser As String = String.Empty
        Try
            Dim gaCASEAPV As New MBSC.MB_OP.GA_CASEAPV(m_databaseManager)
            If gaCASEAPV.LoadByPK(Me.m_sCaseId) Then
                sPreUser = gaCASEAPV.getString("CPL_UID_5500")
            End If
        Catch ex As Exception
            Throw
        End Try
        Return sPreUser
    End Function

    'Function rollBack_04101100() As String
    '    Dim sPreUser As String = String.Empty
    '    Try
    '        '先關閉父流程及所有的子流程
    '        Dim flowFacade As FLOW_OP.FlowFacade = FLOW_OP.FlowFacade.getNewInstance(m_databaseManager)
    '        Dim FlowStepInfoItem As FLOW_OP.StepInfoItemExt
    '        FlowStepInfoItem = flowFacade.getSYFlowStep.GetStepInfoItem(m_sCaseId, 0, "04101200", Nothing, True)
    '        flowFacade.CloseAllSubFlowIncident(m_sCaseId, FlowStepInfoItem.subflowSeq, True)

    '        Dim enCaseBrdivRoles As New Eloan.EN_OP.EN_CASEBRDIVROLES(m_databaseManager)
    '        If enCaseBrdivRoles.loadEN_CASEBRDIVROLESByCaseId(m_sCaseId) Then
    '            sPreUser = enCaseBrdivRoles.getAttribute("OPUID_0900")
    '        End If
    '    Catch ex As Exception
    '        Throw
    '    End Try
    '    Return sPreUser
    'End Function
End Class
