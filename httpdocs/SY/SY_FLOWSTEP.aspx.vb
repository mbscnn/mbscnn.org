Imports com.Azion.EloanUtility

''' <summary>
''' 程式說明：流程步驟
''' 建立者：Lake
''' 建立日期：2012-04-09
''' </summary>

Public Class SY_FLOWSTEP
    Inherits SYUIBase

#Region "PageLoad"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        ' 初始化參數
        initParas()

        If Not IsPostBack Then

            ' 初始化頁面數據
            initData()
        End If
    End Sub
#End Region

#Region "Function"
    ''' <summary>
    ''' 初始化參數
    ''' </summary>
    ''' <remarks>[Lake] 2012/05/09 Created</remarks>
    Sub initParas()

        If m_bTesting Then
            m_sCaseId = "SY0111019000066"
        End If
    End Sub

    ''' <summary>
    ''' 初始化頁面數據
    ''' </summary>
    ''' <remarks>[Lake] 2012/05/09 Created</remarks>
    Sub initData()
        Try
            Dim syFlowStepList As New AUTH_OP.SY_FLOWSTEPList(GetDatabaseManager())

            ' 查詢到案件資料,流程案件綁定資料
            If syFlowStepList.getFlowCaseStatus(m_sCaseId) Then
                dgFlowCaseStatus.DataSource = syFlowStepList.getCurrentDataSet()
                dgFlowCaseStatus.DataBind()
            Else
                com.Azion.EloanUtility.FieldValidator.ShowEmptyGridView(dgFlowCaseStatus, syFlowStepList.getCurrentDataSet().Tables(0))
            End If
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

#End Region

End Class