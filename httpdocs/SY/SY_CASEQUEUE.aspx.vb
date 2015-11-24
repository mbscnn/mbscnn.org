''' <summary>
''' 程式說明：佇列文件
''' 建立者：Lake
''' 建立日期：2012-05-14
''' </summary>

Imports com.Azion.UITools
Imports AUTH_OP
Imports FLOW_OP

Public Class SY_CASEQUEUE
    Inherits SYUIBase

#Region "PageLoad"
    ''' <summary>
    ''' 頁面加載
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>[Lake] 2012/05/14 Created</remarks>
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        ' 初始化參數
        initParas()

        If Not IsPostBack Then

            ' 綁定案件資料
            initData()
        End If
    End Sub
#End Region

#Region "Function"
    ''' <summary>
    ''' 初始化參數
    ''' </summary>
    ''' <remarks>[Lake] 2012/05/14 Created</remarks>
    Sub initParas()

        ' 測試數據
        If m_bTesting Then
            m_sWorkingUserid = "S000023"
            m_sWorkingTopDepNo = "75"
        End If
    End Sub

    ''' <summary>
    ''' 綁定案件資料
    ''' </summary>
    ''' <remarks>[Lake] 2012/05/14 Created</remarks>
    Sub initData()
        Try
            Dim flowFacade As New FlowFacade(GetDatabaseManager())

            dgCase.DataSource = flowFacade.GetCaseListFromQueue(m_sWorkingUserid, m_sWorkingBrid)
            dgCase.DataBind()
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub
#End Region

#Region "Event"
    ''' <summary>
    ''' 取得案件
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>[Lake] 2012/05/14 Created</remarks>
    Protected Sub btnGetCase_Click(sender As Object, e As EventArgs) Handles btnGetCase.Click

        Dim returnValue As StepInfoItemExt

        Try
            'Dim syFlowStep As New AUTH_OP.SY_FLOWSTEP(GetDatabaseManager())
            Dim syFlowFacade As New FLOW_OP.FlowFacade(GetDatabaseManager)
            Dim listCaseid As New List(Of String)


            For Each dgi As DataGridItem In dgCase.Items
                Dim chkFlag As CheckBox = dgi.FindControl("chkFlag")
                Dim sCaseId As Label = dgi.FindControl("lblCaseId")

                If chkFlag.Checked Then
                    listCaseid.Add(sCaseId.Text)
                End If
            Next

            If listCaseid.Count = 0 Then
                com.Azion.EloanUtility.UIUtility.alert("請至少勾選一筆案件資料！")
                Return
            End If

            GetDatabaseManager.beginTran()

            For Each caseid As String In listCaseid
                returnValue = syFlowFacade.SetOwnerOfCaseFromQueue(caseid, m_sWorkingUserid, m_sWorkingBrid)
            Next

            GetDatabaseManager.commit()

            com.Azion.EloanUtility.UIUtility.closeWindow()
        Catch ex As Exception
            GetDatabaseManager.Rollback()
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 關閉視窗
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>[Lake] 2012/05/14 Created</remarks>
    Protected Sub btnCloseWin_Click(sender As Object, e As EventArgs) Handles btnCloseWin.Click
        Try
            com.Azion.EloanUtility.UIUtility.closeWindow()
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub
#End Region
End Class