Imports MBSC.UICtl

''' <summary>
''' 程式說明：代理人查詢
''' 建立者：Lake
''' 建立日期：2012-05-12
''' </summary>

Imports com.Azion.EloanUtility
Imports AUTH_OP.TABLE
Imports AUTH_OP

Public Class SY_QRYAGENT
    Inherits SYUIBase

    

#Region "PageLoad"
    ''' <summary>
    ''' 頁面加載
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>[Lake] 2012/05/12 Created</remarks>
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        ' 為分頁用戶控件指定關聯的GirdView和綁定數據的方法
        Me.PagerMenu1.SetTarget(Me.dgCurrentAgent, New BindDataDelegate(AddressOf initData))

        ' 初始化參數
        initParas()

        If Not IsPostBack Then

            ' 初始化數據
            initData()
        End If
    End Sub
#End Region

#Region "Function"

    ''' <summary>
    ''' 初始化參數
    ''' </summary>
    ''' <remarks>[Lake] 2012/05/12 Created</remarks>
    Sub initParas()

        
        ' 測試數據
        If m_bTesting OrElse Request("TESTMODE") = "1" Then
            m_bTesting = True

            Dim loginUser As New SY_LOGIN(GetDatabaseManager)
            loginUser.getUserInfo("904", "S000035")
            m_sHoFlag = "1"
            MyBase.InitParas()
        End If


        If m_sHoFlag = "1" OrElse SY_UPDAGENT.getSelfStatus() = "1" Then
            m_sHoFlag = "3"
            MyBase.InitParas()
        End If

    End Sub

    ''' <summary>
    ''' 初始化數據
    ''' </summary>
    ''' <remarks>[Lake] 2012/05/12 Created</remarks>
    Sub initData()
        Try
            Dim syUserRoleAgentList As New SY_USERROLEAGENTList(GetDatabaseManager())
            Dim syBranchList As New AUTH_OP.SY_BRANCHList(GetDatabaseManager())

            ' “目前代理狀況”選中
            rdoCurrent.Checked = True

            ' 被代理人單位
            If syBranchList.getBranchByParas("Y", m_sHoFlag, m_sWorkingBrid) Then
                ddlStaffBranch.DataSource = syBranchList.getCurrentDataSet()
                ddlStaffBranch.DataValueField = "BRA_DEPNO"
                ddlStaffBranch.DataTextField = "BRCNAME"
                ddlStaffBranch.DataBind()

                ddlStaffBranch.Items.Insert(0, New ListItem("請選擇", "0"))
            End If

            If m_sHoFlag <> "1" Then
                ddlStaffBranch.SelectedIndex = 1
            Else
                ddlStaffBranch.SelectedIndex = 0
            End If

            ' 綁定目前代理狀況
            If syUserRoleAgentList.getCurrentAgentInfo(ddlStaffBranch.SelectedValue, "C") Then
                dgCurrentAgent.DataSource = syUserRoleAgentList.getCurrentDataSet()
                dgCurrentAgent.DataBind()
            Else
                com.Azion.EloanUtility.FieldValidator.ShowEmptyGridView(dgCurrentAgent, syUserRoleAgentList.getCurrentDataSet().Tables(0))
            End If
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 臺灣日期取得西元日期
    ''' </summary>
    ''' <param name="twDate">臺灣日期</param>
    ''' <returns></returns>
    ''' <remarks>[Lake] 2012/06/04 Created</remarks>
    Private Function getADDate(ByVal sTWDate As String) As String
        Try
            Dim sYear As String = sTWDate.Split(".")(0)
            Dim iYear As Integer = 0
            Dim sADDate As String = String.Empty

            If sYear.IndexOf("-") > 0 Then
                iYear = 1911 - sYear.Split("-")(1)
            Else
                iYear = 1911 + sYear
            End If

            sADDate = iYear.ToString & sTWDate.Substring(sTWDate.IndexOf(".")).Replace(".", "-")

            Return sADDate
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Function

    ''' <summary>
    ''' 西元日期取得台灣日期
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>[Lake] 2012/06/13 Created</remarks>
    Private Function getTWData(ByVal sADDate As String) As String
        Try
            If sADDate <> "" Then
                Dim tYear As Integer = sADDate.Split(".")(0) - 1911

                If tYear < 0 Then
                    tYear = "0" & tYear.ToString()
                ElseIf tYear = 0 Then
                    tYear = "00"
                Else
                    tYear = tYear.ToString()
                End If

                Return tYear.ToString() & sADDate.Substring(sADDate.IndexOf("."))
            Else
                Return ""
            End If
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Function
#End Region

#Region "Event"
    ''' <summary>
    ''' 選擇一個被代理人單位
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>[Lake] 2012/06/26 Created</remarks>
    Protected Sub ddlStaffBranch_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlStaffBranch.SelectedIndexChanged
        Try
            Dim syUserList As New SY_USERList(GetDatabaseManager())

            ddlStaff.Items.Clear()

            ' 綁定被代理人列表
            Dim sStaffid As String = ""
            If SY_UPDAGENT.getSelfStatus() = "1" Then
                sStaffid = m_sLoginUserid
            End If

            If syUserList.getStaffList(ddlStaffBranch.SelectedValue, sStaffid) Then
                ddlStaff.DataSource = syUserList.getCurrentDataSet()
                ddlStaff.DataValueField = "BRADEPNO"
                ddlStaff.DataTextField = "USERINFO"
                ddlStaff.DataBind()
            End If

            If ddlStaff.Items.Count = 1 Then
                ddlStaff.SelectedIndex = 0
            End If


            ddlStaff.Items.Insert(0, New ListItem("請選擇", "0"))
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 選擇一個代理人單位
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>[Lake] 2012/06/26 Created</remarks>
    Protected Sub ddlAgentStaffBranch_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlAgentStaffBranch.SelectedIndexChanged
        Try
            Dim syUserList As New SY_USERList(GetDatabaseManager())

            ddlAgentStaff.Items.Clear()

            ' 綁定代理人列表
            If syUserList.getAgentStaffList(ddlAgentStaffBranch.SelectedValue) Then
                ddlAgentStaff.DataSource = syUserList.getCurrentDataSet()
                ddlAgentStaff.DataValueField = "AGENTBRADEPNO"
                ddlAgentStaff.DataTextField = "AGENTUSERINFO"
                ddlAgentStaff.DataBind()
            End If

            ddlAgentStaff.Items.Insert(0, New ListItem("請選擇", "0"))
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 選擇當前代理情況
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>[Lake] 2012/05/12 Created</remarks>
    Protected Sub rdoCurrent_CheckedChanged(sender As Object, e As EventArgs) Handles rdoCurrent.CheckedChanged
        Try
            Dim oCurrent As CheckBox = sender

            ' 隱藏被代理人，代理人單位，代理人，代理期間
            trStaff.Visible = False
            trAgentStaffBranch.Visible = False
            trAgentStaff.Visible = False
            trAgentDateTime.Visible = False

            ' 刷新頁面
            initData()
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 選擇歷史代理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>[Lake] 2012/05/12 Created</remarks>
    Protected Sub rdoHistory_CheckedChanged(sender As Object, e As EventArgs) Handles rdoHistory.CheckedChanged
        Try
            Dim syBranchList As New AUTH_OP.SY_BRANCHList(GetDatabaseManager())
            Dim syUserList As New SY_USERList(GetDatabaseManager())
            Dim tNowDate As Date = Date.Now

            ddlStaffBranch.SelectedIndex = 0

            ' 清空代理人和被代理人下拉選單
            ddlStaff.Items.Clear()
            ddlAgentStaff.Items.Clear()

            txtStartDate.Text = (tNowDate.Year - 1911).ToString() & "." & tNowDate.Month.ToString() & "." & tNowDate.Day.ToString()
            txtStartTime.Text = tNowDate.ToString("HH:mm")
            txtEndDate.Text = (tNowDate.Year - 1911).ToString() & "." & tNowDate.Month.ToString() & "." & tNowDate.Day.ToString()
            txtEndTime.Text = "23:59"

            ' 顯示被代理人，代理人單位，代理人，代理期間
            trStaff.Visible = True
            trAgentStaffBranch.Visible = True
            trAgentStaff.Visible = True
            trAgentDateTime.Visible = True

            If m_sHoFlag <> "1" Then
                ddlStaffBranch.SelectedIndex = 1

                ' 綁定被代理人列表
                Dim sStaffid As String = ""
                If SY_UPDAGENT.getSelfStatus() = "1" Then
                    sStaffid = m_sLoginUserid
                End If

                If syUserList.getStaffList(ddlStaffBranch.SelectedValue, sStaffid) Then
                    ddlStaff.DataSource = syUserList.getCurrentDataSet()
                    ddlStaff.DataValueField = "BRADEPNO"
                    ddlStaff.DataTextField = "USERINFO"
                    ddlStaff.DataBind()
                End If

                If ddlStaff.Items.Count = 1 Then
                    ddlStaff.SelectedIndex = 0
                End If
            End If

            ' 代理人部門getBranchByParas
            If syBranchList.getBranchByParas("N", "", "") Then
                ddlAgentStaffBranch.DataSource = syBranchList.getCurrentDataSet()
                ddlAgentStaffBranch.DataValueField = "BRA_DEPNO"
                ddlAgentStaffBranch.DataTextField = "BRCNAME"
                ddlAgentStaffBranch.DataBind()
            End If

            ddlAgentStaffBranch.Items.Insert(0, New ListItem("請選擇", "0"))

            ' 代理人和被代理人下拉選單，添加“請選擇”顯示
            ddlStaff.Items.Insert(0, New ListItem("請選擇", "0"))
            ddlAgentStaff.Items.Insert(0, New ListItem("請選擇", "0"))
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 根據查詢條件查詢代理信息
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>[Lake] 2012/05/12 Created</remarks>
    Protected Sub btnConfirm_Click(sender As Object, e As EventArgs) Handles btnConfirm.Click
        Try
            Dim syUserRoleAgentList As New SY_USERROLEAGENTList(GetDatabaseManager())
            Dim dtUserRoleAgent As DataTable
            Dim sStaffId As String = Nothing
            Dim sAgentStaffId As String = Nothing
            Dim sStartDateTime As String = Nothing
            Dim sEndDateTime As String = Nothing
            Dim sSearchFlag As String = String.Empty

            ' 選擇歷史代理資料
            If rdoHistory.Checked Then

                sSearchFlag = "H"

                ' 被代理人編號
                If ddlStaff.SelectedIndex > 0 Then
                    sStaffId = ddlStaff.SelectedItem.Text.Substring(0, 7)
                End If

                ' 代理人編號
                If ddlAgentStaff.SelectedIndex > 0 Then
                    sAgentStaffId = ddlAgentStaff.SelectedItem.Text.Substring(0, 7)
                End If

                ' 代理開始日期時間
                If txtStartDate.Text.Trim() <> "" AndAlso txtStartTime.Text.Trim() <> "" Then
                    sStartDateTime = getADDate(txtStartDate.Text.Trim()) & " " & txtStartTime.Text.Trim()
                End If

                ' 代理結束日期時間
                If txtEndDate.Text.Trim() <> "" AndAlso txtEndTime.Text.Trim() <> "" Then
                    sEndDateTime = getADDate(txtEndDate.Text.Trim()) & " " & txtEndTime.Text.Trim()
                End If
            Else
                sSearchFlag = "C"
            End If

            ' 根據條件查詢代理資料，沒有資料時，顯示空表頭
            If syUserRoleAgentList.getCurrentAgentInfo(ddlStaffBranch.SelectedValue, sSearchFlag, ddlAgentStaffBranch.SelectedValue, sStaffId, sAgentStaffId, sStartDateTime, sEndDateTime) Then
                dtUserRoleAgent = syUserRoleAgentList.getCurrentDataSet().Tables(0)

                dgCurrentAgent.DataSource = dtUserRoleAgent
                dgCurrentAgent.DataBind()
            Else
                com.Azion.EloanUtility.FieldValidator.ShowEmptyGridView(dgCurrentAgent, syUserRoleAgentList.getCurrentDataSet().Tables(0))
            End If
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 代理資料綁定
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>[Lake] 2012/06/04 Created</remarks>
    Protected Sub dgCurrentAgent_ItemDataBound(sender As Object, e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles dgCurrentAgent.ItemDataBound
        Try
            If e.Item.ItemType = ListItemType.AlternatingItem OrElse e.Item.ItemType = ListItemType.Item Then
                Dim oStartDateTime As Label = e.Item.FindControl("lblStartDateTime")
                Dim oEndDateTime As Label = e.Item.FindControl("lblEndDateTime")
                Dim oCreateDateTime As Label = e.Item.FindControl("lblCreateDateTime")
                Dim oCancelDateTime As Label = e.Item.FindControl("lblCancelDateTime")


                oStartDateTime.Text = getTWData(oStartDateTime.Text)
                oEndDateTime.Text = getTWData(oEndDateTime.Text)
                oCreateDateTime.Text = getTWData(oCreateDateTime.Text)
                oCancelDateTime.Text = getTWData(oCancelDateTime.Text)
            End If
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub
#End Region
End Class