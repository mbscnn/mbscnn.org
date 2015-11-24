Imports MBSC.UICtl

''' <summary>
''' 程式說明：代理人設定
''' 建立者：Lake
''' 建立日期：2012-05-10
''' </summary>

Imports AUTH_OP.TABLE
Imports com.Azion.EloanUtility
Imports AUTH_OP

Public Class SY_UPDAGENT
    Inherits SYUIBase
 
#Region "PageLoad"
    ''' <summary>
    ''' 頁面加載
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>[Lake] 2012/05/10 Created</remarks>
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        ' 為分頁用戶控件指定關聯的GirdView和綁定數據的方法
        Me.PagerMenu1.SetTarget(Me.dgAgentInfo, New BindDataDelegate(AddressOf initData))

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
    ''' <remarks>[Lake] 2012/05/10 Created</remarks>
    Sub initParas()
 
        ' 測試數據
        'If m_bTesting Then
        '    m_sCaseId = "049251000000818"
        '    m_sWorkingBrid = "904"
        '    m_sWorkingUserid = "S000023"
        'End If

        If m_bTesting OrElse Request("TESTMODE") = "1" Then
            m_bTesting = True

            Dim loginUser As New SY_LOGIN(GetDatabaseManager)
            loginUser.getUserInfo("904", "S000035")
            m_sHoFlag = "1"
            'm_bCheck = False

            'm_sWorkingTopDepNo = "75"
            'm_bDisplayMode = False
        End If


        If getSelfStatus() = "1" Then
            m_sHoFlag = "3"
        End If


    End Sub

    ''' <summary>
    ''' 初始化數據
    ''' </summary>
    ''' <remarks>[Lake] 2012/05/10 Created</remarks>
    Sub initData()
        Try
            Dim syUserRoleAgentList As New SY_USERROLEAGENTList(GetDatabaseManager())

            ' 隱藏代理人設定新增區塊，顯示代理人設定清單區塊
            trAgentSet.Visible = False
            trAgentSetDetail.Visible = True

            ' True，綁定代理清單資料
            If syUserRoleAgentList.getAgentInfo(m_sWorkingBrid, "0", m_sHoFlag) Then
                dgAgentInfo.DataSource = syUserRoleAgentList.getCurrentDataSet()
                dgAgentInfo.DataBind()
            Else
                ' 暫時按照專案中統一方式顯示空表頭
                com.Azion.EloanUtility.FieldValidator.ShowEmptyGridView(dgAgentInfo, syUserRoleAgentList.getCurrentDataSet().Tables(0))
            End If
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 設置新增模塊為初始狀態
    ''' </summary>
    ''' <remarks>[Lake] 2012/05/11 Created</remarks>
    Sub setDefaultStatus()

        ' 新增區塊欄位初始化
        ddlStaff.SelectedIndex = 0
        ddlAgentStaff.SelectedIndex = 0
        txtStartDate.Text = ""
        txtStartTime.Text = ""
        txtEndDate.Text = ""
        txtEndTime.Text = ""

        ' 顯示代理人資料明細，隱藏新增區塊
        trAgentSetDetail.Visible = True
        trAgentSet.Visible = False
    End Sub

    ''' <summary>
    ''' 臺灣日期取得西元日期
    ''' </summary>
    ''' <param name="twDate">臺灣日期</param>
    ''' <returns></returns>
    ''' <remarks>[Lake] 2012/06/04 Created</remarks>
    Private Function getADDate(ByVal twDate As String) As String
        Try
            Dim sYear As String = twDate.Split(".")(0)
            Dim iYear As Integer = 0
            Dim sADDate As String = String.Empty

            If sYear.IndexOf("-") > 0 Then
                iYear = 1911 - sYear.Split("-")(1)
            Else
                iYear = 1911 + sYear
            End If

            sADDate = iYear.ToString & twDate.Substring(twDate.IndexOf(".")).Replace(".", "-")

            Return sADDate
        Catch ex As Exception
            Throw
        End Try
    End Function

    Protected Function StringToDate(ByVal dtDate As String) As Date

        Dim sDate() As String = dtDate.Split("-")

        If sDate.Length <> 3 Then
            sDate = dtDate.Split(".")
        End If

        If sDate.Length <> 3 Then
            Throw New Exception("日期格式不對")
        End If

        Try
            Return New Date(CInt(sDate(0)), CInt(sDate(1)), CInt(sDate(2)))
        Catch ex As Exception
            Throw New Exception("日期格式不對")
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
            Throw
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
            If syUserList.getStaffList(ddlStaffBranch.SelectedValue) Then
                ddlStaff.DataSource = syUserList.getCurrentDataSet()
                ddlStaff.DataValueField = "BRADEPNO"
                ddlStaff.DataTextField = "USERINFO"
                ddlStaff.DataBind()
            End If

            ddlStaff.Items.Insert(0, New ListItem("請選擇", "0"))


            Dim bResult As Boolean
            If m_sHoFlag = 1 AndAlso ddlStaffBranch.SelectedValue <> ddlAgentStaffBranch.SelectedValue Then
                Dim syBranchList As New AUTH_OP.SY_BRANCHList(GetDatabaseManager())

                '不跨區代理
                bResult = syBranchList.getBranchByParas("Y", m_sHoFlag, FLOW_OP.TABLE.SY_BRANCH.getNewInstance(GetDatabaseManager).GetBRID(ddlStaffBranch.SelectedValue))

                ddlAgentStaffBranch.DataSource = syBranchList.getCurrentDataSet()
                ddlAgentStaffBranch.DataValueField = "BRA_DEPNO"
                ddlAgentStaffBranch.DataTextField = "BRCNAME"
                ddlAgentStaffBranch.DataBind()

                ddlAgentStaffBranch.SelectedIndex = 0
                ddlAgentStaffBranch.Items.Insert(0, New ListItem("請選擇", "0"))
                ddlAgentStaffBranch_SelectedIndexChanged()
            End If


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
        ddlAgentStaffBranch_SelectedIndexChanged()
    End Sub


    Protected Sub ddlAgentStaffBranch_SelectedIndexChanged()
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
    ''' "取消"代理人
    ''' </summary>
    ''' <param name="source"></param>
    ''' <param name="e"></param>
    ''' <remarks>[Lake] 2012/05/10 Created</remarks>
    Protected Sub dgFlowCaseStatus_ItemCommand(source As Object, e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles dgAgentInfo.ItemCommand
        Try
            If e.Item.ItemType = ListItemType.Item OrElse e.Item.ItemType = ListItemType.AlternatingItem Then
                If e.CommandName = "Cancel" Then
                    Dim syUserRoleAgent As New SY_USERROLEAGENT(GetDatabaseManager())
                    Dim oStaff As Label = e.Item.FindControl("lblStaffId")
                    Dim oAgentStaff As Label = e.Item.FindControl("lblAgentStaffId")
                    Dim oBraDepNo As HiddenField = e.Item.FindControl("hidBraDepNo")
                    Dim oAgentBraDepNo As HiddenField = e.Item.FindControl("hidAgentBraDepNo")
                    Dim oCreateTime As HiddenField = e.Item.FindControl("hidCreateTime")

                    If syUserRoleAgent.loadByPK(oStaff.Text, oBraDepNo.Value, oAgentStaff.Text, oAgentBraDepNo.Value, oCreateTime.Value) Then
                        syUserRoleAgent.setAttribute("CANCELTIME", Date.Now())

                        syUserRoleAgent.save()
                    End If
                End If
            End If

            ' 刷新數據
            initData()
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 代理人信息綁定
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>[lake] 2012/06/04 Created</remarks>
    Protected Sub dgAgentInfo_ItemDataBound(sender As Object, e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles dgAgentInfo.ItemDataBound
        Try
            If e.Item.ItemType = ListItemType.AlternatingItem OrElse e.Item.ItemType = ListItemType.Item Then
                Dim oStartDateTime As Label = e.Item.FindControl("lblStartDateTime")
                Dim oEndDateTime As Label = e.Item.FindControl("lblEndDateTime")

                oStartDateTime.Text = getTWData(oStartDateTime.Text)
                oEndDateTime.Text = getTWData(oEndDateTime.Text)
            End If
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 新增
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>[Lake] 2012/05/10 Created</remarks>
    Protected Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        Try
            Dim tNowDate As Date = Date.Now
            Dim syBranchList As New AUTH_OP.SY_BRANCHList(GetDatabaseManager())
            Dim syUserList As New SY_USERList(GetDatabaseManager())

            ' 清空代理人及被代理人下拉選單
            ddlStaff.Items.Clear()
            ddlAgentStaff.Items.Clear()

            ' 被代理人部門
            If syBranchList.getBranchByParas("Y", m_sHoFlag, m_sWorkingBrid) Then
                ddlStaffBranch.DataSource = syBranchList.getCurrentDataSet()
                ddlStaffBranch.DataValueField = "BRA_DEPNO"
                ddlStaffBranch.DataTextField = "BRCNAME"
                ddlStaffBranch.DataBind()
            End If

            ddlStaffBranch.Items.Insert(0, New ListItem("請選擇", "0"))

            ' 代理人部門
            Dim bResult As Boolean
            If m_sHoFlag = 1 Then
                bResult = syBranchList.getBranchByParas("Y", "10", "ABCDEFG")
            Else
                bResult = syBranchList.getBranchByParas("Y", m_sHoFlag, m_sWorkingBrid)
            End If

            If bResult Then
                ddlAgentStaffBranch.DataSource = syBranchList.getCurrentDataSet()
                ddlAgentStaffBranch.DataValueField = "BRA_DEPNO"
                ddlAgentStaffBranch.DataTextField = "BRCNAME"
                ddlAgentStaffBranch.DataBind()
            End If

            If m_sHoFlag <> "1" OrElse getSelfStatus() = "1" Then

                ' 默認選擇單位
                ddlStaffBranch.SelectedIndex = 1

                Dim sStaffid As String = ""
                If getSelfStatus() = "1" Then
                    sStaffid = m_sLoginUserid
                End If

                ' 綁定被代理人列表
                If syUserList.getStaffList(ddlStaffBranch.SelectedValue, sStaffid) Then
                    ddlStaff.DataSource = syUserList.getCurrentDataSet()
                    ddlStaff.DataValueField = "BRADEPNO"
                    ddlStaff.DataTextField = "USERINFO"
                    ddlStaff.DataBind()
                End If
            End If

            ' 代理人及被代理人下拉選單，添加“請選擇”顯示
            ddlStaff.Items.Insert(0, New ListItem("請選擇", "0"))
            If ddlStaff.Items.Count = 2 Then
                ddlStaff.SelectedIndex = 1
            End If

            ddlAgentStaffBranch.Items.Insert(0, New ListItem("請選擇", "0"))
            If ddlAgentStaffBranch.Items.Count = 2 Then
                ddlAgentStaffBranch.SelectedIndex = 1
                ddlAgentStaffBranch_SelectedIndexChanged()
            End If

            ddlAgentStaff.Items.Insert(0, New ListItem("請選擇", "0"))

            Dim tsFiveMinutes As New TimeSpan(0, 5, 0)
            tNowDate = tNowDate + tsFiveMinutes

            txtStartDate.Text = (tNowDate.Year - 1911).ToString() & "." & IIf(tNowDate.Month < 10, "0" & tNowDate.Month.ToString(), tNowDate.Month.ToString()) & "." & tNowDate.Day.ToString()
            txtStartTime.Text = tNowDate.ToString("HH:mm")
            txtEndDate.Text = (tNowDate.Year - 1911).ToString() & "." & IIf(tNowDate.Month < 10, "0" & tNowDate.Month.ToString(), tNowDate.Month.ToString()) & "." & tNowDate.Day.ToString()
            txtEndTime.Text = "23:59"

            ' 顯示代理人設定新增區塊，隱藏代理人設定清單區塊
            trAgentSet.Visible = True
            trAgentSetDetail.Visible = False
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 儲存
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>[Lake] 2012/05/10 Created</remarks>
    Protected Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Try
            Dim syUserRoleAgent As New SY_USERROLEAGENT(GetDatabaseManager())
            Dim sErrorMsg As String = String.Empty
            Dim oErrorControl As Object = Nothing
            Dim sStartDateTime As String = String.Empty
            Dim sEndDateTime As String = String.Empty
            Dim sSysDateTime As String = String.Empty




            ' 訖日不可小於起日
            If StringToDate(getADDate(txtStartDate.Text.Trim())) > StringToDate(getADDate(txtEndDate.Text.Trim())) Then
                com.Azion.EloanUtility.UIUtility.alert("訖日期不可小於起日期！")
                Return
            End If

            sStartDateTime = getADDate(txtStartDate.Text.Trim()) + " " + txtStartTime.Text.Trim()
            sEndDateTime = getADDate(txtEndDate.Text.Trim()) + " " + txtEndTime.Text.Trim()
            sSysDateTime = Date.Now.ToString("yyyy-MM-dd HH:mm")

            ' 起訖日期不可小於系統日期
            If sStartDateTime < sSysDateTime OrElse sEndDateTime < sSysDateTime Then
                com.Azion.EloanUtility.UIUtility.alert("起訖日期不可小於系統日期！")

                Return
            End If

            ' 選擇有意義的代理人及被代理人
            If ddlStaff.SelectedIndex > 0 AndAlso ddlAgentStaff.SelectedIndex > 0 Then

                ' 判斷代理時間沒有交集
                If syUserRoleAgent.checkAgentRelation(ddlStaff.SelectedItem.Text.Substring(0, 7), _
                                                          ddlAgentStaff.SelectedItem.Text.Substring(0, 7), _
                                                          getADDate(txtStartDate.Text.Trim()) & " " & txtStartTime.Text.Trim(), _
                                                          getADDate(txtEndDate.Text.Trim()) & " " & txtEndTime.Text.Trim(), _
                                                          ddlStaffBranch.SelectedValue, ddlAgentStaffBranch.SelectedValue) Then

                    If sErrorMsg = "" Then
                        oErrorControl = txtStartDate
                    End If

                    sErrorMsg = sErrorMsg & "設定代理期間不可有交集！"
                End If
            End If

            ' 提示錯誤信息
            If sErrorMsg <> "" Then
                com.Azion.EloanUtility.UIUtility.alert(sErrorMsg)

                CType(oErrorControl, TextBox).Text = ""
                CType(oErrorControl, TextBox).Focus()

                Return
            End If

            ' 代理關係已經存在，更新代理信息；不存在，新增代理關係
            If Not syUserRoleAgent.loadByPK(ddlStaff.SelectedItem.Text.Substring(0, 7), _
                                        ddlStaff.SelectedValue.Split(";")(1), _
                                        ddlAgentStaff.SelectedItem.Text.Substring(0, 7), _
                                        ddlAgentStaff.SelectedValue.Split(";")(1), Date.Now) Then

                syUserRoleAgent.setAttribute("STAFFID", ddlStaff.SelectedItem.Text.Substring(0, 7))
                syUserRoleAgent.setAttribute("BRA_DEPNO", ddlStaff.SelectedValue.Split(";")(1))
                syUserRoleAgent.setAttribute("AGENT_STAFFID", ddlAgentStaff.SelectedItem.Text.Substring(0, 7))
                syUserRoleAgent.setAttribute("AGENT_BRADEPNO", ddlAgentStaff.SelectedValue.Split(";")(1))
                syUserRoleAgent.setAttribute("STARTTIME", getADDate(txtStartDate.Text.Trim()) & " " & txtStartTime.Text.Trim())
                syUserRoleAgent.setAttribute("ENDTIME", getADDate(txtEndDate.Text.Trim()) & " " & txtEndTime.Text.Trim())
                syUserRoleAgent.setAttribute("CREATEUSER", m_sWorkingUserid)
                syUserRoleAgent.setAttribute("CREATETIME", Date.Now)

                syUserRoleAgent.save()
            End If

            com.Azion.EloanUtility.UIUtility.alert("儲存成功！")

            ' 設置新增區塊為初始狀態
            setDefaultStatus()

            ' 刷新代理人清檔
            initData()
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 取消
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>[Lake] 2012/05/10 Created</remarks>
    Protected Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Try
            ' 設置新增區塊為初始狀態
            setDefaultStatus()
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub
#End Region


    Public Shared Function getSelfStatus() As String
        Dim currContext As System.Web.HttpContext = System.Web.HttpContext.Current
        Dim sSelfStatus As String = ""

        If IsNothing(currContext) OrElse IsNothing(currContext.Session) Then
            Return sSelfStatus
        End If

        If ValidateUtility.isValidateData(currContext.Request("SELF")) Then
            sSelfStatus = currContext.Request("SELF")
        ElseIf ValidateUtility.isValidateData(currContext.Request("SELF")) Then
            sSelfStatus = CStr(currContext.Request("SELF"))
        End If

        Return UIUtility.getReqString(sSelfStatus)
    End Function

End Class