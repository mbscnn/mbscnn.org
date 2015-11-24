Imports MBSC.UICtl

''' <summary>
''' 程式說明：人員分派查詢
''' 建立者：Lake
''' 建立日期：2012-05-15
''' </summary>

Imports com.Azion.EloanUtility
Imports AUTH_OP.TABLE
Imports AUTH_OP
Imports System.IO

Public Class SY_QRYAUTH
    Inherits SYUIBase

    ' 屬性標識
    Dim m_sHoFlag As String = String.Empty

#Region "PageLoad"
    ''' <summary>
    ''' 頁面加載
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>[Lake] 2012/05/15 Created</remarks>
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        ' 顯示分頁控件
        trPagerMenu.Visible = True

        ' 根據選擇狀態，設置分頁控件
        If rdoUser.Checked Then

            ' 為分頁用戶控件指定關聯的GirdView和綁定數據的方法
            Me.PagerMenu1.SetTarget(Me.dgUser, New BindDataDelegate(AddressOf initData))
        ElseIf rdoRole.Checked Then

            ' 為分頁用戶控件指定關聯的GirdView和綁定數據的方法
            Me.PagerMenu1.SetTarget(Me.dgRole, New BindDataDelegate(AddressOf initData))
        ElseIf rdoChange.Checked Then

            ' 為分頁用戶控件指定關聯的GirdView和綁定數據的方法
            Me.PagerMenu1.SetTarget(Me.dgChange, New BindDataDelegate(AddressOf initData))
        End If

        ' 初始化參數
        initParas()

        If Not IsPostBack Then

            ' 綁定期間列表
            getEndTimeList()

            ' 初始化單位列表
            initBranch()

            ' 初始化數據
            initData()
        End If
    End Sub
#End Region

#Region "Function"
    ''' <summary>
    ''' 初始化參數
    ''' </summary>
    ''' <remarks>[Lake] 2012/05/15 Created</remarks>
    Sub initParas()

        ' 取得屬性標識
        m_sHoFlag = Request.QueryString("hoflag")

        ' 測試數據
        If m_bTesting OrElse Request("TESTMODE") = "1" Then
            m_bTesting = True

            Dim loginUser As New SY_LOGIN(GetDatabaseManager)
            loginUser.getUserInfo("904", "S000035")
            m_bDisplayMode = False
        End If

        If m_bDisplayMode Then
            com.Azion.EloanUtility.UIUtility.setControlRead(Me)
        End If

    End Sub

    ''' <summary>
    ''' 初始化數據
    ''' </summary>
    ''' <remarks>[Lake] 2012/05/15 Created</remarks>
    Sub initData()
        Try
            If rdoUser.Checked Then

                ' 綁定人員角色清單
                bindStaffRole()
            ElseIf rdoRole.Checked Then

                ' 綁定角色人員清單
                bindRoleStaff()
            ElseIf rdoChange.Checked Then

                ' 綁定異動人員清單
                bindChangeStaff()
            End If
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 初始化單位列表
    ''' </summary>
    ''' <remarks>[Lake] 2012/06/28 Created</remarks>
    Sub initBranch()
        Try
            Dim syRelRoleUser As New AUTH_OP.SY_REL_ROLE_USER(GetDatabaseManager())
            Dim syBranchList As New AUTH_OP.SY_BRANCHList(GetDatabaseManager())
            Dim syBranch As New AUTH_OP.SY_BRANCH(GetDatabaseManager())

            Dim dt As DataTable = syBranchList.loadBranchInfo(m_sWorkingBrid, m_sHoFlag, "0", 5)

            ' 若有單位資料
            If Not dt Is Nothing Then
                ddlBranch.DataSource = syBranchList.getCurrentDataSet
                ddlBranch.DataTextField = "NAME"
                ddlBranch.DataValueField = "BRA_DEPNO"
                ddlBranch.DataBind()
            End If

            ddlBranch.Items.Insert(0, New ListItem("請選擇", "0"))

            ' 若只有一個單位，默認被選中
            If ddlBranch.Items.Count = 2 Then
                ddlBranch.SelectedIndex = 1
            End If
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 【異動查詢】綁定期間列表
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>[Lake] 2012/06/05 Created</remarks>
    Private Function getEndTimeList() As DataTable
        Try
            Dim sBraDepNo As String = String.Empty
            Dim aMonths As New ArrayList()

            aMonths.Add("01")
            aMonths.Add("02")
            aMonths.Add("03")
            aMonths.Add("04")
            aMonths.Add("05")
            aMonths.Add("06")
            aMonths.Add("07")
            aMonths.Add("08")
            aMonths.Add("09")
            aMonths.Add("10")
            aMonths.Add("11")
            aMonths.Add("12")

            ddlStartMonth.DataSource = aMonths
            ddlStartMonth.DataBind()

            ddlEndMonth.DataSource = aMonths
            ddlEndMonth.DataBind()

            ddlStartMonth.Items.Insert(0, New ListItem("請選擇", "0"))
            ddlEndMonth.Items.Insert(0, New ListItem("請選擇", "0"))


            Dim aDays(30) As String
            For n = 1 To 31
                aDays(n - 1) = n.ToString("D2")
            Next
            ddlStartDay.DataSource = aDays
            ddlStartDay.DataBind()
            ddlEndDay.DataSource = aDays
            ddlEndDay.DataBind()

            ddlStartDay.Items.Insert(0, New ListItem("請選擇", "0"))
            ddlEndDay.Items.Insert(0, New ListItem("請選擇", "0"))


            Dim aYears As New List(Of String)

            For nYear As Integer = 2012 To Year(Now())
                aYears.Add(nYear.ToString())
            Next

            ddlStartYear.DataSource = aYears
            ddlStartYear.DataBind()

            ddlEndYear.DataSource = aYears
            ddlEndYear.DataBind()

            ddlStartYear.Items.Insert(0, New ListItem("請選擇", "0"))
            ddlEndYear.Items.Insert(0, New ListItem("請選擇", "0"))

        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Function

    ''' <summary>
    ''' 創建人員角色表結構
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>[Lake] 2012/06/06 Created</remarks>
    Public Function createNewTableSchema() As DataTable
        Try
            Dim dtStaffRole As New DataTable()
            Dim dcNew As DataColumn = New DataColumn("STAFF", System.Type.GetType("System.String"))
            dtStaffRole.Columns.Add(dcNew)

            dcNew = New DataColumn("ROLENAME", System.Type.GetType("System.String"))
            dtStaffRole.Columns.Add(dcNew)

            Return dtStaffRole
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Function

    ''' <summary>
    ''' 綁定人員角色清單
    ''' </summary>
    ''' <remarks>[Lake] 2012/06/06 Created</remarks>
    Sub bindStaffRole()
        Try
            Dim syRelRoleUserList As New AUTH_OP.SY_REL_ROLE_USERList(GetDatabaseManager())
            Dim dtStaffRole As DataTable = createNewTableSchema()
            Dim sStaff As String = String.Empty
            Dim sRole As String = String.Empty

            trUser.Visible = True
            trTime.Visible = False
            trChange.Visible = False
            trRole.Visible = False

            ' 查詢人員角色列表
            If syRelRoleUserList.getStaffRoleList(ddlBranch.SelectedValue) Then
                For Each drStaffRole As DataRow In syRelRoleUserList.getCurrentDataSet().Tables(0).Rows
                    If sStaff <> drStaffRole("STAFF").ToString() Then
                        Dim drCollection As DataRow() = syRelRoleUserList.getCurrentDataSet().Tables(0).Select("STAFF='" & drStaffRole("STAFF").ToString() & "'")
                        Dim drNew As DataRow = dtStaffRole.NewRow()

                        For Each drRole In drCollection
                            sRole = sRole & "、" & drRole("ROLENAME").ToString()
                        Next

                        drNew("STAFF") = drStaffRole("STAFF").ToString()

                        If sRole <> "" Then
                            drNew("ROLENAME") = sRole.Substring(1)
                        End If

                        dtStaffRole.Rows.Add(drNew)

                        sStaff = drStaffRole("STAFF").ToString()
                        sRole = ""
                    Else
                        Continue For
                    End If
                Next

                dgUser.DataSource = dtStaffRole
                dgUser.DataBind()

                If dtStaffRole.Rows.Count > 0 Then
                    Me.trExcel.Visible = True
                    '添加另存EXCEL
                    Dim sURL1 As String = ""
                    sURL1 = sURL1 & "&Branch=" & ddlBranch.SelectedValue
                    Dim sURL As String = com.Azion.EloanUtility.UIUtility.getSYPath() & "SY_QRYAUTHEXCEL.aspx?Report=1" & sURL1
                    btntoexcel.Attributes("onclick") = "pop_DetailPage('" & sURL & "')"
                Else
                    Me.trExcel.Visible = false
                End If
            Else
                Me.trExcel.Visible = False
                trPagerMenu.Visible = False

                ' 顯示空表頭
                com.Azion.EloanUtility.FieldValidator.ShowEmptyGridView(dgUser, dtStaffRole)
            End If
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 綁定角色人員清單
    ''' </summary>
    ''' <remarks></remarks>
    Sub bindRoleStaff()
        Try
            Dim syRelRoleUserList As New AUTH_OP.SY_REL_ROLE_USERList(GetDatabaseManager())
            Dim dtRoleStaff As DataTable = createNewTableSchema()
            Dim sStaff As String = String.Empty
            Dim sRole As String = String.Empty

            trRole.Visible = True
            trUser.Visible = False
            trTime.Visible = False
            trChange.Visible = False

            ' 查詢人員角色列表
            If syRelRoleUserList.getRoleStaffList(ddlBranch.SelectedValue) Then
                For Each drRoleStaff As DataRow In syRelRoleUserList.getCurrentDataSet().Tables(0).Rows
                    If sRole <> drRoleStaff("ROLENAME").ToString() Then
                        Dim drCollection As DataRow() = syRelRoleUserList.getCurrentDataSet().Tables(0).Select("ROLENAME='" & drRoleStaff("ROLENAME").ToString() & "'")
                        Dim drNew As DataRow = dtRoleStaff.NewRow()

                        For Each drStaff In drCollection
                            sStaff = sStaff & "、" & drStaff("STAFF").ToString()
                        Next

                        drNew("ROLENAME") = drRoleStaff("ROLENAME").ToString()

                        If sStaff <> "" Then
                            drNew("STAFF") = sStaff.Substring(1)
                        End If

                        dtRoleStaff.Rows.Add(drNew)

                        sRole = drRoleStaff("ROLENAME").ToString()
                        sStaff = ""
                    Else
                        Continue For
                    End If
                Next

                dgRole.DataSource = dtRoleStaff
                dgRole.DataBind()

                If dtRoleStaff.Rows.Count > 0 Then
                    Me.trExcel.Visible = True
                    '添加另存EXCEL
                    Dim sURL1 As String = ""
                    sURL1 = sURL1 & "&Branch=" & ddlBranch.SelectedValue
                    Dim sURL As String = com.Azion.EloanUtility.UIUtility.getSYPath() & "SY_QRYAUTHEXCEL.aspx?Report=2" & sURL1
                    btntoexcel.Attributes("onclick") = "pop_DetailPage('" & sURL & "')"
                Else
                    Me.trExcel.Visible = False
                End If
            Else
                Me.trExcel.Visible = False
                trPagerMenu.Visible = False

                ' 顯示空表頭
                com.Azion.EloanUtility.FieldValidator.ShowEmptyGridView(dgRole, dtRoleStaff)
            End If
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 綁定異動人員清單
    ''' </summary>
    ''' <remarks></remarks>
    Sub bindChange()
        Try
            Dim syRelRoleUserList As New AUTH_OP.SY_REL_ROLE_USER(GetDatabaseManager())
            Dim dtChangeStaff As DataTable

            trTime.Visible = True
            trChange.Visible = True
            trRole.Visible = False
            trUser.Visible = False

            ddlStartYear.SelectedIndex = 0
            ddlStartMonth.SelectedIndex = 0
            ddlStartDay.SelectedIndex = 0
            ddlEndYear.SelectedIndex = 0
            ddlEndMonth.SelectedIndex = 0
            ddlEndDay.SelectedIndex = 0

            ' 查詢人員角色列表
            trPagerMenu.Visible = False

            dtChangeStaff = syRelRoleUserList.getChangeStaffRoleList(ddlBranch.SelectedValue)
            dtChangeStaff.Rows.Clear()

            ' 顯示空表頭
            For Each col As DataColumn In dtChangeStaff.Columns
                col.AllowDBNull = True
            Next

            com.Azion.EloanUtility.FieldValidator.ShowEmptyGridView(dgChange, dtChangeStaff)
        Catch ex As Exception
            SYUIBase.ShowErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 點擊“查詢”時查詢方法
    ''' </summary>
    ''' <remarks>[Lake] 2012/07/16 Created</remarks>
    Sub bindChangeStaff()
        Dim syRelRoleUserList As New AUTH_OP.SY_REL_ROLE_USER(GetDatabaseManager())

        ' 查詢人員角色列表
        Dim dtRelRoleUser As DataTable

        dtRelRoleUser = syRelRoleUserList.getChangeStaffRoleList(
            ddlBranch.SelectedValue,
            ddlStartYear.SelectedValue, ddlStartMonth.SelectedValue, ddlStartDay.SelectedValue,
            ddlEndYear.SelectedValue, ddlEndMonth.SelectedValue, ddlEndDay.SelectedValue)

        If IsNothing(dtRelRoleUser) = False Then
            dgChange.DataSource = dtRelRoleUser
            dgChange.DataBind()

            If dtRelRoleUser.Rows.Count > 0 Then
                Me.trExcel.Visible = True
                '添加另存EXCEL
                Dim sURL1 As String = ""
                sURL1 = sURL1 & "&Branch=" & ddlBranch.SelectedValue & "&StartYear=" & ddlStartYear.SelectedValue
                sURL1 = sURL1 & "&StartMonth=" & ddlStartMonth.SelectedValue & "&StartDay=" & ddlStartDay.SelectedValue
                sURL1 = sURL1 & "&EndYear=" & ddlEndYear.SelectedValue + "&EndMonth=" & ddlEndMonth.SelectedValue
                sURL1 = sURL1 & "&EndDay=" & ddlEndDay.SelectedValue
                Dim sURL As String = com.Azion.EloanUtility.UIUtility.getSYPath() & "SY_QRYAUTHEXCEL.aspx?Report=3" & sURL1
                btntoexcel.Attributes("onclick") = "pop_DetailPage('" & sURL & "')"
            Else
                Me.trExcel.Visible = False
            End If
        Else
            Me.trExcel.Visible = False
            trPagerMenu.Visible = False

            ' 顯示空表頭
            com.Azion.EloanUtility.FieldValidator.ShowEmptyGridView(dgChange, dtRelRoleUser)
        End If
    End Sub
#End Region

#Region "Event"
    ''' <summary>
    ''' 依人員查詢
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>[Lake] 2012/05/15 Created</remarks>
    Protected Sub rdoUser_CheckedChanged(sender As Object, e As EventArgs) Handles rdoUser.CheckedChanged, ddlBranch.SelectedIndexChanged
        Try
            If rdoUser.Checked Then

                ' 綁定人員角色清單
                bindStaffRole()
            End If
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 依角色查詢
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>[Lake] 2012/05/15 Created</remarks>
    Protected Sub rdoRole_CheckedChanged(sender As Object, e As EventArgs) Handles rdoRole.CheckedChanged, ddlBranch.SelectedIndexChanged
        Try
            If rdoRole.Checked Then

                ' 綁定角色人員清單
                bindRoleStaff()
            End If
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 異動查詢
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>[Lake] 2012/05/15 Created</remarks>
    Protected Sub rdoChange_CheckedChanged(sender As Object, e As EventArgs) Handles rdoChange.CheckedChanged, ddlBranch.SelectedIndexChanged
        Try
            If rdoChange.Checked Then
                'Dim syRelRoleUserHisList As New AUTH_OP.SY_REL_ROLE_USER_HISList(GetDatabaseManager())
                'Dim dtStaffRole As New DataTable()

                'If syRelRoleUserHisList.getEndTimeList(ddlBranch.SelectedValue) Then
                '    dtStaffRole = syRelRoleUserHisList.getCurrentDataSet().Tables(0)
                'End If

                'ddlStartYear.DataSource = dtStaffRole
                'ddlStartYear.DataValueField = "ENDYEAR"
                'ddlStartYear.DataTextField = "ENDYEAR"
                'ddlStartYear.DataBind()

                'ddlEndYear.DataSource = dtStaffRole
                'ddlEndYear.DataValueField = "ENDYEAR"
                'ddlEndYear.DataTextField = "ENDYEAR"
                'ddlEndYear.DataBind()

                'ddlStartYear.Items.Insert(0, New ListItem("請選擇", "0"))
                'ddlEndYear.Items.Insert(0, New ListItem("請選擇", "0"))

                bindChange()
            End If
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 查詢
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>[Lake] 2012/06/05 Created</remarks>
    Protected Sub btSearch_Click(sender As Object, e As EventArgs) Handles btSearch.Click
        Try
            ' 查詢異動資料
            bindChangeStaff()
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub
#End Region

  

End Class