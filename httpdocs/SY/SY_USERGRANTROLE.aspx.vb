''' <summary>
''' 程式說明：人員分派
''' 建立者：Zack 
''' 建立日期：2012-05-16
''' </summary>

Imports com.Azion.NET.VB
Imports AUTH_OP.TABLE
Imports com.Azion.EloanUtility
Imports AUTH_OP
Imports System.Xml
Imports FLOW_OP

Public Class SY_USERGRANTROLE
    Inherits SYUIBase

#Region "維護區"
#Region "Page_load"


    Dim m_dtFlow As New DataTable

    ''' <summary>
    ''' 頁面初始加載事件
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            m_bDisableAllControlsInDisplayMode = True

            If m_bTesting OrElse Request("TESTMODE") = "1" Then
                m_bTesting = True

                Dim loginUser As New SY_LOGIN(GetDatabaseManager)
                loginUser.getUserInfo("716", "S001443")
                'm_bCheck = False
                Session("HOFLAG") = "1"
                'm_bDisplayMode = False
                MyBase.InitParas()
            End If


            ' 初始化頁面參數
            initParas()

            If Not IsPostBack Then

                ' 設置初始焦點
                rdolistSetStyle.Focus()

                ' 設定角色選中默認值
                rdolistSetStyle.SelectedValue = "2"

                ' 沒有異動的數據集合
                m_dtFlow.Columns.Add("LValue", Type.GetType("System.String"))
                m_dtFlow.Columns.Add("LText", Type.GetType("System.String"))

                ViewState("dataFlow") = m_dtFlow

                ' 異動過的數據集合
                Dim dtListBoxMove As New DataTable

                dtListBoxMove.Columns.Add("VALUE", Type.GetType("System.String"))
                dtListBoxMove.Columns.Add("TEXT", Type.GetType("System.String"))
                dtListBoxMove.Columns.Add("OPERATION", Type.GetType("System.String"))

                ViewState("dtListBoxMove") = dtListBoxMove

                ' 隱藏分派清單
                trByRoleDetail.Style.Value = "display:none"
                trByPeopleDetail.Style.Value = "display:none"

                ' 綁定單位下拉選單
                initDDlDepart()

                ' 初始化頁面數據
                initData()

                bindReivewer()
            End If

            If m_bDisplayMode Then
                btnAgree.Enabled = False

            End If

        Catch ex As Exception
            SYUIBase.ShowErrMsg(Me, ex)
        End Try
    End Sub

    Private Sub Page_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.PreRender
        updateDDLRoleColor()
    End Sub


#End Region

#Region "Function"

    ''' <summary>
    ''' 初始化參數
    ''' </summary>
    ''' <remarks></remarks>
    Sub initParas()

        Dim syBranch As New AUTH_OP.SY_BRANCH(GetDatabaseManager())

        If Not String.IsNullOrEmpty(m_sCaseId) Then

            ' 案號
            lblCaseId.Text = m_sCaseId

            ' 單位
            If syBranch.loadByCaseId(m_sCaseId) Then
                lblBranch.Text = syBranch.getAttribute("BRCNAME")
            End If
        End If


        ' 測試數據
        'm_bTesting = True
        'If m_bTesting Then
        '    m_sWorkingUserid = "S000035"
        '    m_sWorkingTopDepNo = "1"
        '    m_sFuncCode = "41"
        '    m_sStepNo = ""
        '    m_sCaseId = ""
        '    m_sFourStepNo = ""
        '    m_sHoFlag = "1"
        '    m_sLoginUserid = "S000035"
        '    m_sWorkingBrid = "904"
        '    m_bDisplayMode = False

        '    Dim syLogin As New SY_LOGIN(GetDatabaseManager)
        '    syLogin.getUserInfo("924", "S000914")

        '    FLOW_OP.ELoanFlow.m_callbackUserInfo = New com.Azion.UITools.ExportUserInfo
        '    MyBase.InitParas()
        'End If
    End Sub

    ''' <summary>
    ''' 綁定單位下拉選單
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub initDDlDepart()
        Try
            Dim syBranchLIst As New AUTH_OP.SY_BRANCHList(GetDatabaseManager())
            Dim dt As DataTable = syBranchLIst.loadBranchInfo(m_sWorkingBrid, m_sHoFlag, "0", 5)


            '顯示模式只列出變更所在的分行
            If m_bDisplayMode Then
                Dim drc As DataRowCollection = BosBase.getNewBosBase("SY_REL_ROLE_USER_HIS", GetDatabaseManager).GetDataRowCollection( _
                    "select BRID" & vbCrLf & _
                    "  from SY_REL_ROLE_USER_HIS RRU" & vbCrLf & _
                    " inner join SY_BRANCH BR" & vbCrLf & _
                    "    on RRU.BRA_DEPNO = BR.BRA_DEPNO" & vbCrLf & _
                    " where CASEID = @CASEID@" & vbCrLf & _
                    "   and OPERATION <> 'N'" & vbCrLf,
                    "CASEID", m_sCaseId)

                If IsNothing(drc) = False Then
                    Dim sFilter As New StringBuilder()

                    For Each dr As DataRow In drc
                        If sFilter.Length = 0 Then
                            sFilter.Append("'" & CStr(dr("BRID")) & "'")
                        Else
                            sFilter.Append(",'" & CStr(dr("BRID")) & "'")
                        End If
                    Next

                    Dim dt2 As DataTable = dt.Clone
                    Dim dra() As DataRow = dt.Select("BRID in (" & sFilter.ToString() & ")")

                    For Each dr In dra
                        dt2.ImportRow(dr)
                    Next

                    dt = dt2
                Else
                    dt = dt.Clone
                End If
            End If


            ' 綁定"單位"下拉選單
            If Not dt Is Nothing Then
                ddlDepart.DataSource = dt
                ddlDepart.DataTextField = "NAME"
                ddlDepart.DataValueField = "BRA_DEPNO"
                ddlDepart.DataBind()
            End If

            ddlDepart.Items.Insert(0, New ListItem("--請選擇--", ""))

            ' 若單位只有一筆資料，默認選中此筆資料
            If m_sHoFlag <> "1" Then
                ddlDepart.SelectedIndex = 1
            End If

            ' 綁定人員名稱或角色名稱
            bindDropDownList()
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    Sub bindReivewer()

        Dim flowFacade As FLOW_OP.FlowFacade
        Dim dataTable As DataTable
        Dim dbTable As BosBase

        Try
            flowFacade = New FLOW_OP.FlowFacade(GetDatabaseManager)

            Dim sRoleIds() As String = m_sWorkingRoleId.Split(",")

            '如果是系統角色，送件不需要覆核
            For Each sRID As String In sRoleIds
                If flowFacade.getSYRole.GetRoleType(CInt(sRID)) = "1" Then
                    divReviewer.Style.Value = "display:none"
                    Return
                End If
            Next

            dataTable = flowFacade.getSYRelRoleFlowMap.GetDataTable( _
                        "select distinct UR.STAFFID, UR.STAFFID + '  ' + UR.USERNAME as DISPLAY " & vbCrLf & _
                        "  from SY_REL_ROLE_FLOWMAP RRF " & vbCrLf & _
                        " inner join SY_REL_ROLE_USER RRU " & vbCrLf & _
                        "    on RRF.ROLEID = RRU.ROLEID " & vbCrLf & _
                        " inner join SY_BRANCH BR " & vbCrLf & _
                        "    on BR.BRA_DEPNO = RRU.BRA_DEPNO " & vbCrLf & _
                        " inner join SY_USER UR " & vbCrLf & _
                        "    on UR.STAFFID = RRU.STAFFID " & vbCrLf & _
                        " inner join SY_FLOW_ID FI " & vbCrLf & _
                        "    on FI.FLOW_ID = RRF.FLOW_ID " & vbCrLf & _
                        " where FLOW_NAME = 'SY_ROLEUSER' " & vbCrLf & _
                        "   and STEP_NO = 'SY000400' " & vbCrLf & _
                        "   and BRID = @BRID@ " & vbCrLf & _
                        "   and UR.STAFFID <> @STAFFID@ " & vbCrLf & _
                        " order by STAFFID " & vbCrLf,
                        "BRID", m_sWorkingBrid,
                        "STAFFID", m_sWorkingUserid)

            ddlReviewer.DataSource = dataTable
            ddlReviewer.DataTextField = "DISPLAY"
            ddlReviewer.DataValueField = "STAFFID"

            ddlReviewer.DataBind()
        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' 綁定ListBox里的值
    ''' </summary>
    ''' <param name="sBRID">分行代號</param>
    ''' <param name="sRoleOrStaffID">角色編號或者人員編號</param>
    ''' <remarks></remarks>
    Sub bindListBox()
        Try
            Dim syUserList As New AUTH_OP.SY_USERList(GetDatabaseManager())
            Dim syRoleList As New AUTH_OP.SY_ROLEList(GetDatabaseManager())

            Dim dsUnSelect As New DataSet ' 帶選擇資料集
            Dim dsSelect As New DataSet  ' 已選擇資料集

            ' 依人員
            If rdolistSetStyle.SelectedValue = "1" Then

                ' 若有選擇人員信息
                If Not String.IsNullOrEmpty(ddlUser.SelectedItem.Value) Then

                    Dim syDynamic As New FLOW_OP.TABLE.SY_DYNAMICSQL(GetDatabaseManager())
                    Dim sDynamic As String = syDynamic.GetDynamicSQL("SY_USERGRANTROLE:BINDDROPDOWNLIST")

                    If String.IsNullOrEmpty(sDynamic) = False Then
                        sDynamic = DynamicSQL(sDynamic,
                                              "@BRID@", syDynamic.getSYBranch.GetBRID(ddlDepart.SelectedItem.Value))
                    End If

                    ' 如果能查詢出“待選擇角色”資料
                    If syRoleList.loadByBraDepnoStaffid("Y", m_sHoFlag, "," & m_sWorkingRoleId & ",", ddlDepart.SelectedItem.Value, ddlUser.SelectedItem.Value.Split(";")(0), "N", sDynamic) Then
                        dsUnSelect = syRoleList.getCurrentDataSet
                    End If

                    ' 如果能查詢出“已選擇角色”資料
                    If syRoleList.loadByBraDepnoStaffid("Y", m_sHoFlag, "," & m_sWorkingRoleId & ",", ddlDepart.SelectedItem.Value, ddlUser.SelectedItem.Value.Split(";")(0), "Y") Then
                        dsSelect = syRoleList.getCurrentDataSet
                    End If
                End If
            Else

                ' 若有選擇角色信息
                If Not String.IsNullOrEmpty(ddlRole.SelectedItem.Value) Then

                    ' 如果能查詢出"待選人員資料"
                    If syUserList.loadByDepnoRoleID(ddlDepart.SelectedItem.Value, ddlRole.SelectedItem.Value) Then
                        dsUnSelect = syUserList.getCurrentDataSet
                    End If

                    ' 如果能查詢出"已選擇人員資料" 
                    If syUserList.loadInFlow(ddlDepart.SelectedItem.Value, ddlRole.SelectedItem.Value) Then
                        dsSelect = syUserList.getCurrentDataSet
                    End If
                End If
            End If

            ' 若有已選擇資料
            If dsSelect.Tables.Count > 0 Then
                m_dtFlow = dsSelect.Tables(0)
                ViewState("dataFlow") = m_dtFlow

                lstBoxRight.DataSource = dsSelect
                lstBoxRight.DataTextField = "LText"
                lstBoxRight.DataValueField = "LValue"
                lstBoxRight.DataBind()
            End If

            ' 若有待選擇資料
            If dsUnSelect.Tables.Count > 0 Then
                lstBoxLeft.DataSource = dsUnSelect
                lstBoxLeft.DataTextField = "LText"
                lstBoxLeft.DataValueField = "LValue"
                lstBoxLeft.DataBind()
            End If
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 綁定人員名稱或角色名稱下拉選單
    ''' </summary>
    ''' <remarks></remarks>
    Sub bindDropDownList()

        Try
            Dim syRoleList As New AUTH_OP.SY_ROLEList(GetDatabaseManager())
            Dim syUserList As New AUTH_OP.SY_USERList(GetDatabaseManager())
            Dim syDynamic As New FLOW_OP.TABLE.SY_DYNAMICSQL(GetDatabaseManager())

            ' 查詢角色名稱
            If ddlDepart.SelectedIndex <> 0 Then

                Dim sDynamic As String = ""

                If String.IsNullOrEmpty(m_sHoFlag) = False Then
                    sDynamic = DynamicSQL(syDynamic.GetDynamicSQL("SY_USERGRANTROLE:BINDDROPDOWNLIST"),
                                          "@BRID@", syDynamic.getSYBranch.GetBRID(ddlDepart.SelectedItem.Value))
                End If


                If syRoleList.loadByBraDepnoStaffid("N", m_sHoFlag, "," & m_sWorkingRoleId & ",", ddlDepart.SelectedItem.Value, "", "", sDynamic) Then

                    Dim dt As DataTable = syRoleList.getCurrentDataSet.Tables(0)

                    '顯示模式只列出變更所在的分行
                    If m_bDisplayMode Then
                        Dim drc As DataRowCollection = BosBase.getNewBosBase("SY_REL_ROLE_USER_HIS", GetDatabaseManager).GetDataRowCollection( _
                            "select ROLEID" & vbCrLf & _
                            "  from SY_REL_ROLE_USER_HIS RRU" & vbCrLf & _
                            " where CASEID = @CASEID@" & vbCrLf & _
                            "   and OPERATION <> 'N'" & vbCrLf & _
                            "   and BRA_DEPNO = @BRA_DEPNO@" & vbCrLf,
                            "CASEID", m_sCaseId,
                            "BRA_DEPNO", ddlDepart.SelectedItem.Value)

                        If IsNothing(drc) = False Then
                            Dim sFilter As New StringBuilder()

                            For Each dr As DataRow In drc
                                If sFilter.Length = 0 Then
                                    sFilter.Append(CStr(dr("ROLEID")))
                                Else
                                    sFilter.Append("," & CStr(dr("ROLEID")))
                                End If
                            Next

                            Dim dt2 As DataTable = dt.Clone
                            Dim dra() As DataRow = dt.Select("LValue in (" & sFilter.ToString() & ")")

                            For Each dr In dra
                                dt2.ImportRow(dr)
                            Next

                            dt = dt2
                        Else
                            dt = dt.Clone
                        End If
                    End If

                    ddlRole.DataSource = dt
                    ddlRole.DataTextField = "LText"
                    ddlRole.DataValueField = "LValue"
                    ddlRole.DataBind()
                End If
            End If

            ddlRole.Items.Insert(0, New ListItem("--請選擇--", ""))

            ' 綁定【人員名稱】下拉選單
            If ddlDepart.SelectedIndex <> 0 Then
                If syUserList.loadByBRA_DEPNO(ddlDepart.SelectedItem.Value) Then
                    ddlUser.DataSource = syUserList.getCurrentDataSet
                    ddlUser.DataTextField = "USERNAME"
                    ddlUser.DataValueField = "BRA_DEPNO"
                    ddlUser.DataBind()
                End If
            End If

            ddlUser.Items.Insert(0, New ListItem("--請選擇--", ""))

            ' 若角色下拉選單只有一筆資料
            If ddlRole.Items.Count = 2 Then
                ddlRole.SelectedIndex = 1

                ' 依角色
                If rdolistSetStyle.SelectedValue = "2" Then

                    ' 清空左右兩個ListBox的的值
                    lstBoxLeft.Items.Clear()
                    lstBoxRight.Items.Clear()

                    ' 重新綁定ListBox
                    initXmlListBox("2")

                    ' 將移動過的數據變色
                    editListItemColor()
                End If
            End If
        Catch ex As Exception
            SYUIBase.ShowErrMsg(Me, ex)
        End Try
    End Sub


    ''' <summary>
    ''' 設定上層角色的顏色為GRAY，且不可以選入員工
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub updateDDLRoleColor()
        Try
            btnIn.Enabled = True
            Dim roleTable As New BosBase("SY_ROLE", GetDatabaseManager)
            Dim drc As DataRowCollection
            Dim hash As New Hashtable

            If rdolistSetStyle.SelectedValue = "2" Then
                If ddlDepart.SelectedIndex <> 0 Then

                    drc = roleTable.GetDataRowCollection(
                                        "select ROLEID " & vbCrLf & _
                                        "  from SY_ROLE " & vbCrLf & _
                                        " where ROLEID not in " & vbCrLf & _
                                        "       (select distinct parent from SY_ROLE where PARENT <> 0) " & vbCrLf, Nothing)

                    For Each dr As DataRow In drc
                        hash.Add(CStr(dr("ROLEID")), 0)
                    Next

                    For Each ddlRoleitem As ListItem In ddlRole.Items
                        If String.IsNullOrEmpty(ddlRoleitem.Value) = False Then
                            If hash.ContainsKey(ddlRoleitem.Value) = False Then
                                ddlRoleitem.Attributes.Add("style", "color:grey")
                            End If
                        End If
                    Next

                End If

                '上層角色不可設定人員
                If String.IsNullOrEmpty(ddlRole.SelectedValue) = False Then
                    Dim nRoleId As Integer = 0
                    Dim nCount As Integer = 0

                    Try
                        nRoleId = CInt(ddlRole.SelectedValue)
                    Catch ex As Exception
                    End Try

                    nCount = roleTable.ExecuteScalar(
                                "select COUNT(*) " & vbCrLf & _
                                "  from SY_ROLE " & vbCrLf & _
                                " where ROLEID not in " & vbCrLf & _
                                "       (select distinct parent from SY_ROLE where PARENT <> 0) " & vbCrLf & _
                                "   and ROLEID = @ROLEID@ " & vbCrLf,
                                "ROLEID", nRoleId)
                    If nCount = 0 Then
                        btnIn.Enabled = False
                    End If
                End If
            Else
                If String.IsNullOrEmpty(ddlUser.SelectedValue) = False Then
                    drc = roleTable.GetDataRowCollection(
                                        "select ROLEID " & vbCrLf & _
                                        "  from SY_ROLE " & vbCrLf & _
                                        " where ROLEID not in " & vbCrLf & _
                                        "       (select distinct parent from SY_ROLE where PARENT <> 0) " & vbCrLf, Nothing)

                    For Each dr As DataRow In drc
                        hash.Add(CStr(dr("ROLEID")), 0)
                    Next

                    For Each listItem As ListItem In lstBoxLeft.Items
                        If hash.ContainsKey(listItem.Value) = False Then
                            listItem.Attributes.Add("style", "color:grey")
                        End If
                    Next

                End If
            End If

        Catch ex As Exception
            Throw
        End Try
    End Sub


    ''' <summary>
    ''' 正常區塊:如果臨時檔中有資料，擇顯示在ListBox中
    ''' </summary>
    ''' <param name="sType"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function initXmlListBox(ByVal sType As String)

        Dim syUserList As New AUTH_OP.SY_USERList(GetDatabaseManager())
        Dim syRoleList As New AUTH_OP.SY_ROLEList(GetDatabaseManager())
        Dim syTempInfoXml As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager(), IIf(String.IsNullOrEmpty(m_sCaseId), "SY_TEMPINFO", "SY_TEMPINFO2"))

        Dim oldXmlData As String = String.Empty ' 臨時檔資料
        Dim XmlDocument As New XmlDocument
        Dim bXmlhasData As Boolean = False

        ' 沒有異動數據的集合 
        Dim dtFlow As DataTable = ViewState("dataFlow")

        ' 若沒有已選擇資料
        If dtFlow.Columns.Count = 0 Then
            dtFlow.Columns.Add("LValue", Type.GetType("System.String"))
            dtFlow.Columns.Add("LText", Type.GetType("System.String"))
        End If

        ' 員工編號
        Dim sStaffids As String = String.Empty

        ' 角色編號
        Dim sRoleids As String = String.Empty

        ' 如果案件有被退回
        If m_sFourStepNo = "0300" AndAlso syTempInfoXml.loadByCaseId(m_sCaseId) Then

            ' 依據CaseID查詢數據
            oldXmlData = syTempInfoXml.getAttribute("TEMPDATA")
        ElseIf m_sFourStepNo = "" AndAlso syTempInfoXml.loadByPK(m_sWorkingUserid, m_sWorkingTopDepNo, m_sFuncCode) Then
            oldXmlData = syTempInfoXml.getAttribute("TEMPDATA")
        Else
            ' 綁定角色ListBox
            bindListBox()
        End If

        ' 若臨時檔有資料
        If oldXmlData <> "" Then

            ' document對象載入XML文件
            XmlDocument.LoadXml(oldXmlData)

            If XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER").Count > 0 Then
                For i As Integer = 0 To XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER").Count - 1
                    Dim sRID As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("ROLEID"), System.Xml.XmlElement).InnerText
                    Dim sSTAFFID As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("STAFFID"), System.Xml.XmlElement).InnerText
                    Dim sBRA_DEPNO As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("BRA_DEPNO"), System.Xml.XmlElement).InnerText
                    Dim sOPERATION As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("OPERATION"), System.Xml.XmlElement).InnerText
                    Dim TYPE As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("TYPE"), System.Xml.XmlElement).InnerText
                    Dim sBRANAME As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("BRANAME"), System.Xml.XmlElement).InnerText
                    Dim sUSERNAME As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("USERNAME"), System.Xml.XmlElement).InnerText
                    Dim sROLENAME As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("ROLENAME"), System.Xml.XmlElement).InnerText

                    ' 如果設定方式和臨時檔中的值相同
                    If TYPE = sType Then

                        ' 依角色
                        If TYPE = 2 Then
                            If sRID = ddlRole.SelectedItem.Value AndAlso ddlDepart.SelectedItem.Text.Trim() = sBRANAME Then
                                bXmlhasData = True

                                ' 不是刪除的資料
                                If sOPERATION <> "D" Then
                                    lstBoxRight.Items.Add(New ListItem(sUSERNAME, sSTAFFID + ";" + sBRA_DEPNO))
                                    sStaffids = sStaffids & "'" & sSTAFFID & "',"

                                    ' 如果是沒有異動的數據
                                    If sOPERATION = "N" Then
                                        Dim row As DataRow = dtFlow.NewRow
                                        row("LValue") = sSTAFFID & ";" & sBRA_DEPNO
                                        row("LText") = sUSERNAME

                                        dtFlow.Rows.Add(row)
                                    End If
                                End If
                            End If
                        Else
                            If sSTAFFID = ddlUser.SelectedItem.Value.Split(";")(0) AndAlso ddlDepart.SelectedItem.Text.Trim() = sBRANAME Then
                                bXmlhasData = True

                                ' 不是刪除的資料
                                If sOPERATION <> "D" Then
                                    lstBoxRight.Items.Add(New ListItem(sROLENAME, sRID))
                                    sRoleids = sRoleids & sRID & ","

                                    ' 如果是沒有異動的數據
                                    If sOPERATION = "N" Then
                                        Dim row As DataRow = dtFlow.NewRow
                                        row("LValue") = sRID
                                        row("LText") = sROLENAME

                                        dtFlow.Rows.Add(row)
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next

                ViewState("dataFlow") = dtFlow

                ' 如果設定方式為依角色
                If sType = "2" Then

                    If sStaffids <> "" Then
                        ' 如果有人員編號
                        If sStaffids.Length > 0 Then
                            sStaffids = sStaffids.Substring(0, sStaffids.Length - 1)
                        End If

                        ' 如果能查詢出"待選人員資料"
                        If syUserList.loadByRoleEdit(sStaffids, ddlDepart.SelectedItem.Value) Then
                            lstBoxLeft.DataSource = syUserList.getCurrentDataSet
                            lstBoxLeft.DataTextField = "LText"
                            lstBoxLeft.DataValueField = "LValue"
                            lstBoxLeft.DataBind()
                        End If
                    Else

                        ' 若臨時檔中沒有資料
                        If bXmlhasData = False OrElse String.IsNullOrEmpty(sStaffids) = True Then
                            bindListBox()
                        Else

                            ' 如果能查詢出"待選人員資料"
                            If syUserList.loadByRoleEdit(sStaffids, ddlDepart.SelectedItem.Value) Then
                                lstBoxLeft.DataSource = syUserList.getCurrentDataSet
                                lstBoxLeft.DataTextField = "LText"
                                lstBoxLeft.DataValueField = "LValue"
                                lstBoxLeft.DataBind()
                            End If
                        End If

                    End If
                Else

                    If sRoleids <> "" Then

                        ' 如果有人員編號
                        If sRoleids.Length > 0 Then
                            sRoleids = sRoleids.Substring(0, sRoleids.Length - 1)
                        End If

                        ' 如果能查詢出“待選擇角色”資料
                        If syRoleList.loadByPeopleEdit(ddlDepart.SelectedItem.Value, sRoleids, m_sHoFlag, "," & m_sWorkingRoleId & ",") Then
                            lstBoxLeft.DataSource = syRoleList.getCurrentDataSet
                            lstBoxLeft.DataTextField = "LText"
                            lstBoxLeft.DataValueField = "LValue"
                            lstBoxLeft.DataBind()
                        End If
                    Else

                        ' 若臨時檔中沒有資料
                        If bXmlhasData = False Then
                            bindListBox()
                        Else

                            ' 如果能查詢出“待選擇角色”資料
                            If syRoleList.loadByPeopleEdit(ddlDepart.SelectedItem.Value, sRoleids, m_sHoFlag, "," & m_sWorkingRoleId & ",") Then
                                lstBoxLeft.DataSource = syRoleList.getCurrentDataSet
                                lstBoxLeft.DataTextField = "LText"
                                lstBoxLeft.DataValueField = "LValue"
                                lstBoxLeft.DataBind()
                            End If
                        End If
                    End If
                End If
            Else

                ' 綁定角色ListBox
                bindListBox()
            End If
        Else

            ' 綁定角色ListBox
            bindListBox()
        End If
    End Function

    ''' <summary>
    ''' 正常區塊取得XmlData
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getXmlData(ByVal syTempInfo As AUTH_OP.SY_TEMPINFO) As String
        Try
            ' 如果案件有被退回
            If m_sFourStepNo = "" Then
                If syTempInfo.loadByPK(m_sWorkingUserid, m_sWorkingTopDepNo, m_sFuncCode) Then
                    Return syTempInfo.getAttribute("TEMPDATA")
                End If
            Else
                If syTempInfo.loadByCaseId(m_sCaseId) Then
                    Return syTempInfo.getAttribute("TEMPDATA")
                End If
            End If
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Function

    ''' <summary>
    ''' 初始化數據
    ''' </summary>
    ''' <remarks></remarks>
    Sub initData()
        Try

            'Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager())
            Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager(), IIf(String.IsNullOrEmpty(m_sCaseId), "SY_TEMPINFO", "SY_TEMPINFO2"))


            Dim sXmlData As String = String.Empty
            Dim xmlDocument As New XmlDocument
            Dim sStepNO As String = String.Empty  ' 步驟號

            ' 取得臨時檔資料
            sXmlData = getXmlData(syTempInfo)

            ' 若臨時檔有資料
            If sXmlData <> "" Then
                ' document對象載入XML文件
                xmlDocument.LoadXml(sXmlData)

                If xmlDocument.GetElementsByTagName("SY_REL_ROLE_USER").Count > 0 Then
                    rdolistSetStyle.SelectedValue = DirectCast(xmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(0)("TYPE"), System.Xml.XmlElement).InnerText
                End If
            End If

            ' 顯示維護頁面
            If m_sFourStepNo = "" Or m_sFourStepNo = "0300" Then

                ' 顯示編輯區
                divEdit.Style.Value = "display:block"
                divChecked.Style.Value = "display:none"

                ' 綁定分派清單
                BindDetail("1")

                ' 如果設定方式為“依角色”
                If rdolistSetStyle.SelectedValue = "2" Then

                    ' 顯示【角色名稱】下拉選單欄位，【人員名稱】下拉選單隱藏
                    trRole.Style.Value = "display:block"
                    trUser.Style.Value = "display:none"

                    ' 如果臨時檔有資料，綁定ListBox
                    initXmlListBox(rdolistSetStyle.SelectedValue)
                Else

                    ' 顯示【人員名稱】下拉選單欄位，【角色名稱】下拉選單隱藏
                    trRole.Style.Value = "display:none"
                    trUser.Style.Value = "display:block"

                    ' 如果臨時檔有資料，綁定ListBox
                    initXmlListBox(rdolistSetStyle.SelectedValue)
                End If

                ' 有移動的數據變紅
                editListItemColor()
            Else

                ' 顯示復核區
                divQuery.Style.Value = "display:none"
                divEdit.Style.Value = "display:none"
                divChecked.Style.Value = "display:block"
                divSend.Style.Value = "display:none"

                ' 綁定審核區的資歷
                initCheckListBox()

                ' 顯示按鈕區
                divBtnChecked.Style.Value = "display:block;text-align :center"
            End If
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 綁定分派清單
    ''' </summary>
    ''' <remarks></remarks>
    Sub BindDetail(ByVal sSelected As String)
        Try
            'Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager())
            Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager(), IIf(String.IsNullOrEmpty(m_sCaseId), "SY_TEMPINFO", "SY_TEMPINFO2"))

            Dim sXmlData As String = String.Empty
            Dim xmlDocument As New XmlDocument

            ' 標記是否有數據
            Dim sCheckType As String = String.Empty

            ' 取得臨時檔資料
            sXmlData = getXmlData(syTempInfo)

            ' 若臨時檔有資料
            If sXmlData <> "" Then

                ' document對象載入XML文件
                xmlDocument.LoadXml(sXmlData)

                If xmlDocument.GetElementsByTagName("SY_REL_ROLE_USER").Count > 0 Then
                    sCheckType = DirectCast(xmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(0)("TYPE"), System.Xml.XmlElement).InnerText
                End If
            Else
                ' 依人員
                If rdolistSetStyle.SelectedValue = "1" Then
                    trByPeopleDetail.Style.Value = "display:none"
                Else
                    trByRoleDetail.Style.Value = "display:none"
                End If

                If m_sHoFlag <> "1" Then
                    ddlDepart.SelectedIndex = 1

                    ' 重新绑定角色或人员信息
                    bindDropDownList()
                End If
            End If

            ' 若設定方式不為空,且與當前選擇地設定方式相同
            If sCheckType <> "" AndAlso sCheckType = rdolistSetStyle.SelectedValue Then
                ' 設定角色設置默認值
                rdolistSetStyle.SelectedValue = sCheckType

                Dim sRoleNames As String = String.Empty ' 角色名稱
                Dim sUserNames As String = String.Empty   ' 人員名稱
                Dim sStaffid As String = String.Empty  ' 員工編號
                Dim sInRoleID As String = String.Empty  ' 需存入的角色編號
                Dim sInDepart As String = String.Empty  ' 需存入的部門
                Dim sInStaffid As String = String.Empty  ' 需存入的員工編號
                Dim sInUsername As String = String.Empty  ' 需存入的角色名稱
                Dim sDeparType As String = String.Empty   ' 部門
                Dim sInBRA_DEPNO As String = String.Empty
                Dim sRoleType As String = String.Empty
                Dim sInRoleName As String = String.Empty
                Dim sInBraDepno As String = String.Empty

                ' 定義數據表存儲人員分派的資料
                Dim dtPeopleDetail As New DataTable
                dtPeopleDetail.Columns.Add("DEPART", Type.GetType("System.String"))
                dtPeopleDetail.Columns.Add("USERNAME", Type.GetType("System.String"))
                dtPeopleDetail.Columns.Add("ROLENAME", Type.GetType("System.String"))
                dtPeopleDetail.Columns.Add("STAFFID", Type.GetType("System.String"))
                dtPeopleDetail.Columns.Add("BRA_DEPNO", Type.GetType("System.String"))

                ' 定義數據表存儲角色分派的資料
                Dim dtRoleDetail As New DataTable
                dtRoleDetail.Columns.Add("DEPART", Type.GetType("System.String"))
                dtRoleDetail.Columns.Add("BRA_DEPNO", Type.GetType("System.String"))
                dtRoleDetail.Columns.Add("USERNAME", Type.GetType("System.String"))
                dtRoleDetail.Columns.Add("ROLENAME", Type.GetType("System.String"))
                dtRoleDetail.Columns.Add("ROLEID", Type.GetType("System.String"))

                ' 如果有數據
                If xmlDocument.GetElementsByTagName("SY_REL_ROLE_USER").Count > 1 Then
                    sStaffid = DirectCast(xmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(0)("STAFFID"), System.Xml.XmlElement).InnerText
                    sDeparType = DirectCast(xmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(0)("BRANAME"), System.Xml.XmlElement).InnerText
                    sRoleType = DirectCast(xmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(0)("ROLEID"), System.Xml.XmlElement).InnerText

                    For i As Integer = 0 To xmlDocument.GetElementsByTagName("SY_REL_ROLE_USER").Count - 1
                        Dim ROLEID As String = DirectCast(xmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("ROLEID"), System.Xml.XmlElement).InnerText
                        Dim STAFFID As String = DirectCast(xmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("STAFFID"), System.Xml.XmlElement).InnerText
                        Dim BRA_DEPNO As String = DirectCast(xmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("BRA_DEPNO"), System.Xml.XmlElement).InnerText
                        Dim OPERATION As String = DirectCast(xmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("OPERATION"), System.Xml.XmlElement).InnerText
                        Dim TYPE As String = DirectCast(xmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("TYPE"), System.Xml.XmlElement).InnerText
                        Dim BRANAME As String = DirectCast(xmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("BRANAME"), System.Xml.XmlElement).InnerText
                        Dim USERNAME As String = DirectCast(xmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("USERNAME"), System.Xml.XmlElement).InnerText
                        Dim ROLENAME As String = DirectCast(xmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("ROLENAME"), System.Xml.XmlElement).InnerText

                        If TYPE = sCheckType Then

                            ' 依人員
                            If sCheckType = "1" Then

                                ' 將員工編號，部門編號相同的組合成一筆資料
                                If sStaffid = STAFFID AndAlso sDeparType = BRANAME Then
                                    sInStaffid = STAFFID
                                    sInRoleID = ROLEID
                                    sInDepart = BRANAME
                                    sInUsername = USERNAME
                                    sInBRA_DEPNO = BRA_DEPNO

                                    ' 如果是新增的數據
                                    If OPERATION = "I" Then
                                        sRoleNames = sRoleNames + "(選入) " + ROLENAME + ","
                                    ElseIf OPERATION = "D" Then
                                        sRoleNames = sRoleNames + "(選出) " + ROLENAME + ","
                                    End If

                                    ' 若是最後一條數據
                                    If i = xmlDocument.GetElementsByTagName("SY_REL_ROLE_USER").Count - 1 Then
                                        Dim row As DataRow = dtPeopleDetail.NewRow
                                        row("DEPART") = sInDepart
                                        row("STAFFID") = sInStaffid
                                        row("USERNAME") = sInUsername
                                        row("BRA_DEPNO") = sInBRA_DEPNO

                                        If sRoleNames.Length > 0 Then
                                            row("ROLENAME") = sRoleNames.Substring(0, sRoleNames.Length - 1)
                                        End If

                                        dtPeopleDetail.Rows.Add(row)
                                    End If
                                Else
                                    ' 新增一條數據
                                    Dim row As DataRow = dtPeopleDetail.NewRow
                                    row("DEPART") = sInDepart
                                    row("STAFFID") = sInStaffid
                                    row("USERNAME") = sInUsername
                                    row("BRA_DEPNO") = sInBRA_DEPNO

                                    If sRoleNames.Length > 0 Then
                                        row("ROLENAME") = sRoleNames.Substring(0, sRoleNames.Length - 1)
                                    End If

                                    dtPeopleDetail.Rows.Add(row)

                                    sDeparType = BRANAME
                                    sInBRA_DEPNO = BRA_DEPNO
                                    sStaffid = STAFFID
                                    sInStaffid = STAFFID
                                    sInRoleID = ROLEID
                                    sInDepart = BRANAME
                                    sInUsername = USERNAME
                                    sRoleNames = ""

                                    ' 如果是新增的數據
                                    If OPERATION = "I" Then
                                        sRoleNames = sRoleNames + "(選入) " + ROLENAME + ","
                                    ElseIf OPERATION = "D" Then
                                        sRoleNames = sRoleNames + "(選出) " + ROLENAME + ","
                                    End If

                                    ' 如果是最後一條數據
                                    If i = xmlDocument.GetElementsByTagName("SY_REL_ROLE_USER").Count - 1 Then
                                        Dim rowData As DataRow = dtPeopleDetail.NewRow
                                        rowData("DEPART") = sInDepart
                                        rowData("STAFFID") = sInStaffid
                                        rowData("USERNAME") = sInUsername
                                        rowData("BRA_DEPNO") = sInBRA_DEPNO

                                        If sRoleNames.Length > 0 Then
                                            rowData("ROLENAME") = sRoleNames.Substring(0, sRoleNames.Length - 1)
                                        End If

                                        dtPeopleDetail.Rows.Add(rowData)
                                    End If
                                End If
                            Else

                                ' 如果角色，單位與第一條數據不同
                                If sRoleType = ROLEID AndAlso sDeparType = BRANAME Then
                                    sInRoleID = ROLEID
                                    sInDepart = BRANAME
                                    sInRoleName = ROLENAME
                                    sInBraDepno = BRA_DEPNO

                                    ' 如果是新增的數據
                                    If OPERATION = "I" Then
                                        sUserNames = sUserNames + "(選入) " + USERNAME + ","
                                    ElseIf OPERATION = "D" Then
                                        sUserNames = sUserNames + "(選出) " + USERNAME + ","
                                    End If

                                    ' 如果是最後一條數據
                                    If i = xmlDocument.GetElementsByTagName("SY_REL_ROLE_USER").Count - 1 Then

                                        ' 新增一條數據
                                        Dim row As DataRow = dtRoleDetail.NewRow
                                        row("DEPART") = sInDepart
                                        row("ROLENAME") = sInRoleName
                                        row("ROLEID") = sInRoleID
                                        row("BRA_DEPNO") = sInBraDepno

                                        If sUserNames.Length > 0 Then
                                            row("USERNAME") = sUserNames.Substring(0, sUserNames.Length - 1)
                                        End If

                                        dtRoleDetail.Rows.Add(row)
                                    End If
                                Else
                                    ' 新增一條數據
                                    Dim row As DataRow = dtRoleDetail.NewRow
                                    row("DEPART") = sInDepart
                                    row("ROLENAME") = sInRoleName
                                    row("ROLEID") = sInRoleID
                                    row("BRA_DEPNO") = sInBraDepno

                                    If sUserNames.Length > 0 Then
                                        row("USERNAME") = sUserNames.Substring(0, sUserNames.Length - 1)
                                    End If

                                    dtRoleDetail.Rows.Add(row)

                                    sDeparType = BRANAME
                                    sRoleType = ROLEID
                                    sInRoleID = ROLEID
                                    sInDepart = BRANAME
                                    sInRoleName = ROLENAME
                                    sInBraDepno = BRA_DEPNO
                                    sUserNames = ""

                                    ' 如果是新增的數據
                                    If OPERATION = "I" Then
                                        sUserNames = sUserNames + "(選入) " + USERNAME + ","
                                    ElseIf OPERATION = "D" Then
                                        sUserNames = sUserNames + "(選出) " + USERNAME + ","
                                    End If

                                    ' 如果是最後一條數據
                                    If i = xmlDocument.GetElementsByTagName("SY_REL_ROLE_USER").Count - 1 Then
                                        Dim rowData As DataRow = dtRoleDetail.NewRow
                                        rowData("DEPART") = sInDepart
                                        rowData("ROLENAME") = sInRoleName
                                        rowData("ROLEID") = sInRoleID
                                        rowData("BRA_DEPNO") = sInBraDepno

                                        If sUserNames.Length > 0 Then
                                            rowData("USERNAME") = sUserNames.Substring(0, sUserNames.Length - 1)
                                        End If

                                        dtRoleDetail.Rows.Add(rowData)
                                    End If
                                End If
                            End If
                        End If
                    Next
                ElseIf xmlDocument.GetElementsByTagName("SY_REL_ROLE_USER").Count = 1 Then
                    ' 取出第一筆資料
                    Dim ROLEID As String = DirectCast(xmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(0)("ROLEID"), System.Xml.XmlElement).InnerText
                    Dim STAFFID As String = DirectCast(xmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(0)("STAFFID"), System.Xml.XmlElement).InnerText
                    Dim BRA_DEPNO As String = DirectCast(xmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(0)("BRA_DEPNO"), System.Xml.XmlElement).InnerText
                    Dim OPERATION As String = DirectCast(xmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(0)("OPERATION"), System.Xml.XmlElement).InnerText
                    Dim TYPE As String = DirectCast(xmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(0)("TYPE"), System.Xml.XmlElement).InnerText
                    Dim BRANAME As String = DirectCast(xmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(0)("BRANAME"), System.Xml.XmlElement).InnerText
                    Dim USERNAME As String = DirectCast(xmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(0)("USERNAME"), System.Xml.XmlElement).InnerText
                    Dim ROLENAME As String = DirectCast(xmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(0)("ROLENAME"), System.Xml.XmlElement).InnerText

                    If sCheckType = TYPE Then

                        If sCheckType = "1" Then
                            ' 員工編號
                            sInStaffid = STAFFID

                            ' 角色編號
                            sInRoleID = ROLEID
                            sInDepart = BRANAME
                            sInUsername = USERNAME
                            sInBRA_DEPNO = BRA_DEPNO

                            ' 如果是新增的數據
                            If OPERATION = "I" Then
                                sRoleNames = sRoleNames + "(選入) " + ROLENAME + ","
                            ElseIf OPERATION = "D" Then
                                sRoleNames = sRoleNames + "(選出) " + ROLENAME + ","
                            End If

                            ' 定義一個新的數據行
                            Dim rowData As DataRow = dtPeopleDetail.NewRow
                            rowData("DEPART") = sInDepart
                            rowData("STAFFID") = sInStaffid
                            rowData("USERNAME") = sInUsername
                            rowData("BRA_DEPNO") = sInBRA_DEPNO

                            If sRoleNames.Length > 0 Then
                                rowData("ROLENAME") = sRoleNames.Substring(0, sRoleNames.Length - 1)
                            End If

                            dtPeopleDetail.Rows.Add(rowData)
                        Else
                            sInRoleID = ROLEID
                            sInDepart = BRANAME
                            sInRoleName = ROLENAME
                            sInBraDepno = BRA_DEPNO

                            ' 如果是新增的數據
                            If OPERATION = "I" Then
                                sUserNames = sUserNames + "(選入) " + USERNAME + ","
                            ElseIf OPERATION = "D" Then
                                sUserNames = sUserNames + "(選出) " + USERNAME + ","
                            End If

                            Dim row As DataRow = dtRoleDetail.NewRow
                            row("DEPART") = sInDepart
                            row("ROLENAME") = sInRoleName
                            row("ROLEID") = sInRoleID
                            row("BRA_DEPNO") = sInBraDepno

                            If sUserNames.Length > 0 Then
                                row("USERNAME") = sUserNames.Substring(0, sUserNames.Length - 1)
                            End If

                            dtRoleDetail.Rows.Add(row)
                        End If
                    End If
                End If

                ' 依人員
                If sCheckType = "1" Then
                    If dtPeopleDetail.Rows.Count > 0 Then

                        ' 顯示依角色分派清單
                        trByPeopleDetail.Style.Value = "display:block"

                        ' 頁面初始加載
                        If sSelected = "1" Then

                            ddlDepart.SelectedItem.Selected = False

                            ' 單位，人員選中默認值
                            ddlDepart.Items.FindByValue(dtPeopleDetail.Rows(0)("BRA_DEPNO")).Selected = True

                            ' 重新綁定人員下拉選單
                            bindDropDownList()
                            ddlUser.SelectedItem.Selected = False

                            ddlUser.Items.FindByText(dtPeopleDetail.Rows(0)("USERNAME")).Selected = True
                        End If
                    Else
                        trByPeopleDetail.Style.Value = "display:none"
                    End If

                    ' 綁定分派清單
                    If dtPeopleDetail.Rows.Count > 0 Then
                        dgByPeopleDetail.DataSource = dtPeopleDetail
                        dgByPeopleDetail.DataBind()
                    Else
                        lstBoxLeft.Items.Clear()
                        lstBoxRight.Items.Clear()
                    End If
                Else
                    If dtRoleDetail.Rows.Count > 0 Then

                        ' 顯示依角色分派清單
                        trByRoleDetail.Style.Value = "display:block"

                        If sSelected = "1" Then

                            ddlDepart.SelectedItem.Selected = False

                            ' 單位，角色選中默認值
                            If IsNothing(ddlDepart.Items.FindByValue(dtRoleDetail.Rows(0)("BRA_DEPNO"))) = False Then
                                ddlDepart.Items.FindByValue(dtRoleDetail.Rows(0)("BRA_DEPNO")).Selected = True
                            End If

                            ' 重新綁定人員下拉選單
                            bindDropDownList()
                            ddlRole.SelectedItem.Selected = False

                            If IsNothing(ddlRole.Items.FindByValue(dtRoleDetail.Rows(0)("ROLEID"))) = False Then
                                ddlRole.Items.FindByValue(dtRoleDetail.Rows(0)("ROLEID")).Selected = True
                            End If

                        End If
                    Else
                        trByRoleDetail.Style.Value = "display:none"
                    End If

                    ' 若有數據
                    If dtRoleDetail.Rows.Count > 0 Then

                        ' 綁定分派清單
                        dgByRoleDetail.DataSource = dtRoleDetail
                        dgByRoleDetail.DataBind()
                    Else
                        lstBoxLeft.Items.Clear()
                        lstBoxRight.Items.Clear()
                    End If
                End If
            Else
                ' 依人員
                If rdolistSetStyle.SelectedValue = "1" Then
                    trByPeopleDetail.Style.Value = "display:none"
                Else
                    trByRoleDetail.Style.Value = "display:none"
                End If

                If m_sHoFlag <> "1" Then
                    ddlDepart.SelectedIndex = 1

                    ' 重新绑定角色或人员
                    bindDropDownList()
                End If
            End If
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 查詢出沒有異動的數據
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getOLstBoxRright() As String
        Try
            Dim dtFlow As DataTable = ViewState("dataFlow")
            Dim sbXmlData As New StringBuilder

            ' 若已選擇清單中沒有資料
            If IsNothing(dtFlow) Then
                sbXmlData.Append("")
            Else
                ' 檢核是否有沒有異動的已選擇的資料
                If dtFlow.Rows.Count > 0 AndAlso lstBoxRight.Items.Count > 0 Then

                    ' 如果設定方式為“依人員”
                    If rdolistSetStyle.SelectedValue = "1" Then

                        ' 循環找出沒有異動的數據
                        For Each item As ListItem In lstBoxRight.Items
                            For i As Integer = 0 To dtFlow.Rows.Count - 1
                                If item.Value = dtFlow.Rows(i)("LValue").ToString() Then
                                    sbXmlData.Append("<SY_REL_ROLE_USER>")
                                    sbXmlData.Append("<ROLEID>" & dtFlow.Rows(i)("LValue").ToString() & "</ROLEID>")
                                    sbXmlData.Append("<STAFFID>" & ddlUser.SelectedValue.Split(";")(0) & "</STAFFID>")
                                    sbXmlData.Append("<BRA_DEPNO>" & ddlDepart.SelectedItem.Value & "</BRA_DEPNO>")
                                    sbXmlData.Append("<OPERATION>N</OPERATION>")
                                    sbXmlData.Append("<TYPE>1</TYPE>")
                                    sbXmlData.Append("<BRANAME>" & ddlDepart.SelectedItem.Text.Trim() & "</BRANAME>")
                                    sbXmlData.Append("<USERNAME>" & ddlUser.SelectedItem.Text & "</USERNAME>")
                                    sbXmlData.Append("<ROLENAME>" & dtFlow.Rows(i)("LText") & "</ROLENAME>")
                                    sbXmlData.Append("</SY_REL_ROLE_USER>")
                                End If
                            Next
                        Next
                    Else

                        ' 循環找出沒有異動的數據
                        For Each item As ListItem In lstBoxRight.Items
                            For i As Integer = 0 To dtFlow.Rows.Count - 1
                                If item.Value = dtFlow.Rows(i)("LValue") Then
                                    sbXmlData.Append("<SY_REL_ROLE_USER>")
                                    sbXmlData.Append("<ROLEID>" & ddlRole.SelectedItem.Value & "</ROLEID>")
                                    sbXmlData.Append("<STAFFID>" & dtFlow.Rows(i)("LValue").ToString().Split(";")(0) & "</STAFFID>")
                                    sbXmlData.Append("<BRA_DEPNO>" & ddlDepart.SelectedItem.Value & "</BRA_DEPNO>")
                                    sbXmlData.Append("<OPERATION>N</OPERATION>")
                                    sbXmlData.Append("<TYPE>2</TYPE>")
                                    sbXmlData.Append("<BRANAME>" & ddlDepart.SelectedItem.Text.Trim() & "</BRANAME>")
                                    sbXmlData.Append("<USERNAME>" & dtFlow.Rows(i)("LText").ToString() & "</USERNAME>")
                                    sbXmlData.Append("<ROLENAME>" & ddlRole.SelectedItem.Text & "</ROLENAME>")
                                    sbXmlData.Append("</SY_REL_ROLE_USER>")
                                End If
                            Next
                        Next
                    End If
                End If

                ' 返回沒有異動的數據
                Return sbXmlData.ToString()
            End If
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Function

    ''' <summary>
    ''' 取得員工編號集合或角色編號集合
    ''' </summary>
    ''' <param name="sType">設定方式</param>
    ''' <param name="sCheck">比較角色編號，或人員編號</param>
    ''' <param name="sDepart">單位</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function getRoleIDsOrStaffids(ByVal sType As String, ByVal sCheck As String, ByVal sDepart As String) As String

        'Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager())
        Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager(), IIf(String.IsNullOrEmpty(m_sCaseId), "SY_TEMPINFO", "SY_TEMPINFO2"))

        Dim oldXmlData As String = String.Empty
        Dim XmlDocument As New XmlDocument
        Dim sStaffids As String = String.Empty ' 員工編號
        Dim sRoleids As String = String.Empty ' 角色編號
        Dim sValue As String = String.Empty


        ' 沒有異動數據的集合 
        Dim dtFlow As DataTable = ViewState("dataFlow")

        If dtFlow.Columns.Count = 0 Then
            dtFlow.Columns.Add("LValue", Type.GetType("System.String"))
            dtFlow.Columns.Add("LText", Type.GetType("System.String"))
        End If

        ' 取得臨時檔資料
        oldXmlData = getXmlData(syTempInfo)

        If oldXmlData <> "" Then

            ' document對象載入XML文件
            XmlDocument.LoadXml(oldXmlData)

            For i As Integer = 0 To XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER").Count - 1
                Dim sRID As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("ROLEID"), System.Xml.XmlElement).InnerText
                Dim sSTAFFID As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("STAFFID"), System.Xml.XmlElement).InnerText
                Dim sBRA_DEPNO As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("BRA_DEPNO"), System.Xml.XmlElement).InnerText
                Dim sOPERATION As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("OPERATION"), System.Xml.XmlElement).InnerText
                Dim TYPE As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("TYPE"), System.Xml.XmlElement).InnerText
                Dim sBRANAME As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("BRANAME"), System.Xml.XmlElement).InnerText
                Dim sUSERNAME As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("USERNAME"), System.Xml.XmlElement).InnerText
                Dim sROLENAME As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("ROLENAME"), System.Xml.XmlElement).InnerText

                If sType = "3" AndAlso TYPE = sCheck Then

                    ' 如果不是刪除掉的數據
                    If sOPERATION <> "D" Then
                        ' 依人員
                        If TYPE = "1" Then
                            If sSTAFFID + ";" + sBRA_DEPNO = ddlUser.SelectedItem.Value AndAlso ddlDepart.SelectedItem.Text.Trim() = sBRANAME Then
                                lstBoxRight.Items.Add(New ListItem(sROLENAME, sRID))

                                sValue = sValue & "'" & sRID & "',"
                            End If
                        Else
                            If sRID = ddlRole.SelectedItem.Value.Split(";")(0) AndAlso ddlDepart.SelectedItem.Text.Trim() = sBRANAME Then
                                lstBoxRight.Items.Add(New ListItem(sUSERNAME, sSTAFFID & ";" & sBRA_DEPNO))

                                sValue = sValue & "'" & sSTAFFID & "',"
                            End If
                        End If
                    End If
                Else

                    ' 依角色
                    If TYPE = "2" Then
                        If sRID = sCheck AndAlso sDepart = sBRANAME Then
                            If sOPERATION = "I" Or sOPERATION = "N" Then
                                lstBoxRight.Items.Add(New ListItem(sUSERNAME, sSTAFFID & ";" & sBRA_DEPNO))

                                ' 如果是沒有異動的數據
                                If sOPERATION = "N" Then
                                    Dim row As DataRow = dtFlow.NewRow
                                    row("LValue") = sSTAFFID & ";" & sBRA_DEPNO
                                    row("LText") = sUSERNAME

                                    dtFlow.Rows.Add(row)
                                End If

                                sStaffids = sStaffids & "'" & sSTAFFID & "',"
                            End If
                        End If
                    Else
                        If sSTAFFID = sCheck AndAlso sDepart = sBRANAME Then

                            If sOPERATION = "I" Or sOPERATION = "N" Then
                                lstBoxRight.Items.Add(New ListItem(sROLENAME, sRID))

                                ' 如果是沒有異動的數據
                                If sOPERATION = "N" Then
                                    Dim row As DataRow = dtFlow.NewRow
                                    row("LValue") = sRID
                                    row("LText") = sROLENAME

                                    dtFlow.Rows.Add(row)
                                End If

                                sRoleids = sRoleids & "'" & sRID & "',"
                            End If
                        End If
                    End If
                End If
            Next
        End If

        ' 依角色
        If sType = "1" Then
            Return sRoleids
        ElseIf sType = "2" Then
            Return sStaffids
        ElseIf sType = "3" Then
            Return sValue
        End If
    End Function

    ''' <summary>
    ''' 選入選出時存儲Xml到臨時檔(去除沒有異動過的數據)
    ''' </summary>
    ''' <param name="sInOrOut">選入還是選出</param>
    ''' <param name="dtMove">移動過的數據集合</param>
    ''' <remarks></remarks>
    Sub editInOutXml(ByVal sInOrOut As String, ByVal dtMove As DataTable)

        ' 聲明參數
        'Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager())
        Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager(), IIf(String.IsNullOrEmpty(m_sCaseId), "SY_TEMPINFO", "SY_TEMPINFO2"))

        Dim sXmlData As String = String.Empty
        Dim XmlDocument As New XmlDocument
        Dim item As XmlNode
        Dim iIndexDtToLeft As String = String.Empty
        Dim nodeList As IList(Of XmlNode) = New List(Of XmlNode)() ' Xml節點的集合

        If Not IsNothing(ViewState("inOrOutXml")) AndAlso ViewState("inOrOutXml").ToString() <> "" Then
            sXmlData = ViewState("inOrOutXml").ToString()
        Else
            ' 取得臨時檔資料
            sXmlData = getXmlData(syTempInfo)
        End If

        If sXmlData <> "" Then

            ' document對象載入XML文件
            XmlDocument.LoadXml(sXmlData)

            ' 如果有相關人員信息
            For i As Integer = 0 To XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER").Count - 1
                Dim sRID As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("ROLEID"), System.Xml.XmlElement).InnerText
                Dim STAFFID As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("STAFFID"), System.Xml.XmlElement).InnerText
                Dim sBRA_DEPNO As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("BRA_DEPNO"), System.Xml.XmlElement).InnerText
                Dim sOPERATION As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("OPERATION"), System.Xml.XmlElement).InnerText
                Dim TYPE As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("TYPE"), System.Xml.XmlElement).InnerText
                Dim sBRANAME As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("BRANAME"), System.Xml.XmlElement).InnerText
                Dim sUSERNAME As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("USERNAME"), System.Xml.XmlElement).InnerText
                Dim sROLENAME As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("ROLENAME"), System.Xml.XmlElement).InnerText

                ' 如果設定方式與 當前方式相同
                If TYPE = rdolistSetStyle.SelectedValue Then

                    ' 如果是選出
                    If sInOrOut = "out" Then
                        ' 依角色
                        If TYPE = "1" Then
                            For j As Integer = 0 To dtMove.Rows.Count - 1
                                If dtMove.Rows(j)("VALUE") = sRID AndAlso
                                    sOPERATION = "I" AndAlso
                                    ddlDepart.SelectedItem.Text.Trim() = sBRANAME AndAlso
                                    ddlUser.SelectedItem.Value = STAFFID & ";" & sBRA_DEPNO Then
                                    iIndexDtToLeft = iIndexDtToLeft & j & ","
                                    item = XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)
                                    nodeList.Add(item)
                                End If
                            Next
                        Else
                            For j As Integer = 0 To dtMove.Rows.Count - 1
                                If dtMove.Rows(j)("VALUE") = STAFFID & ";" & sBRA_DEPNO AndAlso
                                    sOPERATION = "I" AndAlso
                                    ddlDepart.SelectedItem.Text.Trim() = sBRANAME AndAlso
                                    ddlRole.SelectedItem.Value = sRID Then
                                    iIndexDtToLeft = iIndexDtToLeft & j & ","
                                    item = XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)
                                    nodeList.Add(item)
                                End If
                            Next
                        End If
                    Else
                        ' 依角色
                        If TYPE = "1" Then
                            For j As Integer = 0 To dtMove.Rows.Count - 1
                                If dtMove.Rows(j)("VALUE") = sRID AndAlso
                                    sOPERATION = "D" AndAlso
                                    ddlDepart.SelectedItem.Text.Trim() = sBRANAME AndAlso
                                    ddlUser.SelectedItem.Value = STAFFID & ";" & sBRA_DEPNO Then
                                    iIndexDtToLeft = iIndexDtToLeft & j & ","
                                    item = XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)
                                    nodeList.Add(item)
                                End If
                            Next
                        Else
                            For j As Integer = 0 To dtMove.Rows.Count - 1
                                If dtMove.Rows(j)("VALUE") = STAFFID & ";" & sBRA_DEPNO AndAlso
                                    sOPERATION = "D" AndAlso
                                    ddlDepart.SelectedItem.Text.Trim() = sBRANAME AndAlso
                                    ddlRole.SelectedItem.Value = sRID Then
                                    iIndexDtToLeft = iIndexDtToLeft & j & ","
                                    item = XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)
                                    nodeList.Add(item)
                                End If
                            Next
                        End If
                    End If
                End If
            Next

            ' 如果有 則將其替換掉 然後累加上新的XML
            If nodeList.Count > 0 Then
                For i As Integer = 0 To nodeList.Count - 1
                    If Not nodeList.Item(i) Is Nothing Then
                        Dim oldXML As String = nodeList.Item(i).InnerXml.ToString()

                        ' 若選出
                        If sInOrOut = "out" Then
                            sXmlData = sXmlData.Replace(oldXML.ToString(), "").Replace("<SY_REL_ROLE_USER></SY_REL_ROLE_USER>", "")
                        Else
                            sXmlData = sXmlData.Replace(oldXML.ToString(), oldXML.ToString().Replace("<OPERATION>D</OPERATION>", "<OPERATION>N</OPERATION>"))
                        End If
                    End If
                Next

                ViewState("inOrOutXml") = sXmlData

                Dim iDtToleft As Integer = iIndexDtToLeft.LastIndexOf(",")
                iIndexDtToLeft = iIndexDtToLeft.Substring(0, iDtToleft)

                ' 循環刪除沒有異動的資料
                For i As Integer = iIndexDtToLeft.Split(",").Length - 1 To 0 Step -1
                    Dim j As Integer = Convert.ToInt32(iIndexDtToLeft.Split(",")(i))
                    dtMove.Rows(j).Delete()
                Next
            End If
        End If
    End Sub

    ''' <summary>
    ''' 頁面加載時ListBox中移動過的資料變紅
    ''' </summary>
    ''' <remarks></remarks>
    Sub editListItemColor()
        'Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager())
        Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager(), IIf(String.IsNullOrEmpty(m_sCaseId), "SY_TEMPINFO", "SY_TEMPINFO2"))

        Dim sXmlData As String = String.Empty
        Dim XmlDocument As New XmlDocument

        If Not IsNothing(ViewState("inOrOutXml")) AndAlso ViewState("inOrOutXml").ToString() <> "" Then
            sXmlData = ViewState("inOrOutXml").ToString()
        Else
            ' 取得臨時檔資料
            sXmlData = getXmlData(syTempInfo)
        End If

        ' 若有臨時數據
        If sXmlData <> "" Then
            ' document對象載入XML文件
            XmlDocument.LoadXml(sXmlData)

            ' 如果有相關人員信息
            For i As Integer = 0 To XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER").Count - 1
                Dim sRID As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("ROLEID"), System.Xml.XmlElement).InnerText
                Dim STAFFID As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("STAFFID"), System.Xml.XmlElement).InnerText
                Dim sBRA_DEPNO As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("BRA_DEPNO"), System.Xml.XmlElement).InnerText
                Dim sOPERATION As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("OPERATION"), System.Xml.XmlElement).InnerText
                Dim TYPE As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("TYPE"), System.Xml.XmlElement).InnerText
                Dim sBRANAME As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("BRANAME"), System.Xml.XmlElement).InnerText
                Dim sUSERNAME As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("USERNAME"), System.Xml.XmlElement).InnerText
                Dim sROLENAME As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("ROLENAME"), System.Xml.XmlElement).InnerText

                ' 設定方式與當前角色相同
                If rdolistSetStyle.SelectedValue = TYPE Then

                    ' 依人員
                    If TYPE = "1" Then

                        ' 如果是新增的資料
                        If sOPERATION = "I" Then
                            If STAFFID + ";" + sBRA_DEPNO = ddlUser.SelectedItem.Value AndAlso
                                ddlDepart.SelectedItem.Text.Trim() = sBRANAME Then
                                Dim item As ListItem = lstBoxRight.Items.FindByValue(sRID)
                                item.Attributes.Add("style", "color:red")
                            End If
                        End If

                        ' 如果是刪除的資料
                        If sOPERATION = "D" Then
                            If STAFFID + ";" + sBRA_DEPNO = ddlUser.SelectedItem.Value AndAlso
                                ddlDepart.SelectedItem.Text.Trim() = sBRANAME Then
                                Dim item As ListItem = lstBoxLeft.Items.FindByValue(sRID)
                                item.Attributes.Add("style", "color:red")
                            End If
                        End If
                    Else
                        ' 如果是新增的資料
                        If sOPERATION = "I" Then
                            If sRID = ddlRole.SelectedItem.Value.Split(";")(0) AndAlso
                                ddlDepart.SelectedItem.Text.Trim() = sBRANAME Then
                                Dim item As ListItem = lstBoxRight.Items.FindByValue(STAFFID & ";" & sBRA_DEPNO)
                                item.Attributes.Add("style", "color:red")
                            End If
                        End If

                        ' 如果是刪除的資料
                        If sOPERATION = "D" Then
                            If sRID = ddlRole.SelectedItem.Value.Split(";")(0) Andalso ddlDepart.SelectedItem.Text.Trim() = sBRANAME Then
                                Dim item As ListItem = lstBoxLeft.Items.FindByValue(STAFFID & ";" & sBRA_DEPNO)
                                If IsNothing(item) = False Then
                                    item.Attributes.Add("style", "color:red")
                                End If
                            End If
                        End If
                    End If
                End If
            Next
        End If
    End Sub

    ''' <summary>
    ''' ListBox里的數據左右異動時變紅
    ''' </summary>
    ''' <remarks></remarks>
    Sub editMoweItemColor()

        ' 定義數據表接收移動的數據
        Dim dtListBoxMove As DataTable = ViewState("dtListBoxMove")

        If Not IsNothing(dtListBoxMove) Then
            For i As Integer = 0 To dtListBoxMove.Rows.Count - 1

                ' 如果是左移的數據
                If dtListBoxMove.Rows(i)("OPERATION") = "D" Then
                    Dim itemDelete As ListItem = lstBoxLeft.Items.FindByValue(dtListBoxMove.Rows(i)("VALUE"))
                    itemDelete.Attributes.Add("style", "color:red")
                ElseIf dtListBoxMove.Rows(i)("OPERATION") = "I" Then
                    Dim item As ListItem = lstBoxRight.Items.FindByValue(dtListBoxMove.Rows(i)("VALUE"))
                    item.Attributes.Add("style", "color:red")
                End If
            Next
        End If
    End Sub

    ''' <summary>
    ''' 新增資料到歷史檔
    ''' </summary>
    ''' <param name="sApproved">是否同意</param>
    ''' <param name="sCheck">是否點擊同意按鈕</param>
    ''' <param name="sFlag">是否走流程方法</param>
    ''' <remarks></remarks>
    Public Sub flowFuntion(ByVal sApproved As String, ByVal sCheck As String, ByRef stepinfo As StepInfo)
        ' 實例化
        'Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager())
        Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager(), IIf(String.IsNullOrEmpty(m_sCaseId), "SY_TEMPINFO", "SY_TEMPINFO2"))

        Dim syRelRoleUser As New AUTH_OP.SY_REL_ROLE_USER(GetDatabaseManager())
        Dim XmlDocument As New XmlDocument

        Dim oldXmlData As String = String.Empty ' 聲明參數接收查詢出的XML
        Dim sRoleID As String = String.Empty  ' 角色編號
        Dim sUserID As String = String.Empty ' 用戶編號
        Dim sBraDepno As String = String.Empty  ' 部門編號

        ' 取得臨時檔資料
        oldXmlData = getXmlData(syTempInfo)

        ' 若有臨時檔資料
        If oldXmlData <> "" Then

            ' document對象載入XML文件
            XmlDocument.LoadXml(oldXmlData)

            For i As Integer = 0 To XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER").Count - 1
                Dim sRID As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("ROLEID"), System.Xml.XmlElement).InnerText
                Dim sSTAFFID As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("STAFFID"), System.Xml.XmlElement).InnerText
                Dim sBRA_DEPNO As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("BRA_DEPNO"), System.Xml.XmlElement).InnerText
                Dim sOPERATION As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("OPERATION"), System.Xml.XmlElement).InnerText
                Dim TYPE As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("TYPE"), System.Xml.XmlElement).InnerText
                Dim sBRANAME As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("BRANAME"), System.Xml.XmlElement).InnerText
                Dim sUSERNAME As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("USERNAME"), System.Xml.XmlElement).InnerText
                Dim sROLENAME As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("ROLENAME"), System.Xml.XmlElement).InnerText

                ' 同意
                If sCheck = "agree" Then

                    ' 依人員
                    If TYPE = "1" Then
                        syRelRoleUser.deleteBySTAFFIDDepNO(sSTAFFID, sBRA_DEPNO, sRID)
                    Else
                        If sRoleID <> sRID Or sBraDepno <> sBRA_DEPNO Then
                            sBraDepno = sBRA_DEPNO
                            sRoleID = sRID
                            syRelRoleUser.deleteByRoleIDDepNO(sRID, sBRA_DEPNO)
                        End If
                    End If

                    ' 新增數據到實際表
                    If sOPERATION = "I" Or sOPERATION = "N" Then

                        ' 如果SY_REL_ROLE_USER裱中沒有數據
                        If Not syRelRoleUser.loadByPK(sSTAFFID, sRID, sBRA_DEPNO) Then
                            syRelRoleUser.setAttribute("STAFFID", sSTAFFID)
                            syRelRoleUser.setAttribute("ROLEID", sRID)
                            syRelRoleUser.setAttribute("BRA_DEPNO", sBRA_DEPNO)
                        End If

                        syRelRoleUser.save()
                        syRelRoleUser.clear()
                    End If
                End If

                ' 新增SY_REL_ROLE_USER_HIS數據
                insertsyRelRoleUserHis(GetDatabaseManager(), sRID, sOPERATION, sSTAFFID, sBRA_DEPNO, sApproved _
                                       , stepinfo.currentStepInfo.caseId _
                                       , stepinfo.currentStepInfo.stepNo _
                                       , stepinfo.currentStepInfo.subflowSeq _
                                       , stepinfo.currentStepInfo.subflowCount)
            Next
        End If
    End Sub

    ''' <summary>
    ''' 將數據增加到歷史檔
    ''' </summary>
    ''' <param name="databaseManager">DB鏈接對象</param>
    ''' <param name="sRoleID">角色編號</param>
    ''' <param name="sOperation"></param>
    ''' <param name="sStaffid">員工編號</param>
    ''' <param name="sBraDepno">部門編號</param>
    ''' <param name="sApproved">是否同意</param>
    ''' <param name="sFlag">是否走流程</param>
    ''' <param name="sCaseId">案件編號</param>
    ''' <param name="sStepNo">步驟號</param>
    ''' <param name="iSubFlowSeq"></param>
    ''' <param name="iSubFlowCount"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function insertsyRelRoleUserHis(ByVal databaseManager As DatabaseManager, ByVal sRoleID As String, ByVal sOperation As String, ByVal sStaffid As String _
                  , ByVal sBraDepno As String, ByVal sApproved As String, Optional ByVal sCaseId As String = "SY0000000000000" _
                  , Optional ByVal sStepNo As String = "00000000", Optional ByVal iSubFlowSeq As Integer = 0 _
                  , Optional ByVal iSubFlowCount As Integer = 0)

        Dim syRelRoleUserHis As New AUTH_OP.SY_REL_ROLE_USER_HIS(databaseManager)

        ' 如果SY_REL_ROLE_USER_HIS裱中沒有數據
        syRelRoleUserHis.laodByPK(sRoleID, sCaseId, sStepNo, iSubFlowSeq, iSubFlowCount, sStaffid, sBraDepno)

        syRelRoleUserHis.setAttribute("ROLEID", sRoleID)
        syRelRoleUserHis.setAttribute("CASEID", sCaseId)
        syRelRoleUserHis.setAttribute("STEP_NO", sStepNo)
        syRelRoleUserHis.setAttribute("SUBFLOW_SEQ", iSubFlowSeq)
        syRelRoleUserHis.setAttribute("SUBFLOW_COUNT", iSubFlowCount)
        syRelRoleUserHis.setAttribute("STAFFID", sStaffid)
        syRelRoleUserHis.setAttribute("BRA_DEPNO", sBraDepno)

        syRelRoleUserHis.setAttribute("APPROVED", sApproved)
        syRelRoleUserHis.setAttribute("OPERATION", sOperation)

        syRelRoleUserHis.save()
        syRelRoleUserHis.clear()
    End Function
#End Region

#Region "Event"

    ''' <summary>
    ''' 點擊"選入"按鈕事件
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnIn_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnIn.Click
        Try
            ' 定義數據表接收移動的數據
            Dim dtListBoxMove As DataTable = ViewState("dtListBoxMove")

            ' 定義數據表接收向右移動的數據
            Dim dtToRight As New DataTable
            dtToRight.Columns.Add("VALUE", Type.GetType("System.String"))
            dtToRight.Columns.Add("TEXT", Type.GetType("System.String"))

            Dim iCount As Integer = lstBoxLeft.Items.Count
            Dim iIndex As Integer = 0
            Dim sMove As String = String.Empty
            Dim sTRight As String = String.Empty

            ' 循環得到被選中的人員
            For i As Integer = 0 To iCount - 1

                ' 獲得當前人員
                Dim item As ListItem = lstBoxLeft.Items(iIndex)

                ' 如果當前人員被選中
                If item.Selected Then

                    ' 檢核依角色時登錄者不可以選入自己
                    If (rdolistSetStyle.SelectedValue = "2" AndAlso item.Value.Split(";")(0) = m_sLoginUserid) OrElse
                        (rdolistSetStyle.SelectedValue = "1" AndAlso ddlUser.SelectedItem.Value.Split(";")(0) = m_sLoginUserid) Then
                        com.Azion.EloanUtility.UIUtility.alert("權限管理人員不可更改自己的角色！")
                        iIndex = iIndex + 1

                        Continue For
                    End If

                    '不可選入有子角色的角色
                    If rdolistSetStyle.SelectedValue = "1" Then
                        Dim obj As Object
                        obj = BosBase.getNewBosBase("SY_ROLE", GetDatabaseManager).ExecuteScalar(
                                                "select COUNT(*) " & vbCrLf & _
                                                "  from SY_ROLE " & vbCrLf & _
                                                " where ROLEID not in " & vbCrLf & _
                                                "       (select distinct parent from SY_ROLE where PARENT <> 0) " & vbCrLf & _
                                                "   and ROLEID = @ROLEID@ " & vbCrLf,
                                                "ROLEID", item.Value)

                        If CInt(obj) = 0 Then
                            iIndex = iIndex + 1
                            Continue For
                        End If
                    End If

                    lstBoxLeft.Items.Remove(item)
                    lstBoxRight.Items.Add(item)

                    Dim row As DataRow = dtToRight.NewRow
                    row("VALUE") = item.Value
                    row("TEXT") = item.Text
                    dtToRight.Rows.Add(row)

                    iIndex = iIndex - 1
                End If

                iIndex = iIndex + 1
            Next

            ' 取消右邊ListBox的選中狀態
            lstBoxRight.SelectedIndex = -1

            ' 篩選掉沒有異動的數據
            editInOutXml("in", dtToRight)

            ' 尋找并記錄下不是絕對異動的數據
            If dtListBoxMove.Rows.Count > 0 Then
                For i As Integer = 0 To dtListBoxMove.Rows.Count - 1
                    For j As Integer = 0 To dtToRight.Rows.Count - 1
                        If dtListBoxMove.Rows(i)("VALUE") = dtToRight.Rows(j)("VALUE") Then
                            sMove = sMove & i.ToString() & ","
                            sTRight = sTRight & j.ToString() & ","
                        End If
                    Next
                Next
            End If

            ' 去掉不是絕對異動的數據
            If sMove.Length > 0 Then
                Dim iM As Integer = sMove.LastIndexOf(",")
                Dim iTR As Integer = sTRight.LastIndexOf(",")

                sMove = sMove.Substring(0, iM)
                sTRight = sTRight.Substring(0, iTR)

                For i As Integer = sMove.Split(",").Length - 1 To 0 Step -1
                    Dim j As Integer = Convert.ToInt32(sMove.Split(",")(i))
                    dtListBoxMove.Rows(j).Delete()
                Next

                For i As Integer = sTRight.Split(",").Length - 1 To 0 Step -1
                    Dim j As Integer = Convert.ToInt32(sTRight.Split(",")(i))
                    dtToRight.Rows(j).Delete()
                Next
            End If

            ' 存儲絕異動的數據
            If dtToRight.Rows.Count > 0 Then
                For i As Integer = 0 To dtToRight.Rows.Count - 1
                    Dim row As DataRow = dtListBoxMove.NewRow
                    row("VALUE") = dtToRight.Rows(i)("VALUE")
                    row("TEXT") = dtToRight.Rows(i)("TEXT")
                    row("OPERATION") = "I"

                    dtListBoxMove.Rows.Add(row)
                Next
            End If

            ' 臨時檔中異動數據顯示紅色
            editListItemColor()

            ' 移動數據變紅
            editMoweItemColor()

            ViewState("dtListBoxMove") = dtListBoxMove
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 點擊“選出”按鈕事件
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnOut_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnOut.Click
        Try
            ' 聲明參數
            Dim dtListBoxMove As DataTable = ViewState("dtListBoxMove")  ' 左右移動的數據
            Dim dtFlow As DataTable = ViewState("dataFlow")   ' 已選擇的資料
            Dim iCount As Integer = lstBoxRight.Items.Count   ' 已選擇數據總數
            Dim iIndex As Integer = 0
            Dim sMove As String = String.Empty
            Dim sTleft As String = String.Empty ' 左移資料的編號

            ' 向左移動的資歷表
            Dim dtToLeft As New DataTable

            dtToLeft.Columns.Add("VALUE", Type.GetType("System.String"))
            dtToLeft.Columns.Add("TEXT", Type.GetType("System.String"))

            ' 循環得到被選中的數據
            For i As Integer = 0 To iCount - 1

                ' 獲得當前數據
                Dim lstitem As ListItem = lstBoxRight.Items(iIndex)

                ' 如果當前數據被選中
                If lstitem.Selected Then

                    ' 檢核依角色時登錄者不可以選出自己
                    If (rdolistSetStyle.SelectedValue = "2" AndAlso lstitem.Value.Split(";")(0) = m_sLoginUserid) OrElse
                       (rdolistSetStyle.SelectedValue = "1" AndAlso ddlUser.SelectedItem.Value.Split(";")(0) = m_sLoginUserid) Then
                        com.Azion.EloanUtility.UIUtility.alert("權限管理人員不可更改自己的角色！")
                        iIndex = iIndex + 1

                        Continue For
                    End If

                    lstBoxRight.Items.Remove(lstitem)
                    lstBoxLeft.Items.Add(lstitem)

                    Dim row As DataRow = dtToLeft.NewRow
                    row("VALUE") = lstitem.Value
                    row("TEXT") = lstitem.Text
                    dtToLeft.Rows.Add(row)

                    iIndex = iIndex - 1
                End If

                iIndex = iIndex + 1
            Next

            ' 取消左邊ListBox的選中狀態
            lstBoxLeft.SelectedIndex = -1

            ' 存儲絕對移動的數據
            editInOutXml("out", dtToLeft)

            ' 尋找并記錄下不是絕對異動的數據
            If dtListBoxMove.Rows.Count > 0 Then
                For i As Integer = 0 To dtListBoxMove.Rows.Count - 1
                    For j As Integer = 0 To dtToLeft.Rows.Count - 1
                        If dtListBoxMove.Rows(i)("VALUE") = dtToLeft.Rows(j)("VALUE") Then
                            sMove = sMove & i.ToString() & ","
                            sTleft = sTleft & j.ToString() & ","
                        End If
                    Next
                Next
            End If

            ' 去掉不是絕對異動的數據
            If sMove.Length > 0 Then
                Dim iM As Integer = sMove.LastIndexOf(",")
                Dim iTL As Integer = sTleft.LastIndexOf(",")

                sMove = sMove.Substring(0, iM)
                sTleft = sTleft.Substring(0, iTL)

                For i As Integer = sMove.Split(",").Length - 1 To 0 Step -1
                    Dim j As Integer = Convert.ToInt32(sMove.Split(",")(i))
                    dtListBoxMove.Rows(j).Delete()
                Next

                For i As Integer = sTleft.Split(",").Length - 1 To 0 Step -1
                    Dim j As Integer = Convert.ToInt32(sTleft.Split(",")(i))
                    dtToLeft.Rows(j).Delete()
                Next
            End If

            ' 如果已選擇ListBox中有資料
            If Not IsNothing(dtFlow) Then
                If dtToLeft.Rows.Count > 0 And dtFlow.Rows.Count > 0 Then
                    For i As Integer = 0 To dtToLeft.Rows.Count - 1
                        For j As Integer = 0 To dtFlow.Rows.Count - 1
                            If dtFlow.Rows(j)("LValue") = dtToLeft.Rows(i)("VALUE") Then
                                Dim row As DataRow = dtListBoxMove.NewRow
                                row("VALUE") = dtToLeft.Rows(i)("VALUE")
                                row("TEXT") = dtToLeft.Rows(i)("TEXT")
                                row("OPERATION") = "D"

                                dtListBoxMove.Rows.Add(row)
                            End If
                        Next
                    Next
                End If
            End If

            ' 將移動的數據變紅
            editMoweItemColor()

            ' 數據變紅
            editListItemColor()

            ViewState("dtListBoxMove") = dtListBoxMove
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 點擊"存儲"按鈕事件
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnSave_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnSave.Click
        Try
            ' 實例化
            'Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager())
            Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager(), IIf(String.IsNullOrEmpty(m_sCaseId), "SY_TEMPINFO", "SY_TEMPINFO2"))

            Dim sXmlData As String = String.Empty   ' 臨時檔資料
            Dim xmlDocument As New XmlDocument
            Dim sType As String = String.Empty
            Dim item As XmlNode  ' 臨時檔節點資料
            Dim newXmlData As String = String.Empty
            Dim nodeList As IList(Of XmlNode) = New List(Of XmlNode)()

            ' 取得臨時檔資料
            sXmlData = getXmlData(syTempInfo)

            ' 檢核同一批異動設定方式必須相同
            If sXmlData <> "" Then

                Dim nodeListMove As IList(Of XmlNode) = New List(Of XmlNode)()

                xmlDocument.LoadXml(sXmlData)

                If xmlDocument.GetElementsByTagName("SY_REL_ROLE_USER").Count <> 0 Then
                    ' 取得設定方式
                    sType = DirectCast(xmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(0)("TYPE"), System.Xml.XmlElement).InnerText

                    ' 如果同一批異動設定方式不相同
                    If rdolistSetStyle.SelectedValue <> sType Then
                        com.Azion.EloanUtility.UIUtility.alert("設定方式需一致！")

                        initXmlListBox(rdolistSetStyle.SelectedValue)

                        lstBoxLeft.Items.Clear()
                        lstBoxRight.Items.Clear()

                        ddlDepart.SelectedIndex = -1

                        ' 依人員
                        If rdolistSetStyle.SelectedValue = "1" Then

                            ddlUser.Items.Clear()
                            bindDropDownList()
                        Else

                            ddlRole.Items.Clear()
                            bindDropDownList()
                        End If

                        Return
                    End If
                End If
            End If

            If ViewState("inOrOutXml") <> "" Then
                GetDatabaseManager.beginTran()

                'Dim syTempInfoEdit As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager())
                Dim syTempInfoEdit As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager(), IIf(String.IsNullOrEmpty(m_sCaseId), "SY_TEMPINFO", "SY_TEMPINFO2"))

                If ViewState("inOrOutXml") = "<SY></SY>" Then
                    ViewState("inOrOutXml") = ""
                End If

                ' 若案件已在流程中
                If m_sFourStepNo = "" Then

                    ' 存儲修改過的Xml資料
                    If Not syTempInfoEdit.loadByPK(m_sWorkingUserid, m_sWorkingTopDepNo, m_sFuncCode) Then

                        ' 新增異動數據到臨時表
                        syTempInfoEdit.setAttribute("STAFFID", m_sWorkingUserid)
                        syTempInfoEdit.setAttribute("BRA_DEPNO", m_sWorkingTopDepNo)
                        syTempInfoEdit.setAttribute("FUNCCODE", m_sFuncCode)
                    End If

                    If ViewState("inOrOutXml").ToString() = "" Then
                        syTempInfoEdit.remove()
                    Else
                        syTempInfoEdit.setAttribute("TEMPDATA", ViewState("inOrOutXml").ToString())

                        If String.IsNullOrEmpty(m_sCaseId) = False Then
                            syTempInfo.setAttribute("CASEID", m_sCaseId)
                        Else
                            syTempInfo.setAttribute("CASEID", DBNull.Value)
                        End If

                        syTempInfoEdit.save()
                        syTempInfoEdit.clear()
                    End If
                Else

                    ' 存儲修改過的Xml資料
                    If Not syTempInfoEdit.loadByCaseId(m_sCaseId) Then
                        syTempInfoEdit.setAttribute("CASEID", m_sCaseId)
                    End If

                    If ViewState("inOrOutXml").ToString() = "" Then
                        syTempInfoEdit.remove()
                    Else
                        syTempInfoEdit.setAttribute("TEMPDATA", ViewState("inOrOutXml").ToString())

                        syTempInfoEdit.save()
                        syTempInfoEdit.clear()
                    End If

                    GetDatabaseManager.commit()

                    ViewState("inOrOutXml") = ""
                End If
            End If

            GetDatabaseManager.beginTran()

            ' 如果“設定方式”為依人員
            If rdolistSetStyle.SelectedValue = "1" Then

                ' 獲得絕對異動的數據
                Dim dtListBoxMove As DataTable = ViewState("dtListBoxMove")
                Dim sbXmlData As New StringBuilder
                Dim oldXmlData As String = String.Empty

                ' 如果有絕對異動的數據
                If dtListBoxMove.Rows.Count > 0 Then

                    ' 組XML文件
                    sbXmlData.Append("<SY>")
                    For i As Integer = 0 To dtListBoxMove.Rows.Count - 1
                        sbXmlData.Append("<SY_REL_ROLE_USER>")
                        sbXmlData.Append("<ROLEID>" + dtListBoxMove.Rows(i)("VALUE") + "</ROLEID>")
                        sbXmlData.Append("<STAFFID>" + ddlUser.SelectedValue.Split(";")(0) + "</STAFFID>")
                        sbXmlData.Append("<BRA_DEPNO>" + ddlDepart.SelectedItem.Value + "</BRA_DEPNO>")
                        sbXmlData.Append("<OPERATION>" + dtListBoxMove.Rows(i)("OPERATION") + "</OPERATION>")
                        sbXmlData.Append("<TYPE>1</TYPE>")
                        sbXmlData.Append("<BRANAME>" + ddlDepart.SelectedItem.Text.Trim() + "</BRANAME>")
                        sbXmlData.Append("<USERNAME>" + ddlUser.SelectedItem.Text + "</USERNAME>")
                        sbXmlData.Append("<ROLENAME>" + dtListBoxMove.Rows(i)("TEXT") + "</ROLENAME>")
                        sbXmlData.Append("</SY_REL_ROLE_USER>")
                    Next
                    sbXmlData.Append(getOLstBoxRright())
                    sbXmlData.Append("</SY>")

                    ' 取得臨時檔資料
                    oldXmlData = getXmlData(syTempInfo)

                    If oldXmlData <> "" Then

                        ' document對象載入XML文件
                        xmlDocument.LoadXml(oldXmlData)

                        ' 找到臨時檔中已經存在的資料
                        For i As Integer = 0 To xmlDocument.GetElementsByTagName("SY_REL_ROLE_USER").Count - 1
                            If DirectCast(xmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("STAFFID"), System.Xml.XmlElement).InnerText = ddlUser.SelectedItem.Value.Split(";")(0) _
                                      And DirectCast(xmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("BRANAME"), System.Xml.XmlElement).InnerText = ddlDepart.SelectedItem.Text.Trim() Then
                                item = xmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)
                                nodeList.Add(item)
                            End If
                        Next

                        ' 如果有 則將其替換掉 然後累加上新的XML
                        If nodeList.Count > 0 Then
                            Dim newXML As String = oldXmlData.ToString()
                            Dim oldXML As String = String.Empty

                            ' 循環取得相同的欄位
                            For i As Integer = 0 To nodeList.Count - 1
                                If Not nodeList.Item(i) Is Nothing Then
                                    Dim oldItemXML As String = nodeList.Item(i).InnerXml.ToString()

                                    ' 若是沒有異動過的資料
                                    If Not oldItemXML.Contains("<OPERATION>N</OPERATION>") Then
                                        oldXML = oldXML & "<SY_REL_ROLE_USER>" & oldItemXML.ToString() & "</SY_REL_ROLE_USER>"
                                    End If

                                    newXML = newXML.Replace(oldItemXML.ToString(), "").Replace("<SY_REL_ROLE_USER></SY_REL_ROLE_USER>", "")
                                End If
                            Next
                            newXmlData = newXML.ToString().Replace("</SY>", "") & oldXML & sbXmlData.Replace("<SY>", "").ToString()
                        Else

                            ' 如果找不到，直接進行累加
                            newXmlData = oldXmlData.ToString().Replace("</SY>", "") & sbXmlData.Replace("<SY>", "").ToString()
                        End If
                    Else

                        ' 新增資料
                        newXmlData = sbXmlData.ToString()
                    End If

                    If newXmlData = "<SY></SY>" Then
                        newXmlData = ""
                    End If

                    ' 如果案件已在流程中
                    If m_sFourStepNo = "" Then

                        ' 修改臨時檔資料
                        If Not syTempInfo.loadByPK(m_sWorkingUserid, m_sWorkingTopDepNo, m_sFuncCode) Then

                            ' 新增異動數據到臨時表
                            syTempInfo.setAttribute("STAFFID", m_sWorkingUserid)
                            syTempInfo.setAttribute("BRA_DEPNO", m_sWorkingTopDepNo)
                            syTempInfo.setAttribute("FUNCCODE", m_sFuncCode)
                        End If

                        If newXmlData = "" Then
                            syTempInfo.remove()
                        Else
                            syTempInfo.setAttribute("TEMPDATA", newXmlData.ToString())
                            syTempInfo.save()
                        End If
                    Else

                        '修噶臨時檔資料
                        If syTempInfo.loadByCaseId(m_sCaseId) Then
                            syTempInfo.setAttribute("CASEID", m_sCaseId)
                        End If

                        If newXmlData = "" Then
                            syTempInfo.remove()
                        Else
                            syTempInfo.setAttribute("TEMPDATA", newXmlData.ToString())
                            syTempInfo.save()
                        End If
                    End If

                    com.Azion.EloanUtility.UIUtility.alert("存儲成功！")
                End If

                dtListBoxMove.Rows.Clear()

                ViewState("dtListBoxMove") = dtListBoxMove
            Else
                ' 獲得絕對異動的數據
                Dim dtListBoxMove As DataTable = ViewState("dtListBoxMove")
                Dim sbXmlData As New StringBuilder
                Dim oldXmlData As String = String.Empty

                ' 如果有絕對異動的數據
                If dtListBoxMove.Rows.Count > 0 Then

                    ' 組XML文件
                    sbXmlData.Append("<SY>")
                    For i As Integer = 0 To dtListBoxMove.Rows.Count - 1
                        sbXmlData.Append("<SY_REL_ROLE_USER>")
                        sbXmlData.Append("<ROLEID>" + ddlRole.SelectedItem.Value + "</ROLEID>")
                        sbXmlData.Append("<STAFFID>" + dtListBoxMove.Rows(i)("VALUE").ToString().Split(";")(0) + "</STAFFID>")
                        sbXmlData.Append("<BRA_DEPNO>" + ddlDepart.SelectedItem.Value + "</BRA_DEPNO>")
                        sbXmlData.Append("<OPERATION>" + dtListBoxMove.Rows(i)("OPERATION") + "</OPERATION>")
                        sbXmlData.Append("<TYPE>2</TYPE>")
                        sbXmlData.Append("<BRANAME>" + ddlDepart.SelectedItem.Text.Trim() + "</BRANAME>")
                        sbXmlData.Append("<USERNAME>" + dtListBoxMove.Rows(i)("TEXT").ToString().Split(";")(0) + "</USERNAME>")
                        sbXmlData.Append("<ROLENAME>" + ddlRole.SelectedItem.Text + "</ROLENAME>")
                        sbXmlData.Append("</SY_REL_ROLE_USER>")
                    Next
                    sbXmlData.Append(getOLstBoxRright())
                    sbXmlData.Append("</SY>")

                    ' 如果案件已在流程中
                    If m_sFourStepNo = "" Then

                        ' 判斷DB中是否有資料
                        If syTempInfo.loadByPK(m_sWorkingUserid, m_sWorkingTopDepNo, m_sFuncCode) Then

                            ' 如果是Insert資料，則需要進行追加 如果是修改 則修改
                            oldXmlData = syTempInfo.getAttribute("TEMPDATA").ToString()
                        End If
                    Else

                        ' 判斷DB中是否有資料
                        If syTempInfo.loadByCaseId(m_sCaseId) Then

                            ' 如果是Insert資料，則需要進行追加 如果是修改 則修改
                            oldXmlData = syTempInfo.getAttribute("TEMPDATA").ToString()
                        End If
                    End If

                    ' 若有臨時檔資料
                    If oldXmlData <> "" Then

                        ' document對象載入XML文件
                        xmlDocument.LoadXml(oldXmlData)

                        ' 找到臨時檔中已經存在的資料
                        For i As Integer = 0 To xmlDocument.GetElementsByTagName("SY_REL_ROLE_USER").Count - 1
                            If DirectCast(xmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("ROLEID"), System.Xml.XmlElement).InnerText = ddlRole.SelectedItem.Value _
                                And DirectCast(xmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("BRANAME"), System.Xml.XmlElement).InnerText = ddlDepart.SelectedItem.Text.Trim() Then

                                item = xmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)
                                nodeList.Add(item)
                            End If
                        Next

                        ' 如果有 則將其替換掉 然後累加上新的XML
                        If nodeList.Count > 0 Then
                            Dim newXML As String = oldXmlData.ToString()
                            Dim oldXML As String = String.Empty

                            ' 循環取得相同的欄位
                            For i As Integer = 0 To nodeList.Count - 1
                                If Not nodeList.Item(i) Is Nothing Then
                                    Dim oldItemXML As String = nodeList.Item(i).InnerXml.ToString()

                                    ' 若是沒有異動過的資料
                                    If Not oldItemXML.Contains("<OPERATION>N</OPERATION>") Then
                                        oldXML = oldXML & "<SY_REL_ROLE_USER>" & oldItemXML.ToString() & "</SY_REL_ROLE_USER>"
                                    End If

                                    newXML = newXML.Replace(oldItemXML.ToString(), "").Replace("<SY_REL_ROLE_USER></SY_REL_ROLE_USER>", "")
                                End If
                            Next
                            newXmlData = newXML.ToString().Replace("</SY>", "") & oldXML & sbXmlData.Replace("<SY>", "").ToString()
                        Else

                            ' 如果找不到，直接進行累加
                            newXmlData = oldXmlData.ToString().Replace("</SY>", "") & sbXmlData.Replace("<SY>", "").ToString()
                        End If
                    Else

                        ' 新增資料
                        newXmlData = sbXmlData.ToString()
                    End If

                    If newXmlData = "<SY></SY>" Then
                        newXmlData = ""
                    End If

                    ' 如果案件已在流程中
                    If m_sFourStepNo = "" Then

                        ' 修改臨時檔資料
                        If Not syTempInfo.loadByPK(m_sWorkingUserid, m_sWorkingTopDepNo, m_sFuncCode) Then

                            ' 新增異動數據到臨時表
                            syTempInfo.setAttribute("STAFFID", m_sWorkingUserid)
                            syTempInfo.setAttribute("BRA_DEPNO", m_sWorkingTopDepNo)
                            syTempInfo.setAttribute("FUNCCODE", m_sFuncCode)
                            syTempInfo.setAttribute("CASEID", DBNull.Value)
                        End If

                        If newXmlData = "" Then
                            syTempInfo.remove()
                        Else
                            syTempInfo.setAttribute("TEMPDATA", newXmlData.ToString())
                            syTempInfo.setAttribute("CASEID", DBNull.Value)

                            syTempInfo.save()
                        End If
                Else

                    '修噶臨時檔資料
                    If syTempInfo.loadByCaseId(m_sCaseId) Then
                        syTempInfo.setAttribute("CASEID", m_sCaseId)
                    Else
                        syTempInfo.setAttribute("CASEID", DBNull.Value)
                    End If

                        If String.IsNullOrEmpty(m_sCaseId) = False Then
                            syTempInfo.setAttribute("CASEID", m_sCaseId)
                        End If

                        syTempInfo.setAttribute("STAFFID", m_sWorkingUserid)
                        syTempInfo.setAttribute("BRA_DEPNO", m_sWorkingTopDepNo)
                        syTempInfo.setAttribute("FUNCCODE", m_sFuncCode)

                        If newXmlData = "" Then
                            syTempInfo.remove()
                        Else
                            syTempInfo.setAttribute("TEMPDATA", newXmlData.ToString())
                            syTempInfo.save()
                        End If
                    End If

                    com.Azion.EloanUtility.UIUtility.alert("儲存成功！")
                End If

                dtListBoxMove.Rows.Clear()

                ViewState("dtListBoxMove") = dtListBoxMove
            End If

            ' 綁定分派清單
            BindDetail("")

            ' 異動資料變紅
            editListItemColor()

            GetDatabaseManager.commit()
        Catch ex As Exception
            GetDatabaseManager.Rollback()

            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' ”單位“下拉選單變化時的事件
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub ddlDepart_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles ddlDepart.SelectedIndexChanged

        ' 清空存儲絕對異動的數據
        Dim dtListBoxMove As DataTable = ViewState("dtListBoxMove")
        dtListBoxMove.Rows.Clear()
        ViewState("dtListBoxMove") = dtListBoxMove

        Dim dtFlow As DataTable = ViewState("dataFlow")
        dtFlow.Rows.Clear()
        ViewState("dataFlow") = dtFlow

        ViewState("inOrOutXml") = ""

        ddlRole.Items.Clear()
        ddlUser.Items.Clear()

        ' 重新綁定人員名稱或角色名稱下拉選單
        bindDropDownList()

        ' 清空兩個ListBox的值
        lstBoxRight.Items.Clear()
        lstBoxLeft.Items.Clear()

        '清除選擇
        ddlRole.SelectedIndex = 0
        ddlUser.SelectedIndex = 0
    End Sub

    ''' <summary>
    ''' "角色名稱"下拉選單變化時的事件
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub ddlRole_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles ddlRole.SelectedIndexChanged

        ' 依角色
        If rdolistSetStyle.SelectedValue = "2" Then

            ' 清空存儲絕對異動的數據
            Dim dtListBoxMove As DataTable = ViewState("dtListBoxMove")
            dtListBoxMove.Rows.Clear()
            ViewState("dtListBoxMove") = dtListBoxMove

            Dim dtFlow As DataTable = ViewState("dataFlow")
            dtFlow.Rows.Clear()
            ViewState("dataFlow") = dtFlow

            ViewState("inOrOutXml") = ""

            ' 清空左右兩個ListBox的的值
            lstBoxLeft.Items.Clear()
            lstBoxRight.Items.Clear()

            ' 重新綁定ListBox
            initXmlListBox("2")

            ' 將移動過的數據變色
            editListItemColor()

        End If

    End Sub

    ''' <summary>
    ''' "人員名稱下拉選單變化時的事件"
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub ddlUser_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles ddlUser.SelectedIndexChanged

        ' 依人員
        If rdolistSetStyle.SelectedValue = "1" Then

            ' 清空存儲絕對異動的數據
            Dim dtListBoxMove As DataTable = ViewState("dtListBoxMove")
            dtListBoxMove.Rows.Clear()
            ViewState("dtListBoxMove") = dtListBoxMove

            Dim dtFlow As DataTable = ViewState("dataFlow")
            dtFlow.Rows.Clear()
            ViewState("dataFlow") = dtFlow

            ViewState("inOrOutXml") = ""

            ' 清空左右兩個ListBox的的值
            lstBoxLeft.Items.Clear()
            lstBoxRight.Items.Clear()

            ' 重新綁定ListBox
            initXmlListBox("1")

            ' 將移動過的數據變色
            editListItemColor()
        End If
    End Sub

    ''' <summary>
    ''' "設定方式"的選擇項變化時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub rdolistSetStyle_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles rdolistSetStyle.SelectedIndexChanged

        ' 清空存儲絕對異動的數據
        Dim dtListBoxMove As DataTable = ViewState("dtListBoxMove")
        dtListBoxMove.Rows.Clear()

        ' 如果設定方式為“依人員”
        If rdolistSetStyle.SelectedValue = "1" Then

            ' 顯示【人員名稱】下拉選單欄位，【角色名稱】下拉選單隱藏
            trRole.Style.Value = "display:none"
            trUser.Style.Value = "display:block"
            trByPeopleDetail.Style.Value = "display:block"
            trByRoleDetail.Style.Value = "display:none"

            ddlDepart.SelectedIndex = 0
            ddlUser.Items.Clear()
            ddlUser.Items.Insert(0, New ListItem("--請選擇--", ""))
        Else
            ' 顯示【人員名稱】下拉選單欄位，【角色名稱】下拉選單隱藏
            trRole.Style.Value = "display:block"
            trUser.Style.Value = "display:none"
            trByRoleDetail.Style.Value = "display:block"
            trByPeopleDetail.Style.Value = "display:none"

            ddlDepart.SelectedIndex = 0
            ddlRole.Items.Clear()
            ddlRole.Items.Insert(0, New ListItem("--請選擇--", ""))
        End If

        ViewState("inOrOutXml") = ""

        ' 清空兩個ListBox的值
        lstBoxRight.Items.Clear()
        lstBoxLeft.Items.Clear()

        ' 重新綁定清單
        BindDetail("1")

        ' 綁定ListBox
        initXmlListBox(rdolistSetStyle.SelectedValue)

        ' 顯示顏色
        editListItemColor()
    End Sub

    ''' <summary>
    ''' "依角色"明細檔編輯或刪除事件
    ''' </summary>
    ''' <param name="source"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub dgByRoleDetail_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles dgByRoleDetail.ItemCommand
        Try
            ' 實例化
            'Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager())
            Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager(), IIf(String.IsNullOrEmpty(m_sCaseId), "SY_TEMPINFO", "SY_TEMPINFO2"))

            Dim syUserList As New AUTH_OP.SY_USERList(GetDatabaseManager())

            Dim oldXmlData As String = String.Empty   ' 臨時檔資料
            Dim sRoleId As String = String.Empty    ' 角色編號
            Dim XmlDocument As New XmlDocument
            Dim sType As String = String.Empty
            Dim item As XmlNode
            Dim newXmlData As String = String.Empty
            Dim newXML As String = String.Empty
            Dim nodeList As IList(Of XmlNode) = New List(Of XmlNode)() ' Xml 節點集合

            ' 若命令被觸發
            If e.CommandArgument <> "" Then

                ' 點擊編輯按鈕，顯示編輯行
                If e.CommandName = "Edit" Then

                    Dim dtFlow As DataTable = ViewState("dataFlow")
                    dtFlow.Rows.Clear()
                    ViewState("dataFlow") = dtFlow

                    ' 清空已選擇的人員
                    lstBoxRight.Items.Clear()
                    lstBoxLeft.Items.Clear()

                    ddlDepart.SelectedItem.Selected = False
                    ddlRole.SelectedItem.Selected = False

                    ' 得到當前行的索引
                    Dim index As Integer = e.Item.ItemIndex
                    Dim sStaffids As String = String.Empty

                    Dim lblDepart As Label = dgByRoleDetail.Items(index).FindControl("lblDepart")
                    Dim hidBraDepno As HiddenField = dgByRoleDetail.Items(index).FindControl("hidBraDepno")
                    Dim lblRoleName As Label = dgByRoleDetail.Items(index).FindControl("lblRoleName")
                    Dim hilRoleID As HiddenField = dgByRoleDetail.Items(index).FindControl("hilRoleID")

                    ' 下拉選單重新綁定
                    ddlDepart.Items.FindByValue(hidBraDepno.Value).Selected = True

                    ' 重新綁定角色下拉選單
                    bindDropDownList()

                    ddlRole.Items.FindByValue(hilRoleID.Value).Selected = True

                    ' 取得臨時檔中的員工編號
                    sStaffids = getRoleIDsOrStaffids("2", hilRoleID.Value, lblDepart.Text)

                    ' 如果有人員編號
                    If sStaffids.Length > 0 Then
                        sStaffids = sStaffids.Substring(0, sStaffids.Length - 1)
                    End If

                    ' 如果能查詢出"待選人員資料"
                    If syUserList.loadByRoleEdit(sStaffids, hidBraDepno.Value) Then
                        lstBoxLeft.DataSource = syUserList.getCurrentDataSet
                        lstBoxLeft.DataTextField = "LText"
                        lstBoxLeft.DataValueField = "LValue"
                        lstBoxLeft.DataBind()
                    End If

                    ' 異動數據顯示顏色
                    editListItemColor()
                End If

                ' 如果是刪除
                If Me.hidOutAction.Value = "D" Then

                    ' 得到當前行的索引
                    Dim index As Integer = e.Item.ItemIndex
                    Dim lblDepart As Label = dgByRoleDetail.Items(index).FindControl("lblDepart")

                    ' 獲取此行的主鍵
                    sRoleId = CType(e.CommandArgument, Integer)

                    ' 取得臨時檔資料
                    oldXmlData = getXmlData(syTempInfo)

                    ' 若臨時檔有資料
                    If oldXmlData <> "" Then

                        ' document對象載入XML文件
                        XmlDocument.LoadXml(oldXmlData)

                        ' 找到臨時檔中已經存在的資料
                        For i As Integer = 0 To XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER").Count - 1
                            If DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("ROLEID"), System.Xml.XmlElement).InnerText = sRoleId _
                                 AndAlso DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("BRANAME"), System.Xml.XmlElement).InnerText = lblDepart.Text Then

                                item = XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)
                                nodeList.Add(item)
                            End If
                        Next

                        ' 如果有 則將其替換掉 然後累加上新的XML
                        If nodeList.Count > 0 Then

                            newXML = oldXmlData.ToString()

                            For i As Integer = 0 To nodeList.Count - 1
                                If Not nodeList(i) Is Nothing Then
                                    Dim oldXML As String = nodeList(i).InnerXml.ToString()

                                    newXML = newXML.Replace(oldXML.ToString(), "").Replace("<SY_REL_ROLE_USER></SY_REL_ROLE_USER>", "")
                                End If
                            Next
                        End If
                    End If

                    If newXML.Equals("<SY></SY>") Then
                        newXML = ""
                    End If

                    If newXML = "" Then
                        syTempInfo.remove()
                    Else
                        syTempInfo.setAttribute("TEMPDATA", newXML.ToString())
                        syTempInfo.save()
                    End If

                    ' 提示信息
                    com.Azion.EloanUtility.UIUtility.alert("刪除成功！")

                    ddlDepart.SelectedIndex = -1
                    ddlRole.Items.Clear()
                    ddlRole.Items.Insert(0, New ListItem("--請選擇--", ""))
                    lstBoxLeft.Items.Clear()
                    lstBoxRight.Items.Clear()

                    ' 重新綁定分派清單
                    BindDetail("1")

                    ' 重新綁定ListBox
                    initXmlListBox(rdolistSetStyle.SelectedValue)

                    ' 顯示顏色
                    editListItemColor()

                    Me.hidOutAction.Value = ""
                End If
            End If
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' "依人員"明細檔編輯或刪除事件
    ''' </summary>
    ''' <param name="source"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub dgByPeopleDetail_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles dgByPeopleDetail.ItemCommand
        Try
            ' 實例化
            'Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager())
            Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager(), IIf(String.IsNullOrEmpty(m_sCaseId), "SY_TEMPINFO", "SY_TEMPINFO2"))

            Dim syRoleList As New AUTH_OP.SY_ROLEList(GetDatabaseManager())
            Dim oldXmlData As String = String.Empty
            Dim sStaffid As String = String.Empty ' 員工編號
            Dim XmlDocument As New XmlDocument
            Dim item As XmlNode
            Dim newXML As String = String.Empty
            Dim sRoleids As String = String.Empty ' 角色編號
            Dim nodeList As IList(Of XmlNode) = New List(Of XmlNode)() ' Xml節點集合

            ' 若有命令事件
            If e.CommandArgument <> "" Then

                ' 點擊編輯按鈕，顯示編輯行
                If e.CommandName = "Edit" Then

                    ' 清空已選擇資料
                    Dim dtFlow As DataTable = ViewState("dataFlow")
                    dtFlow.Rows.Clear()
                    ViewState("dataFlow") = dtFlow

                    ddlDepart.SelectedItem.Selected = False
                    ddlUser.SelectedItem.Selected = False

                    ' 清空已選擇的人員
                    lstBoxRight.Items.Clear()
                    lstBoxLeft.Items.Clear()

                    ' 得到當前行的索引
                    Dim index As Integer = e.Item.ItemIndex
                    Dim sStaffids As String = String.Empty

                    Dim lblDepart As Label = dgByPeopleDetail.Items(index).FindControl("lblDepart")
                    Dim lblUserName As Label = dgByPeopleDetail.Items(index).FindControl("lblUserName")
                    Dim hidUserID As HiddenField = dgByPeopleDetail.Items(index).FindControl("hidUserID")
                    Dim hidBraDepno As HiddenField = dgByPeopleDetail.Items(index).FindControl("hidBraDepno")

                    ddlDepart.Items.FindByValue(hidBraDepno.Value).Selected = True

                    ' 重新綁定人員下拉選單
                    bindDropDownList()

                    ddlUser.Items.FindByValue(hidUserID.Value & ";" & hidBraDepno.Value).Selected = True

                    ' 取得臨時檔中的員工編號
                    sRoleids = getRoleIDsOrStaffids("1", hidUserID.Value, lblDepart.Text)

                    ' 如果有人員編號
                    If sRoleids.Length > 0 Then
                        sRoleids = sRoleids.Substring(0, sRoleids.Length - 1)
                    End If

                    ' 如果能查詢出"待選資料"
                    If syRoleList.loadByPeopleEdit(hidBraDepno.Value, sRoleids, m_sHoFlag, "," & m_sWorkingRoleId & ",") Then
                        lstBoxLeft.DataSource = syRoleList.getCurrentDataSet
                        lstBoxLeft.DataTextField = "LText"
                        lstBoxLeft.DataValueField = "LValue"
                        lstBoxLeft.DataBind()
                    End If

                    ' 異動數據顯示顏色
                    editListItemColor()
                End If

                ' 獲取此行的主鍵
                sStaffid = CType(e.CommandArgument, String)

                ' 如果是刪除
                If Me.hidOutAction.Value = "D" Then

                    ' 得到當前行的索引
                    Dim index As Integer = e.Item.ItemIndex
                    Dim lblDepart As Label = dgByPeopleDetail.Items(index).FindControl("lblDepart")

                    ' 取得臨時檔資料
                    oldXmlData = getXmlData(syTempInfo)

                    If oldXmlData <> "" Then

                        ' document對象載入XML文件
                        XmlDocument.LoadXml(oldXmlData)

                        ' 找到臨時檔中已經存在的資料
                        For i As Integer = 0 To XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER").Count - 1
                            If DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("STAFFID"), System.Xml.XmlElement).InnerText = sStaffid _
                                And DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("BRANAME"), System.Xml.XmlElement).InnerText = lblDepart.Text Then

                                item = XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)
                                nodeList.Add(item)
                            End If
                        Next

                        ' 如果有 則將其替換掉 然後累加上新的XML
                        If nodeList.Count > 0 Then
                            newXML = oldXmlData.ToString()
                            For i As Integer = 0 To nodeList.Count - 1
                                If Not nodeList(i) Is Nothing Then
                                    Dim oldXML As String = nodeList(i).InnerXml.ToString()

                                    newXML = newXML.Replace(oldXML.ToString(), "").Replace("<SY_REL_ROLE_USER></SY_REL_ROLE_USER>", "")
                                End If
                            Next
                        End If
                    End If

                    If newXML.Equals("<SY></SY>") Then
                        newXML = ""
                    End If

                    If newXML = "" Then
                        syTempInfo.remove()
                    Else
                        syTempInfo.setAttribute("TEMPDATA", newXML.ToString())
                        syTempInfo.save()
                    End If

                    ' 提示信息
                    com.Azion.EloanUtility.UIUtility.alert("刪除成功！")

                    ddlDepart.SelectedIndex = -1
                    ddlUser.Items.Clear()
                    ddlUser.Items.Insert(0, New ListItem("--請選擇--", ""))
                    lstBoxLeft.Items.Clear()
                    lstBoxRight.Items.Clear()

                    ' 重新綁定分派清單
                    BindDetail("1")

                    ' 重新綁定ListBox
                    initXmlListBox(rdolistSetStyle.SelectedValue)

                    ' 顯示顏色
                    editListItemColor()

                    Me.hidOutAction.Value = ""
                End If
            End If
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 點擊"確認送出"按鈕事件
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnSendFlow_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnSendFlow.Click
        Try
            ' 實例化
            'Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager())
            Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager(), IIf(String.IsNullOrEmpty(m_sCaseId), "SY_TEMPINFO", "SY_TEMPINFO2"))

            Dim syRelRoleUserHis As New AUTH_OP.SY_REL_ROLE_USER_HIS(GetDatabaseManager())

            Dim oldXmlData As String = String.Empty  ' 臨時檔資歷
            Dim XmlDocument As New XmlDocument
            Dim sRoleID As String = String.Empty   ' 角色編號
            Dim sStaffid As String = String.Empty   ' 員工編號
            Dim sRoleName As String = String.Empty   ' 角色名稱

            Dim sUserID As String = String.Empty
            Dim sBraDepno As String = String.Empty
            Dim stepInfo As New FLOW_OP.StepInfo()
            Dim RoleType As String


            Dim sRoleIdArray As String()
            sRoleIdArray = m_sWorkingRoleId.Split(New [Char]() {","c})

            Dim bNeedReview As Boolean = True

            For Each item As String In sRoleIdArray
                RoleType = FLOW_OP.FlowFacade.getNewInstance(GetDatabaseManager).getSYRole.GetRoleType(item)

                If RoleType = "1" Then
                    bNeedReview = False
                    Exit For
                End If
            Next

            '檢查是否有覆核人員
            If bNeedReview = True AndAlso ddlReviewer.Items.Count = 0 Then
                com.Azion.EloanUtility.UIUtility.alert("沒有覆核人員，不能進行送出！")
                Return
            End If


            ' 取得臨時檔資料
            oldXmlData = getXmlData(syTempInfo)

            If oldXmlData <> "" Then

                ' document對象載入XML文件
                XmlDocument.LoadXml(oldXmlData)

                For i As Integer = 0 To XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER").Count - 1
                    Dim RoleID As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("ROLEID"), System.Xml.XmlElement).InnerText
                    Dim STAFFID As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("STAFFID"), System.Xml.XmlElement).InnerText
                    Dim ROLENAME As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("ROLENAME"), System.Xml.XmlElement).InnerText
                    Dim operation As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("OPERATION"), System.Xml.XmlElement).InnerText

                    If STAFFID = ddlReviewer.SelectedValue AndAlso operation <> "N" Then
                        com.Azion.EloanUtility.UIUtility.alert("覆核人員不能覆核自己的角色設定，不能進行送出！")
                        Return
                    End If

                    sRoleName = ROLENAME
                    sRoleID = sRoleID & RoleID & ","
                    sStaffid = sStaffid & STAFFID & ","
                Next
            Else
                com.Azion.EloanUtility.UIUtility.alert("沒有送簽資料，不能進行送出！")

                Return
            End If

            ' 如果有角色ID
            If sRoleID.Length > 0 Then
                sRoleID = sRoleID.Substring(0, sRoleID.Length - 1)
            End If

            ' 如果有人員編號
            If sStaffid.Length > 0 Then
                sStaffid = sStaffid.Substring(0, sStaffid.Length - 1)
            End If

            ' 如果已在處理中
            If syRelRoleUserHis.loadByRoleIDStaffid(m_sStepNo, m_sCaseId, sRoleID, sStaffid) Then
                com.Azion.EloanUtility.UIUtility.alert(sRoleName + "已在處理中，無法送件！")

                Return
            End If

            ' 如果是測試模式

            'If RoleType = "1" Then
            '    m_bCheck = True
            'Else
            '    m_bCheck = False
            'End If


            If RoleType = "1" Then
                ' 開始事務
                GetDatabaseManager().beginTran()

                'stepInfo = com.Azion.UITools.UIShareFun.startFlow(GetDatabaseManager(), m_sFlowName, True, , True)
                stepInfo = FlowFacade.getNewInstance(GetDatabaseManager).StartFlow(
                    Nothing, m_sFlowName, True, ddlReviewer.SelectedValue, True)

                flowFuntion("Y", "agree", stepInfo)

                ' 如果案件没在流程中
                If m_sFourStepNo = "" Then

                    ' 依據主鍵刪除
                    syTempInfo.loadByPK(m_sWorkingUserid, m_sWorkingTopDepNo, m_sFuncCode)
                    syTempInfo.remove()
                Else

                    ' 根據CaseID刪除數據
                    'syTempInfo.deleteByCaseID(m_sCaseId)
                    BosBase.getNewBosBase("SY_TEMPINFO", GetDatabaseManager()).Delete("CASEID", m_sCaseId)
                    BosBase.getNewBosBase("SY_TEMPINFO2", GetDatabaseManager()).Delete("CASEID", m_sCaseId)
                End If

                ' 提交事務
                GetDatabaseManager().commit()

                ' 成功提示信息
                com.Azion.EloanUtility.UIUtility.showJSMsg("案件編號：" + stepInfo.currentStepInfo.caseId + " 儲存成功！")

                ' 跳轉到待辦事項頁面
                com.Azion.EloanUtility.UIUtility.goMainPage("")

                '' 綁定單位下拉選單
                'initDDlDepart()

                '' 依人員
                'If rdolistSetStyle.SelectedValue = "1" Then
                '    ddlUser.Items.Clear()
                '    ddlUser.Items.Insert(0, New ListItem("--請選擇--", ""))
                'Else
                '    ddlRole.Items.Clear()
                '    ddlRole.Items.Insert(0, New ListItem("--請選擇--", ""))
                'End If

                'lstBoxLeft.Items.Clear()
                'lstBoxRight.Items.Clear()

                '' 綁定異動數據清單
                'BindDetail("1")

                '' 重新綁定數據
                initData()
            Else
                ' 開始事務
                GetDatabaseManager().beginTran()

                ' 如果是起案
                If m_sStepNo = "" Then

                    ' 呼叫StartFlow（）方法,得到返回值
                    'stepInfo = com.Azion.UITools.UIShareFun.startFlow(GetDatabaseManager(), m_sFlowName, True)
                    stepInfo = FlowFacade.getNewInstance(GetDatabaseManager).StartFlow(
                        Nothing, m_sFlowName, True, ddlReviewer.SelectedValue, False)


                    ' 更改SY_TEMPINFO裱中的CaseID欄位
                    syTempInfo.setAttribute("CASEID", stepInfo.currentStepInfo.caseId.ToString())
                    syTempInfo.save()

                    '有案號時將存入SY_TEMPINFO改存入SY_TEMPINFO2
                    Dim dr As DataRow
                    dr = syTempInfo.GetDataRow(Nothing, "CASEID", stepInfo.currentStepInfo.caseId)

                    Dim tempInfo2 As New SY_TEMPINFO(GetDatabaseManager(), "SY_TEMPINFO2")
                    tempInfo2.InsertUpdate(dr)

                    syTempInfo.Delete("CASEID", stepInfo.currentStepInfo.caseId.ToString())

                    m_sCaseId = stepInfo.currentStepInfo.caseId
                    m_sFourStepNo = Right(stepInfo.currentStepInfo.stepNo, 4)
                Else
                    stepInfo = MBSC.UICtl.UIShareFun.pushFlow(GetDatabaseManager(), m_sCaseId)
                End If

                ' 新增歷史檔資料
                flowFuntion("", "", stepInfo)

                ' 提交事務
                GetDatabaseManager().commit()

                ' 成功提示信息
                com.Azion.EloanUtility.UIUtility.showJSMsg("案件編號：" + stepInfo.currentStepInfo.caseId + " 儲存成功！")

                ' 跳轉到待辦事項頁面
                com.Azion.EloanUtility.UIUtility.goMainPage("")

            End If
        Catch ex As Exception

            ' 回滾事務
            GetDatabaseManager().Rollback()

            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 點擊取消按鈕
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnCancel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnCancel.Click
        Try
            lstBoxLeft.Items.Clear()
            lstBoxRight.Items.Clear()

            Dim dtFlow As DataTable = ViewState("dataFlow")
            dtFlow.Rows.Clear()
            ViewState("dataFlow") = dtFlow

            ' 重新綁定ListBox
            initXmlListBox(rdolistSetStyle.SelectedValue)

            ' 清空存儲絕對異動的數據
            Dim dtListBoxMove As DataTable = ViewState("dtListBoxMove")
            dtListBoxMove.Rows.Clear()
            ViewState("dtListBoxMove") = dtListBoxMove

            ViewState("inOrOutXml") = ""

            ' 異動數據顯示紅色
            editListItemColor()
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 點擊全部取消按鈕事件
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnAllCancel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnAllCancel.Click
        Try
            GetDatabaseManager().beginTran()
            'Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager())
            Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager(), IIf(String.IsNullOrEmpty(m_sCaseId), "SY_TEMPINFO", "SY_TEMPINFO2"))

            lstBoxLeft.Items.Clear()
            lstBoxRight.Items.Clear()

            ' 依人員
            If rdolistSetStyle.SelectedValue = "1" Then
                ' 隱藏資料清單
                trByPeopleDetail.Style.Value = "display:none"

                ' 清空下拉選單
                ddlUser.Items.Clear()
            Else
                ' 隱藏資料清單
                trByRoleDetail.Style.Value = "display:none"

                ' 清空下拉選單
                ddlRole.Items.Clear()
            End If

            Dim dtFlow As DataTable = ViewState("dataFlow")
            dtFlow.Rows.Clear()
            ViewState("dataFlow") = dtFlow

            ' 清空存儲絕對異動的數據
            Dim dtListBoxMove As DataTable = ViewState("dtListBoxMove")
            dtListBoxMove.Rows.Clear()
            ViewState("dtListBoxMove") = dtListBoxMove

            ViewState("inOrOutXml") = ""

            ' 刪除臨時檔
            If m_sFourStepNo = "" Then
                If syTempInfo.loadByPK(m_sWorkingUserid, m_sWorkingTopDepNo, m_sFuncCode) Then
                    syTempInfo.remove()
                End If
            Else
                If syTempInfo.loadByCaseId(m_sCaseId) Then
                    syTempInfo.remove()
                End If
            End If

            ' 綁定單位下拉選單
            initDDlDepart()
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
            GetDatabaseManager.Rollback()
        Finally
            GetDatabaseManager.commit()
        End Try
    End Sub
#End Region
#End Region

#Region "複核區"

#Region "Fucntion"

    ''' <summary>
    ''' 綁定審核區的資料
    ''' </summary>
    ''' <remarks></remarks>
    Sub initCheckListBox()
        'Dim syTempInfoXml As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager())
        Dim syTempInfoXml As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager(), IIf(String.IsNullOrEmpty(m_sCaseId), "SY_TEMPINFO", "SY_TEMPINFO2"))


        ' 聲明分派清單表
        Dim dtGrantDetail As New DataTable
        dtGrantDetail.Columns.Add("DEPNO", Type.GetType("System.String"))
        dtGrantDetail.Columns.Add("BRANAME", Type.GetType("System.String"))
        dtGrantDetail.Columns.Add("NAME", Type.GetType("System.String"))
        dtGrantDetail.Columns.Add("ID", Type.GetType("System.String"))

        Dim sDepartName As String = String.Empty ' 單位編號
        Dim sUserID As String = String.Empty  ' 人員編號
        Dim sRoleID As String = String.Empty   ' 角色編號
        Dim sType As String = String.Empty   ' 設定方式
        Dim sStaffid As String = String.Empty   ' 人員編號
        Dim sDepartNO As String = String.Empty   ' 部門編號
        Dim sRoleNO As String = String.Empty    ' 角色編號
        Dim sXmlData As String = String.Empty
        Dim XmlDocument As New XmlDocument

        If syTempInfoXml.loadByCaseId(m_sCaseId) Then
            sXmlData = syTempInfoXml.getAttribute("TEMPDATA")

            If sXmlData <> "" Then

                ' document對象載入XML文件
                XmlDocument.LoadXml(sXmlData)

                ' 如果有數據
                If XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER").Count > 0 Then
                    Dim sOneStaffid = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(0)("STAFFID"), System.Xml.XmlElement).InnerText
                    Dim sOneRoleID = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(0)("ROLEID"), System.Xml.XmlElement).InnerText
                    Dim sOneBRAName = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(0)("BRANAME"), System.Xml.XmlElement).InnerText
                    Dim sOneUserName = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(0)("USERNAME"), System.Xml.XmlElement).InnerText
                    Dim sOneROLENAME = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(0)("ROLENAME"), System.Xml.XmlElement).InnerText

                    For i As Integer = 0 To XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER").Count - 1
                        Dim ROLEID As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("ROLEID"), System.Xml.XmlElement).InnerText
                        Dim STAFFID As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("STAFFID"), System.Xml.XmlElement).InnerText
                        Dim BRA_DEPNO As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("BRA_DEPNO"), System.Xml.XmlElement).InnerText
                        Dim OPERATION As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("OPERATION"), System.Xml.XmlElement).InnerText
                        Dim TYPE As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("TYPE"), System.Xml.XmlElement).InnerText
                        Dim BRANAME As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("BRANAME"), System.Xml.XmlElement).InnerText
                        Dim USERNAME As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("USERNAME"), System.Xml.XmlElement).InnerText
                        Dim ROLENAME As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("ROLENAME"), System.Xml.XmlElement).InnerText

                        ' 如果設定方式為“依人員”
                        If TYPE = "1" Then

                            ' 顯示人員分派清單
                            trUserGrant.Style.Value = "display:block"
                            trRoleGrant.Style.Value = "display:none"
                            sType = "1"
                            sStaffid = STAFFID
                            sDepartNO = BRA_DEPNO
                            lblMsg.Text = sOneUserName

                            If BRANAME <> sDepartName Or sUserID <> STAFFID Then
                                sUserID = STAFFID
                                sDepartName = BRANAME

                                Dim row As DataRow = dtGrantDetail.NewRow
                                row("DEPNO") = BRA_DEPNO
                                row("BRANAME") = BRANAME
                                row("NAME") = USERNAME
                                row("ID") = STAFFID

                                dtGrantDetail.Rows.Add(row)
                            End If

                            ' 取出第一個人員的資料
                            If i > 0 Then
                                If sOneStaffid = STAFFID And sOneBRAName = BRANAME Then

                                    ' 沒有異動的數據
                                    If OPERATION = "N" Then
                                        lstPreEdit.Items.Add(New ListItem(ROLENAME, ROLEID))
                                        lstAfterEdit.Items.Add(New ListItem(ROLENAME, ROLEID))
                                    ElseIf OPERATION = "I" Then
                                        Dim item As ListItem = New ListItem(ROLENAME, ROLEID)
                                        item.Attributes.Add("style", "color:red")
                                        lstAfterEdit.Items.Add(item)
                                    ElseIf OPERATION = "D" Then
                                        Dim item As ListItem = New ListItem(ROLENAME, ROLEID)
                                        item.Attributes.Add("style", "color:red")
                                        lstPreEdit.Items.Add(item)
                                    End If
                                End If
                            Else

                                ' 沒有異動的數據
                                If OPERATION = "N" Then
                                    lstPreEdit.Items.Add(New ListItem(ROLENAME, ROLEID))
                                    lstAfterEdit.Items.Add(New ListItem(ROLENAME, ROLEID))
                                ElseIf OPERATION = "I" Then
                                    Dim item As ListItem = New ListItem(ROLENAME, ROLEID)
                                    item.Attributes.Add("style", "color:red")
                                    lstAfterEdit.Items.Add(item)
                                ElseIf OPERATION = "D" Then
                                    Dim item As ListItem = New ListItem(ROLENAME, ROLEID)
                                    item.Attributes.Add("style", "color:red")
                                    lstPreEdit.Items.Add(item)
                                End If
                            End If
                        Else

                            ' 顯示角色分派清單
                            trRoleGrant.Style.Value = "display:block"
                            trUserGrant.Style.Value = "display:none"
                            sType = "2"
                            sDepartNO = BRA_DEPNO
                            sRoleNO = ROLEID
                            lblMsg.Text = sOneROLENAME

                            If BRANAME <> sDepartName Or sRoleID <> ROLEID Then
                                sRoleID = ROLEID
                                sDepartName = BRANAME

                                Dim row As DataRow = dtGrantDetail.NewRow
                                row("DEPNO") = BRA_DEPNO
                                row("BRANAME") = BRANAME
                                row("NAME") = ROLENAME
                                row("ID") = ROLEID

                                dtGrantDetail.Rows.Add(row)
                            End If

                            ' 取出第一個人員的資料
                            If i > 0 Then
                                If sOneRoleID = ROLEID And sOneBRAName = BRANAME Then

                                    ' 如果資料沒有變動
                                    If OPERATION = "N" Then
                                        lstPreEdit.Items.Add(New ListItem(USERNAME, STAFFID))
                                        lstAfterEdit.Items.Add(New ListItem(USERNAME, STAFFID))
                                    ElseIf OPERATION = "I" Then
                                        Dim item As ListItem = New ListItem(USERNAME, STAFFID)
                                        item.Attributes.Add("style", "color:red")
                                        lstAfterEdit.Items.Add(item)
                                    ElseIf OPERATION = "D" Then
                                        Dim item As ListItem = New ListItem(USERNAME, STAFFID)
                                        item.Attributes.Add("style", "color:red")
                                        lstPreEdit.Items.Add(item)
                                    End If
                                End If
                            Else

                                ' 如果資料沒有變動
                                If OPERATION = "N" Then
                                    lstPreEdit.Items.Add(New ListItem(USERNAME, STAFFID))
                                    lstAfterEdit.Items.Add(New ListItem(USERNAME, STAFFID))
                                ElseIf OPERATION = "I" Then
                                    Dim item As ListItem = New ListItem(USERNAME, STAFFID)
                                    item.Attributes.Add("style", "color:red")
                                    lstAfterEdit.Items.Add(item)
                                ElseIf OPERATION = "D" Then
                                    Dim item As ListItem = New ListItem(USERNAME, STAFFID)
                                    item.Attributes.Add("style", "color:red")
                                    lstPreEdit.Items.Add(item)
                                End If
                            End If
                        End If
                    Next
                End If
            End If
        End If

        ' 如果設定方式為“依人員”
        If sType = "1" Then
            Me.lblGrantStyle.Text = "人員分派"

            ' 綁定清單
            dgUserGrant.DataSource = dtGrantDetail
            dgUserGrant.DataBind()
        Else
            Me.lblGrantStyle.Text = "角色分派"

            ' 綁定清單
            dgRoleGrant.DataSource = dtGrantDetail
            dgRoleGrant.DataBind()
        End If
    End Sub
#End Region

#Region "Event"

    ''' <summary>
    ''' 點擊“同意”按鈕事件
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnAgree_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnAgree.Click
        Try
            'Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager())
            Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager(), IIf(String.IsNullOrEmpty(m_sCaseId), "SY_TEMPINFO", "SY_TEMPINFO2"))

            Dim stepInfo As New FLOW_OP.StepInfo()

            ' 開始事務
            GetDatabaseManager().beginTran()

            ' 調用SendFlow方法
            stepInfo = MBSC.UICtl.UIShareFun.pushFlow(GetDatabaseManager(), m_sCaseId)

            ' 對目前資料的流程進行存儲
            flowFuntion("Y", "agree", stepInfo)

            ' 根據CaseID刪除數據
            syTempInfo.deleteByCaseID(m_sCaseId)

            ' 提交事務
            GetDatabaseManager().commit()

            ' 關閉視窗
            com.Azion.EloanUtility.UIUtility.closeWindow()
        Catch ex As Exception

            ' 提交事務
            GetDatabaseManager().Rollback()

            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 點擊"不同意”按鈕事件
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnDiffer_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnDiffer.Click
        Try
            'Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager())
            Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager(), IIf(String.IsNullOrEmpty(m_sCaseId), "SY_TEMPINFO", "SY_TEMPINFO2"))

            Dim stepInfo As New FLOW_OP.StepInfo()

            ' 開始事務
            GetDatabaseManager().beginTran()

            ' 調用SendFlow方法
            stepInfo = MBSC.UICtl.UIShareFun.pushFlow(GetDatabaseManager(), m_sCaseId)

            ' 對目前資料的流程進行存儲
            flowFuntion("N", "", stepInfo)

            ' 刪除臨時檔里的資料
            syTempInfo.deleteByCaseID(m_sCaseId)

            ' 提交事務
            GetDatabaseManager().commit()

            ' 關閉本頁面
            com.Azion.EloanUtility.UIUtility.closeWindow()
        Catch ex As Exception

            ' 回滾事務
            GetDatabaseManager().Rollback()

            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 點擊"修正補充"按鈕事件
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnReviseFlow_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnReviseFlow.Click
        Try
            Dim stepInfo As New FLOW_OP.StepInfo()

            ' 開始事務
            GetDatabaseManager.beginTran()

            ' 調用JumpRollBack方法
            stepInfo = MBSC.UICtl.UIShareFun.rollBack(GetDatabaseManager(), m_sCaseId)

            ' 對目前資料的流程進行存儲
            flowFuntion("", "", stepInfo)

            ' 提交事務
            GetDatabaseManager.commit()

            ' 關閉視窗
            'com.Azion.EloanUtility.UIUtility.closeWindow()
        Catch ex As Exception

            ' 回滾事務
            GetDatabaseManager.Rollback()

            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' "人員分派"清單創建行時的事件
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub dgUserGrant_ItemCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles dgUserGrant.ItemCreated

        ' 當是數據行時或者交替行
        If e.Item.ItemType = ListItemType.Item Or e.Item.ItemType = ListItemType.AlternatingItem Then
            Dim lnkBtnUserName As LinkButton = e.Item.FindControl("lnkBtnUserName")
            AddHandler lnkBtnUserName.Click, AddressOf lnkBtnUserName_Click
        End If
    End Sub

    ''' <summary>
    ''' "角色分派"清單創建行時的事件
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub dgRoleGrant_ItemCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles dgRoleGrant.ItemCreated

        ' 當是數據行時或者交替行
        If e.Item.ItemType = ListItemType.Item Or e.Item.ItemType = ListItemType.AlternatingItem Then
            Dim lnkBtnRoleName As LinkButton = e.Item.FindControl("lnkBtnRoleName")
            AddHandler lnkBtnRoleName.Click, AddressOf lnkBtnRoleName_Click
        End If
    End Sub

    ''' <summary>
    ''' 點擊"人員分派"清單中的人員名稱時的事件
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub lnkBtnUserName_Click(ByVal sender As Object, ByVal e As EventArgs)
        Try
            ' 清空左右兩個ListBox
            lstPreEdit.Items.Clear()
            lstAfterEdit.Items.Clear()

            ' 實例化
            'Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager())
            Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager(), IIf(String.IsNullOrEmpty(m_sCaseId), "SY_TEMPINFO", "SY_TEMPINFO2"))

            Dim sXmlData As String = String.Empty
            Dim XmlDocument As New XmlDocument()

            ' 獲得當前點擊的按鈕
            Dim lnkBtnUserName As LinkButton = sender

            ' 取得單位和人員編號
            Dim hidBtnUserID As HiddenField = lnkBtnUserName.Parent.Parent.FindControl("hidBtnUserID")
            Dim lblDepart As Label = lnkBtnUserName.Parent.Parent.FindControl("lblDepart")

            ' 如果臨時裱中有資料
            If syTempInfo.loadByCaseId(m_sCaseId) Then
                sXmlData = syTempInfo.getAttribute("TEMPDATA")
            End If

            If sXmlData <> "" Then
                ' document對象載入XML文件
                XmlDocument.LoadXml(sXmlData)

                ' 如果有數據
                If XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER").Count > 0 Then
                    For i As Integer = 0 To XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER").Count - 1
                        Dim ROLEID As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("ROLEID"), System.Xml.XmlElement).InnerText
                        Dim STAFFID As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("STAFFID"), System.Xml.XmlElement).InnerText
                        Dim BRA_DEPNO As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("BRA_DEPNO"), System.Xml.XmlElement).InnerText
                        Dim OPERATION As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("OPERATION"), System.Xml.XmlElement).InnerText
                        Dim TYPE As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("TYPE"), System.Xml.XmlElement).InnerText
                        Dim BRANAME As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("BRANAME"), System.Xml.XmlElement).InnerText
                        Dim USERNAME As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("USERNAME"), System.Xml.XmlElement).InnerText
                        Dim ROLENAME As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("ROLENAME"), System.Xml.XmlElement).InnerText

                        ' 如果設定方式為"依人員"
                        If TYPE = "1" Then

                            ' 取出第一個人員的資料
                            If hidBtnUserID.Value = STAFFID And lblDepart.Text = BRANAME Then

                                lblMsg.Text = USERNAME

                                ' 如果是沒有異動的數據
                                If OPERATION = "N" Then
                                    lstPreEdit.Items.Add(New ListItem(ROLENAME, ROLEID))
                                    lstAfterEdit.Items.Add(New ListItem(ROLENAME, ROLEID))
                                ElseIf OPERATION = "I" Then
                                    Dim item As ListItem = New ListItem(ROLENAME, ROLEID)
                                    item.Attributes.Add("style", "color:red")
                                    lstAfterEdit.Items.Add(item)
                                ElseIf OPERATION = "D" Then
                                    Dim item As ListItem = New ListItem(ROLENAME, ROLEID)
                                    item.Attributes.Add("style", "color:red")
                                    lstPreEdit.Items.Add(item)
                                End If
                            End If
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 點擊"角色分派"清單中的角色名稱時的事件
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub lnkBtnRoleName_Click(ByVal sender As Object, ByVal e As EventArgs)
        Try
            ' 清空左右兩個ListBox
            lstPreEdit.Items.Clear()
            lstAfterEdit.Items.Clear()

            ' 獲得當前點擊的按鈕
            Dim lnkBtnRoleName As LinkButton = sender

            ' 取得單位和角色編號
            Dim hidRoleID As HiddenField = lnkBtnRoleName.Parent.Parent.FindControl("hidRoleID")
            Dim lblDepart As Label = lnkBtnRoleName.Parent.Parent.FindControl("lblDepart")

            ' 實例化
            'Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager())
            Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager(), IIf(String.IsNullOrEmpty(m_sCaseId), "SY_TEMPINFO", "SY_TEMPINFO2"))

            Dim sXmlData As String = String.Empty
            Dim XmlDocument As New XmlDocument()

            ' 如果臨時裱中有資料
            If syTempInfo.loadByCaseId(m_sCaseId) Then
                sXmlData = syTempInfo.getAttribute("TEMPDATA")
            End If

            If sXmlData <> "" Then
                ' document對象載入XML文件
                XmlDocument.LoadXml(sXmlData)
            End If

            ' 如果有數據
            If XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER").Count > 0 Then
                For i As Integer = 0 To XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER").Count - 1
                    Dim ROLEID As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("ROLEID"), System.Xml.XmlElement).InnerText
                    Dim STAFFID As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("STAFFID"), System.Xml.XmlElement).InnerText
                    Dim BRA_DEPNO As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("BRA_DEPNO"), System.Xml.XmlElement).InnerText
                    Dim OPERATION As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("OPERATION"), System.Xml.XmlElement).InnerText
                    Dim TYPE As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("TYPE"), System.Xml.XmlElement).InnerText
                    Dim BRANAME As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("BRANAME"), System.Xml.XmlElement).InnerText
                    Dim USERNAME As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("USERNAME"), System.Xml.XmlElement).InnerText
                    Dim ROLENAME As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("ROLENAME"), System.Xml.XmlElement).InnerText

                    ' 如果設定方式為"依角色"
                    If TYPE = "2" Then

                        ' 取出第一個人員的資料
                        If hidRoleID.Value = ROLEID And lblDepart.Text = BRANAME Then

                            lblMsg.Text = ROLENAME

                            ' 如果數據沒有異動
                            If OPERATION = "N" Then
                                lstPreEdit.Items.Add(New ListItem(USERNAME, STAFFID))
                                lstAfterEdit.Items.Add(New ListItem(USERNAME, STAFFID))
                            ElseIf OPERATION = "I" Then
                                Dim item As ListItem = New ListItem(USERNAME, STAFFID)
                                item.Attributes.Add("style", "color:red")
                                lstAfterEdit.Items.Add(item)
                            ElseIf OPERATION = "D" Then
                                Dim item As ListItem = New ListItem(USERNAME, STAFFID)
                                item.Attributes.Add("style", "color:red")
                                lstPreEdit.Items.Add(item)
                            End If
                        End If
                    End If
                Next
            End If
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub
#End Region
#End Region

End Class