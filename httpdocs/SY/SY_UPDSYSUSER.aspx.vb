''' <summary>
''' 程式說明：應用系統人員分派
''' 建立者：Zack 
''' 建立日期：2012-06-01
''' </summary>

Imports com.Azion.NET.VB
Imports AUTH_OP.TABLE
Imports com.Azion.EloanUtility
Imports AUTH_OP
Imports MBSC.MB_OP
Imports System.Xml
Imports FLOW_OP

Public Class SY_UPDSYSUSER
    Inherits SYUIBase

    Dim m_sFuncCode As String = String.Empty
    Dim m_sUpCode2384 As String = String.Empty
    
#Region "Page_load"

    ''' <summary>
    ''' 頁面加載事件
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try

            ' 初始化頁面參數
            initParas()

            If Not IsPostBack Then

                ' 設置初始焦點
                btnSend.Focus()

                ' 初始化頁面數據
                initData()
            End If

            ' 若設定角色沒初始值
            If rdolistRole.SelectedValue = "" Then
                rdolistRole.SelectedValue = "2"
            End If

            ' 設置頁面控件的初始狀態
            If m_bDisplayMode Then

                ' 只讀模式
                com.Azion.EloanUtility.UIUtility.setControlRead(Me)
            End If
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub
#End Region

#Region "維護區"
#Region "Function"

    ''' <summary>
    ''' 初始化參數
    ''' </summary>
    ''' <remarks></remarks>
    Sub initParas()

        Dim syBranch As New AUTH_OP.SY_BRANCH(GetDatabaseManager())

        m_sUpCode2384 = com.Azion.EloanUtility.CodeList.getAppSettings("UpCode2384")

        If Not String.IsNullOrEmpty(m_sCaseId) Then

            ' 案號
            lblCaseId.Text = m_sCaseId

            ' 單位
            If syBranch.loadByCaseId(m_sCaseId) Then
                lblBranch.Text = syBranch.getAttribute("BRCNAME")
            End If
        End If

        ' 測試數據
        If m_bTesting Then
            m_sWorkingUserid = "S002834"
            m_sWorkingTopDepNo = "2"
            m_sFuncCode = "4"
            m_sCaseId = "SY9251019000413"
            m_sFourStepNo = "0300"
            m_bDisplayMode = False
        End If
    End Sub

    ''' <summary>
    ''' 初始化數據
    ''' </summary>
    ''' <remarks></remarks>
    Sub initData()

        Try
            ' 顯示"維護區"
            If m_sFourStepNo = "" Or m_sFourStepNo = "0300" Then

                '顯示“維護區”，隱藏“審核區”
                divEdit.Style.Value = "display:block"
                divChecked.Style.Value = "display:none"

                ' 綁定維護角色
                initRole()

                ' 綁定單位，使用者下拉選單
                initDropList()

                ' 綁定人員信息清單
                initAddDetail("Y")
            Else
                ' 顯示"審核區"

                '顯示“審核區”，隱藏“維護區”
                divEdit.Style.Value = "display:none"
                divChecked.Style.Value = "display:block"

                ' 初始化人員角色清單
                initCheckData()

                ' 如果有角色名稱清單
                If dgRoleDetail.Items.Count > 0 Then
                    Dim hidRoleID As HiddenField = dgRoleDetail.Items(0).FindControl("hidRoleID")
                    Dim lnkBtnRoleName As LinkButton = dgRoleDetail.Items(0).FindControl("lnkBtnRoleName")

                    ' 綁定修改前後的人員信息
                    initListBox(hidRoleID.Value)

                    lblMsg.Text = lnkBtnRoleName.Text
                End If
            End If
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 綁定"維護角色"
    ''' </summary>
    ''' <remarks></remarks>
    Sub initRole()
        Try
            Dim apCodeList As New AP_CODEList(GetDatabaseManager())

            ' 綁定維護角色
            If apCodeList.loadBuUpCodeValue(m_sUpCode2384) Then
                rdolistRole.DataSource = apCodeList.getCurrentDataSet
                rdolistRole.DataTextField = "TEXT"
                rdolistRole.DataValueField = "VALUE"
                rdolistRole.DataBind()
            End If
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 使用者兩個下拉選單
    ''' </summary>
    ''' <remarks></remarks>
    Sub initddlUser()
        Try
            Dim syUserList As New AUTH_OP.SY_USERList(GetDatabaseManager())

            ' 綁定使用者信息
            If syUserList.loadByBRA_DEPNO(ddlDepart.SelectedItem.Value) Then
                ddlUser.DataSource = syUserList.getCurrentDataSet
                ddlUser.DataTextField = "USERNAME"
                ddlUser.DataValueField = "BRA_DEPNO"
                ddlUser.DataBind()
            End If

            ddlUser.Items.Insert(0, New ListItem("--請選擇--", ""))
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 綁定單位
    ''' </summary>
    ''' <remarks></remarks>
    Sub initDropList()
        Try
            Dim syBranchList As New AUTH_OP.SY_BRANCHList(GetDatabaseManager())
            Dim dt As DataTable = syBranchList.loadBranchInfo(m_sWorkingBrid, "1", "0", 5)

            ' 若有單位資料
            If Not dt Is Nothing Then
                ddlDepart.DataSource = syBranchList.getCurrentDataSet
                ddlDepart.DataTextField = "NAME"
                ddlDepart.DataValueField = "BRA_DEPNO"
                ddlDepart.DataBind()
            End If

            ddlDepart.Items.Insert(0, New ListItem("--請選擇--", ""))

            ' 使用者兩個下拉選單
            initddlUser()
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 綁定新增的“人員，角色”信息
    ''' </summary>
    ''' <param name="sCheck">是否要默認選中下拉選單信息</param>
    ''' <remarks></remarks>
    Sub initAddDetail(ByVal sCheck As String)
        Try
            Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager())
            Dim syUserList As New AUTH_OP.SY_USERList(GetDatabaseManager())
            Dim sXmlData As String = String.Empty
            Dim XmlDocument As New XmlDocument

            ' 定義數據表接收新增的人員信息
            Dim dtRoleUser As New DataTable
            dtRoleUser.Columns.Add("STAFFID", Type.GetType("System.String"))
            dtRoleUser.Columns.Add("USERNAME", Type.GetType("System.String"))
            dtRoleUser.Columns.Add("BRANAME", Type.GetType("System.String"))
            dtRoleUser.Columns.Add("ROLEID", Type.GetType("System.String"))
            dtRoleUser.Columns.Add("BRA_DEPNO", Type.GetType("System.String"))
            dtRoleUser.Columns.Add("ROLENAME", Type.GetType("System.String"))

            ' 取得臨時檔資料
            sXmlData = getXmlData(syTempInfo)

            ' 若臨時檔有資料
            If sXmlData <> "" Then

                ' document對象載入XML文件
                XmlDocument.LoadXml(sXmlData)

                For i As Integer = 0 To XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER").Count - 1
                    Dim ROLEID As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("ROLEID"), System.Xml.XmlElement).InnerText
                    Dim STAFFID As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("STAFFID"), System.Xml.XmlElement).InnerText
                    Dim BRA_DEPNO As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("BRA_DEPNO"), System.Xml.XmlElement).InnerText
                    Dim OPERATION As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("OPERATION"), System.Xml.XmlElement).InnerText
                    Dim BRANAME As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("BRANAME"), System.Xml.XmlElement).InnerText
                    Dim USERNAME As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("USERNAME"), System.Xml.XmlElement).InnerText
                    Dim ROLENAME As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("ROLENAME"), System.Xml.XmlElement).InnerText

                    ' 如果角色為當前選中的維護角色
                    Dim row As DataRow = dtRoleUser.NewRow
                    row("STAFFID") = STAFFID
                    row("USERNAME") = USERNAME
                    row("BRANAME") = BRANAME
                    row("ROLEID") = ROLEID
                    row("BRA_DEPNO") = BRA_DEPNO
                    row("ROLENAME") = ROLENAME

                    dtRoleUser.Rows.Add(row)
                Next
            End If

            ' 選中單位，人員的默認選項
            If dtRoleUser.Rows.Count > 0 Then

                ' 若是第一次加載時。臨時擋中有資料
                If sCheck = "Y" Then
                    ddlDepart.SelectedItem.Selected = False
                    ddlUser.SelectedItem.Selected = False

                    rdolistRole.Items.FindByValue(dtRoleUser.Rows(0)("ROLEID").ToString()).Selected = True
                    ddlDepart.Items.FindByValue(dtRoleUser.Rows(0)("BRA_DEPNO").ToString()).Selected = True

                    ' 綁定使用者信息
                    If syUserList.loadByBRA_DEPNO(ddlDepart.SelectedItem.Value) Then
                        ddlUser.DataSource = syUserList.getCurrentDataSet
                        ddlUser.DataTextField = "USERNAME"
                        ddlUser.DataValueField = "BRA_DEPNO"
                        ddlUser.DataBind()

                        ddlUser.Items.FindByValue(dtRoleUser.Rows(0)("STAFFID").ToString() & ";" & dtRoleUser.Rows(0)("BRA_DEPNO").ToString()).Selected = True
                    End If

                    ddlUser.Items.Insert(0, New ListItem("--請選擇--", ""))
                End If
            End If

            ' 若有資料
            If dtRoleUser.Rows.Count > 0 Then

                trDetail.Visible = True

                ' 綁定人員信息清單
                dgAddDetail.DataSource = dtRoleUser
                dgAddDetail.DataBind()
            Else
                trDetail.Visible = False
                ddlDepart.SelectedIndex = -1
                ddlUser.SelectedIndex = -1
                rdolistRole.SelectedIndex = 0
            End If
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 初始化角色名稱清單
    ''' </summary>
    ''' <remarks></remarks>
    Sub initCheckData()
        Try
            Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager())
            Dim syFlowStep As New AUTH_OP.SY_FLOWSTEP(GetDatabaseManager())

            Dim sRoleNameFlag As String = String.Empty  ' 角色名稱
            Dim oldXmlData As String = String.Empty
            Dim XmlDocument As New XmlDocument

            ' 定義數據表接收角色名稱
            Dim dtRole As New DataTable

            dtRole.Columns.Add("ROLENAME", Type.GetType("System.String"))
            dtRole.Columns.Add("ROLEID", Type.GetType("System.String"))

            If syTempInfo.loadByCaseId(m_sCaseId) Then
                oldXmlData = syTempInfo.getAttribute("TEMPDATA")

                If oldXmlData <> "" Then
                    ' document對象載入XML文件
                    XmlDocument.LoadXml(oldXmlData)

                    For i As Integer = 0 To XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER").Count - 1
                        Dim sRID As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("ROLEID"), System.Xml.XmlElement).InnerText
                        Dim sSTAFFID As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("STAFFID"), System.Xml.XmlElement).InnerText
                        Dim sBRA_DEPNO As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("BRA_DEPNO"), System.Xml.XmlElement).InnerText
                        Dim sOPERATION As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("OPERATION"), System.Xml.XmlElement).InnerText
                        Dim sBRANAME As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("BRANAME"), System.Xml.XmlElement).InnerText
                        Dim sUSERNAME As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("USERNAME"), System.Xml.XmlElement).InnerText
                        Dim sROLENAME As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("ROLENAME"), System.Xml.XmlElement).InnerText

                        ' 如果當前的角色名稱和上一個名稱不同
                        If Not sRoleNameFlag.Contains(sROLENAME) Then
                            Dim row As DataRow = dtRole.NewRow
                            row("ROLENAME") = sROLENAME
                            row("ROLEID") = sRID

                            dtRole.Rows.Add(row)
                            sRoleNameFlag = sRoleNameFlag & ";" & sROLENAME
                        End If
                    Next
                End If
            End If

            ' 綁定角色名稱清單
            dgRoleDetail.DataSource = dtRole
            dgRoleDetail.DataBind()
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 初始化修改前後的人員信息
    ''' </summary>
    ''' <param name="sRoleID"></param>
    ''' <remarks></remarks>
    Sub initListBox(ByVal sRoleID As String)
        Try
            Dim syRelRoleUserList As New AUTH_OP.SY_REL_ROLE_USERList(GetDatabaseManager())
            Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager())
            Dim sXmlData As String = String.Empty
            Dim XmlDocument As New XmlDocument

            ' 清除修改前後的人員信息
            lstPreEdit.Items.Clear()
            lstAfterEdit.Items.Clear()

            ' 綁定"修改前"的人員信息
            If syRelRoleUserList.loadByRoleID(sRoleID) Then
                For i As Integer = 0 To syRelRoleUserList.getCurrentDataSet.Tables(0).Rows.Count - 1
                    lstPreEdit.Items.Add(New ListItem(syRelRoleUserList.getCurrentDataSet.Tables(0).Rows(i)("LstName"), syRelRoleUserList.getCurrentDataSet.Tables(0).Rows(i)("LstValue")))
                    lstAfterEdit.Items.Add(New ListItem(syRelRoleUserList.getCurrentDataSet.Tables(0).Rows(i)("LstName"), syRelRoleUserList.getCurrentDataSet.Tables(0).Rows(i)("LstValue")))
                Next
            End If

            ' 查詢臨時檔中的數據
            If syTempInfo.loadByCaseId(m_sCaseId) Then

                sXmlData = syTempInfo.getAttribute("TEMPDATA")

                If sXmlData <> "" Then

                    ' document對象載入XML文件
                    XmlDocument.LoadXml(sXmlData)

                    For i As Integer = 0 To XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER").Count - 1
                        Dim ROLEID As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("ROLEID"), System.Xml.XmlElement).InnerText
                        Dim STAFFID As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("STAFFID"), System.Xml.XmlElement).InnerText
                        Dim BRA_DEPNO As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("BRA_DEPNO"), System.Xml.XmlElement).InnerText
                        Dim OPERATION As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("OPERATION"), System.Xml.XmlElement).InnerText
                        Dim BRANAME As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("BRANAME"), System.Xml.XmlElement).InnerText
                        Dim USERNAME As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("USERNAME"), System.Xml.XmlElement).InnerText
                        Dim ROLENAME As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("ROLENAME"), System.Xml.XmlElement).InnerText

                        ' 如果是當前要查詢的角色編號
                        If ROLEID = sRoleID Then
                            Dim LstValue As String = BRA_DEPNO & ";" & STAFFID
                            Dim LstName As String = "(" & BRANAME & ") " & STAFFID & " " & USERNAME

                            ' 如果在修改后的人員信息裱中不存在當前人員信息
                            If IsNothing(lstAfterEdit.Items.FindByValue(LstValue)) Then
                                Dim item As ListItem = New ListItem(LstName, LstValue)
                                item.Attributes.Add("style", "color:red")
                                lstAfterEdit.Items.Add(item)
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
    ''' 新增資料到歷史檔
    ''' </summary>
    ''' <param name="sApproved">是否同意</param>
    ''' <param name="sCheck">是否點擊同意按鈕</param>
    ''' <param name="sFlag">是否走流程方法</param>
    ''' <remarks></remarks>
    Public Sub flowFuntion(ByVal sApproved As String, ByVal sCheck As String, ByRef stepinfo As StepInfo)
        ' 實例化
        Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager())
        Dim syRelRoleUser As New AUTH_OP.SY_REL_ROLE_USER(GetDatabaseManager())
        Dim XmlDocument As New XmlDocument

        Dim oldXmlData As String = String.Empty ' 聲明參數接收查詢出的XML

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
                Dim sBRANAME As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("BRANAME"), System.Xml.XmlElement).InnerText
                Dim sUSERNAME As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("USERNAME"), System.Xml.XmlElement).InnerText
                Dim sROLENAME As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("ROLENAME"), System.Xml.XmlElement).InnerText

                ' 同意
                If sCheck = "agree" Then

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
    ''' 點擊維護區的“新增”按鈕事件
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnAdd_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnAdd.Click
        Try
            ' 實例化
            Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager())
            Dim syRelRoelUserList As New AUTH_OP.SY_REL_ROLE_USERList(GetDatabaseManager())

            Dim sXmlData As String = String.Empty   ' 臨時檔資料
            Dim xmlDocument As New XmlDocument
            Dim item As XmlNode   ' Xml節點
            Dim newXmlData As String = String.Empty
            Dim oldXmlData As String = String.Empty
            Dim sbXmlData As New StringBuilder

            ' 定義Xml節點集合，接收xml節點
            Dim nodeList As IList(Of XmlNode) = New List(Of XmlNode)()

            ' 如果沒有選擇單位
            If ddlDepart.SelectedIndex = 0 Then
                com.Azion.EloanUtility.UIUtility.alert("請選擇單位！")

                Return
            End If

            ' 如果沒有選擇人員
            If ddlUser.SelectedIndex = 0 Then
                com.Azion.EloanUtility.UIUtility.alert("請選擇人員！")

                Return
            End If

            ' 如果當前新增人員已被設定
            If syRelRoelUserList.loadByStaffidDepnoROLEID(ddlUser.SelectedItem.Value.Split(";")(0), ddlDepart.SelectedItem.Value, rdolistRole.SelectedValue) Then
                com.Azion.EloanUtility.UIUtility.alert("所設定使用者已存在，請重新設定！")

                Return
            End If

            ' 如果當前人員沒有被設定,檢核臨時檔中是否已有當前人資料
            If syTempInfo.loadByPK(m_sWorkingUserid, m_sWorkingTopDepNo, m_sFuncCode) Then
                sXmlData = syTempInfo.getAttribute("TEMPDATA")

                ' 如果有人員信息
                If sXmlData <> "" Then

                    ' document對象載入XML文件
                    xmlDocument.LoadXml(sXmlData)

                    ' 找到臨時檔中已經存在的資料
                    For i As Integer = 0 To xmlDocument.GetElementsByTagName("SY_REL_ROLE_USER").Count - 1
                        If DirectCast(xmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("ROLEID"), System.Xml.XmlElement).InnerText = rdolistRole.SelectedValue _
                           And DirectCast(xmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("STAFFID"), System.Xml.XmlElement).InnerText = ddlUser.SelectedItem.Value.Split(";")(0) _
                                  And DirectCast(xmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("BRA_DEPNO"), System.Xml.XmlElement).InnerText = ddlDepart.SelectedItem.Value Then

                            com.Azion.EloanUtility.UIUtility.alert("所設定使用者已存在，請重新設定！")

                            Return
                        End If
                    Next
                End If
            End If

            ' 組XML文件
            sbXmlData.Append("<SY>")
            sbXmlData.Append("<SY_REL_ROLE_USER>")
            sbXmlData.Append("<ROLEID>" + rdolistRole.SelectedValue + "</ROLEID>")
            sbXmlData.Append("<STAFFID>" + ddlUser.SelectedValue.Split(";")(0) + "</STAFFID>")
            sbXmlData.Append("<BRA_DEPNO>" + ddlDepart.SelectedItem.Value + "</BRA_DEPNO>")
            sbXmlData.Append("<OPERATION>I</OPERATION>")
            sbXmlData.Append("<BRANAME>" + ddlDepart.SelectedItem.Text.Trim() + "</BRANAME>")
            sbXmlData.Append("<USERNAME>" + ddlUser.SelectedItem.Text.Split(" ")(1) + "</USERNAME>")
            sbXmlData.Append("<ROLENAME>" + rdolistRole.SelectedItem.Text + "</ROLENAME>")
            sbXmlData.Append("</SY_REL_ROLE_USER>")
            sbXmlData.Append("</SY>")

            ' 如果案件已在流程中
            If m_sFourStepNo = "" Then

                If syTempInfo.loadByPK(m_sWorkingUserid, m_sWorkingTopDepNo, m_sFuncCode) Then
                    sXmlData = syTempInfo.getAttribute("TEMPDATA")

                    ' 如果有內容
                    If sXmlData <> "" Then
                        ' document對象載入XML文件
                        xmlDocument.LoadXml(sXmlData)

                        newXmlData = sXmlData.Replace("</SY>", "") & sbXmlData.ToString().Replace("<SY>", "")
                    Else

                        ' 新增資料
                        newXmlData = sbXmlData.ToString()
                    End If
                Else

                    ' 新增異動數據到臨時表
                    syTempInfo.setAttribute("STAFFID", m_sWorkingUserid)
                    syTempInfo.setAttribute("BRA_DEPNO", m_sWorkingTopDepNo)
                    syTempInfo.setAttribute("FUNCCODE", m_sFuncCode)

                    newXmlData = sbXmlData.ToString()
                End If

                syTempInfo.setAttribute("TEMPDATA", newXmlData.ToString())
                syTempInfo.save()
            Else

                ' 根據案號查詢
                If syTempInfo.loadByCaseId(m_sCaseId) Then
                    sXmlData = syTempInfo.getAttribute("TEMPDATA")

                    ' 如果有內容
                    If sXmlData <> "" Then
                        ' document對象載入XML文件
                        xmlDocument.LoadXml(sXmlData)

                        newXmlData = sXmlData.Replace("</SY>", "") & sbXmlData.ToString().Replace("<SY>", "")
                    Else

                        ' 新增資料
                        newXmlData = sbXmlData.ToString()
                    End If
                End If

                syTempInfo.setAttribute("TEMPDATA", newXmlData.ToString())
                syTempInfo.save()
            End If

            com.Azion.EloanUtility.UIUtility.alert("儲存成功！")

            ' 重新綁定人員角色清單
            initAddDetail("")
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 點擊“確認送出"按鈕事件
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnSend_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnSend.Click
        Try
            ' 實例化
            Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager())
            Dim syRelRoelUserHis As New AUTH_OP.SY_REL_ROLE_USER_HIS(GetDatabaseManager())
            Dim flowFacade As New FLOW_OP.FlowFacade(GetDatabaseManager())
            Dim syRelRoleUser As New AUTH_OP.SY_REL_ROLE_USER(GetDatabaseManager())
            Dim stepInfo As New FLOW_OP.StepInfo()

            Dim XmlDocument As New XmlDocument
            Dim sXmlData As String = String.Empty

            ' 員工編號
            Dim sStaffid As String = String.Empty
            Dim sPreStaffid As String = String.Empty

            ' 維護角色的編號
            Dim sRoleID As String = "'" & rdolistRole.Items(0).Value.ToString() & "','" & rdolistRole.Items(1).Value.ToString() & "'"

            ' 取得臨時檔資料
            sXmlData = getXmlData(syTempInfo)

            ' 若有臨時資料
            If sXmlData <> "" Then

                ' document對象載入XML文件
                XmlDocument.LoadXml(sXmlData)

                If XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER").Count <> 0 Then
                    For i As Integer = 0 To XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER").Count - 1
                        Dim ROLEID As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("ROLEID"), System.Xml.XmlElement).InnerText
                        Dim STAFFID As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("STAFFID"), System.Xml.XmlElement).InnerText
                        Dim BRA_DEPNO As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("BRA_DEPNO"), System.Xml.XmlElement).InnerText
                        Dim OPERATION As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("OPERATION"), System.Xml.XmlElement).InnerText
                        Dim BRANAME As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("BRANAME"), System.Xml.XmlElement).InnerText
                        Dim USERNAME As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("USERNAME"), System.Xml.XmlElement).InnerText
                        Dim ROLENAME As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("ROLENAME"), System.Xml.XmlElement).InnerText

                        If ROLEID = rdolistRole.Items(0).Value Or ROLEID = rdolistRole.Items(1).Value Then
                            If sPreStaffid <> STAFFID Then
                                sStaffid = sStaffid & "'" & STAFFID & "',"
                                sPreStaffid = STAFFID
                            End If
                        End If
                    Next
                Else
                    com.Azion.EloanUtility.UIUtility.alert("沒有送簽資料，不能進行送出！")

                    Return
                End If
            Else
                com.Azion.EloanUtility.UIUtility.alert("沒有送簽資料，不能進行送出！")

                Return
            End If

            ' 檢核現有資料是否已在處理中
            If sStaffid.Length > 0 Then
                sStaffid = sStaffid.Substring(0, sStaffid.Length - 1)

                ' 若資料已在處理中
                If syRelRoelUserHis.loadByRoleIDStaffid(m_sStepNo, m_sCaseId, sRoleID, sStaffid) Then
                    com.Azion.EloanUtility.UIUtility.alert(rdolistRole.Items(0).Text & "已在處理中，無法送件！")

                    Return
                End If
            End If

            ' 如果是測試模式
            If m_bCheck Then
                ' 開始事務
                GetDatabaseManager().beginTran()

                stepInfo = MBSC.UICtl.UIShareFun.startFlow(GetDatabaseManager(), m_sFlowName, True, , True)

                flowFuntion("Y", "agree", stepInfo)

                '  如果案件没在流程中
                If m_sFourStepNo = "" Then

                    ' 依據主鍵刪除數據
                    syTempInfo.loadByPK(m_sWorkingUserid, m_sWorkingTopDepNo, m_sFuncCode)
                    syTempInfo.remove()
                Else

                    ' 刪除臨時檔的資料
                    syTempInfo.deleteByCaseID(m_sCaseId)
                End If

                ' 提交事務
                GetDatabaseManager().commit()

                ' 重新綁定數據
                initData()
            Else
                ' 開始事務
                GetDatabaseManager().beginTran()

                ' 若是起案
                If m_sStepNo = "" Then

                    ' 調用流程方法，取得案件編號
                    stepInfo = MBSC.UICtl.UIShareFun.startFlow(GetDatabaseManager(), m_sFlowName, True)

                    ' 更改SY_TEMPINFO裱中的CaseID欄位
                    If syTempInfo.loadByPK(m_sWorkingUserid, m_sWorkingTopDepNo, m_sFuncCode) Then
                        syTempInfo.setAttribute("CASEID", stepInfo.currentStepInfo.caseId.ToString())
                        syTempInfo.save()
                    End If
                Else

                    ' 調用流程方法
                    stepInfo = MBSC.UICtl.UIShareFun.pushFlow(GetDatabaseManager(), m_sCaseId)
                End If

                flowFuntion("", "", stepInfo)

                ' 提交事務
                GetDatabaseManager().commit()

                ' 跳轉到代辦清單
                com.Azion.EloanUtility.UIUtility.Redirect("SY_CASELIST.aspx")
            End If
        Catch ex As Exception

            ' 回滾事務
            GetDatabaseManager().Rollback()

            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 點擊“全部取消”按鈕事件
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnCancel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnCancel.Click
        Try
            ' 實例化
            Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager())
            Dim syRelRoleUserHis As New AUTH_OP.SY_REL_ROLE_USER_HIS(GetDatabaseManager())
            Dim flowFacade As New FLOW_OP.FlowFacade(GetDatabaseManager())

            ' 開始事務
            GetDatabaseManager().beginTran()

            ' 若資料沒有被送出
            If m_sFourStepNo = "" Then

                ' 如果臨時檔中有資料
                If syTempInfo.loadByPK(m_sWorkingUserid, m_sWorkingTopDepNo, m_sFuncCode) Then
                    syTempInfo.DeleteByPK(m_sWorkingUserid, m_sWorkingTopDepNo, m_sFuncCode)
                End If
            Else

                ' 依據案件編號刪除臨時檔資料
                syTempInfo.deleteByCaseID(m_sCaseId)
            End If

            ' 刪除歷史檔資料
            syRelRoleUserHis.deleteByCaseID(m_sCaseId)

            ' 刪除流程資料
            flowFacade.DeleteFlow(m_sCaseId)

            ' 提交事務
            GetDatabaseManager.commit()

            ddlDepart.Items.Clear()
            ddlUser.Items.Clear()

            ' 重新加載資料
            initData()

            ' 若設定角色沒初始值
            If rdolistRole.SelectedValue = "" Then
                rdolistRole.SelectedValue = "2"
            End If
        Catch ex As Exception

            ' 回滾事務
            GetDatabaseManager().Rollback()

            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 單位下拉選單的選擇項變化時事件
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub ddlDepart_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles ddlDepart.SelectedIndexChanged

        ' 清空人員下拉選單
        ddlUser.Items.Clear()

        ' 使用者下拉選單
        initddlUser()
    End Sub

    ''' <summary>
    ''' 點擊人員信息清單中的刪除按鈕的事件
    ''' </summary>
    ''' <param name="source"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub dgAddDetail_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles dgAddDetail.ItemCommand
        Try
            ' 實例化
            Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager())
            Dim XmlDocument As New XmlDocument
            Dim sXmlData As String = String.Empty
            Dim item As XmlNode

            ' 如果是刪除命令
            If hidOutAction.Value = "D" Then

                ' 得到當前行的索引
                Dim index As Integer = e.Item.ItemIndex

                Dim lblStaffid As Label = dgAddDetail.Items(index).FindControl("lblSTAFFID")
                Dim hidRoleID As HiddenField = dgAddDetail.Items(index).FindControl("hidROLEID")
                Dim hidBraDepno As HiddenField = dgAddDetail.Items(index).FindControl("hidBRA_DEPNO")

                ' 取得臨時檔資料
                sXmlData = getXmlData(syTempInfo)

                If sXmlData <> "" Then

                    ' document對象載入XML文件
                    XmlDocument.LoadXml(sXmlData)

                    For i As Integer = 0 To XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER").Count - 1
                        Dim sRID As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("ROLEID"), System.Xml.XmlElement).InnerText
                        Dim STAFFID As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("STAFFID"), System.Xml.XmlElement).InnerText
                        Dim sBRA_DEPNO As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("BRA_DEPNO"), System.Xml.XmlElement).InnerText
                        Dim sOPERATION As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("OPERATION"), System.Xml.XmlElement).InnerText
                        Dim sBRANAME As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("BRANAME"), System.Xml.XmlElement).InnerText
                        Dim sUSERNAME As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("USERNAME"), System.Xml.XmlElement).InnerText
                        Dim sROLENAME As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)("ROLENAME"), System.Xml.XmlElement).InnerText

                        ' 循環找到要刪除的數據
                        If lblStaffid.Text = STAFFID And hidBraDepno.Value = sBRA_DEPNO And hidRoleID.Value = sRID Then
                            item = XmlDocument.GetElementsByTagName("SY_REL_ROLE_USER")(i)
                            Exit For
                        End If
                    Next
                End If

                ' 如果找到將要刪除的數據
                If Not IsNothing(item) Then
                    sXmlData = sXmlData.Replace(item.OuterXml.ToString(), "")

                    If sXmlData.Equals("<SY></SY>") Then
                        sXmlData = ""
                    End If

                    If sXmlData = "" Then
                        syTempInfo.remove()
                    Else
                        syTempInfo.setAttribute("TEMPDATA", sXmlData.ToString())
                        syTempInfo.save()
                    End If

                    com.Azion.EloanUtility.UIUtility.alert("刪除成功！")

                    ' 若沒有資歷送出
                    If dgAddDetail.Items.Count = 0 Then
                        trDetail.Visible = False
                        ddlUser.SelectedIndex = -1
                        ddlDepart.SelectedIndex = -1
                    End If

                    ' 重新綁定人員信息表
                    initAddDetail("Y")

                    hidOutAction.Value = ""
                End If
            End If
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 設定角色選擇項變化時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub rdolistRole_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles rdolistRole.SelectedIndexChanged
        ddlDepart.SelectedIndex = -1
        ddlUser.SelectedIndex = -1
    End Sub
#End Region
#End Region

#Region "審核區"

#Region "Event"

    ''' <summary>
    ''' 點擊"同意"按鈕事件
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnAgree_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnAgree.Click
        Try
            ' 實例化
            Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager())
            Dim StepInfo As New FLOW_OP.StepInfo()

            ' 開始事務
            GetDatabaseManager().beginTran()

            ' 調用SendFlow方法 TO DO 方法報異常
            StepInfo = MBSC.UICtl.UIShareFun.pushFlow(GetDatabaseManager(), m_sCaseId)

            ' 對目前資料的流程進行存儲
            flowFuntion("Y", "agree", StepInfo)

            ' 刪除臨時檔的資料
            syTempInfo.deleteByCaseID(m_sCaseId)

            ' 提交事務
            GetDatabaseManager().commit()

            ' 關閉視窗
            com.Azion.EloanUtility.UIUtility.closeWindow()
        Catch ex As Exception

            ' 回滾事務
            GetDatabaseManager().Rollback()

            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 點擊"不同意"按鈕事件
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnDiffer_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnDiffer.Click
        Try
            ' 實例化
            Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager())
            Dim StepInfo As New FLOW_OP.StepInfo()

            ' 開始事務
            GetDatabaseManager().beginTran()

            ' 調用SendFlow方法 TO DO 該方法有異常
            StepInfo = MBSC.UICtl.UIShareFun.pushFlow(GetDatabaseManager(), m_sCaseId)

            ' 對目前資料的流程進行存儲
            flowFuntion("N", "", StepInfo)

            ' 刪除臨時檔資料
            syTempInfo.deleteByCaseID(m_sCaseId)

            ' 提交事務
            GetDatabaseManager().commit()

            ' 關閉視窗
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
    Protected Sub btnUpdate_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnUpdate.Click
        Try
            Dim StepInfo As New FLOW_OP.StepInfo()

            ' 開始事務
            GetDatabaseManager.beginTran()

            ' 調用JumpRollBack方法，TO DO 有異常
            StepInfo = MBSC.UICtl.UIShareFun.rollBack(GetDatabaseManager(), m_sCaseId)

            ' 對目前資料的流程進行存儲
            flowFuntion("", "", StepInfo)

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
    ''' 角色清單創建行是的事件
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub dgRoleDetail_ItemCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles dgRoleDetail.ItemCreated
        Try
            ' 當是數據行時或者交替行
            If e.Item.ItemType = ListItemType.Item Or e.Item.ItemType = ListItemType.AlternatingItem Then
                Dim lnkBtnRoleName As LinkButton = e.Item.FindControl("lnkBtnRoleName")
                AddHandler lnkBtnRoleName.Click, AddressOf lnkBtnRoleName_Click
            End If
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 點擊角色清單中的某個角色名稱時的事件
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
            Dim lnkBtnUserName As LinkButton = sender

            ' 取得單位編號和人員編號
            Dim hidRoleID As HiddenField = lnkBtnUserName.Parent.Parent.FindControl("hidRoleID")

            ' 綁定當前所點擊角色的人員信息
            initListBox(hidRoleID.Value)

            lblMsg.Text = lnkBtnUserName.Text
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub
#End Region
#End Region
End Class