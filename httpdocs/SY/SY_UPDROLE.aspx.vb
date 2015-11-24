Imports FLOW_OP.TABLE
Imports MBSC.UICtl

''' <summary>
''' 程式說明：角色建立
''' 建立者：Avril
''' 建立日期：2012-04-18
''' </summary>
Imports com.Azion.NET.VB
Imports AUTH_OP
Imports AUTH_OP.TABLE
Imports MBSC.MB_OP
Imports System.Xml
Imports System.IO
Imports FLOW_OP

Public Class SY_UPDROLE
    Inherits SYUIBase

    Dim m_sUpCode2366 As String = String.Empty
    Dim m_sOldCodeList As String = String.Empty
    Dim m_sTreeNode As TreeNode

#Region "Page Load"

    ''' <summary>
    ''' 頁面載入
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        m_bCheck = True     '不再複核
        ' 調用初始化參數方法
        initParas()

        ' 頁面第一次加載
        If Not IsPostBack Then
            ' Modify by 2012/06/28
            'btnConfirmSend.Visible = False '確認送出
            'btnCancelAll.Visible = False '全部取消

            ' 綁定資料
            bindData()

            ' 初始加載 顯示右邊維護區資料
            bindBank()

            ' 添加Onclick事件
            treeViewSys.Attributes.Add("onclick", "postBackByObject();")
            treeView.Nodes(0).Select()
            treeView_SelectedNodeChanged(Nothing, Nothing)
        Else
            If Not ViewState("LastNodePath") Is Nothing Then
                treeView.FindNode(ViewState("LastNodePath")).Select()
            End If

            If m_sFourStepNo <> "0400" Then

                If treeView.SelectedNode Is Nothing Then treeView.Nodes(0).Select()

                '名稱
                Dim iRoleId As Integer = CInt(treeView.SelectedNode.Value.Split(";")(0))
                'Dim iRoleId As Integer = lblID.Text
                If iRoleId > 0 Then
                    txtName.Visible = False
                    lblName.Visible = True
                Else
                    txtName.Visible = True
                    lblName.Visible = False
                End If

                treeView.Visible = True
                treeView.Enabled = True

                treeViewSys.Visible = True
                treeViewSys.Enabled = True

                trRoleMgr.Visible = True
                trSubSys.Visible = True '所屬子系統
                trDisable.Visible = True '啟用狀態

                '   btnAdd.Visible = True '新增
                btnSave.Visible = True '儲存
                btnDelete.Visible = False '刪除
                btnCancel.Visible = True '取消
            End If
        End If
    End Sub
#End Region

#Region "Function"

    ''' <summary>
    ''' 初始化參數
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function initParas()

        '#If _INDEV Then
        '        Dim a As New com.Azion.EloanUtility.StaffInfo
        '        a.WorkingStaffid = "S000035"
        '        a.LoginUserId = "S000035"
        '        Session("StaffInfo") = a
        '        FLOW_OP.ELoanFlow.m_callbackUserInfo = New ExportUserInfo
        '        m_bTesting = True
        '#End If

        ' 測試模式
        If m_bTesting OrElse Request("TESTMODE") = "1" Then
            m_bTesting = True

            Dim loginUser As New SY_LOGIN(GetDatabaseManager)
            loginUser.getUserInfo("904", "S000035")
            MyBase.InitParas()
            m_sHoFlag = 1
            'm_bCheck = True
        End If

        If Not String.IsNullOrEmpty(m_sCaseId) Then

            ' 案號
            lblCaseId.Text = m_sCaseId

            ' 單位
            Dim syBranch As New AUTH_OP.SY_BRANCH(GetDatabaseManager())
            If syBranch.loadByCaseId(m_sCaseId) Then
                lblBranch.Text = syBranch.getAttribute("BRCNAME")
            End If
        End If

        ' 取得CodeList中設定的參數值
        m_sUpCode2366 = com.Azion.EloanUtility.CodeList.getAppSettings("UpCode2366")
    End Function

    ''' <summary>
    ''' 綁定資料
    ''' </summary>
    ''' <remarks></remarks>
    Sub bindData()
        Try
            If m_sFourStepNo = "" Or m_sFourStepNo = "0300" Then

                ' 添加節點【存在于Live表】
                initRoleTreeView(treeView)

                ' 添加節點【存在于Temp表】
                initTempInfoTreeView(treeView)

                ' 綁定系統資料
                bindNormal()

                treeView.ExpandDepth = m_sTreeLevel

                ' 選擇“安泰銀行”
                treeView.Nodes(0).Select()

                divCheck.Visible = False
                divNormal.Visible = True
            Else

                ' 審核模塊資料加載
                bindCheck()
                initCheckData()

                divCheck.Visible = True
                divNormal.Visible = False

                divBranch.Visible = True
            End If
        Catch ex As Exception
            showErrMsg(ex)
        End Try
    End Sub

    ''' <summary>
    ''' 操作XML部份
    ''' </summary>
    ''' <param name="sFlag"></param>
    ''' <param name="sApproved"></param>
    ''' <param name="trNode"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function operationXml(ByVal sFlag As String, Optional sApproved As String = "", Optional trNode As TreeNode = Nothing) As String
        Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager())
        Dim syFlowStep As New AUTH_OP.SY_FLOWSTEP(GetDatabaseManager())
        Dim syRoleList As New AUTH_OP.SY_ROLEList(GetDatabaseManager())
        Dim xmlDocument As New XmlDocument
        Dim xmlData As String = String.Empty
        Dim sStepNo As String = String.Empty
        Dim sSubFlowSeq As String = String.Empty
        Dim sSubFlowCount As String = String.Empty
        Dim sCaseId As String = String.Empty
        Dim syRole As New AUTH_OP.SY_ROLE(GetDatabaseManager())
        Dim sParentNow As String = String.Empty

        ' 查詢臨時檔資料
        If m_sFourStepNo = "" Then
            If syTempInfo.loadByPK(m_sWorkingUserid, m_sWorkingTopDepNo, m_sFuncCode) Then
                xmlData = syTempInfo.getAttribute("TEMPDATA")
                sCaseId = syTempInfo.getAttribute("CASEID")
            End If
        Else
            If syTempInfo.loadByCaseId(m_sCaseId) Then
                xmlData = syTempInfo.getAttribute("TEMPDATA")
                sCaseId = m_sCaseId
            End If
        End If

        ' 根據案號查詢資料
        If syFlowStep.loadByCaseId(sCaseId) Then
            sStepNo = syFlowStep.getAttribute("STEP_NO")
            sSubFlowCount = syFlowStep.getAttribute("SUBFLOW_COUNT")
            sSubFlowSeq = syFlowStep.getAttribute("SUBFLOW_SEQ")
        End If

        If Not xmlData Is Nothing Then

            ' document對象載入XML文件
            If xmlData.Length > 0 Then
                xmlDocument.LoadXml(xmlData)
            End If
        End If

        For i As Integer = 0 To xmlDocument.GetElementsByTagName("SY_ROLE").Count - 1

            Dim sRoleId As String = DirectCast(xmlDocument.GetElementsByTagName("SY_ROLE")(i)("ROLEID"), System.Xml.XmlElement).InnerText
            Dim sParent As String = DirectCast(xmlDocument.GetElementsByTagName("SY_ROLE")(i)("PARENT"), System.Xml.XmlElement).InnerText
            Dim sRoleName As String = DirectCast(xmlDocument.GetElementsByTagName("SY_ROLE")(i)("ROLENAME"), System.Xml.XmlElement).InnerText
            Dim sDisabled As String = DirectCast(xmlDocument.GetElementsByTagName("SY_ROLE")(i)("DISABLED"), System.Xml.XmlElement).InnerText
            Dim sRoleType As String = DirectCast(xmlDocument.GetElementsByTagName("SY_ROLE")(i)("ROLETYPE"), System.Xml.XmlElement).InnerText
            Dim sSubSysId As String = DirectCast(xmlDocument.GetElementsByTagName("SY_ROLE")(i)("SUBSYSID"), System.Xml.XmlElement).InnerText
            Dim sSysIds As String = DirectCast(xmlDocument.GetElementsByTagName("SY_ROLE")(i)("SYSID"), System.Xml.XmlElement).InnerText

            Select Case sFlag
                Case "1"
                    ' 新增歷史檔部份
                    Dim sMaxRoleId As String = String.Empty
                    Dim sParentNew As String = String.Empty
                    Dim syRoleHis As New SY_ROLE_HIS(GetDatabaseManager())

                    If syRole.loadByRoleName(sRoleName) Then
                        sMaxRoleId = syRole.getAttribute("ROLEID").ToString()
                    End If

                    ' 如果此節點是其他節點的父節點 ，則需要更新子節點的parent=sMaxRoleId
                    sParentNow += sRoleId & "*" & sMaxRoleId & "|"

                    If Not syRoleHis.loadByPK(sRoleId, sCaseId, sStepNo, sSubFlowSeq, sSubFlowCount) Then
                        syRoleHis.setAttribute("ROLEID", sRoleId)
                        syRoleHis.setAttribute("CASEID", sCaseId)
                        syRoleHis.setAttribute("STEP_NO", sStepNo)
                        syRoleHis.setAttribute("SUBFLOW_SEQ", sSubFlowSeq)
                        syRoleHis.setAttribute("SUBFLOW_COUNT", sSubFlowCount)
                    End If

                    If String.IsNullOrEmpty(sParentNew) Then
                        For m As Integer = 0 To sParentNow.Split("|").Count - 1
                            If sParentNow.Split("|")(m).ToString.Contains(sParent) Then
                                sParentNew = sParentNow.Split("|")(m).Split("*")(1).ToString()
                            Else
                                If String.IsNullOrEmpty(sParentNew) Then
                                    sParentNew = sParent
                                End If
                            End If
                        Next
                    End If

                    If m_sFourStepNo = "" Or m_sFourStepNo = "0300" Then
                        syRoleHis.setAttribute("PARENT", sParent)
                    Else
                        If sParent = "0" Then
                            syRoleHis.setAttribute("PARENT", "0")
                        Else
                            syRoleHis.setAttribute("PARENT", sParentNew)
                        End If
                    End If

                    syRoleHis.setAttribute("ROLENAME", sRoleName)
                    syRoleHis.setAttribute("DISABLED", sDisabled)
                    syRoleHis.setAttribute("ROLETYPE", sRoleType)

                    If Not String.IsNullOrEmpty(sApproved) Then
                        syRoleHis.setAttribute("APPROVED", sApproved)
                    End If

                    syRoleHis.save()

                    ' 欄位賦值部份
                    If sSubSysId.Length > 0 Then
                        Dim sSubsys() As String = sSubSysId.Split(",")
                        Dim sSysIdsNew() As String = sSysIds.Split(",")

                        For j As Integer = 0 To sSubsys.Count - 1

                            Dim syRelSysId As New AUTH_OP.SY_REL_SYSID_SUBSYSID(GetDatabaseManager())
                            Dim syRoleSuitSysHis As New AUTH_OP.SY_ROLESUITSYS_HIS(GetDatabaseManager())
                            Dim sSysId As String = String.Empty

                            sSysId = sSysIdsNew(j).ToString()

                            ' 欄位賦值部份
                            If Not syRoleSuitSysHis.loadByPK(sRoleId, sCaseId, sStepNo, sSubFlowSeq, sSubFlowCount, sSubsys(j).ToString(), sSysId) Then
                                syRoleSuitSysHis.setAttribute("ROLEID", sRoleId)
                                syRoleSuitSysHis.setAttribute("CASEID", sCaseId)
                                syRoleSuitSysHis.setAttribute("STEP_NO", sStepNo)
                                syRoleSuitSysHis.setAttribute("SUBFLOW_SEQ", sSubFlowSeq)
                                syRoleSuitSysHis.setAttribute("SUBFLOW_COUNT", sSubFlowCount)
                                syRoleSuitSysHis.setAttribute("SUBSYSID", sSubsys(j).ToString())
                                syRoleSuitSysHis.setAttribute("SYSID", sSysId)
                            End If

                            syRoleSuitSysHis.save()
                        Next
                    End If

            End Select
        Next
    End Function
#End Region

#Region "Tree View op"

#Region "Role Tree"
    ''' <summary>
    ''' 給TreeView添加相關的節點
    ''' </summary>
    ''' <param name="treeView"></param>
    ''' <remarks></remarks>
    Sub initRoleTreeView(ByVal treeView As TreeView)
        Try
            ' 實例化
            Dim apCodelist As New AP_CODEList(GetDatabaseManager())
            Dim syRoleList As New AUTH_OP.SY_ROLEList(GetDatabaseManager())

            ' root節點
            Dim rootNode As TreeNode = New TreeNode()

            ' (1)綁定root節點
            If apCodelist.loadDataByCodeId(m_sUpCode2366) Then
                For Each dr As DataRow In apCodelist.getCurrentDataSet.Tables(0).Rows 'must be a record in ap_code
                    rootNode.Value = dr("VALUE").ToString() & ";0;0;0"
                    rootNode.Text = dr("TEXT").ToString
                Next
            End If

            ' (2)查詢角色資料的第一層資料：parent=0
            If Not syRoleList.genAllRoleList Is Nothing Then

                '將syRoleList存入ViewState
                ViewState("SY_Role") = syRoleList.getCurrentDataSet.Tables(0)

                ' (4)添加子節點
                addNodes(rootNode, 0, ViewState("SY_Role"))
            End If

            '添加到treeView中()
            treeView.Nodes.Add(rootNode)

        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' 添加子節點
    ''' </summary>
    ''' <param name="tNode">操作的父節點</param>
    ''' <param name="PId">父節點的編號</param>
    ''' <param name="dt">數據源</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function addNodes(ByRef tNode As TreeNode, ByVal PId As Integer, ByRef dt As DataTable) As String

        '定義DataRow承接DataTable篩選的結果
        Dim rows() As DataRow

        '定義篩選的條件
        Dim filterExpr As String
        filterExpr = "ParentId = " & PId

        '資料篩選並把結果傳入Rows
        rows = dt.Select(filterExpr)

        '如果篩選結果有資料
        If rows.GetUpperBound(0) >= 0 Then

            Dim childNode As TreeNode
            Dim rc As String

            '逐筆取出篩選後資料
            For Each row As DataRow In rows

                '實體化新節點
                childNode = New TreeNode

                '設定節點各屬性
                childNode.Text = row("ROLENAME").ToString & "  (" & row("ROLEID").ToString & ")"
                childNode.ToolTip = childNode.Text

                '子節點的Value組成=ROLEID;PARENTID;ROLETYPE;DISABLED
                childNode.Value = row("ROLEID").ToString() & ";" & row("PARENTID").ToString() & ";" & row("ROLETYPE").ToString() & ";" & row("DISABLED").ToString()

                '將節點加入Tree中
                tNode.ChildNodes.Add(childNode)

                '呼叫遞回取得子節點
                rc = addNodes(childNode, row("roleid"), dt)
            Next
        End If
    End Function

    ''' <summary>
    ''' 將臨時檔中的資料添加到左邊目錄樹
    ''' </summary>
    ''' <param name="treeView"></param>
    ''' <remarks></remarks>
    Sub initTempInfoTreeView(ByVal treeView As TreeView)
        Try

            Dim dt As DataTable = loadSYTempInfo()

            ' 將資料來源存在ViewState中  
            ViewState("SY_TEMPINFO") = dt

            If dt Is Nothing Then Return


            '<SY>
            '<SY_ROLE>
            '<ROLEID>11</ROLEID>
            '<PARENT>0</PARENT>
            '<ROLENAME>後台擔保品經辦(法金)</ROLENAME>
            '<DISABLED>0</DISABLED>
            '<ROLETYPE>2</ROLETYPE>
            '<SYSID>D,F,F</SYSID>
            '<SUBSYSID>04,05,06</SUBSYSID>
            '</SY_ROLE>
            '</SY>
            Dim sTempList As String = String.Empty

            ' (8) 循環取得sy_tempinfo中的資料
            For Each row As DataRow In dt.Rows
                Dim sRoleId As String = row("ROLEID")
                Dim sParent As String = row("PARENT")
                Dim sDISABLED As String = row("DISABLED")
                Dim sROLETYPE As String = row("ROLETYPE")
                Dim sSYSID As String = row("SYSID")
                Dim sSUBSYSID As String = row("SUBSYSID")
                Dim sRoleName As String = row("ROLENAME")

                Dim subNode As TreeNode = New TreeNode
                subNode.Text = String.Format("<font color='red'>{0}</font>", sRoleName & "  (" & sRoleId & ")")
                subNode.Value = sRoleId & ";" & sParent & ";" & sROLETYPE & ";" & sDISABLED

                Dim sParentValuePath As String = String.Empty

                ' 查找此節點的ValuePath '0;0;0;0/1;0;1;0
                sParentValuePath = findParentNodeValuePath(treeView, sParent)

                If Not treeView.FindNode(sParentValuePath & "/" & subNode.Value) Is Nothing Then
                    treeView.FindNode(sParentValuePath & "/" & subNode.Value).Text = String.Format("<font color='red'>{0}</font>", treeView.FindNode(sParentValuePath & "/" & subNode.Value).Text)
                    treeView.FindNode(sParentValuePath & "/" & subNode.Value).Selected = True
                    Continue For
                End If

                ' 根據ValuePath查找treeview中的節點
                Dim parent As TreeNode = treeView.FindNode(sParentValuePath)

                ' 如果此節點為Nothing 說明Treeview中無此節點 直接新增 否則給TreeView中parent節點新增子節點
                If parent Is Nothing Then

                    If sTempList.Length > 0 Then
                        If Not sTempList.Contains(sParentValuePath.Split(";")(0)) Then
                            treeView.Nodes(0).ChildNodes.Add(subNode)
                        End If
                    Else
                        treeView.Nodes(0).ChildNodes.Add(subNode)
                    End If
                Else

                    'If hidFlag.Value = "Modify" Then
                    '    If parent.ChildNodes.Count > 0 Then
                    '        For i = 0 To parent.ChildNodes.Count - 1
                    '            If parent.ChildNodes(i).Value.Split(";")(0) <> subNode.Value.Split(";")(0) Then
                    '                parent.ChildNodes.Add(subNode)
                    '            End If
                    '        Next

                    '        parent.ChildNodes.Add(subNode)
                    '    Else
                    '        parent.ChildNodes.Add(subNode)
                    '    End If
                    'Else
                    '    parent.ChildNodes.Add(subNode)
                    'End If

                    For Each node As TreeNode In parent.ChildNodes
                        sTempList &= node.Value.Split(";")(0).ToString() & "*"
                    Next

                    If Not sTempList.Contains(subNode.Value.Split(";")(0)) Then
                        parent.ChildNodes.Add(subNode)
                    End If

                    ' 查詢左邊目錄樹是否有資料，若有資料，則不需要新增，直接變紅即可。若沒有資料，則新增到父節點，然後變紅
                    'If parent.ChildNodes.Count > 0 Then
                    '    For Each node As TreeNode In parent.ChildNodes
                    '        If node.Value.Split(";")(0) <> subNode.Value.Split(";")(0) Then
                    '            parent.ChildNodes.Add(subNode)
                    '        End If
                    '    Next
                    'Else
                    '    parent.ChildNodes.Add(subNode)
                    'End If

                    For Each item As TreeNode In parent.ChildNodes
                        If item.Value.Split(";")(0) = subNode.Value.Split(";")(0) Then
                            item.Text = String.Format("<font color='red'>{0}</font>", item.Text)
                        Else
                        End If
                    Next
                End If
            Next
        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' 查找節點的ValuePath
    ''' </summary>
    ''' <param name="tree"></param>
    ''' <param name="iParentId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function findParentNodeValuePath(ByRef tree As TreeView, ByVal iParentId As Integer) As String
        Dim sValuePath As String = String.Empty

        ' 若Parent>0 說明資料來源是Live表
        If iParentId > 0 Then

            ' 查找存在于ViewState中Roleid為Parent的資料【父節點的資料來源是Live表】
            Dim row() As DataRow = CType(ViewState("SY_Role"), DataTable).Select("roleid=" & iParentId)

            If row.Length > 0 Then

                ' 查找父節點的ValuePath
                sValuePath = findParentNodeValuePath(tree, row(0)("PARENTID"))

                ' 組織此節點的ValuePath
                sValuePath &= "/" & row(0)("ROLEID") & ";" & row(0)("PARENTID") & ";" & row(0)("ROLETYPE") & ";" & row(0)("DISABLED")
            End If
        ElseIf iParentId < 0 Then

            ' 查找存在于ViewState中Roleid為Parent的資料【父節點的資料來源是Temp表】
            Dim row() As DataRow = CType(ViewState("SY_TEMPINFO"), DataTable).Select("roleid='" & iParentId & "'")

            If row.Length > 0 Then

                ' 查找父節點的ValuePath
                sValuePath = findParentNodeValuePath(tree, row(0)("PARENT"))

                ' 組織此節點的ValuePath
                sValuePath &= "/" & row(0)("ROLEID") & ";" & row(0)("PARENT") & ";" & row(0)("ROLETYPE") & ";" & row(0)("DISABLED")
            End If
        Else
            ' 如果Parent為0，為父節點，直接取其ValuePath
            'ROLEID;PARENTID;ROLETYPE;DISABLED
            sValuePath = tree.Nodes(0).ValuePath
        End If

        Return sValuePath
    End Function

#End Region

#Region "SYS Tree"
    ''' <summary>
    ''' 綁定系統資料
    ''' </summary>
    ''' <remarks></remarks>
    Sub bindNormal()
        Dim sySysIdList As New AUTH_OP.SY_SYSIDList(GetDatabaseManager())
        Dim dtSysData As New DataTable

        ' (9) 查詢第一層的數據資料
        If sySysIdList.loadUseData() Then
            dtSysData = sySysIdList.getCurrentDataSet.Tables(0)
        End If

        ' (10)循環 新增到TreeView中
        If Not dtSysData Is Nothing Then

            Dim sySubSysIdList As New AUTH_OP.SY_SUBSYSIDList(GetDatabaseManager())
            Dim dtData As New DataTable

            If sySubSysIdList.loadAllSubSys Then
                dtData = sySubSysIdList.getCurrentDataSet.Tables(0)
            End If

            For Each dr As DataRow In dtSysData.Rows

                ' 父節點
                Dim parentNode As TreeNode = New TreeNode()
                parentNode.Value = dr("SYSID").ToString()
                parentNode.Text = dr("SYSNAME").ToString

                ' 不可點擊
                parentNode.SelectAction = TreeNodeSelectAction.None

                ' (4)添加子節點
                addSysNode(parentNode, dr("SYSID").ToString(), dtData)
                treeViewSys.Nodes.Add(parentNode)
            Next
        End If
    End Sub

    ''' <summary>
    ''' 新增TreeView節點
    ''' </summary>
    ''' <param name="node">節點</param>
    ''' <param name="item">parent</param>
    ''' <remarks></remarks>
    Private Function addSysNode(ByVal tNode As TreeNode, ByVal sSysId As String, ByVal dt As DataTable) As String
        Try
            '定義DataRow承接DataTable篩選的結果
            Dim rows() As DataRow

            '定義篩選的條件
            Dim filterExpr As String
            filterExpr = "SYSID = '" & sSysId & "'"

            '資料篩選並把結果傳入Rows
            rows = dt.Select(filterExpr)

            If rows.Length > 0 Then

                Dim childNode As TreeNode
                Dim rc As String

                For Each row As DataRow In rows

                    childNode = New TreeNode

                    childNode.Text = row("SHORTSSNAME").ToString()
                    childNode.Value = row("SUBSYSID").ToString

                    ' 不可點擊
                    childNode.SelectAction = TreeNodeSelectAction.None

                    tNode.ChildNodes.Add(childNode)
                Next
            End If

        Catch ex As Exception
            showErrMsg(ex)
        End Try
    End Function
#End Region

    Sub selectRootNode(ByVal trNode As TreeNode)
        If trNode.Parent Is Nothing Then
            '上層項目
            lblParentName.Text = ""

            '編號 
            lblID.Text = trNode.Value.Split(";")(0)

            '名稱
            lblName.Text = trNode.Text.Replace("<font color='red'>", "").Replace("</font>", "")

            '所屬子系統
            trRoleMgr.Visible = False
            trSubSys.Visible = False

            '啟用狀態
            trDisable.Visible = False
        End If
    End Sub

    ''' <summary>
    ''' reset treeviewsys 
    ''' </summary>
    ''' <remarks></remarks>
    Sub resetTreeViewSys()
        ' 清空之前選中的資料
        ' 將checkBox全部恢復成不可用
        For Each nodes As TreeNode In treeViewSys.Nodes
            nodes.Checked = False
            nodes.ShowCheckBox = False
            nodes.Text = nodes.Text.Replace("<font color='red'>", "").Replace("</font>", "")
            nodes.Text = nodes.Text.Replace("<input type='checkbox' disabled = 'disabled' />", "")
            nodes.Text = String.Format("<input type='checkbox' disabled = 'disabled' />{0}", nodes.Text)

            For Each childNode As TreeNode In nodes.ChildNodes
                childNode.Checked = False
                childNode.ShowCheckBox = False
                childNode.Text = childNode.Text.Replace("<font color='red'>", "").Replace("</font>", "")
                childNode.Text = childNode.Text.Replace("<input type='checkbox' disabled = 'disabled' />", "")
                childNode.Text = String.Format("<input type='checkbox' disabled = 'disabled' />{0}", childNode.Text)
            Next
        Next
    End Sub

    Function resetTreeViewSys(ByRef nodes As TreeNodeCollection) As String
        For Each node As TreeNode In nodes
            node.Checked = False
            node.ShowCheckBox = False
            node.Text = node.Text.Replace("<font color='red'>", "").Replace("</font>", "")
            node.Text = node.Text.Replace("<input type='checkbox' disabled = 'disabled' />", "")
            node.Text = String.Format("<input type='checkbox' disabled = 'disabled' />{0}", node.Text)
            resetTreeViewSys(node.ChildNodes)
        Next
    End Function

    '上層有勾選的下層才可勾選
    Sub initSyRoleSuitSys(ByVal iParentId As Integer)

        '全部都可勾選
        If iParentId = 0 Then
            For Each node As TreeNode In treeViewSys.Nodes
                node.ShowCheckBox = True
                node.Checked = False
                node.Text = node.Text.Replace("<input type='checkbox' disabled = 'disabled' />", "")
                For Each childnode As TreeNode In node.ChildNodes
                    If Not childnode.ShowCheckBox Then
                        childnode.ShowCheckBox = True
                        childnode.Checked = False
                        childnode.Text = childnode.Text.Replace("<input type='checkbox' disabled = 'disabled' />", "")
                    End If
                Next
            Next
        Else

            'temp 有以temp主，temp 沒有以live主
            Dim row As DataRow = getSYTempInfoRow(iParentId)
            If Not row Is Nothing Then '由Temp
                Dim sSysIds As String = row("SYSID").ToString
                Dim sSubSysIds As String = row("SUBSYSID").ToString

                For i As Integer = 0 To sSysIds.Split(",").Length - 1
                    Dim sSysId As String = sSysIds.Split(",")(i)
                    Dim sSubSysId As String = sSubSysIds.Split(",")(i)
                    Dim node As TreeNode = treeViewSys.FindNode(sSysId & "/" & sSubSysId)
                    If Not node Is Nothing Then
                        node.ShowCheckBox = True
                        node.Checked = False
                        node.Text = node.Text.Replace("<input type='checkbox' disabled = 'disabled' />", "")
                        If Not node.Parent Is Nothing Then
                            node.Parent.ShowCheckBox = True
                            node.Parent.Checked = False
                            node.Parent.Text = node.Parent.Text.Replace("<input type='checkbox' disabled = 'disabled' />", "")
                        End If
                    End If
                Next
            Else '由live
                Dim syRoleSuitSysList As New AUTH_OP.SY_ROLESUITSYSList(GetDatabaseManager())
                If syRoleSuitSysList.loadDataByRoleId(iParentId) > 0 Then
                    Dim dtData As DataTable = syRoleSuitSysList.getCurrentDataSet.Tables(0)
                    For Each row In dtData.Rows
                        Dim sSysId As String = row("SysId").ToString
                        Dim sSubSysId As String = row("SubSysId").ToString
                        Dim node As TreeNode = treeViewSys.FindNode(sSysId & "/" & sSubSysId)
                        If Not node Is Nothing Then
                            node.ShowCheckBox = True
                            node.Checked = False
                            node.Text = node.Text.Replace("<input type='checkbox' disabled = 'disabled' />", "")
                            If Not node.Parent Is Nothing Then
                                node.Parent.ShowCheckBox = True
                                node.Parent.Checked = False
                                node.Parent.Text = node.Parent.Text.Replace("<input type='checkbox' disabled = 'disabled' />", "")
                            End If
                        End If
                    Next
                End If
            End If

        End If
    End Sub

#End Region

#Region "load data"

    Function loadSYTempInfo() As DataTable

        If Not ViewState("SY_TEMPINFO") Is Nothing Then
            Return CType(ViewState("SY_TEMPINFO"), DataTable)
        End If

        Dim oldXmlData As String = String.Empty
        Dim sCaseid As String = String.Empty

        Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager())

        If m_sFourStepNo = "" Then
            If syTempInfo.loadByPK(m_sWorkingUserid, m_sWorkingTopDepNo, m_sFuncCode) Then
                btnCancelAll.Visible = True '全部取消
                btnConfirmSend.Visible = True '確認送出
                ' 如果是Insert資料，則需要進行追加 如果是修改 則修改
                oldXmlData = syTempInfo.getAttribute("TEMPDATA").ToString()
                'sCaseid = syTempInfo.getString("CASEID")
            End If
        Else
            If syTempInfo.loadByCaseId(m_sCaseId) Then
                ' 如果是Insert資料，則需要進行追加 如果是修改 則修改
                oldXmlData = syTempInfo.getAttribute("TEMPDATA").ToString()
                'sCaseid = syTempInfo.getString("CASEID")
            End If
        End If

        '有案號
        If m_sCaseId = "" AndAlso m_sStepNo = "" Then '從左側進來
            m_sCaseId = sCaseid
            If m_sCaseId <> "" Then
                m_sStepNo = MBSC.UICtl.UIShareFun.getCurrentStepInfo(GetDatabaseManager(), m_sCaseId).stepNo
                m_sFourStepNo = Right(m_sStepNo, 4)
                ViewState("currentFourStepNo") = m_sFourStepNo
            End If
        Else
            ViewState("currentFourStepNo") = m_sFourStepNo
        End If

        If oldXmlData <> Nothing Then
            Dim dt As DataTable = com.Azion.EloanUtility.XmlUtility.convertXML2DataTable(oldXmlData)
            If Not dt Is Nothing Then
                Dim dv As System.Data.DataView = dt.DefaultView
                dv.Sort = "Roleid asc"
                Return dv.ToTable
            End If
        End If
        Return Nothing
    End Function

    Function getSYTempInfoRow(ByVal sRoleId As String) As DataRow
        Dim dt As DataTable = CType(ViewState("SY_TEMPINFO"), DataTable)   'loadSYTempInfo()

        If Not dt Is Nothing Then
            If dt.Select("roleid='" & sRoleId & "'").Length > 0 Then Return dt.Select("roleid='" & sRoleId & "'")(0)
        End If

        Return Nothing
    End Function

    Function bindTreeViewSYSForSYTempInfo(ByVal iRoleId As Integer) As Boolean
        Dim row As DataRow = getSYTempInfoRow(iRoleId)
        Dim bHasTempData As Boolean = False

        If row Is Nothing Then Return bHasTempData

        Dim sSysIds As String = row("SYSID").ToString
        Dim sSubSysIds As String = row("SUBSYSID").ToString

        For i As Integer = 0 To sSysIds.Split(",").Length - 1
            Dim sSysId As String = sSysIds.Split(",")(i)
            Dim sSubSysId As String = sSubSysIds.Split(",")(i)

            Dim node As TreeNode = treeViewSys.FindNode(sSysId & "/" & sSubSysId)

            If Not node Is Nothing Then
                Dim sDisabled As String = row("Disabled").ToString
                rdoDisable.SelectedValue = sDisabled
                If node.ShowCheckBox Then
                    'node.ShowCheckBox = True'不可在這設定，因為initSyRoleSuitSys已經將下層可用的設為ShowCheckBox = True
                    node.Checked = True
                    node.Text = node.Text.Replace("<input type='checkbox' disabled = 'disabled' />", "")
                    node.Text = String.Format("<font color='red'>{0}</font>", node.Text)
                    If Not node.Parent Is Nothing Then
                        'node.Parent.ShowCheckBox = True'不可在這設定，因為initSyRoleSuitSys已經將下層可用的設為ShowCheckBox = True
                        node.Parent.Checked = True
                        node.Parent.Text = node.Parent.Text.Replace("<input type='checkbox' disabled = 'disabled' />", "")
                        node.Parent.Text = String.Format("<font color='red'>{0}</font>", node.Parent.Text)
                    End If
                    bHasTempData = True
                End If
            End If
        Next

        Return bHasTempData
    End Function

    Function loadSYRoleSuitSys(ByVal iRoleId As Integer) As DataTable
        Dim syRoleSuitSysList As New AUTH_OP.SY_ROLESUITSYSList(GetDatabaseManager())
        'loadDataByRoleId 一定要 order by SYSID,SUBSYSID 否則產生的XML順序會不同
        If syRoleSuitSysList.loadDataByRoleId(iRoleId) > 0 Then
            Return syRoleSuitSysList.getCurrentDataSet.Tables(0)
        End If
        Return Nothing
    End Function

    Sub bindTreeViewSYSForSYRoleSuitSys(ByVal iRoleId As Integer, ByRef bHasTempData As Boolean)

        Dim dtData As DataTable = loadSYRoleSuitSys(iRoleId)

        If dtData Is Nothing Then
            Return
        End If

        'loadDataByRoleId 一定要 order by SYSID,SUBSYSID 否則產生的XML順序會不同
        If dtData.Rows.Count > 0 Then

            Dim sSysIds As String = String.Empty
            Dim sSubSysIds As String = String.Empty

            'iRoleId<=10代表系統預建角色不可更改
            If iRoleId <= 10 Then
                btnAdd.Visible = False '新增
                btnSave.Visible = False '儲存
                btnDelete.Visible = False '刪除
                btnCancel.Visible = False '取消
                treeViewSys.Enabled = False
            End If

            For Each row As DataRow In dtData.Rows
                Dim sSysId As String = row("SysId").ToString
                Dim sSubSysId As String = row("SubSysId").ToString
                sSysIds &= row("SysId").ToString & ","
                sSubSysIds &= row("SubSysId").ToString & ","
                Dim node As TreeNode = treeViewSys.FindNode(sSysId & "/" & sSubSysId)

                If Not node Is Nothing Then
                    Dim sDisabled As String = row("Disabled").ToString
                    rdoDisable.SelectedValue = sDisabled
                    If node.ShowCheckBox Then
                        'node.ShowCheckBox = True '不可在這設定，因為initSyRoleSuitSys已經將下層可用的設為ShowCheckBox = True
                        If Not bHasTempData Then node.Checked = True 'temp 有資料，以temp資料為主
                        node.Text = node.Text.Replace("<input type='checkbox' disabled = 'disabled' />", "")
                        node.Text = node.Text.Replace("<font color='red'>", "").Replace("</font>", "")
                        If Not node.Checked Then
                            node.Text = String.Format("<font color='red'>{0}</font>", node.Text)
                            If Not node.Parent Is Nothing Then
                                If Not node.Parent.Checked Then
                                    bHasTempData = True
                                    node.Parent.Text = String.Format("<font color='red'>{0}</font>", node.Parent.Text)
                                End If
                            End If
                        Else
                            If Not node.Parent Is Nothing Then
                                'node.Parent.ShowCheckBox = True'不可在這設定，因為initSyRoleSuitSys已經將下層可用的設為ShowCheckBox = True
                                If Not bHasTempData Then node.Parent.Checked = True 'temp 有資料，以temp資料為主
                                node.Parent.Text = node.Parent.Text.Replace("<input type='checkbox' disabled = 'disabled' />", "")
                                node.Parent.Text = node.Parent.Text.Replace("<font color='red'>", "").Replace("</font>", "")
                            End If
                        End If
                    End If
                End If
            Next
        End If

        '產生一份live 的XML
        genLiveTreeViewSYS2XML(dtData)
    End Sub

    ''' <summary>
    ''' <SY>
    ''' <SY_ROLE>
    ''' <ROLEID>11</ROLEID>
    ''' <PARENT>0</PARENT>
    ''' <ROLENAME>後台擔保品經辦(法金)</ROLENAME>
    ''' <DISABLED>0</DISABLED>
    ''' <ROLETYPE>2</ROLETYPE>
    ''' <SYSID>D,F,F</SYSID>
    ''' <SUBSYSID>04,05,06</SUBSYSID>
    ''' </SY_ROLE>
    ''' </SY> 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function genTempInfoTreeViewSYS2XML() As DataTable
        Dim iRoleId As Integer = lblID.Text
        Dim sParent As String = String.Empty

        ' Modify by 2012/06/28
        'If iRoleId = "-1" Then
        '    display()
        '    bindBank()
        'Else
        Dim ds As New DataSet

        Dim xmlDocument As New XmlDocument
        Dim sbXmlData As New Text.StringBuilder
        Dim sSysId As String
        Dim sSubSysId As String

        sbXmlData.Append("<SY>")
        sbXmlData.Append("<SY_ROLE>")
        sbXmlData.Append("<ROLEID>" & iRoleId & "</ROLEID>")

        If treeView.SelectedNode.Parent Is Nothing Then
            'sbXmlData.Append("<PARENT>0</PARENT>")
            ' Modify by 2012/06/28
            If hidId.Value <> "" Then
                If lblParentName.Text = "安泰銀行" Then
                    sbXmlData.Append("<PARENT>0</PARENT>")
                    sParent = "0"
                Else
                    sbXmlData.Append("<PARENT>" & hidId.Value & "</PARENT>")
                    sParent = hidId.Value
                End If
            Else
                sbXmlData.Append("<PARENT>0</PARENT>")
                sParent = "0"
            End If
        Else
            If hidFlag.Value = "Modify" Then
                sbXmlData.Append("<PARENT>" & treeView.SelectedNode.Value.Split(";")(1) & "</PARENT>")
                sParent = treeView.SelectedNode.Value.Split(";")(1)
            Else
                sbXmlData.Append("<PARENT>" & treeView.SelectedNode.Value.Split(";")(0) & "</PARENT>")
                sParent = treeView.SelectedNode.Value.Split(";")(0)
            End If
        End If

        sbXmlData.Append("<ROLEMGR>" & tbRoleMgr.Text.Trim & "</ROLEMGR>")


        sbXmlData.Append("<ROLENAME>" & txtName.Text & "</ROLENAME>")
        sbXmlData.Append("<DISABLED>" & rdoDisable.SelectedValue & "</DISABLED>")
        sbXmlData.Append("<ROLETYPE>2</ROLETYPE>")

        For Each node As TreeNode In treeViewSys.CheckedNodes
            If node.ChildNodes.Count > 0 Then
                For Each childNode As TreeNode In node.ChildNodes
                    If childNode.Checked Then
                        sSysId &= node.Value & ","
                        sSubSysId &= childNode.Value & ","
                    End If
                Next
            End If
        Next

        ' 所屬子系統的父節點
        If sSysId.Length > 0 Then
            sbXmlData.Append("<SYSID>" & sSysId.Substring(0, sSysId.Length - 1) & "</SYSID>")
        Else
            sbXmlData.Append("<SYSID></SYSID>")
        End If

        ' 所屬子系統的子節點
        If sSubSysId.Length > 0 Then
            sbXmlData.Append("<SUBSYSID>" & sSubSysId.Substring(0, sSubSysId.Length - 1) & "</SUBSYSID>")
        Else
            sbXmlData.Append("<SUBSYSID></SUBSYSID>")
        End If

        sbXmlData.Append("</SY_ROLE>")

        sbXmlData.Append("</SY>")

        hidValue.Value = iRoleId.ToString & ";" & sParent & ";" & "2" & ";" & rdoDisable.SelectedValue

        '與live的XML做比較一樣則取代
        If Not ViewState("LiveTreeViewSYS2XML") Is Nothing Then
            sbXmlData = sbXmlData.Replace(ViewState("LiveTreeViewSYS2XML").ToString, "")
            ViewState.Remove("LiveTreeViewSYS2XML") '用完清掉
        End If

        xmlDocument.LoadXml(sbXmlData.ToString)
        ds.ReadXml(New XmlNodeReader(xmlDocument))

        If ds.Tables.Count > 0 Then Return ds.Tables(0)

        Return Nothing
    End Function

    Sub genLiveTreeViewSYS2XML(ByVal dt As DataTable)
        ViewState.Remove("LiveTreeViewSYS2XML")

        If dt Is Nothing Then Return

        Dim sbXmlData As New Text.StringBuilder
        Dim sSysId As String
        Dim sSubSysId As String

        sbXmlData.Append("<SY_ROLE>")
        sbXmlData.Append("<ROLEID>" & dt.Rows(0)("ROLEID") & "</ROLEID>")
        sbXmlData.Append("<PARENT>" & dt.Rows(0)("PARENT") & "</PARENT>")
        sbXmlData.Append("<ROLENAME>" & dt.Rows(0)("ROLENAME") & "</ROLENAME>")
        sbXmlData.Append("<DISABLED>" & dt.Rows(0)("DISABLED") & "</DISABLED>")
        sbXmlData.Append("<ROLETYPE>" & dt.Rows(0)("ROLETYPE") & "</ROLETYPE>")

        For Each row As DataRow In dt.Rows
            sSysId &= row("SysId").ToString & ","
            sSubSysId &= row("SubSysId").ToString & ","
        Next

        ' 所屬子系統的父節點
        sbXmlData.Append("<SYSID>" & sSysId.Substring(0, sSysId.Length - 1) & "</SYSID>")

        ' 所屬子系統的子節點
        sbXmlData.Append("<SUBSYSID>" & sSubSysId.Substring(0, sSubSysId.Length - 1) & "</SUBSYSID>")

        sbXmlData.Append("</SY_ROLE>")

        ViewState("LiveTreeViewSYS2XML") = sbXmlData.ToString
    End Sub

    'insert SY_ROLE_HIS
    '寫法一
    Sub insertSyRoleHisXXX(ByVal databaseManager As DatabaseManager, ByVal syRole As AUTH_OP.SY_ROLE, ByVal sOperation As String, Optional sAPPROVED As String = "Y", _
                  Optional ByVal sCaseId As String = "", Optional ByVal sStepNo As String = "00000000", _
                  Optional ByVal iSubFlowSeq As Integer = 0, Optional ByVal iSubFlowCount As Integer = 0)

        Dim syRoleHis As New AUTH_OP.SY_ROLE_HIS(databaseManager)


        syRoleHis.loadByPK(syRole.getInt("ROLEID"), sCaseId, sStepNo, iSubFlowSeq, iSubFlowCount)

        For i As Integer = 0 To syRole.getBosAttributes.size - 1
            Dim bosAttribute As BosAttribute = syRole.getBosAttributes.item(i)
            Dim sFieldName As String = bosAttribute.getColName()
            syRoleHis.setAttribute(sFieldName, syRole.getAttribute(sFieldName))
        Next

        syRoleHis.setAttribute("CASEID", sCaseId)
        syRoleHis.setAttribute("STEP_NO", sStepNo)
        syRoleHis.setAttribute("SUBFLOW_SEQ", iSubFlowSeq)
        syRoleHis.setAttribute("SUBFLOW_COUNT", iSubFlowCount)
        syRoleHis.setAttribute("APPROVED", sAPPROVED)
        syRoleHis.setAttribute("OPERATION", sOperation)
        syRoleHis.save()
    End Sub

    '寫法二
    Sub insertSyRoleHis(ByVal databaseManager As DatabaseManager, ByVal row As DataRow, ByVal sOperation As String, Optional sAPPROVED As String = "Y", _
                  Optional ByVal sCaseId As String = "SY0000000000000", Optional ByVal sStepNo As String = "00000000", _
                  Optional ByVal iSubFlowSeq As Integer = 0, Optional ByVal iSubFlowCount As Integer = 0)

        Dim syRoleHis As New AUTH_OP.SY_ROLE_HIS(databaseManager)

        If row("ROLEID") < 0 Then
            sOperation = "I"
        Else
            sOperation = "N"
        End If

        syRoleHis.loadByPK(row("ROLEID"), sCaseId, sStepNo, iSubFlowSeq, iSubFlowCount)

        '因為欄位相同
        For Each column As DataColumn In row.Table.Columns
            syRoleHis.setAttribute(column.ColumnName.ToUpper, row(column.ColumnName))
        Next

        syRoleHis.setAttribute("CASEID", sCaseId) ' m_sCaseId
        syRoleHis.setAttribute("STEP_NO", sStepNo)
        syRoleHis.setAttribute("SUBFLOW_SEQ", iSubFlowSeq)
        syRoleHis.setAttribute("SUBFLOW_COUNT", iSubFlowCount)
        syRoleHis.setAttribute("APPROVED", sAPPROVED)
        syRoleHis.setAttribute("OPERATION", sOperation)
        syRoleHis.save()

        ' 如果父節點的啟用狀態改了 ，則子節點也需要新增歷史檔
        Dim syRoleList As New AUTH_OP.SY_ROLEList(GetDatabaseManager)

        ' 如果目前只是修改的父節點，則需要根據其roleid查找是否有子節點，將其disable設置為與
        If syRoleList.loadDataByType(row("ROLEID").ToString) Then

            For Each item As AUTH_OP.SY_ROLE In syRoleList
                iSubFlowSeq = iSubFlowSeq + 1
                insertSyRoleHisXXX(databaseManager, item, sOperation, sAPPROVED, sCaseId, sStepNo, iSubFlowSeq, iSubFlowCount)
            Next
        End If
    End Sub
    '
    Sub insertSyRoleSuitSysHis(ByVal databaseManager As DatabaseManager, ByVal row As DataRow, ByVal sOperation As String, _
                  Optional ByVal sCaseId As String = "SY0000000000000", Optional ByVal sStepNo As String = "00000000", _
                  Optional ByVal iSubFlowSeq As Integer = 0, Optional ByVal iSubFlowCount As Integer = 0, Optional ByVal sSubSysId As String = "", Optional ByVal sSysId As String = "")

        Dim syRoleSuitSysHis As New AUTH_OP.SY_ROLESUITSYS_HIS(databaseManager)

        '歷史檔一定新增，不會更新
        'Dim iTempSubFlowSeq As Integer = 0
        'iTempSubFlowSeq = syRoleSuitSysHis.getMaxSubFlowSeq(row("ROLEID").ToString(), sStepNo, sCaseId, iSubFlowCount, sSubSysId, sSysId) + 1

        syRoleSuitSysHis.loadByPK(syRoleSuitSysHis.getInt("ROLEID"), sCaseId, sStepNo, iSubFlowSeq, iSubFlowCount, sSubSysId, sSysId)

        For Each column As DataColumn In row.Table.Columns
            syRoleSuitSysHis.setAttribute(column.ColumnName.ToUpper, row(column.ColumnName))
        Next

        syRoleSuitSysHis.setAttribute("CASEID", sCaseId) ' m_sCaseId
        syRoleSuitSysHis.setAttribute("STEP_NO", sStepNo)
        syRoleSuitSysHis.setAttribute("SUBFLOW_SEQ", iSubFlowSeq)
        syRoleSuitSysHis.setAttribute("SUBFLOW_COUNT", iSubFlowCount)
        syRoleSuitSysHis.setAttribute("SUBSYSID", sSubSysId)
        syRoleSuitSysHis.setAttribute("SYSID", sSysId)
        syRoleSuitSysHis.setAttribute("OPERATION", sOperation)
        syRoleSuitSysHis.save()
    End Sub


    Sub updatTemp2Live(ByVal databaseManager As DatabaseManager, ByRef stepinfo As StepInfo, ByVal sAPPROVED As String)
        Dim dt As DataTable = loadSYTempInfo()
        Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(databaseManager)
        Dim syRoleSuitSysList As New AUTH_OP.SY_ROLESUITSYSList(databaseManager)
        Dim syRole As New AUTH_OP.SY_ROLE(databaseManager)
        Dim syRoleSuitSys As New AUTH_OP.SY_ROLESUITSYS(databaseManager)

        Dim iNewRoleId As Integer = syRole.getMaxRoleId()

        If sAPPROVED = "Y" Then

            For Each row As DataRow In dt.Rows
                Dim iRoleId As Integer = CInt(row("roleid"))
                Dim iParentId As Integer = CInt(row("parent"))

                If iRoleId > 0 Then Continue For

                If iRoleId < 0 Then '取號
                    iNewRoleId += 1
                End If

                For Each childRow As DataRow In dt.Select("[parent]='" & iRoleId & "'")
                    childRow.BeginEdit()
                    childRow("parent") = iNewRoleId
                    childRow.EndEdit()
                    dt.AcceptChanges()
                Next

                row.BeginEdit()
                row("roleid") = iNewRoleId
                row.EndEdit()
                dt.AcceptChanges()
            Next
        End If

        If IsNothing(dt) = False Then

            For Each row As DataRow In dt.Rows
                Dim iRoleid As Integer = CInt(row("ROLEID"))
                '取得原本SY_RoleSuitSys的資料
                '整理沒有在Temp裡的角色
                Dim dtOldData As DataTable = loadSYRoleSuitSys(iRoleid)
                Dim sRoleOperation As String = String.Empty
                Dim sSysOperation As String = String.Empty
                'If Not dtOldData Is Nothing Then

                '    For Each olrRow As DataRow In dtOldData.Select("roleid=" & iRoleid)

                '        If Not (CInt(row("ROLEID")) = CInt(olrRow("ROLEID")) _
                '             AndAlso row("SUBSYSID").ToString = olrRow("SUBSYSID").ToString _
                '             AndAlso row("SYSID").ToString = olrRow("SYSID").ToString) Then
                '            '1.刪除沒有在Temp裡的角色與交易的關聯(SY_REL_ROLE_FUNCTION)-->f前面已經有做檢核，假使角色已經有交易不可移除權限，所以這裡已經不用delete SY_REL_ROLE_FUNCTION
                '            '2.刪除角色與子系統、系統的關聯(SY_ROLESUITSYS)                        
                '            '3.刪除角色
                '            '4.紀錄刪除(D)的 his(SY_REL_ROLE_FUNCTION,SY_ROLESUITSYS,SY_ROLE)
                '            syRole.clear()

                '            ' TODO: 2012/06/29  暫時Remark掉 會報錯 設計到關聯檔
                '            ' syRole.delData(iRoleid, olrRow("SUBSYSID").ToString, olrRow("SYSID").ToString)
                '            syRoleSuitSys.clear()
                '            syRoleSuitSys.delRoleSys(iRoleid, olrRow("SUBSYSID").ToString, olrRow("SYSID").ToString)

                '            If stepinfo Is Nothing Then '不走流程
                '                insertSyRoleSuitSysHis(databaseManager, olrRow, "D")
                '                insertSyRoleHis(databaseManager, row, "D")
                '            Else
                '                insertSyRoleSuitSysHis(databaseManager, olrRow, "I", _
                '                                       stepinfo.currentStepInfo.caseId, _
                '                                       stepinfo.currentStepInfo.stepNo, _
                '                                       stepinfo.currentStepInfo.subflowSeq, _
                '                                       stepinfo.currentStepInfo.subflowCount)
                '                insertSyRoleHis(databaseManager, row, "I", sAPPROVED, _
                '                                       stepinfo.currentStepInfo.caseId, _
                '                                       stepinfo.currentStepInfo.stepNo, _
                '                                       stepinfo.currentStepInfo.subflowSeq, _
                '                                       stepinfo.currentStepInfo.subflowCount)
                '            End If
                '        End If
                '    Next
                'End If


                '5更新角色SY_ROLE
                syRole.clear()
                If Not syRole.loadByPK(CInt(row("ROLEID"))) Then
                    syRole.setInt("ROLEID", iRoleid)
                    sRoleOperation = "I"
                Else
                    sRoleOperation = "N" '思考一下，會不會有更新的狀況
                End If

                If sAPPROVED = "Y" Then

                    syRole.setString("ROLENAME", row("ROLENAME").ToString)
                    syRole.setString("DISABLED", row("DISABLED").ToString)
                    syRole.setString("ROLETYPE", row("ROLETYPE").ToString)

                    If String.IsNullOrEmpty(row("ROLEMGR").ToString) Then
                        syRole.setAttribute("ROLEMGR", Nothing)
                    Else
                        syRole.setInt("ROLEMGR", row("ROLEMGR"))
                    End If


                    syRole.setInt("PARENT", CInt(row("PARENT")))
                    syRole.save()
                End If

                hidParId.Value &= iRoleid.ToString & ","

                ' 若父節點有更新啟用狀態，則子節點也需要修改
                modifyLiveData(iRoleid, row("DISABLED").ToString)


                Dim sParentIdList As String = hidParId.Value
                If sParentIdList.Length > 0 Then
                    sParentIdList = sParentIdList.Substring(0, sParentIdList.Length - 1)
                End If

                ' 修改子節點的啟用狀態與父節點一樣
                If sAPPROVED = "Y" Then
                    syRole.updateDisabledByRoleIdList(sParentIdList, row("DISABLED").ToString())
                End If

                '6.紀錄新增(I)或修改(U) his(SY_ROLE)
                'If stepinfo Is Nothing Then '不走流程
                '    insertSyRoleHis(databaseManager, row, sRoleOperation)
                'Else
                insertSyRoleHis(databaseManager, row, sRoleOperation, sAPPROVED, _
                                       stepinfo.currentStepInfo.caseId, _
                                       stepinfo.currentStepInfo.stepNo, _
                                       stepinfo.currentStepInfo.subflowSeq, _
                                       stepinfo.currentStepInfo.subflowCount)
                'End If

                '7.更新角色綁子系統SY_ROLESUITSYS
                Dim dtTempData As New DataTable
                Dim sTempData As String = String.Empty '目前DB中的集合
                Dim sTempList As String = String.Empty '記錄本次編輯的節點

                If syRoleSuitSysList.loadDataByRoleId(row("ROLEID")) Then
                    dtTempData = syRoleSuitSysList.getCurrentDataSet.Tables(0)
                    If Not dtTempData Is Nothing And dtTempData.Rows.Count > 0 Then
                        For Each dr As DataRow In dtTempData.Rows
                            sTempData &= dr("SUBSYSID") & "*" & dr("SYSID") & ","
                        Next
                    End If
                End If

                ' 此部份為新增與未修改的部份 
                For i As Integer = 0 To row("SUBSYSID").ToString.Split(",").Length - 1
                    Dim sSubSysId As String = row("SUBSYSID").ToString.Split(",")(i)
                    Dim sSysId As String = row("SYSID").ToString.Split(",")(i)

                    sTempList &= sSubSysId & "*" & sSysId & ","

                    syRoleSuitSys.clear()
                    If Not syRoleSuitSys.LoadByPK(CInt(row("ROLEID")), sSubSysId, sSysId) Then
                        sSysOperation = "I"
                    Else
                        '  sSysOperation = "N" '思考一下，會不會有更新的狀況

                        If sTempData.Length > 0 Then
                            If sTempData.Contains(sSubSysId & "*" & sSysId) Then
                                sSysOperation = "N"
                            End If
                        End If
                    End If

                    If sAPPROVED = "Y" Then
                        syRoleSuitSys.setInt("ROLEID", CInt(row("ROLEID")))
                        syRoleSuitSys.setString("SUBSYSID", sSubSysId)
                        syRoleSuitSys.setString("SYSID", sSysId)
                        syRoleSuitSys.save()
                    End If

                    '8.紀錄新增(I)或修改(U) his(SY_ROLESUITSYS)
                    'If stepinfo Is Nothing Then '不走流程
                    '    insertSyRoleSuitSysHis(databaseManager, row, sSysOperation, , , , , sSubSysId, sSysId)
                    'Else
                    insertSyRoleSuitSysHis(databaseManager, row, sSysOperation, _
                                           stepinfo.currentStepInfo.caseId, _
                                           stepinfo.currentStepInfo.stepNo, _
                                           stepinfo.currentStepInfo.subflowSeq, _
                                           stepinfo.currentStepInfo.subflowCount, sSubSysId, sSysId)
                    'End If
                Next


                ' 將真實檔中的資料里排除到操作過的未修改的資料，則剩下刪除的資料
                For Each Str As String In sTempList.Split(",")
                    If Str <> "" Then
                        If sTempData.Contains(Str) Then
                            sTempData = sTempData.Replace(Str, "")
                        End If
                    End If
                Next

                ' 目前只剩下刪除的資料
                If sTempData.Length > 0 Then
                    For Each Str As String In sTempData.Split(",")
                        If Str <> "" Then

                            If sAPPROVED = "Y" Then
                                If syRoleSuitSys.LoadByPK(CInt(row("ROLEID")), Str.Split("*")(0), Str.Split("*")(1)) Then
                                    syRoleSuitSys.remove()
                                End If
                            End If

                            '8.紀錄新增(I)或修改(U) his(SY_ROLESUITSYS)
                            'If stepinfo Is Nothing Then '不走流程
                            '    insertSyRoleSuitSysHis(databaseManager, row, "D", , , , , Str.Split("*")(0), Str.Split("*")(1))
                            'Else
                            insertSyRoleSuitSysHis(databaseManager, row, "D", _
                                                   stepinfo.currentStepInfo.caseId, _
                                                   stepinfo.currentStepInfo.stepNo, _
                                                   stepinfo.currentStepInfo.subflowSeq, _
                                                   stepinfo.currentStepInfo.subflowCount, Str.Split("*")(0), Str.Split("*")(1))
                            'End If
                        End If
                    Next
                End If
            Next
        End If

        ' 同意。不同意都要刪除臨時檔
        If sAPPROVED = "Y" OrElse sAPPROVED = "N" Then
            If m_sFourStepNo = "" Or m_sFourStepNo = "0300" Then
                If syTempInfo.loadTempData(m_sWorkingUserid, m_sWorkingTopDepNo, m_sFuncCode) Then
                    syTempInfo.remove()
                End If
            Else
                If syTempInfo.loadByCaseId(m_sCaseId) Then
                    syTempInfo.remove()
                End If
            End If
        End If
    End Sub

    ''' <summary>
    ''' 若父節點的啟用狀態改變，遞歸設置其子節點的狀態，應與其父節點一致---->臨時檔資料
    ''' </summary>
    ''' <param name="dtSYTempInfo"></param>
    ''' <param name="sRoleId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function modifyData(ByVal dtSYTempInfo As DataTable, ByVal sRoleId As String) As DataTable

        ' 若父節點為停用 則子節點也為停用,若父節點為啟用，所有的都為啟用
        For Each childRow As DataRow In dtSYTempInfo.Select("[parent]='" & sRoleId & "'")
            childRow.BeginEdit()
            Dim rowsT() As DataRow = dtSYTempInfo.Select("ROLEID='" & sRoleId & "'")
            childRow("DISABLED") = rowsT(0)("DISABLED")
            childRow.EndEdit()
            dtSYTempInfo.AcceptChanges()

            If dtSYTempInfo.Select("[parent]='" & childRow("ROLEID") & "'").Count > 0 Then
                modifyData(dtSYTempInfo, childRow("ROLEID").ToString())
            End If
        Next

        Return dtSYTempInfo
    End Function

    ''' <summary>
    ''' 若父節點的啟用狀態改變，遞歸設置其子節點的狀態，應與其父節點一致---->真實檔資料
    ''' </summary>
    ''' <param name="sRoleId">角色編號</param>
    ''' <param name="sDisabled">父節點的啟用狀態</param>
    ''' <remarks></remarks>
    Sub modifyLiveData(ByRef sRoleId As String, ByVal sDisabled As String)
        Dim syRoleList As New AUTH_OP.SY_ROLEList(GetDatabaseManager)
        Dim dtData As New DataTable

        ' 如果目前只是修改的父節點，則需要根據其roleid查找是否有子節點，將其disable設置為與
        If syRoleList.loadDataByType(sRoleId) Then
            For Each item As AUTH_OP.SY_ROLE In syRoleList

                hidParId.Value &= item.getAttribute("ROLEID") & ","

                Dim syRoleTempList As New AUTH_OP.SY_ROLEList(GetDatabaseManager)

                If syRoleTempList.loadDataByType(item.getAttribute("ROLEID")) Then
                    modifyLiveData(item.getAttribute("ROLEID"), sDisabled)
                End If
            Next
        End If
    End Sub
#End Region

#Region "正常區塊Function"

    ''' <summary>
    ''' 初始加載 顯示右邊維護區資料
    ''' </summary>
    ''' <remarks></remarks>
    Sub bindBank()
        Dim apCodelist As New AP_CODEList(GetDatabaseManager())
        Dim sBankTitle As String = String.Empty

        If apCodelist.loadDataByCodeId(m_sUpCode2366) Then
            sBankTitle = apCodelist.getCurrentDataSet.Tables(0).Rows(0)("TEXT").ToString()
        End If

        ' 顯示成頁面初始加載狀態
        lblParentName.Text = "-"
        lblID.Text = "0"
        lblName.Text = sBankTitle
    End Sub
#End Region

#Region "審核區塊Function"

    ''' <summary>
    ''' 公用方法 新增歷史檔
    ''' </summary>
    ''' <param name="sApproved"></param>
    ''' <remarks></remarks>
    Sub insertDataToHis(ByVal sApproved As String)
        Try
            ' 開始事物
            GetDatabaseManager().beginTran()

            ' 新增歷史檔
            operationXml("1", sApproved)

            ' 提交事物
            GetDatabaseManager().commit()
        Catch ex As Exception

            ' 回滾事務
            GetDatabaseManager.Rollback()
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' 綁定審核模塊
    ''' </summary>
    ''' <remarks></remarks>
    Sub bindCheck()
        Dim sySysIdList As New AUTH_OP.SY_SYSIDList(GetDatabaseManager())
        Dim dtSysData As New DataTable

        ' (9) 查詢第一層的數據資料
        If sySysIdList.loadUseData() Then
            dtSysData = sySysIdList.getCurrentDataSet.Tables(0)
        End If

        ' (10)循環 新增到TreeView中
        If Not dtSysData Is Nothing Then
            For Each dr As DataRow In dtSysData.Rows

                ' 父節點
                Dim parentNode As TreeNode = New TreeNode()
                parentNode.Value = dr("SYSID").ToString()
                parentNode.Text = dr("SYSNAME").ToString

                ' 不可點擊
                parentNode.SelectAction = TreeNodeSelectAction.None

                ' (4)添加子節點
                addSysNodeCheck(parentNode, dr("SYSID").ToString())
                treeViewSysBefore.Nodes.Add(parentNode)
            Next
        End If

        ' (10)循環 新增到TreeView中
        If Not dtSysData Is Nothing Then
            For Each dr As DataRow In dtSysData.Rows

                ' 父節點
                Dim parentNode As TreeNode = New TreeNode()
                parentNode.Value = dr("SYSID").ToString()
                parentNode.Text = dr("SYSNAME").ToString

                ' 不可點擊
                parentNode.SelectAction = TreeNodeSelectAction.None

                ' (4)添加子節點
                addSysNodeCheck(parentNode, dr("SYSID").ToString())
                treeViewSysAfter.Nodes.Add(parentNode)
            Next
        End If
    End Sub

#End Region

#Region "綁定正常區塊"

    ''' <summary>
    ''' 驗證節點是否存在,如果沒有 將其新增到TreeView中
    ''' </summary>
    ''' <param name="node">節點</param>
    ''' <param name="sValue">角色編號</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function checkExistNode(ByVal node As TreeNode, ByVal subNode As TreeNode, ByVal sParent As String) As String

        If node.ChildNodes.Count > 0 Then
            For i As Integer = 0 To node.ChildNodes.Count - 1

                If node.ChildNodes(i).Value.Split(";")(0) = sParent Then
                    node.ChildNodes(i).ChildNodes.Add(subNode)
                End If

                checkExistNode(node.ChildNodes(i), subNode, sParent)
            Next
        End If
    End Function

    ''' <summary>
    '''  判斷DB中是否有此值 若有 將其設置為紅色
    ''' </summary>
    ''' <param name="node"></param>
    ''' <param name="sRoleId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function checkNode(ByVal node As TreeNode, ByVal sRoleId As String)
        For Each item As TreeNode In node.ChildNodes
            If item.Value.Split(";")(0) = sRoleId Then
                item.Text = String.Format("<font color='red'>{0}</font>", item.Text)
            End If

            checkNode(item, sRoleId)
        Next
    End Function

    Function compareTreeViewSYS() As Boolean
        For Each node As TreeNode In treeViewSys.Nodes
            If node.Text.Contains("red") Then
                Return False
            End If
            For Each childnode As TreeNode In node.ChildNodes
                If childnode.Text.Contains("red") Then
                    Return False
                End If
            Next
        Next
        Return True
    End Function

    ''' <summary>
    ''' 初始狀態
    ''' </summary>
    ''' <remarks></remarks>
    Sub display()

        ' 顯示成頁面初始加載狀態
        lblParentName.Text = "-"
        lblID.Text = "0"
        lblName.Text = "安泰銀行"

        ' Modify by 2012/06/28
        lblName.Visible = True
        txtName.Visible = False

        ' 控件可用
        txtName.Enabled = True
        treeViewSys.Enabled = True
        rdoDisable.Enabled = True
        btnSave.Enabled = True
        btnAdd.Enabled = True
        btnDelete.Enabled = True

        ' 控制按鈕
        trRoleMgr.Visible = False
        trSubSys.Visible = False
        trDisable.Visible = False
        btnSave.Visible = False
        btnDelete.Visible = False
        btnCancel.Visible = False
        btnAdd.Visible = True
    End Sub

#End Region

#Region "綁定審核區塊"

    ''' <summary>
    ''' 審核資料加載
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function initCheckData()
        Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager())

        Dim apCodelist As New AP_CODEList(GetDatabaseManager())
        Dim xmlData As String = String.Empty

        Dim dtRootData As New DataTable

        Dim dtData As New DataTable
        Dim iCount As Integer = 0
        Dim syRole As New AUTH_OP.SY_ROLE(GetDatabaseManager())
        Dim syRoleList As New AUTH_OP.SY_ROLEList(GetDatabaseManager)

        ' (1)綁定root節點
        If apCodelist.loadDataByCodeId(m_sUpCode2366) Then
            dtRootData = apCodelist.getCurrentDataSet.Tables(0)
        End If

        If Not dtRootData Is Nothing Then

            For Each dr As DataRow In dtRootData.Rows

                ' root節點
                Dim rootNode As TreeNode = New TreeNode()
                rootNode.Value = dr("VALUE").ToString()
                rootNode.Text = dr("TEXT").ToString
                treeViewCheck.Nodes.Add(rootNode)
            Next
        End If

        '  查詢XML的資料 
        If m_sFourStepNo = "" Then
            If syTempInfo.loadByPK(m_sWorkingUserid, m_sWorkingTopDepNo, m_sFuncCode) Then
                xmlData = syTempInfo.getAttribute("TEMPDATA")
            End If
        Else
            If syTempInfo.loadByCaseId(m_sCaseId) Then
                xmlData = syTempInfo.getAttribute("TEMPDATA")
            End If
        End If

        If xmlData <> Nothing Then
            dtData = com.Azion.EloanUtility.XmlUtility.convertXML2DataTable(xmlData)
            If Not dtData Is Nothing Then
                Dim dv As System.Data.DataView = dtData.DefaultView
                dv.Sort = "roleid desc"
                dtData = dv.ToTable
            End If
        End If

        Dim sParentL As String = String.Empty
        Dim sParentE As String = String.Empty

        For Each dr As DataRow In dtData.Rows
            If dr("PARENT") > 0 Then
                sParentL &= dr("PARENT").ToString() & ","
            ElseIf dr("PARENT") = 0 Then
                sParentE = "0"
            End If
        Next

        ' 如果其父節點大於0 說明是修改或者是在其下新增的子節點
        If sParentL.Length > 0 Then

            ' 1 追溯到其PARENT=0的部份 累加其上層樹
            loadNodeByParent(sParentL)

            ' 記錄本節點之上的所有的RoleID 可以根據角色集合累加其上層樹結構
            If hidNodeParent.Value <> "" Then
                Dim dtDataLive As New DataTable

                hidNodeParent.Value = hidNodeParent.Value & sParentL
                hidNodeParent.Value = hidNodeParent.Value.Substring(0, hidNodeParent.Value.Length - 1)

                If syRoleList.loadDataByRoleId(hidNodeParent.Value) Then
                    dtDataLive = syRoleList.getCurrentDataSet.Tables(0)
                End If

                addChildNode(treeViewCheck.Nodes(0), dtDataLive, dtData)
            End If
        End If

        If sParentE.Length > 0 Then ' 若父節點=0 說明是在“安泰銀行”跟節點下新增的資料

            Dim rows() As DataRow = dtData.Select("[parent]='" & sParentE & "'")

            For Each dr As DataRow In rows
                Dim childNode As TreeNode = New TreeNode()
                childNode.Value = dr("ROLEID").ToString
                childNode.Text = String.Format("<font color='red'>{0}</font>", dr("ROLENAME").ToString & "  (" & dr("ROLEID").ToString & ")")

                If dtData.Select("[parent]='" & childNode.Value & "'").Count > 0 Then
                    addChildNode(childNode, dtData, Nothing)
                Else
                    m_sTreeNode = childNode
                End If

                treeViewCheck.Nodes(0).ChildNodes.Add(childNode)
            Next
        End If

        If Not m_sTreeNode Is Nothing Then
            loadCheck(m_sTreeNode)
        End If
    End Function

    ''' <summary>
    ''' 添加子項
    ''' </summary>
    ''' <param name="node"></param>
    ''' <param name="dtData"></param>
    ''' <remarks></remarks>
    Sub addChildNode(ByVal node As TreeNode, ByVal dtData As DataTable, ByVal dtTempData As DataTable)

        Dim rows() As DataRow = dtData.Select("[parent]='" & node.Value & "'")
        Dim sNodeList As String = String.Empty

        If rows.Count > 0 Then
            For Each drRow As DataRow In rows

                Dim childNode As TreeNode = New TreeNode()
                childNode.Value = drRow("ROLEID").ToString

                If CInt(drRow("ROLEID")) < 0 Then
                    childNode.Text = String.Format("<font color='red'>{0}</font>", drRow("ROLENAME").ToString & "  (" & drRow("ROLEID").ToString & ")")
                Else
                    ' 判斷是否存在于臨時檔，若存在，則需要變紅，若不存在，則不需要變紅
                    If dtTempData.Select("roleid='" & childNode.Value & "'").Count > 0 Then
                        childNode.Text = String.Format("<font color='red'>{0}</font>", drRow("ROLENAME").ToString & "  (" & drRow("ROLEID").ToString & ")")
                    Else
                        childNode.Text = drRow("ROLENAME").ToString & "  (" & drRow("ROLEID").ToString & ")"
                        childNode.SelectAction = TreeNodeSelectAction.None
                    End If
                End If

                If m_sTreeNode Is Nothing Then
                    If childNode.Text.Contains("red") Then
                        m_sTreeNode = childNode
                    End If
                End If

                If dtData.Select("[parent]='" & childNode.Value & "'").Count > 0 Then
                    addChildNode(childNode, dtData, dtTempData)
                End If

                If Not dtTempData Is Nothing Then
                    Dim rowsTemp() As DataRow = dtTempData.Select("[parent]='" & childNode.Value & "'")
                    Dim sNodeIdList As String = String.Empty

                    If rowsTemp.Count > 0 Then

                        For Each drRowTemp As DataRow In dtTempData.Select("[parent]='" & childNode.Value & "'")

                            Dim childNodeTemp As TreeNode = New TreeNode()
                            childNodeTemp.Value = drRowTemp("ROLEID").ToString
                            childNodeTemp.Text = String.Format("<font color='red'>{0}</font>", drRowTemp("ROLENAME").ToString & "  (" & drRowTemp("ROLEID").ToString & ")")

                            If m_sTreeNode Is Nothing Then
                                If childNodeTemp.Text.Contains("red") Then
                                    m_sTreeNode = childNodeTemp
                                End If
                            End If

                            '
                            If dtTempData.Select("[parent]='" & childNodeTemp.Value & "'").Count > 0 Then
                                addChildNode(childNodeTemp, dtTempData, dtTempData)
                            End If

                            If childNode.ChildNodes.Count > 0 Then
                                For Each item As TreeNode In childNode.ChildNodes
                                    sNodeIdList &= item.Value
                                Next
                            End If

                            If Not sNodeIdList.Contains(childNodeTemp.Value) Then
                                childNode.ChildNodes.Add(childNodeTemp)
                            End If
                        Next
                    End If
                End If

                node.ChildNodes.Add(childNode)
            Next
        End If
    End Sub

    ''' <summary>
    ''' 新增TreeView節點
    ''' </summary>
    ''' <param name="node">節點</param>
    ''' <param name="item">parent</param>
    ''' <remarks></remarks>
    Private Sub addSysNodeCheck(ByVal node As TreeNode, ByVal item As String)
        Try
            Dim sySubSysIdList As New AUTH_OP.SY_SUBSYSIDList(GetDatabaseManager())
            Dim dtData As New DataTable

            If sySubSysIdList.loadSubSys(item) Then
                dtData = sySubSysIdList.getCurrentDataSet.Tables(0)
            End If

            If Not dtData Is Nothing Then
                If dtData.Rows.Count > 0 Then
                    For Each dr As DataRow In dtData.Rows

                        Dim subNode As TreeNode = New TreeNode
                        subNode.Text = dr("SHORTSSNAME").ToString()
                        subNode.Value = dr("SUBSYSID").ToString

                        ' 不可點擊
                        subNode.SelectAction = TreeNodeSelectAction.None
                        addSysNodeCheck(subNode, dr("SUBSYSID").ToString())
                        node.ChildNodes.Add(subNode)
                    Next
                End If
            End If
        Catch ex As Exception
            showErrMsg(ex)
        End Try
    End Sub

    ''' <summary>
    ''' 審核區塊更新資料加載
    ''' </summary>
    ''' <remarks></remarks>
    Sub loadCheck(ByVal trNode As TreeNode)
        Dim nodeValue As Integer = trNode.Value
        Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager())
        Dim syRoleSuitSysList As New AUTH_OP.SY_ROLESUITSYSList(GetDatabaseManager())
        Dim syRole As New AUTH_OP.SY_ROLE(GetDatabaseManager())
        Dim dtData As New DataTable
        Dim xmlDocument As New XmlDocument
        Dim xmlData As String = String.Empty
        Dim syRoleHisList As New SY_ROLE_HISList(GetDatabaseManager())


        ' 第一次新增資料
        If nodeValue < 0 Then

            '將查詢出資料顯示到修改後區塊中()
            ' 查詢臨時檔資料
            If m_sFourStepNo = "" Then
                If syTempInfo.loadByPK(m_sWorkingUserid, m_sWorkingTopDepNo, m_sFuncCode) Then
                    xmlData = syTempInfo.getAttribute("TEMPDATA")
                End If
            Else
                If syTempInfo.loadByCaseId(m_sCaseId) Then
                    xmlData = syTempInfo.getAttribute("TEMPDATA")
                End If
            End If

            If Not xmlData Is Nothing Then

                ' document對象載入XML文件
                xmlDocument.LoadXml(xmlData)
            End If

            For i As Integer = 0 To xmlDocument.GetElementsByTagName("SY_ROLE").Count - 1

                Dim syRoleT As New AUTH_OP.SY_ROLE(GetDatabaseManager())

                Dim sRoleId As String = DirectCast(xmlDocument.GetElementsByTagName("SY_ROLE")(i)("ROLEID"), System.Xml.XmlElement).InnerText
                Dim sParent As String = DirectCast(xmlDocument.GetElementsByTagName("SY_ROLE")(i)("PARENT"), System.Xml.XmlElement).InnerText
                Dim sRoleName As String = DirectCast(xmlDocument.GetElementsByTagName("SY_ROLE")(i)("ROLENAME"), System.Xml.XmlElement).InnerText
                Dim sDisabled As String = DirectCast(xmlDocument.GetElementsByTagName("SY_ROLE")(i)("DISABLED"), System.Xml.XmlElement).InnerText
                Dim sRoleType As String = DirectCast(xmlDocument.GetElementsByTagName("SY_ROLE")(i)("ROLETYPE"), System.Xml.XmlElement).InnerText
                Dim sSubSysId As String = DirectCast(xmlDocument.GetElementsByTagName("SY_ROLE")(i)("SUBSYSID"), System.Xml.XmlElement).InnerText
                Dim sSysId As String = DirectCast(xmlDocument.GetElementsByTagName("SY_ROLE")(i)("SYSID"), System.Xml.XmlElement).InnerText

                If sRoleId = nodeValue.ToString Then
                    If Convert.ToInt32(sParent) > 0 Then
                        If syRoleT.loadByPK(sParent) Then
                            lblParentNameAfter.Text = syRoleT.getAttribute("ROLENAME")
                        End If
                    Else
                        If Not trNode.Parent Is Nothing Then
                            lblParentNameAfter.Text = trNode.Parent.Text.Replace("<font color='red'>", "").Replace("</font>", "")
                        End If
                    End If

                    lblIDAfter.Text = nodeValue
                    lblNameAfter.Text = sRoleName
                    rdoDisableAfter.SelectedValue = sDisabled

                    ' 清空之前選中的資料
                    For Each nodes As TreeNode In treeViewSysAfter.Nodes
                        nodes.Checked = False
                        If nodes.ChildNodes.Count > 0 Then
                            For j As Integer = 0 To nodes.ChildNodes.Count - 1
                                nodes.ChildNodes(j).Checked = False
                            Next
                        End If
                    Next

                    ' 欄位賦值部份
                    If sSubSysId.Length > 0 Then
                        Dim sSubsys() As String = sSubSysId.Split(",")
                        Dim sSysIdNew() As String = sSysId.Split(",")
                        Dim iCount As Integer = 0
                        '   For j As Integer = 0 To sSubsys.Count - 1

                        For m As Integer = 0 To sSysIdNew.Count - 1

                            For Each node As TreeNode In treeViewSysAfter.Nodes
                                If node.ChildNodes.Count > 0 Then

                                    For t As Integer = 0 To node.ChildNodes.Count - 1

                                        ' 如果子節點中的value等於子系統 則選中
                                        If node.ChildNodes(t).Value = sSubsys(m).ToString() Then
                                            If node.ChildNodes(t).Parent.Value = sSysIdNew(m).ToString() Then
                                                node.ChildNodes(t).Checked = True
                                                node.ChildNodes(t).Parent.Checked = True
                                            End If
                                        End If
                                    Next
                                End If
                            Next
                        Next
                        ' Next
                    End If
                End If

                ' 清空修改前的資料，因為之前live檔無資料
                lblParentNameBefore.Text = ""
                lblIDBefore.Text = ""
                lblNameBefore.Text = ""
                rdoDisableBefore.SelectedIndex = -1

                ' 清空之前選中的資料
                For Each nodes As TreeNode In treeViewSysBefore.Nodes
                    nodes.Checked = False
                    If nodes.ChildNodes.Count > 0 Then
                        For j As Integer = 0 To nodes.ChildNodes.Count - 1
                            nodes.ChildNodes(j).Checked = False
                        Next
                    End If
                Next
            Next
        Else
            ' 將查詢出資料顯示到修改前區塊中
            If syRole.loadByPK(nodeValue.ToString) Then
                lblParentNameBefore.Text = trNode.Parent.Text.Replace("<font color='red'>", "").Replace("</font>", "")
                lblIDBefore.Text = nodeValue
                lblNameBefore.Text = syRole.getAttribute("ROLENAME")
                rdoDisableBefore.SelectedValue = syRole.getAttribute("DISABLED")

                ' 查詢維護檔是否有資料
                If syRoleSuitSysList.loadDataByRoleId(lblIDBefore.Text) > 0 Then
                    dtData = syRoleSuitSysList.getCurrentDataSet.Tables(0)
                End If

                ' 清空之前選中的資料
                For Each nodes As TreeNode In treeViewSysBefore.Nodes
                    nodes.Checked = False
                    If nodes.ChildNodes.Count > 0 Then
                        For j As Integer = 0 To nodes.ChildNodes.Count - 1
                            nodes.ChildNodes(j).Checked = False
                        Next
                    End If
                Next

                ' 若有資料
                If Not dtData Is Nothing Then
                    If dtData.Rows.Count > 0 Then
                        Dim iCount As Integer = 0

                        ' 循環 給“所屬子系統”的TreeView賦值
                        For Each dr As DataRow In dtData.Rows
                            For Each node As TreeNode In treeViewSysBefore.Nodes

                                If node.Value = dr("SYSID").ToString() Then

                                    ' 若有子節點 則子節點也選中
                                    If node.ChildNodes.Count > 0 Then

                                        For m As Integer = 0 To node.ChildNodes.Count - 1

                                            ' 如果子節點中的value等於子系統 則選中
                                            If node.ChildNodes(m).Value = dr("SUBSYSID").ToString() Then
                                                node.ChildNodes(m).Checked = True
                                                node.Checked = True
                                            End If
                                        Next
                                    End If
                                End If
                            Next
                        Next
                    End If
                End If
            End If

            ' 查詢臨時檔資料
            If m_sFourStepNo = "" Then
                If syTempInfo.loadByPK(m_sWorkingUserid, m_sWorkingTopDepNo, m_sFuncCode) Then
                    xmlData = syTempInfo.getAttribute("TEMPDATA")
                End If
            Else
                If syTempInfo.loadByCaseId(m_sCaseId) Then
                    xmlData = syTempInfo.getAttribute("TEMPDATA")
                End If
            End If

            If Not xmlData Is Nothing Then

                ' document對象載入XML文件
                xmlDocument.LoadXml(xmlData)
            End If

            For i As Integer = 0 To xmlDocument.GetElementsByTagName("SY_ROLE").Count - 1

                Dim sRoleId As String = DirectCast(xmlDocument.GetElementsByTagName("SY_ROLE")(i)("ROLEID"), System.Xml.XmlElement).InnerText
                Dim sParent As String = DirectCast(xmlDocument.GetElementsByTagName("SY_ROLE")(i)("PARENT"), System.Xml.XmlElement).InnerText
                Dim sRoleName As String = DirectCast(xmlDocument.GetElementsByTagName("SY_ROLE")(i)("ROLENAME"), System.Xml.XmlElement).InnerText
                Dim sDisabled As String = DirectCast(xmlDocument.GetElementsByTagName("SY_ROLE")(i)("DISABLED"), System.Xml.XmlElement).InnerText
                Dim sRoleType As String = DirectCast(xmlDocument.GetElementsByTagName("SY_ROLE")(i)("ROLETYPE"), System.Xml.XmlElement).InnerText
                Dim sSubSysId As String = DirectCast(xmlDocument.GetElementsByTagName("SY_ROLE")(i)("SUBSYSID"), System.Xml.XmlElement).InnerText
                Dim sSysId As String = DirectCast(xmlDocument.GetElementsByTagName("SY_ROLE")(i)("SYSID"), System.Xml.XmlElement).InnerText

                If sRoleId = nodeValue.ToString Then
                    lblParentNameAfter.Text = trNode.Parent.Text
                    lblIDAfter.Text = nodeValue
                    lblNameAfter.Text = sRoleName
                    rdoDisableAfter.SelectedValue = sDisabled

                    ' 清空之前選中的資料
                    For Each nodes As TreeNode In treeViewSysAfter.Nodes
                        nodes.Checked = False
                        If nodes.ChildNodes.Count > 0 Then
                            For j As Integer = 0 To nodes.ChildNodes.Count - 1
                                nodes.ChildNodes(j).Checked = False
                            Next
                        End If
                    Next

                    ' 欄位賦值部份
                    If sSubSysId.Length > 0 Then
                        Dim sSubsys() As String = sSubSysId.Split(",")
                        Dim sSysIdNew() As String = sSysId.Split(",")
                        Dim iCount As Integer = 0

                        For m As Integer = 0 To sSysIdNew.Count - 1

                            '  For j As Integer = 0 To sSubsys.Count - 1
                            For Each node As TreeNode In treeViewSysAfter.Nodes

                                If node.Value = sSysIdNew(m) Then
                                    node.Checked = True
                                End If

                                If node.ChildNodes.Count > 0 Then

                                    For t As Integer = 0 To node.ChildNodes.Count - 1

                                        ' 如果子節點中的value等於子系統 則選中
                                        If node.ChildNodes(t).Value = sSubsys(m).ToString() Then
                                            If node.ChildNodes(t).Parent.Value = sSysIdNew(m).ToString() Then
                                                node.ChildNodes(t).Checked = True
                                            End If
                                        End If
                                    Next
                                End If
                                'Next
                            Next
                        Next
                    End If
                End If
            Next
        End If
    End Sub

    ''' <summary>
    ''' 根據父節點往上推到其父節點
    ''' </summary>
    ''' <param name="sParent"></param>
    ''' <remarks></remarks>
    Sub loadNodeByParent(ByVal sParent As String)
        Dim syRole As New AUTH_OP.SY_ROLE(GetDatabaseManager())

        For Each Str As String In sParent.Split(",")
            If syRole.loadByPK(Str) Then
                If syRole.getAttribute("PARENT") = "0" Then
                    hidNodeParent.Value &= syRole.getAttribute("ROLEID").ToString() & ","
                Else
                    loadNodeByParent(syRole.getAttribute("PARENT"))
                    hidNodeParent.Value &= syRole.getAttribute("PARENT").ToString & ","
                End If
            End If
        Next

    End Sub
#End Region

#Region "正常區塊Event"

    ''' <summary>
    ''' treeView SelectedNodeChanged
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub treeView_SelectedNodeChanged(ByVal sender As Object, ByVal e As EventArgs) Handles treeView.SelectedNodeChanged
        Dim flag As Boolean = False
        Dim iRoleId As Integer = CInt(treeView.SelectedNode.Value.Split(";")(0))
        Dim nRoleMgr As Integer = 0

        Try
            ' -1；0；2；0 分別對應ROLEID;PARENTID;ROLETYPE;DISABLED
            ' 控件賦值
            hidName.Value = ""
            hidId.Value = ""
            txtName.Text = treeView.SelectedNode.Text.Replace("<font color='red'>", "").Replace("</font>", "").Replace("  (", "").Replace(treeView.SelectedNode.Value.Split(";")(0) & ")", "")
            lblName.Text = treeView.SelectedNode.Text.Replace("<font color='red'>", "").Replace("</font>", "")
            lblID.Text = iRoleId.ToString

            ' 初始為黑色
            txtName.ForeColor = Drawing.Color.Black
            For Each item As ListItem In rdoDisable.Items
                item.Text = item.Text.Replace("<font color='red'>", "").Replace("</font>", "")
            Next

            If Not treeView.SelectedNode.Parent Is Nothing Then
                lblParentName.Text = treeView.SelectedNode.Parent.Text.Replace("<font color='red'>", "").Replace("</font>", "")
            End If

            resetTreeViewSys()

            If Not treeView.SelectedNode.Parent Is Nothing Then
                'sRoleId & ";" & sParent & ";" & sROLETYPE & ";" & sDISABLED
                initSyRoleSuitSys(treeView.SelectedNode.Value.Split(";")(1))
            Else
                initSyRoleSuitSys(0)
                trRoleMgr.Visible = False
                trSubSys.Visible = False
                trDisable.Visible = False

                btnSave.Visible = False
                btnDelete.Visible = False
                btnCancel.Visible = False
            End If

            nRoleMgr = GetRoleMgrId(iRoleId)
            'UpdateRoleMgrName()

            btnAdd.Visible = True

            ' Add by 2012/06/28
            If iRoleId < 0 Then
                btnDelete.Visible = True
                hidId.Value = iRoleId
            End If

            ' 查詢臨時檔中的資料
            Dim dt As DataTable = loadSYTempInfo()

            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                For Each dr As DataRow In dt.Rows
                    If dr("ROLEID").ToString() = lblID.Text Then
                        flag = True
                        Exit For
                    End If
                Next
            End If

            '更新ROLEMGR
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                For Each dr As DataRow In dt.Rows
                    Dim sRoleId As String
                    sRoleId = dr("ROLEID").ToString().Trim()

                    If sRoleId = lblID.Text Then
                        If String.IsNullOrEmpty(sRoleId) Then
                            nRoleMgr = 0
                        Else
                            nRoleMgr = CInt(sRoleId)
                        End If
                        Exit For
                    End If
                Next
            End If

            If nRoleMgr = 0 Then
                tbRoleMgr.Text = ""
            Else
                tbRoleMgr.Text = nRoleMgr.ToString
            End If
            UpdateRoleMgr()

            ' 根據主鍵查詢角色檔 查看父節點是否有更新過

            Dim sDisabled As String = String.Empty

            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                For Each dr As DataRow In dt.Rows
                    If dr("ROLEID").ToString() = treeView.SelectedNode.Parent.Value.Split(";")(0) Then
                        sDisabled = dr("DISABLED").ToString
                    End If
                Next
            End If

            If Not String.IsNullOrEmpty(sDisabled) Then
                rdoDisable.SelectedValue = sDisabled
            Else
                rdoDisable.SelectedValue = treeView.SelectedNode.Value.Split(";")(3)
            End If

            If flag = True Then
                txtName.ForeColor = Drawing.Color.Red
                rdoDisable.SelectedItem.Text = String.Format("<font color='red'>{0}</font>", rdoDisable.SelectedItem.Text)
            Else
                txtName.ForeColor = Drawing.Color.Black
                rdoDisable.SelectedItem.Text = rdoDisable.SelectedItem.Text.Replace("<font color='red'>", "").Replace("</font>", "")
            End If

            Dim bHasTempData As Boolean = bindTreeViewSYSForSYTempInfo(iRoleId)

            bindTreeViewSYSForSYRoleSuitSys(iRoleId, bHasTempData)

            If Not compareTreeViewSYS() Then
                treeView.SelectedNode.Text = String.Format("<font color='red'>{0}</font>", treeView.SelectedNode.Text)
            Else
                treeView.SelectedNode.Text = treeView.SelectedNode.Text.Replace("<font color='red'>", "").Replace("</font>", "")
            End If

            If treeView.SelectedNode.Value = "0;0;0;0" Then
                display()
                hidName.Value = ""
                hidId.Value = ""
            End If

            '將select node設為disable
            treeView.SelectedNode.SelectAction = TreeNodeSelectAction.None

            If Not ViewState("LastNodePath") Is Nothing Then
                '將select node設為enble
                Dim lastNode As TreeNode = treeView.FindNode(ViewState("LastNodePath").ToString())
                lastNode.SelectAction = TreeNodeSelectAction.Select
                If Not ViewState("hasSave") Is Nothing Then
                    If Not CBool(ViewState("hasSave")) Then
                        lastNode.Text = ViewState("LastNodeText").ToString
                    End If
                    ViewState.Remove("hasSave")
                End If
            End If

            ViewState("LastNodePath") = treeView.SelectedNode.ValuePath

            hidFlag.Value = "Modify"
            hidOldData.Value = treeView.SelectedNode.Value
        Catch ex As Exception
            showErrMsg(ex)
        End Try
    End Sub

    ''' <summary>
    ''' 所屬子系統
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub treeViewSys_TreeNodeCheckChanged(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.TreeNodeEventArgs) Handles treeViewSys.TreeNodeCheckChanged
        Dim func As New SY_REL_ROLE_FUNCTION(GetDatabaseManager())
        Dim syRoleSuitSys As New AUTH_OP.SY_ROLESUITSYS(GetDatabaseManager())
        Dim syRoleSuitSysList As New AUTH_OP.SY_ROLESUITSYSList(GetDatabaseManager())
        Dim syRelRoleUserList As New AUTH_OP.SY_REL_ROLE_USERList(GetDatabaseManager())
        Dim iRoleId As Integer = lblID.Text

        btnAdd.Visible = False '新增
        btnSave.Visible = True '儲存
        btnDelete.Visible = False '刪除
        btnCancel.Visible = True '取消

        '父節點選中，判斷是否有子項，子項也要選中
        For Each item As TreeNode In e.Node.ChildNodes
            If item.ShowCheckBox Then '有開checkbox才可勾選
                item.Checked = e.Node.Checked
            End If
        Next
        e.Node.Text = e.Node.Text.Replace("<font color='red'>", "").Replace("</font>", "")
        If Not e.Node.Parent Is Nothing Then e.Node.Parent.Text = e.Node.Parent.Text.Replace("<font color='red'>", "").Replace("</font>", "")

        If e.Node.Checked Then
            e.Node.Text = String.Format("<font color='red'>{0}</font>", e.Node.Text)
            If Not e.Node.Parent Is Nothing Then e.Node.Parent.Text = String.Format("<font color='red'>{0}</font>", e.Node.Parent.Text)
        Else
            ' 判斷roleid是否在SY_REL_ROLE_USER與SY_REL_ROLE_FUNCTION中，若存在，則提示信息，不可進行修改
            If syRelRoleUserList.loadDataByRoleId(treeView.SelectedValue.Split(";")(0)) > 0 Then

                If Not e.Node.Parent Is Nothing Then
                    If syRoleSuitSys.LoadByPK(treeView.SelectedValue.Split(";")(0), e.Node.Value, e.Node.Parent.Value) Then
                        com.Azion.EloanUtility.UIUtility.alert("此角色已經設定人員，不可取消勾選！")
                        e.Node.Checked = True

                        If e.Node.ChildNodes.Count > 0 Then
                            For Each item As TreeNode In e.Node.ChildNodes
                                item.Checked = True
                            Next
                        End If
                        Return
                    End If
                Else
                    ' 說明是勾選的父節點進行的取消
                    If syRoleSuitSysList.loadDataByCon(treeView.SelectedValue.Split(";")(0), e.Node.Value) Then
                        Dim dtTemp As DataTable = syRoleSuitSysList.getCurrentDataSet.Tables(0)
                        Dim sSubSysIdList As String = String.Empty

                        ' 查詢其裏面的子節點
                        If Not dtTemp Is Nothing Then
                            For Each dr As DataRow In dtTemp.Rows
                                sSubSysIdList &= dr("SUBSYSID") & ","
                            Next
                        End If

                        If e.Node.ChildNodes.Count > 0 Then
                            For Each childNode As TreeNode In e.Node.ChildNodes
                                If sSubSysIdList.Contains(childNode.Value) Then
                                    com.Azion.EloanUtility.UIUtility.alert("此角色已經設定人員，不可取消勾選！")
                                    e.Node.Checked = True

                                    If e.Node.ChildNodes.Count > 0 Then
                                        For Each item As TreeNode In e.Node.ChildNodes
                                            If sSubSysIdList.Contains(item.Value) Then
                                                item.Checked = True
                                            End If
                                        Next
                                    End If
                                    Return
                                End If
                            Next
                        End If
                    End If
                End If
            End If

            ' 記錄刪除的資料
            hidDelData.Value &= e.Node.Value & ";"
        End If


        '若同一级都没有选择，同时取消父节点的选择。
        If Not e.Node.Parent Is Nothing Then
            For Each node As TreeNode In e.Node.Parent.ChildNodes
                If e.Node.Parent.ShowCheckBox Then
                    e.Node.Parent.Checked = False
                    If node.ShowCheckBox AndAlso node.Checked Then
                        e.Node.Parent.Checked = True
                        Exit For
                    End If
                End If
            Next
        End If

        ' 根據主鍵查詢角色檔
        'rdoDisable.SelectedValue = treeView.SelectedNode.Value.Split(";")(3)

        For Each nodes As TreeNode In treeViewSys.Nodes
            nodes.Text = nodes.Text.Replace("<font color='red'>", "").Replace("</font>", "")

            If nodes.ChildNodes.Count > 0 Then
                For Each childNode As TreeNode In nodes.ChildNodes
                    childNode.Text = childNode.Text.Replace("<font color='red'>", "").Replace("</font>", "")
                Next
            End If
        Next

        For Each node As TreeNode In treeViewSys.CheckedNodes
            node.Text = node.Text.Replace("<font color='red'>", "").Replace("</font>", "")
            node.Text = String.Format("<font color='red'>{0}</font>", node.Text)

            If Not node.Parent Is Nothing Then
                node.Parent.Text = node.Parent.Text.Replace("<font color='red'>", "").Replace("</font>", "")
                node.Parent.Text = String.Format("<font color='red'>{0}</font>", node.Parent.Text)
            End If
        Next


        If iRoleId > 0 Then '存在live
            bindTreeViewSYSForSYRoleSuitSys(iRoleId, True)
        End If

        ViewState("LastNodeText") = treeView.SelectedNode.Text
        If Not compareTreeViewSYS() Then
            treeView.SelectedNode.Text = String.Format("<font color='red'>{0}</font>", treeView.SelectedNode.Text)
            ViewState("hasSave") = False '此節點是否已經儲存，如果未存檔，又變紅色，當在選取其它節點，需變回原色
        Else
            treeView.SelectedNode.Text = treeView.SelectedNode.Text.Replace("<font color='red'>", "").Replace("</font>", "")
        End If

        If iRoleId < 0 Then
            txtName.Visible = True
            lblName.Visible = False
        Else
            txtName.Visible = False
            lblName.Visible = True
        End If

        For Each nodes As TreeNode In treeViewSys.Nodes
            For Each childNode As TreeNode In nodes.ChildNodes
                If childNode.Text.Contains("red") Then

                    Dim sSysId As String = childNode.Parent.Value
                    Dim sSubSysId As String = childNode.ValuePath.Split("/")(1)

                    func.clear()
                    If func.loadFunction(iRoleId, sSubSysId, sSysId) Then

                        childNode.Checked = True
                        childNode.Parent.Checked = True

                        childNode.Text = childNode.Text.Replace("<font color='red'>", "").Replace("</font>", "")
                        childNode.Parent.Text = childNode.Parent.Text.Replace("<font color='red'>", "").Replace("</font>", "")

                        com.Azion.EloanUtility.UIUtility.alert("角色【" & childNode.Parent.Text & " ─" & childNode.Text & "】已經分派交易，新先至【交易分派】將該角色所屬之交易移除！！")
                        Exit For
                    End If
                End If
            Next
        Next

    End Sub

    ''' <summary>
    ''' 新增
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnAdd_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnAdd.Click


        If IsNothing(ViewState("LastNodePath")) OrElse String.IsNullOrEmpty(ViewState("LastNodePath")) Then
            com.Azion.EloanUtility.UIUtility.alert("請先從左側角色列表內選擇父節點！")
            Return
        End If


        '名稱
        Dim iRoleId As Integer = CInt(treeView.SelectedNode.Value.Split(";")(0))

        ' 如果此節點的啟用狀態為停用 則不可新增子節點
        If treeView.SelectedNode.Value.Split(";")(3) = "1" Then
            com.Azion.EloanUtility.UIUtility.alert("此節點的啟用狀態為停用，不允許新增子節點！")
            Return
        End If

        '名稱
        txtName.Text = ""
        lblID.Text = ""

        txtName.Visible = True
        lblName.Visible = False

        btnAdd.Visible = False '新增
        btnSave.Visible = True '儲存
        btnDelete.Visible = False '刪除
        btnCancel.Visible = True '取消

        '如果只有安泰銀行 還沒有其他Child
        If treeView.SelectedNode.Parent Is Nothing Then
            'lblParentName.Text = treeView.SelectedNode.Text

            If treeView.SelectedNode.Value.Split(";")(0) = 0 Then
                If hidName.Value <> "" Then
                    lblParentName.Text = hidName.Value
                Else
                    lblParentName.Text = treeView.SelectedNode.Text
                End If
            Else
                If hidName.Value <> "" Then
                    ' Modify by 2012/06/28
                    lblParentName.Text = hidName.Value
                Else
                    lblParentName.Text = treeView.SelectedNode.Text
                End If
            End If
        Else
            'lblParentName.Text = treeView.SelectedNode.Parent.Text
            'txtName.Text = ""
            ' Modify by 2012/06/28
            If hidName.Value <> "" Then
                lblParentName.Text = hidName.Value
            Else
                lblParentName.Text = treeView.SelectedNode.Text
                txtName.Text = ""
            End If
        End If

        ' Modify by 2012/06/28
        If hidName.Value <> "" Then
            iRoleId = hidId.Value
        End If

        '清空TreeViewSYS之前選中的資料
        resetTreeViewSys()

        ' 如果父節點不是“安泰銀行” 則系統Tree所能勾選的權限 只能是父節點勾選的部份
        'sRoleId & ";" & sParent & ";" & sROLETYPE & ";" & sDISABLED
        initSyRoleSuitSys(iRoleId)

        '狀態默認為啟用
        rdoDisable.SelectedValue = "0"

        ' 初始為黑色，異動才會變色
        rdoDisable.SelectedItem.Text = rdoDisable.SelectedItem.Text.Replace("<font color='red'>", "").Replace("</font>", "")

        Dim dt As DataTable = CType(ViewState("SY_TEMPINFO"), DataTable)   'loadSYTempInfo()

        ' 取得編號
        lblID.Text = "-1"
        If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then

            '類型轉換
            If Not dt.Columns.Contains("iRoleId") Then dt.Columns.Add("iRoleId", System.Type.GetType("System.Int32"))
            dt.Columns("iRoleId").Expression = "Convert(roleid, 'System.Int32')"
            Dim iMin As Integer = dt.Compute("min(iRoleId)", "")
            If iMin < 0 Then
                lblID.Text = iMin - 1
            End If
        End If
        treeView.SelectedNode.SelectAction = TreeNodeSelectAction.Select

        hidFlag.Value = "Add"
    End Sub

    ''' <summary>
    ''' 儲存
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Sub hasRoleUser(ByVal iRoleId As Integer)

    End Sub

    ''' <summary>
    ''' 儲存
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnSave_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnSave.Click

        Dim iRoleId As Integer = lblID.Text

        Try
            If String.IsNullOrEmpty(txtName.Text) Then
                com.Azion.EloanUtility.UIUtility.alert("請選擇角色名稱！")
                txtName.Visible = True
                lblName.Visible = False
                Return
            End If

            ' Modify by 2012/06/28
            If treeViewSys.CheckedNodes.Count = 0 Then
                com.Azion.EloanUtility.UIUtility.alert("請選擇所屬子系統！")
                Return
            End If

            ' 查看其父節點的啟用狀態 如果為停用 則不可以修改子節點的狀態
            If Not treeView.SelectedNode.Parent Is Nothing Then
                If treeView.SelectedNode.Parent.Value.Split(";")(3) = "1" Then
                    com.Azion.EloanUtility.UIUtility.alert("父節點已經停用，子節點不可改為啟用！")
                    rdoDisable.SelectedValue = "1"

                    For Each item As ListItem In rdoDisable.Items
                        item.Text = item.Text.Replace("<font color='red'>", "").Replace("</font>", "")
                    Next

                    Return
                End If
            End If

            ' 組XML文件           
            Dim dtNewData As DataTable = genTempInfoTreeViewSYS2XML()
            Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager())

            If dtNewData Is Nothing Then
                If syTempInfo.loadByPK(m_sWorkingUserid, m_sWorkingTopDepNo, m_sFuncCode) Then
                    syTempInfo.remove()
                End If
                com.Azion.EloanUtility.UIUtility.alert("資料無變更，請確認是否已作修改!")
                Return
            End If

            '驗證角色名稱是否存在
            If iRoleId < 0 Then
                Dim rows() As DataRow = CType(ViewState("SY_Role"), DataTable).Select("rolename='" & txtName.Text.Trim & "'")
                If rows.Length > 0 Then
                    com.Azion.EloanUtility.UIUtility.alert("角色名稱已經存在不可重複，請重新輸入！！")

                    txtName.Visible = True
                    lblName.Visible = False

                    txtName.Focus()
                    Return
                End If
            End If

            Dim dtSYTempInfo As DataTable = CType(ViewState("SY_TEMPINFO"), DataTable)   'loadSYTempInfo() 

            ' 查找臨時檔中的資料 進行修改
            If Not dtSYTempInfo Is Nothing Then
                Dim rows() As DataRow = dtSYTempInfo.Select("ROLEID='" & lblID.Text & "'")
                ' 找到臨時檔中已經存在的資料
                ' 如果有 則將其替換掉 然後累加上新的XML
                If rows.Length > 0 Then
                    dtSYTempInfo.Rows.Remove(rows(0))
                    dtSYTempInfo.AcceptChanges()
                End If

                '進行累加
                If Not dtNewData Is Nothing Then
                    Dim rowsNewData() As DataRow = dtNewData.Select("ROLEID='" & lblID.Text & "'")
                    For Each row As DataRow In rowsNewData
                        dtSYTempInfo.ImportRow(row)
                    Next
                End If

                dtSYTempInfo = modifyData(dtSYTempInfo, lblID.Text)

                dtSYTempInfo.AcceptChanges()
            Else ' 如果找不到臨時檔，直接新增資料
                dtSYTempInfo = dtNewData
            End If

            Dim dv As System.Data.DataView = dtSYTempInfo.DefaultView
            dv.Sort = "Roleid asc"
            dtSYTempInfo = dv.ToTable

            'dtSYTempInfo一定會有資料，不做檢核
            '檢核temp有無重複ROLENAME
            Dim rowsRoleName() As DataRow = dtSYTempInfo.Select("ROLENAME='" & txtName.Text.Trim & "'")
            If rowsRoleName.Length > 1 Then
                com.Azion.EloanUtility.UIUtility.alert("角色名稱已經存在不可重複，請重新輸入！！")
                txtName.Focus()
                Return
            End If

            If dtSYTempInfo.Columns.Contains("iRoleId") Then dtSYTempInfo.Columns.Remove("iRoleId")

            syTempInfo.clear()
            If m_sCaseId = "" Then
                syTempInfo.loadByPK(m_sWorkingUserid, m_sWorkingTopDepNo, m_sFuncCode)
            Else
                syTempInfo.loadByCaseId(m_sCaseId)
            End If

            If Not syTempInfo.isLoaded Then
                syTempInfo.setAttribute("STAFFID", m_sWorkingUserid)
                syTempInfo.setAttribute("BRA_DEPNO", m_sWorkingTopDepNo)
                syTempInfo.setAttribute("FUNCCODE", m_sFuncCode)
            End If
            syTempInfo.setAttribute("TEMPDATA", com.Azion.EloanUtility.XmlUtility.convertDataTable2XML(dtSYTempInfo).ToString)
            syTempInfo.save()

            '更新ViewState("SY_TEMPINFO")
            ViewState.Remove("SY_TEMPINFO")
            ViewState("SY_TEMPINFO") = dtSYTempInfo

            ' 重新載入TempInfoTreeView數據
            If iRoleId <= 0 Then
                initTempInfoTreeView(treeView) '代表節點有被修改                
                btnDelete.Visible = True
            End If

            Dim bHasTempData As Boolean = bindTreeViewSYSForSYTempInfo(iRoleId)
            bindTreeViewSYSForSYRoleSuitSys(iRoleId, bHasTempData)

            ' 狀態為展開
            treeView.SelectedNode.ExpandAll()

            hidName.Value = txtName.Text.Trim
            hidId.Value = lblID.Text.Trim
            lblName.Text = txtName.Text.Trim
            btnAdd.Visible = True
            com.Azion.EloanUtility.UIUtility.alert("儲存成功！")

            If Not compareTreeViewSYS() Then
                If iRoleId <= 0 Then

                    treeView.SelectedNode.SelectAction = TreeNodeSelectAction.Select

                    If hidFlag.Value = "Add" Then
                        treeView.FindNode(ViewState("LastNodePath").ToString() & "/" & hidValue.Value).Selected = True
                    Else
                        treeView.FindNode(ViewState("LastNodePath").ToString()).Selected = True
                    End If

                    ' 重新給最後一次編輯的資料賦值
                    ViewState("LastNodePath") = treeView.SelectedNode.ValuePath
                End If
            Else
                If Not treeView.SelectedNode.Text.Replace("<font color='red'>", "").Replace("</font>", "").Contains(lblID.Text) Then
                    treeView.SelectedNode.Text = treeView.SelectedNode.Text.Replace("<font color='red'>", "").Replace("</font>", "") & "  (" & lblID.Text & ")"
                Else
                    treeView.SelectedNode.Text = treeView.SelectedNode.Text.Replace("<font color='red'>", "").Replace("</font>", "")
                End If
            End If

            If Not dtSYTempInfo Is Nothing Then
                For Each dr As DataRow In dtSYTempInfo.Rows
                    rdoDisable.SelectedValue = dr("DISABLED").ToString
                Next
            End If

            ViewState.Remove("hasSave")
            btnConfirmSend.Visible = True
            btnCancelAll.Visible = True
            hidFlag.Value = "Modify"
        Catch ex As Exception
            showErrMsg(ex)
        End Try
    End Sub

    ''' <summary>
    ''' 取消
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnCancel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnCancel.Click
        Dim iRoleId As Integer = lblID.Text

        Try
            If treeViewSys.CheckedNodes.Count = 0 Then
                display()
                bindBank()
            Else

                ' 組XML文件           
                Dim dtNewData As DataTable = genTempInfoTreeViewSYS2XML()

                If dtNewData Is Nothing Then
                    com.Azion.EloanUtility.UIUtility.alert("目前角色已恢復成未編輯前的狀態！")
                    Return
                End If

                Dim rowNewDatas() As DataRow = dtNewData.Select("ROLEID='" & lblID.Text & "'")
                ' 找到臨時檔中已經存在的資料
                ' 如果有刪除temp data
                If rowNewDatas.Length > 0 Then
                    dtNewData.Rows.Remove(rowNewDatas(0))
                    dtNewData.AcceptChanges()
                End If

                If dtNewData.Columns.Contains("iRoleId") Then dtNewData.Columns.Remove("iRoleId")

                Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager())

                If m_sCaseId = "" Then
                    syTempInfo.loadByPK(m_sWorkingUserid, m_sWorkingTopDepNo, m_sFuncCode)
                Else
                    syTempInfo.loadByCaseId(m_sCaseId)
                End If

                ViewState.Remove("SY_TEMPINFO")

                If syTempInfo.isLoaded Then
                    If dtNewData.Rows.Count > 0 Then
                        syTempInfo.setAttribute("TEMPDATA", com.Azion.EloanUtility.XmlUtility.convertDataTable2XML(dtNewData).ToString)
                        syTempInfo.save()
                        '更新ViewState("SY_TEMPINFO")
                        ViewState("SY_TEMPINFO") = dtNewData
                    Else
                        syTempInfo.remove()
                    End If
                End If

                For Each nodes As TreeNode In treeViewSys.Nodes
                    nodes.Checked = False
                    nodes.Text = nodes.Text.Replace("<font color='red'>", "").Replace("</font>", "")
                    For Each childNode As TreeNode In nodes.ChildNodes
                        childNode.Checked = False
                        childNode.Text = childNode.Text.Replace("<font color='red'>", "").Replace("</font>", "")
                    Next
                Next

                ' 重新載入TempInfoTreeView數據
                If iRoleId <= 0 Then

                    If iRoleId = 0 Then
                        display()
                        bindBank()
                    End If
                    initTempInfoTreeView(treeView) '代表節點有被修改             
                End If

                Dim bHasTempData As Boolean = bindTreeViewSYSForSYTempInfo(iRoleId)
                bindTreeViewSYSForSYRoleSuitSys(iRoleId, False)

                ' 狀態為展開
                treeView.SelectedNode.ExpandAll()

                treeView.SelectedNode.Text = treeView.SelectedNode.Text.Replace("<font color='red'>", "").Replace("</font>", "")

                com.Azion.EloanUtility.UIUtility.alert("目前角色已恢復成未編輯前的狀態！")

                btnConfirmSend.Visible = True
                btnCancelAll.Visible = True
                hidFlag.Value = "Modify"
            End If
        Catch ex As Exception
            showErrMsg(ex)
        End Try

    End Sub

    ''' <summary>
    ''' 刪除
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnDelete_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnDelete.Click
        Try
            Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager())
            Dim sNewXml As String = String.Empty

            ' 查詢臨時檔中的資料
            Dim dtData As DataTable = loadSYTempInfo()

            ' 查找子幾點，刪除子節點
            For Each childRow As DataRow In dtData.Select("[parent]='" & lblID.Text & "'")
                childRow.BeginEdit()
                'childRow("parent") = iNewFuncId
                childRow.Delete()
                childRow.EndEdit()
                dtData.AcceptChanges()
            Next

            ' 查找子幾點，刪除子節點
            For Each childRow As DataRow In dtData.Select("[roleid]='" & lblID.Text & "'")
                childRow.BeginEdit()
                'childRow("parent") = iNewFuncId
                childRow.Delete()
                childRow.EndEdit()
                dtData.AcceptChanges()
            Next

            sNewXml = com.Azion.EloanUtility.XmlUtility.convertDataTable2XML(dtData)

            sNewXml = "<SY>" & sNewXml.Replace("<DocumentElement>", "").Replace("</DocumentElement>", "") & "</SY>"

            If sNewXml.Contains("<DocumentElement />") Then
                sNewXml = ""
            End If

            If m_sFourStepNo = "" Then

                If syTempInfo.loadByPK(m_sWorkingUserid, m_sWorkingTopDepNo, m_sFuncCode) Then

                    If sNewXml = "" Then
                        syTempInfo.remove()
                    Else
                        syTempInfo.setAttribute("TEMPDATA", sNewXml)

                        syTempInfo.save()
                    End If
                End If
            End If

            ' 重新給臨時檔的ViewState賦值
            ViewState("SY_TEMPINFO") = dtData

            ViewState("LastNodePath") = Nothing

            treeView.Nodes.Clear()

            ' 重新加載
            initRoleTreeView(treeView)
            initTempInfoTreeView(treeView) '代表節點有被修改   

            treeView.ExpandAll()
            display()

            bindBank()

            hidName.Value = ""
            hidId.Value = ""
            hidParentId.Value = ""
            hidParentName.Value = ""
            hidFlag.Value = "Modify"
        Catch ex As Exception
            showErrMsg(ex)
        End Try
    End Sub

    ''' <summary>
    ''' 確認送出
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnConfirmSend_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnConfirmSend.Click
        Try
            Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager())
            Dim syRoleHis As New SY_ROLE_HIS(GetDatabaseManager())
            Dim flowFacade As New FLOW_OP.FlowFacade(GetDatabaseManager())
            Dim xmlDocument As New XmlDocument
            Dim xmlData As String = String.Empty
            Dim sFuncNameList As String = String.Empty
            Dim sRoleIdList As String = String.Empty

            ' 取得臨時檔中的資料
            If m_sFourStepNo = "" Then
                If syTempInfo.loadByPK(m_sWorkingUserid, m_sWorkingTopDepNo, m_sFuncCode) Then
                    xmlData = syTempInfo.getAttribute("TEMPDATA")
                End If
            Else
                If syTempInfo.loadByCaseId(m_sCaseId) Then
                    xmlData = syTempInfo.getAttribute("TEMPDATA")
                End If
            End If

            ' 沒有送審資料，提示
            If String.IsNullOrEmpty(xmlData) Then
                com.Azion.EloanUtility.UIUtility.alert("無送審資料，不能進行送出！")

                Return
            End If

            If Not xmlData Is Nothing Then

                ' document對象載入XML文件
                xmlDocument.LoadXml(xmlData)
            End If

            ' 找到臨時檔中已經存在的資料
            For i As Integer = 0 To xmlDocument.GetElementsByTagName("SY_ROLE").Count - 1
                If Convert.ToInt32(DirectCast(xmlDocument.GetElementsByTagName("SY_ROLE")(i)("ROLEID"), System.Xml.XmlElement).InnerText) < 0 Then

                    ' 名稱集合
                    sFuncNameList += "'" & DirectCast(xmlDocument.GetElementsByTagName("SY_ROLE")(i)("ROLENAME"), System.Xml.XmlElement).InnerText & "'" & ","
                Else
                    ' 編號集合
                    sRoleIdList += DirectCast(xmlDocument.GetElementsByTagName("SY_ROLE")(i)("ROLEID"), System.Xml.XmlElement).InnerText + ","
                End If
            Next

            If sFuncNameList.Length > 0 Then
                sFuncNameList = sFuncNameList.Substring(0, sFuncNameList.Length - 1)
            End If

            If sRoleIdList.Length > 0 Then
                sRoleIdList = sRoleIdList.Substring(0, sRoleIdList.Length - 1)
            End If

            ' 查詢資料是否在送簽中
            If syRoleHis.loadDataByCon(sFuncNameList, sRoleIdList, m_sStepNo, m_sCaseId) > 0 Then

                ' ROLENAME已在處理中，無法送件！
                com.Azion.EloanUtility.UIUtility.alert(lblName.Text & "已在處理中，無法送件！")
                Return
            End If

            GetDatabaseManager.beginTran()

            Dim stepInfo As FLOW_OP.StepInfo

            m_bCheck = True     '不再複核

            If m_bCheck = False Then
                If m_sStepNo = "" Then
                    ' stepInfo = com.Azion.UITools.UIShareFun.startFlow(GetDatabaseManager(), m_sWorkingBrid, m_sFlowName, True)
                    stepInfo = MBSC.UICtl.UIShareFun.startFlow(GetDatabaseManager(), m_sFlowName, True)

                    Dim sNewCaseId = stepInfo.currentStepInfo.caseId

                    '不可能沒資料 
                    If syTempInfo.loadByPK(m_sWorkingUserid, m_sWorkingTopDepNo, m_sFuncCode) Then
                        syTempInfo.setAttribute("CASEID", sNewCaseId)
                        syTempInfo.save()
                    End If
                    ' 跳轉到待處理頁面
                    'com.Azion.EloanUtility.UIUtility.Redirect("SY_CASELIST.aspx")

                Else
                    stepInfo = MBSC.UICtl.UIShareFun.pushFlow(GetDatabaseManager(), m_sCaseId)
                    com.Azion.EloanUtility.UIUtility.closeWindow()
                End If

                updatTemp2Live(GetDatabaseManager(), stepInfo, "")

                'com.Azion.EloanUtility.UIUtility.alert("存檔成功！")

                ' 成功提示信息
                com.Azion.EloanUtility.UIUtility.showJSMsg("案件編號：" + stepInfo.currentStepInfo.caseId + " 送出成功！")

                ' 跳轉到待辦事項頁面
                If String.IsNullOrEmpty(m_sStepNo) Then
                    com.Azion.EloanUtility.UIUtility.goMainPage("")
                Else
                    com.Azion.EloanUtility.UIUtility.closeWindow()
                End If


            Else
                stepInfo = MBSC.UICtl.UIShareFun.startFlow(GetDatabaseManager(), m_sFlowName, True, , True)
                updatTemp2Live(GetDatabaseManager(), stepInfo, "Y")

                treeView.Nodes.Clear()

                ' 添加節點【存在于Live表】
                initRoleTreeView(treeView)

                treeView.ExpandDepth = m_sTreeLevel

                ' 顯示成初始狀態
                'display()

                ' 成功提示信息
                com.Azion.EloanUtility.UIUtility.showJSMsg("案件編號：" + stepInfo.currentStepInfo.caseId + " 送出成功！")

                ' 跳轉到待辦事項頁面
                If String.IsNullOrEmpty(m_sStepNo) Then
                    com.Azion.EloanUtility.UIUtility.goMainPage("")
                Else
                    com.Azion.EloanUtility.UIUtility.closeWindow()
                End If

            End If

            GetDatabaseManager.commit()
        Catch ex As Exception
            GetDatabaseManager.Rollback()
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' 將選中的節點的顏色恢復成黑色
    ''' </summary>
    ''' <param name="nodes"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function resetTreeViewState(ByRef nodes As TreeNodeCollection) As String
        For Each node As TreeNode In nodes
            node.Text = node.Text.Replace("<font color='red'>", "").Replace("</font>", "")
            resetTreeViewState(node.ChildNodes)
        Next
    End Function

    ''' <summary>
    ''' 全部取消->OK
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnCancelAll_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnCancelAll.Click
        Try
            ' 開始事物
            GetDatabaseManager().beginTran()

            Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager())
            Dim syRoleHis As New AUTH_OP.SY_ROLE_HIS(GetDatabaseManager())
            Dim syRoleSuitSysHis As New AUTH_OP.SY_ROLESUITSYS_HIS(GetDatabaseManager())
            Dim flowFacade As New FLOW_OP.FlowFacade(GetDatabaseManager())

            ' 刪除sy_tempinfo
            If syTempInfo.loadByPK(m_sWorkingUserid, m_sWorkingTopDepNo, m_sFuncCode) Then
                syTempInfo.remove()
            End If

            ' 刪除HIS
            syRoleHis.delHisDataByCaseId(m_sCaseId)
            syRoleSuitSysHis.delHisDataByCaseId(m_sCaseId)

            ' 刪除流程資料
            flowFacade.DeleteFlow(m_sCaseId)

            ' 提交
            GetDatabaseManager().commit()

            ViewState.Remove("SY_TEMPINFO")

            ' 添加節點【存在于Live表】
            treeView.Nodes.Clear()
            initRoleTreeView(treeView)

            resetTreeViewSys()

            trRoleMgr.Visible = False
            trSubSys.Visible = False '所屬子系統
            trDisable.Visible = False '啟用狀態

            ' Modify by 2012/06/28
            'btnConfirmSend.Visible = False '確認送出
            'btnCancelAll.Visible = False '全部取消

            btnSave.Visible = False '儲存
            btnDelete.Visible = False '刪除
            btnCancel.Visible = False '取消           

            txtName.Visible = False
            lblName.Visible = True

            bindBank()

            treeView.Nodes(0).Select()

            'treeView.ExpandAll()
            treeView.ExpandDepth = m_sTreeLevel

            treeView.ExpandAll()
            resetTreeViewSys()

            hidName.Value = ""
        Catch ex As Exception

            ' 事物回滾
            GetDatabaseManager.Rollback()
            showErrMsg(ex)
        End Try
    End Sub

    ''' <summary>
    ''' 記錄“是否停用”欄位
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub rdoDisable_SelectedIndexChanged(sender As Object, e As EventArgs) Handles rdoDisable.SelectedIndexChanged
        If hidFlag.Value = "Add" Then

            txtName.Visible = True
            lblName.Visible = False
            For Each item As ListItem In rdoDisable.Items
                If item.Selected Then
                    item.Text = String.Format("<font color='red'>{0}</font>", item.Text)
                Else
                    item.Text = item.Text.Replace("<font color='red'>", "").Replace("</font>", "")
                End If
            Next
        ElseIf hidFlag.Value = "Modify" Then
            If hidOldData.Value <> "" Then
                ' If hidOldData.Value.Split(";")(3).ToString() <> rdoDisable.SelectedValue Then

                For Each item As ListItem In rdoDisable.Items
                    If item.Selected Then
                        item.Text = String.Format("<font color='red'>{0}</font>", item.Text)
                    Else
                        item.Text = item.Text.Replace("<font color='red'>", "").Replace("</font>", "")
                    End If
                Next
                'End If
            End If
        End If
    End Sub

    ''' <summary>
    ''' 名稱欄位
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub txtName_TextChanged(sender As Object, e As EventArgs) Handles txtName.TextChanged
        If hidFlag.Value = "Add" Then
            txtName.ForeColor = Drawing.Color.Red
            txtName.Visible = True
            lblName.Visible = False
        ElseIf hidFlag.Value = "Modify" Then

            If hidOldData.Value <> "" Then
                If hidOldData.Value.Split(";")(0).ToString() <> txtName.Text Then
                    txtName.ForeColor = Drawing.Color.Red
                Else
                    txtName.ForeColor = Drawing.Color.Black
                End If
            End If
        End If
    End Sub

    ''' <summary>
    ''' 發生於伺服器控制項從記憶體卸載時。
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub Page_Unload(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Unload
        'If m_sFourStepNo = "0400" OrElse (Not ViewState("currentFourStepNo") Is Nothing OrElse ("currentFourStepNo") = "0400") Then
        '    com.Azion.EloanUtility.UIUtility.setControlRead(Me)
        '    showErrMsg("目前流程在審核階段無法修改")
        'End If
    End Sub
#End Region

#Region "審核區塊Event"

    ''' <summary>
    ''' 所屬子系統修改前
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub treeViewSysBefore_TreeNodeCheckChanged(sender As Object, e As System.Web.UI.WebControls.TreeNodeEventArgs) Handles treeViewSysBefore.TreeNodeCheckChanged

        ' 記錄子項選中的個數，與父節點的子項的個數進行比較，如果相同，則父節點選中
        Dim iCount As Integer = 0

        ' 父節點選中，判斷是否有子項，子項也要選中
        If e.Node.Checked = True Then
            If e.Node.ChildNodes.Count > 0 Then
                For Each item As TreeNode In e.Node.ChildNodes
                    item.Checked = True
                Next
            End If
        ElseIf e.Node.Checked = False Then
            If e.Node.ChildNodes.Count > 0 Then
                For Each item As TreeNode In e.Node.ChildNodes
                    item.Checked = False
                Next
            End If
        End If

        ' 若同一级都没有选择，同时取消父节点的选择。
        If Not e.Node.Parent Is Nothing Then
            For Each node As TreeNode In e.Node.Parent.ChildNodes
                If node.Checked = True Then
                    iCount = iCount + 1
                    If iCount = e.Node.Parent.ChildNodes.Count Then
                        e.Node.Parent.Checked = True
                    End If
                Else
                    e.Node.Parent.Checked = False
                End If
            Next
        End If
    End Sub

    ''' <summary>
    ''' 所屬子系統修改后
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub treeViewSysAfter_TreeNodeCheckChanged(sender As Object, e As System.Web.UI.WebControls.TreeNodeEventArgs) Handles treeViewSysAfter.TreeNodeCheckChanged

        ' 記錄子項選中的個數，與父節點的子項的個數進行比較，如果相同，則父節點選中
        Dim iCount As Integer = 0

        ' 父節點選中，判斷是否有子項，子項也要選中
        If e.Node.Checked = True Then
            If e.Node.ChildNodes.Count > 0 Then
                For Each item As TreeNode In e.Node.ChildNodes
                    item.Checked = True
                Next
            End If
        ElseIf e.Node.Checked = False Then
            If e.Node.ChildNodes.Count > 0 Then
                For Each item As TreeNode In e.Node.ChildNodes
                    item.Checked = False
                Next
            End If
        End If

        ' 若同一级都没有选择，同时取消父节点的选择。
        If Not e.Node.Parent Is Nothing Then
            For Each node As TreeNode In e.Node.Parent.ChildNodes
                If node.Checked = True Then
                    iCount = iCount + 1
                    If iCount = e.Node.Parent.ChildNodes.Count Then
                        e.Node.Parent.Checked = True
                    End If
                Else
                    e.Node.Parent.Checked = False
                End If
            Next
        End If
    End Sub

    ''' <summary>
    ''' 同意->SendFlow
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnArgee_Click(sender As Object, e As EventArgs) Handles btnArgee.Click
        Try
            GetDatabaseManager.beginTran()

            Dim stepInfo As FLOW_OP.StepInfo

            stepInfo = MBSC.UICtl.UIShareFun.pushFlow(GetDatabaseManager(), m_sCaseId)

            updatTemp2Live(GetDatabaseManager(), stepInfo, "Y")

            GetDatabaseManager.commit()

            ' 關閉本頁面
            com.Azion.EloanUtility.UIUtility.closeWindow()
        Catch ex As Exception
            GetDatabaseManager.Rollback()
            showErrMsg(ex)
        End Try
    End Sub

    ''' <summary>
    ''' 不同意
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnNotArgee_Click(sender As Object, e As EventArgs) Handles btnNotArgee.Click
        Try
            GetDatabaseManager.beginTran()

            Dim stepInfo As FLOW_OP.StepInfo
            stepInfo = MBSC.UICtl.UIShareFun.pushFlow(GetDatabaseManager(), m_sCaseId)

            updatTemp2Live(GetDatabaseManager(), stepInfo, "N")
            GetDatabaseManager.commit()

            ' 關閉本頁面
            com.Azion.EloanUtility.UIUtility.closeWindow()
        Catch ex As Exception
            GetDatabaseManager.Rollback()
            showErrMsg(ex)
        End Try
    End Sub

    ''' <summary>
    ''' 修正補充
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnReviseFlow_Click(sender As Object, e As EventArgs) Handles btnReviseFlow.Click
        Try
            GetDatabaseManager.beginTran()
            Dim stepInfo As FLOW_OP.StepInfo
            stepInfo = MBSC.UICtl.UIShareFun.rollBack(GetDatabaseManager(), m_sCaseId)

            updatTemp2Live(GetDatabaseManager(), stepInfo, "")
            GetDatabaseManager.commit()

            ' 關閉本頁面
            'com.Azion.EloanUtility.UIUtility.closeWindow()
        Catch ex As Exception
            GetDatabaseManager.Rollback()
            showErrMsg(ex)
        End Try
    End Sub

    ''' <summary>
    ''' 更新
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub treeViewCheck_SelectedNodeChanged(sender As Object, e As EventArgs) Handles treeViewCheck.SelectedNodeChanged
        Try
            If treeViewCheck.SelectedNode.Value.Split(";")(0) <> "0" Then
                loadCheck(treeViewCheck.SelectedNode)
            Else

                ' 清空之前選中的資料
                For Each node As TreeNode In treeViewCheck.Nodes
                    node.Checked = False
                    If node.ChildNodes.Count > 0 Then
                        For i As Integer = 0 To node.ChildNodes.Count - 1
                            node.ChildNodes(i).Checked = False
                        Next
                    End If
                Next
            End If
        Catch ex As Exception
            showErrMsg(ex)
        End Try
    End Sub
#End Region


    Protected Sub tbRoleMgr_TextChanged(sender As Object, e As EventArgs) Handles tbRoleMgr.TextChanged
        UpdateRoleMgr()
    End Sub

    Protected Function GetRoleMgrId(ByVal iRoleId As Integer) As Integer
        Try
            Dim obj As Object
            obj = FLOW_OP.FlowFacade.getNewInstance(GetDatabaseManager).getSYRole.ExecuteScalar( _
                "select top 1 ROLEMGR from SY_ROLE where ROLEID = @ROLEID@", _
                "ROLEID", iRoleId)

            Return BosBase.CDbType(Of Integer)(obj, 0)
        Catch ex As Exception
            Throw
        End Try
    End Function



    Protected Sub UpdateRoleMgr()

        Dim nRoleId As Integer = 0
        Dim nRoleMgrId As Integer = 0
        Dim nRoleMgrId2 As Integer = 0

        Try
            Try
                nRoleId = CInt(lblID.Text)
            Catch ex As Exception
            End Try

            Try
                nRoleMgrId = CInt(tbRoleMgr.Text)
            Catch ex As Exception
            End Try

            Try
                nRoleMgrId2 = GetRoleMgrId(nRoleId)
            Catch ex As Exception
            End Try

            If nRoleMgrId <> nRoleMgrId2 Then
                tbRoleMgr.ForeColor = Drawing.Color.Red
                lbRoleMgr.ForeColor = Drawing.Color.Red
            Else
                tbRoleMgr.ForeColor = Drawing.Color.Black
                lbRoleMgr.ForeColor = Drawing.Color.Black
            End If

            If nRoleId = 0 Then
                tbRoleMgr.Text = ""
                lbRoleMgr.Text = ""
                Return
            End If

            lbRoleMgr.Text = ""

            Dim obj As Object
            obj = FLOW_OP.FlowFacade.getNewInstance(GetDatabaseManager).getSYRole.ExecuteScalar( _
                "select top 1 ROLENAME from SY_ROLE where ROLEID = @ROLEID@", _
                "ROLEID", nRoleMgrId)

            If IsDBNull(obj) OrElse IsNothing(obj) Then

                If nRoleMgrId = 0 Then
                    lbRoleMgr.Text = ""
                Else
                    lbRoleMgr.Text = "不存在"
                End If

                tbRoleMgr.Text = ""
                Return
            End If

            lbRoleMgr.Text = CStr(obj)

        Catch ex As Exception
            Throw
        End Try

    End Sub

End Class

