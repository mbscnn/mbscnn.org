''' <summary>
''' 程式說明：組織維護作業
''' 建立者：Lake
''' 建立日期：2012-05-16
''' </summary>

Imports com.Azion.EloanUtility
Imports com.Azion.NET.VB
Imports AUTH_OP
Imports AUTH_OP.TABLE
Imports MBSC.MB_OP
Imports MBSC.UICtl

Public Class SY_UPDBRANCH
    Inherits SYUIBase

    ' 安泰銀行UpCode
    Dim m_sUpCode2370 As String = String.Empty
    Dim m_sUpCode2366 As String = String.Empty
     
    ' 用於替代名稱中可能出現的“/”
    Public COL_DELIM As String = Microsoft.VisualBasic.ChrW(1)

#Region "PageLoad"
    ''' <summary>
    ''' 頁面加載
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>[Lake] 2012/05/16 Created</remarks>
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        ' 初始化參數
        initParas()

        If Not IsPostBack Then

            ' 是否是新增標識
            ViewState("IsAdd") = ""

            ViewState("Disabled") = ""

            ' 單位代碼
            ViewState("Brid") = ""

            ' 初始化數據
            initData()
        End If

        ' 單位下進入組織維護作業，且選擇“安泰銀行”，隱藏“新增"按鈕
        If m_sHoFlag <> "1" Then

            If tvBranch.SelectedNode.Value = "0" Then

                btnAdd.Visible = False
            Else
                btnAdd.Visible = True
            End If
        End If
 
    End Sub
#End Region

#Region "Function"
    ''' <summary>
    ''' 初始化參數
    ''' </summary>
    ''' <remarks>[Lake] 2012/05/16 Created</remarks>
    Sub initParas()

        ' 測試數據
        If m_bTesting OrElse Request("TESTMODE") = "1" Then
            m_bTesting = True

            Dim loginUser As New SY_LOGIN(GetDatabaseManager)
            loginUser.getUserInfo("904", "S000035")
            m_sHoFlag = 1
            m_bCheck = True     '不再複核
            'Dim a As New com.Azion.EloanUtility.StaffInfo
            'a.WorkingStaffid = "S000035"
            'a.LoginUserId = "S000035"
            'Session("StaffInfo") = a
            'FLOW_OP.ELoanFlow.m_callbackUserInfo = New ExportUserInfo

            'm_sWorkingUserid = "S000035"
            'm_sWorkingTopDepNo = "1"
            'm_sLoginUserid = "S000035"
            ''m_sDepNo = "1"
            'm_sCaseId = "SY0111019000066"
            m_sHoFlag = "1"
            m_bDisplayMode = False
        End If

        ' 取得CodeList中設定的參數值
        m_sUpCode2370 = com.Azion.EloanUtility.CodeList.getAppSettings("UpCode2370")
        m_sUpCode2366 = com.Azion.EloanUtility.CodeList.getAppSettings("UpCode2366")
  
    End Sub

    ''' <summary>
    ''' 初始化數據
    ''' </summary>
    ''' <remarks>[Lake] 2012/05/16 Created</remarks>
    Sub initData()
        Try
            ' 選擇“依單位”
            rdoBranch.Checked = True

            ' 隱藏刪除
            btnDelete.Visible = False

            ' 顯示儲存
            btnSave.Visible = False

            ' 隱藏單位代碼
            txtBrid.Enabled = False

            ' 顯示新增
            btnAdd.Visible = True

            ' 綁定一級單位到部門下拉選單
            bindFirstLevelBranchToDll()

            ' 綁定單位樹
            bindBranchByParentToTv()

            ' 綁定屬性列表
            bindPropertyList()

            ' 選擇第一個屬性
            rdolProperty.SelectedValue = 1

            ' 選擇“安泰銀行”
            tvBranch.Nodes(0).Select()

            ' 選中單位顯示“黑色”
            tvBranch.SelectedNodeStyle.ForeColor = Drawing.Color.Black

            ' 顯示第一個單位的詳細資料
            showDetail()
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 綁定屬性列表
    ''' </summary>
    ''' <remarks>[Lake] 2012/05/23</remarks>
    Sub bindPropertyList()
        Try
            Dim apCodeList As New AP_CODEList(GetDatabaseManager())

            If apCodeList.loadByUpcodeNotValue(m_sUpCode2370, "0") > 0 Then
                rdolProperty.DataSource = apCodeList.getCurrentDataSet()
                rdolProperty.DataValueField = "VALUE"
                rdolProperty.DataTextField = "TEXT"
                rdolProperty.DataBind()
            End If
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 取得第一層單位列表
    ''' </summary>
    ''' <remarks>[Lake] 2012/05/16 Created</remarks>
    Sub bindFirstLevelBranchToDll()
        Dim syBranchList As New AUTH_OP.SY_BRANCHList(GetDatabaseManager())

        If syBranchList.getBrancByParentForDdl("0", m_sWorkingBrid, m_sHoFlag, "T") Then
            ddlBranch.DataSource = syBranchList.getCurrentDataSet()
            ddlBranch.DataValueField = "BRA_DEPNO"
            ddlBranch.DataTextField = "BRCNAME"
            ddlBranch.DataBind()
        End If

        ddlBranch.Items.Insert(0, New ListItem("請選擇", "-1"))

        ' 非總管理處，下拉選單只顯示當前登陸者所在單位
        If m_sHoFlag <> "1" Then
            rdoAll.Enabled = False
            rdoBranch.Enabled = False

            If ddlBranch.Items.Count > 1 Then

                ' 顯示當前登陸者所在單位
                ddlBranch.SelectedIndex = 1
            End If
        End If
    End Sub

    ''' <summary>
    ''' 綁定第一級單位
    ''' </summary>
    ''' <remarks>[Lake] 2012/05/22 Created</remarks>
    Sub bindFirstLevelBranchToTv()
        Try
            Dim syBranchList As New AUTH_OP.SY_BRANCHList(GetDatabaseManager())
            Dim apCode As New AP_CODE(GetDatabaseManager())
            Dim dtBranch As New DataTable()

            tvBranch.Nodes.Clear()

            If apCode.loadByPK(m_sUpCode2366) Then
                ' root節點
                Dim rootNode As TreeNode = New TreeNode()
                Dim node As TreeNode

                rootNode.Value = apCode.getAttribute("VALUE").ToString().Trim()
                rootNode.Text = apCode.getAttribute("TEXT").ToString.Trim()

                If syBranchList.getBrancByParent("0", "T", "1") Then
                    dtBranch = syBranchList.getCurrentDataSet().Tables(0)

                    For Each dr As DataRow In dtBranch.Rows
                        node = New TreeNode()

                        node.Value = dr("BRA_DEPNO").ToString().Replace("/", COL_DELIM)
                        node.Text = dr("BRCNAME").ToString()

                        If hidAddNodeValuePath.Value.Split("/").Count >= 2 Then
                            If hidAddNodeValuePath.Value.Split("/")(1).Split(";")(1) = node.Value.Split(";")(1).Replace("/", COL_DELIM) Then
                                node.Selected = True
                            End If
                        End If

                        rootNode.ChildNodes.Add(node)
                    Next
                End If

                ' 添加根節點到單位樹
                tvBranch.Nodes.Add(rootNode)
            End If

            tvBranch.ExpandAll()
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 根據上級單位綁定所有下屬單位
    ''' </summary>
    ''' <param name="sParent"></param>
    ''' <remarks>[Lake] 2012/05/16 Created</remarks>
    Sub bindBranchByParentToTv()
        Try
            Dim syBranchList As New AUTH_OP.SY_BRANCHList(GetDatabaseManager())
            Dim apCode As New AP_CODE(GetDatabaseManager())
            Dim dtBranch As New DataTable()

            ' 清空單位樹節點
            tvBranch.Nodes.Clear()

            ' 取得根節點
            If apCode.loadByPK(m_sUpCode2366) Then

                ' root節點
                Dim rootNode As TreeNode = New TreeNode()
                Dim node As New TreeNode()

                rootNode.Value = apCode.getAttribute("VALUE").ToString().Trim()
                rootNode.Text = apCode.getAttribute("TEXT").ToString.Trim()

                ' 添加根節點到單位樹
                rootNode.Expanded = True
                tvBranch.Nodes.Add(rootNode)

                ' 選擇一個部門
                If ddlBranch.SelectedIndex <> 0 Then
                    Dim sBraDepNo() As String = ddlBranch.SelectedValue.Split(";")

                    node.Value = ddlBranch.SelectedValue.Replace("/", COL_DELIM)

                    If sBraDepNo(3) = "WRITE" Then
                        node.Text = sBraDepNo(1) & "(" & sBraDepNo(0) & ")" & "  " & sBraDepNo(2)
                    Else
                        node.Text = sBraDepNo(1) & "(" & sBraDepNo(0) & ")" & "  " & sBraDepNo(2) & "(" & sBraDepNo(3) & ")"
                    End If

                    ' 添加當前選擇單位到單位樹
                    node.Expanded = True
                    rootNode.ChildNodes.Add(node)

                    ' 綁定當前選擇單位的所有下屬單位
                    addNode(node)

                    node.Select()
                End If
            End If
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 綁定所有單位及其下屬單位到單位列表
    ''' </summary>
    ''' <remarks>[Lake] 2012/05/16 Created</remarks>
    Sub bindAllBranchToTv()
        Try
            Dim syBranchList As New AUTH_OP.SY_BRANCHList(GetDatabaseManager())
            Dim apCode As New AP_CODE(GetDatabaseManager())
            Dim dtBranch As New DataTable()

            ' 清空單位樹節點
            tvBranch.Nodes.Clear()

            ' 取得根節點
            If apCode.loadByPK(m_sUpCode2366) Then

                ' root節點
                Dim rootNode As TreeNode = New TreeNode()
                rootNode.Value = apCode.getAttribute("VALUE").ToString().Trim()
                rootNode.Text = apCode.getAttribute("TEXT").ToString.Trim()

                tvBranch.Nodes.Add(rootNode)

                ' 綁定節點的下級節點
                addNode(rootNode)
            End If

            ' 展開第一層單位
            tvBranch.ExpandDepth = 1
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 綁定節點的下級節點
    ''' </summary>
    ''' <param name="node">節點</param>
    ''' <param name="item">parent</param>
    ''' <remarks>[Lake] 2012/05/16 Created</remarks>
    Private Sub addNode(ByVal node As TreeNode)
        Try
            Dim syBranchList As New AUTH_OP.SY_BRANCHList(GetDatabaseManager())
            Dim dtBranch As New DataTable()

            ' 取得該部門的下級子部門
            If syBranchList.getBrancByParent(node.Value.Split(";")(0), "T", "1") Then
                dtBranch = syBranchList.getCurrentDataSet().Tables(0)

                For Each dr As DataRow In dtBranch.Rows
                    Dim subNode As TreeNode = New TreeNode()

                    subNode.Value = dr("BRA_DEPNO").ToString().Trim().Replace("/", COL_DELIM)
                    subNode.Text = dr("BRCNAME").ToString().Trim()

                    If hidAddNodeValuePath.Value <> "" AndAlso hidAddNodeValuePath.Value <> "0" Then
                        If hidAddNodeValuePath.Value.Split("/")(1).Split(";")(1) = node.Value.Split(";")(1).Replace("/", COL_DELIM) Then
                            node.Selected = True
                        End If
                    End If

                    node.ChildNodes.Add(subNode)

                    addNode(subNode)
                Next
            End If
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 顯示單位的詳細資料
    ''' </summary>
    ''' <remarks>[Lake] 2012/05/16 Created</remarks>
    Sub showDetail()
        Try
            Dim node As TreeNode = tvBranch.SelectedNode
            Dim syBranch As New AUTH_OP.SY_BRANCH(GetDatabaseManager())
            Dim syBranchMgrs As New AUTH_OP.SY_BRANCHMGRList(GetDatabaseManager())

            ' 點選根節點”安泰銀行“
            If tvBranch.SelectedNode.Value.Split(";")(0) = 0 Then

                ' 上層單位
                lblParentName.Text = ""

                ' 單位名稱
                lblBranchName.Text = node.Text

                ' 電話區號
                txtBraTel.Text = ""

                ' 電話號碼
                txtBrTel.Text = ""

                ' 地址
                brAddr.FirstLoad(GetDatabaseManager(), "", "", "", "", "", "1", "30")

                '單位代碼 
                txtBrid.Text = ""

                '管理單位
                txtMgrBrid.Text = ""
            Else

                ' 上層單位為“安泰銀行”
                If tvBranch.SelectedNode.Parent.Value.Split(";")(0) = "0" Then

                    ' 上層單位
                    lblParentName.Text = node.Parent.Text

                    ' 單位編號
                    txtBrid.Text = tvBranch.SelectedNode.Value.Split(";")(1)

                    ' 管理單位
                    'txtMgrBrid.Text = tvBranch.SelectedNode.Value.Split(";")(4)
                    '一開始自己管理自己
                    Dim sMgDepNo As String = tvBranch.SelectedNode.Value.Split(";")(0)
                    If syBranchMgrs.loadMgDepNo(tvBranch.SelectedNode.Value.Split(";")(0)) > 0 Then
                        sMgDepNo = ""
                        For Each row As DataRow In syBranchMgrs.getCurrentDataSet.Tables(0).Rows
                            sMgDepNo &= row("MGR_BRADEPNO").ToString() & ","
                        Next
                        If sMgDepNo.EndsWith(",") Then
                            sMgDepNo = sMgDepNo.Remove(sMgDepNo.Length - 1, 1)
                        End If
                    End If
                    txtMgrBrid.Text = sMgDepNo
                Else

                    ' 上層單位
                    lblParentName.Text = node.Parent.Value.Split(";")(2).Replace(COL_DELIM, "/")
                End If

                ' “單位名稱”可編輯
                If tvBranch.SelectedNode.Value.Split(";")(3) = "WRITE" Then
                    lblBranchName.Visible = False

                    txtBranchName.Visible = True

                    txtBranchName.Text = node.Value.Split(";")(2).Replace(COL_DELIM, "/")
                Else
                    lblBranchName.Text = node.Value.Split(";")(2).Replace(COL_DELIM, "/")
                End If

                If tvBranch.SelectedNode.Value.Split(";")(0) = 0 Then

                    txtBrid.Enabled = False
                    txtMgrBrid.Enabled = False
                ElseIf tvBranch.SelectedNode.Parent.Value.Split(";")(0) = "0" Then
                    txtBrid.Enabled = True
                    'txtMgrBrid.Enabled = True

                    txtBrid.Text = tvBranch.SelectedNode.Value.Split(";")(1)
                    If Not IsNothing(syBranchMgrs.getCurrentDataSet()) Then
                        syBranchMgrs.loadMgDepNo(tvBranch.SelectedNode.Value.Split(";")(0))
                    End If

                    ' 管理單位
                    'txtMgrBrid.Text = tvBranch.SelectedNode.Value.Split(";")(4)
                    '一開始自己管理自己
                    Dim sMgDepNo As String = tvBranch.SelectedNode.Value.Split(";")(0)
                    If IsNothing(syBranchMgrs.getCurrentDataSet()) Then
                        syBranchMgrs.loadMgDepNo(tvBranch.SelectedNode.Value.Split(";")(0))
                    End If

                    sMgDepNo = ""
                    For Each row As DataRow In syBranchMgrs.getCurrentDataSet.Tables(0).Rows
                        sMgDepNo &= row("MGR_BRADEPNO").ToString() & ","
                    Next
                    If sMgDepNo.EndsWith(",") Then
                        sMgDepNo = sMgDepNo.Remove(sMgDepNo.Length - 1, 1)
                    End If

                    txtMgrBrid.Text = sMgDepNo
                Else
                    txtBrid.Enabled = False
                    'txtMgrBrid.Enabled = False

                    ' 單位編號
                    txtBrid.Text = tvBranch.SelectedNode.Value.Split(";")(1)

                    ' 管理單位
                    ' 管理單位
                    'txtMgrBrid.Text = tvBranch.SelectedNode.Value.Split(";")(4)
                    '一開始自己管理自己
                    Dim sMgDepNo As String = tvBranch.SelectedNode.Value.Split(";")(0)
                    If IsNothing(syBranchMgrs.getCurrentDataSet()) Then
                        syBranchMgrs.loadMgDepNo(tvBranch.SelectedNode.Value.Split(";")(0))
                    End If
                    sMgDepNo = ""
                    For Each row As DataRow In syBranchMgrs.getCurrentDataSet.Tables(0).Rows
                        sMgDepNo &= row("MGR_BRADEPNO").ToString() & ","
                    Next
                    If sMgDepNo.EndsWith(",") Then
                        sMgDepNo = sMgDepNo.Remove(sMgDepNo.Length - 1, 1)
                    End If
                    txtMgrBrid.Text = sMgDepNo
                End If

                ' 查詢點選單位的資料
                If syBranch.loadByPK(node.Value.Split(";")(0)) Then
                    Dim oBraTel As Object = syBranch.getAttribute("BRATEL")
                    Dim oBrTel As Object = syBranch.getAttribute("BRTEL")
                    Dim oDisabled As Object = syBranch.getAttribute("DISABLED")
                    Dim oCity As Object = syBranch.getAttribute("BRCCITY")
                    Dim oTown As Object = syBranch.getAttribute("BRCAREA")
                    Dim oRoad As Object = syBranch.getAttribute("BRCADDR")
                    Dim sHoFlag As String = syBranch.getAttribute("HOFLAG").ToString()


                    If IsNothing(oCity) Then
                        oCity = ""
                    Else
                        oCity = oCity.ToString().Trim()
                    End If

                    If IsNothing(oTown) Then
                        oTown = ""
                    Else
                        oTown = oTown.ToString().Trim()
                    End If

                    If IsNothing(oRoad) Then
                        oRoad = ""
                    Else
                        oRoad = oRoad.ToString().Trim()
                    End If

                    ' 地址
                    brAddr.FirstLoad(GetDatabaseManager(), oCity, oTown, "", oRoad, "", "1", "30")

                    ' 電話區號
                    If Not String.IsNullOrEmpty(oBraTel) Then
                        txtBraTel.Text = oBraTel.ToString().Trim()
                    Else
                        txtBraTel.Text = ""
                    End If

                    ' 電話號碼
                    If Not String.IsNullOrEmpty(oBrTel) Then
                        txtBrTel.Text = oBrTel.ToString().Trim()
                    Else
                        txtBrTel.Text = ""
                    End If

                    ' 屬性
                    rdolProperty.SelectedValue = sHoFlag

                    ' 啟用狀態
                    If Not String.IsNullOrEmpty(oDisabled) Then
                        If oDisabled.ToString() = "0" Then
                            rdoDisable.SelectedValue = 0
                        ElseIf oDisabled.ToString() = "1" Then
                            rdoDisable.SelectedValue = 1
                        End If
                    End If
                End If
            End If

            ' 詳細資料部份顯示欄位
            initPageStyleSet()
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 修改后，刷新數據
    ''' </summary>
    ''' <remarks>[Lake] 2012/05/28 Created</remarks>
    Sub reFreshData()
        Try
            ' 當前新增單位的一級單位路徑
            Dim sFirstLevelNodeValuePath As String = String.Empty

            If txtBranchName.Visible = True Then
                txtBranchName.Visible = False
                lblBranchName.Visible = True
                lblParentName.Text = ""
            End If

            If rdoBranch.Checked Then

                btnSave.Visible = True
                btnDelete.Visible = True

                bindFirstLevelBranchToDll()
            Else
                bindFirstLevelBranchToTv()
            End If

            ' 非刪除操作
            If hidOperateFlag.Value <> "1" Then
                If hidAddNodeValuePath.Value.Split("/").Count = 2 OrElse hidAddNodeValuePath.Value.Split("/").Count = 1 Then

                    If rdoBranch.Checked Then
                        If hidBraDepNo.Value <> "" Then

                            ' 一級單位的下屬單位
                            ddlBranch.SelectedValue = hidBraDepNo.Value & ";" & txtBrid.Text.Trim() & ";" & txtBranchName.Text.Replace(" ", "") & ";" & "WRITE" ' & ";" & txtMgrBrid.Text.Trim()
                        Else
                            ddlBranch.SelectedValue = hidDdlBraDepNo.Value
                        End If

                        ' 重新綁定單位
                        bindBranchByParentToTv()

                        ' 展開全部單位節點
                        tvBranch.ExpandAll()
                    Else
                        tvBranch.FindNode(hidAddNodeValuePath.Value).Select()

                        ' 添加該節點的下級節點
                        addNode(tvBranch.SelectedNode)
                    End If
                Else
                    If rdoBranch.Checked Then
                        If hidDdlBraDepNo.Value <> "" Then
                            ddlBranch.SelectedValue = hidDdlBraDepNo.Value
                        End If

                        ' 重新綁定單位
                        bindBranchByParentToTv()

                        ' 展開全部單位節點
                        tvBranch.ExpandAll()
                    Else

                        ' 找到當前異動過單位的一級單位路勁
                        sFirstLevelNodeValuePath = "0/" & hidAddNodeValuePath.Value.Substring(2).Substring(0, hidAddNodeValuePath.Value.Substring(2).IndexOf("/"))

                        ' 選擇一級單位
                        tvBranch.FindNode(sFirstLevelNodeValuePath).Select()

                        ' 添加該節點的下級節點
                        addNode(tvBranch.SelectedNode)

                        ' 展開節點
                        tvBranch.SelectedNode.ExpandAll()
                    End If
                End If

                tvBranch.FindNode(hidAddNodeValuePath.Value).Select()
            Else
                If rdoBranch.Checked Then
                    If hidAddNodeValuePath.Value = "" Then
                        ddlBranch.SelectedIndex = 0
                    Else
                        ddlBranch.SelectedValue = hidDdlBraDepNo.Value
                    End If

                    ' 重新綁定單位
                    bindBranchByParentToTv()

                    ' 展開全部單位節點
                    tvBranch.ExpandAll()
                Else
                    If hidAddNodeValuePath.Value = "" Then
                        bindFirstLevelBranchToTv()
                    Else
                        ' 找到當前異動過單位的一級單位路勁
                        sFirstLevelNodeValuePath = "0/" & hidAddNodeValuePath.Value.Substring(2).Substring(0, hidAddNodeValuePath.Value.Substring(2).IndexOf("/"))

                        ' 選擇一級單位
                        tvBranch.FindNode(sFirstLevelNodeValuePath).Select()


                        ' 添加該節點的下級節點
                        addNode(tvBranch.SelectedNode)

                        ' 展開節點
                        tvBranch.SelectedNode.ExpandAll()
                    End If
                End If

                If hidAddNodeValuePath.Value = "" Then
                    tvBranch.Nodes(0).Select()
                Else
                    tvBranch.FindNode(hidAddNodeValuePath.Value.Substring(0, hidAddNodeValuePath.Value.LastIndexOf("/"))).Select()
                End If
            End If

            If Not tvBranch.SelectedNode Is Nothing Then

                ' 控制“刪除”按鈕的顯示及隱藏
                If tvBranch.SelectedNode.Value = "0" Then
                    txtBrid.Enabled = False
                    btnDelete.Visible = False
                    btnSave.Visible = False
                Else
                    If tvBranch.SelectedNode.Value.Split(";")(3) = "WRITE" Then

                        btnDelete.Visible = True
                    Else
                        btnDelete.Visible = False
                    End If
                End If
            End If

            hidOperateFlag.Value = ""
            hidBraDepNo.Value = ""

            ' 顯示詳細資料
            showDetail()
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 選擇“全部”或“依單位”時數據加載
    ''' </summary>
    ''' <remarks></remarks>
    Sub initChangeData()

        ' 按鈕及欄位的狀態設置
        rdolProperty.SelectedValue = 1
        rdoDisable.SelectedValue = 0
        btnDelete.Visible = False
        btnAdd.Visible = True
        lblBranchName.Visible = True
        txtBranchName.Visible = False
        txtBrid.Enabled = False

        ' 隱藏“儲存”按鈕
        btnSave.Visible = False

        If rdoAll.Checked Then

            ' 綁定第一級節點
            bindFirstLevelBranchToTv()

            ' 下拉選單為“請選擇”
            ddlBranch.SelectedIndex = 0

            ' 下拉選單不可用
            ddlBranch.Enabled = False
        Else
            bindFirstLevelBranchToDll()

            ' 選擇第一筆
            ddlBranch.SelectedIndex = 0

            tvBranch.Nodes(0).ChildNodes.Clear()

            ' 下拉選單可用
            ddlBranch.Enabled = True
        End If

        ' 選擇安泰銀行
        tvBranch.Nodes(0).Select()

        ' 選擇單位”黑色“顯示
        tvBranch.SelectedNodeStyle.ForeColor = Drawing.Color.Black

        ' 顯示節點資料
        showDetail()
    End Sub

    ''' <summary>
    ''' 取得傳入部門的所有下屬部門代碼
    ''' </summary>
    ''' <param name="sBraDepNo">傳入部門代碼</param>
    ''' <param name="sAllBraDepNo">所有子部門代碼</param>
    ''' <returns></returns>
    ''' <remarks>[Lake] 2012/06/26 Created</remarks>
    Public Function getSubBraDepNo(ByVal sBraDepNo As String, ByVal sAllBraDepNo As String) As String
        Try
            Dim syBranchList As New AUTH_OP.SY_BRANCHList(GetDatabaseManager())

            If syBranchList.getBrancByParent(sBraDepNo, "T", "1") Then
                For Each dr As DataRow In syBranchList.getCurrentDataSet().Tables(0).Rows
                    sAllBraDepNo = sAllBraDepNo & "','" & dr("BRA_DEPNO").ToString().Split(";")(0)

                    sAllBraDepNo = getSubBraDepNo(dr("BRA_DEPNO").ToString().Split(";")(0), sAllBraDepNo)
                Next
            End If

            Return sAllBraDepNo
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Function

    ''' <summary>
    ''' 新增部門歷史檔
    ''' </summary>
    ''' <param name="sBraDepNo">部門代碼</param>
    ''' <param name="sFlag">操作標識</param>
    ''' <param name="sCaseId">案件編號</param>
    ''' <param name="sStepNo"></param>
    ''' <param name="sSubFlowSeq"></param>
    ''' <param name="sSubFlowCount"></param>
    ''' <remarks></remarks>
    Sub insertSyBranchHis(ByVal sBraDepNo As String, ByVal sFlag As String, Optional ByVal sCaseId As String = "SY0000000000000", _
               Optional ByVal sStepNo As String = "00000000", Optional ByVal sSubFlowSeq As String = "0", _
               Optional ByVal sSubFlowCount As String = "0")
        Dim syBranchHis As New AUTH_OP.SY_BRANCH_HIS(GetDatabaseManager())

        ' 新增操作，取得最大部門代碼加1；非新增取得所選單位部門代碼 
        If sFlag = "I" Then
            sBraDepNo = hidBraDepNo.Value
        Else
            sBraDepNo = tvBranch.SelectedNode.Value.Split(";")(0)
        End If

        syBranchHis.loadByPK(sBraDepNo, sCaseId, sStepNo, sSubFlowSeq, sSubFlowCount)

        ' 新增操作，取得最大部門代碼加1；非新增取得所選單位部門代碼 
        If sFlag = "I" Then
            syBranchHis.setAttribute("BRA_DEPNO", sBraDepNo)
        Else
            syBranchHis.setAttribute("BRA_DEPNO", sBraDepNo)
        End If

        syBranchHis.setAttribute("CASEID", sCaseId)
        syBranchHis.setAttribute("STEP_NO", sStepNo)
        syBranchHis.setAttribute("SUBFLOW_SEQ", sSubFlowSeq)
        syBranchHis.setAttribute("SUBFLOW_COUNT", sSubFlowCount)

        ' 新增一級單位，取得輸入單位代碼；新增下級單位或者更新單位取得已有分行代碼
        syBranchHis.setAttribute("BRID", txtBrid.Text)


        ' 新增一級單位，取得輸入管理單位；新增下級單位或者更新單位取得已有分行代碼
        syBranchHis.setAttribute("MGR_BRID", txtMgrBrid.Text)

        ' 選擇“安泰銀行”
        If tvBranch.SelectedNode.Value.Split(";")(0) = 0 Then
            syBranchHis.setAttribute("PARENT", tvBranch.SelectedNode.Value.Split(";")(0))
        Else
            syBranchHis.setAttribute("PARENT", tvBranch.SelectedNode.Parent.Value.Split(";")(0))
        End If

        ' 更新操作
        If lblBranchName.Visible = True Then
            syBranchHis.setAttribute("BRCNAME", lblBranchName.Text)
        Else
            syBranchHis.setAttribute("BRCNAME", txtBranchName.Text.Trim())
        End If

        syBranchHis.setAttribute("BRCCITY", brAddr.getCityValue)
        syBranchHis.setAttribute("BRCAREA", brAddr.getAreaValue)
        syBranchHis.setAttribute("BRCADDR", brAddr.getRoad())
        syBranchHis.setAttribute("BRATEL", txtBraTel.Text.Trim())
        syBranchHis.setAttribute("BRTEL", txtBrTel.Text.Trim())
        syBranchHis.setAttribute("OPERATION", sFlag)
        syBranchHis.setAttribute("APPROVED", "Y")

        If rdoDisable.SelectedValue = 0 Then
            syBranchHis.setAttribute("DISABLED", "0")
        ElseIf rdoDisable.SelectedValue = 1 Then
            syBranchHis.setAttribute("DISABLED", "1")
        End If

        syBranchHis.setAttribute("HOFLAG", rdolProperty.SelectedValue)

        syBranchHis.save()
    End Sub

    ''' <summary>
    ''' 頁面呈現樣式設置
    ''' </summary>
    ''' <remarks>[Lake] 2012/07/16 Created</remarks>
    Sub initPageStyleSet()

        ' 安泰銀行，隱藏欄位顯示
        If tvBranch.SelectedNode.Value = "0" Then
            trBrid.Visible = False
            trMgrBrid.Visible = False
            trAddr.Visible = False
            trTel.Visible = False
            trProperty.Visible = False
            trStatus.Visible = False
            btnDelete.Visible = False
            btnSave.Visible = False
            btnCancel.Visible = False
        Else
            trBrid.Visible = True
            trMgrBrid.Visible = True
            trAddr.Visible = True
            trTel.Visible = True
            trProperty.Visible = True
            trStatus.Visible = True
            'btnDelete.Visible = True
            btnSave.Visible = True
            btnCancel.Visible = True
        End If
    End Sub

    ''' <summary>
    ''' 設置參數初始狀態
    ''' </summary>
    ''' <remarks>[Lake] 2012/07/16 Created</remarks>
    Sub setParasStatus()
        ViewState("IsAdd") = ""
        hidOperateFlag.Value = ""
        hidBraDepNo.Value = ""
    End Sub
#End Region

#Region "Event"
    ''' <summary>
    ''' 查詢所有單位
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>[Lake] 2012/05/16 Created</remarks>
    Protected Sub rdoAll_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles rdoAll.CheckedChanged
        Try
            txtBrid.Text = ""
            txtMgrBrid.Text = ""

            ' 選擇“全部”時數據加載
            initChangeData()

            ' 設置參數初始狀態
            setParasStatus()
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 選擇依單位
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>[Lake] 2012/05/18 Created</remarks>
    Protected Sub rdoBranch_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles rdoBranch.CheckedChanged
        Try
            txtBrid.Text = ""
            txtMgrBrid.Text = ""

            ' 選擇“依單位”時數據加載
            initChangeData()

            ' 設置參數初始狀態
            setParasStatus()
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 查詢某一個單位
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>[Lake] 2012/05/16 Created</remarks>
    Protected Sub ddlBranch_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles ddlBranch.SelectedIndexChanged
        Try
            txtBranchName.Visible = False
            lblBranchName.Visible = True

            ' 依單位分派人員
            If rdoBranch.Checked Then
                If ddlBranch.SelectedIndex <> 0 Then
                    btnAdd.Visible = True
                    btnSave.Visible = True

                    ViewState("Brid") = ddlBranch.SelectedItem.Value.Split(";")(1)

                    Dim syBranchMgrs As New AUTH_OP.SY_BRANCHMGRList(GetDatabaseManager())
                    Dim sMgDepNo As String = ddlBranch.SelectedItem.Value.Split(";")(0)
                    If syBranchMgrs.loadMgDepNo(ddlBranch.SelectedItem.Value.Split(";")(0)) > 0 Then
                        sMgDepNo = ""
                        For Each row As DataRow In syBranchMgrs.getCurrentDataSet.Tables(0).Rows
                            sMgDepNo &= row("MGR_BRADEPNO").ToString() & ","
                        Next
                        If sMgDepNo.EndsWith(",") Then
                            sMgDepNo = sMgDepNo.Remove(sMgDepNo.Length - 1, 1)
                        End If
                    End If

                    ViewState("MgrBrid") = sMgDepNo

                    ' 單位代碼，管理單位可以輸入
                    txtBrid.Enabled = True
                    'txtMgrBrid.Enabled = True

                    ' 記錄當前選擇一級單位Value值
                    hidDdlBraDepNo.Value = ddlBranch.SelectedValue

                    ' 重新綁定單位
                    bindBranchByParentToTv()

                    tvBranch.Nodes(0).ChildNodes(0).Select()
                Else
                    btnSave.Visible = False
                    btnDelete.Visible = False

                    tvBranch.Nodes(0).ChildNodes.Clear()

                    tvBranch.Nodes(0).Select()
                End If

                ' 選擇單位預設顯示“黑色”
                tvBranch.SelectedNodeStyle.ForeColor = Drawing.Color.Black

                ' 顯示節點資料
                showDetail()

                ' 記錄初始啟用狀態
                ViewState("Disabled") = rdoDisable.SelectedValue

                ' 設置參數初始狀態
                setParasStatus()

                If Not tvBranch.SelectedNode.Value = "0" Then
                    If ddlBranch.SelectedIndex <> 0 Then

                        ' 單位名稱可編輯
                        If tvBranch.SelectedNode.Value.Split(";")(3) = "WRITE" Then

                            ' 無組織代碼，可刪除
                            btnDelete.Visible = True
                        Else

                            ' 有組織代碼，不可刪除
                            btnDelete.Visible = False
                        End If
                    End If
                End If
            End If

            ' 單位下進入組織維護作業，且選擇“安泰銀行”，隱藏“新增"按鈕
            If m_sHoFlag <> "1" Then

                If tvBranch.SelectedNode.Value = "0" Then

                    btnAdd.Visible = False
                Else
                    btnAdd.Visible = True
                End If
            End If
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 選擇單位樹種某個單位時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>[Lake] 2012/05/16 Created</remarks>
    Protected Sub tvBranch_SelectedNodeChanged(ByVal sender As Object, ByVal e As EventArgs) Handles tvBranch.SelectedNodeChanged
        Try
            Dim syBranch As New AUTH_OP.SY_BRANCH(GetDatabaseManager())
            Dim syBranchMgrs As New AUTH_OP.SY_BRANCHMGRList(GetDatabaseManager())

            ' 顯示新增按鈕
            btnAdd.Visible = True

            ' 選擇"安泰銀行"
            If tvBranch.SelectedNode.Value.Split(";")(0) = "0" Then

                ' 隱藏“刪除”按鈕
                btnDelete.Visible = False

                ' 隱藏“儲存”
                btnSave.Visible = False

                ' 隱藏Labe顯示的單位名稱
                lblBranchName.Visible = True

                ' 顯示TextBox顯示的單位名稱
                txtBranchName.Visible = False

                ' 安泰銀行時，屬性選擇第一個
                rdolProperty.SelectedValue = 1
            Else

                ' 顯示新增、刪除
                btnSave.Visible = True
                btnDelete.Visible = True

                ' 單位名稱可編輯
                If tvBranch.SelectedNode.Value.Split(";")(3) = "WRITE" Then

                    ' 隱藏Labe顯示的單位名稱
                    lblBranchName.Visible = False

                    ' 顯示TextBox顯示的單位名稱
                    txtBranchName.Visible = True

                    ' 無組織代碼，可刪除
                    btnDelete.Visible = True
                Else
                    ' 顯示Labe顯示的單位名稱
                    lblBranchName.Visible = True

                    ' 隱藏TextBox顯示的單位名稱
                    txtBranchName.Visible = False

                    ' 有組織代碼，不可刪除
                    btnDelete.Visible = False
                End If

                ' 選擇一級單位，單位代碼和管理代碼文本框顯示
                If tvBranch.SelectedNode.Parent.Value.Split(";")(0) = "0" Then
                    ViewState("Brid") = tvBranch.SelectedNode.Value.Split(";")(1)

                    txtBrid.Enabled = True
                Else
                    txtBrid.Enabled = False
                End If

                Dim sMgDepNo As String = tvBranch.SelectedNode.Value.Split(";")(0)
                If syBranchMgrs.loadMgDepNo(tvBranch.SelectedNode.Value.Split(";")(0)) > 0 Then
                    sMgDepNo = ""
                    For Each row As DataRow In syBranchMgrs.getCurrentDataSet.Tables(0).Rows
                        sMgDepNo &= row("MGR_BRADEPNO").ToString() & ","
                    Next
                    If sMgDepNo.EndsWith(",") Then
                        sMgDepNo = sMgDepNo.Remove(sMgDepNo.Length - 1, 1)
                    End If
                End If

                ViewState("MgrBrid") = sMgDepNo
                txtMgrBrid.Text = ViewState("MgrBrid")
            End If

            ' 記錄選擇節點的路徑
            hidAddNodeValuePath.Value = tvBranch.SelectedNode.ValuePath

            ' 選擇“全部”，加載選擇單位下級單位
            If rdoAll.Checked Then
                If tvBranch.SelectedNode.ChildNodes.Count = 0 Then

                    ' 顯示當前選擇節點的子節點
                    addNode(tvBranch.SelectedNode)

                    tvBranch.SelectedNode.CollapseAll()
                End If
            End If

            ' 顯示選擇單位的詳細資料
            showDetail()

            ' 記錄初始啟用狀態
            ViewState("Disabled") = rdoDisable.SelectedValue

            ' 設置參數初始狀態
            setParasStatus()

            ' 單位下進入組織維護作業，且選擇“安泰銀行”，隱藏“新增"按鈕
            If m_sHoFlag <> "1" Then

                If tvBranch.SelectedNode.Value = "0" Then

                    btnAdd.Visible = False
                Else
                    btnAdd.Visible = True
                End If
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
    ''' <remarks>[Lake] 2012/05/16 Created</remarks>
    Protected Sub btnAdd_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnAdd.Click
        Try
            Dim syBranch As New AUTH_OP.SY_BRANCH(GetDatabaseManager())

            ' 沒有選擇單位，直接跳出
            If IsNothing(tvBranch.SelectedNode) Then
                Return
            End If

            ' 選擇單位非“安泰銀行”
            If tvBranch.SelectedNode.Value <> "0" Then

                ' 查詢啟用狀態
                syBranch.loadByPK(tvBranch.SelectedNode.Value.Split(";")(0))

                ' 單位停用，不能新增下級單位
                If syBranch.getAttribute("DISABLED") = "1" Then
                    com.Azion.EloanUtility.UIUtility.alert("單位已停用，不能新增下級單位！")

                    Return
                End If
            End If

            ' 按鈕及欄位的初始設置
            trBrid.Visible = True
            trMgrBrid.Visible = True
            trAddr.Visible = True
            trTel.Visible = True
            trProperty.Visible = True
            trStatus.Visible = True
            btnDelete.Visible = True
            btnSave.Visible = True
            btnCancel.Visible = True
            lblBranchName.Visible = False
            txtBranchName.Visible = True
            txtBranchName.Text = ""
            btnSave.Visible = True
            btnAdd.Visible = False
            btnDelete.Visible = False
            rdoDisable.SelectedValue = 0

            ' 新增標識
            ViewState("IsAdd") = "Y"

            ' 安泰銀行或第一層單位，則屬性欄位可以編輯，且屬性默認選擇第一個
            If tvBranch.SelectedNode.Value.Split(";")(0) = 0 Then
                rdolProperty.SelectedValue = 1
            Else
                hidDdlBraDepNo.Value = ddlBranch.SelectedValue
            End If

            ' 上層單位
            If tvBranch.SelectedNode.Value = "0" Then
                lblParentName.Text = tvBranch.SelectedNode.Text
            Else
                lblParentName.Text = tvBranch.SelectedNode.Value.Split(";")(2).Replace(COL_DELIM, "/")
            End If

            ' 取得新部門代碼
            hidBraDepNo.Value = syBranch.getMaxBraDepNo()

            ' 若新增一級單位
            If tvBranch.SelectedNode.Parent Is Nothing Then

                ' 單位代碼和管理單位可用
                txtBrid.Enabled = True
                'txtMgrBrid.Enabled = True
            Else
                ' 單位代碼和管理單位可用
                txtBrid.Enabled = False
                'txtMgrBrid.Enabled = False
            End If
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 刪除
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>[Lake] 2012/05/16 Created</remarks>
    Protected Sub btnDelete_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnDelete.Click
        Try
            Dim syBranch As New AUTH_OP.SY_BRANCH(GetDatabaseManager())
            Dim syBranchhis As New AUTH_OP.SY_BRANCH_HIS(GetDatabaseManager())
            Dim syFlowStep As New AUTH_OP.SY_FLOWSTEP(GetDatabaseManager())
            Dim syBranchMgrList As New AUTH_OP.SY_BRANCHMGRList(GetDatabaseManager())
            Dim syRelBranchUserList As New AUTH_OP.SY_REL_BRANCH_USERList(GetDatabaseManager())
            Dim stepInfo As New FLOW_OP.StepInfo()

            ' 操作標識
            hidOperateFlag.Value = "1"

            If tvBranch.SelectedNode.Parent.Value.Split(";")(0) = 0 Then

                ' 刪除一級節點，標識為空
                hidAddNodeValuePath.Value = ""
            Else

                ' 記錄當前選擇節點的ValuePath
                hidAddNodeValuePath.Value = tvBranch.SelectedNode.ValuePath
            End If

            If tvBranch.SelectedNode Is Nothing Then
                com.Azion.EloanUtility.UIUtility.alert("請選擇要刪除的單位！")

                Return
            End If

            ' 判斷該單位是否為其他單位管理單位
            If syBranchMgrList.loadBymgrBraDepNo(tvBranch.SelectedNode.Value.Split(";")(0)) Then
                com.Azion.EloanUtility.UIUtility.alert("該單位是其它單位的管理單位，不得刪除！")

                Return
            End If

            ' 該單位還有人員存在，不能刪除
            If syRelBranchUserList.getUserCountByBraDepNo(tvBranch.SelectedNode.Value.Split(";")(0)) Then
                com.Azion.EloanUtility.UIUtility.alert("此組織尚有人員編列，不得刪除！")

                Return
            End If

            ' 該單位還有下屬單位，不能刪除
            If tvBranch.SelectedNode.ChildNodes.Count > 0 Then
                com.Azion.EloanUtility.UIUtility.alert("此組織尚有子節點, 不得刪除！")

                Return
            End If

            ' 事務開始
            GetDatabaseManager().beginTran()

            ' 1.刪除BRANCH表
            syBranch.deleteByPk(tvBranch.SelectedNode.Value.Split(";")(0))

            m_bCheck = True     '不再複核

            ' 2.調用流程方法
            ' 測試模式
            If m_bCheck Then
                stepInfo = MBSC.UICtl.UIShareFun.startFlow(GetDatabaseManager(), m_sFlowName, True, Nothing, True)
            Else
                ' 正常模式
                stepInfo = MBSC.UICtl.UIShareFun.startFlow(GetDatabaseManager(), m_sFlowName, True, Nothing, True)
            End If

            ' 3.新增歷史表
            insertSyBranchHis(tvBranch.SelectedNode.Value, "D", stepInfo.currentStepInfo.caseId, stepInfo.currentStepInfo.stepNo, _
                                        stepInfo.currentStepInfo.subflowSeq, stepInfo.currentStepInfo.subflowCount)

            ' 事務提交
            GetDatabaseManager().commit()

            'com.Azion.EloanUtility.UIUtility.alert("刪除成功！")

            ' 成功提示信息
            com.Azion.EloanUtility.UIUtility.showJSMsg("案件編號：" + stepInfo.currentStepInfo.caseId + " 送出成功！")

            ' 跳轉到待辦事項頁面
            If String.IsNullOrEmpty(m_sStepNo) Then
                com.Azion.EloanUtility.UIUtility.goMainPage("")
            Else
                com.Azion.EloanUtility.UIUtility.closeWindow()
            End If


            ' 刷新數據
            reFreshData()

            ' 一級單位
            If (Not IsNothing(tvBranch.SelectedNode.Parent)) AndAlso tvBranch.SelectedNode.Parent.Value = "0" Then

                ' 記錄單位代碼
                ViewState("Brid") = tvBranch.SelectedNode.Value.Split(";")(1)
            End If

            ViewState("IsAdd") = ""
        Catch ex As Exception

            ' 事務回滾
            GetDatabaseManager().Rollback()
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 儲存
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>[Lake] 2012/05/16 Created</remarks>
    Protected Sub btnSave_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnSave.Click
        Try
            Dim syBranch As New AUTH_OP.SY_BRANCH(GetDatabaseManager())
            Dim syBranchhis As New AUTH_OP.SY_BRANCH_HIS(GetDatabaseManager())
            Dim syBranchMgr As New AUTH_OP.SY_BRANCHMGR(GetDatabaseManager())
            Dim syBranchMgrList As New AUTH_OP.SY_BRANCHMGRList(GetDatabaseManager())
            Dim stepInfo As New FLOW_OP.StepInfo()
            Dim sBraDepNo As String = String.Empty
            Dim sBrid As String = String.Empty
            Dim sMgrBrid As String = String.Empty
            Dim bIsEdit As Boolean = False
            Dim oControl As Control = Nothing
            Dim sMsg As String = String.Empty

            ' 臨時存儲管理單位
            Dim sTempMgrBrid As String = String.Empty

            ' 重組管理單位集合
            Dim sFinalMgrBrid As String = String.Empty

            ' 有效管理單位個數
            Dim iMgrBridNum As Integer = 1

            ' 操作標識
            Dim sOperation As String = String.Empty

            If tvBranch.SelectedNode Is Nothing Then
                tvBranch.Nodes(0).Selected = True
            End If

            ' 管理單位為空則默認為其自身
            If txtMgrBrid.Text.Trim() = "" Then

                ' 管理單位
                sFinalMgrBrid = hidBraDepNo.Value
            Else

                ' 替換全角符號
                sMgrBrid = txtMgrBrid.Text.Trim().Replace("，", ",")

                ' 如果包含多個管理單位
                If sMgrBrid.Contains(",") Then

                    ' 去掉重複的管理單位
                    For Each mgrBrid As String In sMgrBrid.Split(",")
                        If sTempMgrBrid <> mgrBrid Then
                            sFinalMgrBrid = sFinalMgrBrid & "," & mgrBrid
                        End If

                        sTempMgrBrid = mgrBrid
                    Next

                    iMgrBridNum = sFinalMgrBrid.Substring(1).Split(",").Count

                    ' 替換掉分隔符“,”為“','”
                    sFinalMgrBrid = sFinalMgrBrid.Substring(1).Replace(",", "','")
                Else
                    sFinalMgrBrid = txtMgrBrid.Text.Trim()
                End If

                ' 判斷輸入的管理單位是否有效
                If Convert.ToInt32(syBranch.existMgrBrid(sFinalMgrBrid)) <> iMgrBridNum Then
                    com.Azion.EloanUtility.UIUtility.alert("存在無效的管理單位！")

                    Return
                End If
            End If

            ' 取得單位代碼
            sBrid = txtBrid.Text

            ' 單位代碼的存在性檢核
            If txtBrid.Enabled Then

                ' 單位代碼有變更
                If ViewState("Brid").ToString() <> txtBrid.Text.Trim() Then

                    ' 判斷單位代碼不能重複
                    If syBranch.existBrid(txtBrid.Text.Trim()) Then
                        com.Azion.EloanUtility.UIUtility.alert("單位代碼已經存在！")

                        Return
                    End If
                End If
            End If

            ' 啟用狀態變更，更新其下屬單位
            If Not ((Not IsNothing(tvBranch.SelectedNode.Parent)) AndAlso tvBranch.SelectedNode.Parent.Value = "0") Then
                If ViewState("Disabled").ToString() = "1" Then

                    syBranch.loadByPK(tvBranch.SelectedNode.Parent.Value.Split(";")(0))

                    If syBranch.getAttribute("DISABLED") = "1" Then
                        com.Azion.EloanUtility.UIUtility.alert("上級單位停用，不能啟用該單位！")

                        rdoDisable.SelectedValue = "1"

                        Return
                    End If
                End If
            End If

            ' 依单位
            If rdoBranch.Checked Then

                ' 新增操作
                If ViewState("IsAdd").ToString() = "Y" Then
                    If tvBranch.SelectedNode.Value = "0" Then

                        ' 記錄下拉選單Value值
                        hidDdlBraDepNo.Value = hidBraDepNo.Value & ";" & txtBrid.Text & ";" & txtBranchName.Text.Replace(" ", "") _
                                             & ";" & "WRITE"
                    Else

                        ' 記錄下拉選單Value值
                        hidDdlBraDepNo.Value = ddlBranch.SelectedValue
                    End If
                Else

                    If (Not IsNothing(tvBranch.SelectedNode.Parent)) AndAlso tvBranch.SelectedNode.Parent.Value = "0" Then

                        If lblBranchName.Visible Then

                            ' 記錄下拉選單Value值
                            hidDdlBraDepNo.Value = ddlBranch.SelectedValue.Split(";")(0) & ";" & txtBrid.Text & _
                                ";" & ddlBranch.SelectedValue.Split(";")(2) & ";" & ddlBranch.SelectedValue.Split(";")(3)
                        Else

                            ' 記錄下拉選單Value值
                            hidDdlBraDepNo.Value = ddlBranch.SelectedValue.Split(";")(0) & ";" & txtBrid.Text & _
                                ";" & txtBranchName.Text.Replace(" ", "") & ";" & "WRITE"
                        End If
                    Else

                        ' 記錄下拉選單Value值
                        hidDdlBraDepNo.Value = ddlBranch.SelectedValue
                    End If
                End If
            End If

            ' 新增操作
            If ViewState("IsAdd").ToString() = "Y" Then

                bIsEdit = False

                ' 新增操作
                sOperation = "I"

                sBraDepNo = hidBraDepNo.Value

                If txtMgrBrid.Text.Trim() = "" Then
                    txtMgrBrid.Text = sBraDepNo
                End If

                If tvBranch.SelectedNode.Value = "0" Then
                    ' 單位代碼和管理單位可編輯
                    txtBrid.Enabled = True
                    'txtMgrBrid.Enabled = True
                Else
                    ' 單位代碼和管理單位不可編輯
                    txtBrid.Enabled = False
                    'txtMgrBrid.Enabled = False
                End If

                ' 記錄新增節點的ValuePath
                hidAddNodeValuePath.Value = tvBranch.SelectedNode.ValuePath & "/" & _
                    (sBraDepNo & ";" & txtBrid.Text & ";" & txtBranchName.Text.Replace(" ", "").Replace("/", COL_DELIM) & ";WRITE")
                '(sBraDepNo & ";" & txtBrid.Text & ";" & txtBranchName.Text.Replace(" ", "") & ";WRITE" & ";" & txtMgrBrid.Text.Trim())
            Else

                bIsEdit = True

                ' 新增操作
                sOperation = "U"

                sBraDepNo = tvBranch.SelectedNode.Value.Split(";")(0)

                If txtMgrBrid.Text.Trim() = "" Then
                    txtMgrBrid.Text = sBraDepNo
                End If

                Dim sBridTemp As String = String.Empty
                ' 記錄新增節點的ValuePath
                If tvBranch.SelectedNode.Value.Split(";")(1) <> txtBrid.Text Then
                    sBridTemp = txtBrid.Text
                Else
                    sBridTemp = tvBranch.SelectedNode.Value.Split(";")(1)
                End If

                ' EPSDEP為NULL，單位名稱不可編輯
                If lblBranchName.Visible Then

                    hidAddNodeValuePath.Value = tvBranch.SelectedNode.Parent.ValuePath & "/" & _
                    (tvBranch.SelectedNode.Value.Split(";")(0) & ";" & _
                     sBridTemp & ";" & lblBranchName.Text.Replace(" ", "").Replace("/", COL_DELIM) & ";" & tvBranch.SelectedNode.Value.Split(";")(3))
                Else
                    hidAddNodeValuePath.Value = tvBranch.SelectedNode.Parent.ValuePath & "/" & _
                    (tvBranch.SelectedNode.Value.Split(";")(0) & ";" & _
                     sBridTemp & ";" & txtBranchName.Text.Replace(" ", "").Replace("/", COL_DELIM) & ";" & tvBranch.SelectedNode.Value.Split(";")(3))
                End If
            End If

            ' 事務開始
            GetDatabaseManager.beginTran()

            ' 1.新增 or 更新部門表
            If Not syBranch.loadByPK(sBraDepNo) Then
                syBranch.setAttribute("BRA_DEPNO", sBraDepNo)
            End If

            syBranch.setAttribute("BRID", sBrid)
            '已將MGR_BRID另建一個新SY_BRANCHMGR Table
            '管理行可多個逗號隔開
            'syBranch.setAttribute("MGR_BRID", sMgrBrid)

            If txtBranchName.Visible = True Then
                syBranch.setAttribute("BRCNAME", txtBranchName.Text.Replace(" ", ""))
            End If

            If sOperation = "I" Then
                syBranch.setAttribute("PARENT", tvBranch.SelectedNode.Value.Split(";")(0))
            End If

            syBranch.setAttribute("BRCCITY", brAddr.getCityValue)
            syBranch.setAttribute("BRCAREA", brAddr.getAreaValue)
            syBranch.setAttribute("BRCADDR", brAddr.getRoad())
            syBranch.setAttribute("BRATEL", txtBraTel.Text.Trim())
            syBranch.setAttribute("BRTEL", txtBrTel.Text.Trim())
            syBranch.setAttribute("HOFLAG", rdolProperty.SelectedValue)

            If rdoDisable.SelectedValue = 0 Then
                syBranch.setAttribute("DISABLED", "0")
            ElseIf rdoDisable.SelectedValue = 1 Then
                syBranch.setAttribute("DISABLED", "1")
            End If

            syBranch.save()

            m_bCheck = True     '不再複核

            ' 2.走流程
            If Not m_bCheck Then

                ' 非測試模式
                stepInfo = MBSC.UICtl.UIShareFun.startFlow(GetDatabaseManager(), m_sFlowName, True, , True)
            Else

                ' 測試模式
                stepInfo = MBSC.UICtl.UIShareFun.startFlow(GetDatabaseManager(), m_sFlowName, True, , True)
            End If

            ' 更新操作，更新單位代碼或啟用狀態，並新增HIS表
            If ViewState("IsAdd").ToString() <> "Y" Then

                ' 停用/啟用單位操作
                Dim sAllBraDepNo As String = String.Empty

                ' 取得所有下屬單位
                sAllBraDepNo = getSubBraDepNo(tvBranch.SelectedNode.Value.Split(";")(0), sAllBraDepNo)

                If sAllBraDepNo <> "" Then
                    sAllBraDepNo = sAllBraDepNo.Substring(3)
                End If

                ' 修改一級單位单位代码，更新其下屬單位Brid
                If (Not tvBranch.SelectedNode.Parent Is Nothing) AndAlso tvBranch.SelectedNode.Parent.Value = "0" Then

                    If ViewState("Brid").ToString() <> txtBrid.Text.Trim().Replace(" ", "") Then
                        ' 更新SY_BRANCH表BRID欄位
                        syBranch.updateByBrid(txtBrid.Text.Trim(), ViewState("Brid").ToString())
                    End If
                End If

                ' 啟用狀態變更，更新其下屬單位
                If ViewState("Disabled").ToString() <> rdoDisable.SelectedValue.ToString() Then

                    If sAllBraDepNo <> "" Then

                        ' 更新SY_BRANCH表Disabled欄位
                        syBranch.updateDisabled(sAllBraDepNo, rdoDisable.SelectedValue.ToString())
                    End If
                End If

                ' 單位或啟用狀態有過變更
                If ViewState("Brid").ToString() <> txtBrid.Text.Trim().Replace(" ", "") OrElse ViewState("Disabled").ToString() <> rdoDisable.SelectedValue.ToString() Then

                    ' 新增HIS
                    If sAllBraDepNo <> "" Then

                        ' 1.取得所有下屬節點
                        sAllBraDepNo = sAllBraDepNo.Replace("','", ",")

                        ' 2.遍歷所有下屬單位，新增HIS
                        For Each braDepNo As String In sAllBraDepNo.Split(",")
                            Dim sMgrBraDepNo As String = String.Empty

                            ' (1)取得該單位的管理單位
                            If syBranchMgrList.loadMgDepNo(Convert.ToInt32(braDepNo)) > 0 Then
                                For Each dr As DataRow In syBranchMgrList.getCurrentDataSet().Tables(0).Rows
                                    sMgrBraDepNo = sMgrBraDepNo & "," & dr("MGR_BRADEPNO").ToString()
                                Next

                                If sMgrBraDepNo <> "" Then
                                    sMgrBraDepNo = sMgrBraDepNo.Substring(1)
                                End If
                            End If

                            ' 新增HIS表
                            syBranchhis.insertByParas(braDepNo, rdoDisable.SelectedValue, sMgrBraDepNo, stepInfo.currentStepInfo.caseId, stepInfo.currentStepInfo.stepNo, _
                                        stepInfo.currentStepInfo.subflowSeq, stepInfo.currentStepInfo.subflowCount)
                        Next
                    End If
                End If
            End If

            ' 3.新增歷史表
            insertSyBranchHis(tvBranch.SelectedNode.Value, sOperation, stepInfo.currentStepInfo.caseId, stepInfo.currentStepInfo.stepNo, _
                                        stepInfo.currentStepInfo.subflowSeq, stepInfo.currentStepInfo.subflowCount)

            'txtMgrBrid.Text.Trim()個號隔開多個管理行
            '已將MGR_BRID另建一個新SY_BRANCHMGR Table 
            If Not IsNothing(ViewState("MgrBrid")) Then
                If ViewState("MgrBrid").ToString() <> "" Then
                    For Each sBrMgr As String In ViewState("MgrBrid").ToString.Split(",") 'titan
                        If sBrMgr.Trim() <> "" Then
                            If IsNumeric(sBrMgr) Then
                                If syBranchMgr.loadByPK(syBranch.getInt("BRA_DEPNO"), sBrMgr) Then
                                    syBranchMgr.remove()
                                End If
                            End If
                        End If
                    Next
                End If
            End If

            For Each sBrMgr As String In txtMgrBrid.Text.Trim().ToString.Split(",") 'titan
                If sBrMgr.Trim() <> "" Then
                    If IsNumeric(sBrMgr) Then
                        syBranchMgr.loadByPK(syBranch.getInt("BRA_DEPNO"), sBrMgr)
                        syBranchMgr.setInt("BRA_DEPNO", syBranch.getInt("BRA_DEPNO"))
                        syBranchMgr.setInt("MGR_BRADEPNO", sBrMgr)
                        syBranchMgr.save()
                    End If
                End If
            Next

            ' 事務提交
            GetDatabaseManager.commit()

            'com.Azion.EloanUtility.UIUtility.alert("儲存成功！")

            ' 成功提示信息
            com.Azion.EloanUtility.UIUtility.showJSMsg("案件編號：" + stepInfo.currentStepInfo.caseId + " 送出成功！")

            ' 跳轉到待辦事項頁面
            If String.IsNullOrEmpty(m_sStepNo) Then
                com.Azion.EloanUtility.UIUtility.goMainPage("")
            Else
                com.Azion.EloanUtility.UIUtility.closeWindow()
            End If


            ' 單位下，修改一級單位的單位代碼
            If m_sHoFlag <> "1" Then
                If txtBrid.Enabled Then
                    m_sWorkingBrid = txtBrid.Text.Replace(" ", "")
                End If
            End If

            ' 刷新數據
            reFreshData()

            btnAdd.Visible = True
            btnDelete.Visible = True

            ' 如果一級單位
            If tvBranch.SelectedNode.Parent.Value = "0" Then

                ' 記錄單位代碼
                ViewState("Brid") = tvBranch.SelectedNode.Value.Split(";")(1)
            End If

            ' 清空管理單位
            ViewState("MgrBrid") = txtMgrBrid.Text.Trim()

            ' 清空新增標識
            ViewState("IsAdd") = ""

            ' 取得最新的啟用狀態
            ViewState("Disabled") = rdoDisable.SelectedValue
        Catch ex As Exception

            ' 事務回滾
            GetDatabaseManager.Rollback()

            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 取消
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>[Lake] 2012/05/16 Created</remarks>
    Protected Sub btnCancel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnCancel.Click
        Try
            ' 變量及頁面狀態初始化
            txtBranchName.Visible = False
            lblBranchName.Visible = True
            btnAdd.Visible = True
            btnDelete.Visible = True
            lblBranchName.Text = lblParentName.Text

            If tvBranch.SelectedNode.Value.Split(";")(0) <> "0" Then
                lblParentName.Text = tvBranch.SelectedNode.Parent.Text
            Else
                lblParentName.Text = ""
            End If

            ' 顯示當前選中單位資料或者顯示“安泰銀行”資料
            showDetail()

            ' 設置參數初始狀態
            setParasStatus()

            If Not tvBranch.SelectedNode.Value = "0" Then
                If ddlBranch.SelectedIndex <> 0 Then

                    ' 單位名稱可編輯
                    If tvBranch.SelectedNode.Value.Split(";")(3) = "WRITE" Then

                        ' 無組織代碼，可刪除
                        btnDelete.Visible = True
                    Else

                        ' 有組織代碼，不可刪除
                        btnDelete.Visible = False
                    End If
                End If
            End If
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub
#End Region
End Class