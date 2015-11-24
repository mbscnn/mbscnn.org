''' <summary>
''' 程式說明：組織人員分派
''' 建立者：Lake
''' 建立日期：2012-05-18
''' </summary>

Imports com.Azion.NET.VB
Imports MBSC.MB_OP
Imports System.Xml

Public Class SY_UPGBRANCHUSER
    Inherits SYUIBase

    ' 安泰銀行UpCode
    Dim m_sUpCode2366 As String = String.Empty
    ' 右側原始人員資料
    Dim m_dtFlow As New DataTable
 
#Region "PageLoad"
    ''' <summary>
    ''' 頁面加載
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>[Lake] 2012/05/18</remarks>
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ' 初始化參數
        initParas()

        ' 初次加載
        If Not IsPostBack Then

            ' 存儲人員變動單位及人員變動情況
            Dim dtChange As New DataTable()

            ' 異動過的數據集合
            Dim dtListBoxMove As New DataTable

            ViewState("dataFlow") = m_dtFlow

            dtChange.Columns.Add("VALUE", Type.GetType("System.String"))
            dtChange.Columns.Add("TEXT", Type.GetType("System.String"))

            ' 存儲人員變動單位及人員變動情況
            ViewState("dtChange") = dtChange

            dtListBoxMove.Columns.Add("VALUE", Type.GetType("System.String"))
            dtListBoxMove.Columns.Add("TEXT", Type.GetType("System.String"))
            dtListBoxMove.Columns.Add("OPERATION", Type.GetType("System.String"))

            ViewState("dtListBoxMove") = dtListBoxMove

            ' 取得當前登錄人員的異動單位資料
            getChangeStaffInfo(dtChange)

            ' 綁定正常數據or審核數據
            If m_sFourStepNo = "" OrElse m_sFourStepNo = "0300" Then

                ' 初始化數據
                initNormalData()

                If m_sHoFlag <> "1" Then

                    tvBranch.Nodes(0).ChildNodes(0).Select()
                Else

                    tvBranch.Nodes(0).Select()
                End If
            End If
        End If

        ' 審核頁面點擊同意單位，設置樣式
        If trCheckContent.Visible Then

            If tvBranchCheck.Nodes.Count > 0 Then

                ' 清空安泰銀行下節點
                tvBranchCheck.Nodes.Clear()
            End If

            ' 加載審核人員
            initCheckData()
        Else

            ' 正常模塊，點擊同一個單位時，異動人員變色
            If Not IsNothing(tvBranch.SelectedNode) Then

                ' 異動數據變紅
                editListItemColor()
                editMoweItemColor()
            End If
        End If

        ' 設置頁面為檢視狀態
        If m_bDisplayMode Then
            com.Azion.EloanUtility.UIUtility.setControlRead(Me)
        End If
    End Sub
#End Region

#Region "Function"

    ''' <summary>
    ''' 初始化參數
    ''' </summary>
    ''' <remarks>[Lake] 2012/05/18 Created</remarks>
    Sub initParas()
        Dim syBranch As New AUTH_OP.SY_BRANCH(GetDatabaseManager())
        Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager())
        Dim syFlowStep As New AUTH_OP.SY_FLOWSTEP(GetDatabaseManager())

        ' 流程編號
        Dim sStepNO As String = String.Empty

        ' 案件編號
        Dim sCaseID As String = String.Empty

        ' 取得所在區域的標識值
        If Request.QueryString("hoflag") <> Nothing Then
            m_sHoFlag = Request.QueryString("hoflag")
        Else
            Dim syfunctionCode As New AUTH_OP.SY_FUNCTION_CODE(GetDatabaseManager())

            ' 若能查詢出數據
            If syfunctionCode.getHoflag(m_sCaseId) Then
                m_sHoFlag = syfunctionCode.getAttribute("HOFLAG")
            End If
        End If
 
      
        ' 測試數據
        If m_bTesting Then
            m_sWorkingUserid = "S000035"
            m_sWorkingTopDepNo = "75"
            m_sLoginUserid = "S000035"
            m_sLoginTopDepNo = "75"
            m_sCaseId = "SY9251019000118"
            m_sFourStepNo = Right(m_sStepNo, 4)
            m_bDisplayMode = False
            m_sFunccode = "87"
            m_sWorkingBrid = "925"

            m_sFourStepNo = "0400"
        End If

        ' 根據流程編號控制頁面可編輯模式
        If m_sStepNo = String.Empty Then

            ' 取出流程步驟編號
            If syFlowStep.loadRelSYTEMPINFO(m_sWorkingUserid, m_sWorkingTopDepNo, m_sFuncCode, "1") Then
                sStepNO = syFlowStep.getAttribute("STEP_NO")

                ' 如果有步驟編號
                If sStepNO <> String.Empty Then

                    ' 如果步驟號位400
                    If sStepNO.Substring(sStepNO.Length - 4, 4) = "0400" Then
                        m_bDisplayMode = True
                    End If
                End If
            End If
        End If

        ' 控制正常模塊，審核模塊的顯示隱藏
        If m_sFourStepNo = "" OrElse m_sFourStepNo = "0300" Then
            trNormalCondition.Visible = True
            trNormalContent.Visible = True
            trHeader.Visible = False
            trCheckContent.Visible = False
        Else
            trNormalCondition.Visible = False
            trNormalContent.Visible = False
            trHeader.Visible = True
            trCheckContent.Visible = True
        End If

        ' 案號
        lblCaseId.Text = m_sCaseId

        ' 單位
        If syBranch.loadByCaseId(m_sCaseId) Then
            lblBranch.Text = syBranch.getAttribute("BRCNAME")
        End If

        ' 取得CodeList中設定的參數值
        m_sUpCode2366 = com.Azion.EloanUtility.CodeList.getAppSettings("UpCode2366")
 
    End Sub

    ''' <summary>
    ''' 初始化數據
    ''' </summary>
    ''' <remarks>[Lake] 2012/05/18 Created</remarks>
    Sub initNormalData()
        Dim syBranchList As New AUTH_OP.SY_BRANCHList(GetDatabaseManager())
        Dim dtSubBranch As New DataTable

        ' 綁定一級單位到部門下拉選單
        bindFirstLevelBranchToDll()

        ' 初始化單位樹
        initTvBranch()
    End Sub

    ''' <summary>
    ''' 取得當前登錄人員的異動人員資料，
    ''' 並按單位儲存到視圖中
    ''' </summary>
    ''' <remarks>[Lake] 2012/07/09 Created</remarks>
    Sub getChangeStaffInfo(ByVal dtChange As DataTable)
        Try
            Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager())
            Dim sStaffXML As String = String.Empty
            Dim xmlDocument As New XmlDocument
            Dim xmlNodeCount As Integer = 0
            Dim sTempBraDepNo As String = String.Empty
            Dim sTempBrach As String = String.Empty
            Dim sBraDepNo As String = String.Empty

            If m_sFourStepNo = "" Then

                ' 取得登錄人員對部門員工的調整信息，保存到ViewState("dtChange")
                If syTempInfo.loadByPK(m_sWorkingUserid, m_sWorkingTopDepNo, m_sFuncCode) Then

                    If Not String.IsNullOrEmpty(syTempInfo.getAttribute("TEMPDATA")) Then
                        sStaffXML = syTempInfo.getAttribute("TEMPDATA").ToString()
                    Else
                        Return
                    End If
                End If
            Else

                ' 取得登錄人員對部門員工的調整信息，保存到ViewState("dtChange")
                If syTempInfo.loadByCaseId(m_sCaseId) Then

                    If Not String.IsNullOrEmpty(syTempInfo.getAttribute("TEMPDATA")) Then
                        sStaffXML = syTempInfo.getAttribute("TEMPDATA").ToString()
                    Else
                        Return
                    End If
                End If
            End If

            ' 不存在人員變更資料，跳出
            If String.IsNullOrEmpty(sStaffXML) Then
                Return
            End If

            ' 加載XML資料
            xmlDocument.LoadXml(sStaffXML)

            ' 計算XML人員節的數量
            xmlNodeCount = xmlDocument.GetElementsByTagName("SY_REL_BRANCH_USER").Count - 1

            ' 存在人員變動
            If xmlNodeCount >= 0 Then

                ' 取得各部門的人員變動情況，並按部門儲存到DataTable中
                For index = 0 To xmlNodeCount
                    sBraDepNo = DirectCast(xmlDocument.GetElementsByTagName("SY_REL_BRANCH_USER")(index)("BRA_DEPNO"), System.Xml.XmlElement).InnerText

                    ' 取得相同部門人員
                    If sTempBraDepNo <> sBraDepNo Then

                        If sTempBrach <> "" Then
                            Dim drBranch As DataRow = dtChange.NewRow()
                            drBranch("VALUE") = sTempBraDepNo
                            drBranch("TEXT") = sTempBrach

                            dtChange.Rows.Add(drBranch)
                        End If

                        ' 最後一個單位只有一個人員變動時在此記錄，跳出循環后，會在下邊的方法中被添加到DataTable中
                        sTempBrach = xmlDocument.GetElementsByTagName("SY_REL_BRANCH_USER")(index).OuterXml
                        sTempBraDepNo = sBraDepNo
                    Else
                        sTempBrach = sTempBrach & xmlDocument.GetElementsByTagName("SY_REL_BRANCH_USER")(index).OuterXml
                    End If
                Next

                ' 最後一個單位代碼及其人員變動資料
                Dim drNewBranch As DataRow = dtChange.NewRow()
                drNewBranch("VALUE") = sBraDepNo
                drNewBranch("TEXT") = sTempBrach

                dtChange.Rows.Add(drNewBranch)
            End If

            ViewState("dtChange") = dtChange
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 新增歷史檔
    ''' </summary>
    ''' <param name="sBraDepNo">部門代碼</param>
    ''' <param name="sStaffId">員工編號</param>
    ''' <param name="sStepNo">流程步驟</param>
    ''' <param name="sSubFlowSeq"></param>
    ''' <param name="sSubFlowCount"></param>
    ''' <param name="sTempData">XML檔</param>
    ''' <param name="sOperateFlag">"Y"：同意，"N"“不同意，""：其他</param>
    ''' <param name="sCaseId"></param>
    ''' <remarks>[Lake] 2012/06/26 Created</remarks>
    Sub insertSyRelBranchUserHis(ByVal sBraDepNo As String, ByVal sStaffId As String, ByVal sOperation As String, _
               ByVal sOperateFlag As String, Optional ByVal sCaseId As String = "SY0000000000000", _
               Optional ByVal sStepNo As String = "00000000", Optional ByVal sSubFlowSeq As String = "0", _
               Optional ByVal sSubFlowCount As String = "0")
        Try
            Dim syRelBranchUserHis As New AUTH_OP.SY_REL_BRANCH_USER_HIS(GetDatabaseManager())

            syRelBranchUserHis.loadByPk(sStaffId, sBraDepNo, sCaseId, sStepNo, sSubFlowSeq, sSubFlowCount)

            syRelBranchUserHis.setAttribute("STAFFID", sStaffId)
            syRelBranchUserHis.setAttribute("BRA_DEPNO", sBraDepNo)
            syRelBranchUserHis.setAttribute("CASEID", sCaseId)
            syRelBranchUserHis.setAttribute("STEP_NO", sStepNo)
            syRelBranchUserHis.setAttribute("SUBFLOW_SEQ", sSubFlowSeq)
            syRelBranchUserHis.setAttribute("SUBFLOW_COUNT", sSubFlowCount)
            syRelBranchUserHis.setAttribute("OPERATION", sOperation)

            ' 操作功能
            If sOperateFlag <> "" Then
                syRelBranchUserHis.setAttribute("APPROVED", sOperateFlag)
            End If

            syRelBranchUserHis.save()
            syRelBranchUserHis.clear()
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

#End Region

#Region "正常模塊Function"
    ''' <summary>
    ''' 初始化單位樹
    ''' </summary>
    ''' <remarks>[Lake] 2012/06/26 Created</remarks>
    Sub initTvBranch()
        Dim apCode As New AP_CODE(GetDatabaseManager())
        Dim syBranchList As New AUTH_OP.SY_BRANCHList(GetDatabaseManager())
        Dim dtSubBranch As New DataTable

        ' 清空單位樹節點
        tvBranch.Nodes.Clear()

        ' 取得根節點
        If apCode.loadByPK(m_sUpCode2366) Then

            ' 安泰銀行節點
            Dim rootNode As TreeNode = New TreeNode()

            rootNode.Value = apCode.getAttribute("VALUE").ToString().Trim()
            rootNode.Text = apCode.getAttribute("TEXT").ToString.Trim()

            ' 添加根節點到單位樹
            tvBranch.Nodes.Add(rootNode)

            ' 登錄人員所在一級單位處於選擇裝填
            ddlBranch.SelectedValue = m_sWorkingTopDepNo.ToString() & ";" & m_sWorkingBrid

            ' 登錄人員所在一級單位添加到單位樹
            If ddlBranch.SelectedItem.Value <> "-1" Then
                tvBranch.Nodes(0).ChildNodes.Add(New TreeNode(ddlBranch.SelectedItem.Text, ddlBranch.SelectedValue))
            End If

            ' 登錄人員所在一級單位處於選擇狀態
            If tvBranch.Nodes(0).ChildNodes.Count > 0 Then
                tvBranch.Nodes(0).ChildNodes(0).Select()

                ' [正常模塊]人員列表
                initNormalStaff()
            End If

            ' 取得當前單位的所有子單位
            If (Not tvBranch.SelectedNode Is Nothing) AndAlso tvBranch.SelectedNode.Value.Split(";").Length > 1 Then
                If syBranchList.getAllSubBraDepNo(tvBranch.SelectedNode.Value.Split(";")(1), tvBranch.SelectedNode.Value.Split(";")(0)) Then
                    dtSubBranch = syBranchList.getCurrentDataSet().Tables(0)
                End If
            End If

            If dtSubBranch.Rows.Count > 0 Then

                ' 添加下屬單位
                addNode(tvBranch.SelectedNode, dtSubBranch)
            End If
        End If
    End Sub

    ''' <summary>
    ''' 取得一級單位，
    ''' 綁定到下拉選單
    ''' </summary>
    ''' <remarks>[Lake] 2012/05/18 Created</remarks>
    Sub bindFirstLevelBranchToDll()
        Dim syBranchList As New AUTH_OP.SY_BRANCHList(GetDatabaseManager())
        Dim syBranch As New AUTH_OP.SY_BRANCH(GetDatabaseManager())

        ' 取得一級單位的資料
        If syBranchList.getBrancByParentForDdl("0", m_sWorkingBrid, m_sHoFlag) Then
            ddlBranch.DataSource = syBranchList.getCurrentDataSet()
            ddlBranch.DataValueField = "BRA_DEPNO"
            ddlBranch.DataTextField = "BRCNAME"
            ddlBranch.DataBind()

            ddlBranch.Items.Insert(0, New ListItem("請選擇", "-1"))
        End If

        ' 非總管理處，下拉選單只顯示當前登陸者所在單位
        If m_sHoFlag <> "1" Then
            rdolNormalCondition.Enabled = False

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
        Dim syBranchList As New AUTH_OP.SY_BRANCHList(GetDatabaseManager())
        Dim apCode As New AP_CODE(GetDatabaseManager())
        Dim dtBranch As New DataTable()

        ' 清空安泰銀行下所有單位
        tvBranch.Nodes(0).ChildNodes.Clear()

        ' 取得一級單位資料，添加到安泰銀行下
        If syBranchList.getBrancByParent("0", "", "1") Then
            dtBranch = syBranchList.getCurrentDataSet().Tables(0)

            For Each dr As DataRow In dtBranch.Rows
                Dim node = New TreeNode()

                ' BRA_DEPNO;BRID格式
                node.Value = dr("BRA_DEPNO").ToString()
                node.Text = dr("BRCNAME").ToString()

                tvBranch.Nodes(0).ChildNodes.Add(node)
            Next
        End If

        tvBranch.ExpandAll()
    End Sub

    ''' <summary>
    ''' 給單位添加下級單位
    ''' </summary>
    ''' <param name="node">節點</param>
    ''' <param name="item">parent</param>
    ''' <remarks>[Lake] 2012/05/18 Created</remarks>
    Private Sub addNode(ByVal node As TreeNode, ByVal dtSubBranch As DataTable)
        Dim syBranchList As New AUTH_OP.SY_BRANCHList(GetDatabaseManager())

        ' 該部門存在子部門
        If dtSubBranch.Select("[PARENT]='" & node.Value.Split(";")(0) & "'").Count > 0 Then

            ' 循環其子部門添加到部門樹
            For Each dr As DataRow In dtSubBranch.Select("[PARENT]='" & node.Value.Split(";")(0) & "'")
                Dim subNode As TreeNode = New TreeNode()
                subNode.Text = dr("BRCNAME").ToString().Trim()
                subNode.Value = dr("BRA_DEPNO").ToString().Trim()

                node.ChildNodes.Add(subNode)

                If dtSubBranch.Rows.Count > 0 Then

                    addNode(subNode, dtSubBranch)
                End If
            Next
        End If
    End Sub

    ''' <summary>
    ''' 根據下拉選單選擇，綁定下屬單位
    ''' </summary>
    ''' <remarks>[Lake] 2012/06/26 Created</remarks>
    Sub bindBranchByParentToTv()
        Dim syBranchList As New AUTH_OP.SY_BRANCHList(GetDatabaseManager())
        Dim apCode As New AP_CODE(GetDatabaseManager())
        Dim dtSubBranch As New DataTable()

        ' 清空單位樹節點
        tvBranch.Nodes(0).ChildNodes.Clear()

        Dim node As New TreeNode()

        ' 選擇一個部門
        If ddlBranch.SelectedIndex <> 0 Then
            node.Value = ddlBranch.SelectedValue
            node.Text = ddlBranch.SelectedItem.Text

            ' 添加當前選擇單位到單位樹
            tvBranch.Nodes(0).ChildNodes.Add(node)

            ' 取得當前單位的所有子單位
            If syBranchList.getAllSubBraDepNo(node.Value.Split(";")(1), node.Value.Split(";")(0)) Then
                dtSubBranch = syBranchList.getCurrentDataSet().Tables(0)
            End If

            If dtSubBranch.Rows.Count > 0 Then

                ' 添加下屬單位
                addNode(node, dtSubBranch)
            End If
        End If
    End Sub

    ''' <summary>
    ''' 正常模塊人員加載
    ''' </summary>
    ''' <remarks>[Lake] 2012/06/25 Created</remarks>
    Sub initNormalStaff()

        ' 清空人員列表
        lstLeft.Items.Clear()
        lstRight.Items.Clear()

        ' 加載左右兩側人員列表
        getStaff(getNotInStaff(tvBranch.SelectedNode.Value.Split(";")(0)), "L")
        getStaff("", "R")
    End Sub

    ''' <summary>
    ''' 取得部門人員，若有過人員變動取XML中人員，若沒有取得DB中人員
    ''' </summary>
    ''' <param name="sBraDepNO">部門代碼</param>
    ''' <param name="sStaffId">人員列表</param>
    ''' <returns></returns>
    ''' <remarks>[Lake] 2012/06/26 Created</remarks>
    Private Function getStaffByBraDepNo(ByVal sBraDepNO As String, ByVal sStaffId As String) As Object
        Dim syUserList As New AUTH_OP.SY_USERList(GetDatabaseManager())

        ' 取得變動人員Table
        Dim dtChangeBranch As DataTable = getChangeBranch()

        ' 人員集合物件
        Dim oStaff As Object = Nothing

        ' 該單位存在人員變動
        If Not dtChangeBranch Is Nothing Then
            If dtChangeBranch.Rows.Count > 0 AndAlso dtChangeBranch.Select("BRA_DEPNO='" & sBraDepNO & "'").Count > 0 Then
                oStaff = dtChangeBranch.Select("BRA_DEPNO='" & sBraDepNO & "' AND OPERATION IN('I','N') AND STAFFID NOT IN ('" & sStaffId & "')")

                Dim dtNew As DataTable = dtChangeBranch.Clone()
                For Each dr As DataRow In oStaff
                    Dim drNew As DataRow = dtNew.NewRow()
                    drNew("BRA_DEPNO") = dr("BRA_DEPNO")
                    drNew("STAFFID") = dr("STAFFID")
                    drNew("NAME") = dr("NAME")
                    drNew("OPERATION") = dr("OPERATION")
                    dtNew.Rows.Add(drNew)
                Next

                oStaff = dtNew
            Else
                ' DB中取得某單位的人員
                If syUserList.getStaffByBraDepNo(sBraDepNO, sStaffId) Then
                    oStaff = syUserList.getCurrentDataSet.Tables(0)
                End If
            End If
        Else
            ' DB中取得某單位的人員
            If syUserList.getStaffByBraDepNo(sBraDepNO, sStaffId) Then
                oStaff = syUserList.getCurrentDataSet.Tables(0)
            End If
        End If

        Return oStaff
    End Function

    ''' <summary>
    ''' 取得要排除的人員
    ''' </summary>
    ''' <param name="sBraDepNo">部門代碼</param>
    ''' <remarks>[Lake] 2012/06/26 Created</remarks>
    Private Function getNotInStaff(ByVal sBraDepNo As String) As String

        ' 取得右側人員集合
        Dim oStaff As Object = getStaffByBraDepNo(sBraDepNo, "")

        Dim dtFlow As DataTable = ViewState("dataFlow")

        If dtFlow.Columns.Count = 0 Then
            dtFlow.Columns.Add("STAFFID", Type.GetType("System.String"))
            dtFlow.Columns.Add("NAME", Type.GetType("System.String"))
        End If

        If dtFlow.Rows.Count > 0 Then
            dtFlow.Rows.Clear()
        End If

        ' 要排除的人員
        Dim sStaffId As String = String.Empty

        If oStaff Is Nothing Then
            Return ""
        End If

        If CType(oStaff, DataTable).Columns.Count = 2 Then
            For Each dr As DataRow In CType(oStaff, DataTable).Rows
                sStaffId = sStaffId & "','" & dr("STAFFID").ToString().Substring(0, 7)

                Dim row As DataRow = dtFlow.NewRow
                row("STAFFID") = dr("STAFFID")
                row("NAME") = dr("NAME")

                dtFlow.Rows.Add(row)
            Next

            ViewState("dataFlow") = dtFlow
        Else
            For Each dr As DataRow In CType(oStaff, DataTable).Rows
                sStaffId = sStaffId & "','" & dr("STAFFID").ToString().Substring(0, 7)

                ' 若是沒有異動的資料
                If dr("OPERATION") = "N" Then
                    Dim row As DataRow = dtFlow.NewRow
                    row("STAFFID") = dr("STAFFID")
                    row("NAME") = dr("NAME")

                    dtFlow.Rows.Add(row)
                End If
            Next

            ViewState("dataFlow") = dtFlow
        End If

        If sStaffId <> "" Then
            sStaffId = sStaffId.Substring(3)
        End If

        Return sStaffId
    End Function

    ''' <summary>
    ''' 取得部門人員
    ''' </summary>
    ''' <param name="sBraDepNo">部門代碼;左側為其上級部門代碼，右側為本級部門代碼</param>
    ''' <param name="sStaffId">要排除人員清單</param>
    ''' <param name="sFlag">加載左側或右側人員標識，L：代表左側；R：代表右側</param>
    ''' <remarks>[Lake] 2012/05/26 Created</remarks>
    Sub getStaff(ByVal sStaffId As String, ByVal sFlag As String)
        Dim syUserList As New AUTH_OP.SY_USERList(GetDatabaseManager())

        ' 人員集合物件
        Dim oStaff As Object = Nothing

        ' 選擇安泰銀行，直接跳出
        If tvBranch.SelectedNode.Value.Split(";")(0) = 0 Then
            Return
        End If

        If sFlag = "R" Then

            ' 取得該部門人員，若有過人員變動取XML中人員，若沒有取得DB中人員
            oStaff = getStaffByBraDepNo(tvBranch.SelectedNode.Value.Split(";")(0), sStaffId)

            lstRight.DataSource = oStaff

            lstRight.DataValueField = "STAFFID"
            lstRight.DataTextField = "NAME"
            lstRight.DataBind()
        ElseIf sFlag = "L" Then

            ' 選擇一級單位
            If tvBranch.SelectedNode.Parent.Value.Split(";")(0) = 0 Then

                ' 取得所有人員，排除點選單位人員
                If syUserList.getAllStaff(sStaffId) Then
                    oStaff = syUserList.getCurrentDataSet()
                End If
            Else
                ' 取得該部門人員，若有過人員變動取XML中人員，若沒有取得DB中人員
                oStaff = getStaffByBraDepNo(tvBranch.SelectedNode.Parent.Value.Split(";")(0), sStaffId)
            End If

            lstLeft.DataSource = oStaff
            lstLeft.DataValueField = "STAFFID"
            lstLeft.DataTextField = "NAME"
            lstLeft.DataBind()
        End If
    End Sub

    ''' <summary>
    ''' 選入選出時存儲Xml到臨時檔(去除沒有異動過的數據)
    ''' </summary>
    ''' <param name="sInOrOut">選入還是選出</param>
    ''' <param name="dtMove">移動過的數據集合</param>
    ''' <remarks></remarks>
    Sub editInOutXml(ByVal sInOrOut As String, ByVal dtMove As DataTable)

        ' 聲明參數
        Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager())
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
            For i As Integer = 0 To XmlDocument.GetElementsByTagName("SY_REL_BRANCH_USER").Count - 1
                Dim sNAME As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_BRANCH_USER")(i)("NAME"), System.Xml.XmlElement).InnerText
                Dim STAFFID As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_BRANCH_USER")(i)("STAFFID"), System.Xml.XmlElement).InnerText
                Dim sBRA_DEPNO As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_BRANCH_USER")(i)("BRA_DEPNO"), System.Xml.XmlElement).InnerText
                Dim sOPERATION As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_BRANCH_USER")(i)("OPERATION"), System.Xml.XmlElement).InnerText

                ' 如果是選出
                If sInOrOut = "out" Then
                    ' 
                    For j As Integer = 0 To dtMove.Rows.Count - 1
                        If sOPERATION = "I" AndAlso tvBranch.SelectedNode.Value.Split(";")(0) = sBRA_DEPNO AndAlso dtMove.Rows(j)("VALUE") = STAFFID Then
                            iIndexDtToLeft = iIndexDtToLeft & j & ","
                            item = XmlDocument.GetElementsByTagName("SY_REL_BRANCH_USER")(i)
                            nodeList.Add(item)
                        End If
                    Next
                Else
                    ' 
                    For j As Integer = 0 To dtMove.Rows.Count - 1
                        If sOPERATION = "D" AndAlso tvBranch.SelectedNode.Value.Split(";")(0) = sBRA_DEPNO AndAlso dtMove.Rows(j)("VALUE") = STAFFID Then
                            iIndexDtToLeft = iIndexDtToLeft & j & ","
                            item = XmlDocument.GetElementsByTagName("SY_REL_BRANCH_USER")(i)
                            nodeList.Add(item)
                        End If
                    Next
                End If
            Next

            ' 如果有 則將其替換掉 然後累加上新的XML
            If nodeList.Count > 0 Then
                For i As Integer = 0 To nodeList.Count - 1
                    If Not nodeList.Item(i) Is Nothing Then
                        Dim oldXML As String = nodeList.Item(i).InnerXml.ToString()

                        If sInOrOut = "out" Then
                            sXmlData = sXmlData.Replace(oldXML.ToString(), "").Replace("<SY_REL_BRANCH_USER></SY_REL_BRANCH_USER>", "")
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
    ''' 查詢出沒有異動的數據
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getOLstBoxRright() As String
        Dim dtFlow As DataTable = ViewState("dataFlow")
        Dim sbXmlData As New StringBuilder

        ' 若已選擇清單中沒有資料
        If IsNothing(dtFlow) Then
            sbXmlData.Append("")
        Else
            ' 檢核是否有沒有異動的已選擇的資料
            If dtFlow.Rows.Count > 0 And lstRight.Items.Count > 0 Then

                ' 循環找出沒有異動的數據
                For Each item As ListItem In lstRight.Items
                    For i As Integer = 0 To dtFlow.Rows.Count - 1
                        If item.Value = dtFlow.Rows(i)("STAFFID").ToString() Then
                            sbXmlData.Append("<SY_REL_BRANCH_USER>")
                            sbXmlData.Append("<BRA_DEPNO>" & tvBranch.SelectedNode.Value.Split(";")(0) & "</BRA_DEPNO>")
                            sbXmlData.Append("<STAFFID>" & item.Value & "</STAFFID>")
                            sbXmlData.Append("<NAME>" + item.Text + "</NAME>")
                            sbXmlData.Append("<OPERATION>N</OPERATION>")
                            sbXmlData.Append("</SY_REL_BRANCH_USER>")
                        End If
                    Next
                Next
            End If

            ' 返回沒有異動的數據
            Return sbXmlData.ToString()
        End If
    End Function

    ''' <summary>
    ''' 取得變動部門資料
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>[Lake] 2012/06/25 Created</remarks>
    Public Function getChangeBranch() As DataTable
        Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager())
        Dim sStaffXML As String = String.Empty
        Dim dtChangeBranch As New DataTable

        If m_sFourStepNo = "" Then

            ' 取得登錄人員對部門員工的調整信息，保存到ViewState("dtChange")
            If syTempInfo.loadByPK(m_sWorkingUserid, m_sWorkingTopDepNo, m_sFuncCode) Then

                If Not String.IsNullOrEmpty(syTempInfo.getAttribute("TEMPDATA")) Then
                    sStaffXML = syTempInfo.getAttribute("TEMPDATA").ToString()
                End If
            End If
        Else

            ' 取得登錄人員對部門員工的調整信息，保存到ViewState("dtChange")
            If syTempInfo.loadByCaseId(m_sCaseId) Then

                If Not String.IsNullOrEmpty(syTempInfo.getAttribute("TEMPDATA")) Then
                    sStaffXML = syTempInfo.getAttribute("TEMPDATA").ToString()
                End If
            End If
        End If

        If sStaffXML <> "" Then
            ViewState("inOrOutXml") = sStaffXML
        Else
            ViewState("inOrOutXml") = ""
        End If

        If Not String.IsNullOrEmpty(sStaffXML) Then
            dtChangeBranch = com.Azion.EloanUtility.XmlUtility.convertXML2DataTable(sStaffXML)
        End If

        Return dtChangeBranch
    End Function

    ''' <summary>
    ''' 正常區塊取得XmlData
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getXmlData(ByVal syTempInfo As AUTH_OP.SY_TEMPINFO) As String
        ' 如果案件為空
        If m_sFourStepNo = "" Then

            If syTempInfo.loadByPK(m_sWorkingUserid, m_sWorkingTopDepNo, m_sFuncCode) Then
                Return syTempInfo.getAttribute("TEMPDATA")
            End If
        Else
            If syTempInfo.loadByCaseId(m_sCaseId) Then
                Return syTempInfo.getAttribute("TEMPDATA")
            End If
        End If
    End Function

    ''' <summary>
    ''' 頁面加載時ListBox中移動過的資料變紅
    ''' </summary>
    ''' <remarks></remarks>
    Sub editListItemColor()
        Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager())
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

            Dim listBridStaffid As New List(Of FLOW_OP.BRID_STAFFID)
            Dim drcBridStaffid As DataRowCollection
            Dim syTable As New FLOW_OP.TABLE.SY_REL_BRANCH_USER(GetDatabaseManager)


            ' 如果有相關人員信息
            For i As Integer = 0 To XmlDocument.GetElementsByTagName("SY_REL_BRANCH_USER").Count - 1
                Dim sSTAFFID As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_BRANCH_USER")(i)("STAFFID"), System.Xml.XmlElement).InnerText
                Dim sBRA_DEPNO As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_BRANCH_USER")(i)("BRA_DEPNO"), System.Xml.XmlElement).InnerText
                Dim sOPERATION As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_BRANCH_USER")(i)("OPERATION"), System.Xml.XmlElement).InnerText
                Dim sName As String = DirectCast(XmlDocument.GetElementsByTagName("SY_REL_BRANCH_USER")(i)("NAME"), System.Xml.XmlElement).InnerText

                ' 設定方式與當前角色相同
                If tvBranch.SelectedNode.Value.Split(";")(0) = sBRA_DEPNO Then

                    ' 如果是新增的資料
                    If sOPERATION = "I" Then
                        Dim item As ListItem = lstRight.Items.FindByValue(sSTAFFID)

                        If Not IsNothing(item) Then
                            item.Attributes.Add("style", "color:red")

                            '檢查是否已存在其它分行，如果有，顯示警告
                            drcBridStaffid = syTable.GetDataRowCollection(
                                                    "select distinct BU.STAFFID, US.USERNAME, BR.BRID, BR.BRCNAME " & vbCrLf & _
                                                    "  from SY_REL_BRANCH_USER BU " & vbCrLf & _
                                                    " inner join SY_USER US " & vbCrLf & _
                                                    "    on US.STAFFID = BU.STAFFID " & vbCrLf & _
                                                    " inner join SY_BRANCH BR " & vbCrLf & _
                                                    "    on BU.BRA_DEPNO = BR.BRA_DEPNO " & vbCrLf & _
                                                    " where BU.STAFFID = @STAFFID@ " & vbCrLf & _
                                                    "   and BR.BRID <> (select BRID from SY_BRANCH BR2 where BRA_DEPNO = @BRA_DEPNO@) " & vbCrLf,
                                                    "STAFFID", sSTAFFID,
                                                    "BRA_DEPNO", sBRA_DEPNO)

                            If IsNothing(drcBridStaffid) = False Then
                                For Each drBridStaffid As DataRow In drcBridStaffid
                                    listBridStaffid.Add(FLOW_OP.BRID_STAFFID.Pair(
                                                        drBridStaffid("BRID"), drBridStaffid("BRCNAME"),
                                                        drBridStaffid("STAFFID"), drBridStaffid("USERNAME")))
                                Next
                            End If

                        End If
                    End If

                    ' 如果是刪除的資料
                    If sOPERATION = "D" Then
                        Dim item As ListItem = lstLeft.Items.FindByValue(sSTAFFID)

                        If Not IsNothing(item) Then
                            item.Attributes.Add("style", "color:red")
                        End If
                    End If
                End If
            Next

            If listBridStaffid.Count > 0 Then
                Dim sb As New StringBuilder("下列行員已存在於其它分行：")

                For Each item As FLOW_OP.BRID_STAFFID In listBridStaffid
                    sb.Append(vbCrLf & item.Staffname & "(" & item.Staffid & ")  已存在於 " & item.Brcname & "(" & item.Brid & ")")
                Next

                com.Azion.EloanUtility.UIUtility.alert(sb.ToString)
                'SYUIBase.showErrMsg(Me, New Exception(sb.ToString))
            End If

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
                    Dim itemDelete As ListItem = lstLeft.Items.FindByValue(dtListBoxMove.Rows(i)("VALUE"))
                    itemDelete.Attributes.Add("style", "color:red")
                ElseIf dtListBoxMove.Rows(i)("OPERATION") = "I" Then
                    Dim item As ListItem = lstRight.Items.FindByValue(dtListBoxMove.Rows(i)("VALUE"))
                    item.Attributes.Add("style", "color:red")
                End If
            Next
        End If
    End Sub

    ''' <summary>
    ''' 根據m_bCheck，走確認送出操作
    ''' </summary>
    ''' <param name="dtChangeBranchStaff"></param>
    ''' <param name="sOperateFlag"></param>
    ''' <remarks>[Lake] 2012/06/19 Created</remarks>
    Sub AgreeOperate(ByVal dtChangeBranchStaff As DataTable, Optional ByVal sOperateFlag As String = "")
        Dim flowFacade As New FLOW_OP.FlowFacade(GetDatabaseManager())
        Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager())
        Dim syRelBranchUser As New AUTH_OP.TABLE.SY_REL_BRANCH_USER(GetDatabaseManager())
        Dim syRelRoleUser As New AUTH_OP.SY_REL_ROLE_USER(GetDatabaseManager())
        Dim syUserRoleAgent As New AUTH_OP.TABLE.SY_USERROLEAGENT(GetDatabaseManager())
        Dim syRelRoleUserHis As New AUTH_OP.SY_REL_ROLE_USER_HIS(GetDatabaseManager())
        Dim stepInfo As New FLOW_OP.StepInfo()

        ' XML中循環遍歷的部門代碼
        Dim sCurrentBraDepNo As String = String.Empty

        ' 當前部門及其所有下屬部門
        Dim sAllBraDepNo As String = String.Empty

        If Not m_bCheck Then

            ' 調用流程方法
            stepInfo = MBSC.UICtl.UIShareFun.pushFlow(GetDatabaseManager(), m_sCaseId, , , , False)
        Else
            stepInfo = MBSC.UICtl.UIShareFun.startFlow(GetDatabaseManager(), m_sFlowName, True, , True)
        End If

        ' 循環異動資料，新增異動部門人員
        For Each dr As DataRow In dtChangeBranchStaff.Rows

            ' 當前遍歷單位為遍歷過
            If sCurrentBraDepNo <> dr("BRA_DEPNO").ToString() Then
                'Dbg.Assert(dr("STAFFID").ToString() <> "S04592")



                ' 取得當前部門及其下屬部門
                Dim sChilds As String = getSubBraDepNo(dr("BRA_DEPNO").ToString(), sAllBraDepNo)

                If String.IsNullOrEmpty(sChilds) Then
                    sAllBraDepNo = dr("BRA_DEPNO").ToString()
                Else
                    sAllBraDepNo = dr("BRA_DEPNO").ToString() & "," & sChilds
                End If

                sAllBraDepNo = sAllBraDepNo.Replace(",,", ",")
                sCurrentBraDepNo = dr("BRA_DEPNO").ToString()
            End If



            ' 本部門刪除的人員
            If dr("OPERATION").ToString() = "D" Then

                ' 新增人員角色歷史檔
                syRelRoleUserHis.insertSyRelRoleUserHis(stepInfo.currentStepInfo.caseId, stepInfo.currentStepInfo.stepNo, _
                   stepInfo.currentStepInfo.subflowSeq, stepInfo.currentStepInfo.subflowCount, dr("STAFFID").ToString(), sAllBraDepNo)

                ' 刪除人員角色檔
                syRelRoleUser.deleteStaffByBraDepNoAndStaffId(sAllBraDepNo, dr("STAFFID").ToString())

                ' 刪除代理人資料
                syUserRoleAgent.deleteStaffByBraDepNoAndStaffId(dr("BRA_DEPNO").ToString(), dr("STAFFID").ToString())

                ' 刪除組織人員檔
                syRelBranchUser.deleteStaffByBraDepNoAndStaffId(sAllBraDepNo, dr("STAFFID").ToString())
            ElseIf dr("OPERATION").ToString() = "I" Then

                ' 新增部門員工表
                syRelBranchUser.setAttribute("BRA_DEPNO", dr("BRA_DEPNO").ToString())
                syRelBranchUser.setAttribute("STAFFID", dr("STAFFID").ToString())
                syRelBranchUser.save()
            End If

            ' 新增歷史檔
            insertSyRelBranchUserHis(dr("BRA_DEPNO").ToString(), dr("STAFFID").ToString(), dr("OPERATION").ToString(), sOperateFlag, _
                   stepInfo.currentStepInfo.caseId, stepInfo.currentStepInfo.stepNo, _
                   stepInfo.currentStepInfo.subflowSeq, stepInfo.currentStepInfo.subflowCount)

            If Not m_bCheck Then

                ' 刪除Temp資料
                syTempInfo.deleteByCaseID(m_sCaseId)
                syTempInfo.clear()
            Else

                ' 刪除Temp資料
                syTempInfo.DeleteByPK(m_sWorkingUserid, m_sWorkingTopDepNo, m_sFuncCode)
                syTempInfo.clear()
            End If
        Next
    End Sub

    ''' <summary>
    ''' 選擇一個單位或取消操作
    ''' </summary>
    ''' <remarks>[Lake] 2012/07/09 Created</remarks>
    Sub selectNodeOrCancel()
        Dim syBranchList As New AUTH_OP.SY_BRANCHList(GetDatabaseManager())

        ' 清空存儲絕對異動的數據
        Dim dtListBoxMove As DataTable = ViewState("dtListBoxMove")

        ' 一級單位的所有子部門集合
        Dim dtSubBranch As New DataTable()

        dtListBoxMove.Rows.Clear()
        ViewState("dtListBoxMove") = dtListBoxMove

        ' 清空人員列表
        lstLeft.Items.Clear()
        lstRight.Items.Clear()

        ' 選擇”安泰銀行“，跳出
        If tvBranch.SelectedNode.Value.Split(";")(0) = 0 Then
            Return
        End If

        ' “全部”，加載選擇的一級單位的下級單位
        If rdolNormalCondition.SelectedValue = "all" Then

            ' 選擇一級單位且沒有下屬單位
            If tvBranch.SelectedNode.Parent.Value.Split(";")(0) = 0 AndAlso tvBranch.SelectedNode.ChildNodes.Count = 0 Then

                If syBranchList.getAllSubBraDepNo(tvBranch.SelectedNode.Value.Split(";")(1), tvBranch.SelectedNode.Value.Split(";")(0)) Then
                    dtSubBranch = syBranchList.getCurrentDataSet().Tables(0)
                End If

                If dtSubBranch.Rows.Count > 0 Then

                    ' 顯示當前選擇節點的子節點
                    addNode(tvBranch.SelectedNode, dtSubBranch)

                    tvBranch.SelectedNode.CollapseAll()
                End If
            End If
        End If

        ' [正常模塊]人員列表
        initNormalStaff()

        ViewState("inOrOutXml") = ""

        ' 異動數據變紅
        editListItemColor()

        ' 異東數據變紅
        editMoweItemColor()

        ' 當前選擇單位變黑
        tvBranch.SelectedNodeStyle.ForeColor = Drawing.Color.Black
    End Sub
#End Region

#Region "審核模塊Function"
    ''' <summary>
    ''' 初始化審核數據
    ''' </summary>
    ''' <remarks>[Lake] 2012/05/18 Created</remarks>
    Sub initCheckData()
        Try
            ' 綁定有人員變動給的單位
            bindEditBranch()
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' [審核模塊] 加載人員列表
    ''' </summary>
    ''' <param name="node"></param>
    ''' <remarks>[Lake] 2012/05/23</remarks>
    Sub initStaffList(ByVal node As TreeNode)
        Dim syUser As New AUTH_OP.TABLE.SY_USER(GetDatabaseManager())
        Dim syUserList As New AUTH_OP.SY_USERList(GetDatabaseManager())
        Dim syTempIndo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager())
        Dim xmlDocument As New XmlDocument
        Dim xmlNodeCount As Integer = 0

        lstEditBefore.Items.Clear()
        lstEditAfter.Items.Clear()

        ' TempInfo中當前單位的人員變動情況
        Dim dtChange As DataTable = ViewState("dtChange")

        If Not IsNothing(node) Then
            If node.Value = 0 Then
                Return
            End If
        End If

        ' 對人員做過調整
        If Not IsNothing(dtChange) Then

            ' 篩選當前點選單位Temp資料
            Dim drChangeBranch() As DataRow = dtChange.Select("VALUE='" & node.Value & "'")

            ' 該單位有人員異動
            If drChangeBranch.Count = 1 Then
                If Not String.IsNullOrEmpty(drChangeBranch(0)("TEXT")) Then
                    Dim sTempData As String = drChangeBranch(0)("TEXT").ToString()
                    Dim sTempDepNo As String = String.Empty
                    Dim sBraDepNo As String = String.Empty
                    Dim sName As String = String.Empty

                    Dim listBridStaffid As New List(Of FLOW_OP.BRID_STAFFID)
                    Dim drcBridStaffid As DataRowCollection
                    Dim syTable As New FLOW_OP.TABLE.SY_REL_BRANCH_USER(GetDatabaseManager)

                    ' 記錄未變動人員及新增人員編號
                    Dim sTempStaffId As String = String.Empty

                    ' document對象載入XML文件
                    xmlDocument.LoadXml("<SY>" & sTempData & "</SY>")

                    xmlNodeCount = xmlDocument.GetElementsByTagName("SY_REL_BRANCH_USER").Count - 1

                    ' "審核模塊"人員加載
                    For index = 0 To xmlNodeCount
                        Dim sStaffId As String = DirectCast(xmlDocument.GetElementsByTagName("SY_REL_BRANCH_USER")(index)("STAFFID"), System.Xml.XmlElement).InnerText
                        sBraDepNo = DirectCast(xmlDocument.GetElementsByTagName("SY_REL_BRANCH_USER")(index)("BRA_DEPNO"), System.Xml.XmlElement).InnerText
                        sName = DirectCast(xmlDocument.GetElementsByTagName("SY_REL_BRANCH_USER")(index)("NAME"), System.Xml.XmlElement).InnerText
                        Dim sOperation As String = DirectCast(xmlDocument.GetElementsByTagName("SY_REL_BRANCH_USER")(index)("OPERATION"), System.Xml.XmlElement).InnerText

                        ' 沒有異動過人員加載
                        If sOperation = "N" Then
                            lstEditBefore.Items.Add(New ListItem(sName, sStaffId))
                            lstEditAfter.Items.Add(New ListItem(sName, sStaffId))
                        ElseIf sOperation = "I" Then

                            ' 新增人員加載到”修改後“列表
                            lstEditAfter.Items.Add(New ListItem(sName, sStaffId))

                            lstEditAfter.Items.FindByValue(sStaffId).Attributes.Add("style", "color:Red;")

                            '檢查是否已存在其它分行，如果有，顯示警告
                            drcBridStaffid = syTable.GetDataRowCollection(
                                                    "select distinct BU.STAFFID, US.USERNAME, BR.BRID, BR.BRCNAME " & vbCrLf & _
                                                    "  from SY_REL_BRANCH_USER BU " & vbCrLf & _
                                                    " inner join SY_USER US " & vbCrLf & _
                                                    "    on US.STAFFID = BU.STAFFID " & vbCrLf & _
                                                    " inner join SY_BRANCH BR " & vbCrLf & _
                                                    "    on BU.BRA_DEPNO = BR.BRA_DEPNO " & vbCrLf & _
                                                    " where BU.STAFFID = @STAFFID@ " & vbCrLf & _
                                                    "   and BR.BRID <> (select BRID from SY_BRANCH BR2 where BRA_DEPNO = @BRA_DEPNO@) " & vbCrLf,
                                                    "STAFFID", sStaffId,
                                                    "BRA_DEPNO", sBraDepNo)

                            If IsNothing(drcBridStaffid) = False Then
                                For Each drBridStaffid As DataRow In drcBridStaffid
                                    listBridStaffid.Add(FLOW_OP.BRID_STAFFID.Pair(
                                                        drBridStaffid("BRID"), drBridStaffid("BRCNAME"),
                                                        drBridStaffid("STAFFID"), drBridStaffid("USERNAME")))
                                Next
                            End If

                        ElseIf sOperation = "D" Then

                            ' 刪除人員加載到”修改前“列表
                            lstEditBefore.Items.Add(New ListItem(sName, sStaffId))

                            lstEditBefore.Items.FindByValue(sStaffId).Attributes.Add("style", "color:Red;")
                        End If
                    Next

                    If listBridStaffid.Count > 0 Then
                        Dim sb As New StringBuilder("下列行員已存在於其它分行：")

                        For Each item As FLOW_OP.BRID_STAFFID In listBridStaffid
                            sb.Append(vbCrLf & item.Staffname & "(" & item.Staffid & ")  已存在於 " & item.Brcname & "(" & item.Brid & ")")
                        Next

                        com.Azion.EloanUtility.UIUtility.alert(sb.ToString)
                        'SYUIBase.showErrMsg(Me, New Exception(sb.ToString))
                    End If

                End If
            End If
        End If
    End Sub

    ''' <summary>
    ''' 取得異動部門的父子關係表
    ''' </summary>
    ''' <param name="sChangeBraDepNo"></param>
    ''' <returns></returns>
    ''' <remarks>[Lake] 2012/06/15</remarks>
    Private Function getFCNode(ByVal dtBranch As DataTable, ByVal sbraDepNo As String) As DataTable
        Try
            Dim syBranch As New AUTH_OP.SY_BRANCH(GetDatabaseManager())
            Dim drNew As DataRow = dtBranch.NewRow()

            If sbraDepNo <> "" Then
                If syBranch.getBrancByBraDepNo(sbraDepNo) Then

                    If dtBranch.Select("BRADEPNO='" & sbraDepNo & "'").Count = 0 Then
                        drNew("BRADEPNO") = syBranch.getAttribute("BRA_DEPNO")
                        drNew("BRCNAME") = syBranch.getAttribute("BRCNAME")
                        drNew("PARENT") = syBranch.getAttribute("PARENT")

                        dtBranch.Rows.Add(drNew)
                    End If

                    If syBranch.getAttribute("PARENT").ToString() <> "0" Then
                        dtBranch = getFCNode(dtBranch, syBranch.getAttribute("PARENT").ToString())
                    End If
                End If
            End If

            Return dtBranch
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Function

    ''' <summary>
    ''' [審核模塊]添加父子節點
    ''' </summary>
    ''' <param name="node"></param>
    ''' <param name="dtBranch"></param>
    ''' <param name="sBraDepNo"></param>
    ''' <remarks></remarks>
    Sub addNodeToTv(ByVal node As TreeNode, ByVal dtBranch As DataTable, ByVal sBraDepNo As String)
        Dim dtChange As DataTable = ViewState("dtChange")

        For Each dr As DataRow In dtBranch.Select("[PARENT]='" & sBraDepNo & "'")
            Dim newNode As New TreeNode()
            newNode.Text = dr("BRCNAME").ToString()
            newNode.Value = dr("BRADEPNO").ToString()

            ' 一級單位展開
            If sBraDepNo = "0" Then
                newNode.Expanded = True
            End If

            node.ChildNodes.Add(newNode)

            ' 存在人員異動的部門變紅
            If dtChange.Select("VALUE='" & newNode.Value & "'").Count > 0 Then

                If IsNothing(ViewState("FirstRedNode")) Then
                    ViewState("FirstRedNode") = newNode.ValuePath.ToString()
                End If

                newNode.Text = "<span style='color:red;'>" & newNode.Text & "</span>"
            Else
                newNode.SelectAction = TreeNodeSelectAction.None
                newNode.Text = "<span style='color:black;'>" & newNode.Text & "</span>"
            End If

            addNodeToTv(newNode, dtBranch, newNode.Value)
        Next
    End Sub

    ''' <summary>
    ''' 綁定存在人員變動的單位
    ''' </summary>
    ''' <remarks>[Lake] 2012/05/23 Created</remarks>
    Sub bindEditBranch()
        Dim syBranchList As New AUTH_OP.SY_BRANCHList(GetDatabaseManager())
        Dim syBranch As New AUTH_OP.SY_BRANCH(GetDatabaseManager())
        Dim apCode As New AP_CODE(GetDatabaseManager())
        Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager())
        Dim dtChange As DataTable = ViewState("dtChange")
        Dim iCount As Integer = 0
        Dim iRowCount As Integer = dtChange.Rows.Count - 1
        Dim sBraDepNoList As String = String.Empty

        ' 部門父子關係表
        Dim dtBranch As New DataTable()
        dtBranch.Columns.Add("BRADEPNO", Type.GetType("System.String"))
        dtBranch.Columns.Add("BRCNAME", Type.GetType("System.String"))
        dtBranch.Columns.Add("PARENT", Type.GetType("System.String"))

        For Each dr As DataRow In dtChange.Rows
            sBraDepNoList = sBraDepNoList & ";" & dr("VALUE").ToString()
        Next

        If sBraDepNoList <> "" Then
            sBraDepNoList = sBraDepNoList.Substring(1)
        Else
            Return
        End If

        ' 將變動單位的各層級父子關係保存
        For Each sBraDepNo As String In sBraDepNoList.Split(";")
            dtBranch = getFCNode(dtBranch, sBraDepNo)
        Next

        If apCode.loadByPK(m_sUpCode2366) Then

            ' root節點
            Dim rootNode As TreeNode = New TreeNode()
            Dim node As TreeNode

            rootNode.Value = apCode.getAttribute("VALUE").ToString().Trim()
            rootNode.Text = apCode.getAttribute("TEXT").ToString.Trim()

            rootNode.Expanded = True

            ' 添加根節點到單位樹
            tvBranchCheck.Nodes.Add(rootNode)

            ' 展開預設層次
            tvBranchCheck.ExpandDepth = m_sTreeLevel

            rootNode.SelectAction = TreeNodeSelectAction.None
            rootNode.Text = "<span style='color:black;'>" & rootNode.Text & "</span>"
        End If

        ' 綁定異動單位到單位樹
        addNodeToTv(tvBranchCheck.Nodes(0), dtBranch, "0")

        ' 選擇第一個異動單位
        If ViewState("FirstRedNode").ToString() <> "" Then

            tvBranchCheck.FindNode(ViewState("FirstRedNode").ToString()).Select()

            ' 初始化修改前後人員列表
            initStaffList(tvBranchCheck.SelectedNode)
        End If
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

            If syBranchList.getBrancByParent(sBraDepNo, "", "1") Then
                For Each dr As DataRow In syBranchList.getCurrentDataSet().Tables(0).Rows
                    sAllBraDepNo = sAllBraDepNo & "," & dr("BRA_DEPNO").ToString().Split(";")(0)

                    sAllBraDepNo = getSubBraDepNo(dr("BRA_DEPNO").ToString().Split(";")(0), sAllBraDepNo)
                Next
            End If

            Return sAllBraDepNo
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Function
#End Region

#Region "正常模塊Event"
    ''' <summary>
    ''' 正常模塊，
    ''' “全部”或“依單位”選擇狀態變動時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>[Lake] 2012/06/25 Created</remarks>
    Protected Sub rdolNormalCondition_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles rdolNormalCondition.SelectedIndexChanged
        Try
            Dim dtListBoxMove As DataTable = ViewState("dtListBoxMove")
            Dim syBranchList As New AUTH_OP.SY_BRANCHList(GetDatabaseManager())

            ' 一級單位的所有子部門集合
            Dim dtSubBranch As New DataTable()

            dtListBoxMove.Rows.Clear()
            ViewState("dtListBoxMove") = dtListBoxMove

            ' 選擇“全部”
            If rdolNormalCondition.SelectedValue = "all" Then

                ' 設置單位下拉選單狀態
                ddlBranch.SelectedIndex = 0
                ddlBranch.Enabled = False

                ' 綁定一級單位到單位樹
                bindFirstLevelBranchToTv()
            ElseIf rdolNormalCondition.SelectedValue = "branch" Then

                ' 設置單位下拉選單狀態
                ddlBranch.Enabled = True

                ' 清空安泰銀行下屬單位
                tvBranch.Nodes(0).ChildNodes.Clear()

                ' 清空安泰銀行下屬單位
                tvBranch.Nodes(0).Select()

                tvBranch.SelectedNodeStyle.Reset()
            End If

            ' 安泰銀行有下屬單位
            If tvBranch.Nodes(0).ChildNodes.Count > 0 Then

                ' 第一個單位處於選擇狀態
                tvBranch.Nodes(0).ChildNodes(0).Select()

                ' 選擇一級單位且沒有下屬單位
                If tvBranch.SelectedNode.Parent.Value.Split(";")(0) = 0 AndAlso tvBranch.SelectedNode.ChildNodes.Count = 0 Then

                    If syBranchList.getAllSubBraDepNo(tvBranch.SelectedNode.Value.Split(";")(1), tvBranch.SelectedNode.Value.Split(";")(0)) Then
                        dtSubBranch = syBranchList.getCurrentDataSet().Tables(0)
                    End If

                    If dtSubBranch.Rows.Count > 0 Then

                        ' 顯示當前選擇節點的子節點
                        addNode(tvBranch.SelectedNode, dtSubBranch)

                        tvBranch.SelectedNode.CollapseAll()
                    End If
                End If

                tvBranch.SelectedNodeStyle.ForeColor = Drawing.Color.Black
            End If

            ' 加載人員列表
            initNormalStaff()

            ' 記錄選擇單位
            hidSelectedNode.Value = tvBranch.SelectedNode.Value.Split(";")(0)
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 下拉選單選擇一個單位
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>[Lake] 2012/05/18 Created</remarks>
    Protected Sub ddlBranch_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles ddlBranch.SelectedIndexChanged
        Try

            ' 異動過的數據集合
            Dim dtListBoxMove As New DataTable

            dtListBoxMove.Columns.Add("VALUE", Type.GetType("System.String"))
            dtListBoxMove.Columns.Add("TEXT", Type.GetType("System.String"))
            dtListBoxMove.Columns.Add("OPERATION", Type.GetType("System.String"))

            ' 重新設置ViewState
            ViewState("dtListBoxMove") = dtListBoxMove

            ' 清空左右人員列表
            lstLeft.Items.Clear()
            lstRight.Items.Clear()

            ' 添加下屬單位
            bindBranchByParentToTv()

            ' 選擇”請選擇“
            If ddlBranch.SelectedIndex = "0" Then
                tvBranch.Nodes(0).Select()

                tvBranch.SelectedNodeStyle.Reset()
            Else
                tvBranch.Nodes(0).ChildNodes(0).Select()

                ' 按照設定值展開單位樹層級
                tvBranch.Nodes(0).ChildNodes(0).Expand()

                tvBranch.SelectedNodeStyle.ForeColor = Drawing.Color.Black
            End If

            ' [正常模塊]人員列表
            initNormalStaff()

            ' 異動數據變紅
            editListItemColor()

            ' 異東數據變紅
            editMoweItemColor()

            hidSelectedNode.Value = ddlBranch.SelectedValue.Split(";")(0)
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' "正常模塊"單位樹選擇一個單位
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>[Lake] 2012/05/18 Created</remarks>
    Protected Sub tvBranch_SelectedNodeChanged(ByVal sender As Object, ByVal e As EventArgs) Handles tvBranch.SelectedNodeChanged
        Try
            Dim dtFlow As DataTable = ViewState("dataFlow")
            dtFlow.Rows.Clear()
            ViewState("dataFlow") = dtFlow

            ' 選擇一個節點
            selectNodeOrCancel()

            hidSelectedNode.Value = tvBranch.SelectedNode.Value.Split(";")(0)
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 點擊"取消"按鈕
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnCancel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnCancel.Click
        Try
            ' 取消操作
            selectNodeOrCancel()
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 選入
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>[Lake] 2012/05/18 Created</remarks>
    Protected Sub btnCheckIn_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnCheckIn.Click

        Dim listBridStaffid As New List(Of FLOW_OP.BRID_STAFFID)
        Dim drcBridStaffid As DataRowCollection
        Dim syTable As FLOW_OP.TABLE.SY_REL_BRANCH_USER

        Try

            syTable = New FLOW_OP.TABLE.SY_REL_BRANCH_USER(GetDatabaseManager)

            ' 定義數據表接收移動的數據
            Dim dtListBoxMove As DataTable = ViewState("dtListBoxMove")

            ' 定義數據表接收向右移動的數據
            Dim dtToRight As New DataTable
            dtToRight.Columns.Add("VALUE", Type.GetType("System.String"))
            dtToRight.Columns.Add("TEXT", Type.GetType("System.String"))

            Dim iCount As Integer = lstLeft.Items.Count
            Dim iIndex As Integer = 0
            Dim sMove As String = String.Empty
            Dim sTRight As String = String.Empty

            ' 循環得到被選中的人員
            For i As Integer = 0 To iCount - 1

                ' 獲得當前人員
                Dim item As ListItem = lstLeft.Items(iIndex)

                ' 如果當前人員被選中
                If item.Selected Then

                    If item.Attributes.CssStyle.Value <> Nothing AndAlso item.Attributes.CssStyle.Value = "color:red" Then
                        item.Attributes.CssStyle.Value = Nothing
                    End If

                    lstLeft.Items.Remove(item)
                    lstRight.Items.Add(item)

                    Dim row As DataRow = dtToRight.NewRow
                    row("VALUE") = item.Value
                    row("TEXT") = item.Text

                    '檢查是否已存在其它分行，如果有，顯示警告
                    drcBridStaffid = syTable.GetDataRowCollection(
                                            "select distinct BU.STAFFID, US.USERNAME, BR.BRID, BR.BRCNAME " & vbCrLf & _
                                            "  from SY_REL_BRANCH_USER BU " & vbCrLf & _
                                            " inner join SY_USER US " & vbCrLf & _
                                            "    on US.STAFFID = BU.STAFFID " & vbCrLf & _
                                            " inner join SY_BRANCH BR " & vbCrLf & _
                                            "    on BU.BRA_DEPNO = BR.BRA_DEPNO " & vbCrLf & _
                                            " where BU.STAFFID = @STAFFID@ " & vbCrLf & _
                                            "   and BR.BRID <> @BRID@ " & vbCrLf,
                                            "STAFFID", item.Value,
                                            "BRID", tvBranch.SelectedNode.Value.Split(";")(1))

                    If IsNothing(drcBridStaffid) = False Then
                        For Each drBridStaffid As DataRow In drcBridStaffid
                            listBridStaffid.Add(FLOW_OP.BRID_STAFFID.Pair(
                                                drBridStaffid("BRID"), drBridStaffid("BRCNAME"),
                                                drBridStaffid("STAFFID"), drBridStaffid("USERNAME")))
                        Next
                    End If

                    dtToRight.Rows.Add(row)

                    iIndex = iIndex - 1
                End If

                iIndex = iIndex + 1
            Next

            If listBridStaffid.Count > 0 Then
                Dim sb As New StringBuilder("下列行員已存在於其它分行：")

                For Each item As FLOW_OP.BRID_STAFFID In listBridStaffid
                    sb.Append(vbCrLf & item.Staffname & "(" & item.Staffid & ")  已存在於 " & item.Brcname & "(" & item.Brid & ")")
                Next

                com.Azion.EloanUtility.UIUtility.alert(sb.ToString)
                'SYUIBase.showErrMsg(Me, New Exception(sb.ToString))
            End If

            ' 取消右邊ListBox的選中狀態
            lstRight.SelectedIndex = -1

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

            ' 異動數據變紅
            editListItemColor()

            ' 移動數據變紅
            editMoweItemColor()

            ViewState("dtListBoxMove") = dtListBoxMove
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 選出
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>[Lake] 2012/05/18 Created</remarks>
    Protected Sub btnCheckOut_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnCheckOut.Click
        Try
            ' 聲明參數
            Dim dtListBoxMove As DataTable = ViewState("dtListBoxMove")  ' 左右移動的數據
            Dim dtFlow As DataTable = ViewState("dataFlow")   ' 已選擇的資料
            Dim iCount As Integer = lstRight.Items.Count   ' 已選擇數據總數
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
                Dim lstitem As ListItem = lstRight.Items(iIndex)

                ' 如果當前數據被選中
                If lstitem.Selected = True Then

                    If lstitem.Attributes.CssStyle.Value <> Nothing AndAlso lstitem.Attributes.CssStyle.Value = "color:red" Then
                        lstitem.Attributes.CssStyle.Value = Nothing
                    End If

                    ' 不能選出自己
                    If lstitem.Text.Substring(0, 7) = m_sWorkingUserid Then
                        com.Azion.EloanUtility.UIUtility.alert("登入者不可選出自己！")

                        iIndex = iIndex + 1

                        Continue For
                    End If

                    lstRight.Items.Remove(lstitem)
                    lstLeft.Items.Add(lstitem)

                    Dim row As DataRow = dtToLeft.NewRow
                    row("VALUE") = lstitem.Value
                    row("TEXT") = lstitem.Text
                    dtToLeft.Rows.Add(row)

                    iIndex = iIndex - 1
                End If

                iIndex = iIndex + 1
            Next

            ' 取消左邊ListBox的選中狀態
            lstLeft.SelectedIndex = -1

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
                            If dtFlow.Rows(j)("STAFFID") = dtToLeft.Rows(i)("VALUE") Then
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

            ' 異動數據變紅
            editListItemColor()

            ' 移動數據變紅
            editMoweItemColor()

            ViewState("dtListBoxMove") = dtListBoxMove
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 儲存
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>[Lake] 2012/05/18 Created</remarks>
    Protected Sub btnSave_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnSave.Click
        Try
            Dim syRelRoleUserList As New AUTH_OP.SY_REL_ROLE_USERList(GetDatabaseManager())
            Dim syUserRoleAgentList As New AUTH_OP.SY_USERROLEAGENTList(GetDatabaseManager())

            Dim sAllBraDepNo As String = String.Empty

            ' 本次刪除人員
            Dim sToDeleteStaff As String = String.Empty

            ' 本次操作人員
            Dim dtListBoxMove As DataTable = ViewState("dtListBoxMove")

            'If Not (dtListBoxMove.Rows.Count > 0) Then
            '    Return
            'End If

            ' 取得本次刪除人員
            For Each dr As DataRow In dtListBoxMove.Rows
                If dr("OPERATION").ToString() = "D" Then
                    sToDeleteStaff = sToDeleteStaff & "','" & dr("TEXT").ToString().Substring(0, 7)
                End If
            Next

            ' 本次刪除人員
            If sToDeleteStaff <> "" Then
                sToDeleteStaff = sToDeleteStaff.Substring(3)
            End If

            ' 判斷是否分配角色
            If syRelRoleUserList.getByToDeleteStaffId(tvBranch.SelectedNode.Value.Split(";")(0), sToDeleteStaff) Then

                ' 是否分配角色標識
                hidRoleFlag.Value = "1"
            End If

            ' 存在代理關係
            If syUserRoleAgentList.getToDeleteStaff(tvBranch.SelectedNode.Value.Split(";")(0), sToDeleteStaff) Then

                ' 是否存在代理關係
                hidAgentFlag.Value = "1"
            End If

            ' 給隱藏按鈕註冊儲存異動人員的方法
            btnSaveStaff.OnClientClick = "return SaveChangeToXML();"

            ' 調用儲存異動人員的方法
            ClientScript.RegisterStartupScript(Me.GetType(), Guid.NewGuid().ToString(), "CallStaffSave();", True)
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 保存人員
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>[Lake] 2012/06/18 Crated</remarks>
    Protected Sub btnSaveStaff_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnSaveStaff.Click
        Try
            ' 實例化
            Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager())
            Dim xmlDocument As New XmlDocument
            Dim sType As String = String.Empty
            Dim item As XmlNode  ' 臨時檔節點資料
            Dim newXmlData As String = String.Empty
            Dim nodeList As IList(Of XmlNode) = New List(Of XmlNode)()
            Dim syTempInfoEdit As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager())

            ' 獲得絕對異動的數據
            Dim dtListBoxMove As DataTable = ViewState("dtListBoxMove")

            Dim sbXmlData As New StringBuilder
            Dim oldXmlData As String = String.Empty

            If ViewState("inOrOutXml") <> "" Then

                ' 若案件已在流程中
                If m_sFourStepNo = "" Then

                    ' 存儲修改過的Xml資料
                    If Not syTempInfoEdit.loadByPK(m_sWorkingUserid, m_sWorkingTopDepNo, m_sFuncCode) Then

                        ' 新增異動數據到臨時表
                        syTempInfoEdit.setAttribute("STAFFID", m_sWorkingUserid)
                        syTempInfoEdit.setAttribute("BRA_DEPNO", m_sWorkingTopDepNo)
                        syTempInfoEdit.setAttribute("FUNCCODE", m_sFuncCode)
                    End If

                    syTempInfoEdit.setAttribute("TEMPDATA", ViewState("inOrOutXml").ToString())

                    syTempInfoEdit.save()
                    syTempInfoEdit.clear()
                Else

                    ' 存儲修改過的Xml資料
                    If Not syTempInfoEdit.loadByCaseId(m_sCaseId) Then
                        syTempInfoEdit.setAttribute("CASEID", m_sCaseId)
                    End If

                    syTempInfoEdit.setAttribute("TEMPDATA", ViewState("inOrOutXml").ToString())

                    syTempInfoEdit.save()
                    syTempInfoEdit.clear()
                End If

                ViewState("inOrOutXml") = ""
            End If

            ' 如果有絕對異動的數據
            If dtListBoxMove.Rows.Count > 0 Then

                ' 組XML文件
                sbXmlData.Append("<SY>")
                For i As Integer = 0 To dtListBoxMove.Rows.Count - 1
                    sbXmlData.Append("<SY_REL_BRANCH_USER>")
                    sbXmlData.Append("<STAFFID>" + dtListBoxMove.Rows(i)("VALUE") + "</STAFFID>")
                    sbXmlData.Append("<BRA_DEPNO>" + tvBranch.SelectedNode.Value.Split(";")(0) + "</BRA_DEPNO>")
                    sbXmlData.Append("<OPERATION>" + dtListBoxMove.Rows(i)("OPERATION") + "</OPERATION>")
                    sbXmlData.Append("<NAME>" + dtListBoxMove.Rows(i)("TEXT") + "</NAME>")
                    sbXmlData.Append("</SY_REL_BRANCH_USER>")
                Next
                sbXmlData.Append(getOLstBoxRright())
                sbXmlData.Append("</SY>")

                ' 取得臨時檔資料
                oldXmlData = getXmlData(syTempInfo)

                If oldXmlData <> "" Then

                    ' document對象載入XML文件
                    xmlDocument.LoadXml(oldXmlData)

                    For i As Integer = 0 To xmlDocument.GetElementsByTagName("SY_REL_BRANCH_USER").Count - 1
                        If DirectCast(xmlDocument.GetElementsByTagName("SY_REL_BRANCH_USER")(i)("BRA_DEPNO"), System.Xml.XmlElement).InnerText = tvBranch.SelectedNode.Value.Split(";")(0) Then
                            item = xmlDocument.GetElementsByTagName("SY_REL_BRANCH_USER")(i)
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
                                    oldXML = oldXML & "<SY_REL_BRANCH_USER>" & oldItemXML.ToString() & "</SY_REL_BRANCH_USER>"
                                End If

                                newXML = newXML.Replace(oldItemXML.ToString(), "").Replace("<SY_REL_BRANCH_USER></SY_REL_BRANCH_USER>", "")
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

                ' 如果數據有被退回
                If m_sFourStepNo = "" Then

                    ' 修改臨時檔數據
                    If Not syTempInfo.loadByPK(m_sWorkingUserid, m_sWorkingTopDepNo, m_sFuncCode) Then

                        ' 新增異動數據到臨時表
                        syTempInfo.setAttribute("STAFFID", m_sWorkingUserid)
                        syTempInfo.setAttribute("BRA_DEPNO", m_sWorkingTopDepNo)
                        syTempInfo.setAttribute("FUNCCODE", m_sFuncCode)
                    End If

                    syTempInfo.setAttribute("TEMPDATA", newXmlData.ToString())
                    syTempInfo.save()
                Else

                    ' 修改臨時檔數據
                    If Not syTempInfo.loadByCaseId(m_sCaseId) Then
                        syTempInfo.setAttribute("CASEID", m_sCaseId)
                    End If

                    syTempInfo.setAttribute("TEMPDATA", newXmlData.ToString())
                    syTempInfo.save()
                End If
            End If

            ' 如果數據有被退回
            If m_sFourStepNo = "" Then

                ' 修改臨時檔數據
                If syTempInfo.loadByPK(m_sWorkingUserid, m_sWorkingTopDepNo, m_sFuncCode) Then

                    If syTempInfo.getAttribute("TEMPDATA").ToString() = "<SY></SY>" Then
                        syTempInfo.remove()
                    End If
                End If
            Else
                If syTempInfo.loadByCaseId(m_sCaseId) Then

                    If syTempInfo.getAttribute("TEMPDATA").ToString() = "<SY></SY>" Then
                        syTempInfo.remove()
                    End If
                End If
            End If

            com.Azion.EloanUtility.UIUtility.alert("儲存成功！")

            dtListBoxMove.Rows.Clear()

            ViewState("dtListBoxMove") = dtListBoxMove

            ' 異動數據變紅
            editListItemColor()

            ' 移動數據變紅
            editMoweItemColor()

            ' 設置是否刪除角色及代理關係標識為空
            hidRoleFlag.Value = ""
            hidAgentFlag.Value = ""
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 確認送出
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>[Lake] 2012/05/18 Created</remarks>
    Protected Sub btnSendFlow_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnSendFlow.Click
        Try
            Dim flowFacade As New FLOW_OP.FlowFacade(GetDatabaseManager())
            Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager())
            Dim stepInfo As New FLOW_OP.StepInfo()
            Dim sTempData As String = String.Empty

            ' 異動人員XML資料轉換后存儲到此變量
            Dim dtChangeBranchStaff As DataTable

            ' m_sFourStepNo=""時，根據“登入者編號”，“部門代碼”，“功能編號”查詢人員變動資料；
            ' 300或0400時, 根據m_sCaseId查詢
            If m_sFourStepNo = "" Then
                If syTempInfo.loadTempData(m_sWorkingUserid, m_sWorkingTopDepNo, m_sFuncCode) Then
                    sTempData = syTempInfo.getAttribute("TEMPDATA")
                End If
            Else
                If syTempInfo.loadByCaseId(m_sCaseId) Then
                    sTempData = syTempInfo.getAttribute("TEMPDATA")
                End If
            End If

            ' 沒有送審資料，提示
            If String.IsNullOrEmpty(sTempData) Then
                com.Azion.EloanUtility.UIUtility.alert("無送審資料，不能進行送出！")

                Return
            Else
                ' 轉換XML資料為Table
                If Not String.IsNullOrEmpty(sTempData) Then
                    dtChangeBranchStaff = com.Azion.EloanUtility.XmlUtility.convertXML2DataTable(sTempData)
                End If
            End If

            ' 事務開始
            GetDatabaseManager().beginTran()

            If m_bCheck Then

                ' 走【審核模塊】同意操作，不走流程，不增臨時檔
                AgreeOperate(dtChangeBranchStaff, "Y")

                ' 初始化數據
                ddlBranch.SelectedIndex = -1
                lstLeft.Items.Clear()
                lstRight.Items.Clear()
                tvBranch.Nodes(0).ChildNodes.Clear()
            Else

                ' 流程送出，更新Temp表CaseId
                If m_sStepNo = "" Then
                    stepInfo = MBSC.UICtl.UIShareFun.startFlow(GetDatabaseManager(), m_sFlowName, True)

                    ' 更新人員變動臨時檔的CASEID
                    syTempInfo.setAttribute("CASEID", stepInfo.currentStepInfo.caseId)
                    syTempInfo.save()
                Else
                    stepInfo = MBSC.UICtl.UIShareFun.pushFlow(GetDatabaseManager(), m_sCaseId, , , , False)
                End If

                ' 循環異動資料，新增異動部門人員
                For Each dr As DataRow In dtChangeBranchStaff.Rows
                    ' 新增歷史檔
                    insertSyRelBranchUserHis(dr("BRA_DEPNO").ToString(), dr("STAFFID").ToString(), dr("OPERATION").ToString(), "", _
                           stepInfo.currentStepInfo.caseId, stepInfo.currentStepInfo.stepNo, _
                           stepInfo.currentStepInfo.subflowSeq, stepInfo.currentStepInfo.subflowCount)
                Next

                ' 跳轉到代辦清單
                'com.Azion.EloanUtility.UIUtility.Redirect("SY_CASELIST.aspx")
            End If

            ' 事務提交
            GetDatabaseManager().commit()
            com.Azion.EloanUtility.UIUtility.showJSMsg("案件編號：" + stepInfo.currentStepInfo.caseId + " 送出成功！")
            CloseWindow()
        Catch ex As Exception

            ' 事務回滾
            GetDatabaseManager().Rollback()

            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 全部取消
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>[Lake] 2012/07/20 Created</remarks>
    Protected Sub btnCancelAll_Click(sender As Object, e As EventArgs) Handles btnCancelAll.Click
        Try
            Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager())
            Dim dtListBoxMove As DataTable = ViewState("dtListBoxMove")
            Dim dtFlow As DataTable = ViewState("dataFlow")

            lstLeft.Items.Clear()
            lstRight.Items.Clear()


            dtFlow.Rows.Clear()
            ViewState("dataFlow") = dtFlow

            ' 清空存儲絕對異動的數據
            dtListBoxMove.Rows.Clear()
            ViewState("dtListBoxMove") = dtListBoxMove

            ViewState("inOrOutXml") = ""

            ' 選擇“全部”
            If rdolNormalCondition.SelectedValue = "all" Then

                If tvBranch.Nodes.Count > 0 Then
                    If tvBranch.Nodes(0).ChildNodes.Count > 0 Then
                        tvBranch.Nodes(0).ChildNodes(0).Select()

                        initNormalStaff()
                    End If
                End If
            Else
                If m_sHoFlag <> "1" Then
                    ddlBranch.SelectedIndex = 1

                    If tvBranch.Nodes.Count > 0 Then
                        If tvBranch.Nodes(0).ChildNodes.Count > 0 Then
                            tvBranch.Nodes(0).ChildNodes(0).Select()

                            initNormalStaff()
                        End If
                    End If
                Else
                    ddlBranch.SelectedIndex = 0

                    tvBranch.Nodes(0).ChildNodes.Clear()

                    tvBranch.Nodes(0).Select()
                End If
            End If

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
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub
#End Region

#Region "審核模塊Event"
    ''' <summary>
    ''' "審核模塊"單位樹
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>[Lake] 2012/05/24</remarks>
    Protected Sub tvBranchCheck_SelectedNodeChanged(ByVal sender As Object, ByVal e As EventArgs) Handles tvBranchCheck.SelectedNodeChanged
        Try
            ' 清空修改前後人員列表
            lstEditBefore.Items.Clear()
            lstEditAfter.Items.Clear()

            ' 初始化左右人員列表
            initStaffList(tvBranchCheck.SelectedNode)

            ' 當前選擇單位變黑
            tvBranchCheck.SelectedNodeStyle.ForeColor = Drawing.Color.Black
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 同意
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>[Lake] 2012/06/26 Created</remarks>
    Protected Sub btnYes_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnYes.Click
        Try
            Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager())

            ' TempInfo中XML人員變動資料
            Dim sTempData As String = String.Empty

            ' 異動人員XML資料轉換后存儲到此變量
            Dim dtChangeBranchStaff As DataTable

            ' 事務開始
            GetDatabaseManager().beginTran()

            ' 查詢Temp資料,取得人員變動資料
            If syTempInfo.loadByCaseId(m_sCaseId) Then
                sTempData = syTempInfo.getAttribute("TEMPDATA")

                ' 轉換XML資料為Table
                If Not String.IsNullOrEmpty(sTempData) Then
                    dtChangeBranchStaff = com.Azion.EloanUtility.XmlUtility.convertXML2DataTable(sTempData)
                End If
            End If

            ' 同意操作
            AgreeOperate(dtChangeBranchStaff, "Y")

            ' 事務提交
            GetDatabaseManager().commit()

            ' 關閉窗體
            com.Azion.EloanUtility.UIUtility.showJSMsg("案件編號：" + m_sCaseId + " 送出成功！")
            CloseWindow()
        Catch ex As Exception

            ' 事務回滾
            GetDatabaseManager().Rollback()

            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 不同意
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>[Lake] 2012/05/22</remarks>
    Protected Sub btnNo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnNo.Click
        Try
            Dim flowFacade As New FLOW_OP.FlowFacade(GetDatabaseManager())
            Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager())
            Dim sTempData As String = String.Empty
            Dim stepInfo As New FLOW_OP.StepInfo()

            ' 異動人員XML資料轉換后存儲到此變量
            Dim dtChangeBranchStaff As DataTable

            ' 事務開始 
            GetDatabaseManager().beginTran()

            ' 調用流程方法
            stepInfo = MBSC.UICtl.UIShareFun.pushFlow(GetDatabaseManager(), m_sCaseId, , , , False)

            ' 查詢Temp資料,取得人員變動資料
            If syTempInfo.loadByCaseId(m_sCaseId) Then
                sTempData = syTempInfo.getAttribute("TEMPDATA")

                ' 轉換XML資料為Table
                If Not String.IsNullOrEmpty(sTempData) Then
                    dtChangeBranchStaff = com.Azion.EloanUtility.XmlUtility.convertXML2DataTable(sTempData)
                End If
            End If

            ' 循環異動資料，新增異動部門人員
            If IsNothing(dtChangeBranchStaff) = False Then
                For Each dr As DataRow In dtChangeBranchStaff.Rows

                    ' 新增歷史檔
                    insertSyRelBranchUserHis(dr("BRA_DEPNO").ToString(), dr("STAFFID").ToString(), dr("OPERATION").ToString(), "N", _
                                             stepInfo.currentStepInfo.caseId, stepInfo.currentStepInfo.stepNo, _
                                             stepInfo.currentStepInfo.subflowSeq, stepInfo.currentStepInfo.subflowCount)
                Next
            End If

            ' 刪除Temp資料
            syTempInfo.deleteByCaseID(m_sCaseId)
            syTempInfo.clear()

            ' 事務提交
            GetDatabaseManager().commit()


            ' 關閉窗體
            com.Azion.EloanUtility.UIUtility.showJSMsg("案件編號：" + m_sCaseId + " 送出成功！")
            CloseWindow()
        Catch ex As Exception

            ' 事務回滾
            GetDatabaseManager().Rollback()

            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 修正補充
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>[Lake] 2012/05/22</remarks>
    Protected Sub btnReviseFlow_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnReviseFlow.Click
        Try
            Dim flowFacade As New FLOW_OP.FlowFacade(GetDatabaseManager())
            Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager())
            Dim stepInfo As New FLOW_OP.StepInfo()

            ' 人員變動臨時資料
            Dim sTempData As String = String.Empty

            ' 異動人員XML資料轉換后存儲到此變量
            Dim dtChangeBranchStaff As DataTable

            ' 事務開始
            GetDatabaseManager().beginTran()

            ' 查詢Temp資料,取得人員變動資料
            If syTempInfo.loadByCaseId(m_sCaseId) Then
                sTempData = syTempInfo.getAttribute("TEMPDATA")

                ' 轉換XML資料為Table
                If Not String.IsNullOrEmpty(sTempData) Then
                    dtChangeBranchStaff = com.Azion.EloanUtility.XmlUtility.convertXML2DataTable(sTempData)
                End If
            End If

            ' 流程方法
            stepInfo = MBSC.UICtl.UIShareFun.rollBack(GetDatabaseManager(), m_sCaseId, , , , , , False)

            ' 循環異動資料，新增異動部門人員
            For Each dr As DataRow In dtChangeBranchStaff.Rows

                ' 新增歷史檔
                insertSyRelBranchUserHis(dr("BRA_DEPNO").ToString(), dr("STAFFID").ToString(), dr("OPERATION").ToString(), "", _
                       stepInfo.currentStepInfo.caseId, stepInfo.currentStepInfo.stepNo, _
                       stepInfo.currentStepInfo.subflowSeq, stepInfo.currentStepInfo.subflowCount)
            Next

            ' 事務提交
            GetDatabaseManager().commit()
            com.Azion.EloanUtility.UIUtility.showJSMsg("案件編號：" + m_sCaseId + " 送出成功！")
            CloseWindow()
        Catch ex As Exception

            ' 事務回滾
            GetDatabaseManager().Rollback()
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub
#End Region
End Class