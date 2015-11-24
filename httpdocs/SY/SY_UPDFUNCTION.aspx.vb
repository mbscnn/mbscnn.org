''' <summary>
''' 程式說明：交易建立
''' 建立者：Avril
''' 建立日期：2012-04-24
''' </summary>
Imports com.Azion.NET.VB
Imports AUTH_OP
Imports AUTH_OP.TABLE
Imports MBSC.MB_OP
Imports System.Xml
Imports System.IO
Imports com.Azion.EloanUtility
Imports MBSC.UICtl
Imports System.Web.UI

Public Class SY_UPDFUNCTION
    Inherits SYUIBase

    Dim m_sUpCode2366 As String = String.Empty
    Dim m_sUpCode2370 As String = String.Empty

    Public dtDataTable As New DataTable
    Public m_sFlag As String = String.Empty

    Dim m_sItem As XmlNode
    Dim m_sXmlData As String = String.Empty
    Dim m_sFuncNameList As String = String.Empty
    Dim m_sFuncCodeList As String = String.Empty

    Dim m_sTreeNode As TreeNode

    ' 用於替代名稱中可能出現的“/”
    Public COL_DELIM As String = Microsoft.VisualBasic.ChrW(1)

#Region "Page Load"

    ''' <summary>
    ''' 頁面載入
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        ' 調用初始化參數方法
        initParas()

        ' 頁面第一次加載
        If Not IsPostBack Then

            bindData()
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
        Dim syBranch As New AUTH_OP.SY_BRANCH(GetDatabaseManager())
       
        ' 測試模式
        'If m_bTesting Then
        '    m_sWorkingUserid = "S000035"
        '    m_sWorkingTopDepNo = "2"
        '    m_sLoginUserid = "S000035"
        '    m_sLoginTopDepNo = "1"
        '    m_sFunccode = "7"
        'End If

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
            If syBranch.loadByCaseId(m_sCaseId) Then
                lblBranch.Text = syBranch.getAttribute("BRCNAME")
            End If
        End If

        ' 取得CodeList中設定的參數值
        m_sUpCode2366 = com.Azion.EloanUtility.CodeList.getAppSettings("UpCode2366")
        m_sUpCode2370 = com.Azion.EloanUtility.CodeList.getAppSettings("UpCode2370")
      
    End Function

    ''' <summary>
    ''' 綁定資料
    ''' </summary>
    ''' <remarks></remarks>
    Sub bindData()
        Try
             
            If m_sFourStepNo = "" Or m_sFourStepNo = "0300" Then

                ' 綁定屬性
                bindHoFlag()

                ' 添加節點【存在于Live表】
                initFunctionTreeView(treeView)

                ' 添加節點【存在于Temp表】
                initTempInfoTreeView(treeView)

                ' 控件顯示
                displayControl()

                treeView.ExpandDepth = m_sTreeLevel

                ' 選擇“安泰銀行”
                treeView.Nodes(0).Select()
                divCheck.Visible = False
                divNormal.Visible = True
            Else
                ' 綁定屬性
                bindCheck()

                ' 審核模塊資料加載
                initCheckData()

                divBranch.Visible = True
                divCheck.Visible = True
                divNormal.Visible = False
            End If
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 根據不同的標識操作XML
    ''' </summary>
    ''' <param name="sFlag"></param>
    ''' <remarks></remarks>
    Function operationXML(ByVal sFlag As String, Optional sApproved As String = "", Optional trNode As TreeNode = Nothing) As String
        Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager())
        Dim syFunctionList As New AUTH_OP.SY_FUNCTION_CODEList(GetDatabaseManager())
        Dim syFlowStep As New AUTH_OP.SY_FLOWSTEP(GetDatabaseManager())
        Dim syFunctionCode As New AUTH_OP.SY_FUNCTION_CODE(GetDatabaseManager())
        Dim xmlDocument As New XmlDocument
        Dim elementList As XmlNodeList
        Dim sStepNo As String = String.Empty
        Dim sSubFlowSeq As String = String.Empty
        Dim sSubFlowCount As String = String.Empty
        Dim iCount As Integer = 0
        Dim sCaseId As String = String.Empty
        Dim sParentNow As String = String.Empty

        If sFlag = "9" Then

            ' 將查詢出資料顯示到修改前區塊中
            If syFunctionCode.loadByPK(trNode.Value.ToString) Then

                lblIDBefore.Text = trNode.Value
                lblNameBefore.Text = syFunctionCode.getAttribute("FUCNAME")
                txtFuncUrlBefore.Text = syFunctionCode.getAttribute("FUCURL")
                txtSortCtrlBefore.Text = syFunctionCode.getAttribute("SORTCTRL")
                If Not String.IsNullOrEmpty(txtFuncUrlBefore.Text) Then
                    rdoListHoFlagBefore.Visible = True
                    rdoListHoFlagBefore.SelectedValue = syFunctionCode.getAttribute("HOFLAG")
                Else
                    rdoListHoFlagBefore.Visible = False
                End If

                rdoDisableBefore.SelectedValue = syFunctionCode.getAttribute("DISABLED")

                If Not trNode.Parent Is Nothing Then
                    lblParentNameBefore.Text = trNode.Parent.Text.Replace("<font color='red'>", "").Replace("</font>", "")
                Else
                    If syFunctionCode.loadByPK(syFunctionCode.getAttribute("PARENT")) Then
                        lblParentNameBefore.Text = syFunctionCode.getAttribute("FUCNAME")
                    End If
                End If
            End If
        End If

        ' 查詢XML的資料 如果是正常頁面 則根據主鍵進行查詢，否則根據案號進行查詢
        If m_sFourStepNo = "" Then
            If syTempInfo.loadTempData(m_sWorkingUserid, m_sWorkingTopDepNo, m_sFuncCode) Then
                m_sXmlData = syTempInfo.getAttribute("TEMPDATA")
                sCaseId = syTempInfo.getAttribute("CASEID")
            End If
        Else
            If syTempInfo.loadByCaseId(m_sCaseId) Then
                m_sXmlData = syTempInfo.getAttribute("TEMPDATA")
                sCaseId = m_sCaseId
            End If
        End If

        ' 根據案號查詢資料
        If syFlowStep.loadByCaseId(sCaseId) Then

            sStepNo = syFlowStep.getAttribute("STEP_NO")
            sSubFlowCount = syFlowStep.getAttribute("SUBFLOW_COUNT")
            sSubFlowSeq = syFlowStep.getAttribute("SUBFLOW_SEQ")
        End If

        If Not m_sXmlData Is Nothing Then
            If m_sXmlData.Length > 0 Then

                ' document對象載入XML文件
                xmlDocument.LoadXml(m_sXmlData)
            End If
        End If

        '  循環取得sy_tempinfo中的資料
        For i As Integer = 0 To xmlDocument.GetElementsByTagName("SY_FUNCTION_CODE").Count - 1
            Dim sFuncCode As String = DirectCast(xmlDocument.GetElementsByTagName("SY_FUNCTION_CODE")(i)("FUNCCODE"), System.Xml.XmlElement).InnerText
            Dim sParent As String = DirectCast(xmlDocument.GetElementsByTagName("SY_FUNCTION_CODE")(i)("PARENT"), System.Xml.XmlElement).InnerText
            Dim sFuncName As String = DirectCast(xmlDocument.GetElementsByTagName("SY_FUNCTION_CODE")(i)("FUCNAME"), System.Xml.XmlElement).InnerText
            Dim sFuncUrl As String = DirectCast(xmlDocument.GetElementsByTagName("SY_FUNCTION_CODE")(i)("FUCURL"), System.Xml.XmlElement).InnerText
            Dim sHoFlag As String = DirectCast(xmlDocument.GetElementsByTagName("SY_FUNCTION_CODE")(i)("HOFLAG"), System.Xml.XmlElement).InnerText
            Dim sDisabled As String = DirectCast(xmlDocument.GetElementsByTagName("SY_FUNCTION_CODE")(i)("DISABLED"), System.Xml.XmlElement).InnerText
            Dim sSortCtrl As String = DirectCast(xmlDocument.GetElementsByTagName("SY_FUNCTION_CODE")(i)("SORTCTRL"), System.Xml.XmlElement).InnerText

            '------------------------------ 正常區域加載資料,查詢xml中的資料，將其新增到左邊目錄樹---------------------------------------------------
            If sFlag = "1" Then
                If sFuncUrl <> "" Then
                    trHoFlag.Visible = True
                    rdoListHoFlag.SelectedValue = sHoFlag
                End If

                If Not syFunctionList.loadDataByFunccode(sFuncCode) Then

                    ' 臨時檔中組織起來的節點
                    Dim subNode As TreeNode = New TreeNode
                    subNode.Text = String.Format("<font color='red'>{0}</font>", sFuncName)
                    subNode.Value = sFuncCode

                    For Each node As TreeNode In treeView.Nodes

                        ' 如果節點=臨時檔中的節點 則給其添加子節點
                        If node.Value = sParent Then
                            node.ChildNodes.Add(subNode)
                        Else
                            ' 驗證節點是否存在,如果沒有 將其新增到TreeView中
                            checkExistNode(node, subNode, sParent)
                        End If
                    Next
                Else
                    For Each node As TreeNode In treeView.Nodes
                        If node.Value = sFuncCode Then
                            node.Text = String.Format("<font color='red'>{0}</font>", sFuncName)
                        Else
                            ' 判斷DB中是否有此值 若有 將其設置為紅色
                            checkNode(node, sFuncCode)
                        End If
                    Next
                End If
            ElseIf sFlag = "2" Then

                '------------------------------ 新增歷史檔部份---------------------------------------------------

                Dim syFunctionCodeHis As New SY_FUNCTION_CODE_HIS(GetDatabaseManager())
                If Not syFunctionCodeHis.loadByPK(sFuncCode, sCaseId, sStepNo, sSubFlowSeq, sSubFlowCount) Then
                    syFunctionCodeHis.setAttribute("FUNCCODE", sFuncCode)
                    syFunctionCodeHis.setAttribute("CASEID", sCaseId)
                    syFunctionCodeHis.setAttribute("STEP_NO", sStepNo)
                    syFunctionCodeHis.setAttribute("SUBFLOW_SEQ", sSubFlowSeq)
                    syFunctionCodeHis.setAttribute("SUBFLOW_COUNT", sSubFlowCount)
                End If

                syFunctionCodeHis.setAttribute("FUCNAME", sFuncName)
                syFunctionCodeHis.setAttribute("FUCURL", sFuncUrl)
                syFunctionCodeHis.setAttribute("PARENT", sParent)
                syFunctionCodeHis.setAttribute("DISABLED", sDisabled)
                syFunctionCodeHis.setAttribute("SORTCTRL", sSortCtrl)
                syFunctionCodeHis.setAttribute("HOFLAG", sHoFlag)
                syFunctionCodeHis.setAttribute("XMLDATA", m_sXmlData)

                If Not String.IsNullOrEmpty(sApproved) Then
                    syFunctionCodeHis.setAttribute("APPROVED", sApproved)
                End If

                syFunctionCodeHis.save()
            ElseIf sFlag = "3" Then
                '------------------------------ 正常區域，更新操作，根據Funccode查詢出相關的資料，給維護區域賦值---------------------------------------------------
                If DirectCast(xmlDocument.GetElementsByTagName("SY_FUNCTION_CODE")(i)("FUNCCODE"), System.Xml.XmlElement).InnerText = treeView.SelectedNode.Value.Split(";")(0) Then
                    sFuncCode = DirectCast(xmlDocument.GetElementsByTagName("SY_FUNCTION_CODE")(i)("FUNCCODE"), System.Xml.XmlElement).InnerText
                    sParent = DirectCast(xmlDocument.GetElementsByTagName("SY_FUNCTION_CODE")(i)("PARENT"), System.Xml.XmlElement).InnerText
                    rdoDisable.SelectedValue = DirectCast(xmlDocument.GetElementsByTagName("SY_FUNCTION_CODE")(i)("DISABLED"), System.Xml.XmlElement).InnerText
                    txtFuncUrl.Text = DirectCast(xmlDocument.GetElementsByTagName("SY_FUNCTION_CODE")(i)("FUCURL"), System.Xml.XmlElement).InnerText
                    txtSortCtrl.Text = DirectCast(xmlDocument.GetElementsByTagName("SY_FUNCTION_CODE")(i)("SORTCTRL"), System.Xml.XmlElement).InnerText

                    If Not String.IsNullOrEmpty(DirectCast(xmlDocument.GetElementsByTagName("SY_FUNCTION_CODE")(i)("HOFLAG"), System.Xml.XmlElement).InnerText) Then
                        rdoListHoFlag.SelectedValue = DirectCast(xmlDocument.GetElementsByTagName("SY_FUNCTION_CODE")(i)("HOFLAG"), System.Xml.XmlElement).InnerText
                    End If

                    If txtFuncUrl.Text <> "" Then
                        trHoFlag.Visible = True
                    End If

                    hidParentId.Value = sParent
                End If
            ElseIf sFlag = "4" Then
                '------------------------------ 正常區域，新增操作，取得最大的Funccode---------------------------------------------------
                If i = xmlDocument.GetElementsByTagName("SY_FUNCTION_CODE").Count - 1 Then
                    sFuncCode = DirectCast(xmlDocument.GetElementsByTagName("SY_FUNCTION_CODE")(i)("FUNCCODE"), System.Xml.XmlElement).InnerText - 1
                    sParent = DirectCast(xmlDocument.GetElementsByTagName("SY_FUNCTION_CODE")(i)("FUNCCODE"), System.Xml.XmlElement).InnerText

                    Return sFuncCode & "*" & sParent
                End If
            ElseIf sFlag = "5" Then
                '------------------------------ 正常區域，刪除操作，查詢臨時檔中是否存在符合刪除條件的資料--------------------------------------------------- 
                If DirectCast(xmlDocument.GetElementsByTagName("SY_FUNCTION_CODE")(i)("FUNCCODE"), System.Xml.XmlElement).InnerText = hidId.Value Then
                    m_sItem = xmlDocument.GetElementsByTagName("SY_FUNCTION_CODE")(i)
                    Exit For
                End If
            ElseIf sFlag = "6" Then
                '------------------------------ 正常區域，確認送出操作，查詢臨時檔中已經存在的資料--------------------------------------------------- 
                If Convert.ToInt32(DirectCast(xmlDocument.GetElementsByTagName("SY_FUNCTION_CODE")(i)("FUNCCODE"), System.Xml.XmlElement).InnerText) < 0 Then

                    ' 名稱集合
                    m_sFuncNameList += "'" & DirectCast(xmlDocument.GetElementsByTagName("SY_FUNCTION_CODE")(i)("FUCNAME"), System.Xml.XmlElement).InnerText & "'" & ","
                Else
                    ' 編號集合
                    m_sFuncCodeList += DirectCast(xmlDocument.GetElementsByTagName("SY_FUNCTION_CODE")(i)("FUNCCODE"), System.Xml.XmlElement).InnerText + ","
                End If
            ElseIf sFlag = "7" Then
                '------------------------------ 正常區域，儲存操作--------------------------------------------------- 
                ' 如果XML中包含Funccode則需要將其刪除掉 然後再進行累加
                Dim sTempId As String = String.Empty

                ' 如果點擊了新增按鈕，則funccode為最新的funccode 否則為當前編輯節點的id
                If hidFlag.Value = "Add" Then
                    sTempId = lblID.Text
                Else
                    sTempId = hidId.Value
                End If

                ' 找到臨時檔中已經存在的資料
                If DirectCast(xmlDocument.GetElementsByTagName("SY_FUNCTION_CODE")(i)("FUNCCODE"), System.Xml.XmlElement).InnerText = sTempId Then
                    m_sItem = xmlDocument.GetElementsByTagName("SY_FUNCTION_CODE")(i)
                    Exit For
                End If

                ' 檢核角色名稱不可與live及temp中資料重複不可重複 
                If hidId.Value <> "" Then
                    If Convert.ToInt32(DirectCast(xmlDocument.GetElementsByTagName("SY_FUNCTION_CODE")(i)("FUNCCODE"), System.Xml.XmlElement).InnerText) <> sTempId Then

                        ' 名稱集合
                        m_sFuncNameList &= "'" & DirectCast(xmlDocument.GetElementsByTagName("SY_FUNCTION_CODE")(i)("FUCNAME"), System.Xml.XmlElement).InnerText & "'" & ","
                    End If
                End If

                ' 編號集合
                m_sFuncCodeList &= DirectCast(xmlDocument.GetElementsByTagName("SY_FUNCTION_CODE")(i)("FUNCCODE"), System.Xml.XmlElement).InnerText + ","
            ElseIf sFlag = "8" Then
                '------------------------------ 審核區域，資料加載--------------------------------------------------- 
                hidNodeParent.Value = sParent

                ' 修改
                If Convert.ToInt32(sFuncCode) > 0 Then
                    loadNodeByParent(sParent)

                    If hidNodeParent.Value <> "" Then
                        If syFunctionCode.loadByPK(hidNodeParent.Value) Then
                            If syFunctionCode.getAttribute("PARENT").ToString() = "0" Then
                                Dim parentNode As TreeNode = New TreeNode()
                                parentNode.Value = syFunctionCode.getAttribute("FUNCCODE")
                                parentNode.Text = syFunctionCode.getAttribute("FUCNAME") & "  (" & syFunctionCode.getAttribute("FUNCCODE") & ")"

                                parentNode.SelectAction = TreeNodeSelectAction.None

                                ' (4)添加子節點
                                addNodeCheck(parentNode, hidNodeParent.Value, sFuncCode)

                                treeViewCheck.Nodes(0).ChildNodes.Add(parentNode)
                            End If
                        End If
                    End If
                Else
                    ' 首次新增狀態
                    If Convert.ToInt32(sParent) > 0 Then
                        loadNodeByParent(sParent)

                        If hidNodeParent.Value <> "" Then
                            If syFunctionCode.loadByPK(hidNodeParent.Value) Then
                                If syFunctionCode.getAttribute("PARENT").ToString() = "0" Then
                                    Dim parentNode As TreeNode = New TreeNode()
                                    parentNode.Value = syFunctionCode.getAttribute("ROLEID")
                                    parentNode.Text = syFunctionCode.getAttribute("ROLENAME") & "  (" & syFunctionCode.getAttribute("FUNCCODE") & ")"

                                    parentNode.SelectAction = TreeNodeSelectAction.None

                                    ' (4)添加子節點
                                    addNodeCheck(parentNode, hidNodeParent.Value, sFuncCode)

                                    treeViewCheck.Nodes(0).ChildNodes.Add(parentNode)
                                End If
                            End If
                        End If
                    End If
                End If

                ' 臨時檔中組織起來的節點
                Dim subNode As TreeNode = New TreeNode
                subNode.Text = String.Format("<font color='red'>{0}</font>", sFuncName)
                subNode.Value = sFuncCode

                ' 添加子項
                'addChildNode(subNode, m_sXmlData)

                iCount += 1

                For Each node As TreeNode In treeViewCheck.Nodes
                    If node.Value = sParent Then
                        node.ChildNodes.Add(subNode)
                    Else
                        ' 驗證節點是否存在,如果沒有 將其新增到TreeView中
                        checkExistNodeCheck(node, subNode, sParent)
                    End If
                Next

                If iCount = 1 Then
                    loadCheck(subNode)
                End If
            ElseIf sFlag = "9" Then
                '------------------------------ 審核區域，更新操作，資料加載--------------------------------------------------- 
                Dim nodeValue As Integer = trNode.Value
                If sFuncCode = nodeValue.ToString Then
                    If Not trNode.Parent Is Nothing Then
                        lblParentNameAfter.Text = trNode.Parent.Text.Replace("<font color='red'>", "").Replace("</font>", "")
                    Else
                        If syFunctionCode.loadByPK(sParent) Then
                            lblParentNameAfter.Text = syFunctionCode.getAttribute("FUCNAME")
                        End If
                    End If

                    lblIDAfter.Text = nodeValue
                    lblNameAfter.Text = sFuncName
                    txtFuncUrlAfter.Text = sFuncUrl
                    txtSortCtrlAfter.Text = sSortCtrl

                    If Not String.IsNullOrEmpty(sFuncUrl) Then
                        rdoListHoFlagAfter.Visible = True
                        rdoListHoFlagAfter.SelectedValue = sHoFlag
                    Else
                        rdoListHoFlagAfter.Visible = False
                    End If

                    rdoDisableAfter.SelectedValue = sDisabled
                End If
            ElseIf sFlag = "10" Then

                '------------------------------ 審核區域，同意按鈕操作，資料更新--------------------------------------------------- 
                ' RoleID<0 新增資料到 sy_role SY_FUNCTIONSUITSYS
                If Convert.ToInt32(sFuncCode) < 0 Then
                    Dim syFunctionCodeHis As New AUTH_OP.SY_FUNCTION_CODE_HIS(GetDatabaseManager())
                    Dim sMaxFuncId As String = String.Empty

                    Dim sParentNew As String = String.Empty

                    sMaxFuncId = syFunctionCode.getMaxFuncCode().ToString()

                    ' 如果此節點是其他節點的父節點 ，則需要更新子節點的parent=sMaxRoleId
                    sParentNow &= sFuncCode & "*" & sMaxFuncId & "|"

                    If Not syFunctionCode.loadByPK(sMaxFuncId) Then
                        syFunctionCode.setAttribute("FUNCCODE", sMaxFuncId)
                    End If

                    If String.IsNullOrEmpty(sParentNew) Then
                        For m As Integer = 0 To sParentNow.Split("|").Count - 1

                            If sParentNow.Split("|")(m).ToString.Contains(sParent) Then
                                If sParent <= 0 Then
                                    sParentNew = sParentNow.Split("|")(m).Split("*")(1).ToString()
                                Else
                                    sParentNew = sParent
                                End If
                            Else
                                If String.IsNullOrEmpty(sParentNew) Then
                                    sParentNew = sParent
                                End If
                            End If
                        Next
                    End If

                    syFunctionCode.setAttribute("PARENT", sParentNew)
                    syFunctionCode.setAttribute("FUCNAME", sFuncName)
                    syFunctionCode.setAttribute("FUCURL", sFuncUrl)
                    syFunctionCode.setAttribute("DISABLED", sDisabled)
                    syFunctionCode.setAttribute("SORTCTRL", sSortCtrl)
                    syFunctionCode.setAttribute("HOFLAG", sHoFlag)
                    syFunctionCode.save()

                Else
                    m_bCheck = True     '不再複核

                    ' 更新資料
                    If m_bCheck Then
                        lblIDAfter.Text = sFuncCode
                    End If

                    If Not syFunctionCode.loadByPK(lblIDAfter.Text) Then
                        syFunctionCode.setAttribute("FUNCCODE", lblIDAfter.Text)
                    End If

                    syFunctionCode.setAttribute("FUCNAME", sFuncName)
                    syFunctionCode.setAttribute("FUCURL", sFuncUrl)
                    syFunctionCode.setAttribute("PARENT", sParent)
                    syFunctionCode.setAttribute("DISABLED", sDisabled)
                    syFunctionCode.setAttribute("SORTCTRL", sSortCtrl)
                    syFunctionCode.setAttribute("HOFLAG", sHoFlag)
                    syFunctionCode.save()
                End If
            End If
        Next
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
                childNode.Value = drRow("FUNCCODE").ToString

                If CInt(drRow("FUNCCODE")) < 0 Then
                    childNode.Text = String.Format("<font color='red'>{0}</font>", drRow("FUCNAME").ToString & "  (" & drRow("FUNCCODE").ToString & ")")
                Else
                    childNode.Text = drRow("FUCNAME").ToString & "  (" & drRow("FUNCCODE").ToString & ")"
                    childNode.SelectAction = TreeNodeSelectAction.None
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
                            childNodeTemp.Value = drRowTemp("FUNCCODE").ToString
                            childNodeTemp.Text = String.Format("<font color='red'>{0}</font>", drRowTemp("FUCNAME").ToString & "  (" & drRowTemp("FUNCCODE").ToString & ")")

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
    ''' 綁定屬性
    ''' </summary>
    ''' <remarks></remarks>
    Sub bindHoFlag()
        Dim apCodelist As New AP_CODEList(GetDatabaseManager())

        ' 綁定屬性
        If apCodelist.getCurUnit(m_sUpCode2370) Then
            rdoListHoFlag.DataSource = apCodelist.getCurrentDataSet
            rdoListHoFlag.DataTextField = "TEXT"
            rdoListHoFlag.DataValueField = "VALUE"
            rdoListHoFlag.DataBind()
        End If
    End Sub

    ''' <summary>
    ''' 查詢臨時檔中的資料
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
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
                sCaseid = syTempInfo.getString("CASEID")
            End If
        Else
            If syTempInfo.loadByCaseId(m_sCaseId) Then
                ' 如果是Insert資料，則需要進行追加 如果是修改 則修改
                oldXmlData = syTempInfo.getAttribute("TEMPDATA").ToString()
                sCaseid = syTempInfo.getString("CASEID")
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
                dv.Sort = "FUNCCODE asc"
                Return dv.ToTable
            End If
        End If
        Return Nothing
    End Function

    ''' <summary>
    ''' 新增真實檔及歷史檔
    ''' </summary>
    ''' <param name="databaseManager"></param>
    ''' <param name="stepinfo"></param>
    ''' <remarks></remarks>
    Sub updatTemp2Live(ByVal databaseManager As DatabaseManager, ByRef stepinfo As FLOW_OP.StepInfo, ByVal sAPPROVED As String)
        Dim dt As DataTable = loadSYTempInfo()
        Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(databaseManager)

        Dim syFunctionCode As New AUTH_OP.SY_FUNCTION_CODE(databaseManager)

        Dim iNewFuncId As Integer = syFunctionCode.getMaxFuncCode()

        If sAPPROVED = "Y" Then
            For Each row As DataRow In dt.Rows
                Dim iRoleId As Integer = CInt(row("FUNCCODE"))
                Dim iParentId As Integer = CInt(row("parent"))

                If iRoleId > 0 Then Continue For

                If iRoleId < 0 Then '取號
                    iNewFuncId += 1
                End If

                For Each childRow As DataRow In dt.Select("[parent]='" & iRoleId & "'")
                    childRow.BeginEdit()
                    childRow("parent") = iNewFuncId
                    childRow.EndEdit()
                    dt.AcceptChanges()
                Next

                row.BeginEdit()
                row("FUNCCODE") = iNewFuncId
                row.EndEdit()
                dt.AcceptChanges()
            Next
        End If

        For Each row As DataRow In dt.Rows
            Dim iFuncCode As Integer = CInt(row("FUNCCODE"))
            '整理沒有在Temp裡的角色

            Dim sFuncOperation As String = String.Empty

            '5更新角色SY_ROLE
            syFunctionCode.clear()

            If Not syFunctionCode.loadByPK(CInt(row("FUNCCODE"))) Then
                sFuncOperation = "I"
            Else
                sFuncOperation = "U" '思考一下，會不會有更新的狀況
            End If

            Dim strFucurl As String = String.Empty
            strFucurl = row("FUCURL").ToString.Replace("＆", "&")

            If sAPPROVED = "Y" Then
                syFunctionCode.setInt("FUNCCODE", iFuncCode)
                syFunctionCode.setString("FUCNAME", row("FUCNAME").ToString)
                'syFunctionCode.setString("FUCURL", row("FUCURL").ToString)
                syFunctionCode.setString("FUCURL", strFucurl)
                syFunctionCode.setInt("PARENT", CInt(row("PARENT")))
                syFunctionCode.setString("DISABLED", row("DISABLED").ToString)
                syFunctionCode.setString("SORTCTRL", row("SORTCTRL").ToString)
                syFunctionCode.setString("HOFLAG", row("HOFLAG").ToString)

                syFunctionCode.save()
            End If

            hidParId.Value &= iFuncCode.ToString & ","

            ' 若父節點有更新啟用狀態，則子節點也需要修改
            modifyLiveData(iFuncCode, row("DISABLED").ToString)

            '6.紀錄新增(I)或修改(U) his(SY_ROLE)
            'If stepinfo Is Nothing Then '不走流程
            '    insertSyFuncCodeHis(databaseManager, row, sFuncOperation, "N", sAPPROVED, , , , )
            'Else
            If hidParentFuncList.Value <> "" Then
                hidParentFuncList.Value = hidParentFuncList.Value.Substring(0, hidParentFuncList.Value.Length - 1)

                ' 更新真實檔
                If sAPPROVED = "Y" Then
                    syFunctionCode.updateDisable(hidParentFuncList.Value, row("DISABLED"))
                End If

                ' 新增歷史檔
                For Each Str As String In hidParentFuncList.Value.Split(",")
                    If Str <> "" Then
                        syFunctionCode.loadByPK(Str)
                        insertSyFunctionCodeHisXXX(databaseManager, syFunctionCode, sFuncOperation, sAPPROVED, _
                                   stepinfo.currentStepInfo.caseId, _
                                   stepinfo.currentStepInfo.stepNo, _
                                   stepinfo.currentStepInfo.subflowSeq, _
                                   stepinfo.currentStepInfo.subflowCount)
                    End If
                Next
            End If

            insertSyFuncCodeHis(databaseManager, row, sFuncOperation, "Y", sAPPROVED, _
                                   stepinfo.currentStepInfo.caseId, _
                                   stepinfo.currentStepInfo.stepNo, _
                                   stepinfo.currentStepInfo.subflowSeq, _
                                   stepinfo.currentStepInfo.subflowCount)
        Next

        ' 刪除臨時檔中的資料
        If sAPPROVED = "Y" OrElse sAPPROVED = "N" Then 'm_bCheck OrElse m_sFourStepNo = "0400"
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
    ''' 新增歷史檔
    ''' </summary>
    ''' <param name="databaseManager"></param>
    ''' <param name="row"></param>
    ''' <param name="sOperation"></param>
    ''' <param name="sAPPROVED"></param>
    ''' <param name="sCaseId"></param>
    ''' <param name="sStepNo"></param>
    ''' <param name="iSubFlowSeq"></param>
    ''' <param name="iSubFlowCount"></param>
    ''' <remarks></remarks>
    Sub insertSyFuncCodeHis(ByVal databaseManager As DatabaseManager, ByVal row As DataRow, ByVal sOperation As String, ByRef sFlag As String, Optional sAPPROVED As String = "Y", _
              Optional ByVal sCaseId As String = "SY0000000000000", Optional ByVal sStepNo As String = "00000000", _
              Optional ByVal iSubFlowSeq As Integer = 0, Optional ByVal iSubFlowCount As Integer = 0)

        Dim syFunctionCodeHis As New AUTH_OP.SY_FUNCTION_CODE_HIS(databaseManager)

        syFunctionCodeHis.loadByPK(row("FUNCCODE"), sCaseId, sStepNo, iSubFlowSeq, iSubFlowCount)

        '因為欄位相同
        For Each column As DataColumn In row.Table.Columns
            syFunctionCodeHis.setAttribute(column.ColumnName.ToUpper, row(column.ColumnName))
        Next

        syFunctionCodeHis.setAttribute("CASEID", sCaseId) ' m_sCaseId
        syFunctionCodeHis.setAttribute("STEP_NO", sStepNo)
        syFunctionCodeHis.setAttribute("SUBFLOW_SEQ", iSubFlowSeq)
        syFunctionCodeHis.setAttribute("SUBFLOW_COUNT", iSubFlowCount)
        syFunctionCodeHis.setAttribute("APPROVED", sAPPROVED)
        syFunctionCodeHis.setAttribute("OPERATION", sOperation)

        syFunctionCodeHis.save()
    End Sub

    Sub insertSyFunctionCodeHisXXX(ByVal databaseManager As DatabaseManager, ByVal syFunctionCode As AUTH_OP.SY_FUNCTION_CODE, ByVal sOperation As String, Optional sAPPROVED As String = "Y", _
                    Optional ByVal sCaseId As String = "", Optional ByVal sStepNo As String = "00000000", _
                    Optional ByVal iSubFlowSeq As Integer = 0, Optional ByVal iSubFlowCount As Integer = 0)

        Dim syFunctionCodeHis As New AUTH_OP.SY_FUNCTION_CODE_HIS(databaseManager)

        syFunctionCodeHis.loadByPK(syFunctionCode.getInt("ROLEID"), sCaseId, sStepNo, iSubFlowSeq, iSubFlowCount)

        For i As Integer = 0 To syFunctionCode.getBosAttributes.size - 1
            Dim bosAttribute As BosAttribute = syFunctionCode.getBosAttributes.item(i)
            Dim sFieldName As String = bosAttribute.getColName()
            syFunctionCodeHis.setAttribute(sFieldName, syFunctionCode.getAttribute(sFieldName))
        Next

        syFunctionCodeHis.setAttribute("CASEID", sCaseId)
        syFunctionCodeHis.setAttribute("STEP_NO", sStepNo)
        syFunctionCodeHis.setAttribute("SUBFLOW_SEQ", iSubFlowSeq)
        syFunctionCodeHis.setAttribute("SUBFLOW_COUNT", iSubFlowCount)
        syFunctionCodeHis.setAttribute("APPROVED", sAPPROVED)
        syFunctionCodeHis.setAttribute("OPERATION", sOperation)
        syFunctionCodeHis.save()
    End Sub

#End Region

#Region "正常區塊Function"

    ''' <summary>
    ''' 給TreeView添加相關的節點
    ''' </summary>
    ''' <param name="treeView"></param>
    ''' <remarks></remarks>
    Sub initFunctionTreeView(ByVal treeView As TreeView)
        Try
            ' 實例化
            Dim apCodelist As New AP_CODEList(GetDatabaseManager())
            Dim syFunctionCodeList As New AUTH_OP.SY_FUNCTION_CODEList(GetDatabaseManager())

            ' root節點
            Dim rootNode As TreeNode = New TreeNode()

            ' (1)綁定root節點
            If apCodelist.loadDataByCodeId(m_sUpCode2366) Then
                For Each dr As DataRow In apCodelist.getCurrentDataSet.Tables(0).Rows 'must be a record in ap_code
                    rootNode.Value = dr("VALUE").ToString() & ";0;;0;0;0;0"
                    rootNode.Text = dr("TEXT").ToString
                Next
            End If

            ' (2)查詢交易建立的資料的第一層資料
            If Not syFunctionCodeList.genAllFunctionCodeList Is Nothing Then

                '將syRoleList存入ViewState
                ViewState("SY_FunctionCodeList") = syFunctionCodeList.getCurrentDataSet.Tables(0)

                ' (4)添加子節點
                addNodes(rootNode, 0, ViewState("SY_FunctionCodeList"))
            End If

            '添加到treeView中()
            treeView.Nodes.Add(rootNode)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ''' <summary>
    ''' 將臨時檔中的資料添加到左邊目錄樹
    ''' </summary>
    ''' <param name="treeView"></param>
    ''' <remarks></remarks>
    Sub initTempInfoTreeView(ByVal treeView As TreeView)
        Dim ds As New DataSet
        Dim dt As New DataTable
        Try
            Dim syTempInfo As New SY_TEMPINFO(GetDatabaseManager())
            Dim xmlData As String = String.Empty

            ' (5) 查詢XML的資料 只有為“”時才按照主鍵進行查詢，否則按鈕caseid進行查詢
            If m_sFourStepNo = "" Then
                If syTempInfo.loadTempData(m_sWorkingUserid, m_sWorkingTopDepNo, m_sFuncCode) Then
                    xmlData = syTempInfo.getString("TEMPDATA")

                    If xmlData <> Nothing Then

                        ' (6) document對象載入XML文件
                        Dim xmlDocument As New XmlDocument
                        xmlDocument.LoadXml(xmlData)
                        ds.ReadXml(New XmlNodeReader(xmlDocument))
                        dt = ds.Tables(0)
                    End If
                End If
            Else
                If syTempInfo.loadByCaseId(m_sCaseId) Then
                    xmlData = syTempInfo.getString("TEMPDATA")

                    If xmlData <> Nothing Then

                        ' (6) document對象載入XML文件
                        Dim xmlDocument As New XmlDocument
                        xmlDocument.LoadXml(xmlData)
                        ds.ReadXml(New XmlNodeReader(xmlDocument))
                        dt = ds.Tables(0)
                    End If
                End If
            End If

            'Dim dtView As New DataView
            'dtView = dt.DefaultView
            'dtView.Sort = "FUNCCODE DESC"

            'dt = dtView.Table

            ' 將資料來源存在ViewState中
            ViewState("SY_FunctionCodeList_Temp") = dt

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
            For Each row As DataRow In CType(ViewState("SY_FunctionCodeList_Temp"), DataTable).Rows
                Dim sFuncCode As String = row("FUNCCODE")
                Dim sFuncName As String = row("FUCNAME")
                Dim sFuncUrl As String = row("FUCURL")
                Dim sParent As String = row("PARENT")
                Dim sDisabled As String = row("DISABLED")
                Dim sSortCtrl As String = row("SORTCTRL")
                Dim sHoFlag As String = row("HOFLAG")
                Dim sFlag As Boolean = False
                Dim sTempName As String = String.Empty

                Dim subNode As TreeNode = New TreeNode
                subNode.Text = String.Format("<font color='red'>{0}</font>", sFuncName.Replace(COL_DELIM, "/") & "  (" & sFuncCode & ")")
                sTempName = subNode.Text

                If sHoFlag = "" Then
                    sHoFlag = ""
                End If

                subNode.Value = Server.HtmlEncode(sFuncCode & ";" & sFuncName.Replace("/", COL_DELIM) & ";" & sFuncUrl.Replace("/", "#") & ";" & sParent & ";" & sDisabled & ";" & sSortCtrl & ";" & sHoFlag)

                Dim sParentValuePath As String = String.Empty

                ' 查找此節點的ValuePath '0;0;0;0/1;0;1;0
                sParentValuePath = findParentNodeValuePath(treeView, sParent)

                If Not treeView.FindNode(sParentValuePath & "/" & subNode.Value) Is Nothing Then
                    treeView.FindNode(sParentValuePath & "/" & subNode.Value).Text = String.Format("<font color='red'>{0}</font>", treeView.FindNode(sParentValuePath & "/" & subNode.Value).Text)
                    '  treeView.FindNode(sParentValuePath & "/" & subNode.Value).Selected = True
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
                        sTempList &= subNode.Value.Split(";")(0).ToString() & "*"
                    End If
                Else
                    For Each node As TreeNode In parent.ChildNodes
                        sTempList &= node.Value.Split(";")(0).ToString() & "*"
                    Next

                    If Not sTempList.Contains(subNode.Value.Split(";")(0)) Then
                        parent.ChildNodes.Add(subNode)
                    Else

                        If Not treeView.SelectedNode Is Nothing Then
                            parent.ChildNodes.Remove(treeView.SelectedNode)

                            subNode.Value = hidValue.Value
                            parent.ChildNodes.Add(subNode)

                            subNode.Select()
                        End If
                    End If

                    For Each item As TreeNode In parent.ChildNodes
                        If subNode.Value.Split(";")(0) > 0 Then
                            If hidValue.Value = "" Then
                                hidValue.Value = subNode.Value
                            End If

                            If item.Value.Split(";")(0) = subNode.Value.Split(";")(0) Then
                                parent.ChildNodes.Remove(item)

                                Dim node As New TreeNode
                                node.Value = hidValue.Value
                                node.Text = String.Format("<font color='red'>{0}</font>", sTempName)
                                parent.ChildNodes.Add(node)

                                Exit For
                            End If

                        End If

                        If item.Value.Split(";")(0) = subNode.Value.Split(";")(0) Then
                            item.Text = String.Format("<font color='red'>{0}</font>", item.Text)
                        End If
                    Next
                End If
            Next
        Catch ex As Exception
            Throw ex
        Finally
            ds.Dispose()
            dt.Dispose()
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
            Dim row() As DataRow = CType(ViewState("SY_FunctionCodeList"), DataTable).Select("FUNCCODE=" & iParentId)

            If row.Length > 0 Then

                ' 查找父節點的ValuePath
                sValuePath = findParentNodeValuePath(tree, row(0)("PARENTID"))

                ' 組織此節點的ValuePath
                sValuePath &= "/" & Server.HtmlEncode(row(0)("FUNCCODE") & ";" & row(0)("FUCNAME").ToString().Replace("/", COL_DELIM) & ";" & row(0)("FUCURL").ToString().Replace("/", "#") &
                                 ";" & row(0)("PARENTID") & ";" & row(0)("DISABLED") &
                                 ";" & row(0)("SORTCTRL") & ";" & row(0)("HOFLAG"))
            End If
        ElseIf iParentId < 0 Then

            ' 查找存在于ViewState中Roleid為Parent的資料【父節點的資料來源是Temp表】
            Dim row() As DataRow = CType(ViewState("SY_FunctionCodeList_Temp"), DataTable).Select("FUNCCODE='" & iParentId & "'")

            If row.Length > 0 Then

                ' 查找父節點的ValuePath
                sValuePath = findParentNodeValuePath(tree, row(0)("PARENT"))

                ' 組織此節點的ValuePath
                sValuePath &= "/" & Server.HtmlEncode(row(0)("FUNCCODE") & ";" & row(0)("FUCNAME").ToString().Replace("/", COL_DELIM) & ";" & row(0)("FUCURL").ToString().Replace("/", "#") &
                                 ";" & row(0)("PARENT") & ";" & row(0)("DISABLED") &
                                 ";" & row(0)("SORTCTRL") & ";" & row(0)("HOFLAG"))
            End If
        Else
            ' 如果Parent為0，為父節點，直接取其ValuePath
            'ROLEID;PARENTID;ROLETYPE;DISABLED
            sValuePath = tree.Nodes(0).ValuePath
        End If

        Return sValuePath
    End Function

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
        filterExpr = "PARENTID = " & PId

        '資料篩選並把結果傳入Rows
        rows = dt.Select(filterExpr)

        '如果篩選結果有資料
        If rows.GetUpperBound(0) >= 0 Then

            Dim childNode As TreeNode
            Dim rc As String

            '逐筆取出篩選後資料
            For i As Integer = 0 To rows.GetLength(0) - 1

                '放入相關變數中
                Dim row As DataRow = rows(i)

                '實體化新節點
                childNode = New TreeNode

                '設定節點各屬性
                childNode.Text = row("FUCNAME").ToString & "  (" & row("FUNCCODE").ToString & ")"
                childNode.ToolTip = childNode.Text

                Dim strFucurlTmp As String = String.Empty
                strFucurlTmp = row("FUCURL").ToString().Replace("/", "#")
                strFucurlTmp = strFucurlTmp.Replace("&", "＆")

                '子節點的Value組成=FUNCCODE;FUCNAME;FUCURL;PARENT;DISABLED;SORTCTRL;HOFLAG
                childNode.Value = Server.HtmlEncode(row("FUNCCODE").ToString() & ";" & row("FUCNAME").ToString().Replace("/", COL_DELIM) & ";" &
                                  strFucurlTmp & ";" & row("PARENTID").ToString() & ";" &
                                 row("DISABLED").ToString() & ";" & row("SORTCTRL").ToString() & ";" &
                                 row("HOFLAG").ToString())

                '將節點加入Tree中
                tNode.ChildNodes.Add(childNode)

                If Not childNode.Parent Is Nothing Then
                    If childNode.Parent.Depth = 0 Then
                        childNode.ImageUrl = com.Azion.EloanUtility.UIUtility.getImgPath & "map2.gif"
                    End If
                End If

                '呼叫遞回取得子節點
                rc = addNodes(childNode, row("FUNCCODE"), dt)

                If childNode.Parent.Depth > 0 AndAlso childNode.ChildNodes.Count = 0 Then
                    childNode.ImageUrl = com.Azion.EloanUtility.UIUtility.getImgPath & "file.gif"
                ElseIf childNode.Parent.Depth > 0 Then
                    childNode.ImageUrl = com.Azion.EloanUtility.UIUtility.getImgPath & "folder.png"
                End If

            Next
        End If
    End Function

    ''' <summary>
    ''' 資料加載
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function initData()
        Try
            Dim apCodelist As New AP_CODEList(GetDatabaseManager())
            'Dim tempFunctionList As New TEMPFUNCTIONLIST(GetDatabaseManager())
            ' Dim tempFunctionListList As New TEMPFUNCTIONLISTList(GetDatabaseManager())
            Dim syFunctionList As New AUTH_OP.SY_FUNCTION_CODEList(GetDatabaseManager())

            Dim sySysId As New AUTH_OP.SY_SYSID(GetDatabaseManager())
            Dim dtRootData As New DataTable
            Dim dtParentData As New DataTable
            Dim dtTempData As New DataTable

            Dim sValue As String = String.Empty
            Dim sStringValue As String = String.Empty

            ' 綁定屬性
            If apCodelist.loadByUpCode(m_sUpCode2370) Then
                rdoListHoFlag.DataSource = apCodelist.getCurrentDataSet
                rdoListHoFlag.DataTextField = "TEXT"
                rdoListHoFlag.DataValueField = "VALUE"
                rdoListHoFlag.DataBind()
            End If

            ' 開始事物
            GetDatabaseManager().beginTran()

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
                    treeView.Nodes.Add(rootNode)
                Next
            End If

            ' (2)查詢角色資料的第一層資料：parent=0
            If syFunctionList.loadDataByType(0) Then
                dtParentData = syFunctionList.getCurrentDataSet.Tables(0)
            End If

            ' (3)循環 新增到TreeView中
            If Not dtParentData Is Nothing Then
                For Each dr As DataRow In dtParentData.Rows

                    ' 父節點
                    Dim parentNode As TreeNode = New TreeNode()
                    parentNode.Value = dr("FUNCCODE").ToString()
                    parentNode.Text = dr("FUCNAME").ToString

                    ' (4)添加子節點
                    addNode(parentNode, dr("FUNCCODE").ToString())

                    ' 添加到treeView中
                    treeView.Nodes(0).ChildNodes.Add(parentNode)
                Next
            End If

            ' 正常區域加載資料,查詢xml中的資料，將其新增到左邊目錄樹
            operationXML("1")

            ' 提交
            GetDatabaseManager().commit()
        Catch ex As Exception

            ' 事物回滾
            GetDatabaseManager.Rollback()
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Function

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
                If i < node.ChildNodes.Count Then
                    If node.ChildNodes(i).Value.Split(";")(0) = sParent Then
                        node.ChildNodes(i).ChildNodes.Add(subNode)
                    End If

                    checkExistNode(node.ChildNodes(i), subNode, sParent)
                End If
            Next
        End If
    End Function

    ''' <summary>
    '''  判斷DB中是否有此值 若有 將其設置為紅色
    ''' </summary>
    ''' <param name="node"></param>
    ''' <param name="sFuncCode"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function checkNode(ByVal node As TreeNode, ByVal sFuncCode As String)
        For Each item As TreeNode In node.ChildNodes
            If item.Value.Split(";")(0) = sFuncCode Then
                item.Text = String.Format("<font color='red'>{0}</font>", item.Text)
            End If

            checkNode(item, sFuncCode)
        Next
    End Function

    ''' <summary>
    ''' 新增TreeView節點
    ''' </summary>
    ''' <param name="node">節點</param>
    ''' <param name="item">parent</param>
    ''' <remarks></remarks>
    Private Sub addNode(ByVal node As TreeNode, ByVal item As String)
        Try
            Dim syFunctionCodeList As New AUTH_OP.SY_FUNCTION_CODEList(GetDatabaseManager())

            Dim dtData As New DataTable

            ' 根據RoleId取得所有的子節點
            If syFunctionCodeList.loadDataByType(item) Then
                dtData = syFunctionCodeList.getCurrentDataSet.Tables(0)
            End If

            ' 循環添加子節點
            If Not dtData Is Nothing Then
                If dtData.Rows.Count > 0 Then
                    For Each dr As DataRow In dtData.Rows

                        Dim subNode As TreeNode = New TreeNode
                        subNode.Text = dr("FUCNAME").ToString()
                        subNode.Value = dr("FUNCCODE").ToString

                        addNode(subNode, dr("FUNCCODE").ToString())
                        node.ChildNodes.Add(subNode)
                    Next
                End If
            End If
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 控件顯示
    ''' </summary>
    ''' <remarks></remarks>
    Sub displayControl()
        lblParentName.Text = "-"
        lblName.Text = "安泰銀行"
        lblName.Visible = True
        trId.Visible = False
        txtName.Visible = False
        trFuncUrl.Visible = False
        trSortCtrl.Visible = False
        trHoFlag.Visible = False
        trDisable.Visible = False
        btnSave.Visible = False
        btnCancel.Visible = False
        btnAdd.Visible = True
        btnDelete.Visible = False
    End Sub

    ''' <summary>
    ''' 公用方法 新增歷史檔
    ''' </summary>
    ''' <param name="sApproved"></param>
    ''' <remarks></remarks>
    Sub insertDataToHis(ByVal sApproved As String)
        Try
            ' 開始事物
            GetDatabaseManager().beginTran()

            operationXML("2", sApproved)

            ' 提交事物
            GetDatabaseManager().commit()
        Catch ex As Exception

            ' 回滾事務
            GetDatabaseManager.Rollback()
            Throw ex
        End Try
    End Sub

    ''' <summary>
    ''' 正常資料加載
    ''' </summary>
    ''' <param name="trNode"></param>
    ''' <remarks></remarks>
    Sub loadNormal(ByVal trNode As TreeNode)
        Try
            ' 聲明控件
            Dim syFunctionCode As New AUTH_OP.SY_FUNCTION_CODE(GetDatabaseManager())

            '   treeView.SelectedNode.Value = trNode.Value

            Dim dt As DataTable = loadSYTempInfo()
            Dim sParent As String = String.Empty

            If Not dt Is Nothing Then
                Dim rows() As DataRow = dt.Select("FUNCCODE='" & trNode.Value.Split(";")(0) & "'")
                If rows.Length > 0 Then
                    lblID.Text = rows(0)("FUNCCODE").ToString
                    txtName.Text = rows(0)("FUCNAME").ToString
                    txtFuncUrl.Text = rows(0)("FUCURL").ToString
                    txtSortCtrl.Text = rows(0)("SORTCTRL").ToString

                    ' 屬性
                    If Not String.IsNullOrEmpty(txtFuncUrl.Text) Then
                        trHoFlag.Visible = True

                        If Not String.IsNullOrEmpty(rows(0)("HOFLAG").ToString) Then

                            rdoListHoFlag.SelectedValue = rows(0)("HOFLAG").ToString
                        Else
                            rdoListHoFlag.SelectedIndex = 0
                        End If
                    Else
                        trHoFlag.Visible = False

                        rdoListHoFlag.SelectedIndex = 0
                    End If

                    rdoDisable.SelectedValue = rows(0)("DISABLED").ToString
                    sParent = rows(0)("PARENT").ToString

                    hidId.Value = rows(0)("FUNCCODE").ToString
                    hidName.Value = rows(0)("FUCNAME").ToString
                End If
            End If

            ' 初始為黑色
            txtName.ForeColor = Drawing.Color.Black
            txtFuncUrl.ForeColor = Drawing.Color.Black
            txtSortCtrl.ForeColor = Drawing.Color.Black

            For Each item As ListItem In rdoDisable.Items
                item.Text = item.Text.Replace("<font color='red'>", "").Replace("</font>", "")
            Next

            For Each item As ListItem In rdoListHoFlag.Items
                item.Text = item.Text.Replace("<font color='red'>", "").Replace("</font>", "")
            Next

            ' 查詢臨時檔中的資料
            Dim flag As Boolean = False
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                For Each dr As DataRow In dt.Rows
                    If dr("FUNCCODE").ToString() = lblID.Text Then
                        flag = True
                        Exit For
                    End If
                Next
            End If

            If flag = True Then
                txtName.ForeColor = Drawing.Color.Red
                txtFuncUrl.ForeColor = Drawing.Color.Red
                txtSortCtrl.ForeColor = Drawing.Color.Red

                If trHoFlag.Visible = True Then
                    If rdoListHoFlag.SelectedIndex <> "-1" Then
                        rdoListHoFlag.SelectedItem.Text = String.Format("<font color='red'>{0}</font>", rdoListHoFlag.SelectedItem.Text)
                    End If
                End If

                rdoDisable.SelectedItem.Text = String.Format("<font color='red'>{0}</font>", rdoDisable.SelectedItem.Text)
            Else
                txtName.ForeColor = Drawing.Color.Black
                txtFuncUrl.ForeColor = Drawing.Color.Black
                txtSortCtrl.ForeColor = Drawing.Color.Black

                If trHoFlag.Visible = True Then
                    If rdoListHoFlag.SelectedIndex <> "-1" Then
                        rdoListHoFlag.SelectedItem.Text = rdoListHoFlag.SelectedItem.Text.Replace("<font color='red'>", "").Replace("</font>", "")
                    End If
                End If

                rdoDisable.SelectedItem.Text = rdoDisable.SelectedItem.Text.Replace("<font color='red'>", "").Replace("</font>", "")
            End If

            ' 父節點的值
            Dim rowsParent() As DataRow = dt.Select("FUNCCODE='" & sParent & "'")
            If rowsParent.Length > 0 Then
                lblParentName.Text = rowsParent(0)("FUCNAME").ToString & "  (" & rowsParent(0)("FUNCCODE").ToString & ")"
            End If

            ' 父節點
            If Not treeView.SelectedNode.Parent Is Nothing Then
                hidParentId.Value = sParent ' treeView.SelectedNode.Parent.Value.Split(";")(0)
            End If

            '  正常區域，更新操作，根據Funccode查詢出相關的資料，給維護區域賦值
            operationXML("3")

            ' 檢核刪除按鈕是否可用
            If treeView.SelectedNode.Value.Split(";")(0) < 0 AndAlso m_sFourStepNo = "" Then
                btnDelete.Visible = True
            Else
                btnDelete.Visible = False
            End If

            ' 如果 lblID.Text<0 則說明是臨時檔的資料，此時將hidFlag的值改為Add,無論怎樣修改，都以紅色進行標示
            If lblID.Text < 0 Then
                ' hidFlag.Value = "Add"
            End If

            ' 記錄老的數據， lblID.Text>0 則說明是真實檔的資料， 若進行編輯，則以紅色進行標示 名稱&路徑&排序& 屬性&是否可用
            hidOldData.Value = txtName.Text.Replace("/", COL_DELIM) & ";" & txtFuncUrl.Text.Replace("/", "#") & ";" & txtSortCtrl.Text & ";" & rdoListHoFlag.SelectedValue & ";" & rdoDisable.SelectedValue

            btnAdd.Visible = True
            btnSave.Visible = True
            btnCancel.Visible = True
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Sub loadNormalData(ByVal sFunccode As String)
        Try
            ' 聲明控件
            Dim syFunctionCode As New AUTH_OP.SY_FUNCTION_CODE(GetDatabaseManager())

            Dim dt As DataTable = loadSYTempInfo()
            Dim sParent As String = String.Empty

            If Not dt Is Nothing Then
                Dim rows() As DataRow = dt.Select("FUNCCODE='" & sFunccode & "'")
                If rows.Length > 0 Then
                    lblID.Text = rows(0)("FUNCCODE").ToString
                    txtName.Text = rows(0)("FUCNAME").ToString
                    txtFuncUrl.Text = rows(0)("FUCURL").ToString
                    txtSortCtrl.Text = rows(0)("SORTCTRL").ToString

                    ' 屬性
                    If Not String.IsNullOrEmpty(txtFuncUrl.Text) Then
                        trHoFlag.Visible = True

                        If Not String.IsNullOrEmpty(rows(0)("HOFLAG").ToString) Then

                            rdoListHoFlag.SelectedValue = rows(0)("HOFLAG").ToString
                        Else
                            rdoListHoFlag.SelectedIndex = 0
                        End If
                    Else
                        trHoFlag.Visible = False

                        rdoListHoFlag.SelectedIndex = 0
                    End If

                    rdoDisable.SelectedValue = rows(0)("DISABLED").ToString
                    sParent = rows(0)("PARENT").ToString

                    hidId.Value = rows(0)("FUNCCODE").ToString
                    hidName.Value = rows(0)("FUCNAME").ToString
                    hidParentId.Value = rows(0)("PARENT").ToString

                    ' 父節點的值
                    Dim rowsParent() As DataRow = dt.Select("FUNCCODE='" & sParent & "'")
                    If rowsParent.Length > 0 Then
                        lblParentName.Text = rowsParent(0)("FUCNAME").ToString & "  (" & rowsParent(0)("FUNCCODE").ToString & ")"
                    Else
                        lblParentName.Text = treeView.SelectedNode.Parent.Text
                    End If

                    ' 父節點
                    If Not treeView.SelectedNode.Parent Is Nothing Then
                        hidParentId.Value = sParent ' treeView.SelectedNode.Parent.Value.Split(";")(0)
                    End If
                Else

                    lblID.Text = treeView.SelectedNode.Value.Split(";")(0)

                    ' 名稱
                    If treeView.SelectedNode.Value.Split(";")(1).Contains(COL_DELIM) Then
                        txtName.Text = treeView.SelectedNode.Value.Split(";")(1).Replace(COL_DELIM, "/")
                    Else
                        txtName.Text = treeView.SelectedNode.Value.Split(";")(1)
                    End If

                    ' 路徑
                    txtFuncUrl.Text = treeView.SelectedNode.Value.Split(";")(2).Replace("#", "/")

                    ' 排序
                    txtSortCtrl.Text = treeView.SelectedNode.Value.Split(";")(5)

                    ' 屬性
                    If Not String.IsNullOrEmpty(txtFuncUrl.Text) Then
                        trHoFlag.Visible = True

                        If Not String.IsNullOrEmpty(treeView.SelectedNode.Value.Split(";")(6).Trim) Then

                            rdoListHoFlag.SelectedValue = treeView.SelectedNode.Value.Split(";")(6)
                        Else
                            rdoListHoFlag.SelectedIndex = 0
                        End If
                    Else
                        trHoFlag.Visible = False

                        rdoListHoFlag.SelectedIndex = 0
                    End If

                    ' 是否可用
                    rdoDisable.SelectedValue = treeView.SelectedNode.Value.Split(";")(4)

                    hidParentId.Value = treeView.SelectedNode.Parent.Value.Split(";")(0)
                    lblParentName.Text = treeView.SelectedNode.Parent.Text
                    hidName.Value = txtName.Text
                    hidId.Value = lblID.Text
                End If
            Else

                lblID.Text = treeView.SelectedNode.Value.Split(";")(0)

                ' 名稱
                If treeView.SelectedNode.Value.Split(";")(1).Contains(COL_DELIM) Then
                    txtName.Text = treeView.SelectedNode.Value.Split(";")(1).Replace(COL_DELIM, "/")
                Else
                    txtName.Text = treeView.SelectedNode.Value.Split(";")(1)
                End If

                ' 路徑
                txtFuncUrl.Text = treeView.SelectedNode.Value.Split(";")(2).Replace("#", "/")

                ' 排序
                txtSortCtrl.Text = treeView.SelectedNode.Value.Split(";")(5)

                ' 屬性
                If Not String.IsNullOrEmpty(txtFuncUrl.Text) Then
                    trHoFlag.Visible = True

                    If Not String.IsNullOrEmpty(treeView.SelectedNode.Value.Split(";")(6).Trim) Then

                        rdoListHoFlag.SelectedValue = treeView.SelectedNode.Value.Split(";")(6)
                    Else
                        rdoListHoFlag.SelectedIndex = 0
                    End If
                Else
                    trHoFlag.Visible = False

                    rdoListHoFlag.SelectedIndex = 0
                End If

                ' 是否可用
                rdoDisable.SelectedValue = treeView.SelectedNode.Value.Split(";")(4)

                hidParentId.Value = treeView.SelectedNode.Parent.Value.Split(";")(0)
                lblParentName.Text = treeView.SelectedNode.Parent.Text
                hidName.Value = txtName.Text
                hidId.Value = lblID.Text
            End If

            ' 初始為黑色
            txtName.ForeColor = Drawing.Color.Black
            txtFuncUrl.ForeColor = Drawing.Color.Black
            txtSortCtrl.ForeColor = Drawing.Color.Black

            For Each item As ListItem In rdoDisable.Items
                item.Text = item.Text.Replace("<font color='red'>", "").Replace("</font>", "")
            Next

            For Each item As ListItem In rdoListHoFlag.Items
                item.Text = item.Text.Replace("<font color='red'>", "").Replace("</font>", "")
            Next

            ' 查詢臨時檔中的資料
            Dim flag As Boolean = False
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                For Each dr As DataRow In dt.Rows
                    If dr("FUNCCODE").ToString() = lblID.Text Then
                        flag = True
                        Exit For
                    End If
                Next
            End If

            If flag = True Then
                txtName.ForeColor = Drawing.Color.Red
                txtFuncUrl.ForeColor = Drawing.Color.Red
                txtSortCtrl.ForeColor = Drawing.Color.Red

                If trHoFlag.Visible = True Then
                    If rdoListHoFlag.SelectedIndex <> "-1" Then
                        rdoListHoFlag.SelectedItem.Text = String.Format("<font color='red'>{0}</font>", rdoListHoFlag.SelectedItem.Text)
                    End If
                End If

                rdoDisable.SelectedItem.Text = String.Format("<font color='red'>{0}</font>", rdoDisable.SelectedItem.Text)
            Else
                txtName.ForeColor = Drawing.Color.Black
                txtFuncUrl.ForeColor = Drawing.Color.Black
                txtSortCtrl.ForeColor = Drawing.Color.Black

                If trHoFlag.Visible = True Then
                    If rdoListHoFlag.SelectedIndex <> "-1" Then
                        rdoListHoFlag.SelectedItem.Text = rdoListHoFlag.SelectedItem.Text.Replace("<font color='red'>", "").Replace("</font>", "")
                    End If
                End If

                rdoDisable.SelectedItem.Text = rdoDisable.SelectedItem.Text.Replace("<font color='red'>", "").Replace("</font>", "")
            End If

            '  正常區域，更新操作，根據Funccode查詢出相關的資料，給維護區域賦值
            '  operationXML("3")

            ' 檢核刪除按鈕是否可用
            If treeView.SelectedNode.Value.Split(";")(0) < 0 AndAlso m_sFourStepNo = "" Then
                btnDelete.Visible = True
            Else
                btnDelete.Visible = False
            End If

            ' 記錄老的數據， lblID.Text>0 則說明是真實檔的資料， 若進行編輯，則以紅色進行標示 名稱&路徑&排序& 屬性&是否可用
            hidOldData.Value = txtName.Text.Replace("/", COL_DELIM) & ";" & txtFuncUrl.Text.Replace("/", "#") & ";" & txtSortCtrl.Text & ";" & rdoListHoFlag.SelectedValue & ";" & rdoDisable.SelectedValue

            btnAdd.Visible = True
            btnSave.Visible = True
            btnCancel.Visible = True
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ''' <summary>
    ''' 回覆到初始狀態
    ''' </summary>
    ''' <param name="node"></param>
    ''' <param name="sParent"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function displayNode(ByVal node As TreeNode, ByVal sParent As String) As String
        If node.ChildNodes.Count > 0 Then
            For i As Integer = 0 To node.ChildNodes.Count - 1

                If node.ChildNodes(i).Value = sParent Then
                    loadNormal(node.ChildNodes(i))
                Else
                    displayNode(node.ChildNodes(i), sParent)
                End If
            Next
        End If
    End Function

    ''' <summary>
    ''' 若父節點的啟用狀態改變，遞歸設置其子節點的狀態，應與其父節點一致---->臨時檔資料
    ''' </summary>
    ''' <param name="dtSYTempInfo"></param>
    ''' <param name="sRoleId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function modifyData(ByVal dtSYTempInfo As DataTable, ByVal sFunccode As String) As DataTable

        ' 若父節點為停用 則子節點也為停用,若父節點為啟用，所有的都為啟用
        If dtSYTempInfo.Select("[parent]='" & sFunccode & "'").Count > 0 Then
            For Each childRow As DataRow In dtSYTempInfo.Select("[parent]='" & sFunccode & "'")
                childRow.BeginEdit()
                Dim rowsT() As DataRow = dtSYTempInfo.Select("FUNCCODE='" & sFunccode & "'")
                childRow("DISABLED") = rdoDisable.SelectedValue
                childRow.EndEdit()
                dtSYTempInfo.AcceptChanges()

                If dtSYTempInfo.Select("[parent]='" & childRow("FUNCCODE") & "'").Count > 0 Then
                    modifyData(dtSYTempInfo, childRow("FUNCCODE").ToString())
                End If
            Next
        End If

        Return dtSYTempInfo
    End Function

    ''' <summary>
    ''' 若父節點的啟用狀態改變，遞歸設置其子節點的狀態，應與其父節點一致---->真實檔資料
    ''' </summary>
    ''' <param name="sRoleId">角色編號</param>
    ''' <param name="sDisabled">父節點的啟用狀態</param>
    ''' <remarks></remarks>
    Sub modifyLiveData(ByRef sRoleId As String, ByVal sDisabled As String)
        Dim syFunccodeList As New AUTH_OP.SY_FUNCTION_CODEList(GetDatabaseManager)

        ' 如果目前只是修改的父節點，則需要根據其roleid查找是否有子節點，將其disable設置為與
        If syFunccodeList.loadDataByType(sRoleId) Then

            Dim dt As DataTable = syFunccodeList.getCurrentDataSet.Tables(0)
            If Not dt Is Nothing Then
                If dt.Rows.Count > 0 Then
                    For Each dr As DataRow In dt.Rows

                        hidParentFuncList.Value &= dr("FUNCCODE") & ","

                        If syFunccodeList.loadDataByType(dr("FUNCCODE")) Then
                            If Not syFunccodeList.getCurrentDataSet.Tables(0) Is Nothing Then
                                modifyLiveData(dr("FUNCCODE"), sDisabled)
                            End If
                        End If
                    Next
                End If
            End If
        End If
    End Sub
#End Region

#Region "正常區塊Event"

    ''' <summary>
    ''' 新增
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnAdd_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnAdd.Click
        ' 取得編號
        Dim syFunctionCode As New SY_FUNCTION_CODE(GetDatabaseManager())
        Dim syTempInfo As New SY_TEMPINFO(GetDatabaseManager())
        Dim sFuncCode As String = String.Empty
        Dim sParent As String = String.Empty
        Dim dtTempData As New DataTable

        rdoListHoFlag.SelectedIndex = 0
        txtName.ForeColor = Drawing.Color.Black
        txtFuncUrl.ForeColor = Drawing.Color.Black
        txtSortCtrl.ForeColor = Drawing.Color.Black
        rdoListHoFlag.ForeColor = Drawing.Color.Black
        rdoDisable.ForeColor = Drawing.Color.Black

        ' 如果只有安泰銀行 還沒有其他Child
        If treeView.Nodes.Count = 1 Then

            For Each node As TreeNode In treeView.Nodes
                If node.ChildNodes.Count = 0 Then
                    lblParentName.Text = "安泰銀行"
                    hidParentName.Value = lblParentName.Text
                    btnAdd.Visible = False
                Else
                    If hidName.Value <> "" Then
                        lblParentName.Text = hidName.Value & "  (" & hidId.Value & ")"
                    Else
                        lblParentName.Text = "安泰銀行"
                        hidParentName.Value = lblParentName.Text
                    End If

                    txtName.Text = ""
                    btnAdd.Visible = True
                End If
            Next
        End If

        trId.Visible = True
        trName.Visible = True
        trSortCtrl.Visible = True
        trFuncUrl.Visible = True
        trDisable.Visible = True
        trHoFlag.Visible = False

        txtName.Visible = True
        lblName.Visible = False
        btnSave.Visible = True
        btnCancel.Visible = True
        btnAdd.Visible = False

        ' 狀態默認為啟用
        rdoDisable.SelectedValue = "0"
        txtFuncUrl.Text = ""
        txtName.Text = ""
        lblID.Text = 0

        'If Not treeView.SelectedNode Is Nothing Then

        '    If treeView.SelectedNode.Value.Split(";")(2) <> "" Then
        '        com.Azion.EloanUtility.UIUtility.alert("不可新增子節點！")
        '        lblParentName.Text = hidParentName.Value

        '        loadNormal(treeView.SelectedNode)
        '        Exit Sub
        '    End If
        'Else
        '    If hidValue.Value.Length > 0 Then
        '        If hidValue.Value.Split(";")(2) <> "" Then
        '            com.Azion.EloanUtility.UIUtility.alert("不可新增子節點！")
        '            lblParentName.Text = hidParentName.Value

        '            Exit Sub
        '        End If
        '    End If
        'End If

        ' 路徑是否為空檢核
        If hidValue.Value.Length > 0 Then
            If hidValue.Value.Split(";")(2) <> "" Then
                com.Azion.EloanUtility.UIUtility.alert("不可新增子節點！")
                loadNormalData(hidValue.Value.Split(";")(0))
                Exit Sub
            End If
        End If

        ' 新增時，啟用狀態檢核
        If Not IsNothing(treeView.SelectedNode) Then

            ' 選擇非安泰銀行交易
            If treeView.SelectedNode.Value.Split(";")(0) <> "0" Then
                If hidValue.Value.Length > 0 Then

                    ' 交易啟用狀態為停用
                    If hidValue.Value.Split(";")(4) <> "0" Then
                        com.Azion.EloanUtility.UIUtility.alert("交易已停用，不可新增子節點！")
                        loadNormalData(hidValue.Value.Split(";")(0))
                        Exit Sub
                    End If
                End If
            End If
        Else
            Return
        End If

        ' 取得真實檔的最大的順序欄位
        txtSortCtrl.Text = syFunctionCode.getMaxSortCtrl

        ' 取得臨時檔中的最大的順序欄位
        Dim sXmlData As String = String.Empty
        Dim dtData As New DataTable
        Dim dtView As New DataView
        Dim dtViewCode As New DataView
        Dim iSort As Integer = 0

        If syTempInfo.loadByPK(m_sWorkingUserid, m_sWorkingTopDepNo, m_sFuncCode) Then
            sXmlData = syTempInfo.getAttribute("TEMPDATA")
        End If

        lblID.Text = "-1"
        dtData = com.Azion.EloanUtility.XmlUtility.convertXML2DataTable(sXmlData)

        If Not IsNothing(dtData) AndAlso dtData.Rows.Count > 0 Then
            dtView = dtData.DefaultView
            dtView.Sort = "SORTCTRL desc"

            dtData = dtView.ToTable

            iSort = Convert.ToInt32(dtData.Rows(0)("SORTCTRL").ToString()) + 10

            '類型轉換
            If Not dtData.Columns.Contains("iRoleId") Then dtData.Columns.Add("iRoleId", System.Type.GetType("System.Int32"))
            dtData.Columns("iRoleId").Expression = "Convert(funccode, 'System.Int32')"
            Dim iMin As Integer = dtData.Compute("min(iRoleId)", "")
            If iMin < 0 Then
                lblID.Text = iMin - 1
            End If
        End If

        If iSort > syFunctionCode.getMaxSortCtrl Then
            txtSortCtrl.Text = iSort.ToString
        End If
        hidFlag.Value = "Add"
    End Sub

    ''' <summary>
    ''' 取消
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnCancel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnCancel.Click

        treeView_SelectedNodeChanged(sender, e)

        '' 點擊一筆資料，編輯，點擊【取消】，返回當前資料的初始狀態
        'If hidId.Value = "0" Then hidId.Value = ""
        'If (hidFlag.Value = "" AndAlso hidId.Value <> "") Then

        '    If String.IsNullOrEmpty(m_sFlag) Then

        '        ' .頁面進入直接點擊【新增】再點擊【取消】，顯示【安泰銀行】資料
        '        displayControl()

        '        hidName.Value = ""
        '        hidId.Value = ""
        '        hidParentId.Value = ""
        '        hidParentName.Value = ""
        '    Else
        '        ' 通過Hidid的Value值取得對應的TreeNode
        '        For Each node As TreeNode In treeView.Nodes
        '            If node.Value = hidId.Value Then
        '                loadNormal(node)
        '            Else
        '                displayNode(node, hidId.Value)
        '            End If
        '        Next

        '        m_sFlag = ""
        '    End If
        'Else
        '    If hidFlag.Value = "Add" Then

        '        '  點擊一筆資料，點擊新增
        '        If hidId.Value <> "" Then

        '            ' 通過Hidid的Value值取得對應的TreeNode
        '            For Each node As TreeNode In treeView.Nodes
        '                If node.Value = hidId.Value Then
        '                    loadNormal(node)
        '                Else
        '                    displayNode(node, hidId.Value)
        '                End If
        '            Next
        '        Else
        '            ' .頁面進入直接點擊【新增】再點擊【取消】，顯示【安泰銀行】資料
        '            displayControl()

        '            hidName.Value = ""
        '            hidId.Value = ""
        '            hidParentId.Value = ""
        '            hidParentName.Value = ""
        '        End If
        '    Else
        '        displayControl()

        '        hidName.Value = ""
        '        hidId.Value = ""
        '        hidParentId.Value = ""
        '        hidParentName.Value = ""
        '    End If
        'End If

        hidValue.Value = ""
        hidFlag.Value = ""
        hidName.Value = ""
        hidId.Value = ""
        hidParentId.Value = ""
        hidParentName.Value = ""
    End Sub

    ''' <summary>
    ''' 儲存ok
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnSave_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnSave.Click
        Try
            Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager())
            Dim sbXmlData As New StringBuilder
            Dim newXmlData As String = String.Empty
            Dim syFunctionCodeT As New AUTH_OP.SY_FUNCTION_CODE(GetDatabaseManager())
            Dim sFuncName As String = String.Empty
            Dim dtTempData As New DataTable
            Dim sParent As String = String.Empty
            Dim iRoleId As Integer = lblID.Text

            If txtName.Text.Trim = "" Then
                com.Azion.EloanUtility.UIUtility.alert("請輸入交易項目！")
                Return
            End If

            ' 查看其父節點的啟用狀態 如果為停用 則不可以修改子節點的狀態
            If Not treeView.SelectedNode.Parent Is Nothing Then
                If treeView.SelectedNode.Parent.Value.Split(";")(4) = "1" Then
                    com.Azion.EloanUtility.UIUtility.alert("父節點已經停用，子節點不可改為啟用！")
                    rdoDisable.SelectedValue = 1

                    For Each item As ListItem In rdoDisable.Items
                        item.Text = item.Text.Replace("<font color='red'>", "").Replace("</font>", "")
                    Next

                    Return
                End If
            End If

            ' 組XML文件
            sbXmlData.Append("<SY>")
            sbXmlData.Append("<SY_FUNCTION_CODE>")
            sbXmlData.Append("<FUNCCODE>" & lblID.Text.Trim & "</FUNCCODE>")

            If Not treeView.SelectedNode Is Nothing Then

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
                    If iRoleId < 0 Then
                        '  sbXmlData.Append("<PARENT>" & treeView.SelectedNode.Value.Split(";")(0) & "</PARENT>")
                        ' Modify by 2012/06/28
                        If lblID.Text = "-1" Then

                            ' 判斷是新增還是修改
                            If hidFlag.Value = "Modify" Then
                                sbXmlData.Append("<PARENT>" & treeView.SelectedNode.Value.Split(";")(3) & "</PARENT>")
                                sParent = treeView.SelectedNode.Value.Split(";")(3)
                            ElseIf hidFlag.Value = "" Then
                                sbXmlData.Append("<PARENT>" & treeView.SelectedNode.Value.Split(";")(3) & "</PARENT>")
                                sParent = treeView.SelectedNode.Value.Split(";")(3)
                            Else
                                sbXmlData.Append("<PARENT>" & treeView.SelectedNode.Value.Split(";")(0) & "</PARENT>")
                                sParent = treeView.SelectedNode.Value.Split(";")(0)
                            End If
                        Else
                            'sbXmlData.Append("<PARENT>" & treeView.SelectedNode.Value.Split(";")(1) & "</PARENT>")
                            If hidFlag.Value = "Modify" Then
                                sbXmlData.Append("<PARENT>" & treeView.SelectedNode.Value.Split(";")(3) & "</PARENT>")
                                sParent = treeView.SelectedNode.Value.Split(";")(3)
                            ElseIf hidFlag.Value = "" Then
                                sbXmlData.Append("<PARENT>" & treeView.SelectedNode.Value.Split(";")(3) & "</PARENT>")
                                sParent = treeView.SelectedNode.Value.Split(";")(3)
                            Else
                                sbXmlData.Append("<PARENT>" & treeView.SelectedNode.Value.Split(";")(0) & "</PARENT>")
                                sParent = treeView.SelectedNode.Value.Split(";")(0)
                            End If
                        End If

                    Else
                        sbXmlData.Append("<PARENT>" & treeView.SelectedNode.Value.Split(";")(3) & "</PARENT>")
                        sParent = treeView.SelectedNode.Value.Split(";")(3)
                    End If
                End If
            Else
                sbXmlData.Append("<PARENT>0</PARENT>")
                sParent = "0"
            End If

            sbXmlData.Append("<FUCNAME>" & txtName.Text.Trim & "</FUCNAME>")
            sbXmlData.Append("<FUCURL>" & txtFuncUrl.Text.Trim.Replace("&", "＆") & "</FUCURL>")
            sbXmlData.Append("<DISABLED>" & rdoDisable.SelectedValue & "</DISABLED>")
            sbXmlData.Append("<SORTCTRL>" & txtSortCtrl.Text.Trim & "</SORTCTRL>")

            If trHoFlag.Visible = True Then
                sbXmlData.Append("<HOFLAG>" & rdoListHoFlag.SelectedValue & "</HOFLAG>")
            Else
                sbXmlData.Append("<HOFLAG></HOFLAG>")
            End If

            ' 所屬子系統的父節點 TODO: 第二層與第三層取值待調整
            sbXmlData.Append("</SY_FUNCTION_CODE>")
            sbXmlData.Append("</SY>")

            '------------------------------ 正常區域，儲存操作--------------------------------------------------- 
            operationXML("7")

            If m_sFuncCodeList.Length > 0 Then
                m_sFuncCodeList = m_sFuncCodeList.Substring(0, m_sFuncCodeList.Length - 1)
            End If

            If syFunctionCodeT.loadDataByCon(txtName.Text.Trim, lblID.Text, m_sFuncCodeList) Then
                sFuncName = syFunctionCodeT.getAttribute("FUCNAME")
            End If

            ' 驗證角色名稱是否存在
            'If m_sFuncNameList.Contains(txtName.Text.Trim) Or txtName.Text.Trim = sFuncName Then
            '    com.Azion.EloanUtility.UIUtility.alert("交易名稱不可重複，請重新輸入！！")
            '    txtName.Focus()

            '    Return
            'End If

            Dim dtSYTempInfo As DataTable = loadSYTempInfo()
            Dim tempData As DataTable = com.Azion.EloanUtility.XmlUtility.convertXML2DataTable(sbXmlData.ToString)

            ' 查找臨時檔中的資料 進行修改
            If Not dtSYTempInfo Is Nothing Then
                Dim rows() As DataRow = dtSYTempInfo.Select("FUNCCODE='" & lblID.Text & "'")
                ' 找到臨時檔中已經存在的資料
                ' 如果有 則將其替換掉 然後累加上新的XML
                If rows.Length > 0 Then
                    dtSYTempInfo.Rows.Remove(rows(0))
                    dtSYTempInfo.AcceptChanges()
                End If

                '進行累加
                If Not tempData Is Nothing Then
                    Dim rowsNewData() As DataRow = tempData.Select("FUNCCODE='" & lblID.Text & "'")
                    For Each row As DataRow In rowsNewData
                        dtSYTempInfo.ImportRow(row)
                    Next
                End If

                dtSYTempInfo = modifyData(dtSYTempInfo, lblID.Text)

                dtSYTempInfo.AcceptChanges()
            Else ' 如果找不到臨時檔，直接新增資料
                dtSYTempInfo = tempData
            End If

            Dim dv As System.Data.DataView = dtSYTempInfo.DefaultView
            dv.Sort = "FUNCCODE asc"
            dtSYTempInfo = dv.ToTable

            '' 查詢除了此筆交易外，還存在其他交易有此名稱的資料
            'If syFunctionCodeT.loadByFuncName(txtName.Text.Trim, lblID.Text) Then
            '    com.Azion.EloanUtility.UIUtility.alert("交易名稱不可重複，請重新輸入！！")
            '    txtName.Focus()

            '    Return
            'End If

            Dim sHoFlag As String = String.Empty
            If trHoFlag.Visible = True Then
                sHoFlag = rdoListHoFlag.SelectedValue
            Else
                sHoFlag = ""
            End If

            hidValue.Value = lblID.Text & ";" & txtName.Text.Replace("/", COL_DELIM) & ";" & txtFuncUrl.Text.Replace("/", "#") & ";" & sParent & ";" & rdoDisable.SelectedValue & ";" & txtSortCtrl.Text & ";" & sHoFlag

            ' 沒有異動 直接點擊的存儲
            If Not treeView.SelectedNode Is Nothing Then
                If hidValue.Value.Trim = treeView.SelectedNode.Value.Trim Then
                    com.Azion.EloanUtility.UIUtility.alert("資料無變更，請確認是否已作修改!")
                    Return
                End If
            End If

            ' 開始事物
            GetDatabaseManager().beginTran()

            If m_sFourStepNo = "" Then
                If Not syTempInfo.loadByPK(m_sWorkingUserid, m_sWorkingTopDepNo, m_sFuncCode) Then
                    syTempInfo.setAttribute("STAFFID", m_sWorkingUserid)
                    syTempInfo.setAttribute("BRA_DEPNO", m_sWorkingTopDepNo)
                    syTempInfo.setAttribute("FUNCCODE", m_sFuncCode)
                End If

                syTempInfo.setAttribute("TEMPDATA", com.Azion.EloanUtility.XmlUtility.convertDataTable2XML(dtSYTempInfo).ToString)
                syTempInfo.save()
            Else
                If Not syTempInfo.loadByCaseId(m_sCaseId) Then
                    syTempInfo.setAttribute("CASEID", m_sCaseId)
                End If

                syTempInfo.setAttribute("TEMPDATA", com.Azion.EloanUtility.XmlUtility.convertDataTable2XML(dtSYTempInfo).ToString)
                syTempInfo.save()
            End If

            If iRoleId <= 0 Then
                initTempInfoTreeView(treeView)
                btnDelete.Visible = True
            Else
                ' 取得父節點 然後根據funccode找到修改的餓那筆treenode 然後刪除掉 然後再新增上新的
                Dim treeNode As TreeNode = treeView.FindNode(treeView.SelectedNode.Parent.ValuePath)
                If treeNode.ChildNodes.Count > 0 Then
                    For Each node As TreeNode In treeNode.ChildNodes
                        treeNode.ChildNodes.Remove(treeView.SelectedNode)
                        Dim subNode As New TreeNode
                        subNode.Value = hidValue.Value.Replace(COL_DELIM, "/")
                        subNode.Text = String.Format("<font color='red'>{0}</font>", hidValue.Value.Split(";")(1) & "  (" & hidValue.Value.Split(";")(0) & ")")

                        treeNode.ChildNodes.Add(subNode)
                        subNode.Select()

                        Exit For
                    Next
                End If
            End If

            ' 狀態為展開
            treeView.ExpandDepth = m_sTreeLevel

            '' 清空標識欄位
            'hidFlag.Value = ""
            hidName.Value = ""
            hidParentName.Value = ""

            If hidFlag.Value = "Add" Then
                treeView.FindNode(treeView.SelectedNode.ValuePath & "/" & hidValue.Value).Selected = True
            Else
                If iRoleId < 0 Then
                    treeView.FindNode(treeView.SelectedNode.Parent.ValuePath & "/" & hidValue.Value).Selected = True
                End If
            End If

            hidId.Value = lblID.Text
            hidName.Value = txtName.Text
            btnAdd.Visible = True
            btnDelete.Visible = True

            If Not dtSYTempInfo Is Nothing Then
                For Each dr As DataRow In dtSYTempInfo.Rows
                    rdoDisable.SelectedValue = dr("DISABLED").ToString
                Next
            End If

            ' 提交
            GetDatabaseManager().commit()

            com.Azion.EloanUtility.UIUtility.alert("儲存成功！")
        Catch ex As Exception

            ' 事物回滾
            GetDatabaseManager.Rollback()
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 更新
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub treeView_SelectedNodeChanged(ByVal sender As Object, ByVal e As EventArgs) Handles treeView.SelectedNodeChanged
        Try
            hidFlag.Value = "Modify"

            hidValue.Value = ""
            hidId.Value = ""
            hidName.Value = ""
            hidParentId.Value = ""
            hidParentName.Value = ""
            treeView.SelectedNodeStyle.ForeColor = Drawing.Color.Black

            If treeView.SelectedNode.Value.Split(";")(0) <> "0" Then

                trId.Visible = True
                trName.Visible = True
                trSortCtrl.Visible = True
                trFuncUrl.Visible = True
                trDisable.Visible = True

                ' 控件顯示隱藏
                txtName.Visible = True
                lblName.Visible = False
                hidValue.Value = treeView.SelectedNode.Value
                loadNormalData(treeView.SelectedNode.Value.Split(";")(0))
            Else
                'lblParentName.Text = "-"
                'lblName.Text = treeView.SelectedNode.Text
                'trId.Visible = False
                'trName.Visible = True
                'lblName.Visible = True
                'txtName.Visible = False
                'trFuncUrl.Visible = False
                'trSortCtrl.Visible = False
                'trHoFlag.Visible = False
                'trDisable.Visible = False
                'btnSave.Visible = False
                'btnCancel.Visible = False
                'btnDelete.Visible = False
                'hidName.Value = ""
                displayControl()
            End If

        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 刪除OK
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
            For Each childRow As DataRow In dtData.Select("[funccode]='" & lblID.Text & "'")
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

            ' 提示信息
            com.Azion.EloanUtility.UIUtility.alert("刪除成功！")

            treeView.Nodes.Clear()

            initFunctionTreeView(treeView)

            ' 添加節點【存在于Temp表】
            initTempInfoTreeView(treeView)

            treeView.ExpandDepth = m_sTreeLevel

            displayControl()

            btnDelete.Visible = False
            hidName.Value = ""
            hidId.Value = ""
            hidParentId.Value = ""
            hidParentName.Value = ""
            hidFlag.Value = ""
        Catch ex As Exception
            GetDatabaseManager().Rollback()
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 全部取消ok
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnCancelAll_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnCancelAll.Click
        Try
            ' 開始事物
            GetDatabaseManager().beginTran()

            Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager())
            Dim syFunctionCodeHis As New AUTH_OP.SY_FUNCTION_CODE_HIS(GetDatabaseManager())
            Dim flowFacade As New FLOW_OP.FlowFacade(GetDatabaseManager())

            ' 刪除sy_tempinfo
            If m_sFourStepNo = "" Then
                If syTempInfo.loadByPK(m_sWorkingUserid, m_sWorkingTopDepNo, m_sFuncCode) Then
                    syTempInfo.remove()
                End If
            Else
                If syTempInfo.loadByCaseId(m_sCaseId) Then
                    syTempInfo.remove()
                End If
            End If

            ' 刪除HIS
            syFunctionCodeHis.delHisDataByCaseId(m_sCaseId)

            ' 刪除流程
            flowFacade.DeleteFlow(m_sCaseId)

            ' 重新載入數據
            'initData()

            treeView.Nodes.Clear()

            ' 添加節點【存在于Live表】
            initFunctionTreeView(treeView)

            initTempInfoTreeView(treeView)

            ' 狀態為展開
            treeView.ExpandAll()
            ' treeView.ExpandDepth = m_sTreeLevel

            displayControl()

            hidId.Value = ""
            hidName.Value = ""
            hidParent.Value = ""
            hidParentId.Value = ""
            hidParentName.Value = ""
            hidFlag.Value = ""

            ' 提交
            GetDatabaseManager().commit()
        Catch ex As Exception

            ' 事物回滾
            GetDatabaseManager.Rollback()
            SYUIBase.showErrMsg(Me, ex)
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

            Dim syFunctionCodeHis As New SY_FUNCTION_CODE_HIS(GetDatabaseManager())
            Dim syTempInfo As New SY_TEMPINFO(GetDatabaseManager())
            Dim flowFacade As New FLOW_OP.FlowFacade(GetDatabaseManager())
            Dim sTempData As String = String.Empty

            ' 根據“登入者編號”，“部門代碼”，“功能編號”查詢人員變動資料
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
            End If

            operationXML("6")

            If m_sFuncNameList.Length > 0 Then
                m_sFuncNameList = m_sFuncNameList.Substring(0, m_sFuncNameList.Length - 1)
            End If

            If m_sFuncCodeList.Length > 0 Then
                m_sFuncCodeList = m_sFuncCodeList.Substring(0, m_sFuncCodeList.Length - 1)
            End If

            ' 查詢資料是否在送簽中
            If syFunctionCodeHis.loadDataByCon(m_sFuncNameList, m_sFuncCodeList, m_sStepNo, m_sCaseId) > 0 Then

                ' ROLENAME已在處理中，無法送件！
                com.Azion.EloanUtility.UIUtility.alert(lblName.Text & "已在處理中，無法送件！")
                Return
            End If

            GetDatabaseManager().beginTran()
            Dim stepInfo As FLOW_OP.StepInfo

            m_bCheck = True     '不再複核

            If m_bCheck = False Then

                If m_sStepNo = "" Then
                    '  stepInfo = flowFacade.StartFlow(m_sWorkingBrid, m_sFlowName, True, "")
                    stepInfo = MBSC.UICtl.UIShareFun.startFlow(GetDatabaseManager(), m_sFlowName, True)

                    Dim sNewCaseId = stepInfo.currentStepInfo.caseId

                    '不可能沒資料 
                    If syTempInfo.loadByPK(m_sWorkingUserid, m_sWorkingTopDepNo, m_sFuncCode) Then
                        syTempInfo.setAttribute("CASEID", sNewCaseId)
                        syTempInfo.save()
                    End If

                    com.Azion.EloanUtility.UIUtility.alert("存檔成功！")

                    ' 跳轉到代辦清單
                    com.Azion.EloanUtility.UIUtility.Redirect("SY_CASELIST.aspx")
                Else
                    stepInfo = flowFacade.SendFlow(m_sCaseId, 0, "", "")
                    com.Azion.EloanUtility.UIUtility.closeWindow()
                End If

                updatTemp2Live(GetDatabaseManager(), stepInfo, "")
            Else
                stepInfo = MBSC.UICtl.UIShareFun.startFlow(GetDatabaseManager(), m_sFlowName, True, , True)
                updatTemp2Live(GetDatabaseManager(), stepInfo, "Y")

                treeView.Nodes.Clear()

                ' 添加節點【存在于Live表】
                initFunctionTreeView(treeView)

                ' 添加節點【存在于Temp表】
                initTempInfoTreeView(treeView)

                'treeView.ExpandAll()
                treeView.ExpandDepth = m_sTreeLevel

                'displayControl()

                ' 成功提示信息
                com.Azion.EloanUtility.UIUtility.showJSMsg("案件編號：" + stepInfo.currentStepInfo.caseId + " 送出成功！")

                ' 跳轉到待辦事項頁面
                If String.IsNullOrEmpty(m_sStepNo) Then
                    com.Azion.EloanUtility.UIUtility.goMainPage("")
                Else
                    com.Azion.EloanUtility.UIUtility.closeWindow()
                End If

            End If

            GetDatabaseManager().commit()

            hidFlag.Value = ""
        Catch ex As Exception

            GetDatabaseManager.Rollback()
            Throw ex
        End Try
    End Sub

    ''' <summary>
    ''' 路徑輸入資料時，屬性行才顯示
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub txtFuncUrl_TextChanged(ByVal sender As Object, ByVal e As EventArgs) Handles txtFuncUrl.TextChanged

        If Not String.IsNullOrEmpty(txtFuncUrl.Text.Trim) Then

            If hidFlag.Value = "Add" Then
                txtFuncUrl.ForeColor = Drawing.Color.Red
            ElseIf hidFlag.Value = "Modify" Then

                If hidOldData.Value <> "" Then
                    If hidOldData.Value.Split(";")(1).ToString() <> txtFuncUrl.Text.Replace("/", "#") Then
                        txtFuncUrl.ForeColor = Drawing.Color.Red
                    Else
                        txtFuncUrl.ForeColor = Drawing.Color.Black
                    End If
                End If
            End If

            trHoFlag.Visible = True
        Else
            trHoFlag.Visible = False

            rdoListHoFlag.SelectedIndex = 0
        End If
    End Sub

    ''' <summary>
    ''' 名稱改變
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub txtName_TextChanged(sender As Object, e As EventArgs) Handles txtName.TextChanged
        If hidFlag.Value = "Add" Then
            txtName.ForeColor = Drawing.Color.Red
        ElseIf hidFlag.Value = "Modify" Then

            If hidOldData.Value <> "" Then
                If hidOldData.Value.Split(";")(0).ToString() <> txtName.Text.Replace("/", COL_DELIM) Then
                    txtName.ForeColor = Drawing.Color.Red
                Else
                    txtName.ForeColor = Drawing.Color.Black
                End If
            End If
        End If
    End Sub

    ''' <summary>
    ''' 順序改變
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub txtSortCtrl_TextChanged(sender As Object, e As EventArgs) Handles txtSortCtrl.TextChanged
        If hidFlag.Value = "Add" Then
            txtSortCtrl.ForeColor = Drawing.Color.Red
        ElseIf hidFlag.Value = "Modify" Then

            If hidOldData.Value <> "" Then
                If hidOldData.Value.Split(";")(2).ToString() <> txtSortCtrl.Text Then
                    txtSortCtrl.ForeColor = Drawing.Color.Red
                Else
                    txtSortCtrl.ForeColor = Drawing.Color.Black
                End If
            End If
        End If
    End Sub

    ''' <summary>
    ''' 屬性改變
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub rdoListHoFlag_SelectedIndexChanged(sender As Object, e As EventArgs) Handles rdoListHoFlag.SelectedIndexChanged

        If hidFlag.Value = "Add" Then

            For Each item As ListItem In rdoListHoFlag.Items
                If item.Selected Then
                    item.Text = String.Format("<font color='red'>{0}</font>", item.Text)
                Else
                    item.Text = item.Text.Replace("<font color='red'>", "").Replace("</font>", "")
                End If
            Next
        ElseIf hidFlag.Value = "Modify" Then
            If hidOldData.Value <> "" Then
                If hidOldData.Value.Split(";")(3).ToString() = "" Then

                Else
                    For Each item As ListItem In rdoListHoFlag.Items
                        If item.Selected Then
                            item.Text = String.Format("<font color='red'>{0}</font>", item.Text)
                        Else
                            item.Text = item.Text.Replace("<font color='red'>", "").Replace("</font>", "")
                        End If
                    Next
                End If
            End If
        End If
    End Sub

    ''' <summary>
    ''' 是否啟用
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub rdoDisable_SelectedIndexChanged(sender As Object, e As EventArgs) Handles rdoDisable.SelectedIndexChanged
        If hidFlag.Value = "Add" Then
            For Each item As ListItem In rdoDisable.Items
                If item.Selected Then
                    item.Text = String.Format("<font color='red'>{0}</font>", item.Text)
                Else
                    item.Text = item.Text.Replace("<font color='red'>", "").Replace("</font>", "")
                End If
            Next
        ElseIf hidFlag.Value = "Modify" Then
            If hidOldData.Value <> "" Then
                For Each item As ListItem In rdoDisable.Items
                    If item.Selected Then
                        item.Text = String.Format("<font color='red'>{0}</font>", item.Text)
                    Else
                        item.Text = item.Text.Replace("<font color='red'>", "").Replace("</font>", "")
                    End If
                Next
            End If
        End If
    End Sub
#End Region

#Region "審核區塊Function"

    ''' <summary>
    ''' 綁定審核資料
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function initCheckData()
        Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager())
        Dim syFunctionCode As New AUTH_OP.SY_FUNCTION_CODE(GetDatabaseManager())
        Dim syFunctionCodeList As New AUTH_OP.SY_FUNCTION_CODEList(GetDatabaseManager())
        Dim apCodelist As New AP_CODEList(GetDatabaseManager())
        Dim syRole As New AUTH_OP.SY_ROLE(GetDatabaseManager())
        Dim xmlData As String = String.Empty
        Dim dtRootData As New DataTable
        Dim dtData As New DataTable
        Dim iCount As Integer = 0

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
                dv.Sort = "funccode desc"
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

                If syFunctionCodeList.loadDataByParentList(hidNodeParent.Value) Then
                    dtDataLive = syFunctionCodeList.getCurrentDataSet.Tables(0)
                End If

                addChildNode(treeViewCheck.Nodes(0), dtDataLive, dtData)
            End If
        End If

        If sParentE.Length > 0 Then ' 若父節點=0 說明是在“安泰銀行”跟節點下新增的資料

            Dim rows() As DataRow = dtData.Select("[parent]='" & sParentE & "'")

            For Each dr As DataRow In rows
                Dim childNode As TreeNode = New TreeNode()
                childNode.Value = dr("FUNCCODE").ToString
                childNode.Text = String.Format("<font color='red'>{0}</font>", dr("FUCNAME").ToString & "  (" & dr("FUNCCODE").ToString & ")")

                If dtData.Select("[parent]='" & childNode.Value & "'").Count > 0 Then
                    addChildNode(childNode, dtData, Nothing)
                End If

                treeViewCheck.Nodes(0).ChildNodes.Add(childNode)
            Next

            If m_sTreeNode Is Nothing Then
                m_sTreeNode = treeViewCheck.Nodes(0).ChildNodes(0)
            End If
        End If

        If Not m_sTreeNode Is Nothing Then
            loadCheck(m_sTreeNode)
        End If
    End Function

    ''' <summary>
    ''' 根據父節點往上推到其父節點
    ''' </summary>
    ''' <param name="sParent"></param>
    ''' <remarks></remarks>
    Sub loadNodeByParent(ByVal sParent As String)
        Dim syFunctionCode As New AUTH_OP.SY_FUNCTION_CODE(GetDatabaseManager())

        For Each Str As String In sParent.Split(",")
            If syFunctionCode.loadByPK(Str) Then
                If syFunctionCode.getAttribute("PARENT") = "0" Then
                    hidNodeParent.Value &= syFunctionCode.getAttribute("FUNCCODE").ToString() & ","
                Else
                    loadNodeByParent(syFunctionCode.getAttribute("PARENT"))
                    hidNodeParent.Value &= syFunctionCode.getAttribute("PARENT").ToString & ","
                End If
            End If
        Next
    End Sub

    ''' <summary>
    ''' 綁定屬性
    ''' </summary>
    ''' <remarks></remarks>
    Sub bindCheck()

        ' 實例化
        Dim apCodelist As New AP_CODEList(GetDatabaseManager())

        ' 綁定屬性
        If apCodelist.loadByUpCode(m_sUpCode2370) Then
            rdoListHoFlagBefore.DataSource = apCodelist.getCurrentDataSet
            rdoListHoFlagBefore.DataTextField = "TEXT"
            rdoListHoFlagBefore.DataValueField = "VALUE"
            rdoListHoFlagBefore.DataBind()

            rdoListHoFlagAfter.DataSource = apCodelist.getCurrentDataSet
            rdoListHoFlagAfter.DataTextField = "TEXT"
            rdoListHoFlagAfter.DataValueField = "VALUE"
            rdoListHoFlagAfter.DataBind()
        End If
    End Sub

    ''' <summary>
    ''' 新增treeViewCheck節點
    ''' </summary>
    ''' <param name="node">節點</param>
    ''' <param name="item">parent</param>
    ''' <remarks></remarks>
    Private Sub addNodeCheck(ByVal node As TreeNode, ByVal item As String, ByVal sFuncId As String)
        Try
            Dim syFunctionList As New AUTH_OP.SY_FUNCTION_CODEList(GetDatabaseManager())
            Dim dtData As New DataTable

            ' 根據FuncCode取得所有的子節點
            If syFunctionList.loadDataByType(item) Then
                dtData = syFunctionList.getCurrentDataSet.Tables(0)
            End If

            ' 循環添加子節點
            If Not dtData Is Nothing Then
                If dtData.Rows.Count > 0 Then
                    For Each dr As DataRow In dtData.Rows

                        Dim subNode As TreeNode = New TreeNode
                        subNode.Text = dr("FUCNAME").ToString() & "  (" & dr("FUNCCODE").ToString & ")"
                        subNode.Value = dr("FUNCCODE").ToString
                        If subNode.Value <> sFuncId Then
                            addNodeCheck(subNode, dr("FUNCCODE").ToString(), sFuncId)
                            node.ChildNodes.Add(subNode)
                            subNode.SelectAction = TreeNodeSelectAction.None
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 驗證節點是否存在,如果沒有 將其新增到TreeView中
    ''' </summary>
    ''' <param name="node">節點</param>
    ''' <param name="sValue">角色編號</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function checkExistNodeCheck(ByVal node As TreeNode, ByVal subNode As TreeNode, ByVal sParent As String) As String
        If node.ChildNodes.Count > 0 Then
            For i As Integer = 0 To node.ChildNodes.Count - 1

                If i < node.ChildNodes.Count Then
                    If node.ChildNodes(i).Value = sParent Then
                        node.ChildNodes(i).ChildNodes.Add(subNode)
                    End If

                    checkExistNodeCheck(node.ChildNodes(i), subNode, sParent)
                End If
            Next
        End If
    End Function

    ''' <summary>
    ''' 審核區塊更新資料加載
    ''' </summary>
    ''' <remarks></remarks>
    Sub loadCheck(ByVal trNode As TreeNode)
        '------------------------------ 審核區域，更新操作，資料加載--------------------------------------------------- 
        operationXML("9", "", trNode)
    End Sub

    ''' <summary>
    '''  判斷DB中是否有此值 若有 將其設置為紅色
    ''' </summary>
    ''' <param name="node"></param>
    ''' <param name="sRoleId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function checkNodeCheck(ByVal node As TreeNode, ByVal sRoleId As String)
        For Each item As TreeNode In node.ChildNodes
            If item.Value = sRoleId Then
                item.SelectAction = TreeNodeSelectAction.None
            End If

            checkNodeCheck(item, sRoleId)
        Next
    End Function
#End Region

#Region "審核區塊Event"

    ''' <summary>
    ''' 同意OK
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnArgee_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnArgee.Click

        Dim stepInfo As FLOW_OP.StepInfo
        ' 呼叫流程方法
        Try
            GetDatabaseManager().beginTran()
            stepInfo = MBSC.UICtl.UIShareFun.pushFlow(GetDatabaseManager(), m_sCaseId)

            updatTemp2Live(GetDatabaseManager(), stepInfo, "Y")
            ' 關閉視窗
            com.Azion.EloanUtility.UIUtility.closeWindow()
            GetDatabaseManager().commit()
        Catch ex As Exception
            GetDatabaseManager().Rollback()
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 不同意 OK
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnNotArgee_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnNotArgee.Click
        Dim stepInfo As FLOW_OP.StepInfo

        ' 呼叫流程方法
        Try
            GetDatabaseManager().beginTran()
            stepInfo = MBSC.UICtl.UIShareFun.pushFlow(GetDatabaseManager(), m_sCaseId)

            updatTemp2Live(GetDatabaseManager(), stepInfo, "N")

            ' 關閉視窗
            com.Azion.EloanUtility.UIUtility.closeWindow()
            GetDatabaseManager().commit()
        Catch ex As Exception
            GetDatabaseManager().Rollback()
            SYUIBase.showErrMsg(Me, ex)
        End Try

    End Sub

    ''' <summary>
    ''' 修正補充 OK
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnReviseFlow_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnReviseFlow.Click
        Dim stepInfo As FLOW_OP.StepInfo

        Try
            GetDatabaseManager().beginTran()
            ' 呼叫流程方法
            stepInfo = MBSC.UICtl.UIShareFun.rollBack(GetDatabaseManager(), m_sCaseId)

            updatTemp2Live(GetDatabaseManager(), stepInfo, "")

            ' 關閉視窗
            ' com.Azion.EloanUtility.UIUtility.closeWindow()
            GetDatabaseManager().commit()

        Catch ex As Exception
            GetDatabaseManager().Rollback()
            SYUIBase.showErrMsg(Me, ex)
        End Try

    End Sub

    ''' <summary>
    ''' 更新
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub treeViewCheck_SelectedNodeChanged(ByVal sender As Object, ByVal e As EventArgs) Handles treeViewCheck.SelectedNodeChanged
        Try
            If treeViewCheck.SelectedNode.Value.Split(";")(0) <> "0" Then
                loadCheck(treeViewCheck.SelectedNode)
            End If
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub
#End Region


End Class