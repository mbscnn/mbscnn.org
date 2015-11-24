''' <summary>
''' 程式說明：交易分派
''' 建立者：Avril
''' 建立日期：2012-04-27
''' </summary>
Imports com.Azion.NET.VB
Imports AUTH_OP
Imports AUTH_OP.TABLE
Imports MBSC.MB_OP
Imports System.Xml
Imports System.IO
Imports com.Azion.EloanUtility

Public Class SY_UPDFUNCTIONGRANTROLE
    Inherits SYUIBase


    Dim m_sFuncCodeList As String = String.Empty ' 記錄交易編號的集合
    Dim m_sCodeFlag As String = String.Empty ' 記錄臨時檔中每筆資料的標識(N/I/D)
    Dim m_sOldCodeList As String = String.Empty ' 記錄實際表中的交易編號
    Dim m_sXmlData As String = String.Empty ' 記錄臨時檔中的XMLDATA
    Dim m_sbXmlData As New StringBuilder ' 組織要存儲到臨時檔中的XMLDATA
    Dim m_sRoleIdList As String = String.Empty ' 記錄角色編號的集合
    Dim m_sSysIdList As String = String.Empty ' 記錄系統編號的集合
    Dim m_sSubSysIdList As String = String.Empty ' 記錄子系統編號的集合
    Dim m_sInfoData As String = String.Empty ' 記錄角色編號*系統編號*子系統編號
    Dim m_sUpCode2366 As String = String.Empty ' 安泰銀行的UpCode
    Dim m_sParentRole As String = String.Empty

    Dim m_sTreeNode As TreeNode
    Dim m_sFlag As String = String.Empty

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

        m_bCheck = False

        ' 頁面第一次加載
        If Not IsPostBack Then

            ' 資料加載
            bindData()

            ' 添加Onclick事件
            treeViewFun.Attributes.Add("onclick", "postBackByObject();")
        End If
    End Sub

    Protected Sub Page_Error(sender As Object, e As EventArgs)


        Dim ex As Exception = Server.GetLastError()

        If TypeOf HttpContext.Current.Server.GetLastError() Is HttpRequestValidationException Then


            HttpContext.Current.Response.Write("请输入合法的字符串【<a href=""javascript:history.back(0);"">返回</a>】")


            HttpContext.Current.Server.ClearError()
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
        '    m_sWorkingTopDepNo = "1"
        '    m_sLoginUserid = "S000035"
        '    m_sLoginTopDepNo = "1"
        '    m_sFunccode = "1"
        '    m_sStepNo = ""
        'End If

        If m_bTesting OrElse Request("TESTMODE") = "1" Then
            m_bTesting = True

            Dim loginUser As New SY_LOGIN(GetDatabaseManager)
            loginUser.getUserInfo("924", "S000914")
            'm_bCheck = False
            Session("HOFLAG") = "1"
            'm_bDisplayMode = False
            MyBase.InitParas()
        End If


        ' 案號
        lblCaseId.Text = m_sCaseId

        ' 單位
        If syBranch.loadByCaseId(m_sCaseId) Then
            lblBranch.Text = syBranch.getAttribute("BRCNAME")
        End If

        m_sUpCode2366 = com.Azion.EloanUtility.CodeList.getAppSettings("UpCode2366")

    End Function

    ''' <summary>
    ''' 綁定資料
    ''' </summary>
    ''' <remarks></remarks>
    Sub bindData()
        Try

            'Dim bHas400 As Boolean = False

            'If FLOW_OP.FlowFacade.getNewInstance(GetDatabaseManager).getSYFlowStep.GetCount( _
            '    "CASEID", m_sCaseId,
            '    "STEP_NO", "SY000400") > 0 Then
            '    bHas400 = True
            'End If


            If m_sFourStepNo = "" OrElse m_sFourStepNo = "0300" OrElse m_bDisplayMode Then

                ' 正常區域資料加載
                initNormalData()

                treeViewRole.ExpandDepth = m_sTreeLevel

                divCheck.Visible = False
                divNormal.Visible = True
            Else

                ' 審核模塊資料加載
                initCheckData()

                divCheck.Visible = True
                divNormal.Visible = False

                ' 顯示案號與單位
                divBranch.Visible = True
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
        Dim syFlowStep As New AUTH_OP.SY_FLOWSTEP(GetDatabaseManager())
        Dim sySubSysId As New AUTH_OP.SY_SUBSYSIDList(GetDatabaseManager())
        Dim xmlDocument As New XmlDocument
        Dim sStepNo As String = String.Empty
        Dim sSubFlowSeq As String = String.Empty
        Dim sSubFlowCount As String = String.Empty
        Dim sCaseId As String = String.Empty
        Dim sTempData As String = "2,4,8"
        Dim syAllSubSysId As String = String.Empty

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

        If sySubSysId.loadAllData() Then
            For Each item As SY_SUBSYSID In sySubSysId
                syAllSubSysId &= item.getAttribute("SUBSYSID").ToString() & ","
            Next
        End If

        If sFlag = "1" Then

            ' 根據案號查詢資料
            If syFlowStep.loadByCaseId(sCaseId) Then
                sStepNo = syFlowStep.getAttribute("STEP_NO")
                sSubFlowCount = syFlowStep.getAttribute("SUBFLOW_COUNT")
                sSubFlowSeq = syFlowStep.getAttribute("SUBFLOW_SEQ")
            End If
        End If

        If Not m_sXmlData Is Nothing Then

            ' document對象載入XML文件
            If m_sXmlData.Length > 0 Then
                xmlDocument.LoadXml(m_sXmlData)
            End If
        End If

        For i As Integer = 0 To xmlDocument.GetElementsByTagName("SY_REL_ROLE_FUNCTION").Count - 1

            Dim syRelRoleFunctionHis As New SY_REL_ROLE_FUNCTION_HIS(GetDatabaseManager())
            Dim sRoleId As String = DirectCast(xmlDocument.GetElementsByTagName("SY_REL_ROLE_FUNCTION")(i)("ROLEID"), System.Xml.XmlElement).InnerText
            Dim sFuncCode As String = DirectCast(xmlDocument.GetElementsByTagName("SY_REL_ROLE_FUNCTION")(i)("FUNCCODE"), System.Xml.XmlElement).InnerText
            Dim sSysId As String = DirectCast(xmlDocument.GetElementsByTagName("SY_REL_ROLE_FUNCTION")(i)("SYSID"), System.Xml.XmlElement).InnerText
            Dim sSubSysId As String = DirectCast(xmlDocument.GetElementsByTagName("SY_REL_ROLE_FUNCTION")(i)("SUBSYSID"), System.Xml.XmlElement).InnerText

            Select Case sFlag
                Case "1"
                    '------------------------------ 新增歷史檔部份---------------------------------------------------
                    If Not syRelRoleFunctionHis.loadByPK(sFuncCode, sSubSysId, sSysId, sCaseId, sStepNo, sSubFlowSeq, sSubFlowCount, sRoleId) Then
                        syRelRoleFunctionHis.setAttribute("FUNCCODE", sFuncCode)
                        syRelRoleFunctionHis.setAttribute("SUBSYSID", sSubSysId)
                        syRelRoleFunctionHis.setAttribute("SYSID", sSysId)
                        syRelRoleFunctionHis.setAttribute("CASEID", sCaseId)
                        syRelRoleFunctionHis.setAttribute("STEP_NO", sStepNo)
                        syRelRoleFunctionHis.setAttribute("SUBFLOW_SEQ", sSubFlowSeq)
                        syRelRoleFunctionHis.setAttribute("SUBFLOW_COUNT", sSubFlowCount)
                        syRelRoleFunctionHis.setAttribute("ROLEID", sRoleId)
                    End If

                    If Not String.IsNullOrEmpty(sApproved) Then
                        syRelRoleFunctionHis.setAttribute("APPROVED", sApproved)
                    End If

                    syRelRoleFunctionHis.setAttribute("XMLDATA", m_sXmlData)
                    syRelRoleFunctionHis.save()
                Case "2", "4", "8"

                    '------------------------------ 正常區域加載資料,查詢xml中的資料，將其新增到左邊目錄樹---------------------------------------------------
                    If DirectCast(xmlDocument.GetElementsByTagName("SY_REL_ROLE_FUNCTION")(i)("ROLEID"), System.Xml.XmlElement).InnerText = hidRoleId.Value.Split("*")(0) Then
                        m_sFuncCodeList &= DirectCast(xmlDocument.GetElementsByTagName("SY_REL_ROLE_FUNCTION")(i)("FUNCCODE"), System.Xml.XmlElement).InnerText & ","
                        m_sCodeFlag &= DirectCast(xmlDocument.GetElementsByTagName("SY_REL_ROLE_FUNCTION")(i)("FUNCCODE"), System.Xml.XmlElement).InnerText & "," & DirectCast(xmlDocument.GetElementsByTagName("SY_REL_ROLE_FUNCTION")(i)("OPERATION"), System.Xml.XmlElement).InnerText & "|"
                    End If
                Case "3"
                    '------------------------------ 正常區域加載資料,儲存按鈕---------------------------------------------------
                    ' 查找出所有RoleId一樣的資料,刪除掉
                    If DirectCast(xmlDocument.GetElementsByTagName("SY_REL_ROLE_FUNCTION")(i)("ROLEID"), System.Xml.XmlElement).InnerText = hidRoleId.Value.Split("*")(0) Then
                        m_sOldCodeList &= DirectCast(xmlDocument.GetElementsByTagName("SY_REL_ROLE_FUNCTION")(i)("FUNCCODE"), System.Xml.XmlElement).InnerText & ";" & DirectCast(xmlDocument.GetElementsByTagName("SY_REL_ROLE_FUNCTION")(i)("OPERATION"), System.Xml.XmlElement).InnerText & ","
                        xmlDocument.GetElementsByTagName("SY_REL_ROLE_FUNCTION")(i).RemoveAll()
                    Else
                        m_sbXmlData.Append("<SY_REL_ROLE_FUNCTION>" & xmlDocument.GetElementsByTagName("SY_REL_ROLE_FUNCTION")(i).InnerXml & "</SY_REL_ROLE_FUNCTION>")
                    End If
                Case "5"
                    '------------------------------ 正常區域,確認送出操作，查找所有的角色編號---------------------------------------------------
                    m_sRoleIdList += "'" & DirectCast(xmlDocument.GetElementsByTagName("SY_REL_ROLE_FUNCTION")(i)("ROLEID"), System.Xml.XmlElement).InnerText & "'" & ","
                Case "6"
                    '------------------------------ 審核區域,資料加載部份---------------------------------------------------
                    m_sRoleIdList &= DirectCast(xmlDocument.GetElementsByTagName("SY_REL_ROLE_FUNCTION")(i)("ROLEID"), System.Xml.XmlElement).InnerText & ","
                    m_sSysIdList &= "'" & DirectCast(xmlDocument.GetElementsByTagName("SY_REL_ROLE_FUNCTION")(i)("SYSID"), System.Xml.XmlElement).InnerText & "'" & ","
                    m_sSubSysIdList &= "'" & DirectCast(xmlDocument.GetElementsByTagName("SY_REL_ROLE_FUNCTION")(i)("SUBSYSID"), System.Xml.XmlElement).InnerText & "'" & ","
                Case "7"
                    '------------------------------ 審核區域,同意按鈕---------------------------------------------------
                    Dim syRelRoleFunction As New AUTH_OP.SY_REL_ROLE_FUNCTION(GetDatabaseManager())


                    If syAllSubSysId.Contains(sSubSysId) Then

                        If sFuncCode <> "0" Then
                            ' 判斷DB中是否有資料，若沒有，設置主鍵
                            If Not syRelRoleFunction.loadByPK(sRoleId, sSubSysId, sSysId, sFuncCode) Then
                                syRelRoleFunction.setAttribute("ROLEID", sRoleId)
                                syRelRoleFunction.setAttribute("SUBSYSID", sSubSysId)
                                syRelRoleFunction.setAttribute("SYSID", sSysId)
                                syRelRoleFunction.setAttribute("FUNCCODE", sFuncCode)
                                syRelRoleFunction.save()
                            End If
                        End If

                        If String.IsNullOrEmpty(m_sInfoData) Then
                            m_sInfoData = sRoleId & "*" & sSysId & "*" & sSubSysId
                        End If
                    End If
            End Select
        Next
    End Function

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

                ' 如果是Insert資料，則需要進行追加 如果是修改 則修改
                oldXmlData = syTempInfo.getAttribute("TEMPDATA").ToString()
                sCaseid = syTempInfo.getString("CASEID")
            End If
        Else
            If syTempInfo.loadByCaseId(m_sCaseId) Then
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
    ''' 新增資料到真實檔及歷史檔
    ''' </summary>
    ''' <param name="databaseManager"></param>
    ''' <param name="stepinfo"></param>
    ''' <remarks></remarks>
    Sub updatTemp2Live(ByVal databaseManager As DatabaseManager, ByRef stepinfo As FLOW_OP.StepInfo, ByVal sAPPROVED As String)
        Dim dt As DataTable = loadSYTempInfo()
        Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(databaseManager)
        Dim syRoleList As New AUTH_OP.SY_ROLEList(databaseManager)
        Dim syRelRoleFunction As New AUTH_OP.SY_REL_ROLE_FUNCTION(databaseManager)
        Dim syRelRoleFunctionList As New AUTH_OP.SY_REL_ROLE_FUNCTIONList(databaseManager)

        For Each row As DataRow In dt.Rows
            Dim iRoleid As Integer = CInt(row("ROLEID"))

            '取得原本SY_RoleSuitSys的資料
            '整理沒有在Temp裡的角色
            Dim sRoleOperation As String = String.Empty
            Dim sSysOperation As String = String.Empty
            Dim sRoleList As String = String.Empty

            If row("FUNCCODE").ToString() <> "0" Then

                syRelRoleFunction.clear()

                If row("FUNCCODE").ToString() <> "" Then
                    If Not syRelRoleFunction.loadByPK(CInt(row("ROLEID")), CStr(row("SUBSYSID")), CStr(row("SYSID")), CInt(row("FUNCCODE"))) Then
                        syRelRoleFunction.setInt("ROLEID", iRoleid)
                        syRelRoleFunction.setString("SUBSYSID", CStr(row("SUBSYSID")))
                        syRelRoleFunction.setString("SYSID", CStr(row("SYSID")))
                        syRelRoleFunction.setString("FUNCCODE", CInt(row("FUNCCODE")))
                    End If

                    If sAPPROVED = "Y" Then
                        If row("OPERATION").ToString() <> "D" Then
                            syRelRoleFunction.save()
                        End If
                    End If

                    ' 若標識為刪除，則需要刪除live檔的資料
                    If sAPPROVED = "Y" Then
                        If row("OPERATION").ToString = "D" Then
                            If syRelRoleFunction.loadByPK(CInt(row("ROLEID")), CStr(row("SUBSYSID")), CStr(row("SYSID")), CInt(row("FUNCCODE"))) Then
                                syRelRoleFunction.remove()
                            End If
                        End If
                    End If

                    '6.紀錄新增(I)或修改(U) his(SY_ROLE)
                    insertSyRelRoleFunctionHis(databaseManager, row, row("OPERATION").ToString, "Y", sAPPROVED, _
                                           stepinfo.currentStepInfo.caseId, _
                                           stepinfo.currentStepInfo.stepNo, _
                                           stepinfo.currentStepInfo.subflowSeq, _
                                           stepinfo.currentStepInfo.subflowCount)

                    '查詢勾選節點的子節點的集合
                    If syRoleList.loadDataByType(row("ROLEID")) Then
                        For Each item As AUTH_OP.SY_ROLE In syRoleList
                            sRoleList &= item.getAttribute("ROLEID") & ","
                        Next
                    End If

                    If sRoleList.Length > 0 Then
                        sRoleList = sRoleList.Substring(0, sRoleList.Length - 1)

                        ' 若將父節點角色的交易取消勾選，對應子節點的交易也需要在DB中刪除
                        If row("OPERATION").ToString() = "D" Then

                            ' 刪除live檔
                            If syRelRoleFunctionList.loadAddDataByCon(sRoleList, row("SYSID"), row("SUBSYSID")) Then
                                For Each item As SY_REL_ROLE_FUNCTION In syRelRoleFunctionList
                                    If item.getAttribute("FUNCCODE").ToString = row("FUNCCODE") Then

                                        ' 新增一筆歷史檔資料
                                        insertSyRelRoleFunctionHis2(databaseManager, item, row("OPERATION").ToString, "N", sAPPROVED)

                                        If sAPPROVED = "Y" Then
                                            item.remove()
                                        End If
                                    End If
                                Next
                            End If
                        End If
                    End If
                End If
            End If
        Next

        If sAPPROVED = "Y" Or sAPPROVED = "N" Then
            If m_sFourStepNo = "" Then
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
    Sub insertSyRelRoleFunctionHis(ByVal databaseManager As DatabaseManager, ByVal row As DataRow, ByVal sOperation As String, ByRef sFlag As String, Optional sAPPROVED As String = "Y", _
                  Optional ByVal sCaseId As String = "SY0000000000000", Optional ByVal sStepNo As String = "00000000", _
                  Optional ByVal iSubFlowSeq As Integer = 0, Optional ByVal iSubFlowCount As Integer = 0)

        Dim syRelRoleFunctionHis As New AUTH_OP.SY_REL_ROLE_FUNCTION_HIS(databaseManager)

        syRelRoleFunctionHis.loadByPK(row("FUNCCODE"), row("SUBSYSID"), row("SYSID"), sCaseId, sStepNo, iSubFlowSeq, iSubFlowCount, row("ROLEID"))

        '因為欄位相同
        For Each column As DataColumn In row.Table.Columns
            syRelRoleFunctionHis.setAttribute(column.ColumnName.ToUpper, row(column.ColumnName))
        Next

        syRelRoleFunctionHis.setAttribute("CASEID", sCaseId)
        syRelRoleFunctionHis.setAttribute("STEP_NO", sStepNo)
        syRelRoleFunctionHis.setAttribute("SUBFLOW_SEQ", iSubFlowSeq)
        syRelRoleFunctionHis.setAttribute("SUBFLOW_COUNT", iSubFlowCount)
        syRelRoleFunctionHis.setAttribute("APPROVED", sAPPROVED)
        syRelRoleFunctionHis.setAttribute("OPERATION", sOperation)
        syRelRoleFunctionHis.save()
    End Sub

    Sub insertSyRelRoleFunctionHis2(ByVal databaseManager As DatabaseManager, ByVal item As AUTH_OP.SY_REL_ROLE_FUNCTION, ByVal sOperation As String, ByRef sFlag As String, Optional sAPPROVED As String = "Y", _
                  Optional ByVal sCaseId As String = "SY0000000000000", Optional ByVal sStepNo As String = "00000000", _
                  Optional ByVal iSubFlowSeq As Integer = 0, Optional ByVal iSubFlowCount As Integer = 0)
        Dim syRelRoleFunctionHis As New AUTH_OP.SY_REL_ROLE_FUNCTION_HIS(databaseManager)
        Dim iTempSubFlowSeq As Integer = 0

        If sFlag = "N" Then '不走流程
            iTempSubFlowSeq = syRelRoleFunctionHis.getMaxSubFlowSeq(item.getAttribute("FUNCCODE"), item.getAttribute("SUBSYSID"), item.getAttribute("SYSID"), sCaseId, sStepNo, iSubFlowCount, item.getAttribute("ROLEID").ToString()) + 1
        Else
            iTempSubFlowSeq = iSubFlowSeq
        End If

        '因為欄位相同
        syRelRoleFunctionHis.setAttribute("ROLEID", item.getAttribute("ROLEID"))
        syRelRoleFunctionHis.setAttribute("SUBSYSID", item.getAttribute("SUBSYSID"))
        syRelRoleFunctionHis.setAttribute("SYSID", item.getAttribute("SYSID"))
        syRelRoleFunctionHis.setAttribute("FUNCCODE", item.getAttribute("FUNCCODE"))

        syRelRoleFunctionHis.setAttribute("CASEID", sCaseId)
        syRelRoleFunctionHis.setAttribute("STEP_NO", sStepNo)
        syRelRoleFunctionHis.setAttribute("SUBFLOW_SEQ", iTempSubFlowSeq)
        syRelRoleFunctionHis.setAttribute("SUBFLOW_COUNT", iSubFlowCount)
        syRelRoleFunctionHis.setAttribute("APPROVED", sAPPROVED)
        syRelRoleFunctionHis.setAttribute("OPERATION", sOperation)
        syRelRoleFunctionHis.save()
    End Sub

    ''' <summary>
    ''' 檢查父節點的顏色部份
    ''' </summary>
    ''' <param name="node"></param>
    ''' <param name="checkFlag"></param>
    ''' <remarks></remarks>
    Sub checkParentNode(ByVal node As TreeNode, ByVal checkFlag As Boolean, ByVal dtData As DataTable)
        Dim bTemp As Boolean = checkFlag

        If Not node.Parent Is Nothing Then

            ' 只有有一個勾選，則父節點勾選，若沒有勾選的子項，則父節點取消勾選
            If node.Parent.ChildNodes.Count > 0 Then
                For Each item As TreeNode In node.Parent.ChildNodes
                    If item.Checked = True Then
                        checkFlag = True
                    End If
                Next
            End If

            node.Parent.Checked = checkFlag

            If checkFlag = True Then
                If Not node.Parent.Text.Contains("red") Then
                    node.Parent.Text = String.Format("<font color='red'>{0}</font>", node.Parent.Text)
                End If

                If bTemp = False Then
                    node.Parent.Text = node.Parent.Text.Replace("<font color='red'>", "").Replace("</font>", "")
                End If
            Else

                If dtData.Rows.Count > 0 Then

                    Dim drRowTemp() As DataRow = dtData.Select(String.Format("ROLEID ='{0}' and SYSID='{1}' and SUBSYSID='{2}' AND FUNCCODE = '{3}'", treeViewRole.SelectedNode.Value.Split("*")(0), treeViewRole.SelectedNode.Value.Split("*")(1), treeViewRole.SelectedNode.Value.Split("*")(2), node.Value.Split("$")(0)))
                    If drRowTemp.Count > 0 Then
                        For Each dr As DataRow In drRowTemp
                            If dr("OPERATION").ToString() <> "I" Then

                                node.Parent.Text = String.Format("<font color='red'>{0}</font>", node.Parent.Text)

                                ' 記錄刪除的資料
                                hidDelData.Value &= node.Parent.Value.Split("$")(0) & ";"
                            End If
                        Next
                    End If
                Else
                    ' 記錄刪除的資料
                    hidDelData.Value &= node.Parent.Value.Split("$")(0) & ";"
                End If
            End If

            If Not node.Parent.Parent Is Nothing Then
                checkParentNode(node.Parent, checkFlag, dtData)
            End If
        End If
    End Sub
#End Region

#Region "正常資料Function"

    ''' <summary>
    ''' 資料加載
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function initNormalData()
        Try
            ' 實例化
            Dim apCodelist As New AP_CODEList(GetDatabaseManager())
            Dim sySysIdList As New AUTH_OP.SY_SYSIDList(GetDatabaseManager())
            Dim dtRootData As New DataTable
            Dim dtParentData As New DataTable

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

                    rootNode.SelectAction = TreeNodeSelectAction.None
                    treeViewRole.Nodes.Add(rootNode)
                Next
            End If

            ' (2)查詢角色資料的第一層資料：parent=0
            If sySysIdList.loadUseData() Then
                dtParentData = sySysIdList.getCurrentDataSet.Tables(0)
            End If

            ' (3)循環 新增到treeViewRole中
            If Not dtParentData Is Nothing Then
                For Each dr As DataRow In dtParentData.Rows

                    ' 父節點
                    Dim parentNode As TreeNode = New TreeNode()
                    parentNode.Value = dr("SYSID").ToString()
                    parentNode.Text = dr("SYSNAME").ToString & "  (" & dr("SYSID").ToString & ")"
                    parentNode.SelectAction = TreeNodeSelectAction.None

                    ' (4)添加子節點
                    addNode(parentNode, dr("SYSID").ToString())

                    ' 添加到treeViewRole中
                    treeViewRole.Nodes(0).ChildNodes.Add(parentNode)
                Next
            End If

            '#0000FF
            treeViewRole.RootNodeStyle.ForeColor = Drawing.Color.Blue
            treeViewRole.ParentNodeStyle.ForeColor = Drawing.Color.Blue
            treeViewRole.LeafNodeStyle.ForeColor = Drawing.Color.Blue
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Function

    ''' <summary>
    ''' 新增treeViewRole節點
    ''' </summary>
    ''' <param name="node">節點</param>
    ''' <param name="item">parent</param>
    ''' <remarks></remarks>
    Private Sub addNode(ByVal node As TreeNode, ByVal item As String)
        Try
            Dim sySubSysIdList As New AUTH_OP.SY_SUBSYSIDList(GetDatabaseManager())
            Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager())
            Dim dtData As New DataTable
            Dim sNewTempTable As New DataTable
            Dim syRoleList As New SY_ROLEList(GetDatabaseManager())
            Dim sTempTable As New DataTable

            ' 根據SYSID取得所有的子系統
            If sySubSysIdList.loadSubSys(item) Then
                dtData = sySubSysIdList.getCurrentDataSet.Tables(0)
            End If

            ' 循環添加子節點
            If Not dtData Is Nothing Then
                If dtData.Rows.Count > 0 Then
                    For Each dr As DataRow In dtData.Rows

                        Dim subNode As TreeNode = New TreeNode
                        subNode.Text = dr("SUBSYSNAME").ToString() & "  (" & dr("SUBSYSID").ToString & ")"
                        subNode.Value = dr("SUBSYSID").ToString
                        subNode.SelectAction = TreeNodeSelectAction.None

                        ' 根據登入者，部門編號，系統編號取得資料
                        sNewTempTable = syRoleList.genFunList(m_sWorkingUserid, subNode.Value, item)

                        If Not sNewTempTable Is Nothing Then
                            'dtTable = sNewTempTable.DefaultView.ToTable(True)

                            If Not sNewTempTable Is Nothing Then
                                If sNewTempTable.Rows.Count > 0 Then

                                    ' 若有資料，將其新增到子系統下面
                                    addNodes(subNode, 0, 0, sNewTempTable)
                                End If
                            End If
                        End If

                        node.ChildNodes.Add(subNode)
                    Next
                End If
            End If
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 新增子系統下的角色
    ''' </summary>
    ''' <param name="tNode">節點</param>
    ''' <param name="PId">ParentID</param>
    ''' <param name="Level">Leave</param>
    ''' <param name="dt">數據源</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function addNodes(ByRef subNode As TreeNode, ByVal PId As Integer, ByVal Level As Integer, ByVal dt As DataTable) As String

        '定義DataRow承接DataTable篩選的結果
        Dim rows() As DataRow
        Dim sRoleList As String = String.Empty

        ' 查詢臨時檔中的資料
        Dim dtTemp As DataTable = loadSYTempInfo()
        Dim sTempRoleList As String = String.Empty

        ' 查詢此節點是否存在于臨時檔中 若存在 需要變紅
        If Not dtTemp Is Nothing Then
            For Each dr As DataRow In dtTemp.Rows
                sTempRoleList &= dr("ROLEID").ToString & ","
            Next
        End If

        '定義篩選的條件
        Dim filterExpr As String
        filterExpr = "ParentId = " & PId

        '資料篩選並把結果傳入Rows
        rows = dt.Select(filterExpr)

        For Each item As TreeNode In subNode.ChildNodes
            sRoleList &= item.Value & ","
        Next

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
                childNode.Text = row("ROLENAME").ToString & "  (" & row("ROLEID").ToString & ")"
                childNode.ToolTip = childNode.Text
                childNode.Value = row("ROLEID").ToString() & "*" & row("SYSID").ToString() & "*" & row("SUBSYSID").ToString()

                ' 證實 此角色存在于真實檔中 需要變紅
                If sTempRoleList.Contains(row("ROLEID").ToString & ",") Then
                    childNode.Text = String.Format("<font color='red'>{0}</font>", childNode.Text)
                End If

                '將節點加入Tree中
                If PId = "0" Then
                    If row("SUBSYSID").ToString().Equals(subNode.Value.Split("*")(0)) Then
                        subNode.ChildNodes.Add(childNode)
                    End If
                ElseIf Convert.ToInt32(PId) > 0 Then
                    If row("ParentId").ToString().Equals(subNode.Value.Split("*")(0)) Then
                        If Not subNode.Parent Is Nothing Then
                            If subNode.Parent.Value.IndexOf("*") < 0 Then
                                If row("SUBSYSID").ToString().Equals(subNode.Parent.Value) Then
                                    subNode.ChildNodes.Add(childNode)
                                End If
                            Else
                                If row("SUBSYSID").ToString().Equals(subNode.Parent.Value.Split("*")(2)) Then
                                    subNode.ChildNodes.Add(childNode)
                                End If
                            End If
                        End If
                    End If
                End If

                '呼叫遞回取得子節點
                rc = addNodes(childNode, row("ROLEID"), row("Level"), dt)
            Next
        End If
    End Function

    ''' <summary>
    ''' 綁定交易樹
    ''' </summary>
    ''' <param name="treeViewFun"></param>
    ''' <remarks></remarks>
    Sub initFunctionTreeView(ByVal treeViewFun As TreeView)
        Try
            ' 實例化
            Dim apCodelist As New AP_CODEList(GetDatabaseManager())
            Dim syFunctionCodeList As New AUTH_OP.SY_FUNCTION_CODEList(GetDatabaseManager())

            ' root節點
            Dim rootNode As TreeNode = New TreeNode()

            ' (1)綁定root節點
            If apCodelist.loadDataByCodeId(m_sUpCode2366) Then
                For Each dr As DataRow In apCodelist.getCurrentDataSet.Tables(0).Rows 'must be a record in ap_code
                    rootNode.Value = dr("VALUE").ToString() '& ";0;;0;0;0;0"
                    rootNode.Text = dr("TEXT").ToString
                    rootNode.ShowCheckBox = False
                Next
            End If

            ' (2)查詢交易建立的資料的第一層資料
            If Not syFunctionCodeList.genAllFunctionNonHoFlag0() Is Nothing Then

                '將syRoleList存入ViewState
                ViewState("SY_FunctionCodeList") = syFunctionCodeList.getCurrentDataSet.Tables(0)

                ' (4)添加子節點
                addNodes(rootNode, 0, ViewState("SY_FunctionCodeList"))
            End If

            '添加到treeView中()
            treeViewFun.Nodes.Add(rootNode)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ''' <summary>
    ''' 添加子節點
    ''' </summary>
    ''' <param name="tNode"></param>
    ''' <param name="PId"></param>
    ''' <param name="dt"></param>
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

                '子節點的Value組成=FUNCCODE;FUCNAME;FUCURL;PARENT;DISABLED;SORTCTRL;HOFLAG
                childNode.Value = row("FUNCCODE").ToString() & "$" & row("FUCURL").ToString()

                '將節點加入Tree中
                tNode.ChildNodes.Add(childNode)

                If Not childNode.Parent Is Nothing Then
                    If childNode.Parent.Depth = 0 Then
                        childNode.ImageUrl = com.Azion.EloanUtility.UIUtility.getImgPath & "folder.png"
                    End If
                End If

                '呼叫遞回取得子節點
                rc = addNodes(childNode, row("FUNCCODE"), dt)

                If m_sFlag = "" Then
                    For Each node As TreeNode In childNode.ChildNodes
                        node.ShowCheckBox = False
                        node.Parent.ShowCheckBox = False
                    Next
                End If

                If childNode.Parent.Depth > 0 AndAlso childNode.ChildNodes.Count = 0 Then
                ElseIf childNode.Parent.Depth > 0 Then
                    childNode.ImageUrl = com.Azion.EloanUtility.UIUtility.getImgPath & "folder.png"
                End If

                ' 類似于“海外分行”這樣的資料 如果FUCURL為空 並且沒有子節點 則不需要顯示
                If childNode.Parent.Depth = 0 AndAlso childNode.ChildNodes.Count = 0 Then
                    tNode.ChildNodes.Remove(childNode)
                    'ElseIf childNode.Parent.Depth > 0 AndAlso childNode.ChildNodes.Count = 0 Then
                    '    If row("Level").ToString = "1" Then
                    '        tNode.ChildNodes.Remove(childNode)
                    '    End If
                End If
            Next
        End If
    End Function

    ''' <summary>
    ''' 初始化checkbox的狀態
    ''' </summary>
    ''' <remarks></remarks>
    Sub initCheckBox()

        ' 第一層是安泰銀行
        For Each node As TreeNode In treeViewFun.Nodes
            initCheckBox(node)
        Next
    End Sub

    ''' <summary>
    ''' 設置節點的chebox為display
    ''' </summary>
    ''' <param name="node"></param>
    ''' <remarks></remarks>
    Sub initCheckBox(ByVal node As TreeNode)

        ' 交易父節點
        For Each item As TreeNode In node.ChildNodes

            initChildCheckBox(item)

            If m_sFlag <> "" Then
                item.ShowCheckBox = True

                If item.ChildNodes.Count > 0 Then
                    initCheckBox(item)
                Else
                    If item.Value.Split("$")(1) <> "" Then
                        item.ShowCheckBox = True
                    End If
                End If

                m_sFlag = ""

            End If
        Next
    End Sub

    ''' <summary>
    ''' 檢測最後一個節點是否為沒有路徑的節點
    ''' </summary>
    ''' <param name="node"></param>
    ''' <remarks></remarks>
    Sub initChildCheckBox(ByVal node As TreeNode)
        If node.ChildNodes.Count > 0 Then

            For Each item As TreeNode In node.ChildNodes

                If item.ChildNodes.Count > 0 Then
                    initChildCheckBox(item)
                Else
                    If item.Value.Split("$")(1) <> "" Then
                        item.ShowCheckBox = True
                        m_sFlag = "true"
                    Else
                        item.ShowCheckBox = False
                    End If
                End If
            Next
        End If
    End Sub

    ''' <summary>
    ''' 公用方法 新增歷史檔
    ''' </summary>
    ''' <param name="sApproved"></param>
    ''' <remarks></remarks>
    Sub insertDataToHis(ByVal sApproved As String)
        operationXML("1")
    End Sub

    ''' <summary>
    ''' 加載資料
    ''' </summary>
    ''' <remarks></remarks>
    Sub loadNormalData()
        Try
            Dim syTempInfo As New SY_TEMPINFO(GetDatabaseManager())
            Dim syRelRoleFunctionList As New SY_REL_ROLE_FUNCTIONList(GetDatabaseManager())
            Dim xmlData As String = String.Empty
            Dim xmlDocument As New XmlDocument
            hidRoleId.Value = treeViewRole.SelectedValue

            ' 綁定交易
            'bindTreeFun()
            initFunctionTreeView(treeViewFun)

            ' 展開節點
            treeViewFun.ExpandAll()

            ' 查詢臨時表
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
                If xmlData.Length > 0 Then

                    ' document對象載入XML文件
                    xmlDocument.LoadXml(xmlData)
                Else
                    If syRelRoleFunctionList.loadAllDataByRoleId(hidRoleId.Value.Split("*")(0), hidRoleId.Value.Split("*")(1), hidRoleId.Value.Split("*")(2)) Then
                        For Each item As SY_REL_ROLE_FUNCTION In syRelRoleFunctionList
                            m_sFuncCodeList &= item.getAttribute("FUNCCODE") & ","
                        Next
                    End If
                End If
            End If

            '------------------------------ 正常區域加載資料,查詢xml中的資料，將其新增到左邊目錄樹---------------------------------------------------
            operationXML("2")

            If Not String.IsNullOrEmpty(m_sCodeFlag) Then

                ' 第一層是安泰銀行
                For Each firstItem As TreeNode In treeViewFun.Nodes
                    If firstItem.ChildNodes.Count > 0 Then

                        ' 父節點
                        For Each secondItem As TreeNode In firstItem.ChildNodes
                            If secondItem.ChildNodes.Count > 0 Then

                                If Not m_sCodeFlag Is Nothing Then

                                    ' 子系統
                                    For Each item As String In m_sCodeFlag.Split("|")

                                        For Each threeItem As TreeNode In secondItem.ChildNodes

                                            ' 臨時表中有資料，則選中
                                            If item.Split(",")(0).ToString() = threeItem.Value Then
                                                '
                                                If item.Split(",")(1).ToString() = "I" Then
                                                    threeItem.Checked = True
                                                    threeItem.Text = String.Format("<font color='red'>{0}</font>", threeItem.Text)
                                                ElseIf item.Split(",")(1).ToString() = "N" Then
                                                    threeItem.Checked = True
                                                ElseIf item.Split(",")(1).ToString() = "D" Then
                                                    threeItem.ImageUrl = "../img/del.gif"
                                                    threeItem.ShowCheckBox = False
                                                End If
                                            End If
                                        Next
                                    Next
                                End If
                            End If
                        Next
                    End If
                Next
            Else
                ' 已經存在于live表中的資料
                If Not String.IsNullOrEmpty(m_sFuncCodeList) Then

                    ' 第一層是安泰銀行
                    For Each firstItem As TreeNode In treeViewFun.Nodes
                        If firstItem.ChildNodes.Count > 0 Then

                            ' 父節點
                            For Each secondItem As TreeNode In firstItem.ChildNodes
                                If secondItem.ChildNodes.Count > 0 Then

                                    If Not m_sCodeFlag Is Nothing Then

                                        ' 子系統
                                        For Each item As String In m_sFuncCodeList.Split(",")

                                            For Each threeItem As TreeNode In secondItem.ChildNodes

                                                ' 臨時表中有資料，則選中
                                                If item.ToString() = threeItem.Value Then
                                                    threeItem.Checked = True
                                                End If
                                            Next
                                        Next
                                    End If
                                End If
                            Next
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 判斷節點是選中還是取消勾選
    ''' </summary>
    ''' <param name="node"></param>
    ''' <param name="bCheck"></param>
    ''' <remarks></remarks>
    Sub checkItem(ByVal node As TreeNode, ByVal bCheck As Boolean)
        For Each item As TreeNode In node.ChildNodes
            item.Checked = bCheck

            If bCheck = True Then
                If Not item.Text.Contains("red") Then
                    item.Text = String.Format("<font color='red'>{0}</font>", item.Text)
                End If
            Else
                If Not item.Text.Contains("disabled") Then
                    item.Text = String.Format("<font color='red'>{0}</font>", item.Text)
                    hidDelData.Value &= item.Value.Split("$")(0) & ";"
                End If
            End If

            If item.ChildNodes.Count > 0 Then
                checkItem(item, bCheck)
            End If
        Next
    End Sub

    ''' <summary>
    ''' 操作節點是否可用
    ''' </summary>
    ''' <param name="node"></param>
    ''' <param name="sCodeFlag"></param>
    ''' <param name="sFunccodeList"></param>
    ''' <param name="bFlag"></param>
    ''' <remarks></remarks>
    Sub operationNode(ByVal node As TreeNode, ByVal sCodeFlag As String, ByVal sFunccodeList As String, ByVal bFlag As Boolean, Optional ByVal sChildCodeList As String = "")
        If sFunccodeList.Length > 0 Then
            If sFunccodeList.Substring(sFunccodeList.Length - 1, 1) <> "," Then
                sFunccodeList = sFunccodeList & ","
            End If

            sFunccodeList = "," & sFunccodeList
        End If

        If node.ChildNodes.Count > 0 Then
            For Each itemNode As TreeNode In node.ChildNodes
                If Not String.IsNullOrEmpty(sCodeFlag) Then

                    ' 子系統
                    For Each item As String In m_sCodeFlag.Split("|")

                        ' 臨時表中有資料，則選中
                        If m_sCodeFlag.Contains(itemNode.Value.Split("$")(0)) Then
                            If item.Length > 0 Then
                                If item.Split(",")(0).ToString = itemNode.Value.Split("$")(0) Then
                                    If item.Split(",")(1).ToString() = "I" Then
                                        itemNode.Checked = True
                                        itemNode.Parent.Checked = True
                                        itemNode.Text = String.Format("<font color='red'>{0}</font>", itemNode.Text)
                                    ElseIf item.Split(",")(1).ToString() = "N" Then
                                        itemNode.Checked = True
                                        itemNode.Parent.Checked = True
                                    ElseIf item.Split(",")(1).ToString() = "D" Then

                                        'itemNode.ImageUrl = "../img/del.gif"
                                        'itemNode.ShowCheckBox = False
                                        itemNode.Checked = False
                                        itemNode.Text = String.Format("<font color='red'>{0}</font>", itemNode.Text)
                                    ElseIf item.Split(",")(1).ToString() = "X" Then
                                        itemNode.Checked = False
                                    End If
                                End If
                            End If

                            If itemNode.ChildNodes.Count > 0 Then
                                operationNode(itemNode, sCodeFlag, sFunccodeList, bFlag, sChildCodeList)
                            End If
                        Else
                            ' 如果不包含在代號集合中，則此節點設置為disable狀態，但是處於同一父節點的節點，若有選中，則父節點選中，不可將其設置為disable狀態
                            If bFlag = False Then

                                '  itemNode.Checked = False
                                If itemNode.ShowCheckBox = True Then
                                    itemNode.ShowCheckBox = False
                                    If Not itemNode.Text.Contains("<input type='checkbox' disabled = 'disabled' />") Then
                                        itemNode.Text = String.Format("<input type='checkbox' disabled = 'disabled' />{0}", itemNode.Text)
                                    End If
                                End If

                                ' 設置其子節點也不可使用
                                If itemNode.ChildNodes.Count > 0 Then
                                    disableNode(itemNode, sCodeFlag)
                                End If
                            End If
                        End If
                    Next
                Else

                    If itemNode.Text.Contains("<input type='checkbox' disabled = 'disabled' />") Then
                        itemNode.Text = itemNode.Text.Replace("<input type='checkbox' disabled = 'disabled' />", "")
                    End If

                    If itemNode.Parent.Text.Contains("<input type='checkbox' disabled = 'disabled' />") Then
                        itemNode.Parent.Text = itemNode.Parent.Text.Replace("<input type='checkbox' disabled = 'disabled' />", "")
                    End If
                End If
            Next
        Else
            If bFlag = False Then

                ' 沒有子節點的交易，若不在m_sFuncCodeList中 也需要設置成不可是狀態
                If Not sFunccodeList.Contains("," & node.Value & ",") Then
                    node.ShowCheckBox = False
                    If Not node.Text.Contains("<input type='checkbox' disabled = 'disabled' />") Then
                        node.Text = String.Format("<input type='checkbox' disabled = 'disabled' />{0}", node.Text)
                    End If

                    If node.ChildNodes.Count > 0 Then
                        operationNode(node, sCodeFlag, sFunccodeList, bFlag, sChildCodeList)
                    End If
                End If
            Else
                If sFunccodeList.Contains("," & node.Value & ",") Then
                    node.Checked = True
                ElseIf sCodeFlag.Contains(node.Value) Then
                    node.Checked = True
                End If
            End If
        End If
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

#End Region

#Region "正常資料Event"

    ''' <summary>
    ''' 儲存按鈕
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnSave_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnSave.Click
        Try
            Dim syRelRoleFunctionHisList As New SY_REL_ROLE_FUNCTION_HISList(GetDatabaseManager())
            Dim syRelRoleFunctionList As New SY_REL_ROLE_FUNCTIONList(GetDatabaseManager())
            Dim syTempInfo As New SY_TEMPINFO(GetDatabaseManager())

            ' 開始事物
            GetDatabaseManager().beginTran()

            ' 組XML文件
            m_sbXmlData.Append("<SY>")

            '------------------------------ 正常區域加載資料,儲存按鈕---------------------------------------------------
            operationXML("3")

            ' 查找存在于live表中的資料
            If hidRoleId.Value <> "" Then
                If syRelRoleFunctionList.loadAllDataByRoleId(hidRoleId.Value.Split("*")(0), hidRoleId.Value.Split("*")(1), hidRoleId.Value.Split("*")(2)) Then
                    For Each item As SY_REL_ROLE_FUNCTION In syRelRoleFunctionList
                        m_sOldCodeList &= item.getAttribute("FUNCCODE") & ";N,"
                    Next
                End If
            End If

            ' 未修改的資料及刪除的資料
            For Each node As TreeNode In treeViewFun.CheckedNodes
                m_sbXmlData.Append("<SY_REL_ROLE_FUNCTION>")
                m_sbXmlData.Append("<ROLEID>" & hidRoleId.Value.Split("*")(0) & "</ROLEID>")
                m_sbXmlData.Append("<FUNCCODE>" & node.Value.Split("$")(0) & "</FUNCCODE>")
                m_sbXmlData.Append("<FUCNAME>" & node.Text.Replace("<font color='red'>", "").Replace("</font>", "") & "</FUCNAME>")
                m_sbXmlData.Append("<SYSID>" & hidRoleId.Value.Split("*")(1) & "</SYSID>")
                m_sbXmlData.Append("<SUBSYSID>" & hidRoleId.Value.Split("*")(2) & "</SUBSYSID>")

                ' 未改變的資料
                If m_sOldCodeList.Length > 0 Then

                    ' 若Node存在于此字符串
                    If m_sOldCodeList.Contains(node.Value.Split("$")(0) & ";") Then

                        ' 查找定位到這個交易編號的Operation
                        Dim startIndex As Integer = m_sOldCodeList.IndexOf(node.Value.Split("$")(0) & ";")
                        startIndex = startIndex + Convert.ToInt32(node.Value.Split("$")(0).Length) + 1

                        Dim operation As String = m_sOldCodeList.Substring(startIndex, 1)
                        If operation = "I" Then
                            m_sbXmlData.Append("<OPERATION>I</OPERATION>")
                        Else
                            m_sbXmlData.Append("<OPERATION>N</OPERATION>")
                        End If
                    Else
                        m_sbXmlData.Append("<OPERATION>I</OPERATION>")
                    End If
                Else
                    m_sbXmlData.Append("<OPERATION>I</OPERATION>")
                End If

                m_sbXmlData.Append("</SY_REL_ROLE_FUNCTION>")
            Next

            ' 查詢那些刪除過的資料
            If hidDelData.Value.Length > 0 Then
                For Each Str As String In hidDelData.Value.Split(";")
                    m_sbXmlData.Append("<SY_REL_ROLE_FUNCTION>")
                    m_sbXmlData.Append("<ROLEID>" & hidRoleId.Value.Split("*")(0) & "</ROLEID>")
                    m_sbXmlData.Append("<FUNCCODE>" & Str.Split("*")(0) & "</FUNCCODE>")
                    m_sbXmlData.Append("<FUCNAME>" & Str.Replace("<font color='red'>", "").Replace("</font>", "") & "</FUCNAME>")
                    m_sbXmlData.Append("<SYSID>" & hidRoleId.Value.Split("*")(1) & "</SYSID>")
                    m_sbXmlData.Append("<SUBSYSID>" & hidRoleId.Value.Split("*")(2) & "</SUBSYSID>")
                    m_sbXmlData.Append("<OPERATION>D</OPERATION>")
                    m_sbXmlData.Append("</SY_REL_ROLE_FUNCTION>")
                Next
            End If

            m_sbXmlData.Append("</SY>")

            Dim dtNewData As DataTable = com.Azion.EloanUtility.XmlUtility.convertXML2DataTable(m_sbXmlData.ToString)

            '查詢之前的資料()
            Dim dtSYTempInfo As DataTable = loadSYTempInfo()

            ' 如果為同一個角色 則刪除即可 若不是保留 進行累加
            ' 查找臨時檔中的資料 進行修改
            If Not dtSYTempInfo Is Nothing Then
                Dim rows() As DataRow = dtSYTempInfo.Select("ROLEID='" & hidRoleId.Value.Split("*")(0) & "' AND SYSID= '" & hidRoleId.Value.Split("*")(1) & "' AND SUBSYSID = '" & hidRoleId.Value.Split("*")(2) & "' ")

                ' 找到臨時檔中已經存在的資料
                ' 如果有 則將其替換掉 然後累加上新的XML
                If rows.Length > 0 Then
                    For i As Integer = 0 To rows.Count - 1
                        dtSYTempInfo.Rows.Remove(rows(i))
                        dtSYTempInfo.AcceptChanges()
                    Next

                    '進行累加
                    If Not dtNewData Is Nothing Then
                        Dim rowsNewData() As DataRow = dtNewData.Select("ROLEID='" & hidRoleId.Value.Split("*")(0) & "' AND SYSID= '" & hidRoleId.Value.Split("*")(1) & "' AND SUBSYSID = '" & hidRoleId.Value.Split("*")(2) & "' ")

                        For Each row As DataRow In rowsNewData
                            row.AcceptChanges()
                            dtSYTempInfo.ImportRow(row)
                        Next
                    End If
                Else
                    '進行累加
                    If Not dtNewData Is Nothing Then
                        Dim rowsNewData() As DataRow = dtNewData.Select("ROLEID='" & hidRoleId.Value.Split("*")(0) & "'")
                        For Each row As DataRow In rowsNewData
                            dtSYTempInfo.ImportRow(row)
                        Next
                    End If
                End If
                dtSYTempInfo.AcceptChanges()
            Else ' 如果找不到臨時檔，直接新增資料
                dtSYTempInfo = dtNewData
            End If

            If m_sFourStepNo = "" Then
                If Not syTempInfo.loadByPK(m_sWorkingUserid, m_sWorkingTopDepNo, m_sFuncCode) Then
                    syTempInfo.setAttribute("STAFFID", m_sWorkingUserid)
                    syTempInfo.setAttribute("BRA_DEPNO", m_sWorkingTopDepNo)
                    syTempInfo.setAttribute("FUNCCODE", m_sFuncCode)
                End If

                If m_sbXmlData.ToString() = "<SY></SY>" Then
                    syTempInfo.setAttribute("TEMPDATA", "")
                Else
                    syTempInfo.setAttribute("TEMPDATA", com.Azion.EloanUtility.XmlUtility.convertDataTable2XML(dtSYTempInfo).ToString)
                End If

                syTempInfo.setAttribute("CASEID", m_sCaseId)
                syTempInfo.save()
            Else
                If Not syTempInfo.loadByCaseId(m_sCaseId) Then
                    syTempInfo.setAttribute("CASEID", m_sCaseId)
                End If

                If m_sbXmlData.ToString() = "<SY></SY>" Then
                    syTempInfo.setAttribute("TEMPDATA", "")
                Else
                    syTempInfo.setAttribute("TEMPDATA", com.Azion.EloanUtility.XmlUtility.convertDataTable2XML(dtSYTempInfo).ToString)
                End If

                syTempInfo.save()
            End If

            com.Azion.EloanUtility.UIUtility.alert("儲存成功！")

            ' 找到操作的資料 變紅
            If hidValuePath.Value <> "" Then
                Dim nodeSelect As TreeNode = treeViewRole.FindNode(hidValuePath.Value)
                nodeSelect.Text = String.Format("<font color='red'>{0}</font>", nodeSelect.Text)
            End If

            ' 提交
            GetDatabaseManager().commit()
        Catch ex As Exception

            ' 事物回滾
            GetDatabaseManager.Rollback()
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 取消
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnCancel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnCancel.Click
        If hidRoleId.Value <> "" Then
            treeViewFun.Nodes.Clear()
            treeViewRole_SelectedNodeChanged(sender, e)
        End If
    End Sub

    ''' <summary>
    ''' 查詢編輯節點的父節點
    ''' </summary>
    ''' <param name="node"></param>
    ''' <remarks></remarks>
    Sub roleList(ByVal node As TreeNode, ByVal sSysIdList As String)
        If Not node.Parent Is Nothing And node.ValuePath.ToString.Contains("*") Then
            If Not sSysIdList.Contains(node.ValuePath.Split("*")(0).Split("/")(3)) Then

                ' 取得父節點的ValuePath
                Dim str As String = node.Parent.ValuePath

                ' 取得父節點的系統與子系統
                Dim sLastIndex As String = hidRoleId.Value.Split("*")(1) & "*" & hidRoleId.Value.Split("*")(2)

                ' 取得最後一個“/”
                Dim strSum() As String = node.Parent.Value.Split("/")

                ' 取得父節點的ROLEID*SYSID*SUBSYSID
                Dim strLast As String = String.Empty

                If strSum.Length > 0 Then
                    strLast = strSum(strSum.Length - 1).ToString
                End If

                Dim strResult As String = strLast.Replace(sLastIndex, "").Replace("*", "")

                m_sParentRole &= "'" & strResult & "'" & ","
            End If

        End If
    End Sub

    ''' <summary>
    ''' 節點不可用
    ''' </summary>
    ''' <param name="node"></param>
    ''' <param name="sCodeList"></param>
    ''' <remarks></remarks>
    Sub disableNode(ByVal node As TreeNode, ByVal sCodeList As String)
        If node.ChildNodes.Count > 0 Then

            For Each item As TreeNode In node.ChildNodes

                If Not m_sCodeFlag.Contains("|" & item.Value.Split("$")(0) & ",") Then

                    If item.ShowCheckBox = True Then
                        item.ShowCheckBox = False
                        If Not item.Text.Contains("<input type='checkbox' disabled = 'disabled' />") Then
                            item.Text = String.Format("<input type='checkbox' disabled = 'disabled' />{0}", item.Text)
                        End If

                        If item.ChildNodes.Count > 0 Then
                            disableNode(item, m_sCodeFlag)
                        End If
                    End If
                End If
            Next
        End If
    End Sub

    ''' <summary>
    ''' 更新
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub treeViewRole_SelectedNodeChanged(ByVal sender As Object, ByVal e As EventArgs) Handles treeViewRole.SelectedNodeChanged
        Try
            Dim syRelRoleFunctionList As New SY_REL_ROLE_FUNCTIONList(GetDatabaseManager())
            Dim syTempInfo As New SY_TEMPINFO(GetDatabaseManager())
            Dim xmlData As String = String.Empty
            Dim dsData As New DataSet
            Dim dtData As New DataTable
            Dim sChildCodeList As String = String.Empty

            ' 若其text為紅色 說明在真實檔存在 需要先設置其text為乾淨的text 無<font> 然後設置其前景色為黑色
            If treeViewRole.SelectedNode.Text.Contains("red") Then
                treeViewRole.SelectedNode.Text = treeViewRole.SelectedNode.Text.Replace("<font color='red'>", "").Replace("</font>", "")
                ViewState("lastPath") = treeViewRole.SelectedNode.ValuePath
            Else
                If Not ViewState("lastPath") Is Nothing Then
                    treeViewRole.FindNode(ViewState("lastPath")).Text = String.Format("<font color='red'>{0}</font>", treeViewRole.FindNode(ViewState("lastPath")).Text)
                End If
            End If

            treeViewRole.SelectedNodeStyle.ForeColor = Drawing.Color.Black

            hidDelData.Value = ""

            ' 選中值
            hidRoleId.Value = treeViewRole.SelectedNode.Value
            hidRoleName.Value = treeViewRole.SelectedNode.Text.Replace("<font color='red'>", "").Replace("</font>", "")

            m_sCodeFlag = "|"

            treeViewFun.Nodes.Clear()
            m_sFuncCodeList = ""

            ' 綁定交易樹
            initFunctionTreeView(treeViewFun)

            ' 設置CHECKBOX狀態
            initCheckBox()

            ' 展開節點
            treeViewFun.ExpandAll()

            If m_sFourStepNo = "" Then
                If syTempInfo.loadTempData(m_sWorkingUserid, m_sWorkingTopDepNo, m_sFuncCode) Then
                    xmlData = syTempInfo.getString("TEMPDATA")

                    If xmlData <> Nothing Then

                        Dim xmlDocument As New XmlDocument
                        xmlDocument.LoadXml(xmlData)
                        dsData.ReadXml(New XmlNodeReader(xmlDocument))
                        dtData = dsData.Tables(0)
                    End If
                End If
            Else
                If syTempInfo.loadByCaseId(m_sCaseId) Then
                    xmlData = syTempInfo.getString("TEMPDATA")

                    If xmlData <> Nothing Then

                        Dim xmlDocument As New XmlDocument
                        xmlDocument.LoadXml(xmlData)
                        dsData.ReadXml(New XmlNodeReader(xmlDocument))
                        dtData = dsData.Tables(0)
                    End If
                End If
            End If

            ' 查詢出所有的子系統來
            Dim sySubSysId As New AUTH_OP.SY_SUBSYSIDList(GetDatabaseManager())
            Dim sySysId As New AUTH_OP.SY_SYSIDList(GetDatabaseManager())
            Dim syAllSubSysId As String = String.Empty
            Dim bFlag As Boolean = False
            Dim sSysIdList As String = String.Empty

            If sySubSysId.loadAllData() Then
                For Each item As SY_SUBSYSID In sySubSysId
                    syAllSubSysId &= item.getAttribute("SUBSYSID").ToString() & ","
                Next
            End If

            If sySysId.loadAllData() Then
                For Each item As AUTH_OP.SY_SYSID In sySysId
                    sSysIdList &= item.getAttribute("SYSID").ToString() & ","
                Next
            End If

            ' 第一步：根據當前編輯節點的Value在查詢出來的臨時檔中查看是否有符合條件的資料
            If dtData.Rows.Count > 0 Then
                Dim drRowTemp() As DataRow = dtData.Select(String.Format("ROLEID='{0}' and SYSID='{1}' and SUBSYSID='{2}'", treeViewRole.SelectedNode.Value.Split("*")(0), treeViewRole.SelectedNode.Value.Split("*")(1), treeViewRole.SelectedNode.Value.Split("*")(2)))
                If drRowTemp.Count > 0 Then
                    For Each dr As DataRow In drRowTemp

                        ' 第二步：記錄存在于臨時檔中的資料，并記錄其操作方式，方便在右側的TreeView中操作節點
                        m_sCodeFlag &= dr("FUNCCODE").ToString & "," & dr("OPERATION").ToString & "|"
                    Next
                End If
            End If

            ' 第三步：查詢存在于實際表中的交易編號,然後將其追加到m_sCodeFlag中(要找不存在于臨時檔中的未異動的資料)
            If syRelRoleFunctionList.loadAllDataByRoleId(treeViewRole.SelectedNode.Value.Split("*")(0), treeViewRole.SelectedNode.Value.Split("*")(1), treeViewRole.SelectedNode.Value.Split("*")(2)) Then
                For Each item As SY_REL_ROLE_FUNCTION In syRelRoleFunctionList
                    If Not m_sCodeFlag.Contains(item.getAttribute("FUNCCODE")) Then
                        m_sCodeFlag &= item.getAttribute("FUNCCODE").ToString & ",N" & "|"
                    End If
                Next
            End If

            ' 第四步：若臨時檔中未有符合條件的資料，真實檔中亦沒有，則需要查詢此節點的父節點，判斷其選取範圍
            roleList(treeViewRole.SelectedNode, sSysIdList)

            ' 排除掉系統，子系統的資料
            For Each sSysIdTemp As String In sSysIdList.Split(",")
                If m_sParentRole.Contains("'" & sSysIdTemp & "'") Then
                    m_sParentRole = m_sParentRole.Replace("'" & sSysIdTemp & "'" & ",", "")
                End If
            Next

            For Each sSubSysIdTemp As String In syAllSubSysId.Split(",")
                If m_sParentRole.Contains("'" & sSubSysIdTemp & "'") Then
                    m_sParentRole = m_sParentRole.Replace("'" & sSubSysIdTemp & "'" & ",", "")
                End If
            Next

            If m_sParentRole.Length > 0 Then
                m_sParentRole = m_sParentRole.Substring(0, m_sParentRole.Length - 1)

                If dtData.Rows.Count > 0 Then

                    ' 第五步：判斷其父節點是否有在臨時檔總存在資料,設置標誌符為X,表示右邊交易中checkbox可勾選
                    Dim drRowTemp() As DataRow = dtData.Select(String.Format("ROLEID in (" & m_sParentRole & ") and SYSID='{0}' and SUBSYSID='{1}'", treeViewRole.SelectedNode.Value.Split("*")(1), treeViewRole.SelectedNode.Value.Split("*")(2)))
                    If drRowTemp.Count > 0 Then
                        For Each dr As DataRow In drRowTemp
                            If Not m_sCodeFlag.Contains(dr("FUNCCODE")) Then
                                If dr("OPERATION") <> "D" Then
                                    m_sCodeFlag &= dr("FUNCCODE").ToString & ",X" & "|"
                                Else
                                    m_sCodeFlag &= dr("FUNCCODE").ToString & ",D" & "|"
                                End If
                            End If
                        Next
                    End If
                End If

                ' 第六步：若臨時檔中沒有資料，則需要查詢其真實檔中是否有資料
                If syRelRoleFunctionList.loadAddDataByCon(m_sParentRole, treeViewRole.SelectedNode.Value.Split("*")(1), treeViewRole.SelectedNode.Value.Split("*")(2)) Then
                    For Each item As SY_REL_ROLE_FUNCTION In syRelRoleFunctionList

                        If Not m_sCodeFlag.Contains(item.getAttribute("FUNCCODE")) Then
                            m_sCodeFlag &= item.getAttribute("FUNCCODE").ToString & ",X" & "|"
                        End If
                    Next
                End If
            End If

            '---------------------------------------------------------設置交易樹的可用與否------------------------------------------------------------------------------------

            ' 如果選中節點的父節點在子系統範圍內 則說明此節點是根節點，右邊的交易樹全部開放
            If syAllSubSysId.Contains(treeViewRole.SelectedNode.Parent.Value) Then
                ' 右邊交易樹正常顯示即可，不用做額外處理
                bFlag = True
            End If
            '--------------------------------------------------------------------------------------------------------------------------------------------
            '------------------------------ 正常區域,更新操作---------------------------------------------------

            If bFlag Then

                If Not String.IsNullOrEmpty(m_sCodeFlag) Then

                    ' 第一層是安泰銀行
                    For Each firstItem As TreeNode In treeViewFun.Nodes
                        If firstItem.ChildNodes.Count > 0 Then

                            ' 父節點
                            For Each secondItem As TreeNode In firstItem.ChildNodes
                                operationNode(secondItem, m_sCodeFlag, "", True)
                            Next
                        End If
                    Next
                End If
            Else

                ' 查詢出其父節點 所勾選的交易項目，其他的項目為不可用

                ' 第一層是安泰銀行
                For Each firstItem As TreeNode In treeViewFun.Nodes
                    If firstItem.ChildNodes.Count > 0 Then

                        ' 父節點
                        For Each secondItem As TreeNode In firstItem.ChildNodes

                            If m_sCodeFlag.Length > 0 Then
                                If Not m_sCodeFlag.Contains("|" & secondItem.Value.Split("$")(0) & ",") Then

                                    If secondItem.ShowCheckBox = True Then
                                        secondItem.ShowCheckBox = False
                                        If Not secondItem.Text.Contains("<input type='checkbox' disabled = 'disabled' />") Then
                                            secondItem.Text = String.Format("<input type='checkbox' disabled = 'disabled' />{0}", secondItem.Text)
                                        End If

                                        If secondItem.ChildNodes.Count > 0 Then
                                            disableNode(secondItem, m_sCodeFlag)
                                        End If
                                    End If
                                Else
                                    operationNode(secondItem, m_sCodeFlag, m_sFuncCodeList, False)

                                    ' 循環父節點 查看其是否變色或怎樣
                                    For Each item As String In m_sCodeFlag.Split("|")

                                        ' 臨時表中有資料，則選中
                                        If m_sCodeFlag.Contains(secondItem.Value.Split("$")(0)) Then
                                            If item.Length > 0 Then
                                                If item.Split(",")(0).ToString = secondItem.Value.Split("$")(0) Then
                                                    If item.Split(",")(1).ToString() = "I" Then
                                                        secondItem.Checked = True
                                                        secondItem.Parent.Checked = True
                                                        secondItem.Text = String.Format("<font color='red'>{0}</font>", secondItem.Text)
                                                    ElseIf item.Split(",")(1).ToString() = "N" Then
                                                        secondItem.Checked = True
                                                        secondItem.Parent.Checked = True
                                                    ElseIf item.Split(",")(1).ToString() = "D" Then
                                                        secondItem.Checked = False
                                                        secondItem.Text = String.Format("<font color='red'>{0}</font>", secondItem.Text)
                                                    ElseIf item.Split(",")(1).ToString() = "X" Then
                                                        secondItem.Checked = False
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Next
                                End If
                            End If
                        Next
                    End If
                Next
            End If

            ' 遍歷樹 若子節點為disabled 則父節點disabled


            ViewState("m_sFuncCodeList") = m_sFuncCodeList
            ViewState("m_sCodeFlag") = m_sCodeFlag

            hidValuePath.Value = treeViewRole.SelectedNode.ValuePath
        Catch ex As Exception
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

            Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager())
            Dim syRelRoleFunctionHis As New SY_REL_ROLE_FUNCTION_HIS(GetDatabaseManager())
            Dim flowFacade As New FLOW_OP.FlowFacade(GetDatabaseManager())
            Dim syRelRoleFunctionList As New AUTH_OP.SY_REL_ROLE_FUNCTIONList(GetDatabaseManager())

            Dim m_sFlowName As String = String.Empty ' 流程編號
            Dim sTempData As String = String.Empty
            Dim sFuncRoot As String = String.Empty

            m_sFlowName = com.Azion.EloanUtility.CodeList.getAppSettings("SY_UPDFUNCTIONGRANTROLE")

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

            GetDatabaseManager.beginTran()
            Dim stepInfo As FLOW_OP.StepInfo

            m_bCheck = False
            If m_bCheck = False Then

                If m_sStepNo = "" Then
                    stepInfo = MBSC.UICtl.UIShareFun.startFlow(GetDatabaseManager(), m_sFlowName, True)

                    Dim sNewCaseId = stepInfo.currentStepInfo.caseId

                    '不可能沒資料 
                    If syTempInfo.loadByPK(m_sWorkingUserid, m_sWorkingTopDepNo, m_sFuncCode) Then
                        syTempInfo.setAttribute("CASEID", sNewCaseId)
                        syTempInfo.save()
                    End If

                    'com.Azion.EloanUtility.UIUtility.alert("存檔成功！")

                    ' 跳轉到代辦清單
                    'com.Azion.EloanUtility.UIUtility.Redirect("SY_CASELIST.aspx")
                Else
                    stepInfo = flowFacade.SendFlow(m_sCaseId)
                End If

                updatTemp2Live(GetDatabaseManager(), stepInfo, "")
            Else

                stepInfo = MBSC.UICtl.UIShareFun.startFlow(GetDatabaseManager(), m_sFlowName, True, , True)
                updatTemp2Live(GetDatabaseManager(), stepInfo, "Y")

                ' 將選中節點的顏色變黑
                resetTreeViewState(treeViewFun.CheckedNodes)
                resetTreeViewState(treeViewRole.Nodes)

                treeViewRole_SelectedNodeChanged(sender, e)

                ' 將選中節點的資料變黑
                If Not treeViewRole.SelectedNode Is Nothing Then
                    treeViewRole.SelectedNode.Text = treeViewRole.SelectedNode.Text.Replace("<font color='red'>", "").Replace("</font>", "")

                    If Not treeViewRole.SelectedNode.Parent Is Nothing Then
                        treeViewRole.SelectedNode.Parent.Text = treeViewRole.SelectedNode.Parent.Text.Replace("<font color='red'>", "").Replace("</font>", "")
                    End If
                End If
            End If

            GetDatabaseManager.commit()

            ' 成功提示信息
            com.Azion.EloanUtility.UIUtility.showJSMsg("案件編號：" + stepInfo.currentStepInfo.caseId + " 送出成功！")

            ' 跳轉到待辦事項頁面
            If String.IsNullOrEmpty(m_sStepNo) Then
                com.Azion.EloanUtility.UIUtility.goMainPage("")
            Else
                com.Azion.EloanUtility.UIUtility.closeWindow()
            End If

        Catch ex As Exception

            GetDatabaseManager.Rollback()
            Throw ex
        End Try
    End Sub

    ''' <summary>
    ''' 交易是否選中
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub treeViewFun_TreeNodeCheckChanged(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.TreeNodeEventArgs) Handles treeViewFun.TreeNodeCheckChanged

        Dim syTempInfo As New SY_TEMPINFO(GetDatabaseManager())
        Dim syRelRoleFunctionList As New SY_REL_ROLE_FUNCTIONList(GetDatabaseManager())
        Dim xmlData As String = String.Empty
        Dim dsData As New DataSet
        Dim dtData As New DataTable
        Dim dtLiveData As New DataTable

        If String.IsNullOrEmpty(m_sFuncCodeList) Then
            m_sFuncCodeList = ViewState("m_sFuncCodeList").ToString()
        End If

        ' 取得臨時檔中的資料
        If m_sFourStepNo = "" Then
            If syTempInfo.loadTempData(m_sWorkingUserid, m_sWorkingTopDepNo, m_sFuncCode) Then
                xmlData = syTempInfo.getString("TEMPDATA")

                If xmlData <> Nothing Then

                    Dim xmlDocument As New XmlDocument
                    xmlDocument.LoadXml(xmlData)
                    dsData.ReadXml(New XmlNodeReader(xmlDocument))
                    dtData = dsData.Tables(0)
                End If
            End If
        Else
            If syTempInfo.loadByCaseId(m_sCaseId) Then
                xmlData = syTempInfo.getString("TEMPDATA")

                If xmlData <> Nothing Then

                    Dim xmlDocument As New XmlDocument
                    xmlDocument.LoadXml(xmlData)
                    dsData.ReadXml(New XmlNodeReader(xmlDocument))
                    dtData = dsData.Tables(0)
                End If
            End If
        End If

        If syRelRoleFunctionList.loadDataByPK(treeViewRole.SelectedNode.Value.Split("*")(0), treeViewRole.SelectedNode.Value.Split("*")(1), treeViewRole.SelectedNode.Value.Split("*")(2), e.Node.Value.Split("$")(0)) Then
            dtLiveData = syRelRoleFunctionList.getCurrentDataSet.Tables(0)
        End If

        ' 取消勾選
        If e.Node.Checked = True Then

            ' 判斷是否有父節點，如果有，則子節點選中，父節點也選中
            If Not e.Node.Parent Is Nothing Then
                checkParentNode(e.Node, True, dtData)
            End If

            ' 判斷是否有子節點，如果有子節點，則父節點選中，子接地亦選中
            If e.Node.ChildNodes.Count > 0 Then
                For Each item As TreeNode In e.Node.ChildNodes

                    ' 只勾選可用的資料，并將其設置為紅色，disable的資料不允許變為紅色
                    If item.Checked = False AndAlso (Not item.Text.Contains("disabled")) Then
                        item.Checked = True
                        If Not item.Text.Contains("red") Then
                            item.Text = String.Format("<font color='red'>{0}</font>", item.Text)
                        End If

                        If item.ChildNodes.Count > 0 Then
                            checkItem(item, True)
                        End If
                    End If
                Next
            End If

            ' 如果之前有勾選，則需要變色
            If m_sFuncCodeList.Contains(e.Node.Value.Split("$")(0)) Then
                e.Node.Text = String.Format("<font color='red'>{0}</font>", e.Node.Text)
            End If

            ' 如果勾選了 之前沒有勾選的 則也需要變色
            If Not m_sFuncCodeList.Contains(e.Node.Value.Split("$")(0)) Then
                If Not e.Node.Text.Contains("red") Then
                    e.Node.Text = String.Format("<font color='red'>{0}</font>", e.Node.Text)
                End If
            End If

            ' 如果取消過的資料再重新選中 那麼應該保持不變
            If hidDelData.Value.Contains(e.Node.Value.Split("$")(0)) Then
                hidDelData.Value = hidDelData.Value.Replace(e.Node.Value.Split("$")(0) & ";", "")
                e.Node.Text = e.Node.Text.Replace("<font color='red'>", "").Replace("</font>", "")
            End If
        Else

            ' 查看此節點是否在Live檔中存在，是否在臨時檔中存在。 如果不是在這兩個檔存在，說明是新增資料，無需記錄hidDelData
            If Not dtData Is Nothing Then
                If dtData.Rows.Count > 0 Then

                    Dim drRowTemp() As DataRow = dtData.Select(String.Format("ROLEID ='{0}' and SYSID='{1}' and SUBSYSID='{2}' AND FUNCCODE = '{3}'", treeViewRole.SelectedNode.Value.Split("*")(0), treeViewRole.SelectedNode.Value.Split("*")(1), treeViewRole.SelectedNode.Value.Split("*")(2), e.Node.Value.Split("$")(0)))
                    If drRowTemp.Count > 0 Then
                        For Each dr As DataRow In drRowTemp
                            If dr("OPERATION").ToString() <> "I" Then

                                e.Node.Text = String.Format("<font color='red'>{0}</font>", e.Node.Text)
                                ' 記錄刪除的資料
                                hidDelData.Value &= e.Node.Value.Split("$")(0) & ";"
                            End If
                        Next
                    End If
                End If
            Else
                e.Node.Text = e.Node.Text.Replace("<font color='red'>", "").Replace("</font>", "")
            End If

            If Not dtLiveData Is Nothing Then
                If dtLiveData.Rows.Count > 0 Then
                    e.Node.Text = String.Format("<font color='red'>{0}</font>", e.Node.Text)
                    ' 記錄刪除的資料
                    hidDelData.Value &= e.Node.Value.Split("$")(0) & ";"
                End If
            Else
                e.Node.Text = e.Node.Text.Replace("<font color='red'>", "").Replace("</font>", "")
            End If

            If Not e.Node.Parent Is Nothing Then

                checkParentNode(e.Node, False, dtData)
                e.Node.Parent.Text = String.Format("<font color='red'>{0}</font>", e.Node.Parent.Text)

                If e.Node.Parent.ChildNodes.Count = 1 Then
                    hidDelData.Value &= e.Node.Parent.Value.Split("$")(0) & ";"
                End If
            End If

            ' 判斷是否有子節點，如果有的話 ，一併取消勾選
            If e.Node.ChildNodes.Count > 0 Then
                For Each item As TreeNode In e.Node.ChildNodes
                    If item.Checked = True Then

                        hidDelData.Value &= item.Value.Split("$")(0) & ";"
                        item.Checked = False
                        item.Text = String.Format("<font color='red'>{0}</font>", item.Text)
                        checkItem(item, False)
                    End If
                Next
            End If
        End If
    End Sub

    ''' <summary>
    ''' 全部取消
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnCancelAll_Click(sender As Object, e As EventArgs) Handles btnCancelAll.Click
        Try
            ' 開始事物
            GetDatabaseManager().beginTran()

            Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager())
            Dim syRelRoleFunctionHis As New AUTH_OP.SY_REL_ROLE_FUNCTION_HIS(GetDatabaseManager())
            Dim flowFacade As New FLOW_OP.FlowFacade(GetDatabaseManager())

            ' 刪除sy_tempinfo
            If syTempInfo.loadByPK(m_sWorkingUserid, m_sWorkingTopDepNo, m_sFuncCode) Then
                syTempInfo.remove()
            End If

            ' 刪除HIS
            syRelRoleFunctionHis.delHisDataByCaseId(m_sCaseId)

            ' 刪除流程資料
            flowFacade.DeleteFlow(m_sCaseId)

            ' 提交
            GetDatabaseManager().commit()

            ViewState.Remove("SY_TEMPINFO")

            treeViewRole.Nodes.Clear()
            treeViewFun.Nodes.Clear()

            ' 正常區域資料加載
            initNormalData()

            treeViewRole.ExpandDepth = m_sTreeLevel
        Catch ex As Exception

            ' 事物回滾
            GetDatabaseManager.Rollback()
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub
#End Region

#Region "審核資料Function"

    ''' <summary>
    ''' 加載審核資料
    ''' </summary>
    ''' <remarks></remarks>
    Sub initCheckData()
        Dim syRole As New AUTH_OP.SY_ROLE(GetDatabaseManager())
        Dim apCodelist As New AP_CODEList(GetDatabaseManager())
        Dim sySysIdList As New AUTH_OP.SY_SYSIDList(GetDatabaseManager())
        Dim syFunctionCode As New AUTH_OP.SY_FUNCTION_CODE(GetDatabaseManager())
        Dim dtRootData As New DataTable
        Dim dtParentData As New DataTable
        Dim ds As New DataSet
        Dim dtTemp As New DataTable

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
                    dtTemp = ds.Tables(0)
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
                    dtTemp = ds.Tables(0)
                End If
            End If
        End If

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
                rootNode.SelectAction = TreeNodeSelectAction.None

                treeViewRoleCheck.Nodes.Add(rootNode)
            Next
        End If

        If Not dtTemp Is Nothing And dtTemp.Rows.Count > 0 Then
            For Each dr As DataRow In dtTemp.Rows
                m_sSysIdList &= "'" & dr("SYSID") & "'" & ","
                m_sRoleIdList &= dr("ROLEID") & ","
                m_sSubSysIdList &= "'" & dr("SUBSYSID") & "'" & ","
                m_sCodeFlag &= dr("FUNCCODE") & "," & dr("OPERATION") & "|"

                If syRole.loadByPK(dr("ROLEID")) Then
                    If syRole.getAttribute("PARENT") <> 0 Then
                        loadNodeByParent(dr("ROLEID"))
                    End If
                End If

            Next

            hidOldRoleList.Value = m_sRoleIdList

            If hidNodeParent.Value <> "" Then
                m_sRoleIdList &= hidNodeParent.Value
            End If
        End If

        ' (2)查詢角色資料的第一層資料：parent=0
        If m_sSysIdList.Length > 0 Then
            m_sSysIdList = m_sSysIdList.Substring(0, m_sSysIdList.Length - 1)
            m_sRoleIdList = m_sRoleIdList.Substring(0, m_sRoleIdList.Length - 1)
            m_sSubSysIdList = m_sSubSysIdList.Substring(0, m_sSubSysIdList.Length - 1)


            If sySysIdList.loadDataBySysId(m_sSysIdList) Then
                dtParentData = sySysIdList.getCurrentDataSet.Tables(0)
            End If
        End If

        ' (3)循環 新增到treeViewRole中
        If Not dtParentData Is Nothing Then
            If dtParentData.Rows.Count > 0 Then
                For Each dr As DataRow In dtParentData.Rows

                    ' 父節點
                    Dim parentNode As TreeNode = New TreeNode()
                    parentNode.Value = dr("SYSID").ToString()
                    parentNode.Text = dr("SYSNAME").ToString & "  (" & dr("SYSID") & ")"

                    parentNode.SelectAction = TreeNodeSelectAction.None

                    ' (4)添加子節點
                    addNodeCheck(parentNode, dr("SYSID").ToString(), m_sRoleIdList, m_sSysIdList, m_sSubSysIdList, dtTemp)

                    ' 添加到treeViewRole中
                    treeViewRoleCheck.Nodes(0).ChildNodes.Add(parentNode)
                Next
            End If
        End If

        Dim drRowTemp() As DataRow = dtTemp.Select(String.Format("ROLEID='{0}' and SYSID='{1}' and SUBSYSID='{2}'", m_sTreeNode.Value.Split("*")(0), m_sTreeNode.Value.Split("*")(1), m_sTreeNode.Value.Split("*")(2)))
        For Each dr As DataRow In drRowTemp
            If dr("FUNCCODE") <> "" Then
                m_sFuncCodeList &= dr("FUNCCODE") & ","
            End If

            If syFunctionCode.loadByPK(dr("FUNCCODE")) Then
                If syFunctionCode.getAttribute("PARENT") <> 0 Then
                    loadNodeByParentF(dr("FUNCCODE"))
                End If
            End If
        Next

        If hidNodeParentF.Value <> "" Then
            m_sFuncCodeList &= hidNodeParentF.Value
        End If

        ViewState("m_sCodeFlag") = m_sCodeFlag

        ' 加載左邊交易的目錄樹
        loadCheckData(m_sTreeNode, dtTemp, m_sFuncCodeList)
    End Sub

    ''' <summary>
    ''' 根據父節點往上推到其父節點-角色
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

    ''' <summary>
    ''' 根據父節點往上推到其父節點-交易
    ''' </summary>
    ''' <param name="sParent"></param>
    ''' <remarks></remarks>
    Sub loadNodeByParentF(ByVal sParent As String)
        Dim syFunctionCode As New AUTH_OP.SY_FUNCTION_CODE(GetDatabaseManager())

        For Each Str As String In sParent.Split(",")
            If syFunctionCode.loadByPK(Str) Then
                If syFunctionCode.getAttribute("PARENT") = "0" Then
                    hidNodeParentF.Value &= syFunctionCode.getAttribute("FUNCCODE").ToString() & ","
                Else
                    loadNodeByParent(syFunctionCode.getAttribute("PARENT"))
                    hidNodeParentF.Value &= syFunctionCode.getAttribute("PARENT").ToString & ","
                End If
            End If
        Next
    End Sub

    ''' <summary>
    ''' 新增treeViewRole節點
    ''' </summary>
    ''' <param name="node">節點</param>
    ''' <param name="item">parent</param>
    ''' <remarks></remarks>
    Private Sub addNodeCheck(ByVal node As TreeNode, ByVal item As String, ByVal sRoleList As String, ByVal sSysIdList As String, ByVal sSubSysId As String, ByVal dtTemp As DataTable)
        Try
            Dim sySubSysIdList As New AUTH_OP.SY_SUBSYSIDList(GetDatabaseManager())
            Dim syTempInfo As New AUTH_OP.SY_TEMPINFO(GetDatabaseManager())
            Dim dtData As New DataTable
            Dim sNewTempTable As New DataTable
            Dim syRoleList As New SY_ROLEList(GetDatabaseManager())
            Dim sTempTable As New DataTable

            ' 根據SYSID取得所有的子系統
            If sySubSysIdList.loadSubSysList(item, sSubSysId) Then
                dtData = sySubSysIdList.getCurrentDataSet.Tables(0)
            End If

            ' 循環添加子節點
            If Not dtData Is Nothing Then
                If dtData.Rows.Count > 0 Then
                    For Each dr As DataRow In dtData.Rows

                        Dim subNode As TreeNode = New TreeNode
                        subNode.Text = dr("SUBSYSNAME").ToString() & "  (" & dr("SUBSYSID").ToString & ")"
                        subNode.Value = dr("SUBSYSID").ToString

                        subNode.SelectAction = TreeNodeSelectAction.None

                        ' 根據登入者，部門編號，系統編號取得資料
                        If sRoleList.Length > 0 Or sSubSysId.Length > 0 Or sSysIdList.Length > 0 Then
                            sNewTempTable = syRoleList.genFunListCheck(m_sWorkingUserid, m_sWorkingTopDepNo, item, sRoleList, sSysIdList, sSubSysId, subNode.Value)
                            sNewTempTable = sNewTempTable.DefaultView.ToTable(True)
                        End If

                        If Not sNewTempTable Is Nothing Then

                            If Not sNewTempTable Is Nothing Then
                                If sNewTempTable.Rows.Count > 0 Then

                                    ' 若有資料，將其新增到子系統下面
                                    addNodesCheck(subNode, 0, 0, sNewTempTable)
                                End If
                            End If

                            ' 只添加存在于臨時檔中的資料
                            If sSubSysId.Contains(subNode.Value) Then

                                ' 證明為正常資料 
                                node.ChildNodes.Add(subNode)
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
    ''' 新增子系統下的角色
    ''' </summary>
    ''' <param name="tNode">節點</param>
    ''' <param name="PId">ParentID</param>
    ''' <param name="Level">Leave</param>
    ''' <param name="dt">數據源</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function addNodesCheck(ByRef subNode As TreeNode, ByVal PId As Integer, ByVal Level As Integer, ByVal dt As DataTable) As String

        '定義DataRow承接DataTable篩選的結果
        Dim rows() As DataRow

        '定義篩選的條件
        Dim filterExpr As String
        filterExpr = "ParentId = " & PId

        '資料篩選並把結果傳入Rows
        rows = dt.Select(filterExpr)

        ' 查詢臨時檔中的資料
        Dim dtTemp As DataTable = loadSYTempInfo()
        Dim dv As System.Data.DataView = dtTemp.DefaultView
        dv.Sort = "SYSID asc,SUBSYSID ASC "
        dtTemp = dv.ToTable

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
                childNode.Text = row("ROLENAME").ToString & "  (" & row("ROLEID").ToString & ")"
                childNode.ToolTip = childNode.Text
                childNode.Value = row("ROLEID").ToString() & "*" & row("SYSID").ToString() & "*" & row("SUBSYSID").ToString()

                ' 若此編號不再臨時檔中查詢出來的角色編號 ，說明其是父節點
                If Not hidOldRoleList.Value.Contains(childNode.Value.Split("*")(0) & ",") Then
                    childNode.SelectAction = TreeNodeSelectAction.None
                    childNode.Text = row("ROLENAME").ToString & "  (" & row("ROLEID").ToString() & ")"
                Else
                    childNode.Text = String.Format("<font color='red'>{0}</font>", row("ROLENAME").ToString & "  (" & row("ROLEID").ToString() & ")")
                End If

                ' 查詢歷史檔中的第一筆數據
                If m_sTreeNode Is Nothing Then
                    If Not dtTemp Is Nothing Then
                        If dtTemp.Rows.Count > 0 Then
                            For Each dr As DataRow In dtTemp.Rows
                                If dr("FUNCCODE") <> "0" Then
                                    If dr("ROLEID") = childNode.Value.Split("*")(0) Then
                                        m_sTreeNode = childNode
                                        Exit For
                                    End If
                                End If
                            Next
                        End If
                    End If
                End If

                '將節點加入Tree中
                If PId = "0" Then
                    If row("SUBSYSID").ToString().Equals(subNode.Value.Split("*")(0)) Then
                        subNode.ChildNodes.Add(childNode)
                    End If
                ElseIf Convert.ToInt32(PId) > 0 Then
                    If row("ParentId").ToString().Equals(subNode.Value.Split("*")(0)) Then
                        If Not subNode.Parent Is Nothing Then
                            If subNode.Parent.Value.IndexOf("*") < 0 Then
                                If row("SUBSYSID").ToString().Equals(subNode.Parent.Value) Then
                                    subNode.ChildNodes.Add(childNode)
                                End If
                            Else
                                If row("SUBSYSID").ToString().Equals(subNode.Parent.Value.Split("*")(2)) Then
                                    subNode.ChildNodes.Add(childNode)
                                End If
                            End If
                        End If
                    End If
                End If


                '呼叫遞回取得子節點
                rc = addNodesCheck(childNode, row("ROLEID"), row("Level"), dt)
            Next
        End If
    End Function

    ''' <summary>
    ''' 綁定交易樹
    ''' </summary>
    ''' <remarks></remarks>
    Sub bindTreeFunCheck(ByVal sFuncCode As String)
        Try
            Dim apCodelist As New AP_CODEList(GetDatabaseManager())
            Dim syFunctionList As New SY_FUNCTION_CODEList(GetDatabaseManager())
            Dim dtRootData As New DataTable
            Dim dtParentData As DataTable
            Dim sParentList As String = String.Empty

            If m_sCodeFlag = "" Then
                m_sCodeFlag = ViewState("m_sCodeFlag").ToString
            End If

            ' (1)綁定安泰銀行
            If apCodelist.loadDataByCodeId(m_sUpCode2366) Then
                dtRootData = apCodelist.getCurrentDataSet.Tables(0)
            End If

            If Not dtRootData Is Nothing Then
                For Each dr As DataRow In dtRootData.Rows

                    ' root節點
                    Dim rootNode As TreeNode = New TreeNode()
                    rootNode.Value = dr("VALUE").ToString()
                    rootNode.Text = dr("TEXT").ToString

                    For Each item As String In m_sCodeFlag.Split("|")
                        If rootNode.Value = item.Split(",")(0) Then
                            If item.Split(",")(1).ToString() = "I" Then
                                rootNode.Checked = True
                                rootNode.Text = String.Format("<font color='red'>{0}</font>", rootNode.Text)
                            ElseIf item.Split(",")(1).ToString() = "N" Then
                                rootNode.Checked = True
                            ElseIf item.Split(",")(1).ToString() = "D" Then
                                rootNode.ImageUrl = "../img/del.gif"
                                rootNode.ShowCheckBox = False
                            End If
                        End If
                    Next

                    treeViewFunCheck.Nodes.Add(rootNode)
                Next
            End If

            If sFuncCode.Length > 0 Then
                sFuncCode = sFuncCode.Substring(0, sFuncCode.Length - 1)
            End If

            ' (2)查詢交易資料的第一層資料：parent=0
            If sFuncCode <> "" Then
                If syFunctionList.loadDataByParentList(sFuncCode) Then
                    dtParentData = syFunctionList.getCurrentDataSet.Tables(0)
                End If
            End If

            ' (3)循環 新增到TreeView中
            If Not dtParentData Is Nothing Then
                Dim rows() As DataRow = dtParentData.Select("[parent]= '0' ")

                For Each dr As DataRow In rows

                    ' 父節點
                    Dim parentNode As TreeNode = New TreeNode()
                    parentNode.Value = dr("FUNCCODE").ToString()
                    parentNode.Text = dr("FUCNAME").ToString

                    For Each item As String In m_sCodeFlag.Split("|")
                        If parentNode.Value = item.Split(",")(0) Then
                            If item.Split(",")(1).ToString() = "I" Then
                                parentNode.Checked = True
                                parentNode.Text = String.Format("<font color='red'>{0}</font>", parentNode.Text)
                            ElseIf item.Split(",")(1).ToString() = "N" Then
                                parentNode.Checked = True
                            ElseIf item.Split(",")(1).ToString() = "D" Then
                                parentNode.ImageUrl = "../img/del.gif"
                                parentNode.ShowCheckBox = False
                            End If
                        End If
                    Next

                    ' (4)添加子節點
                    addNodeFunCheck(parentNode, dr("FUNCCODE").ToString(), sFuncCode)

                    ' 添加到treeView中
                    treeViewFunCheck.Nodes(0).ChildNodes.Add(parentNode)
                Next
            End If
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 添加子節點
    ''' </summary>
    ''' <param name="node"></param>
    ''' <param name="item"></param>
    ''' <remarks></remarks>
    Sub addNodeFunCheck(ByVal node As TreeNode, ByVal item As String, ByVal sFuncCode As String)
        Try
            Dim syFunctionCodeList As New AUTH_OP.SY_FUNCTION_CODEList(GetDatabaseManager())
            Dim sNodeIdList As String = String.Empty

            Dim dtData As New DataTable

            ' 根據Funccode取得所有的子節點
            If syFunctionCodeList.loadDataByParentAndCode(item, sFuncCode) Then
                dtData = syFunctionCodeList.getCurrentDataSet.Tables(0)
            End If

            ' 循環添加子節點
            If Not dtData Is Nothing Then
                If dtData.Rows.Count > 0 Then
                    For Each dr As DataRow In dtData.Rows

                        Dim subNode As TreeNode = New TreeNode
                        subNode.Text = dr("FUCNAME").ToString()
                        subNode.Value = dr("FUNCCODE").ToString

                        For Each itemT As String In m_sCodeFlag.Split("|")
                            If subNode.Value = itemT.Split(",")(0) Then
                                If itemT.Split(",")(1).ToString() = "I" Then
                                    subNode.Checked = True
                                    subNode.Text = String.Format("<font color='red'>{0}</font>", subNode.Text)
                                ElseIf itemT.Split(",")(1).ToString() = "N" Then
                                    subNode.Checked = True
                                ElseIf itemT.Split(",")(1).ToString() = "D" Then
                                    subNode.ImageUrl = "../img/del.gif"
                                    subNode.ShowCheckBox = False
                                End If
                            End If
                        Next

                        addNodeFunCheck(subNode, dr("FUNCCODE").ToString(), sFuncCode)

                        If treeViewFunCheck.Nodes.Count > 0 Then
                            For Each itemT As TreeNode In treeViewFunCheck.Nodes
                                sNodeIdList &= itemT.Value
                            Next
                        End If

                        If Not sNodeIdList.Contains(subNode.Value) Then
                            node.ChildNodes.Add(subNode)
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 更新操作
    ''' </summary>
    ''' <remarks></remarks>
    Sub loadCheckData(ByVal itemNode As TreeNode, ByVal dtData As DataTable, ByVal m_sFuncCodeList As String)
        treeViewFunCheck.Nodes.Clear()

        ' 綁定目錄樹
        bindTreeFunCheck(m_sFuncCodeList)

        treeViewFunCheck.ExpandAll()
        treeViewFunCheck.Enabled = False
    End Sub
#End Region

#Region "審核資料Event"

    ''' <summary>
    ''' 同意
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnArgee_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnArgee.Click
        Dim stepInfo As FLOW_OP.StepInfo

        Try
            GetDatabaseManager.beginTran()

            ' 呼叫流程方法
            stepInfo = MBSC.UICtl.UIShareFun.pushFlow(GetDatabaseManager(), m_sCaseId)

            updatTemp2Live(GetDatabaseManager(), stepInfo, "Y")

            GetDatabaseManager.commit()

            ' 關閉視窗
            com.Azion.EloanUtility.UIUtility.closeWindow()

        Catch ex As Exception
            GetDatabaseManager.Rollback()
            Throw
        End Try


    End Sub

    ''' <summary>
    ''' 不同意
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnNotArgee_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnNotArgee.Click
        Try
            Dim stepInfo As FLOW_OP.StepInfo

            GetDatabaseManager.beginTran()

            ' 呼叫流程方法
            stepInfo = MBSC.UICtl.UIShareFun.pushFlow(GetDatabaseManager(), m_sCaseId)

            updatTemp2Live(GetDatabaseManager(), stepInfo, "N")

            GetDatabaseManager.commit()

            ' 關閉視窗
            com.Azion.EloanUtility.UIUtility.closeWindow()
        Catch ex As Exception
            GetDatabaseManager.Rollback()
            Throw ex
        End Try
    End Sub

    ''' <summary>
    ''' 修正補充
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnReviseFlow_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnReviseFlow.Click
        Try
            GetDatabaseManager.beginTran()

            Dim stepInfo As FLOW_OP.StepInfo

            ' 呼叫流程方法
            stepInfo = MBSC.UICtl.UIShareFun.rollBack(GetDatabaseManager(), m_sCaseId)

            updatTemp2Live(GetDatabaseManager(), stepInfo, "")

            GetDatabaseManager.commit()
        Catch ex As Exception
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
    Protected Sub treeViewRoleCheck_SelectedNodeChanged(ByVal sender As Object, ByVal e As EventArgs) Handles treeViewRoleCheck.SelectedNodeChanged
        Try
            hidRoleId.Value = treeViewRoleCheck.SelectedNode.Value
            hidNodeParentF.Value = ""

            Dim syTempInfo As New SY_TEMPINFO(GetDatabaseManager())
            Dim syFunctionCode As New SY_FUNCTION_CODE(GetDatabaseManager())
            Dim xmlData As String = String.Empty
            Dim ds As New DataSet
            Dim dtTemp As New DataTable

            ' (5) 查詢XML的資料 只有為“”時才按照主鍵進行查詢，否則按鈕caseid進行查詢

            If syTempInfo.loadByCaseId(m_sCaseId) Then
                xmlData = syTempInfo.getString("TEMPDATA")

                If xmlData <> Nothing Then

                    ' (6) document對象載入XML文件
                    Dim xmlDocument As New XmlDocument
                    xmlDocument.LoadXml(xmlData)
                    ds.ReadXml(New XmlNodeReader(xmlDocument))
                    dtTemp = ds.Tables(0)
                End If
            End If

            Dim drRowTemp() As DataRow = dtTemp.Select(String.Format("ROLEID='{0}' and SYSID='{1}' and SUBSYSID='{2}'", hidRoleId.Value.Split("*")(0), hidRoleId.Value.Split("*")(1), hidRoleId.Value.Split("*")(2)))
            For Each dr As DataRow In drRowTemp
                If dr("FUNCCODE").ToString() <> "" Then
                    m_sFuncCodeList &= dr("FUNCCODE") & ","

                    If syFunctionCode.loadByPK(dr("FUNCCODE")) Then
                        If syFunctionCode.getAttribute("PARENT") <> 0 Then
                            loadNodeByParentF(dr("FUNCCODE"))
                        End If
                    End If
                End If
            Next

            If hidNodeParentF.Value <> "" Then
                m_sFuncCodeList &= hidNodeParentF.Value
            End If

            ' 加載資料
            loadCheckData(treeViewRoleCheck.SelectedNode, dtTemp, m_sFuncCodeList)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

End Class