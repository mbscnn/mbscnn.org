Imports com.Azion.NET.VB
Imports MBSC.MB_OP
Imports MBSC.UICtl

Public Class MBMnt_ItemCode_01_v01
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents ddlSubSys As System.Web.UI.WebControls.DropDownList
    Protected WithEvents tbSearch As System.Web.UI.WebControls.TextBox
    Protected WithEvents btnSearch As System.Web.UI.WebControls.Button
    Protected WithEvents ddlItemClass0 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents ddlItemClass1 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents ddlItemClass2 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents ddlItemClass3 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents ddlItemClass4 As System.Web.UI.WebControls.DropDownList
    Protected WithEvents imgbtnAdd0 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgbtnEdit0 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgbtnDel0 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgbtnAdd1 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgbtnEdit1 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgbtnDel1 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgbtnAdd2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgbtnEdit2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgbtnDel2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgbtnAdd3 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgbtnEdit3 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgbtnDel3 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgbtnAdd4 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgbtnEdit4 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgbtnDel4 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents lblID As System.Web.UI.WebControls.Label
    Protected WithEvents tbValue As System.Web.UI.WebControls.TextBox
    Protected WithEvents tbText As System.Web.UI.WebControls.TextBox
    Protected WithEvents tbSeq As System.Web.UI.WebControls.TextBox
    Protected WithEvents tbPs As System.Web.UI.WebControls.TextBox
    Protected WithEvents ddlVisible As System.Web.UI.WebControls.DropDownList
    Protected WithEvents btnSure As System.Web.UI.WebControls.Button
    Protected WithEvents btnCancel As System.Web.UI.WebControls.Button
    Protected WithEvents divEdit As System.Web.UI.HtmlControls.HtmlGenericControl
    Protected WithEvents hidCurrentUpCode As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents hidCurrentLevel As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents hidSearchCodeID As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents divNormal As System.Web.UI.HtmlControls.HtmlGenericControl
    Protected WithEvents divPopUp As System.Web.UI.HtmlControls.HtmlGenericControl
    Protected WithEvents dg As System.Web.UI.WebControls.DataGrid
    Protected WithEvents tdLevel0 As System.Web.UI.HtmlControls.HtmlTableCell
    Protected WithEvents tdLevel1 As System.Web.UI.HtmlControls.HtmlTableCell
    Protected WithEvents tdLevel2 As System.Web.UI.HtmlControls.HtmlTableCell
    Protected WithEvents tdLevel3 As System.Web.UI.HtmlControls.HtmlTableCell
    Protected WithEvents tdLevel4 As System.Web.UI.HtmlControls.HtmlTableCell
    Protected WithEvents lblUPCODE As System.Web.UI.WebControls.Label
    Protected WithEvents DDL_LEVEL As System.Web.UI.WebControls.DropDownList
    Protected WithEvents DDL_BRID As System.Web.UI.WebControls.DropDownList

    '注意: 下列預留位置宣告是 Web Form 設計工具需要的項目。
    '請勿刪除或移動它。
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: 此為 Web Form 設計工具所需的方法呼叫
        '請勿使用程式碼編輯器進行修改。
        InitializeComponent()
    End Sub

#End Region

    Private m_htSubCodes As Hashtable
    Private m_dbManager As DatabaseManager

    Dim m_AL_BRID As New ArrayList

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化
        Dim bSearchMode As Boolean
        m_dbManager = UIShareFun.getDataBaseManager
        bSearchMode = Utility.isValidateData(Request.QueryString("AdvSearch"))

        Try
            UIShareFun.closeErrMsg()
            If bSearchMode Then 'Pop Up Mode
                Dim apCodes As New AP_CODEList(m_dbManager)
                Dim sQueryText As String
                Dim dt As DataTable
                divPopUp.Style.Item("display") = ""
                divNormal.Style.Item("display") = "none"

                sQueryText = Request.QueryString("QueryText")
                If Not Utility.isValidateData(sQueryText) Then
                    UIShareFun.showErrMsg(Me, "查無資料")
                    Exit Sub
                End If

                If apCodes.loadForAdvanceSearch(sQueryText) > 0 Then
                    dt = apCodes.getCurrentDataSet.Tables(0)
                    dg.DataSource = dt
                    dg.DataBind()
                Else
                    divPopUp.InnerHtml = "<br><br><div style='text-align:center;color:red;font-size:15px'>無資料</div>"
                End If
            Else 'Normal Mode
                divNormal.Style.Item("display") = ""
                divPopUp.Style.Item("display") = "none"
                initSubSys()
            End If

            If Not Page.IsPostBack Then
                '中心別
                Me.DDL_BRID.Items.Clear()
                Dim AP_CODEList As New AP_CODEList(Me.m_dbManager)
                AP_CODEList.loadByUpCodeOrderBySortno("177")
                For Each ROW As DataRow In AP_CODEList.getCurrentDataSet.Tables(0).Rows
                    Dim ListItem As New ListItem(ROW("TEXT").ToString, ROW("VALUE").ToString)
                    Me.DDL_BRID.Items.Add(ListItem)
                Next

                Me.DDL_BRID.Items.Insert(0, New ListItem("總會", "0"))
                Me.DDL_BRID.Items.Insert(0, New ListItem("請選擇", String.Empty))
            End If
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    Private Sub Page_Unload(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Unload
        UIShareFun.releaseConnection(m_dbManager)
    End Sub

    '初始化子系統下拉選單
    Private Sub initSubSys()
        Dim subSys As SUBSYS_CODEList
        If Me.ddlSubSys.Items.Count < 1 Then
            Try
                ddlSubSys.Items.Add(New ListItem("<請選擇>", "-1"))
                subSys = New SUBSYS_CODEList(m_dbManager)
                subSys.setSQLCondition(" ORDER BY SUBSYSID")
                subSys.loadAllData()
                For i As Integer = 0 To subSys.size - 1
                    ddlSubSys.Items.Add(New ListItem(subSys.item(i).getString("SHORTSSNAME"), subSys.item(i).getString("SUBSYSID")))
                Next
                ddlSubSys.SelectedIndex = 0
            Catch ex As Exception
                Throw
            End Try
        End If
    End Sub

    '顯示到哪個層級    Private Sub DisplayLevel(ByVal iLevel As Integer)
        Dim ddl As DropDownList
        Dim Add As ImageButton
        Dim Edit As ImageButton
        Dim Del As ImageButton
        Dim bVisible As Boolean

        For i As Integer = 0 To 4
            bVisible = (i <= iLevel) '顯示與否
            ddl = getObjectLevel(i, ObjectType.ItemClass)
            Add = getObjectLevel(i, ObjectType.Add)
            Edit = getObjectLevel(i, ObjectType.Edit)
            Del = getObjectLevel(i, ObjectType.Del)
            If Utility.isValidateData(ddl) Then ddl.Visible = bVisible
            If Utility.isValidateData(Add) Then Add.Visible = bVisible
            If Utility.isValidateData(Edit) Then Edit.Visible = bVisible
            If Utility.isValidateData(Del) Then Del.Visible = bVisible
        Next
    End Sub

    '是否顯示編輯區
    Private Sub showModifyArea(ByVal bShow As Boolean, Optional ByVal ddlItemClass As DropDownList = Nothing, Optional ByVal bAdd As Boolean = False)
        If bShow Then '顯示
            divEdit.Style.Item("display") = ""
            Try
                Dim apCode As New AP_CODE(Me.m_dbManager)
                Dim iNext = CInt(ddlItemClass.ID.Replace("ddlItemClass", ""))
                Dim sCodeID As String = ddlItemClass.SelectedValue
                Dim td As System.Web.UI.HtmlControls.HtmlTableCell

                hidCurrentLevel.Value = iNext
                If Not bAdd Then '修改
                    If apCode.loadByPK(sCodeID) Then
                        lblID.Text = sCodeID
                        tbPs.Text = apCode.getString("NOTE")
                        tbSeq.Text = apCode.getString("SORTNO")
                        tbValue.Text = apCode.getString("VALUE")
                        tbText.Text = apCode.getString("TEXT")
                        tbText.Attributes.Item("OrigText") = tbText.Text '儲存時，用來檢核，是否有修改過。新增則沒有
                        ddlVisible.SelectedValue = apCode.getString("DISABLED")

                        lblUPCODE.Text = apCode.getString("UPCODE")
                        hidCurrentUpCode.Value = apCode.getString("UPCODE")
                        setSeriesByCodeID(ddlItemClass.SelectedValue)

                        Me.DDL_LEVEL.SelectedIndex = -1
                        If Not IsNothing(Me.DDL_LEVEL.Items.FindByValue(apCode.getString("LEVEL"))) Then
                            Me.DDL_LEVEL.Items.FindByValue(apCode.getString("LEVEL")).Selected = True
                        End If

                        Me.DDL_BRID.SelectedIndex = -1
                        If Not IsNothing(Me.DDL_BRID.Items.FindByValue(apCode.getString("BRID"))) Then
                            Me.DDL_BRID.Items.FindByValue(apCode.getString("BRID")).Selected = True
                        End If
                    End If
                Else '新增
                    lblID.Text = CStr(apCode.getMaxCodeID + 1)
                    lblUPCODE.Text = ddlItemClass.Attributes.Item("UpCode")
                    hidCurrentUpCode.Value = ddlItemClass.Attributes.Item("UpCode")
                    If Not ddlItemClass.SelectedValue.Equals("-1") Then
                        setSeriesByCodeID(ddlItemClass.Attributes.Item("UpCode"))
                    End If
                End If



                '設定反白
                ddlItemClass.BackColor = System.Drawing.Color.Gold
                td = FindControl("tdLevel" & CStr(iNext))
                If Utility.isValidateData(td) Then td.Style.Item("background-color") = "Gold"
            Catch ex As Exception
                Throw
            End Try
        Else '不顯示
            divEdit.Style.Item("display") = "none"
            lblID.Text = ""
            tbPs.Text = ""
            tbSeq.Text = ""
            tbValue.Text = ""
            tbText.Text = ""
            ddlVisible.SelectedValue = "1"
            '反白取消
            tdLevel0.Style.Item("background-color") = ""
            tdLevel1.Style.Item("background-color") = ""
            tdLevel2.Style.Item("background-color") = ""
            tdLevel3.Style.Item("background-color") = ""
            tdLevel4.Style.Item("background-color") = ""
            ddlItemClass0.BackColor = System.Drawing.Color.White
            ddlItemClass1.BackColor = System.Drawing.Color.White
            ddlItemClass2.BackColor = System.Drawing.Color.White
            ddlItemClass3.BackColor = System.Drawing.Color.White
            ddlItemClass4.BackColor = System.Drawing.Color.White

            Me.DDL_LEVEL.SelectedIndex = -1
            Me.DDL_BRID.SelectedIndex = -1
        End If
    End Sub

    '初始化項目類別
    Private Sub initItemClass(ByVal sUpCode As String, ByVal iLevel As Integer)
        Dim apCodes As New AP_CODEList(m_dbManager)
        Dim ddlItemClass As DropDownList = getObjectLevel(iLevel, ObjectType.ItemClass)

        Try
            If Utility.isValidateData(ddlItemClass) Then
                ddlItemClass.Items.Clear()
                ddlItemClass.Items.Add(New ListItem("<請選擇>", "-1"))
                apCodes.loadByUpCode(sUpCode, Me.ddlSubSys.SelectedValue)
                For i As Integer = 0 To apCodes.size - 1
                    ddlItemClass.Items.Add(New ListItem(apCodes.item(i).getString("TEXT"), apCodes.item(i).getString("CODEID")))
                Next
                ddlItemClass.Attributes.Item("UpCode") = sUpCode
                DisplayLevel(iLevel)
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub

    '將sCodeID從UpCode=0，設定下拉選單，到自已本身
    Private Function setSeriesByCodeID(ByVal sCodeID As String) As String
        Dim sUpCode As String = sCodeID
        Dim sbUpCodeList As New System.Text.StringBuilder
        Dim aryUpCodeList As String()
        Dim sSysSub As String = "" '搜尋時SubSysID不確定
        Try
            '取得sCodeID到Level 0
            If sUpCode.Equals("0") Then '如果sCodeID一開始就是0，那表示不是搜尋模式，而是新增/編輯模式，可以取得ddlSubSys.SelectedValue
                initItemClass("0", 0)
            Else
                While Not sUpCode.Equals("0")
                    Dim apCode As New AP_CODE(Me.m_dbManager)
                    If apCode.loadByPK(sUpCode) Then
                        If sSysSub.Equals("") Then sSysSub = apCode.getString("SUBSYSID")
                        sUpCode = apCode.getString("UPCODE")
                        sbUpCodeList.Append(sUpCode & ",")
                        If sbUpCodeList.Length > 100 Then Exit While
                    Else
                        Exit While
                    End If
                End While

                aryUpCodeList = sbUpCodeList.ToString.TrimEnd(","c).Split(","c)
                Array.Reverse(aryUpCodeList)
                initSubSys()
                ddlSubSys.SelectedValue = sSysSub

                Dim i As Integer
                Dim ddl As DropDownList
                Dim iUpperIndex As Integer = aryUpCodeList.Length - 1
                '從第0層DropDownList開始初始化
                For i = 0 To 4
                    If i > iUpperIndex Then Exit For
                    Dim sEachOfCodeID As String
                    initItemClass(aryUpCodeList(i), i)
                    '將DropDownList停在下一層的父類別上面
                    ddl = CType(getObjectLevel(i, ObjectType.ItemClass), DropDownList)
                    If i < iUpperIndex Then
                        sEachOfCodeID = CStr(aryUpCodeList(i + 1))
                    Else
                        sEachOfCodeID = sCodeID
                    End If
                    If Utility.isValidateData(ddl.Items.FindByValue(sEachOfCodeID)) Then ddl.SelectedValue = sEachOfCodeID
                    setHasSon(ddl) '設定有無子項目
                Next

                '設定最後一個ItemClass，顯示請選擇
                If i <= 4 Then
                    initItemClass(sCodeID, i)
                End If
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function

    '取得子系統名稱
    Public Function getSubSysName(ByVal _SubSysID As String) As String
        Try
            If Not Utility.isValidateData(m_htSubCodes) OrElse m_htSubCodes.Count < 1 Then
                Dim m_subCodes As New SUBSYS_CODEList(m_dbManager)
                m_htSubCodes = New Hashtable
                m_subCodes.loadAllData()
                For i As Integer = 0 To m_subCodes.size - 1
                    m_htSubCodes.Add(m_subCodes.item(i).getString("SUBSYSID"), m_subCodes.item(i).getString("SHORTSSNAME"))
                Next
            End If

            Return CStr(m_htSubCodes.Item(_SubSysID))
        Catch ex As Exception
            Throw
        End Try
    End Function

    Private Sub dg_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles dg.ItemDataBound
        Dim apCode As New AP_CODE(Me.m_dbManager)
        If e.Item.ItemType = ListItemType.AlternatingItem OrElse e.Item.ItemType = ListItemType.Item Then
            Try '將UpCode轉成中文名稱
                apCode.loadByPK(e.Item.DataItem("UPCODE"))
                e.Item.Cells(2).Text = apCode.getString("TEXT")
            Catch ex As Exception
                Throw
            End Try
        End If
    End Sub

    '取得指定Level的物件(用select case而不用FindControl來提昇效能)
    Private Enum ObjectType
        Edit
        Add
        Del
        ItemClass
    End Enum
    Private Function getObjectLevel(ByVal iLevel As Integer, ByVal type As ObjectType) As Object
        Select Case type
            Case ObjectType.Add 'imgbtnAddX
                Select Case iLevel
                    Case 0
                        Return imgbtnAdd0
                    Case 1
                        Return imgbtnAdd1
                    Case 2
                        Return imgbtnAdd2
                    Case 3
                        Return imgbtnAdd3
                    Case 4
                        Return imgbtnAdd4
                End Select
            Case ObjectType.Del 'imgbtnDelX
                Select Case iLevel
                    Case 0
                        Return imgbtnDel0
                    Case 1
                        Return imgbtnDel1
                    Case 2
                        Return imgbtnDel2
                    Case 3
                        Return imgbtnDel3
                    Case 4
                        Return imgbtnDel4
                End Select
            Case ObjectType.Edit 'imgbtnEditX
                Select Case iLevel
                    Case 0
                        Return imgbtnEdit0
                    Case 1
                        Return imgbtnEdit1
                    Case 2
                        Return imgbtnEdit2
                    Case 3
                        Return imgbtnEdit3
                    Case 4
                        Return imgbtnEdit4
                End Select
            Case ObjectType.ItemClass 'ddlItemClassX
                Select Case iLevel
                    Case 0
                        Return ddlItemClass0
                    Case 1
                        Return ddlItemClass1
                    Case 2
                        Return ddlItemClass2
                    Case 3
                        Return ddlItemClass3
                    Case 4
                        Return ddlItemClass4
                End Select
        End Select
    End Function

#Region "Event"
    'on SubSys DropDownList be selected
    Private Sub ddlSubSys_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlSubSys.SelectedIndexChanged
        Try
            If Not ddlSubSys.SelectedValue.Equals("-1") Then
                initItemClass("0", 0)
            Else
                Me.DisplayLevel(-1)
            End If
            showModifyArea(False)
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    'on SelectedIndexChanged 所有的項目類別
    Private Sub setItemClass(ByVal sender As Object, ByVal e As System.EventArgs) Handles _
    ddlItemClass0.SelectedIndexChanged, _
    ddlItemClass1.SelectedIndexChanged, _
    ddlItemClass2.SelectedIndexChanged, _
    ddlItemClass3.SelectedIndexChanged
        Dim ddlItemClass As DropDownList
        Dim clickItemClass As DropDownList = CType(sender, DropDownList)
        Dim iNext As Integer

        Try
            iNext = CInt(CType(sender, DropDownList).ID.Replace("ddlItemClass", ""))
            If Not clickItemClass.SelectedValue.Equals("-1") Then
                iNext += 1
                initItemClass(clickItemClass.SelectedValue, iNext)
                setHasSon(clickItemClass) '設定有無子項目
            Else
                DisplayLevel(iNext)
            End If
            showModifyArea(False)

            '中心別
            Me.DDL_BRID.SelectedIndex = -1
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    '設定該下拉選單目前的項目有無子項目
    Private Sub setHasSon(ByRef ddl As DropDownList)
        Dim apCodes As New AP_CODEList(Me.m_dbManager)
        Try
            If Not IsNothing(ddl) Then
                If apCodes.loadByUpCode(ddl.SelectedValue, Me.ddlSubSys.SelectedValue) > 0 Then
                    ddl.Attributes.Add("hasSon", "true")
                Else
                    ddl.Attributes.Add("hasSon", "false")
                End If
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub

    'on Add ImageButton Click
    Private Sub onAddClick(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles _
    imgbtnAdd0.Click, _
    imgbtnAdd1.Click, _
    imgbtnAdd2.Click, _
    imgbtnAdd3.Click, _
    imgbtnAdd4.Click
        Dim iCurr As Integer
        Dim ddlItemClass As DropDownList
        Try
            iCurr = CInt(CType(sender, ImageButton).ID.Replace("imgbtnAdd", ""))
            ddlItemClass = getObjectLevel(iCurr, ObjectType.ItemClass)

            If Utility.isValidateData(ddlItemClass) Then
                Me.showModifyArea(True, ddlItemClass, True)
            End If
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    'on Edit ImageButton Click
    Private Sub onEditClick(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles _
    imgbtnEdit0.Click, _
    imgbtnEdit1.Click, _
    imgbtnEdit2.Click, _
    imgbtnEdit3.Click, _
    imgbtnEdit4.Click
        Dim iCurr As Integer
        Dim ddlItemClass As DropDownList

        Try
            iCurr = CInt(CType(sender, ImageButton).ID.Replace("imgbtnEdit", ""))
            ddlItemClass = getObjectLevel(iCurr, ObjectType.ItemClass)

            If Utility.isValidateData(ddlItemClass) Then
                Me.showModifyArea(True, ddlItemClass)
            End If
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    'on Delete ImageButton Click
    Private Sub onDelClick(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles _
    imgbtnDel0.Click, _
    imgbtnDel1.Click, _
    imgbtnDel2.Click, _
    imgbtnDel3.Click, _
    imgbtnDel4.Click
        Dim iCurr As Integer
        Dim ddlItemClass As DropDownList
        Dim apCode As AP_CODE
        Dim sUpCode As String
        Try
            iCurr = CInt(CType(sender, ImageButton).ID.Replace("imgbtnDel", ""))
            ddlItemClass = getObjectLevel(iCurr, ObjectType.ItemClass)

            If Utility.isValidateData(ddlItemClass) Then
                m_dbManager.beginTran()
                apCode = New AP_CODE(m_dbManager)

                If apCode.loadByPkForUpdate(ddlItemClass.SelectedValue) Then
                    sUpCode = apCode.getString("UPCODE")
                    apCode.remove()
                End If
                m_dbManager.commit()
                UIShareFun.showErrMsg(Me, "刪除成功")

                initItemClass(sUpCode, iCurr)
                showModifyArea(False)

                '設定上一層有無子項目
                ddlItemClass = getObjectLevel(iCurr - 1, ObjectType.ItemClass)
                Me.setHasSon(ddlItemClass)
            End If
        Catch ex As Exception
            m_dbManager.Rollback()
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    '取消
    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            Me.showModifyArea(False)
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    '確定修改 / 新增
    Private Sub btnSure_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSure.Click
        Dim apCode As AP_CODE
        Dim sSubSysID As String = Me.ddlSubSys.SelectedValue
        Dim iLevelEdited As Integer = CInt(hidCurrentLevel.Value)
        Dim sOrigText As String = tbText.Attributes.Item("OrigText")
        Dim bHasEdited As Boolean
        Try
            If Not Utility.isValidateData(tbText.Text) Then
                UIShareFun.showErrMsg(Me, "[文字]不可為空白")
                Exit Sub
            End If

            If hidCurrentUpCode.Value = "76" Then
                '維護人員
                If Not Utility.isValidateData(Me.DDL_BRID.SelectedValue) Then
                    com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "請選擇中心別")
                    com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請選擇中心別")
                    Return
                End If
            End If

            m_dbManager.beginTran()
            apCode = New AP_CODE(m_dbManager)

            '如果存在原來的字串(是編輯模式，不是新增模式)並且和tbText.Text不相同，才可以檢核Ap_Code中是否有相同資料
            If Not Utility.isValidateData(sOrigText) OrElse (Utility.isValidateData(sOrigText) AndAlso Not sOrigText.Equals(tbText.Text.Trim)) Then
                If apCode.loadByText(ddlSubSys.SelectedValue, hidCurrentUpCode.Value, tbText.Text) Then
                    UIShareFun.showErrMsg(Me, "文字[" & tbText.Text & "]已存在，請重新輸入")
                    m_dbManager.Rollback()
                    Exit Sub
                End If
            End If

            If Not apCode.loadByPkForUpdate(lblID.Text) Then apCode.setAttribute("CODEID", lblID.Text)
            apCode.setAttribute("SUBSYSID", sSubSysID)
            apCode.setAttribute("VALUE", tbValue.Text)
            apCode.setAttribute("TEXT", tbText.Text)
            apCode.setAttribute("NOTE", tbPs.Text)
            apCode.setAttribute("UPCODE", hidCurrentUpCode.Value)
            apCode.setAttribute("SORTNO", tbSeq.Text)
            apCode.setAttribute("DISABLED", ddlVisible.SelectedValue)
            apCode.setAttribute("CHGUID", UIShareFun.getLoginUserid)
            apCode.setAttribute("CHGDATE", Now)
            apCode.setAttribute("LEVEL", CInt(Me.DDL_LEVEL.SelectedValue))
            apCode.setAttribute("BRID", Me.DDL_BRID.SelectedValue)
            apCode.save()
            m_dbManager.commit()
            UIShareFun.showErrMsg(Me, "儲存成功")

            initItemClass(Me.hidCurrentUpCode.Value, iLevelEdited)
            initItemClass(lblID.Text, iLevelEdited + 1)
            CType(getObjectLevel(iLevelEdited, ObjectType.ItemClass), DropDownList).SelectedValue = lblID.Text

            ' '設定上一層有無子項目
            Dim ddl As DropDownList = CType(getObjectLevel(iLevelEdited - 1, ObjectType.ItemClass), DropDownList)
            Me.setHasSon(ddl)

            showModifyArea(False)
        Catch ex As Exception
            m_dbManager.Rollback()
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    '搜尋
    Private Sub btnSearch_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        setSeriesByCodeID(hidSearchCodeID.Value)
    End Sub
#End Region
End Class
