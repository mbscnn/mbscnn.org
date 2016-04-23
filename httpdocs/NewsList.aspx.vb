Imports MBSC.MB_OP
Imports MBSC.UICtl
Imports System.IO

Public Class NewsList
    Inherits System.Web.UI.Page

    Dim m_DBManager As com.Azion.NET.VB.DatabaseManager

    Dim m_sCODEID As String = String.Empty

    Dim m_iPageSize As Integer = 20

    Dim m_iUPCODE_LV As Integer = 37

    Dim m_sNEWSODEID As String = String.Empty

    Dim m_sCRETIME As String = String.Empty

    '測試GitHub變更

#Region "Page Events"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            'Clear Cache
            Response.Expires = 0
            Response.Cache.SetNoStore()
            Response.AppendHeader("Pragma", "no-cache")

            Me.m_sCODEID = "" & Request.QueryString("CLASS")

            Me.m_sCRETIME = "" & Request.QueryString("CRETIME")

            Me.m_DBManager = UIShareFun.getDataBaseManager

            'If Not Utility.isValidateData(Me.m_sCODEID) Then
            '    Dim apCODEList As New AP_CODEList(Me.m_DBManager)
            '    apCODEList.setSQLCondition(" ORDER BY SORTNO ")
            '    apCODEList.loadByUpCode(Me.m_iUPCODE_LV)
            '    Dim ROW_LV1() As DataRow = apCODEList.getCurrentDataSet.Tables(0).Select("DISABLED='1'")
            '    If Not IsNothing(ROW_LV1) AndAlso ROW_LV1.Length > 0 Then
            '        Me.m_sCODEID = ROW_LV1(0)("CODEID")
            '    End If
            'End If
            Dim apCODEList As New AP_CODEList(Me.m_DBManager)
            apCODEList.setSQLCondition(" ORDER BY SORTNO ")
            apCODEList.loadByUpCode(Me.m_iUPCODE_LV)
            'For i As Integer = 0 To apCODEList.size - 1
            '    Response.Write(apCODEList.item(i).getString("TEXT") & "<BR/>")
            'Next

            Dim ROW_LV1() As DataRow = apCODEList.getCurrentDataSet.Tables(0).Select("DISABLED='1'")
            If Not IsNothing(ROW_LV1) AndAlso ROW_LV1.Length > 0 Then
                Me.m_sNEWSODEID = ROW_LV1(0)("CODEID")
            End If

            If Not Page.IsPostBack Then
                Dim sYear As String = String.Empty
                Dim MB_YENT As New MB_YENT(Me.m_DBManager)
                sYear = MB_YENT.getMAX_SDATE
                Me.LTL_YEAR_ENT.Text = sYear
                Me.A_YEAR_ENT.Attributes("href") = com.Azion.EloanUtility.UIUtility.getRootPath & "/QRY/MBQry_Year_ENT.aspx?YEAR=" & sYear

                Dim DT_FILES As New DataTable
                Try
                    DT_FILES.Columns.Add("NAME", Type.GetType("System.String"))
                    DT_FILES.Columns.Add("LASTWRITETIME", Type.GetType("System.DateTime"))

                    Dim dirInfo As New DirectoryInfo(Server.MapPath("~/ckfinder/userfiles/images/banner/"))
                    Dim objFilwinfo() As FileInfo = dirInfo.GetFiles
                    For Each FileInfo As FileInfo In objFilwinfo
                        Dim sqlRow As DataRow = DT_FILES.NewRow
                        sqlRow("NAME") = FileInfo.Name
                        sqlRow("LASTWRITETIME") = FileInfo.LastWriteTime
                        DT_FILES.Rows.Add(sqlRow)
                    Next

                    Me.RP_Banner.DataSource = New DataView(DT_FILES, String.Empty, "LASTWRITETIME DESC", DataViewRowState.CurrentRows).ToTable
                    Me.RP_Banner.DataBind()
                Catch ex As Exception
                    Throw
                Finally
                    If Not IsNothing(DT_FILES) Then
                        DT_FILES.Dispose()
                    End If
                End Try

                '報名中
                'Me.IMG_SET.ImageUrl = com.Azion.EloanUtility.UIUtility.getRootPath & "/ckfinder/userfiles/images/Settlement.png"
                Me.Bind_RP_CLASS_1()

                '進行中課程
                Me.Bind_RP_CLASS_4()

                '課程活動預告
                'Me.IMG_PRE.ImageUrl = com.Azion.EloanUtility.UIUtility.getRootPath & "/ckfinder/userfiles/images/Design.png"
                Me.Bind_RP_CLASS_2()

                If IsNumeric(Session("admin")) AndAlso CInt(Session("admin")) >= 5 Then
                    'Me.btAdd.Visible = True
                    Me.IMG_ADD.Visible = True
                End If

                Me.IMG_FST.ImageUrl = com.Azion.EloanUtility.UIUtility.getImgPath & "FST.png"
                Me.IMG_LEFT.ImageUrl = com.Azion.EloanUtility.UIUtility.getImgPath & "LEFT.png"

                Me.IMG_RIGHT.ImageUrl = com.Azion.EloanUtility.UIUtility.getImgPath & "RIGHT.png"
                Me.IMG_LST.ImageUrl = com.Azion.EloanUtility.UIUtility.getImgPath & "LST.png"

                Dim mbNEWS As New MB_NEWS(Me.m_DBManager)

                If IsNumeric(Me.m_sCRETIME) Then
                    Dim mbNEWSList As New MB_NEWSList(Me.m_DBManager)
                    mbNEWSList.LoadByCRETIME(Me.m_sCRETIME)
                    Me.RP_NEWS.DataSource = mbNEWSList.getCurrentDataSet.Tables(0)
                    Me.RP_NEWS.DataBind()

                    '隱藏課程清單
                    Me.PLH_CLASS_LIST.Visible = False
                Else
                    If Me.m_sCODEID = Me.m_sNEWSODEID OrElse Not IsNumeric(Me.m_sCODEID) Then
                        Dim D_NOW As Date = Now
                        Me.Bind_RP_NEWSByTime(Me.getNEW_SEQTIME, 1)
                    Else
                        Me.Bind_RP_NEWS(CInt(Me.m_sCODEID), mbNEWS.getMAX_SEQTIME(CInt(Me.m_sCODEID)), 1)
                    End If
                End If

                '選擇課程
                'Dim sURLClass As String = String.Empty
                'sURLClass = com.Azion.EloanUtility.UIUtility.getRootPath & "/MNT/MBMnt_Class_01_v01.aspx?OPTYPE=Q"
                'Me.btnChoose.Attributes("onclick") = "chooseClass('" & sURLClass & "');"

                Me.IMG_ADD.ImageUrl = com.Azion.EloanUtility.UIUtility.getImgPath & "Pencil.png"
            End If
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

    Private Sub NewsList_Unload(sender As Object, e As System.EventArgs) Handles Me.Unload
        UIShareFun.releaseConnection(Me.m_DBManager)
    End Sub
#End Region

    Sub Bind_RP_CLASS_1()
        Try
            Dim MB_CLASSList As New MB_CLASSList(Me.m_DBManager)
            MB_CLASSList.LoadAPLY_G()
            Me.RP_CLASS_1.DataSource = MB_CLASSList.getCurrentDataSet.Tables(0)
            Me.RP_CLASS_1.DataBind()
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Sub Bind_RP_CLASS_4()
        Try
            Dim MB_CLASSList As New MB_CLASSList(Me.m_DBManager)
            MB_CLASSList.LoadAPLY_4_G()
            Me.RP_CLASS_4.DataSource = MB_CLASSList.getCurrentDataSet.Tables(0)
            Me.RP_CLASS_4.DataBind()
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Sub Bind_RP_CLASS_2()
        Try
            Dim MB_CLASSList As New MB_CLASSList(Me.m_DBManager)
            MB_CLASSList.LoadAPLY_2_G()
            Me.RP_CLASS_2.DataSource = MB_CLASSList.getCurrentDataSet.Tables(0)
            Me.RP_CLASS_2.DataBind()
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Sub Bind_RP_NEWS(ByVal iCODEID As Integer, ByVal iSEQTIME As Decimal, ByVal iMode As Integer)
        Try
            Dim mbNEWSList As New MB_NEWSList(Me.m_DBManager)
            If iMode = 1 Then
                mbNEWSList.LoadByCODEID_PREV(iCODEID, iSEQTIME, Me.m_iPageSize)
            Else
                mbNEWSList.LoadByCODEID_NEXT(iCODEID, iSEQTIME, Me.m_iPageSize)
            End If

            If mbNEWSList.getCurrentDataSet.Tables(0).Rows.Count > 0 Then
                Me.RP_NEWS.DataSource = mbNEWSList.getCurrentDataSet.Tables(0)
                Me.RP_NEWS.DataBind()
            Else
                Me.RP_NEWS.DataSource = Nothing
                Me.RP_NEWS.DataBind()
            End If

            Me.RP_NEWS.Visible = True
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Sub Bind_RP_NEWSByTime(ByVal iSEQTIME As Decimal, ByVal iMode As Integer)
        Try
            Dim mbNEWSList As New MB_NEWSList(Me.m_DBManager)
            If iMode = 1 Then
                mbNEWSList.LoadByTIME_PREV(iSEQTIME, Me.m_iPageSize)
            Else
                mbNEWSList.LoadByTIME_NEXT_ASC(iSEQTIME, Me.m_iPageSize)
            End If

            If mbNEWSList.getCurrentDataSet.Tables(0).Rows.Count > 0 Then
                Me.RP_NEWS.DataSource = mbNEWSList.getCurrentDataSet.Tables(0)
                Me.RP_NEWS.DataBind()
            End If

            Me.RP_NEWS.Visible = True
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Private Sub RP_NEWS_ItemDataBound(sender As Object, e As System.Web.UI.WebControls.RepeaterItemEventArgs) Handles RP_NEWS.ItemDataBound
        Try
            If e.Item.ItemType = ListItemType.Item OrElse e.Item.ItemType = ListItemType.AlternatingItem Then
                Dim DRV As DataRowView = CType(e.Item.DataItem, DataRowView)
                Dim HID_CODEID As HtmlInputHidden = e.Item.FindControl("HID_CODEID")
                Dim HID_BRID As HtmlInputHidden = e.Item.FindControl("HID_BRID")
                Dim apCODE As New AP_CODE(Me.m_DBManager)
                If apCODE.loadByPK(HID_CODEID.Value) Then
                    Dim LTL_TITLE As Literal = e.Item.FindControl("LTL_TITLE")
                    LTL_TITLE.Text = apCODE.getString("TEXT")

                    HID_BRID.Value = apCODE.getString("BRID")
                End If

                Dim LTL_CNTHTML As Literal = e.Item.FindControl("LTL_CNTHTML")
                Dim LTL_CNTHTML_MORE As Literal = e.Item.FindControl("LTL_CNTHTML_MORE")
                Dim DIV_MORE As HtmlGenericControl = e.Item.FindControl("DIV_MORE")
                If Utility.isValidateData(DRV("DESCHTML")) Then
                    LTL_CNTHTML.Text = DRV("DESCHTML").ToString

                    DIV_MORE.Visible = True
                Else
                    LTL_CNTHTML.Text = DRV("CNTHTML").ToString
                End If
                LTL_CNTHTML_MORE.Text = DRV("CNTHTML").ToString

                Dim DIV_LTL_CNTHTML As HtmlGenericControl = e.Item.FindControl("DIV_LTL_CNTHTML")
                Dim DIV_LTL_CNTHTML_MORE As HtmlGenericControl = e.Item.FindControl("DIV_LTL_CNTHTML_MORE")
                Dim btnReadMore As HtmlInputButton = e.Item.FindControl("btnReadMore")
                btnReadMore.Attributes("onclick") = "$('#" & DIV_LTL_CNTHTML.ClientID & "').hide();$('#" & DIV_LTL_CNTHTML_MORE.ClientID & "').show();$('#" & DIV_MORE.ClientID & "').hide();"

                Dim HID_MB_SEQ As HtmlInputHidden = e.Item.FindControl("HID_MB_SEQ")
                Dim sMB_PLACE_VALUE As String = String.Empty
                If IsNumeric(HID_MB_SEQ.Value) Then
                    Dim mbCLASSList As New MB_CLASSList(Me.m_DBManager)
                    mbCLASSList.LoadByMB_SEQ(CDec(HID_MB_SEQ.Value))
                    'Dim RP_SIGN As Repeater = e.Item.FindControl("RP_SIGN")
                    'RP_SIGN.DataSource = mbCLASSList.getCurrentDataSet.Tables(0).Rows(0)
                    'RP_SIGN.DataBind()

                    Dim btnClass As Button = e.Item.FindControl("btnClass")
                    Dim btnModify As Button = e.Item.FindControl("btnModify")
                    Dim btnCanClass As Button = e.Item.FindControl("btnCanClass")
                    Dim LTL_APLY As Label = e.Item.FindControl("LTL_APLY")
                    Me.getAPLY_MSG(HID_MB_SEQ.Value, btnClass, btnModify, btnCanClass, LTL_APLY)

                    Dim PNL_CLASS As Panel = e.Item.FindControl("PNL_CLASS")
                    PNL_CLASS.Visible = True

                    If mbCLASSList.getCurrentDataSet.Tables(0).Rows.Count > 0 Then
                        Dim sMB_PLACE As String = String.Empty
                        sMB_PLACE = mbCLASSList.getCurrentDataSet.Tables(0).Rows(0)("MB_PLACE").ToString

                        If Utility.isValidateData(sMB_PLACE) Then
                            Dim AP_CODEList As New AP_CODEList(Me.m_DBManager)
                            AP_CODEList.loadByUpCode("177")
                            Dim ROW_VALUE() As DataRow = Nothing
                            ROW_VALUE = AP_CODEList.getCurrentDataSet.Tables(0).Select("TEXT='" & sMB_PLACE & "'")
                            If Not IsNothing(ROW_VALUE) AndAlso ROW_VALUE.Length > 0 Then
                                sMB_PLACE_VALUE = ROW_VALUE(0)("VALUE").ToString
                            End If
                        End If
                    End If
                End If

                Dim IMG_EDIT As ImageButton = e.Item.FindControl("IMG_EDIT")
                IMG_EDIT.ImageUrl = com.Azion.EloanUtility.UIUtility.getImgPath & "EDIT.png"

                Dim IMG_TOP As ImageButton = e.Item.FindControl("IMG_TOP")
                IMG_TOP.ImageUrl = com.Azion.EloanUtility.UIUtility.getImgPath & "TOP.png"

                Dim IMG_UP As ImageButton = e.Item.FindControl("IMG_UP")
                IMG_UP.ImageUrl = com.Azion.EloanUtility.UIUtility.getImgPath & "UP.png"

                Dim IMG_DOWN As ImageButton = e.Item.FindControl("IMG_DOWN")
                IMG_DOWN.ImageUrl = com.Azion.EloanUtility.UIUtility.getImgPath & "DOWN.png"

                Dim IMG_DEL As ImageButton = e.Item.FindControl("IMG_DEL")
                IMG_DEL.ImageUrl = com.Azion.EloanUtility.UIUtility.getImgPath & "Trash.png"

                Dim IMG_ADD As ImageButton = e.Item.FindControl("IMG_ADD")
                IMG_ADD.ImageUrl = com.Azion.EloanUtility.UIUtility.getImgPath & "Pencil.png"

                If IsNumeric(Session("admin")) AndAlso CInt(Session("admin")) >= 4 Then
                    Dim PLH_ADMIN As PlaceHolder = e.Item.FindControl("PLH_ADMIN")
                    PLH_ADMIN.Visible = True

                    If Utility.isValidateData(Session("BRID")) AndAlso Session("BRID") = "0" Then
                        '總會可維護所有中心的資料
                    ElseIf Utility.isValidateData(HID_BRID.Value) AndAlso Utility.isValidateData(Session("BRID")) Then
                        If HID_BRID.Value <> Session("BRID") Then
                            PLH_ADMIN.Visible = False
                        End If
                    ElseIf Utility.isValidateData(sMB_PLACE_VALUE) AndAlso Utility.isValidateData(Session("BRID")) Then
                        If sMB_PLACE_VALUE <> Session("BRID") Then
                            PLH_ADMIN.Visible = False
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

    Sub getAPLY_MSG(ByVal iMB_SEQ As Decimal, ByVal btnClass As Button, ByVal btnModify As Button, ByVal btnCanClass As Button, ByVal LTL_APLY As Label)
        Dim MB_CLASSList As New MB_CLASSList(Me.m_DBManager)
        MB_CLASSList.setSQLCondition(" ORDER BY MB_BATCH ")
        MB_CLASSList.LoadByMB_SEQ(iMB_SEQ)
        Dim MB_MEMCLASSList As New MB_MEMCLASSList(Me.m_DBManager)
        Dim iVadSign As Integer = 0
        Dim iMB_YES As Integer = 0
        Dim sb_APLY As New System.Text.StringBuilder
        For Each ROW As DataRow In MB_CLASSList.getCurrentDataSet.Tables(0).Rows
            MB_MEMCLASSList.clear()
            MB_MEMCLASSList.setSQLCondition(" AND IFNULL(A.MB_FWMK,'') NOT IN ('3','4','5') AND IFNULL(B.MB_ELECT,' ')='1' AND IFNULL(B.MB_RESP,' ')<>'N' ")
            MB_MEMCLASSList.loadByMB_SEQ(iMB_SEQ, ROW("MB_BATCH"))
            Dim isFULL As Boolean = False
            If MB_MEMCLASSList.size >= (Utility.CheckNumNull(ROW("MB_FULL")) + Utility.CheckNumNull(ROW("MB_WAIT"))) Then
                '報名已額滿
                isFULL = True

                '1050217 AMY
                '報名已額滿仍顯示報名及取消報名
                iVadSign += 1
            Else
                If Now >= CDate(ROW("MB_SAPLY").ToString) AndAlso Now <= CDate(ROW("MB_EAPLY").ToString) Then
                    iVadSign += 1
                End If
            End If

            Dim isEndSign As Boolean = False
            If ROW("MB_YES").ToString = "N" Then
                isEndSign = True
            Else
                If IsDate(ROW("MB_EAPLY").ToString) Then
                    If Now > CDate(ROW("MB_EAPLY").ToString) Then
                        isEndSign = True
                    End If
                End If

                If Not isEndSign Then
                    iMB_YES += 1
                End If
            End If

            Dim sBatch As String = String.Empty
            If CDec(ROW("MB_BATCH")) > 0 Then
                sBatch = "第" & ROW("MB_BATCH").ToString & "梯次"
            End If
            If isEndSign Then
                sb_APLY.Append(sBatch).Append("已截止報名").Append("<BR/>")
            Else
                If isFULL Then
                    'sb_APLY.Append(sBatch).Append("報名已額滿").Append("<BR/>")
                Else
                    If IsDate(ROW("MB_SAPLY").ToString) Then
                        sb_APLY.Append(sBatch).Append(CDate(ROW("MB_SAPLY").ToString).Month & "月" & CDate(ROW("MB_SAPLY").ToString).Day & "日開始報名").Append("<BR/>")
                    End If
                End If
            End If
        Next

        If iMB_YES = 0 Then
            LTL_APLY.Text = "已截止報名"
            btnClass.Visible = False
            btnModify.Visible = False
            btnCanClass.Visible = False
        Else
            Dim sBR As String = String.Empty
            If iVadSign > 0 Then
                btnClass.Visible = True
                btnModify.Visible = True
                btnCanClass.Visible = True
                sBR = "<BR/>"
            Else
                btnClass.Visible = False
                btnModify.Visible = False
                btnCanClass.Visible = False
            End If

            Dim sAPLY As String = String.Empty
            sAPLY = sb_APLY.ToString
            If Utility.isValidateData(sAPLY) Then
                sAPLY = Left(sAPLY, sAPLY.Length - 5)
                LTL_APLY.Text = sBR & sAPLY
            End If
        End If
    End Sub

    'Sub RP_SIGN_OnItemDataBound(ByVal Sender As Object, ByVal e As RepeaterItemEventArgs)
    '    Try
    '        If e.Item.ItemType = ListItemType.Item OrElse e.Item.ItemType = ListItemType.AlternatingItem Then
    '            Dim DRV As DataRowView = CType(e.Item.DataItem, DataRowView)
    '            Dim LTL_MB_BATCH As Label = e.Item.FindControl("LTL_MB_BATCH")
    '            If Not IsNumeric(DRV("MB_BATCH")) OrElse DRV("MB_BATCH") = 0 Then
    '                LTL_MB_BATCH.Visible = False
    '            Else
    '                'LTL_MB_BATCH.Text = "第【" & DRV("MB_BATCH") & "】梯次"
    '                LTL_MB_BATCH.Text = "第" & DRV("MB_BATCH") & "梯次"
    '            End If

    '            Dim btnClass As Button = e.Item.FindControl("btnClass")
    '            Dim HID_MB_BATCH As HtmlInputHidden = e.Item.FindControl("HID_MB_BATCH")
    '            Dim MB_CLASS As New MB_CLASS(m_DBManager)
    '            If MB_CLASS.loadByPK(DRV("MB_SEQ"), HID_MB_BATCH.Value) Then
    '                Dim MB_MEMCLASSList As New MB_MEMCLASSList(m_DBManager)
    '                MB_MEMCLASSList.setSQLCondition(" AND IFNULL(MB_FWMK,'') NOT IN ('3','4') ")
    '                MB_MEMCLASSList.loadByMB_SEQ(DRV("MB_SEQ"), HID_MB_BATCH.Value)

    '                If MB_MEMCLASSList.size >= (MB_CLASS.getInt("MB_FULL") + MB_CLASS.getInt("MB_WAIT")) Then
    '                    btnClass.Text = "報名已額滿"
    '                    btnClass.Enabled = False
    '                End If
    '            End If

    '            Dim LTL_APLY As Label = e.Item.FindControl("LTL_APLY")
    '            If MB_CLASS.getString("MB_YES") = "N" Then
    '                LTL_APLY.Text = "已截止報名"
    '                Dim PLH_APLY As PlaceHolder = e.Item.FindControl("PLH_APLY")
    '                PLH_APLY.Visible = True
    '                btnClass.Visible = False
    '                Dim btnCanClass As Button = e.Item.FindControl("btnCanClass")
    '                btnCanClass.Visible = False
    '            Else
    '                If IsDate(DRV("MB_SAPLY").ToString) Then
    '                    'LTL_APLY.Text &= CDate(DRV("MB_SAPLY").ToString).Year - 1911 & "年" & CDate(DRV("MB_SAPLY").ToString).Month & "月" & CDate(DRV("MB_SAPLY").ToString).Day & "日"
    '                    LTL_APLY.Text = CDate(DRV("MB_SAPLY").ToString).Month & "月" & CDate(DRV("MB_SAPLY").ToString).Day & "日開始報名"
    '                End If
    '                If IsDate(DRV("MB_EAPLY").ToString) Then
    '                    If Now > CDate(DRV("MB_EAPLY").ToString) Then
    '                        LTL_APLY.Text = "已截止報名"
    '                    End If
    '                End If

    '                'If IsDate(DRV("MB_EAPLY").ToString) Then
    '                '    LTL_APLY.Text &= "~" & CDate(DRV("MB_EAPLY").ToString).Year - 1911 & "年" & CDate(DRV("MB_EAPLY").ToString).Month & "月" & CDate(DRV("MB_EAPLY").ToString).Day & "日"
    '                'End If

    '                Dim btnCanClass As Button = e.Item.FindControl("btnCanClass")
    '                Dim PLH_APLY As PlaceHolder = e.Item.FindControl("PLH_APLY")
    '                If Now >= CDate(DRV("MB_SAPLY").ToString) AndAlso Now <= CDate(DRV("MB_EAPLY").ToString) Then
    '                    PLH_APLY.Visible = False
    '                    btnClass.Visible = True
    '                    btnCanClass.Visible = True
    '                Else
    '                    PLH_APLY.Visible = True
    '                    btnClass.Visible = False
    '                    btnCanClass.Visible = False
    '                End If
    '            End If
    '        End If
    '    Catch ex As Exception
    '        Throw
    '    End Try
    'End Sub

    Private Sub RP_NEWS_ItemCommand(source As Object, e As System.Web.UI.WebControls.RepeaterCommandEventArgs) Handles RP_NEWS.ItemCommand
        Try
            If UCase(e.CommandName) = "EDIT" Then
                Me.PLH_CKEditor.Visible = True

                Me.CKEditorControl1.Text = String.Empty

                Me.CKEditorControl2.Text = String.Empty

                Me.LBL_HTML.Text = String.Empty

                'Me.EVENT.Text = String.Empty

                'Me.btAdd.Visible = False

                Dim HID_CRETIME As HtmlInputHidden = e.Item.FindControl("HID_CRETIME")
                Dim HID_CHGTIME As HtmlInputHidden = e.Item.FindControl("HID_CHGTIME")
                Me.LTL_YYYYMMDD.Text = Left(HID_CHGTIME.Value, 8)

                Me.LTL_HHMMSS.Text = Right(HID_CHGTIME.Value, 6)

                Me.HID_CRETIME.Value = e.CommandArgument
                Dim HID_SEQTIME As HtmlInputHidden = e.Item.FindControl("HID_SEQTIME")
                Me.HID_SEQTIME.Value = HID_SEQTIME.Value

                Me.HID_INDEX.Value = e.Item.ItemIndex

                Dim HID_CODEID As HtmlInputHidden = e.Item.FindControl("HID_CODEID")
                Me.HID_CODEID.Value = HID_CODEID.Value

                'Dim LTL_CNTHTML As Literal = e.Item.FindControl("LTL_CNTHTML")

                Dim mbNEWS As New MB_NEWS(Me.m_DBManager)
                mbNEWS.LoadByPK(HID_CRETIME.Value)

                Me.CKEditorControl1.Text = mbNEWS.getString("CNTHTML")
                Me.CKEditorControl2.Text = mbNEWS.getString("DESCHTML")

                'Me.EVENT.Text = mbNEWS.getString("EVENT")

                Me.RP_NEWS.Visible = False

                '選擇課程
                'Dim sURLClass As String = String.Empty
                'sURLClass = com.Azion.EloanUtility.UIUtility.getRootPath & "/MNT/MBMnt_Class_01_v01.aspx?OPTYPE=Q&CRETIME=" & Me.HID_CRETIME.Value
                'Me.btnChoose.Attributes("onclick") = "chooseClass('" & sURLClass & "');"
            ElseIf UCase(e.CommandName) = "MORE" Then
                Dim mbNEWS As New MB_NEWS(Me.m_DBManager)
                Dim HID_CRETIME As HtmlInputHidden = e.Item.FindControl("HID_CRETIME")
                mbNEWS.LoadByPK(CDec(HID_CRETIME.Value))
                Dim LTL_CNTHTML As Literal = e.Item.FindControl("LTL_CNTHTML")
                LTL_CNTHTML.Text = mbNEWS.getString("CNTHTML")
                Dim DIV_MORE As HtmlGenericControl = e.Item.FindControl("DIV_MORE")
                DIV_MORE.Visible = False
                Dim HID_MB_SEQ As HtmlInputHidden = e.Item.FindControl("HID_MB_SEQ")
                If IsNumeric(HID_MB_SEQ.Value) Then
                    'Dim mbCLASSList As New MB_CLASSList(Me.m_DBManager)
                    'mbCLASSList.LoadByMB_SEQ(CDec(HID_MB_SEQ.Value))
                    'Dim RP_SIGN As Repeater = e.Item.FindControl("RP_SIGN")
                    'RP_SIGN.DataSource = mbCLASSList.getCurrentDataSet.Tables(0)
                    'RP_SIGN.DataBind()

                    'Dim PNL_CLASS As Panel = e.Item.FindControl("PNL_CLASS")
                    'PNL_CLASS.Visible = True

                    Dim btnClass As Button = e.Item.FindControl("btnClass")
                    Dim btnModify As Button = e.Item.FindControl("btnModify")
                    Dim btnCanClass As Button = e.Item.FindControl("btnCanClass")
                    Dim LTL_APLY As Label = e.Item.FindControl("LTL_APLY")
                    Me.getAPLY_MSG(HID_MB_SEQ.Value, btnClass, btnModify, btnCanClass, LTL_APLY)
                End If

                'Dim DIV_TITLE As HtmlGenericControl = e.Item.FindControl("DIV_TITLE")
                'com.Azion.EloanUtility.UIUtility.setObjFocus(Me, DIV_TITLE.ClientID)
            ElseIf UCase(e.CommandName) = "CLASS" Then
                '我要報名
                'If Not com.Azion.EloanUtility.ValidateUtility.isValidateData(Session("USERID")) Then
                '    com.Azion.EloanUtility.UIUtility.alert("請先登入或成為會員")
                '    com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請先登入或成為會員")
                '    Return
                'End If

                'Dim HID_MB_SEQ As HtmlInputHidden = e.Item.FindControl("HID_MB_SEQ")
                'Dim sURL As String = String.Empty
                'sURL = com.Azion.EloanUtility.UIUtility.getRootPath & "/MNT/MBMnt_Sign_01_v01.aspx?MBSEQ=" & HID_MB_SEQ.Value

                'Response.Redirect(sURL)

                Dim HID_MB_SEQ As HtmlInputHidden = e.Item.FindControl("HID_MB_SEQ")
                Dim MB_CLASSList As New MB_CLASSList(Me.m_DBManager)
                MB_CLASSList.setSQLCondition(" ORDER BY MB_BATCH ")
                MB_CLASSList.LoadByMB_SEQ(CDec(HID_MB_SEQ.Value))
                Dim MB_MEMCLASSList As New MB_MEMCLASSList(Me.m_DBManager)
                
                Dim iVadSign As Integer = 0 
                For Each ROW As DataRow In MB_CLASSList.getCurrentDataSet.Tables(0).Rows
                    MB_MEMCLASSList.clear()
                    MB_MEMCLASSList.setSQLCondition(" AND IFNULL(A.MB_FWMK,'') NOT IN ('3','4','5') AND IFNULL(B.MB_ELECT,' ')='1' AND IFNULL(B.MB_RESP,' ')<>'N' ")
                    MB_MEMCLASSList.loadByMB_SEQ(CDec(HID_MB_SEQ.Value), ROW("MB_BATCH"))

                    If MB_MEMCLASSList.size >= (Utility.CheckNumNull(ROW("MB_FULL")) + Utility.CheckNumNull(ROW("MB_WAIT"))) Then
                        '報名已額滿
                    Else
                        If Now >= CDate(ROW("MB_SAPLY").ToString) AndAlso Now <= CDate(ROW("MB_EAPLY").ToString) Then
                            iVadSign += 1
                        End If                        
                    End If
                Next

                If iVadSign=0 Then
                    com.Azion.EloanUtility.UIUtility.showJSMsg(Me,"報名已額滿")
                    com.Azion.EloanUtility.UIUtility.showErrMsg(Me,"報名已額滿")
                    Return
                End If

                If Not com.Azion.EloanUtility.ValidateUtility.isValidateData(Session("USERID")) Then
                    com.Azion.EloanUtility.UIUtility.alert("請先登入或成為會員")
                    com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請先登入或成為會員")
                    Dim sURL_Sign As String = String.Empty
                    sURL_Sign = com.Azion.EloanUtility.UIUtility.getRootPath & "/MNT/MBSignIn_01_v01.aspx"
                    Server.Transfer(sURL_Sign)
                    Return
                End If
                Dim sURL As String = String.Empty
                sURL = com.Azion.EloanUtility.UIUtility.getRootPath &
                       "/MNT/MBMnt_Sign_01_v01.aspx?MBSEQ=" & HID_MB_SEQ.Value '& "&MB_BATCH=" & HID_MB_BATCH.Value

                Response.Redirect(sURL)
            ElseIf UCase(e.CommandName) = "CANCLASS" Then
                '取消報名
                'If Not com.Azion.EloanUtility.ValidateUtility.isValidateData(Session("USERID")) Then
                '    com.Azion.EloanUtility.UIUtility.alert("請先登入或成為會員")
                '    com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請先登入或成為會員")
                '    Return
                'End If

                'Dim HID_MB_SEQ As HtmlInputHidden = e.Item.FindControl("HID_MB_SEQ")
                'Dim sURL As String = String.Empty
                'sURL = com.Azion.EloanUtility.UIUtility.getRootPath & "/MNT/MBMnt_Sign_01_v01.aspx?MBSEQ=" & HID_MB_SEQ.Value

                'Response.Redirect(sURL)

                If Not com.Azion.EloanUtility.ValidateUtility.isValidateData(Session("USERID")) Then
                    com.Azion.EloanUtility.UIUtility.alert("請先登入或成為會員")
                    com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請先登入或成為會員")
                    Dim sURL_Sign As String = String.Empty
                    sURL_Sign = com.Azion.EloanUtility.UIUtility.getRootPath & "/MNT/MBSignIn_01_v01.aspx"
                    Server.Transfer(sURL_Sign)
                    Return
                End If

                Dim HID_MB_SEQ As HtmlInputHidden = e.Item.FindControl("HID_MB_SEQ")
                Dim sURL As String = String.Empty
                sURL = com.Azion.EloanUtility.UIUtility.getRootPath &
                       "/MNT/MBMnt_Sign_01_v01.aspx?OPTYPE=CANCEL&MBSEQ=" & HID_MB_SEQ.Value '& "&MB_BATCH=" & HID_MB_BATCH.Value

                Response.Redirect(sURL)
            ElseIf UCase(e.CommandName) = "DEL" Then
                Dim iINDEX As Integer = e.Item.ItemIndex
                Dim HID_CRETIME As HtmlInputHidden = e.Item.FindControl("HID_CRETIME")
                Dim HID_SEQTIME As HtmlInputHidden = e.Item.FindControl("HID_SEQTIME")
                Dim mbNEWS As New MB_NEWS(Me.m_DBManager)
                If mbNEWS.LoadByPK(CDec(HID_CRETIME.Value)) Then
                    mbNEWS.remove()
                End If

                If Not IsNumeric(Me.m_sCODEID) Then
                    Dim sURL As String = String.Empty
                    sURL = com.Azion.EloanUtility.UIUtility.getRootPath & "/NewsList.aspx"

                    Response.Redirect(sURL)
                Else
                    Dim objRepeaterItem As RepeaterItem = Me.RP_NEWS.Items(0)
                    Dim HID_FST_SORTNO As HtmlInputHidden = objRepeaterItem.FindControl("HID_SORTNO")
                    Me.Bind_RP_NEWS(CInt(Me.m_sCODEID), CDec(HID_SEQTIME.Value), 1)

                    If (iINDEX - 1) >= 0 Then
                        Dim objFocusRepeaterItem = Me.RP_NEWS.Items(iINDEX - 1)
                        Dim IMG_DEL As ImageButton = objFocusRepeaterItem.FindControl("IMG_DEL")
                        Utility.setObjFocus(IMG_DEL.ClientID, Me.Page)
                    End If
                End If
            ElseIf UCase(e.CommandName) = "ADD" Then
                Me.btAdd_Click()
            ElseIf UCase(e.CommandName) = "DOWN" Then
                Dim iNEW_SEQTIME As Decimal = 0

                Dim HID_CRETIME As HtmlInputHidden = e.Item.FindControl("HID_CRETIME")
                Dim HID_SEQTIME As HtmlInputHidden = e.Item.FindControl("HID_SEQTIME")
                Dim mbNEWSList As New MB_NEWSList(Me.m_DBManager)
                If IsNumeric(Me.m_sCODEID) Then
                    mbNEWSList.LoadByCODEID_PREV_FST(CDec(Me.m_sCODEID), CDec(HID_CRETIME.Value), 1)
                    If mbNEWSList.getCurrentDataSet.Tables(0).Rows.Count > 0 Then
                        iNEW_SEQTIME = mbNEWSList.getCurrentDataSet.Tables(0).Rows(0)("SEQTIME")
                        Me.SAVE_NEXT_SEQTIME(mbNEWSList.getCurrentDataSet.Tables(0).Rows(0), HID_CRETIME, HID_SEQTIME)
                    Else
                        com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "已經是此類別最舊的文章了，無法排序")
                        Return
                    End If
                Else
                    mbNEWSList.LoadByTIME_PREV_FST(CDec(HID_SEQTIME.Value), 1)
                    If mbNEWSList.getCurrentDataSet.Tables(0).Rows.Count > 0 Then
                        iNEW_SEQTIME = mbNEWSList.getCurrentDataSet.Tables(0).Rows(0)("SEQTIME")
                        Me.SAVE_NEXT_SEQTIME(mbNEWSList.getCurrentDataSet.Tables(0).Rows(0), HID_CRETIME, HID_SEQTIME)
                    End If
                End If

                If iNEW_SEQTIME > 0 AndAlso Me.RP_NEWS.Items.Count > 0 Then
                    Dim objRepeaterItem As RepeaterItem = Me.RP_NEWS.Items(Me.RP_NEWS.Items.Count - 1)
                    Dim HID_LST_SEQTIME As HtmlInputHidden = objRepeaterItem.FindControl("HID_SEQTIME")
                    If Not com.Azion.EloanUtility.ValidateUtility.isValidateData(Me.m_sCODEID) OrElse Me.m_sCODEID = Me.m_sNEWSODEID Then
                        If iNEW_SEQTIME < CDec(HID_LST_SEQTIME.Value) Then
                            Me.Bind_RP_NEWSByTime(iNEW_SEQTIME + 1, 1)
                        Else
                            Me.Bind_RP_NEWSByTime(CDec(HID_LST_SEQTIME.Value) - 1, 2)
                        End If
                    Else
                        If iNEW_SEQTIME < CDec(HID_LST_SEQTIME.Value) Then
                            Me.Bind_RP_NEWS(CInt(Me.m_sCODEID), iNEW_SEQTIME + 1, 1)
                        Else
                            Me.Bind_RP_NEWS(CInt(Me.m_sCODEID), CDec(HID_LST_SEQTIME.Value) - 1, 2)
                        End If
                    End If
                End If
            ElseIf UCase(e.CommandName) = "UP" Then
                Dim iNEW_SEQTIME As Decimal = 0

                Dim HID_CRETIME As HtmlInputHidden = e.Item.FindControl("HID_CRETIME")
                Dim HID_SEQTIME As HtmlInputHidden = e.Item.FindControl("HID_SEQTIME")
                Dim mbNEWSList As New MB_NEWSList(Me.m_DBManager)
                If IsNumeric(Me.m_sCODEID) Then
                    mbNEWSList.LoadByCODEID_NEXT_FST(CDec(Me.m_sCODEID), CDec(HID_SEQTIME.Value), 1)
                    If mbNEWSList.getCurrentDataSet.Tables(0).Rows.Count > 0 Then
                        iNEW_SEQTIME = mbNEWSList.getCurrentDataSet.Tables(0).Rows(0)("SEQTIME")
                        Me.SAVE_NEXT_SEQTIME(mbNEWSList.getCurrentDataSet.Tables(0).Rows(0), HID_CRETIME, HID_SEQTIME)
                    Else
                        com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "已經是此類別最新文章了，無法排序")
                        Return
                    End If
                Else
                    mbNEWSList.LoadByTIME_NEXT_FST(CDec(HID_SEQTIME.Value), 1)
                    If mbNEWSList.getCurrentDataSet.Tables(0).Rows.Count > 0 Then
                        iNEW_SEQTIME = mbNEWSList.getCurrentDataSet.Tables(0).Rows(0)("SEQTIME")
                        Me.SAVE_NEXT_SEQTIME(mbNEWSList.getCurrentDataSet.Tables(0).Rows(0), HID_CRETIME, HID_SEQTIME)
                    End If
                End If

                If iNEW_SEQTIME > 0 AndAlso Me.RP_NEWS.Items.Count > 0 Then
                    Dim objRepeaterItem As RepeaterItem = Me.RP_NEWS.Items(0)
                    Dim HID_FST_SEQTIME As HtmlInputHidden = objRepeaterItem.FindControl("HID_SEQTIME")
                    If Not com.Azion.EloanUtility.ValidateUtility.isValidateData(Me.m_sCODEID) OrElse Me.m_sCODEID = Me.m_sNEWSODEID Then
                        If iNEW_SEQTIME > CDec(HID_FST_SEQTIME.Value) Then
                            Me.Bind_RP_NEWSByTime(iNEW_SEQTIME - 1, 2)
                        Else
                            Me.Bind_RP_NEWSByTime(CDec(HID_FST_SEQTIME.Value) + 1, 1)
                        End If
                    Else
                        If iNEW_SEQTIME > CDec(HID_FST_SEQTIME.Value) Then
                            Me.Bind_RP_NEWS(CInt(Me.m_sCODEID), iNEW_SEQTIME - 1, 2)
                        Else
                            Me.Bind_RP_NEWS(CInt(Me.m_sCODEID), CDec(HID_FST_SEQTIME.Value) + 1, 1)
                        End If
                    End If
                End If
            ElseIf UCase(e.CommandName) = "TOP" Then
                Dim HID_CRETIME As HtmlInputHidden = e.Item.FindControl("HID_CRETIME")
                Dim mbNEWS As New MB_NEWS(Me.m_DBManager)
                If mbNEWS.LoadByPK(CDec(HID_CRETIME.Value)) Then
                    Dim mbMAXNEWS As New MB_NEWS(Me.m_DBManager)
                    Dim iMAX_SEQTIME As Decimal = 0
                    iMAX_SEQTIME = mbMAXNEWS.getALL_MAX_SEQTIME()
                    mbNEWS.setAttribute("SEQTIME", iMAX_SEQTIME + 1)
                    mbNEWS.save()
                End If

                Dim sURL As String = String.Empty
                sURL = com.Azion.EloanUtility.UIUtility.getRootPath & "/NewsList.aspx"

                Response.Redirect(sURL)
            End If
        Catch ex As System.Threading.ThreadAbortException
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

    'Sub RP_SIGN_OnItemCommand(source As Object, e As System.Web.UI.WebControls.RepeaterCommandEventArgs)
    '    Try
    '        Dim sURL_Sign As String = String.Empty
    '        sURL_Sign = com.Azion.EloanUtility.UIUtility.getRootPath & "/MNT/MBSignIn_01_v01.aspx"

    '        If UCase(e.CommandName) = "CLASS" Then
    '            '我要報名
    '            If Not com.Azion.EloanUtility.ValidateUtility.isValidateData(Session("USERID")) Then
    '                com.Azion.EloanUtility.UIUtility.alert("請先登入或成為會員")
    '                com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請先登入或成為會員")
    '                Server.Transfer(sURL_Sign)
    '                Return
    '            End If

    '            Dim HID_MB_SEQ As HtmlInputHidden = e.Item.FindControl("HID_MB_SEQ")
    '            Dim HID_MB_BATCH As HtmlInputHidden = e.Item.FindControl("HID_MB_BATCH")

    '            Dim sURL As String = String.Empty
    '            sURL = com.Azion.EloanUtility.UIUtility.getRootPath & "/MNT/MBMnt_Sign_01_v01.aspx?MBSEQ=" & HID_MB_SEQ.Value & "&MB_BATCH=" & HID_MB_BATCH.Value

    '            Response.Redirect(sURL)
    '        ElseIf UCase(e.CommandName) = "CANCLASS" Then
    '            '取消報名
    '            If Not com.Azion.EloanUtility.ValidateUtility.isValidateData(Session("USERID")) Then
    '                com.Azion.EloanUtility.UIUtility.alert("請先登入或成為會員")
    '                com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請先登入或成為會員")
    '                Server.Transfer(sURL_Sign)
    '                Return
    '            End If

    '            Dim HID_MB_SEQ As HtmlInputHidden = e.Item.FindControl("HID_MB_SEQ")
    '            Dim HID_MB_BATCH As HtmlInputHidden = e.Item.FindControl("HID_MB_BATCH")

    '            Dim sURL As String = String.Empty
    '            sURL = com.Azion.EloanUtility.UIUtility.getRootPath & "/MNT/MBMnt_Sign_01_v01.aspx?OPTYPE=CANCEL&MBSEQ=" & HID_MB_SEQ.Value & "&MB_BATCH=" & HID_MB_BATCH.Value

    '            Response.Redirect(sURL)
    '        End If
    '    Catch ex As System.Threading.ThreadAbortException
    '    Catch ex As Exception
    '        com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
    '    End Try
    'End Sub

    Sub SAVE_NEXT_SEQTIME(ByVal sqlRow As DataRow, ByVal HID_CRETIME As HtmlInputHidden, ByVal HID_SEQTIME As HtmlInputHidden)
        Dim mbNEWS As New MB_NEWS(Me.m_DBManager)
        If mbNEWS.LoadByPK(sqlRow("CRETIME")) Then
            mbNEWS.remove()
        End If
        mbNEWS.clear()
        If mbNEWS.LoadByPK(CDec(HID_CRETIME.Value)) Then
            mbNEWS.setAttribute("SEQTIME", sqlRow("SEQTIME"))
            mbNEWS.save()
        End If
        mbNEWS.clear()
        For i As Integer = 0 To sqlRow.Table.Columns.Count - 1
            Dim sCOLNAME As String = String.Empty
            sCOLNAME = UCase(sqlRow.Table.Columns(i).ColumnName)
            If Not IsDBNull(sqlRow(sCOLNAME)) Then
                mbNEWS.setAttribute(sCOLNAME, sqlRow(sCOLNAME))
            Else
                mbNEWS.setAttribute(sCOLNAME, Nothing)
            End If
        Next
        mbNEWS.setAttribute("SEQTIME", CDec(HID_SEQTIME.Value))
        mbNEWS.save()
    End Sub

    Function getNEW_SEQTIME() As Decimal
        Try
            Dim D_NOW As Date = Now

            Dim iYYYYMMDD As Decimal = (D_NOW.Year * 10000) + (D_NOW.Month * 100) + D_NOW.Day

            Dim iHHMMSS As Decimal = (D_NOW.Hour * 10000) + (D_NOW.Minute * 100) + D_NOW.Second

            Return iYYYYMMDD * 1000000 + iHHMMSS
        Catch ex As Exception
            Throw
        End Try
    End Function

#Region "Button Events"
    Private Sub IMG_RIGHT_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles IMG_RIGHT.Click
        Try
            If Me.RP_NEWS.Items.Count > 0 Then
                Dim objRepeaterItem As RepeaterItem = Me.RP_NEWS.Items(Me.RP_NEWS.Items.Count - 1)

                If Not com.Azion.EloanUtility.ValidateUtility.isValidateData(Me.m_sCODEID) OrElse Me.m_sCODEID = Me.m_sNEWSODEID Then
                    Dim HID_SEQTIME As HtmlInputHidden = objRepeaterItem.FindControl("HID_SEQTIME")

                    Me.Bind_RP_NEWSByTime(CDec(HID_SEQTIME.Value), 1)
                Else
                    Dim HID_SEQTIME As HtmlInputHidden = objRepeaterItem.FindControl("HID_SEQTIME")

                    Me.Bind_RP_NEWS(CInt(Me.m_sCODEID), CDec(HID_SEQTIME.Value), 1)
                End If
            End If
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

    Private Sub IMG_LST_Click(sender As Object, e As ImageClickEventArgs) Handles IMG_LST.Click
        Try
            Dim MB_NEWS As New MB_NEWS(Me.m_DBManager)
            Dim iSEQTIME As Decimal = 0
            If Me.m_sCODEID = Me.m_sNEWSODEID OrElse Not IsNumeric(Me.m_sCODEID) Then
                iSEQTIME = MB_NEWS.getMIN_SEQTIME() - 1

                Dim MB_NEWSList As New MB_NEWSList(Me.m_DBManager)
                MB_NEWSList.LoadByTIME_NEXT_ASC(iSEQTIME, Me.m_iPageSize)
                If MB_NEWSList.getCurrentDataSet.Tables(0).Rows.Count > 0 Then
                    Me.RP_NEWS.DataSource = MB_NEWSList.getCurrentDataSet.Tables(0)
                    Me.RP_NEWS.DataBind()
                Else
                    Me.RP_NEWS.DataSource = Nothing
                    Me.RP_NEWS.DataBind()
                End If

                Me.RP_NEWS.Visible = True
            Else
                iSEQTIME = MB_NEWS.getMIN_SEQTIME_CODEID(CInt(Me.m_sCODEID)) - 1

                Dim MB_NEWSList As New MB_NEWSList(Me.m_DBManager)
                MB_NEWSList.LoadByCODEID_NEXT_ASC(CInt(Me.m_sCODEID), iSEQTIME, Me.m_iPageSize)

                If MB_NEWSList.getCurrentDataSet.Tables(0).Rows.Count > 0 Then
                    Me.RP_NEWS.DataSource = MB_NEWSList.getCurrentDataSet.Tables(0)
                    Me.RP_NEWS.DataBind()
                Else
                    Me.RP_NEWS.DataSource = Nothing
                    Me.RP_NEWS.DataBind()
                End If

                Me.RP_NEWS.Visible = True
            End If
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

    Private Sub IMG_LEFT_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles IMG_LEFT.Click
        Try
            If Me.RP_NEWS.Items.Count > 0 Then
                Dim objRepeaterItem As RepeaterItem = Me.RP_NEWS.Items(0)

                If Not com.Azion.EloanUtility.ValidateUtility.isValidateData(Me.m_sCODEID) OrElse Me.m_sCODEID = Me.m_sNEWSODEID Then
                    Dim HID_SEQTIME As HtmlInputHidden = objRepeaterItem.FindControl("HID_SEQTIME")

                    Me.Bind_RP_NEWSByTime(CDec(HID_SEQTIME.Value), 2)
                Else
                    Dim HID_SEQTIME As HtmlInputHidden = objRepeaterItem.FindControl("HID_SEQTIME")

                    Me.Bind_RP_NEWS(CInt(Me.m_sCODEID), CDec(HID_SEQTIME.Value), 2)
                End If
            End If
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

    Private Sub IMG_FST_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles IMG_FST.Click
        Try
            If Me.m_sCODEID = Me.m_sNEWSODEID OrElse Not IsNumeric(Me.m_sCODEID) Then
                Dim D_NOW As Date = Now
                Me.Bind_RP_NEWSByTime(Me.getNEW_SEQTIME, 1)
            Else
                Dim mbNEWS As New MB_NEWS(Me.m_DBManager)
                Me.Bind_RP_NEWS(CInt(Me.m_sCODEID), mbNEWS.getMAX_SEQTIME(CInt(Me.m_sCODEID)), 1)
            End If
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

    Private Sub btAdd_Click()
        If Not com.Azion.EloanUtility.ValidateUtility.isValidateData(Me.m_sCODEID) OrElse
            Me.m_sCODEID = Me.m_sNEWSODEID Then
            com.Azion.EloanUtility.UIUtility.alert("無法在最新訊息頁籤內新增文章，請先點擊其他頁籤後再新增文章")
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "無法在最新訊息頁籤內新增文章，請先點擊其他頁籤後再新增文章")
            Return
        End If

        Dim AP_CODE As New AP_CODE(Me.m_DBManager)
        If AP_CODE.loadByPK(Me.m_sCODEID) Then
            If Utility.isValidateData(AP_CODE.getAttribute("BRID")) Then
                If Utility.isValidateData(Session("BRID")) AndAlso Session("BRID") <> "0" Then
                    If AP_CODE.getString("BRID") <> Session("BRID") Then
                        com.Azion.EloanUtility.UIUtility.alert("您無權限新增本中心文章")
                        com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "您無權限新增本中心文章")
                        Return
                    End If
                End If
            End If
        End If

        Me.PLH_CKEditor.Visible = True

        Me.CKEditorControl1.Text = String.Empty

        Me.LBL_HTML.Text = String.Empty

        'Me.btAdd.Visible = False

        Dim iYYYYMMDD As Decimal = 0
        Dim iHHMMSS As Decimal = 0

        Dim D_DATE As Date = Now

        Me.LTL_YYYYMMDD.Text = D_DATE.Year * 10000 + D_DATE.Month * 100 + D_DATE.Day

        Me.LTL_HHMMSS.Text = D_DATE.Hour * 10000 + D_DATE.Minute * 100 + D_DATE.Second

        Me.HID_CRETIME.Value = String.Empty

        Me.HID_SEQTIME.Value = String.Empty

        Me.HID_INDEX.Value = String.Empty

        Me.RP_NEWS.Visible = False

        Me.HID_CODEID.Value = Me.m_sCODEID
    End Sub

    Private Sub btCancel_Click(sender As Object, e As System.EventArgs) Handles btCancel.Click
        Me.PLH_CKEditor.Visible = False

        Me.CKEditorControl1.Text = String.Empty

        Me.LBL_HTML.Text = String.Empty

        'Me.btAdd.Visible = True

        Me.LTL_YYYYMMDD.Text = String.Empty

        Me.LTL_HHMMSS.Text = String.Empty

        Me.HID_CRETIME.Value = String.Empty

        Me.HID_SEQTIME.Value = String.Empty

        Me.HID_INDEX.Value = String.Empty

        Me.RP_NEWS.Visible = True
    End Sub

    'Sub btPreview_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles btPreview.Click
    '    Me.LBL_HTML.Text = Me.CKEditorControl1.Text
    'End Sub

    Private Sub btSAVE_Click(sender As Object, e As System.EventArgs) Handles btSAVE.Click
        Try
            Me.Save_News()

            Dim iSEQTIME As Decimal = 0
            iSEQTIME = CDec(Me.HID_SEQTIME.Value)

            'If IsNumeric(Me.HID_CRETIME.Value) Then
            '    'Dim objRepeaterItem As RepeaterItem = Me.RP_NEWS.Items(0)
            '    'Dim HID_SORTNO As HtmlInputHidden = objRepeaterItem.FindControl("HID_SORTNO")
            '    Me.Bind_RP_NEWS(CInt(Me.m_sCODEID), CDec(Me.HID_CRETIME.Value), 1)
            'Else
            '    Me.Bind_RP_NEWS(CInt(Me.m_sCODEID), iCRETIME, 1)
            'End If

            Dim iTopSEQTIME As Decimal = 0
            If Me.RP_NEWS.Items.Count > 0 Then
                Dim HID_SEQTIME As HtmlInputHidden = Me.RP_NEWS.Items(0).FindControl("HID_SEQTIME")
                iTopSEQTIME = CDec(HID_SEQTIME.Value) + 1
            Else
                iTopSEQTIME = iSEQTIME + 1
            End If

            If Not com.Azion.EloanUtility.ValidateUtility.isValidateData(Me.m_sCODEID) OrElse Me.m_sCODEID = Me.m_sNEWSODEID Then
                Me.Bind_RP_NEWSByTime(iTopSEQTIME, 1)
            Else
                Me.Bind_RP_NEWS(CInt(Me.m_sCODEID), iTopSEQTIME, 1)
            End If

            Me.PLH_CKEditor.Visible = False

            Me.CKEditorControl1.Text = String.Empty

            Me.CKEditorControl2.Text = String.Empty

            Me.LBL_HTML.Text = String.Empty

            'Me.btAdd.Visible = True

            Me.LTL_YYYYMMDD.Text = String.Empty

            Me.LTL_HHMMSS.Text = String.Empty

            Me.HID_CRETIME.Value = String.Empty

            Me.HID_SEQTIME.Value = String.Empty

            Me.HID_INDEX.Value = String.Empty

            'If iINDEX >= 0 Then
            '    Dim objFocisRepeaterItem As RepeaterItem = Me.RP_NEWS.Items(iINDEX)
            '    Dim IMG_EDIT As ImageButton = objFocisRepeaterItem.FindControl("IMG_EDIT")
            '    Utility.setObjFocus(IMG_EDIT.ClientID, Me.Page)
            'End If
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

    Sub Save_News()
        Try
            Dim iINDEX As Integer = -1
            If IsNumeric(Me.HID_INDEX.Value) Then
                iINDEX = CInt(Me.HID_INDEX.Value)
            End If

            Dim mbNEWS As New MB_NEWS(Me.m_DBManager)
            Dim iCRETIME As Decimal = 0
            Dim iSEQTIME As Decimal = 0
            If Not IsNumeric(Me.HID_CRETIME.Value) Then
                iCRETIME = Me.getNEW_SEQTIME
                iSEQTIME = iCRETIME
            Else
                iCRETIME = CDec(Me.HID_CRETIME.Value)
                iSEQTIME = CDec(Me.HID_SEQTIME.Value)
            End If

            mbNEWS.clear()
            mbNEWS.LoadByPK(iCRETIME)

            mbNEWS.setAttribute("CODEID", CInt(Me.HID_CODEID.Value))
            mbNEWS.setAttribute("CRETIME", iCRETIME)
            If IsNothing(mbNEWS.getAttribute("CHGTIME")) OrElse mbNEWS.getDecimal("CHGTIME") = 0 Then
                mbNEWS.setAttribute("CHGTIME", iCRETIME)
            End If

            Dim sHTML As String = String.Empty
            sHTML = Me.CKEditorControl1.Text

            'If Me.RB_CLASS_YES.Checked AndAlso IsNumeric(Me.HID_MB_SEQ.Value) Then
            '    mbNEWS.setAttribute("MB_SEQ", CDec(Me.HID_MB_SEQ.Value))
            'Else
            '    mbNEWS.setAttribute("MB_SEQ", Nothing)
            'End If

            If Me.RB_CLASS_NO.Checked Then
                mbNEWS.setAttribute("MB_SEQ", Nothing)
            End If

            mbNEWS.setAttribute("CNTHTML", sHTML)

            Dim sDESCHTML As String = String.Empty
            sDESCHTML = Me.CKEditorControl2.Text
            mbNEWS.setAttribute("DESCHTML", sDESCHTML)
            mbNEWS.setAttribute("SEQTIME", iSEQTIME)

            'mbNEWS.setAttribute("EVENT", Trim(Me.EVENT.Text))

            mbNEWS.save()

            Me.HID_CRETIME.Value = iCRETIME
            Me.HID_SEQTIME.Value = iSEQTIME
        Catch ex As Exception
            Throw
        End Try
    End Sub
#End Region

    Function getTIME(ByVal iCRETIME As Decimal) As String
        Try
            'Dim sYEAR As String = String.Empty
            'sYEAR = Left(iYYYYMMDD, 4)
            'Dim sDD As String = String.Empty
            'sDD = Right(iYYYYMMDD, 2)
            'Dim sMM As String = String.Empty
            'sMM = iYYYYMMDD.ToString.Substring(4, 2)

            'Dim sHHMMSS As String = String.Empty
            'sHHMMSS = Utility.FillZero(iHHMMSS, 6)
            'Dim sHH As String = String.Empty
            'sHH = Left(sHHMMSS, 2)

            'Dim sMi As String = String.Empty
            'sMi = Left(Right(sHHMMSS, 4), 2)

            'Dim sSS As String = String.Empty
            'sSS = Right(sHHMMSS, 2)

            'Return sYEAR & "/" & sMM & "/" & sDD & " " & sHH & ":" & sMi & ":" & sSS

            Return Left(iCRETIME, 8) & " " & Left(Right(iCRETIME, 6), 2) & ":" & Left(Right(iCRETIME, 4), 2) & ":" & Right(iCRETIME, 2)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Private Sub IMG_ADD_Click(sender As Object, e As ImageClickEventArgs) Handles IMG_ADD.Click
        Me.btAdd_Click()
    End Sub

    Private Sub RP_CLASS_1_ItemCommand(source As Object, e As RepeaterCommandEventArgs) Handles RP_CLASS_1.ItemCommand, RP_CLASS_4.ItemCommand, RP_CLASS_2.ItemCommand
        Try
            If UCase(e.CommandName) = "CONTENT" Then
                Dim MB_SEQ As LinkButton = e.Item.FindControl("MB_SEQ")
                Dim mbNEWSList As New MB_NEWSList(Me.m_DBManager)
                mbNEWSList.setSQLCondition(" ORDER BY CRETIME DESC ")
                mbNEWSList.LoadbyMB_SEQ(CDec(e.CommandArgument))
                If mbNEWSList.getCurrentDataSet.Tables(0).Rows.Count > 0 Then
                    Dim iCRETIME As Decimal = 0
                    iCRETIME = mbNEWSList.getCurrentDataSet.Tables(0).Rows(0)("CRETIME")

                    Dim sURL As String = String.Empty
                    sURL = com.Azion.EloanUtility.UIUtility.getRootPath & "/NewsList.aspx?CRETIME=" & iCRETIME
                    Response.Redirect(sURL)
                Else
                    com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "該課程尚未刊登文章內容")
                    com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "該課程尚未刊登文章內容")
                    Return
                End If
            End If
        Catch ex As System.Threading.ThreadAbortException
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    'Private Sub LB_CLASS_4_Click(sender As Object, e As EventArgs) Handles LB_CLASS_4.Click
    '    Try
    '        Me.PLH_CLASS_4.Visible = True
    '    Catch ex As Exception
    '        UIShareFun.showErrMsg(Me, ex)
    '    End Try
    'End Sub

    'Private Sub LB_CLASS_2_Click(sender As Object, e As EventArgs) Handles LB_CLASS_2.Click
    '    Try
    '        Me.PLH_CLASS_2.Visible = True
    '    Catch ex As Exception
    '        UIShareFun.showErrMsg(Me, ex)
    '    End Try
    'End Sub

    Private Sub RP_Banner_ItemDataBound(sender As Object, e As Web.UI.WebControls.RepeaterItemEventArgs) Handles RP_Banner.ItemDataBound
        Try
            If e.Item.ItemType = ListItemType.Item OrElse e.Item.ItemType = ListItemType.AlternatingItem Then
                Dim DRV As DataRowView = CType(e.Item.DataItem, DataRowView)
                Dim IMG_BANNER As HtmlImage = e.Item.FindControl("IMG_BANNER")

                IMG_BANNER.Src = com.Azion.EloanUtility.UIUtility.getRootPath & "/ckfinder/userfiles/images/banner/" & DRV("NAME").ToString
                IMG_BANNER.Attributes("title") = Path.GetFileNameWithoutExtension(DRV("NAME").ToString)
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Private Sub btnChoose_Click(sender As Object, e As EventArgs) Handles btnChoose.Click
        Try
            If Not Utility.isValidateData(Me.CKEditorControl1.Text) AndAlso Not Utility.isValidateData(Me.CKEditorControl2.Text) Then
                com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "請先輸入簡介或內文")
                com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請先輸入簡介或內文")
                Return
            End If

            Me.Save_News()

            Dim sURLClass As String = String.Empty
            sURLClass = com.Azion.EloanUtility.UIUtility.getRootPath & "/MNT/MBMnt_Class_01_v01.aspx?OPTYPE=Q&CRETIME=" & Me.HID_CRETIME.Value
            'com.Azion.EloanUtility.UIUtility.showOpen(sURLClass)
            Response.Redirect(sURLClass)
        Catch ex As System.Threading.ThreadAbortException
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub
End Class