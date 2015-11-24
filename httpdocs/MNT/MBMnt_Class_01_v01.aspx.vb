Imports com.Azion.NET.VB
Imports MBSC.MB_OP
Imports MBSC.UICtl

Public Class MBMnt_Class_01_v01
    Inherits System.Web.UI.Page

    Dim m_DBManager As DatabaseManager

    Dim m_sCRETIME As String = String.Empty

#Region "Page Events"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.m_sCRETIME = "" & Request.QueryString("CRETIME")

            Me.m_DBManager = UIShareFun.getDataBaseManager

            If Not Page.IsPostBack Then
                Me.INIT_REGTIME_H()
                Me.Bind_DDL_Type()
                Me.Bind_DDL_Place()
                Me.init_Bind_DDL_DATE()
            End If
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    Private Sub MBMnt_Class_01_v01_Unload(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Unload
        UIShareFun.releaseConnection(Me.m_DBManager)
    End Sub
#End Region

#Region "init"
    Sub INIT_REGTIME_H()
        Try
            Me.REGTIME_H.Items.Clear()
            Me.REGTIME_H.Items.Add(New ListItem("請選擇", String.Empty))
            For i As Integer = 1 To 24
                Dim objItem As New ListItem(Utility.FillZero(i.ToString, 2), i)
                Me.REGTIME_H.Items.Add(objItem)
            Next
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Sub init_Bind_DDL_DATE()
        Try
            Me.Bind_DDL_Date(ddl_StartYear, "1")
            Me.Bind_DDL_Date(ddl_EndYear, "1")
            Me.Bind_DDL_Date(ddl_EditStartYear, "1")
            Me.Bind_DDL_Date(ddl_EditEndYear, "1")
            Me.Bind_DDL_Date(ddl_StartMonth, "2")
            Me.Bind_DDL_Date(ddl_EndMonth, "2")
            Me.Bind_DDL_Date(ddl_EditStartMonth, "2")
            Me.Bind_DDL_Date(ddl_EditEndMonth, "2")
            Me.Bind_DDL_Date(ddl_EditStartDay, "3")
            Me.Bind_DDL_Date(ddl_EditEndDay, "3")

            Me.Bind_DDL_Date(Me.DDL_MB_SAPLY_Y, "1")
            Me.Bind_DDL_Date(Me.DDL_MB_SAPLY_M, "2")
            Me.Bind_DDL_Date(Me.DDL_MB_SAPLY_D, "3")

            Me.Bind_DDL_Date(Me.DDL_MB_EAPLY_Y, "1")
            Me.Bind_DDL_Date(Me.DDL_MB_EAPLY_M, "2")
            Me.Bind_DDL_Date(Me.DDL_MB_EAPLY_D, "3")
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    Sub Bind_DDL_Date(ByVal DDL As DropDownList, ByVal sType As String)
        Try
            DDL.Items.Clear()
            Select Case sType
                Case "1"
                    For i As Integer = Now.Year + 2 To 2013 Step -1
                        DDL.Items.Insert(0, New ListItem(i, i))
                    Next
                Case "2"
                    For i As Integer = 12 To 1 Step -1
                        DDL.Items.Insert(0, New ListItem(i, i))
                    Next
                Case "3"
                    For i As Integer = 31 To 1 Step -1
                        DDL.Items.Insert(0, New ListItem(i, i))
                    Next
            End Select

            DDL.DataTextField = "TEXT"
            DDL.DataValueField = "VALUE"
            DDL.DataBind()

            DDL.Items.Insert(0, New ListItem("請選擇", ""))
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    Sub Bind_DDL_Type()
        Try
            ddl_TYPE.Items.Clear()
            ddl_EditTYPE.Items.Clear()

            Dim AP_CODEList As New AP_CODEList(Me.m_DBManager)

            AP_CODEList.loadByUpCode(91)

            ddl_TYPE.DataSource = AP_CODEList.getCurrentDataSet.Tables(0)
            ddl_TYPE.DataTextField = "TEXT"
            ddl_TYPE.DataValueField = "VALUE"
            ddl_TYPE.DataBind()

            ddl_EditTYPE.DataSource = AP_CODEList.getCurrentDataSet.Tables(0)
            ddl_EditTYPE.DataTextField = "TEXT"
            ddl_EditTYPE.DataValueField = "VALUE"
            ddl_EditTYPE.DataBind()

            ddl_TYPE.Items.Insert(0, New ListItem("請選擇", ""))
            ddl_EditTYPE.Items.Insert(0, New ListItem("請選擇", ""))
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    Sub Bind_DDL_Place()
        Dim DT_177 As DataTable = Nothing
        Try
            ddl_PLACE.Items.Clear()
            ddl_EditPLACE.Items.Clear()

            Dim AP_CODEList As New AP_CODEList(Me.m_DBManager)

            AP_CODEList.loadByUpCode(177)

            DT_177 = AP_CODEList.getCurrentDataSet.Tables(0)
            If Utility.isValidateData(Session("BRID")) AndAlso Session("BRID") <> "0" Then
                Dim ROW_OTHBRID() As DataRow = Nothing
                ROW_OTHBRID = DT_177.Select("VALUE<>'" & Session("BRID") & "'")
                If Not IsNothing(ROW_OTHBRID) AndAlso ROW_OTHBRID.Length > 0 Then
                    For i As Integer = 0 To UBound(ROW_OTHBRID)
                        DT_177.Rows.Remove(ROW_OTHBRID(i))
                    Next
                End If
            End If
            ddl_PLACE.DataSource = DT_177
            ddl_PLACE.DataTextField = "TEXT"
            ddl_PLACE.DataValueField = "VALUE"
            ddl_PLACE.DataBind()

            If Me.ddl_PLACE.Items.Count > 1 Then
                Me.ddl_PLACE.Items.Insert(0, New ListItem("請選擇", ""))
            End If


            ddl_EditPLACE.DataSource = DT_177
            ddl_EditPLACE.DataTextField = "TEXT"
            ddl_EditPLACE.DataValueField = "VALUE"
            ddl_EditPLACE.DataBind()

            If Me.ddl_EditPLACE.Items.Count > 1 Then
                Me.ddl_EditPLACE.Items.Insert(0, New ListItem("請選擇", ""))
            End If
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        Finally
            If Not IsNothing(DT_177) Then
                DT_177.Dispose()
            End If
        End Try
    End Sub

    Sub init_TB1_A()
        Try
            ddl_EditTYPE.SelectedValue = ""
            txt_BATCH.Text = String.Empty
            'txt_BATCH.Visible = False
            ddl_EditPLACE.SelectedIndex = 0
            txt_ClassName.Text = ""
            init_Bind_DDL_DATE()
            txt_TEACHER.Text = ""
            If rbt_APV.SelectedValue <> "" Then
                rbt_APV.SelectedItem.Selected = False
            End If
            txt_FULL.Text = ""
            txt_WAIT.Text = ""
            If rbt_YES.SelectedValue <> "" Then
                rbt_YES.SelectedItem.Selected = False
            End If
            lbl_sSEQ.Text = ""
            txt_MEMO.Text = ""
            Me.MB_CDAYS.Text = String.Empty
            '提醒信期限一/二
            Me.MB_ALERT1_DAY.Text = String.Empty
            Me.MB_ALERT2_DAY.Text = String.Empty
            '報到時間
            Me.REGTIME_H.SelectedIndex = -1
            Me.REGTIME_M.Text = String.Empty
            '聯絡人/電話
            Me.CONTACT.Text = String.Empty
            Me.CONTEL.Text = String.Empty
            '上課地點
            Me.CLASS_PLACE.Text = String.Empty
            '交通資訊說明
            Me.TRAFFIC_DESC.Text = String.Empty
            '是否需填初學者
            Me.MB_BEGIN.SelectedIndex = -1
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub
#End Region

#Region "Button Events"
    Private Sub btn_Confirm_Click(sender As Object, e As System.EventArgs) Handles btn_Confirm.Click
        Try
            Dim MB_CLASSList As New MB_CLASSList(Me.m_DBManager)
            MB_CLASSList.clear()
            Dim sSDATE As String = String.Empty
            Dim sEDATE As String = String.Empty
            Dim sFlag As String = String.Empty

            If ddl_StartYear.SelectedValue <> "" AndAlso ddl_StartMonth.SelectedValue <> "" Then
                sSDATE = (CInt(ddl_StartYear.SelectedValue)).ToString & "-" & ddl_StartMonth.SelectedValue.PadLeft(2, "0") & "-01"
            ElseIf ddl_StartYear.SelectedValue = "" AndAlso ddl_StartMonth.SelectedValue = "" Then
                sSDATE = ""
            Else
                sFlag &= "開課起期未輸入完成\n"
            End If

            If ddl_EndYear.SelectedValue <> "" AndAlso ddl_EndMonth.SelectedValue <> "" Then
                sEDATE = (CInt(ddl_EndYear.SelectedValue)).ToString & "-" & ddl_EndMonth.SelectedValue.PadLeft(2, "0") & "-" & DateTime.DaysInMonth(CInt(ddl_EndYear.SelectedValue), CInt(ddl_EndMonth.SelectedValue))
            ElseIf ddl_EndYear.SelectedValue = "" AndAlso ddl_EndMonth.SelectedValue = "" Then
                sEDATE = ""
            Else
                sFlag &= "開課迄期未輸入完成\n"
            End If

            If sFlag.Trim.Length = 0 AndAlso MB_CLASSList.getClassList(ddl_TYPE.SelectedValue, IIf(ddl_PLACE.SelectedValue <> "", ddl_PLACE.SelectedItem.Text, ""), sSDATE, sEDATE) > 0 Then
                dg_Class.DataSource = MB_CLASSList.getCurrentDataSet.Tables(0)
                dg_Class.DataBind()
                If Request("OPTYPE") = "Q" Then
                    tbl_A.Visible = False
                    btn_QSEQ.Visible = True
                    dg_Class.Columns(0).Visible = True
                    dg_Class.Columns(10).Visible = False
                    dg_Class.Columns(11).Visible = False
                Else
                    btn_QSEQ.Visible = False
                    dg_Class.Columns(0).Visible = False
                    dg_Class.Columns(10).Visible = True
                    dg_Class.Columns(11).Visible = True
                End If
            Else
                dg_Class.DataSource = Nothing
                dg_Class.DataBind()
                btn_QSEQ.Visible = False
                If sFlag.Trim.Length > 0 Then
                    com.Azion.EloanUtility.UIUtility.alert(sFlag)
                End If
            End If

            Me.SH_Modify(False)
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

#End Region

#Region "DateBind"
    Private Sub dg_Class_ItemDataBound(sender As Object, e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles dg_Class.ItemDataBound
        Try
            If e.Item.ItemType = ListItemType.Item Or e.Item.ItemType = ListItemType.AlternatingItem Then
                Dim lbl_APV As Label = CType(e.Item.FindControl("lbl_APV"), Label) '是否需核准
                Dim lbl_OPEN As Label = CType(e.Item.FindControl("lbl_OPEN"), Label) '開放報名
                Dim img_Del As ImageButton = CType(e.Item.FindControl("img_Del"), ImageButton)  '刪除按鈕
                Dim lbl_SDATE As Label = CType(e.Item.FindControl("lbl_SDATE"), Label) '課程起日
                Dim lbl_EDATE As Label = CType(e.Item.FindControl("lbl_EDATE"), Label) '課程訖日

                '課程起日
                lbl_SDATE.Text = lbl_SDATE.Text & " - "

                '課程訖日
                lbl_EDATE.Text = lbl_EDATE.Text

                '是否需核准
                If lbl_APV.Text = "Y" Then
                    lbl_APV.Text = "是"
                Else
                    lbl_APV.Text = "否"
                End If

                '刪除
                If lbl_OPEN.Text = "" Then
                    img_Del.Visible = True
                Else
                    img_Del.Visible = False
                End If

                '梯次
                Dim lbl_BATCH As Label = e.Item.FindControl("lbl_BATCH")
                If lbl_BATCH.Text = "0" Then
                    lbl_BATCH.Text = String.Empty
                End If
            End If
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub
#End Region

#Region "按鈕事件"

    'Sub ddl_EditTYPE_SelectChange(ByVal sender As Object, ByVal e As EventArgs)
    '    Try
    '        If ddl_EditTYPE.SelectedValue = "1" Then
    '            txt_BATCH.Visible = True
    '        Else
    '            txt_BATCH.Text = ""
    '            txt_BATCH.Visible = False
    '        End If
    '    Catch ex As Exception
    '        UIShareFun.showErrMsg(Me, ex)
    '    End Try
    'End Sub

    '修改
    Sub img_Edit_Click(sender As Object, e As ImageClickEventArgs)
        Try
            Dim DataGridItem As DataGridItem = CType(sender, ImageButton).NamingContainer
            Dim sSEQ As String = CType(DataGridItem.FindControl("lbl_SEQ"), Label).Text.Trim
            Dim sMB_BATCH As String = CType(DataGridItem.FindControl("HID_MB_BATCH"), HtmlInputHidden).Value.Trim
            Dim MB_CLASS As New MB_CLASS(Me.m_DBManager)
            If MB_CLASS.loadByPK(CInt(sSEQ), CDec(sMB_BATCH)) Then
                Dim MB_CLASS_M As New MB_CLASS_M(Me.m_DBManager)
                MB_CLASS_M.LoadByPK(CDec(sSEQ))
                ddl_EditTYPE.SelectedIndex = -1
                If Not IsNothing(ddl_EditTYPE.Items.FindByValue(MB_CLASS_M.getString("MB_TYPE_NO"))) Then
                    ddl_EditTYPE.Items.FindByValue(MB_CLASS_M.getString("MB_TYPE_NO")).Selected = True
                End If

                'If MB_CLASS.getString("MB_TYPE_NO") = "1" Then
                '    txt_BATCH.Visible = True
                '    txt_BATCH.Text = MB_CLASS.getString("MB_BATCH")
                'Else
                '    txt_BATCH.Visible = False
                '    txt_BATCH.Text = ""
                'End If
                txt_BATCH.Text = MB_CLASS.getString("MB_BATCH")

                ddl_EditPLACE.SelectedIndex = -1
                If Not IsNothing(ddl_EditPLACE.Items.FindByText(MB_CLASS.getString("MB_PLACE"))) Then
                    ddl_EditPLACE.Items.FindByText(MB_CLASS.getString("MB_PLACE")).Selected = True
                End If

                txt_ClassName.Text = MB_CLASS.getString("MB_CLASS_NAME")

                Dim tStart As Date = CDate(MB_CLASS.getAttribute("MB_SDATE").ToString)
                Me.ddl_EditStartYear.SelectedIndex = -1
                If Not IsNothing(ddl_EditStartYear.Items.FindByValue((tStart.Year).ToString)) Then
                    ddl_EditStartYear.Items.FindByValue((tStart.Year).ToString).Selected = True
                End If

                ddl_EditStartMonth.SelectedIndex = -1
                If Not IsNothing(ddl_EditStartMonth.Items.FindByValue(tStart.Month.ToString)) Then
                    ddl_EditStartMonth.Items.FindByValue(tStart.Month.ToString).Selected = True
                End If

                ddl_EditStartDay.SelectedIndex = -1
                If Not IsNothing(ddl_EditStartDay.Items.FindByValue(tStart.Day.ToString)) Then
                    ddl_EditStartDay.Items.FindByValue(tStart.Day.ToString).Selected = True
                End If

                Dim tEnd As Date = CDate(MB_CLASS.getAttribute("MB_EDATE").ToString)
                ddl_EditEndYear.SelectedIndex = -1
                If Not IsNothing(ddl_EditEndYear.Items.FindByValue((tEnd.Year).ToString)) Then
                    ddl_EditEndYear.Items.FindByValue((tEnd.Year).ToString).Selected = True
                End If

                ddl_EditEndMonth.SelectedIndex = -1
                If Not IsNothing(ddl_EditEndMonth.Items.FindByValue(tEnd.Month.ToString)) Then
                    ddl_EditEndMonth.Items.FindByValue(tEnd.Month.ToString).Selected = True
                End If

                ddl_EditEndDay.SelectedIndex = -1
                If Not IsNothing(ddl_EditEndDay.Items.FindByValue(tEnd.Day.ToString)) Then
                    ddl_EditEndDay.Items.FindByValue(tEnd.Day.ToString).Selected = True
                End If

                Me.DDL_MB_SAPLY_Y.SelectedIndex = -1
                Me.DDL_MB_SAPLY_M.SelectedIndex = -1
                Me.DDL_MB_SAPLY_D.SelectedIndex = -1
                If Utility.isValidateData(MB_CLASS.getAttribute("MB_SAPLY")) AndAlso IsDate(MB_CLASS.getAttribute("MB_SAPLY").ToString) AndAlso CDate(MB_CLASS.getAttribute("MB_SAPLY").ToString).Year > 1911 Then
                    Dim D_MB_SAPLY As Date = CDate(MB_CLASS.getAttribute("MB_SAPLY").ToString)

                    If Not IsNothing(Me.DDL_MB_SAPLY_Y.Items.FindByValue(D_MB_SAPLY.Year)) Then
                        Me.DDL_MB_SAPLY_Y.Items.FindByValue(D_MB_SAPLY.Year).Selected = True
                    End If

                    If Not IsNothing(Me.DDL_MB_SAPLY_M.Items.FindByValue(D_MB_SAPLY.Month)) Then
                        Me.DDL_MB_SAPLY_M.Items.FindByValue(D_MB_SAPLY.Month).Selected = True
                    End If

                    If Not IsNothing(Me.DDL_MB_SAPLY_D.Items.FindByValue(D_MB_SAPLY.Day)) Then
                        Me.DDL_MB_SAPLY_D.Items.FindByValue(D_MB_SAPLY.Day).Selected = True
                    End If
                End If

                Me.DDL_MB_EAPLY_Y.SelectedIndex = -1
                Me.DDL_MB_EAPLY_M.SelectedIndex = -1
                Me.DDL_MB_EAPLY_D.SelectedIndex = -1
                If Utility.isValidateData(MB_CLASS.getAttribute("MB_EAPLY")) AndAlso IsDate(MB_CLASS.getAttribute("MB_EAPLY").ToString) AndAlso CDate(MB_CLASS.getAttribute("MB_EAPLY").ToString).Year > 1911 Then
                    Dim D_MB_EAPLY As Date = CDate(MB_CLASS.getAttribute("MB_EAPLY").ToString)

                    If Not IsNothing(Me.DDL_MB_EAPLY_Y.Items.FindByValue(D_MB_EAPLY.Year)) Then
                        Me.DDL_MB_EAPLY_Y.Items.FindByValue(D_MB_EAPLY.Year).Selected = True
                    End If

                    If Not IsNothing(Me.DDL_MB_EAPLY_M.Items.FindByValue(D_MB_EAPLY.Month)) Then
                        Me.DDL_MB_EAPLY_M.Items.FindByValue(D_MB_EAPLY.Month).Selected = True
                    End If

                    If Not IsNothing(Me.DDL_MB_EAPLY_D.Items.FindByValue(D_MB_EAPLY.Day)) Then
                        Me.DDL_MB_EAPLY_D.Items.FindByValue(D_MB_EAPLY.Day).Selected = True
                    End If
                End If

                txt_TEACHER.Text = MB_CLASS.getString("MB_TEACHER")
                rbt_APV.SelectedValue = MB_CLASS.getString("MB_APV")
                txt_FULL.Text = MB_CLASS.getString("MB_FULL")
                txt_WAIT.Text = MB_CLASS.getString("MB_WAIT")
                rbt_YES.SelectedValue = MB_CLASS.getString("MB_YES")
                txt_MEMO.Text = MB_CLASS.getString("MB_MEMO")
                Me.MB_CDAYS.Text = MB_CLASS.getString("MB_CDAYS")
                '提醒信期限一/二
                If Utility.isValidateData(MB_CLASS.getAttribute("MB_ALERT1_DAY")) Then
                    Me.MB_ALERT1_DAY.Text = MB_CLASS.getAttribute("MB_ALERT1_DAY")
                Else
                    If MB_CLASS.isLoaded Then
                        Me.MB_ALERT1_DAY.Text = String.Empty
                    Else
                        Me.MB_ALERT1_DAY.Text = "10"
                    End If
                End If
                If Utility.isValidateData(MB_CLASS.getAttribute("MB_ALERT2_DAY")) Then
                    Me.MB_ALERT2_DAY.Text = MB_CLASS.getAttribute("MB_ALERT2_DAY")
                Else
                    If MB_CLASS.isLoaded Then
                        Me.MB_ALERT2_DAY.Text = String.Empty
                    Else
                        Me.MB_ALERT2_DAY.Text = "3"
                    End If
                End If
                '報到時間
                Me.REGTIME_H.SelectedIndex = -1
                If Utility.isValidateData(MB_CLASS.getAttribute("REGTIME")) Then
                    Dim sREGTIME As String = String.Empty
                    sREGTIME = Utility.FillZero(MB_CLASS.getAttribute("REGTIME"), 4)
                    If Not IsNothing(Me.REGTIME_H.Items.FindByValue(CDec(Left(sREGTIME, 2)))) Then
                        Me.REGTIME_H.Items.FindByValue(CDec(Left(sREGTIME, 2))).Selected = True
                    End If

                    Me.REGTIME_M.Text = CDec(Right(sREGTIME, 2))
                End If
                '聯絡人/電話
                Me.CONTACT.Text = MB_CLASS.getString("CONTACT")
                Me.CONTEL.Text = MB_CLASS.getString("CONTEL")
                '上課地點
                Me.CLASS_PLACE.Text = MB_CLASS.getString("CLASS_PLACE")
                '交通資訊說明
                Me.TRAFFIC_DESC.Text = MB_CLASS.getString("TRAFFIC_DESC")
                '是否需填初學者
                Me.MB_BEGIN.SelectedIndex = -1
                If Not IsNothing(Me.MB_BEGIN.Items.FindByValue(MB_CLASS.getString("MB_BEGIN"))) Then
                    Me.MB_BEGIN.Items.FindByValue(MB_CLASS.getString("MB_BEGIN")).Selected = True
                End If
                lbl_sSEQ.Text = sSEQ
            End If

            Me.SH_Modify(True)
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    '刪除
    Sub img_Del_Click(sender As Object, e As ImageClickEventArgs)
        Try
            Dim griditem As DataGridItem = CType(sender, ImageButton).NamingContainer
            Dim sSEQ As String = CType(griditem.FindControl("lbl_SEQ"), Label).Text.Trim
            Dim sMB_BATCH As String = CType(griditem.FindControl("HID_MB_BATCH"), HtmlInputHidden).Value.Trim
            Try
                Me.m_DBManager.beginTran()

                Dim MB_CLASS As New MB_CLASS(Me.m_DBManager)
                If MB_CLASS.loadByPK(CInt(sSEQ), CDec(sMB_BATCH)) Then
                    MB_CLASS.remove()
                End If

                Dim mbCLASSList As New MB_CLASSList(Me.m_DBManager)
                Dim iSize As Decimal = mbCLASSList.LoadByMB_SEQ(CDec(sSEQ))
                If iSize = 0 Then
                    Dim mbCLASS_M As New MB_CLASS_M(Me.m_DBManager)
                    If mbCLASS_M.LoadByPK(CDec(sSEQ)) Then
                        mbCLASS_M.remove()
                    End If
                End If

                Me.m_DBManager.commit()
            Catch ex As Exception
                Me.m_DBManager.Rollback()
                Throw
            End Try

            com.Azion.EloanUtility.UIUtility.alert("刪除成功")

            Me.btn_Confirm_Click(sender, e)
            Me.SH_Modify(False)
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    Private Sub btn_Add_Click(sender As Object, e As System.EventArgs) Handles btn_Add.Click
        Try
            If CheckMB_CLASS() Then
                Try
                    Me.m_DBManager.beginTran()

                    Dim iMB_SEQ As Decimal = -1
                    iMB_SEQ = Me.Save_MB_CLASS("")

                    Me.SAVE_MB_CLASS_M(iMB_SEQ)

                    Me.m_DBManager.commit()
                Catch ex As Exception
                    Me.m_DBManager.Rollback()
                    Throw
                End Try

                com.Azion.EloanUtility.UIUtility.alert("新增成功")
                Me.btn_Confirm_Click(sender, e)
            End If
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    Sub SAVE_MB_CLASS_M(ByVal iMB_SEQ As Decimal)
        Try
            Dim MB_CLASS_M As New MB_CLASS_M(Me.m_DBManager)
            MB_CLASS_M.LoadByPK(iMB_SEQ)

            'MB_SEQ	decimal(7,0)		NO	PRI	0		select,insert,update,references	課程序號
            MB_CLASS_M.setAttribute("MB_SEQ", iMB_SEQ)

            'MB_TYPE	varchar(50)	utf8_general_ci	YES				select,insert,update,references	課程類別名稱
            MB_CLASS_M.setAttribute("MB_TYPE", Me.ddl_EditTYPE.SelectedItem.Text)

            'MB_TYPE_NO	varchar(2)	utf8_general_ci	YES	MUL			select,insert,update,references	課程類別編號
            MB_CLASS_M.setAttribute("MB_TYPE_NO", Me.ddl_EditTYPE.SelectedValue)

            'MB_CHGUID	varchar(100)	utf8_general_ci	YES				select,insert,update,references	修改人
            MB_CLASS_M.setAttribute("MB_CHGUID", com.Azion.EloanUtility.UIUtility.getLoginUserID)

            'MB_CHGDATE	date		YES				select,insert,update,references	修改日期
            MB_CLASS_M.setAttribute("MB_CHGDATE", Now)

            MB_CLASS_M.save()
        Catch ex As Exception
            Throw
        End Try
    End Sub

    '課程序號
    Function getVadID(ByVal dbManager As com.Azion.NET.VB.DatabaseManager) As String
        Try
            Dim sProcName As String = String.Empty
            sProcName = LCase(com.Azion.NET.VB.Properties.getSchemaName) & ".MB_GET_MAXID"
            Dim inParaAL As New ArrayList
            Dim outParaAL As New ArrayList
            inParaAL.Add("01")
            inParaAL.Add("3")
            outParaAL.Add(7)
            Dim HT_Return As Hashtable = com.Azion.NET.VB.DBObject.getProcedureData(dbManager, sProcName, inParaAL, outParaAL)
            Dim iMAXID As Decimal = HT_Return.Item("@IMAXID")
            Return iMAXID
            'Return com.Azion.EloanUtility.StrUtility.FillZero(iMAXID, 7)
        Catch ex As Exception
            Throw
        Finally
        End Try
    End Function


    Function Save_MB_CLASS(ByVal sSEQ As String) As Decimal
        Try
            Dim MB_CLASS As New MB_CLASS(Me.m_DBManager)
            Dim dSEQ As Decimal = -1

            If sSEQ = "" Then
                dSEQ = getVadID(Me.m_DBManager)
            Else
                dSEQ = CInt(sSEQ)
            End If

            If Not IsNumeric(Me.txt_BATCH.Text) Then
                Me.txt_BATCH.Text = "0"
            End If

            MB_CLASS.loadByPK(dSEQ, CDec(Me.txt_BATCH.Text))

            MB_CLASS.setAttribute("MB_SEQ", dSEQ)   '課程序號
            'MB_CLASS.setAttribute("MB_TYPE", ddl_EditTYPE.SelectedItem.Text)      '課程類別名稱
            'MB_CLASS.setAttribute("MB_TYPE_NO", ddl_EditTYPE.SelectedValue)       '課程類別編號

            'If ddl_EditTYPE.SelectedValue = "1" Then
            '    MB_CLASS.setAttribute("MB_BATCH", CDec(txt_BATCH.Text.Trim))                 '梯次
            'End If
            MB_CLASS.setAttribute("MB_BATCH", CDec(txt_BATCH.Text.Trim))

            MB_CLASS.setAttribute("MB_PLACE", ddl_EditPLACE.SelectedItem.Text)               '課程地點
            MB_CLASS.setAttribute("MB_SDATE", CDate(CDec(ddl_EditStartYear.SelectedValue) & "-" & CDec(ddl_EditStartMonth.SelectedValue) & "-" & CDec(ddl_EditStartDay.SelectedValue)))              '課程起日
            MB_CLASS.setAttribute("MB_EDATE", CDate(CDec(ddl_EditEndYear.SelectedValue) & "-" & CDec(ddl_EditEndMonth.SelectedValue) & "-" & CDec(ddl_EditEndDay.SelectedValue)))              '課程起訖日
            MB_CLASS.setAttribute("MB_SWEEK", Weekday(CDate(CDec(ddl_EditStartYear.SelectedValue) & "-" & CDec(ddl_EditStartMonth.SelectedValue) & "-" & CDec(ddl_EditStartDay.SelectedValue)), Microsoft.VisualBasic.FirstDayOfWeek.Monday))              '課程起日(星期)
            MB_CLASS.setAttribute("MB_EWEEK", Weekday(CDate(CDec(ddl_EditEndYear.SelectedValue) & "-" & CDec(ddl_EditEndMonth.SelectedValue) & "-" & CDec(ddl_EditEndDay.SelectedValue)), Microsoft.VisualBasic.FirstDayOfWeek.Monday))         '課程訖日(星期)
            MB_CLASS.setAttribute("MB_CLASS_NAME", txt_ClassName.Text)          '課程名稱
            MB_CLASS.setAttribute("MB_TEACHER", txt_TEACHER.Text)              '指導老師
            MB_CLASS.setAttribute("MB_MEMO", txt_MEMO.Text)              '備註說明
            MB_CLASS.setAttribute("MB_YES", rbt_YES.SelectedValue)                 '是否開課  Y: 是N:否
            MB_CLASS.setAttribute("MB_FULL", CInt(txt_FULL.Text.Trim))               '額滿人數
            MB_CLASS.setAttribute("MB_WAIT", CInt(txt_WAIT.Text.Trim))             '備取人數
            MB_CLASS.setAttribute("MB_APV", rbt_APV.SelectedValue)                           '是否需核准
            If IsNumeric(Me.MB_CDAYS.Text) Then
                MB_CLASS.setAttribute("MB_CDAYS", CDec(Me.MB_CDAYS.Text))
            Else
                MB_CLASS.setAttribute("MB_CDAYS", Nothing)
            End If

            'MB_SAPLY	date		YES				select,insert,update,references	報名起日
            Dim sMB_SAPLY As String = String.Empty
            If IsNumeric(Me.DDL_MB_SAPLY_Y.SelectedValue) Then
                sMB_SAPLY = CDec(Me.DDL_MB_SAPLY_Y.SelectedValue)
            End If
            If IsNumeric(Me.DDL_MB_SAPLY_M.SelectedValue) Then
                sMB_SAPLY &= "/" & Me.DDL_MB_SAPLY_M.SelectedValue
            End If
            If IsNumeric(Me.DDL_MB_SAPLY_D.SelectedValue) Then
                sMB_SAPLY &= "/" & Me.DDL_MB_SAPLY_D.SelectedValue
            End If
            MB_CLASS.setAttribute("MB_SAPLY", CDate(sMB_SAPLY))

            'MB_EAPLY	date		YES				select,insert,update,references	報名訖日
            Dim sMB_EAPLY As String = String.Empty
            If IsNumeric(Me.DDL_MB_EAPLY_Y.SelectedValue) Then
                sMB_EAPLY = CDec(Me.DDL_MB_EAPLY_Y.SelectedValue)
            End If
            If IsNumeric(Me.DDL_MB_EAPLY_M.SelectedValue) Then
                sMB_EAPLY &= "/" & Me.DDL_MB_EAPLY_M.SelectedValue
            End If
            If IsNumeric(Me.DDL_MB_EAPLY_D.SelectedValue) Then
                sMB_EAPLY &= "/" & Me.DDL_MB_EAPLY_D.SelectedValue
            End If
            MB_CLASS.setAttribute("MB_EAPLY", CDate(sMB_EAPLY))

            '提醒信一期限
            MB_CLASS.setAttribute("MB_ALERT1_DAY", CDec(Me.MB_ALERT1_DAY.Text))
            '提醒信二期限
            MB_CLASS.setAttribute("MB_ALERT2_DAY", CDec(Me.MB_ALERT2_DAY.Text))
            '報到時間
            Dim iREGTIME As Decimal = 0
            iREGTIME = CDec(Me.REGTIME_H.SelectedValue) * 100 + CDec(Me.REGTIME_M.Text)
            MB_CLASS.setAttribute("REGTIME", iREGTIME)
            '聯絡人
            MB_CLASS.setAttribute("CONTACT", Me.CONTACT.Text)
            '聯絡人電話
            MB_CLASS.setAttribute("CONTEL", Me.CONTEL.Text)
            '上課地點
            MB_CLASS.setAttribute("CLASS_PLACE", Me.CLASS_PLACE.Text)
            '交通資訊說明
            MB_CLASS.setAttribute("TRAFFIC_DESC", Me.TRAFFIC_DESC.Text)
            '是否初學者;Y是N:否
            MB_CLASS.setAttribute("MB_BEGIN", Me.MB_BEGIN.SelectedValue)

            MB_CLASS.save()

            Return dSEQ
        Catch ex As Exception
            Throw
        End Try
    End Function

    Function CheckMB_CLASS() As Boolean
        Try
            If ddl_EditTYPE.SelectedValue = "" Then
                com.Azion.EloanUtility.UIUtility.alert("請輸入類別")
                Return False
            End If

            'If ddl_EditTYPE.SelectedValue = "1" AndAlso Not IsNumeric(txt_BATCH.Text.Trim) Then
            '    com.Azion.EloanUtility.UIUtility.alert("梯次請輸入數字")
            '    Return False
            'End If
            If IsNumeric(txt_BATCH.Text.Trim) AndAlso CDec(Me.txt_BATCH.Text) < 0 Then
                com.Azion.EloanUtility.UIUtility.alert("梯次請輸入數字，且不可小於0")
                Return False
            End If

            If ddl_EditPLACE.SelectedItem.Text = "" Then
                com.Azion.EloanUtility.UIUtility.alert("請輸入地點")
                Return False
            End If

            If txt_ClassName.Text.Trim = "" Then
                com.Azion.EloanUtility.UIUtility.alert("請輸入課程名稱")
                Return False
            End If

            If (ddl_EditStartYear.SelectedValue = "" OrElse ddl_EditStartMonth.SelectedValue = "" OrElse ddl_EditStartDay.SelectedValue = "" OrElse ddl_EditEndYear.SelectedValue = "" OrElse ddl_EditEndMonth.SelectedValue = "" OrElse ddl_EditEndDay.SelectedValue = "") Then
                com.Azion.EloanUtility.UIUtility.alert("請輸入完整開課日期")
                Return False
            ElseIf DateTime.DaysInMonth(CInt(ddl_EditEndYear.SelectedValue), CInt(ddl_EditEndMonth.SelectedValue)) < CInt(ddl_EditStartDay.SelectedValue) OrElse DateTime.DaysInMonth(CInt(ddl_EditEndYear.SelectedValue), CInt(ddl_EditEndMonth.SelectedValue)) < CInt(ddl_EditEndDay.SelectedValue) Then
                com.Azion.EloanUtility.UIUtility.alert("請檢核開課日期")
                Return False
                '1040720 AMY 不檢核開課日期起日不可小於等於今天
                'ElseIf (CInt(ddl_EditStartYear.SelectedValue) * 10000 + CInt(ddl_EditStartMonth.SelectedValue) * 100 + CInt(ddl_EditStartDay.SelectedValue)) <= (Now.Year) * 10000 + Now.Month * 100 + Now.Day Then
                '    com.Azion.EloanUtility.UIUtility.alert("開課日期起日不可小於等於今天")
                '    Return False
            ElseIf CInt(ddl_EditStartYear.SelectedValue) * 10000 + CInt(ddl_EditStartMonth.SelectedValue) * 100 + CInt(ddl_EditStartDay.SelectedValue) > CInt(ddl_EditEndYear.SelectedValue) * 10000 + CInt(ddl_EditEndMonth.SelectedValue) * 100 + CInt(ddl_EditEndDay.SelectedValue) Then
                com.Azion.EloanUtility.UIUtility.alert("開課日期起日不可大於開課日期訖日")
                Return False
            End If

            Dim sMB_SDATE As String = String.Empty
            If IsNumeric(Me.ddl_EditStartYear.SelectedValue) Then
                sMB_SDATE = CDec(Me.ddl_EditStartYear.SelectedValue)
            End If
            If IsNumeric(Me.ddl_EditStartMonth.SelectedValue) Then
                sMB_SDATE &= "/" & Me.ddl_EditStartMonth.SelectedValue
            End If
            If IsNumeric(Me.ddl_EditStartDay.SelectedValue) Then
                sMB_SDATE &= "/" & Me.ddl_EditStartDay.SelectedValue
            End If

            Dim sMB_EDATE As String = String.Empty
            If IsNumeric(Me.ddl_EditEndYear.SelectedValue) Then
                sMB_EDATE = CDec(Me.ddl_EditEndYear.SelectedValue)
            End If
            If IsNumeric(Me.ddl_EditEndMonth.SelectedValue) Then
                sMB_EDATE &= CDec(Me.ddl_EditEndMonth.SelectedValue)
            End If
            If IsNumeric(Me.ddl_EditEndDay.SelectedValue) Then
                sMB_EDATE &= CDec(Me.ddl_EditEndDay.SelectedValue)
            End If

            Dim sMB_SAPLY As String = String.Empty
            If IsNumeric(Me.DDL_MB_SAPLY_Y.SelectedValue) Then
                sMB_SAPLY = CDec(Me.DDL_MB_SAPLY_Y.SelectedValue)
            End If
            If IsNumeric(Me.DDL_MB_SAPLY_M.SelectedValue) Then
                sMB_SAPLY &= "/" & Me.DDL_MB_SAPLY_M.SelectedValue
            End If
            If IsNumeric(Me.DDL_MB_SAPLY_D.SelectedValue) Then
                sMB_SAPLY &= "/" & Me.DDL_MB_SAPLY_D.SelectedValue
            End If
            If Not IsDate(sMB_SAPLY) Then
                com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "請輸入完整課程報名起日")
                com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請輸入完整課程報名起日")
                Return False
            Else
                '1040720 AMY 不檢核開課日期起日不可小於等於今天
                '新增時 課程報名起 >= 今天 and
                'Dim D_NOW As New Date(Now.Year, Now.Month, Now.Day)
                'If CDate(sMB_SAPLY) < D_NOW Then
                '    com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "課程報名起日需大於今天")
                '    com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "課程報名起日需大於今天")
                '    Return False
                'End If

                '課程報名起日 >  MB_CLASS.MB_SDATE (課程時間起)
                If CDate(sMB_SAPLY) > CDate(sMB_SDATE) Then
                    com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "課程報名起日需小於課程時間起日")
                    com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "課程報名起日需小於課程時間起日")
                    Return False
                End If
            End If

            Dim sMB_EAPLY As String = String.Empty
            If IsNumeric(Me.DDL_MB_EAPLY_Y.SelectedValue) Then
                sMB_EAPLY = CDec(Me.DDL_MB_EAPLY_Y.SelectedValue)
            End If
            If IsNumeric(Me.DDL_MB_EAPLY_M.SelectedValue) Then
                sMB_EAPLY &= "/" & Me.DDL_MB_EAPLY_M.SelectedValue
            End If
            If IsNumeric(Me.DDL_MB_EAPLY_D.SelectedValue) Then
                sMB_EAPLY &= "/" & Me.DDL_MB_EAPLY_D.SelectedValue
            End If
            If Not IsDate(sMB_EAPLY) Then
                com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "請輸入完整課程報名訖日")
                com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請輸入完整課程報名訖日")
                Return False
            End If

            If CDate(sMB_SAPLY) > CDate(sMB_EAPLY) Then
                com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "報名起日不可大於報名訖日")
                com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "報名起日不可大於報名訖日")
                Return False
            End If

            If txt_TEACHER.Text.Trim = "" Then
                com.Azion.EloanUtility.UIUtility.alert("請輸入指導老師")
                Return False
            End If

            If rbt_APV.SelectedValue = "" Then
                com.Azion.EloanUtility.UIUtility.alert("請選擇是否需核准")
                Return False
            End If

            If txt_FULL.Text.Trim = "" Then
                com.Azion.EloanUtility.UIUtility.alert("請輸入額滿人數")
                Return False
            ElseIf Not IsNumeric(txt_FULL.Text.Trim) Then
                com.Azion.EloanUtility.UIUtility.alert("額滿人數請輸入數字")
                Return False
            End If

            If txt_WAIT.Text.Trim = "" Then
                com.Azion.EloanUtility.UIUtility.alert("請輸入備取人數")
                Return False
            ElseIf Not IsNumeric(txt_WAIT.Text.Trim) Then
                com.Azion.EloanUtility.UIUtility.alert("備取人數請輸入數字")
                Return False
            End If

            If Not IsNumeric(Me.MB_CDAYS.Text) Then
                com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "請輸入取消報名期限")
                com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請輸入取消報名期限")
                Return False
            End If

            '提醒信1期限
            If Not IsNumeric(Me.MB_ALERT1_DAY.Text) Then
                com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "請輸入提醒信1期限")
                com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請輸入提醒信1期限")
                Return False
            End If
            '提醒信2期限
            If Not IsNumeric(Me.MB_ALERT2_DAY.Text) Then
                com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "請輸入提醒信2期限")
                com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請輸入提醒信2期限")
                Return False
            End If
            If CDec(Me.MB_ALERT2_DAY.Text) >= CDec(Me.MB_ALERT1_DAY.Text) Then
                com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "提醒信二期限必須小於提醒信一期限")
                com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "提醒信二期限必須小於提醒信一期限")
                Return False
            End If
            '報到時間
            If Not IsNumeric(Me.REGTIME_H.SelectedValue) OrElse Not IsNumeric(Me.REGTIME_M.Text) Then
                com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "請輸入報到時間")
                com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請輸入報到時間")
                Return False
            End If
            '聯絡人
            If Not Utility.isValidateData(Me.CONTACT.Text) Then
                com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "請輸入聯絡人")
                com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請輸入聯絡人")
                Return False
            End If
            '聯絡人電話
            If Not Utility.isValidateData(Me.CONTEL.Text) Then
                com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "請輸入聯絡人電話")
                com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請輸入聯絡人電話")
                Return False
            End If
            '上課地點
            If Not Utility.isValidateData(Me.CLASS_PLACE.Text) Then
                com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "請輸入上課地點")
                com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請輸入上課地點")
                Return False
            End If
            '交通資訊說明
            If Not Utility.isValidateData(Me.TRAFFIC_DESC.Text) Then
                com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "請輸入交通資訊說明")
                com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請輸入交通資訊說明")
                Return False
            End If

            '是否初學者
            If Not Utility.isValidateData(Me.MB_BEGIN.SelectedValue) Then
                com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "請選擇是否需填初學者")
                com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請選擇是否需填初學者")
                Return False
            End If

            Return True
        Catch ex As Exception
            Throw
        End Try
    End Function

    '確定修改
    Protected Sub btn_Edit_Click(sender As Object, e As EventArgs) Handles btn_Edit.Click
        Try
            If CheckMB_CLASS() Then
                Try
                    Me.m_DBManager.beginTran()

                    Me.Save_MB_CLASS(lbl_sSEQ.Text)

                    Me.m_DBManager.commit()
                    com.Azion.EloanUtility.UIUtility.alert("修改成功")
                    Me.btn_Confirm_Click(sender, e)
                Catch ex As Exception
                    Me.m_DBManager.Rollback()
                    Throw
                End Try
            End If
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    Sub SH_Modify(ByVal isShowModify As Boolean)
        Try
            Me.btn_Add.Visible = Not isShowModify
            Me.btn_Batch.Visible = isShowModify
            Me.btn_Type.Visible = isShowModify
            Me.btn_Edit.Visible = isShowModify

            If Not isShowModify Then
                Me.init_TB1_A()
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub

    '選擇課程
    Sub btn_QSEQ_Click(sender As Object, e As EventArgs) Handles btn_QSEQ.Click
        Try
            Dim sSEQ As String = String.Empty
            For i As Integer = 0 To dg_Class.Items.Count - 1
                Dim rbt_SEQ As RadioButton = dg_Class.Items(i).FindControl("rbt_SEQ")
                If rbt_SEQ.Checked = True Then
                    Dim lbl_SEQ As Label = dg_Class.Items(i).FindControl("lbl_SEQ")
                    sSEQ = lbl_SEQ.Text
                    Exit For
                End If
            Next

            Dim mbNEWS As New MB_NEWS(Me.m_DBManager)
            If mbNEWS.LoadByPK(CDec(Me.m_sCRETIME)) Then
                mbNEWS.setAttribute("MB_SEQ", CDec(sSEQ))
                mbNEWS.setAttribute("CHGUID", Now)
                mbNEWS.save()

                com.Azion.EloanUtility.UIUtility.alert("課程編號【" & sSEQ & "】已選擇")
            Else
                com.Azion.EloanUtility.UIUtility.alert("課程編號【" & sSEQ & "】無法連結")
            End If

            Dim sJscript As String = String.Empty
            sJscript = "<script language='javascript'>" & vbCrLf
            'sJscript &= "window.returnValue='" & sSEQ & "';" & vbCrLf
            sJscript &= "window.close();" & vbCrLf
            sJscript &= "</script" & ">" & vbCrLf

            Response.Write(sJscript)
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub
#End Region

    Function ChangeDate(ByVal dDate As Object) As String
        Try
            Dim sDate As String

            If IsDate(dDate.ToString) Then
                sDate = dDate.year & "." & dDate.Month & "." & dDate.Day
            Else
                sDate = ""
            End If

            Return sDate
        Catch ex As Exception
            Throw
        End Try
    End Function

    Private Sub btn_Batch_Click(sender As Object, e As EventArgs) Handles btn_Batch.Click
        Try
            If Me.CheckMB_CLASS() Then
                Dim mbCLASS As New MB_CLASS(Me.m_DBManager)
                If mbCLASS.loadByPK(CDec(Me.lbl_sSEQ.Text), CDec(Me.txt_BATCH.Text)) Then
                    com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "第【" & Me.txt_BATCH.Text & "】梯次已經存在,不可新增,只能修改")
                    com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "第【" & Me.txt_BATCH.Text & "】梯次已經存在,不可新增,只能修改")
                    Return
                Else
                    Me.Save_MB_CLASS(Me.lbl_sSEQ.Text)

                    com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "第【" & Me.txt_BATCH.Text & "】梯次已新增")
                    com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "第【" & Me.txt_BATCH.Text & "】梯次已新增")

                    Me.btn_Confirm_Click(sender, e)
                End If
            End If
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

    Private Sub btn_Type_Click(sender As Object, e As EventArgs) Handles btn_Type.Click
        Try
            If Not Utility.isValidateData(Me.ddl_EditTYPE.SelectedValue) Then
                com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "請選擇類別")
                com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請選擇類別")
                Return
            End If

            Dim mbCLASS_M As New MB_CLASS_M(Me.m_DBManager)
            If mbCLASS_M.LoadByPK(CDec(Me.lbl_sSEQ.Text)) Then
                'MB_TYPE	varchar(50)	utf8_general_ci	YES				select,insert,update,references	課程類別名稱
                mbCLASS_M.setAttribute("MB_TYPE", Me.ddl_EditTYPE.SelectedItem.Text)
                'MB_TYPE_NO	varchar(2)	utf8_general_ci	YES	MUL			select,insert,update,references	課程類別編號
                mbCLASS_M.setAttribute("MB_TYPE_NO", Me.ddl_EditTYPE.SelectedValue)
                'MB_CHGUID	varchar(100)	utf8_general_ci	YES				select,insert,update,references	修改人
                mbCLASS_M.setAttribute("MB_CHGUID", com.Azion.EloanUtility.UIUtility.getLoginUserID)
                'MB_CHGDATE	date		YES				select,insert,update,references	修改日期
                mbCLASS_M.setAttribute("MB_CHGDATE", Now)
                mbCLASS_M.save()

                com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "儲存成功")
            End If

            Me.btn_Confirm_Click(sender, e)
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub
End Class