Imports MBSC.MB_OP
Imports com.Azion.NET.VB
Imports MBSC.UICtl

Public Class MBQry_TYPE_01_v01
    Inherits System.Web.UI.Page

    Dim m_DBManager As com.Azion.NET.VB.DatabaseManager
    Dim m_sUpCode7 As String = String.Empty
    Dim m_PageSize As Decimal = 15
    Dim m_sUpCode28 As String = String.Empty
    Dim m_DT_UpCode28 As DataTable = Nothing
    Dim m_DT_UpCode7 As DataTable = Nothing
    Dim m_sUpcode15 As String = String.Empty
    Dim DT_UPCODE15 As DataTable = Nothing
    Dim m_sUpcode23 As String = String.Empty
    Dim m_DT_UpCode23 As DataTable = Nothing
    Dim m_sQTYPE As String = String.Empty

#Region "Page Events"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.m_DBManager = UIShareFun.getDataBaseManager

            '所屬區代碼
            Me.m_sUpCode7 = com.Azion.EloanUtility.CodeList.getAppSettings("UpCode7")

            '會員類別
            Me.m_sUpCode28 = com.Azion.EloanUtility.CodeList.getAppSettings("UpCode28")

            '功德項目
            Me.m_sUpcode15 = com.Azion.EloanUtility.CodeList.getAppSettings("UpCode15")

            '繳款方式
            m_sUpcode23 = com.Azion.EloanUtility.CodeList.getAppSettings("UpCode23")

            Me.m_sQTYPE = "" & Request.QueryString("QTYPE")

            If Not Page.IsPostBack Then
                If Me.m_sQTYPE = "1" Then
                    Me.DDL_TYPE.Items.Clear()
                    Me.DDL_TYPE.Items.Add(New ListItem("請選擇", ""))
                    Me.DDL_TYPE.Items.Add(New ListItem("學員資料查詢", "1"))
                    Me.DDL_TYPE.Items.Add(New ListItem("會員繳款資料查詢", "2"))
                    Me.DDL_TYPE.Items.Add(New ListItem("會員收入查詢", "4"))
                    Me.DDL_TYPE.Items.Add(New ListItem("一般會員繳款年報表", "5"))
                    Me.DDL_TYPE.Items.Add(New ListItem("種籽會員繳款報表", "6"))
                ElseIf Me.m_sQTYPE = "2" Then
                    Me.DDL_TYPE.Items.Clear()
                    Me.DDL_TYPE.Items.Add(New ListItem("課程報名查詢", "3"))
                End If

                Me.Bind_DDL_Type()
                Me.Bind_DDL_Place()
                Me.Bind_DDL_Date(ddl_StartYear, "1")
                Me.Bind_DDL_Date(ddl_EndYear, "1")
                Me.Bind_DDL_Date(ddl_StartMonth, "2")
                Me.Bind_DDL_Date(ddl_EndMonth, "2")

                '所屬區
                Me.Bind_DDL_MB_AREA()

                '繳款方式
                Me.Bind_MB_FEETYPE("0")

                '繳款會員種類
                Me.Bind_MB_FAMILY()
            End If
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

    Sub Bind_DDL_Type()
        Try
            Dim AP_CODEList As New AP_CODEList(Me.m_DBManager)
            AP_CODEList.loadByUpCode(91)

            ddl_Class_TYPE.DataSource = AP_CODEList.getCurrentDataSet.Tables(0)
            ddl_CLASS_TYPE.DataTextField = "TEXT"
            ddl_CLASS_TYPE.DataValueField = "VALUE"
            ddl_CLASS_TYPE.DataBind()

            ddl_CLASS_TYPE.Items.Insert(0, New ListItem("請選擇", ""))
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Sub Bind_DDL_Place()
        Dim DT_177 As DataTable = Nothing
        Try
            ddl_PLACE.Items.Clear()
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
        Catch ex As Exception
            Throw
        Finally
            If Not IsNothing(DT_177) Then
                DT_177.Dispose()
            End If
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
            Throw
        End Try
    End Sub

    Private Sub MBQry_TYPE_01_v01_Unload(sender As Object, e As EventArgs) Handles Me.Unload
        UIShareFun.releaseConnection(Me.m_DBManager)
    End Sub
#End Region

#Region "下拉選單"
    Sub Bind_DDL_MB_AREA()
        Try
            Me.DDL_MB_AREA.Items.Clear()

            Dim apCODEList As New AP_CODEList(Me.m_DBManager)
            apCODEList.loadByUpCode(m_sUpCode7)

            Me.DDL_MB_AREA.DataSource = apCODEList.getCurrentDataSet.Tables(0)
            Me.DDL_MB_AREA.DataTextField = "TEXT"
            Me.DDL_MB_AREA.DataValueField = "VALUE"
            Me.DDL_MB_AREA.DataBind()

            Me.DDL_MB_AREA.Items.Insert(0, New ListItem("請選擇", ""))
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Private Sub DDL_MB_AREA_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DDL_MB_AREA.SelectedIndexChanged
        Try
            If Utility.isValidateData(Me.DDL_MB_AREA.SelectedValue) Then
                Me.Bind_DDL_MB_LEADER(Me.DDL_MB_AREA.SelectedValue)
            Else
                Me.DDL_MB_LEADER.Items.Clear()
                Me.DDL_MB_LEADER.Items.Insert(0, New ListItem("所有委員", ""))
            End If
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

    Sub Bind_DDL_MB_LEADER(ByVal sGA_AREA As String)
        Try
            Dim apCODEList As New AP_CODEList(Me.m_DBManager)
            apCODEList.loadByUpCode(Me.m_sUpCode7)
            Dim ROW_Select() As DataRow = apCODEList.getCurrentDataSet.Tables(0).Select("VALUE='" & sGA_AREA & "'")
            If Not IsNothing(ROW_Select) AndAlso ROW_Select.Length > 0 Then
                apCODEList.clear()
                apCODEList.loadByUpCode(ROW_Select(0)("CODEID"))

                Me.DDL_MB_LEADER.Items.Clear()
                Me.DDL_MB_LEADER.DataTextField = "TEXT"
                Me.DDL_MB_LEADER.DataValueField = "VALUE"
                Me.DDL_MB_LEADER.DataSource = apCODEList.getCurrentDataSet.Tables(0)
                Me.DDL_MB_LEADER.DataBind()

                Me.DDL_MB_LEADER.Items.Insert(0, New ListItem("所有委員", ""))
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Sub Bind_MB_FEETYPE(ByVal sMode As String)
        Dim DT_23 As DataTable = Nothing
        Try
            MB_FEETYPE.Items.Clear()

            Dim apCODEList As New AP_CODEList(Me.m_DBManager)
            apCODEList.loadByUpCode(m_sUpcode23)
            DT_23 = apCODEList.getCurrentDataSet.Tables(0)
            If sMode = "1" Then
                ' 一般會員
                Dim ROW_EF() As DataRow = DT_23.Select("VALUE IN ('E','F')")
                If Not IsNothing(ROW_EF) AndAlso ROW_EF.Length > 0 Then
                    For Each ROW As DataRow In ROW_EF
                        DT_23.Rows.Remove(ROW)
                    Next
                End If
            ElseIf sMode = "2" Then
                '種子護法
                Dim ROW_ABCD() As DataRow = DT_23.Select("VALUE IN ('A','B','C','D')")
                If Not IsNothing(ROW_ABCD) AndAlso ROW_ABCD.Length > 0 Then
                    For Each ROW As DataRow In ROW_ABCD
                        DT_23.Rows.Remove(ROW)
                    Next
                End If
            End If

            Me.MB_FEETYPE.DataSource = DT_23
            Me.MB_FEETYPE.DataTextField = "TEXT"
            Me.MB_FEETYPE.DataValueField = "VALUE"
            Me.MB_FEETYPE.DataBind()

            Me.MB_FEETYPE.Items.Insert(0, New ListItem("請選擇", ""))
        Catch ex As Exception
            Throw
        Finally
            If Not IsNothing(DT_23) Then
                DT_23.Dispose()
            End If
        End Try
    End Sub

    Sub Bind_MB_FAMILY()
        Try
            Me.MB_MEMTYP.Items.Clear()

            Dim apCODEList As New AP_CODEList(Me.m_DBManager)
            apCODEList.loadByUpCode(m_sUpCode28)

            Me.MB_MEMTYP.DataSource = apCODEList.getCurrentDataSet.Tables(0)
            Me.MB_MEMTYP.DataTextField = "TEXT"
            Me.MB_MEMTYP.DataValueField = "VALUE"
            Me.MB_MEMTYP.DataBind()

            Me.MB_MEMTYP.Items.Insert(0, New ListItem("請選擇", ""))
        Catch ex As Exception
            Throw
        End Try
    End Sub
#End Region

#Region "Button"
    Private Sub btnQTYPE_Click(sender As Object, e As EventArgs) Handles btnQTYPE.Click
        Try
            If Not Utility.isValidateData(Me.DDL_TYPE.SelectedValue) Then
                com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "請選擇報表查詢格式")
                com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請選擇報表查詢格式")
                Return
            End If

            Select Case Me.DDL_TYPE.SelectedValue
                Case "1"
                    '學員資料查詢
                    Me.PLH_MB_AREA.Visible = True
                    Me.LTL_TYPE.Text = "學員資料查詢"
                    Me.LTL_PERIOD.Text = "輸入起訖日期"
                    Me.TR_BE_YYYMMDD.Visible = True
                    Me.PLH_BUTTON.Visible = True
                    Me.TR_MB_MEMTYP.Visible = False
                    Me.TR_BE_YYY.Visible = True
                    Me.btnMB_FWMK.Visible = False
                Case "2"
                    '會員繳款資料查詢
                    Me.PLH_MB_AREA.Visible = True
                    Me.LTL_TYPE.Text = "會員繳款資料查詢"
                    Me.LTL_PERIOD.Text = "輸入起訖日期"
                    Me.TR_BE_YYYMMDD.Visible = True
                    Me.PLH_BUTTON.Visible = True
                    Me.TR_MB_MEMTYP.Visible = False
                    Me.TR_BE_YYY.Visible = True
                    Me.btnMB_FWMK.Visible = False
                Case "3"
                    '課程報名查詢
                    'Me.PLH_MB_CLASS.Visible = True
                    '查詢課程清單
                    Me.btnQRYClass.Visible = True
                    Me.PLH_BUTTON.Visible = False
                    Me.LTL_TYPE.Text = "課程報名查詢"
                    Me.LTL_PERIOD.Text = "課程起日區間"
                    Me.TR_BE_YYYMMDD.Visible = False
                    Me.TR_MB_MEMTYP.Visible = False
                    Me.TR_BE_YYY.Visible = True
                    Me.TR_BE_YYY.Visible = False
                    Me.PLH_3.Visible = True
                    Me.btnMB_FWMK.Visible = True
                Case "4"
                    '會員收入查詢
                    Me.PLH_MB_AREA.Visible = False
                    Me.LTL_TYPE.Text = "會員收入查詢"
                    Me.LTL_PERIOD.Text = "繳款起訖日期"
                    Me.TR_BE_YYYMMDD.Visible = True
                    Me.PLH_BUTTON.Visible = True
                    Me.TR_MB_MEMTYP.Visible = True
                    Me.TR_BE_YYY.Visible = True
                    Me.btnMB_FWMK.Visible = False
                Case "5"
                    '一般會員繳款年報表
                    Me.LTL_TYPE.Text = "一般會員繳款年報表"
                    Me.TR_MB_MEMTYP.Visible = False
                    Me.PLH_MB_AREA.Visible = True
                    Me.PLH_MB_CLASS.Visible = False
                    Me.PLH_BUTTON.Visible = True
                    Me.TR_BE_YYYMMDD.Visible = False
                    Me.TR_BE_YYY.Visible = True
                    Me.btnMB_FWMK.Visible = False
                Case "6"
                    '種籽會員繳款報表
                    Me.LTL_TYPE.Text = "種籽會員繳款報表"
                    Me.TR_MB_MEMTYP.Visible = False
                    Me.PLH_MB_AREA.Visible = True
                    Me.PLH_MB_CLASS.Visible = False
                    Me.PLH_BUTTON.Visible = True
                    Me.TR_BE_YYYMMDD.Visible = False
                    Me.TR_BE_YYY.Visible = False
                    Me.btnMB_FWMK.Visible = False
            End Select

            Me.PLH_QTYPE.Visible = False
            Me.PLH_PARAS.Visible = True
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Try
            Me.PLH_QTYPE.Visible = True
            Me.PLH_PARAS.Visible = False
            Me.PLH_BUTTON.Visible = False
            Me.PLH_MB_AREA.Visible = False
            Me.PLH_MB_CLASS.Visible = False

            Me.ClearPageValue()
            Me.RP_1.DataSource = Nothing
            Me.RP_1.DataBind()
            Me.RP_2.DataSource = Nothing
            Me.RP_2.DataBind()
            Me.RP_3.DataSource = Nothing
            Me.RP_3.DataBind()
            Me.RP_4.DataSource = Nothing
            Me.RP_4.DataBind()
            Me.RP_5.DataSource = Nothing
            Me.RP_5.DataBind()
            Me.RP_6.DataSource = Nothing
            Me.RP_6.DataBind()

            Me.PLH_DATA.Visible = False
            Me.PLH_DATA_1.Visible = False
            Me.PLH_DATA_2.Visible = False
            Me.PLH_DATA_3.Visible = False
            Me.PLH_DATA_4.Visible = False
            Me.PLH_DATA_5.Visible = False
            Me.PLH_DATA_6.Visible = False

            '查詢課程清單
            Me.btnQRYClass.Visible = False
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

    Private Sub btnQRYClass_Click(sender As Object, e As EventArgs) Handles btnQRYClass.Click
        Dim DT_MB_CLASS As DataTable = Nothing
        Dim DT_UK_MB_CLASS As DataTable = Nothing
        Try
            'Dim sBDateErr As String = String.Empty
            'sBDateErr = Me.isValidDate("起日", Me.TXT_BY.Text, Me.TXT_BM.Text, Me.TXT_BD.Text)
            'If Utility.isValidateData(sBDateErr) Then
            '    com.Azion.EloanUtility.UIUtility.showJSMsg(Me, sBDateErr)
            '    com.Azion.EloanUtility.UIUtility.showErrMsg(Me, sBDateErr)
            '    Return
            'End If

            'Dim sEDateErr As String = String.Empty
            'sEDateErr = Me.isValidDate("訖日", Me.TXT_EY.Text, Me.TXT_EM.Text, Me.TXT_ED.Text)
            'If Utility.isValidateData(sEDateErr) Then
            '    com.Azion.EloanUtility.UIUtility.showJSMsg(Me, sEDateErr)
            '    com.Azion.EloanUtility.UIUtility.showErrMsg(Me, sEDateErr)
            '    Return
            'End If

            'Dim D_BDay As Object = Convert.DBNull
            'D_BDay = Me.ConvertWDate(Me.TXT_BY.Text, Me.TXT_BM.Text, Me.TXT_BD.Text)
            'Dim D_EDay As Object = Convert.DBNull
            'D_EDay = Me.ConvertWDate(Me.TXT_EY.Text, Me.TXT_EM.Text, Me.TXT_ED.Text)

            'If Not IsDate(D_BDay) Then
            '    com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "請輸入課程起日")
            '    com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請輸入課程起日")
            '    Return
            'End If

            'If Not IsDate(D_EDay) Then
            '    com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "請輸入課程訖日")
            '    com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請輸入課程訖日")
            '    Return
            'End If

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

            Dim mbCLASSList As New MB_CLASSList(Me.m_DBManager)
            If sFlag.Trim.Length = 0 AndAlso mbCLASSList.getClassList(ddl_CLASS_TYPE.SelectedValue, IIf(ddl_PLACE.SelectedValue <> "", ddl_PLACE.SelectedItem.Text, ""), sSDATE, sEDATE) > 0 Then
                DT_MB_CLASS = mbCLASSList.getCurrentDataSet.Tables(0)
            End If
            'mbCLASSList.LoadBySEDate(D_BDay, D_EDay)
            'DT_MB_CLASS = mbCLASSList.getCurrentDataSet.Tables(0)

            If Not IsNothing(DT_MB_CLASS) AndAlso DT_MB_CLASS.Rows.Count > 0 Then
                Me.DDL_MB_CLASS.Items.Clear()
                'Me.DDL_MB_CLASS.DataTextField = "MB_SHOW"
                'Me.DDL_MB_CLASS.DataValueField = "MB_SEQ"
                'Me.DDL_MB_CLASS.DataSource = DT_MB_CLASS

                DT_UK_MB_CLASS = New DataView(DT_MB_CLASS, String.Empty, "MB_SEQ,MB_BATCH", DataViewRowState.CurrentRows).ToTable(True, "MB_SEQ")

                For Each ROW As DataRow In DT_UK_MB_CLASS.Rows
                    Dim sCLASS As String = String.Empty

                    Dim ROW_MB_SEQ() As DataRow = Nothing
                    ROW_MB_SEQ = DT_MB_CLASS.Select("MB_SEQ=" & ROW("MB_SEQ"))
                    If Not IsNothing(ROW_MB_SEQ) AndAlso ROW_MB_SEQ.Length > 0 Then
                        sCLASS = "【" & ROW_MB_SEQ(0)("MB_SEQ").ToString & "】" & ROW_MB_SEQ(0)("MB_CLASS_NAME").ToString
                        Dim objListItem As New ListItem(sCLASS, ROW_MB_SEQ(0)("MB_SEQ").ToString)
                        Me.DDL_MB_CLASS.Items.Add(objListItem)
                    End If
                Next

                Me.DDL_MB_CLASS.DataBind()

                Me.DDL_MB_CLASS.Items.Insert(0, New ListItem("請選擇", ""))

                Me.PLH_MB_CLASS.Visible = True
                Me.PLH_BUTTON.Visible = True
            Else
                com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "課程起日區間內查無課程資料")
                com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "課程起日區間內查無課程資料")
                Return
            End If
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        Finally
            If Not IsNothing(DT_MB_CLASS) Then
                DT_MB_CLASS.Dispose()
            End If
        End Try
    End Sub

    Private Sub btnQRY_Click(sender As Object, e As EventArgs) Handles btnQRY.Click
        Try
            If Me.DDL_TYPE.SelectedValue = "5" Then
                If IsNumeric(Me.TXT_B_YYY.Text) AndAlso IsNumeric(Me.TXT_E_YYY.Text) Then
                    If CDec(Me.TXT_B_YYY.Text) > CDec(Me.TXT_E_YYY.Text) Then
                        com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "起日年度不可大於訖日年度")
                        com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "起日年度不可大於訖日年度")
                        Return
                    End If
                ElseIf IsNumeric(Me.TXT_B_YYY.Text) AndAlso Not IsNumeric(Me.TXT_E_YYY.Text) Then
                    com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "請輸入訖日年度")
                    com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請輸入訖日年度")
                    Return
                ElseIf Not IsNumeric(Me.TXT_B_YYY.Text) AndAlso IsNumeric(Me.TXT_E_YYY.Text) Then
                    com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "請輸入起日年度")
                    com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請輸入起日年度")
                    Return
                End If
            Else
                Dim sBDateErr As String = String.Empty
                sBDateErr = Me.isValidDate("起日", Me.TXT_BY.Text, Me.TXT_BM.Text, Me.TXT_BD.Text)
                If Utility.isValidateData(sBDateErr) Then
                    com.Azion.EloanUtility.UIUtility.showJSMsg(Me, sBDateErr)
                    com.Azion.EloanUtility.UIUtility.showErrMsg(Me, sBDateErr)
                    Return
                End If

                Dim sEDateErr As String = String.Empty
                sEDateErr = Me.isValidDate("訖日", Me.TXT_EY.Text, Me.TXT_EM.Text, Me.TXT_ED.Text)
                If Utility.isValidateData(sEDateErr) Then
                    com.Azion.EloanUtility.UIUtility.showJSMsg(Me, sEDateErr)
                    com.Azion.EloanUtility.UIUtility.showErrMsg(Me, sEDateErr)
                    Return
                End If

                Dim D_BDay As Object = Convert.DBNull
                D_BDay = Me.ConvertWDate(Me.TXT_BY.Text, Me.TXT_BM.Text, Me.TXT_BD.Text)
                Dim D_EDay As Object = Convert.DBNull
                D_EDay = Me.ConvertWDate(Me.TXT_EY.Text, Me.TXT_EM.Text, Me.TXT_ED.Text)

                If IsDate(D_BDay) AndAlso IsDate(D_EDay) Then
                    If D_BDay > D_EDay Then
                        com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "起日不可大於訖日")
                        com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "起日不可大於訖日")
                        Return
                    End If
                ElseIf IsDate(D_BDay) AndAlso Not IsDate(D_EDay) Then
                    com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "請輸入訖日")
                    com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請輸入訖日")
                    Return
                ElseIf Not IsDate(D_BDay) AndAlso IsDate(D_EDay) Then
                    com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "請輸入起日")
                    com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請輸入起日")
                    Return
                End If
            End If

            Select Case Me.DDL_TYPE.SelectedValue
                Case "1"
                    '學員資料查詢
                    Me.Bind_TYPE_1(1)
                Case "2"
                    '會員繳款資料查詢
                    Me.Bind_TYPE_2(1)
                Case "3"
                    '課程報名查詢
                    If Not IsNumeric(Me.DDL_MB_CLASS.SelectedValue) Then
                        com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "請選擇課程編號")
                        com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請選擇課程編號")
                        Return
                    End If

                    Me.Bind_TYPE_3(1)
                Case "4"
                    '會員收入查詢
                    Me.Bind_TYPE_4(1)
                Case "5"
                    '一般會員繳款年報表
                    Me.Bind_TYPE_5(1)
                Case "6"
                    '種籽會員繳款報表
                    Me.Bind_TYPE_6(1)
            End Select

            Me.PLH_DATA.Visible = True
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

    Private Sub btnEXCEL_Click(sender As Object, e As EventArgs) Handles btnEXCEL.Click
        Try
            If Me.DDL_TYPE.SelectedValue = "5" Then
                If IsNumeric(Me.TXT_B_YYY.Text) AndAlso IsNumeric(Me.TXT_E_YYY.Text) Then
                    If CDec(Me.TXT_B_YYY.Text) > CDec(Me.TXT_E_YYY.Text) Then
                        com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "起日年度不可大於訖日年度")
                        com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "起日年度不可大於訖日年度")
                        Return
                    End If
                ElseIf IsNumeric(Me.TXT_B_YYY.Text) AndAlso Not IsNumeric(Me.TXT_E_YYY.Text) Then
                    com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "請輸入訖日年度")
                    com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請輸入訖日年度")
                    Return
                ElseIf Not IsNumeric(Me.TXT_B_YYY.Text) AndAlso IsNumeric(Me.TXT_E_YYY.Text) Then
                    com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "請輸入起日年度")
                    com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請輸入起日年度")
                    Return
                End If
            ElseIf Me.DDL_TYPE.SelectedValue = "3" Then
                If Not IsNumeric(Me.DDL_MB_CLASS.SelectedValue) Then
                    com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "請選擇課程編號")
                    com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請選擇課程編號")
                    Return
                End If
            Else
                Dim sBDateErr As String = String.Empty
                sBDateErr = Me.isValidDate("起日", Me.TXT_BY.Text, Me.TXT_BM.Text, Me.TXT_BD.Text)
                If Utility.isValidateData(sBDateErr) Then
                    com.Azion.EloanUtility.UIUtility.showJSMsg(Me, sBDateErr)
                    com.Azion.EloanUtility.UIUtility.showErrMsg(Me, sBDateErr)
                    Return
                End If

                Dim sEDateErr As String = String.Empty
                sEDateErr = Me.isValidDate("訖日", Me.TXT_EY.Text, Me.TXT_EM.Text, Me.TXT_ED.Text)
                If Utility.isValidateData(sEDateErr) Then
                    com.Azion.EloanUtility.UIUtility.showJSMsg(Me, sEDateErr)
                    com.Azion.EloanUtility.UIUtility.showErrMsg(Me, sEDateErr)
                    Return
                End If

                Dim D_BDay As Object = Convert.DBNull
                D_BDay = Me.ConvertWDate(Me.TXT_BY.Text, Me.TXT_BM.Text, Me.TXT_BD.Text)
                Dim D_EDay As Object = Convert.DBNull
                D_EDay = Me.ConvertWDate(Me.TXT_EY.Text, Me.TXT_EM.Text, Me.TXT_ED.Text)

                If IsDate(D_BDay) AndAlso IsDate(D_EDay) Then
                    If D_BDay > D_EDay Then
                        com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "起日不可大於訖日")
                        com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "起日不可大於訖日")
                        Return
                    End If
                ElseIf IsDate(D_BDay) AndAlso Not IsDate(D_EDay) Then
                    com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "請輸入訖日")
                    com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請輸入訖日")
                    Return
                ElseIf Not IsDate(D_BDay) AndAlso IsDate(D_EDay) Then
                    com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "請輸入起日")
                    com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請輸入起日")
                    Return
                End If

                If Me.DDL_TYPE.SelectedValue = "3" Then
                    If Not IsDate(D_BDay) Then
                        com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "請輸入課程起日")
                        com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請輸入課程起日")
                        Return
                    End If

                    If Not IsDate(D_EDay) Then
                        com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "請輸入課程訖日")
                        com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請輸入課程訖日")
                        Return
                    End If
                End If
            End If

            Dim sURL As String = String.Empty

            Select Case Me.DDL_TYPE.SelectedValue
                Case "1"
                    '學員資料查詢
                    sURL = com.Azion.EloanUtility.UIUtility.getRootPath & "/QRY/MBQry_TYPE_01_01_v01.aspx?" & _
                          "BY=" & Me.TXT_BY.Text & "&BM=" & Me.TXT_BM.Text & "&BD=" & Me.TXT_BD.Text & _
                          "&EY=" & Me.TXT_EY.Text & "&EM=" & Me.TXT_EM.Text & "&ED=" & Me.TXT_ED.Text & _
                          "&MB_AREA=" & Me.DDL_MB_AREA.SelectedValue & "&MB_LEADER=" & Me.DDL_MB_LEADER.SelectedValue
                Case "2"
                    '會員繳款資料查詢
                    sURL = com.Azion.EloanUtility.UIUtility.getRootPath & "/QRY/MBQry_TYPE_01_02_v01.aspx?" & _
                          "BY=" & Me.TXT_BY.Text & "&BM=" & Me.TXT_BM.Text & "&BD=" & Me.TXT_BD.Text & _
                          "&EY=" & Me.TXT_EY.Text & "&EM=" & Me.TXT_EM.Text & "&ED=" & Me.TXT_ED.Text & _
                          "&MB_AREA=" & Me.DDL_MB_AREA.SelectedValue & "&MB_LEADER=" & Me.DDL_MB_LEADER.SelectedValue
                Case "3"
                    '課程報名查詢
                    sURL = com.Azion.EloanUtility.UIUtility.getRootPath & "/QRY/MBQry_TYPE_01_03_v01.aspx?MB_SEQ=" & Me.DDL_MB_CLASS.SelectedValue
                Case "4"
                    '會員收入查詢
                    sURL = com.Azion.EloanUtility.UIUtility.getRootPath & "/QRY/MBQry_TYPE_01_04_v01.aspx?" & _
                           "BY=" & Me.TXT_BY.Text & "&BM=" & Me.TXT_BM.Text & "&BD=" & Me.TXT_BD.Text & _
                           "&EY=" & Me.TXT_EY.Text & "&EM=" & Me.TXT_EM.Text & "&ED=" & Me.TXT_ED.Text & _
                           "&MB_MEMTYP=" & Me.MB_MEMTYP.SelectedValue & "&MB_FEETYPE=" & Me.MB_FEETYPE.SelectedValue
                Case "5"
                    '一般會員繳款年報表
                    sURL = com.Azion.EloanUtility.UIUtility.getRootPath & "/QRY/MBQry_TYPE_01_05_v01.aspx?" & _
                           "BY=" & Me.TXT_B_YYY.Text & "&EY=" & Me.TXT_E_YYY.Text & _
                           "&MB_AREA=" & Me.DDL_MB_AREA.SelectedValue & "&MB_LEADER=" & Me.DDL_MB_LEADER.SelectedValue
                Case "6"
                    sURL = com.Azion.EloanUtility.UIUtility.getRootPath & "/QRY/MBQry_TYPE_01_06_v01.aspx?" & _
                           "MB_AREA=" & Me.DDL_MB_AREA.SelectedValue & "&MB_LEADER=" & Me.DDL_MB_LEADER.SelectedValue
            End Select

            com.Azion.EloanUtility.UIUtility.showOpen(sURL)
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

    Sub Bind_TYPE_1(ByVal iPageNUM As Integer)
        Dim DT_MEMBER As DataTable = Nothing
        Try
            Dim D_BDay As Object = Convert.DBNull
            D_BDay = Me.ConvertWDate(Me.TXT_BY.Text, Me.TXT_BM.Text, Me.TXT_BD.Text)
            Dim D_EDay As Object = Convert.DBNull
            D_EDay = Me.ConvertWDate(Me.TXT_EY.Text, Me.TXT_EM.Text, Me.TXT_ED.Text)

            Dim isPeriod As Boolean = False
            If IsDate(D_BDay) AndAlso IsDate(D_EDay) Then
                isPeriod = True
            End If

            Dim iStart As Decimal = 0
            iStart = (iPageNUM - 1) * Me.m_PageSize

            Dim mbMEMBERList As New MB_MEMBERList(Me.m_DBManager)
            If Not isPeriod AndAlso Not Utility.isValidateData(Me.DDL_MB_AREA.SelectedValue) AndAlso Not Utility.isValidateData(Me.DDL_MB_LEADER.SelectedValue) Then
                mbMEMBERList.QRY_1(iStart, Me.m_PageSize)
            ElseIf isPeriod AndAlso Not Utility.isValidateData(Me.DDL_MB_AREA.SelectedValue) AndAlso Not Utility.isValidateData(Me.DDL_MB_LEADER.SelectedValue) Then
                mbMEMBERList.QRY_2(iStart, Me.m_PageSize, D_BDay, D_EDay)
            ElseIf isPeriod AndAlso Utility.isValidateData(Me.DDL_MB_AREA.SelectedValue) AndAlso Not Utility.isValidateData(Me.DDL_MB_LEADER.SelectedValue) Then
                mbMEMBERList.QRY_3(iStart, Me.m_PageSize, D_BDay, D_EDay, Me.DDL_MB_AREA.SelectedValue)
            ElseIf isPeriod AndAlso Not Utility.isValidateData(Me.DDL_MB_AREA.SelectedValue) AndAlso Utility.isValidateData(Me.DDL_MB_LEADER.SelectedValue) Then
                mbMEMBERList.QRY_4(iStart, Me.m_PageSize, D_BDay, D_EDay, Me.DDL_MB_LEADER.SelectedValue)
            ElseIf isPeriod AndAlso Utility.isValidateData(Me.DDL_MB_AREA.SelectedValue) AndAlso Utility.isValidateData(Me.DDL_MB_LEADER.SelectedValue) Then
                mbMEMBERList.QRY_5(iStart, Me.m_PageSize, D_BDay, D_EDay, Me.DDL_MB_AREA.SelectedValue, Me.DDL_MB_LEADER.SelectedValue)
            ElseIf Not isPeriod AndAlso Utility.isValidateData(Me.DDL_MB_AREA.SelectedValue) AndAlso Not Utility.isValidateData(Me.DDL_MB_LEADER.SelectedValue) Then
                mbMEMBERList.QRY_6(iStart, Me.m_PageSize, Me.DDL_MB_AREA.SelectedValue)
            ElseIf Not isPeriod AndAlso Utility.isValidateData(Me.DDL_MB_AREA.SelectedValue) AndAlso Utility.isValidateData(Me.DDL_MB_LEADER.SelectedValue) Then
                mbMEMBERList.QRY_7(iStart, Me.m_PageSize, Me.DDL_MB_AREA.SelectedValue, Me.DDL_MB_LEADER.SelectedValue)
            End If
            DT_MEMBER = mbMEMBERList.getCurrentDataSet.Tables(0)
            Me.RP_1.DataSource = DT_MEMBER
            Me.RP_1.DataBind()

            Me.initCHGPage_1(iPageNUM)

            Me.PLH_DATA_1.Visible = True
        Catch ex As Exception
            Throw
        Finally
            If Not IsNothing(DT_MEMBER) Then
                DT_MEMBER.Dispose()
            End If
        End Try
    End Sub

    Sub Bind_TYPE_2(ByVal iPageNUM As Integer)
        Dim DT_MB_MEMREV As DataTable = Nothing
        Try
            Dim D_BDay As Object = Convert.DBNull
            D_BDay = Me.ConvertWDate(Me.TXT_BY.Text, Me.TXT_BM.Text, Me.TXT_BD.Text)
            Dim D_EDay As Object = Convert.DBNull
            D_EDay = Me.ConvertWDate(Me.TXT_EY.Text, Me.TXT_EM.Text, Me.TXT_ED.Text)

            Dim isPeriod As Boolean = False
            If IsDate(D_BDay) AndAlso IsDate(D_EDay) Then
                isPeriod = True
            End If

            Dim iStart As Decimal = 0
            iStart = (iPageNUM - 1) * Me.m_PageSize

            Dim mbMEMREVList As New MB_MEMREVList(Me.m_DBManager)
            If Not isPeriod AndAlso Not Utility.isValidateData(Me.DDL_MB_AREA.SelectedValue) AndAlso Not Utility.isValidateData(Me.DDL_MB_LEADER.SelectedValue) Then
                mbMEMREVList.QRY_1(iStart, Me.m_PageSize)
            ElseIf isPeriod AndAlso Not Utility.isValidateData(Me.DDL_MB_AREA.SelectedValue) AndAlso Not Utility.isValidateData(Me.DDL_MB_LEADER.SelectedValue) Then
                mbMEMREVList.QRY_2(iStart, Me.m_PageSize, D_BDay, D_EDay)
            ElseIf isPeriod AndAlso Utility.isValidateData(Me.DDL_MB_AREA.SelectedValue) AndAlso Not Utility.isValidateData(Me.DDL_MB_LEADER.SelectedValue) Then
                mbMEMREVList.QRY_3(iStart, Me.m_PageSize, D_BDay, D_EDay, Me.DDL_MB_AREA.SelectedValue)
            ElseIf isPeriod AndAlso Not Utility.isValidateData(Me.DDL_MB_AREA.SelectedValue) AndAlso Utility.isValidateData(Me.DDL_MB_LEADER.SelectedValue) Then
                mbMEMREVList.QRY_4(iStart, Me.m_PageSize, D_BDay, D_EDay, Me.DDL_MB_LEADER.SelectedValue)
            ElseIf isPeriod AndAlso Utility.isValidateData(Me.DDL_MB_AREA.SelectedValue) AndAlso Utility.isValidateData(Me.DDL_MB_LEADER.SelectedValue) Then
                mbMEMREVList.QRY_5(iStart, Me.m_PageSize, D_BDay, D_EDay, Me.DDL_MB_AREA.SelectedValue, Me.DDL_MB_LEADER.SelectedValue)
            ElseIf Not isPeriod AndAlso Utility.isValidateData(Me.DDL_MB_AREA.SelectedValue) AndAlso Not Utility.isValidateData(Me.DDL_MB_LEADER.SelectedValue) Then
                mbMEMREVList.QRY_6(iStart, Me.m_PageSize, Me.DDL_MB_AREA.SelectedValue)
            ElseIf Not isPeriod AndAlso Utility.isValidateData(Me.DDL_MB_AREA.SelectedValue) AndAlso Utility.isValidateData(Me.DDL_MB_LEADER.SelectedValue) Then
                mbMEMREVList.QRY_7(iStart, Me.m_PageSize, Me.DDL_MB_AREA.SelectedValue, Me.DDL_MB_LEADER.SelectedValue)
            End If
            DT_MB_MEMREV = mbMEMREVList.getCurrentDataSet.Tables(0)
            Me.RP_2.DataSource = DT_MB_MEMREV
            Me.RP_2.DataBind()

            Me.initCHGPage_2(iPageNUM)

            Me.PLH_DATA_2.Visible = True
        Catch ex As Exception
            Throw
        Finally
            If Not IsNothing(DT_MB_MEMREV) Then
                DT_MB_MEMREV.Dispose()
            End If
        End Try
    End Sub

    Sub Bind_TYPE_3(ByVal iPageNUM As Integer)
        Dim DT_MB_MEMCLASS As DataTable = Nothing
        Try
            Dim iMB_SEQ As Decimal = 0
            iMB_SEQ = CDec(Me.DDL_MB_CLASS.SelectedValue)

            Dim iStart As Decimal = 0
            iStart = (iPageNUM - 1) * Me.m_PageSize

            Dim mbMEMCLASSList As New MB_MEMCLASSList(Me.m_DBManager)
            mbMEMCLASSList.LoadQData(iStart, Me.m_PageSize, iMB_SEQ)
            DT_MB_MEMCLASS = mbMEMCLASSList.getCurrentDataSet.Tables(0)

            Me.TD_MB_APV.Visible = False
            Dim MB_CLASSList As New MB_CLASSList(Me.m_DBManager)
            MB_CLASSList.LoadByMB_SEQ(iMB_SEQ)
            If MB_CLASSList.getCurrentDataSet.Tables(0).Rows.Count > 0 Then
                Dim ROW As DataRow = MB_CLASSList.getCurrentDataSet.Tables(0).Rows(0)
                If ROW("MB_APV").ToString = "1" OrElse ROW("MB_APV").ToString = "3" Then
                    Me.TD_MB_APV.Visible = True
                End If
            End If

            Me.RP_3.DataSource = DT_MB_MEMCLASS
            Me.RP_3.DataBind()

            Me.initCHGPage_3(iPageNUM)

            Me.PLH_DATA_3.Visible = True
        Catch ex As Exception
            Throw
        Finally
            If Not IsNothing(DT_MB_MEMCLASS) Then
                DT_MB_MEMCLASS.Dispose()
            End If
        End Try
    End Sub

    Sub Bind_TYPE_4(ByVal iPageNUM As Integer)
        Dim DT_MB_MEMREV As DataTable = Nothing
        Try
            Dim D_BDay As Object = Convert.DBNull
            D_BDay = Me.ConvertWDate(Me.TXT_BY.Text, Me.TXT_BM.Text, Me.TXT_BD.Text)
            Dim D_EDay As Object = Convert.DBNull
            D_EDay = Me.ConvertWDate(Me.TXT_EY.Text, Me.TXT_EM.Text, Me.TXT_ED.Text)

            Dim isPeriod As Boolean = False
            Dim iB_YYYYMMDD As Decimal = 0
            Dim iE_YYYYMMDD As Decimal = 0
            If IsDate(D_BDay) AndAlso IsDate(D_EDay) Then
                isPeriod = True

                iB_YYYYMMDD = CDate(D_BDay).Year * 10000 + CDate(D_BDay).Month * 100 + CDate(D_BDay).Day
                iE_YYYYMMDD = CDate(D_EDay).Year * 10000 + CDate(D_EDay).Month * 100 + CDate(D_EDay).Day
            End If

            Dim iStart As Decimal = 0
            iStart = (iPageNUM - 1) * Me.m_PageSize

            Dim mbMEMREVList As New MB_MEMREVList(Me.m_DBManager)
            If isPeriod AndAlso Not Utility.isValidateData(Me.MB_MEMTYP.SelectedValue) AndAlso Not Utility.isValidateData(Me.MB_FEETYPE.SelectedValue) Then
                mbMEMREVList.QRY_INCOME_1(iStart, Me.m_PageSize, iB_YYYYMMDD, iE_YYYYMMDD)
            ElseIf isPeriod AndAlso Utility.isValidateData(Me.MB_MEMTYP.SelectedValue) AndAlso Utility.isValidateData(Me.MB_FEETYPE.SelectedValue) Then
                mbMEMREVList.QRY_INCOME_2(iStart, Me.m_PageSize, iB_YYYYMMDD, iE_YYYYMMDD, Me.MB_MEMTYP.SelectedValue, Me.MB_FEETYPE.SelectedValue)
            ElseIf isPeriod AndAlso Not Utility.isValidateData(Me.MB_MEMTYP.SelectedValue) AndAlso Utility.isValidateData(Me.MB_FEETYPE.SelectedValue) Then
                mbMEMREVList.QRY_INCOME_3(iStart, Me.m_PageSize, iB_YYYYMMDD, iE_YYYYMMDD, Me.MB_FEETYPE.SelectedValue)
            ElseIf Not isPeriod AndAlso Utility.isValidateData(Me.MB_MEMTYP.SelectedValue) AndAlso Utility.isValidateData(Me.MB_FEETYPE.SelectedValue) Then
                mbMEMREVList.QRY_INCOME_4(iStart, Me.m_PageSize, Me.MB_MEMTYP.SelectedValue, Me.MB_FEETYPE.SelectedValue)
            ElseIf Not isPeriod AndAlso Utility.isValidateData(Me.MB_MEMTYP.SelectedValue) AndAlso Not Utility.isValidateData(Me.MB_FEETYPE.SelectedValue) Then
                mbMEMREVList.QRY_INCOME_5(iStart, Me.m_PageSize, Me.MB_MEMTYP.SelectedValue)
            ElseIf Not isPeriod AndAlso Not Utility.isValidateData(Me.MB_MEMTYP.SelectedValue) AndAlso Utility.isValidateData(Me.MB_FEETYPE.SelectedValue) Then
                mbMEMREVList.QRY_INCOME_6(iStart, Me.m_PageSize, Me.MB_FEETYPE.SelectedValue)
            ElseIf Not isPeriod AndAlso Not Utility.isValidateData(Me.MB_MEMTYP.SelectedValue) AndAlso Not Utility.isValidateData(Me.MB_MEMTYP.SelectedValue) Then
                mbMEMREVList.QRY_INCOME_7(iStart, Me.m_PageSize)
            End If

            DT_MB_MEMREV = mbMEMREVList.getCurrentDataSet.Tables(0)
            Me.RP_4.DataSource = DT_MB_MEMREV
            Me.RP_4.DataBind()

            Me.initCHGPage_4(iPageNUM)

            Me.PLH_DATA_4.Visible = True
        Catch ex As Exception
            Throw
        Finally
            If Not IsNothing(DT_MB_MEMREV) Then
                DT_MB_MEMREV.Dispose()
            End If
        End Try
    End Sub

    Sub Bind_TYPE_5(ByVal iPageNUM As Integer)
        Dim DT_MB_MEMFEE As DataTable = Nothing
        Try
            Dim isPeriod As Boolean = False
            If IsNumeric(Me.TXT_B_YYY.Text) AndAlso IsNumeric(Me.TXT_E_YYY.Text) Then
                isPeriod = True
            End If

            Dim iStart As Decimal = 0
            iStart = (iPageNUM - 1) * Me.m_PageSize
            Dim mbMEMFEEList As New MB_MEMFEEList(Me.m_DBManager)
            If Not isPeriod AndAlso Not Utility.isValidateData(Me.DDL_MB_AREA.SelectedValue) AndAlso Not Utility.isValidateData(Me.DDL_MB_LEADER.SelectedValue) Then
                mbMEMFEEList.QRY_1(iStart, Me.m_PageSize)
            ElseIf isPeriod AndAlso Not Utility.isValidateData(Me.DDL_MB_AREA.SelectedValue) AndAlso Not Utility.isValidateData(Me.DDL_MB_LEADER.SelectedValue) Then
                mbMEMFEEList.QRY_2(CDec(Me.TXT_B_YYY.Text), CDec(Me.TXT_E_YYY.Text), iStart, Me.m_PageSize)
            ElseIf isPeriod AndAlso Utility.isValidateData(Me.DDL_MB_AREA.SelectedValue) AndAlso Not Utility.isValidateData(Me.DDL_MB_LEADER.SelectedValue) Then
                mbMEMFEEList.QRY_3(CDec(Me.TXT_B_YYY.Text), CDec(Me.TXT_E_YYY.Text), Me.DDL_MB_AREA.SelectedValue, iStart, Me.m_PageSize)
            ElseIf isPeriod AndAlso Utility.isValidateData(Me.DDL_MB_AREA.SelectedValue) AndAlso Utility.isValidateData(Me.DDL_MB_LEADER.SelectedValue) Then
                mbMEMFEEList.QRY_4(CDec(Me.TXT_B_YYY.Text), CDec(Me.TXT_E_YYY.Text), Me.DDL_MB_AREA.SelectedValue, Me.DDL_MB_LEADER.SelectedValue, iStart, Me.m_PageSize)
            ElseIf Not isPeriod AndAlso Utility.isValidateData(Me.DDL_MB_AREA.SelectedValue) AndAlso Utility.isValidateData(Me.DDL_MB_LEADER.SelectedValue) Then
                mbMEMFEEList.QRY_5(Me.DDL_MB_AREA.SelectedValue, Me.DDL_MB_LEADER.SelectedValue, iStart, Me.m_PageSize)
            ElseIf Not isPeriod AndAlso Utility.isValidateData(Me.DDL_MB_AREA.SelectedValue) AndAlso Not Utility.isValidateData(Me.DDL_MB_LEADER.SelectedValue) Then
                mbMEMFEEList.QRY_6(Me.DDL_MB_AREA.SelectedValue, iStart, Me.m_PageSize)
            ElseIf Not isPeriod AndAlso Not Utility.isValidateData(Me.DDL_MB_AREA.SelectedValue) AndAlso Utility.isValidateData(Me.DDL_MB_LEADER.SelectedValue) Then
                mbMEMFEEList.QRY_7(Me.DDL_MB_LEADER.SelectedValue, iStart, Me.m_PageSize)
            End If

            DT_MB_MEMFEE = mbMEMFEEList.getCurrentDataSet.Tables(0)
            Me.RP_5.DataSource = DT_MB_MEMFEE
            Me.RP_5.DataBind()

            Me.initCHGPage_5(iPageNUM)

            Me.PLH_DATA_5.Visible = True
        Catch ex As Exception
            Throw
        Finally
            If Not IsNothing(DT_MB_MEMFEE) Then
                DT_MB_MEMFEE.Dispose()
            End If
        End Try
    End Sub

    Sub Bind_TYPE_6(ByVal iPageNUM As Integer)
        Dim DT_MB_MEMFEE As DataTable = Nothing
        Try
            Dim iStart As Decimal = 0
            iStart = (iPageNUM - 1) * Me.m_PageSize

            Dim mbVIPList As New MB_VIPList(Me.m_DBManager)
            If Not Utility.isValidateData(Me.DDL_MB_AREA.SelectedValue) AndAlso Not Utility.isValidateData(Me.DDL_MB_LEADER.SelectedValue) Then
                mbVIPList.QRY_1(iStart, Me.m_PageSize)
            ElseIf Utility.isValidateData(Me.DDL_MB_AREA.SelectedValue) AndAlso Not Utility.isValidateData(Me.DDL_MB_LEADER.SelectedValue) Then
                mbVIPList.QRY_2(Me.DDL_MB_AREA.SelectedValue, iStart, Me.m_PageSize)
            ElseIf Utility.isValidateData(Me.DDL_MB_AREA.SelectedValue) AndAlso Utility.isValidateData(Me.DDL_MB_LEADER.SelectedValue) Then
                mbVIPList.QRY_3(Me.DDL_MB_AREA.SelectedValue, Me.DDL_MB_LEADER.SelectedValue, iStart, Me.m_PageSize)
            ElseIf Not Utility.isValidateData(Me.DDL_MB_AREA.SelectedValue) AndAlso Utility.isValidateData(Me.DDL_MB_LEADER.SelectedValue) Then
                mbVIPList.QRY_4(Me.DDL_MB_LEADER.SelectedValue, iStart, Me.m_PageSize)
            End If

            DT_MB_MEMFEE = mbVIPList.getCurrentDataSet.Tables(0)
            Me.RP_6.DataSource = DT_MB_MEMFEE
            Me.RP_6.DataBind()

            Me.initCHGPage_6(iPageNUM)

            Me.PLH_DATA_6.Visible = True
        Catch ex As Exception
            Throw
        Finally
            If Not IsNothing(DT_MB_MEMFEE) Then
                DT_MB_MEMFEE.Dispose()
            End If
        End Try
    End Sub
#End Region

#Region "換頁"
    Sub initCHGPage_1(ByVal iPageNUM As Integer)
        Try
            Dim iRowCount As Decimal = 0

            Dim D_BDay As Object = Convert.DBNull
            D_BDay = Me.ConvertWDate(Me.TXT_BY.Text, Me.TXT_BM.Text, Me.TXT_BD.Text)
            Dim D_EDay As Object = Convert.DBNull
            D_EDay = Me.ConvertWDate(Me.TXT_EY.Text, Me.TXT_EM.Text, Me.TXT_ED.Text)

            Dim isPeriod As Boolean = False
            If IsDate(D_BDay) AndAlso IsDate(D_EDay) Then
                isPeriod = True
            End If

            Dim mbMEMBERList As New MB_MEMBERList(Me.m_DBManager)
            If Not isPeriod AndAlso Not Utility.isValidateData(Me.DDL_MB_AREA.SelectedValue) AndAlso Not Utility.isValidateData(Me.DDL_MB_LEADER.SelectedValue) Then
                iRowCount = mbMEMBERList.COUNT_1()
            ElseIf isPeriod AndAlso Not Utility.isValidateData(Me.DDL_MB_AREA.SelectedValue) AndAlso Not Utility.isValidateData(Me.DDL_MB_LEADER.SelectedValue) Then
                iRowCount = mbMEMBERList.COUNT_2(D_BDay, D_EDay)
            ElseIf isPeriod AndAlso Utility.isValidateData(Me.DDL_MB_AREA.SelectedValue) AndAlso Not Utility.isValidateData(Me.DDL_MB_LEADER.SelectedValue) Then
                iRowCount = mbMEMBERList.COUNT_3(D_BDay, D_EDay, Me.DDL_MB_AREA.SelectedValue)
            ElseIf isPeriod AndAlso Not Utility.isValidateData(Me.DDL_MB_AREA.SelectedValue) AndAlso Utility.isValidateData(Me.DDL_MB_LEADER.SelectedValue) Then
                iRowCount = mbMEMBERList.COUNT_4(D_BDay, D_EDay, Me.DDL_MB_LEADER.SelectedValue)
            ElseIf isPeriod AndAlso Utility.isValidateData(Me.DDL_MB_AREA.SelectedValue) AndAlso Utility.isValidateData(Me.DDL_MB_LEADER.SelectedValue) Then
                iRowCount = mbMEMBERList.COUNT_5(D_BDay, D_EDay, Me.DDL_MB_AREA.SelectedValue, Me.DDL_MB_LEADER.SelectedValue)
            ElseIf Not isPeriod AndAlso Utility.isValidateData(Me.DDL_MB_AREA.SelectedValue) AndAlso Not Utility.isValidateData(Me.DDL_MB_LEADER.SelectedValue) Then
                iRowCount = mbMEMBERList.COUNT_6(Me.DDL_MB_AREA.SelectedValue)
            ElseIf Not isPeriod AndAlso Utility.isValidateData(Me.DDL_MB_AREA.SelectedValue) AndAlso Utility.isValidateData(Me.DDL_MB_LEADER.SelectedValue) Then
                iRowCount = mbMEMBERList.COUNT_7(Me.DDL_MB_AREA.SelectedValue, Me.DDL_MB_LEADER.SelectedValue)
            End If

            '目前在第?頁
            Me.lblCurrentPage.Text = iPageNUM.ToString

            '每頁?筆
            Me.lblPageSize.Text = Me.m_PageSize.ToString

            '共?筆
            Me.lblTotleSize.Text = iRowCount.ToString

            Me.DDL_Page.Items.Clear()

            Dim iTopNum As Integer = iRowCount \ Me.m_PageSize
            If (iRowCount Mod Me.m_PageSize) > 0 Then
                iTopNum += 1
            End If
            Me.lblTotalPage.Text = iTopNum.ToString

            For i As Integer = 1 To iTopNum
                Dim objListItem As New ListItem(i.ToString, i.ToString)
                Me.DDL_Page.Items.Add(objListItem)
            Next

            Me.DDL_Page.SelectedIndex = -1
            If Not IsNothing(Me.DDL_Page.Items.FindByValue(iPageNUM.ToString)) Then
                Me.DDL_Page.Items.FindByValue(iPageNUM.ToString).Selected = True
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Sub initCHGPage_2(ByVal iPageNUM As Integer)
        Try
            Dim iRowCount As Decimal = 0

            Dim D_BDay As Object = Convert.DBNull
            D_BDay = Me.ConvertWDate(Me.TXT_BY.Text, Me.TXT_BM.Text, Me.TXT_BD.Text)
            Dim D_EDay As Object = Convert.DBNull
            D_EDay = Me.ConvertWDate(Me.TXT_EY.Text, Me.TXT_EM.Text, Me.TXT_ED.Text)

            Dim isPeriod As Boolean = False
            If IsDate(D_BDay) AndAlso IsDate(D_EDay) Then
                isPeriod = True
            End If

            Dim mbMEMREVList As New MB_MEMREVList(Me.m_DBManager)
            If Not isPeriod AndAlso Not Utility.isValidateData(Me.DDL_MB_AREA.SelectedValue) AndAlso Not Utility.isValidateData(Me.DDL_MB_LEADER.SelectedValue) Then
                iRowCount = mbMEMREVList.COUNT_1()
            ElseIf isPeriod AndAlso Not Utility.isValidateData(Me.DDL_MB_AREA.SelectedValue) AndAlso Not Utility.isValidateData(Me.DDL_MB_LEADER.SelectedValue) Then
                iRowCount = mbMEMREVList.COUNT_2(D_BDay, D_EDay)
            ElseIf isPeriod AndAlso Utility.isValidateData(Me.DDL_MB_AREA.SelectedValue) AndAlso Not Utility.isValidateData(Me.DDL_MB_LEADER.SelectedValue) Then
                iRowCount = mbMEMREVList.COUNT_3(D_BDay, D_EDay, Me.DDL_MB_AREA.SelectedValue)
            ElseIf isPeriod AndAlso Not Utility.isValidateData(Me.DDL_MB_AREA.SelectedValue) AndAlso Utility.isValidateData(Me.DDL_MB_LEADER.SelectedValue) Then
                iRowCount = mbMEMREVList.COUNT_4(D_BDay, D_EDay, Me.DDL_MB_LEADER.SelectedValue)
            ElseIf isPeriod AndAlso Utility.isValidateData(Me.DDL_MB_AREA.SelectedValue) AndAlso Utility.isValidateData(Me.DDL_MB_LEADER.SelectedValue) Then
                iRowCount = mbMEMREVList.COUNT_5(D_BDay, D_EDay, Me.DDL_MB_AREA.SelectedValue, Me.DDL_MB_LEADER.SelectedValue)
            ElseIf Not isPeriod AndAlso Utility.isValidateData(Me.DDL_MB_AREA.SelectedValue) AndAlso Not Utility.isValidateData(Me.DDL_MB_LEADER.SelectedValue) Then
                iRowCount = mbMEMREVList.COUNT_6(Me.DDL_MB_AREA.SelectedValue)
            ElseIf Not isPeriod AndAlso Utility.isValidateData(Me.DDL_MB_AREA.SelectedValue) AndAlso Utility.isValidateData(Me.DDL_MB_LEADER.SelectedValue) Then
                iRowCount = mbMEMREVList.COUNT_7(Me.DDL_MB_AREA.SelectedValue, Me.DDL_MB_LEADER.SelectedValue)
            End If

            '目前在第?頁
            Me.lblCurrentPage.Text = iPageNUM.ToString

            '每頁?筆
            Me.lblPageSize.Text = Me.m_PageSize.ToString

            '共?筆
            Me.lblTotleSize.Text = iRowCount.ToString

            Me.DDL_Page.Items.Clear()

            Dim iTopNum As Integer = iRowCount \ Me.m_PageSize
            If (iRowCount Mod Me.m_PageSize) > 0 Then
                iTopNum += 1
            End If
            Me.lblTotalPage.Text = iTopNum.ToString

            For i As Integer = 1 To iTopNum
                Dim objListItem As New ListItem(i.ToString, i.ToString)
                Me.DDL_Page.Items.Add(objListItem)
            Next

            Me.DDL_Page.SelectedIndex = -1
            If Not IsNothing(Me.DDL_Page.Items.FindByValue(iPageNUM.ToString)) Then
                Me.DDL_Page.Items.FindByValue(iPageNUM.ToString).Selected = True
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Sub initCHGPage_3(ByVal iPageNUM As Integer)
        Try
            Dim iRowCount As Decimal = 0

            Dim iMB_SEQ As Decimal = 0
            iMB_SEQ = CDec(Me.DDL_MB_CLASS.SelectedValue)

            Dim mbMEMCLASSList As New MB_MEMCLASSList(Me.m_DBManager)
            iRowCount = mbMEMCLASSList.COUNT_MB_MEMCLASS(iMB_SEQ)

            '目前在第?頁
            Me.lblCurrentPage.Text = iPageNUM.ToString

            '每頁?筆
            Me.lblPageSize.Text = Me.m_PageSize.ToString

            '共?筆
            Me.lblTotleSize.Text = iRowCount.ToString

            Me.DDL_Page.Items.Clear()

            Dim iTopNum As Integer = iRowCount \ Me.m_PageSize
            If (iRowCount Mod Me.m_PageSize) > 0 Then
                iTopNum += 1
            End If
            Me.lblTotalPage.Text = iTopNum.ToString

            For i As Integer = 1 To iTopNum
                Dim objListItem As New ListItem(i.ToString, i.ToString)
                Me.DDL_Page.Items.Add(objListItem)
            Next

            Me.DDL_Page.SelectedIndex = -1
            If Not IsNothing(Me.DDL_Page.Items.FindByValue(iPageNUM.ToString)) Then
                Me.DDL_Page.Items.FindByValue(iPageNUM.ToString).Selected = True
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Sub initCHGPage_4(ByVal iPageNUM As Integer)
        Try
            Dim iRowCount As Decimal = 0

            Dim D_BDay As Object = Convert.DBNull
            D_BDay = Me.ConvertWDate(Me.TXT_BY.Text, Me.TXT_BM.Text, Me.TXT_BD.Text)
            Dim D_EDay As Object = Convert.DBNull
            D_EDay = Me.ConvertWDate(Me.TXT_EY.Text, Me.TXT_EM.Text, Me.TXT_ED.Text)

            Dim isPeriod As Boolean = False
            Dim iB_YYYYMMDD As Decimal = 0
            Dim iE_YYYYMMDD As Decimal = 0
            If IsDate(D_BDay) AndAlso IsDate(D_EDay) Then
                isPeriod = True

                iB_YYYYMMDD = CDate(D_BDay).Year * 10000 + CDate(D_BDay).Month * 100 + CDate(D_BDay).Day
                iE_YYYYMMDD = CDate(D_EDay).Year * 10000 + CDate(D_EDay).Month * 100 + CDate(D_EDay).Day
            End If

            Dim mbMEMREVList As New MB_MEMREVList(Me.m_DBManager)
            If isPeriod AndAlso Not Utility.isValidateData(Me.MB_MEMTYP.SelectedValue) AndAlso Not Utility.isValidateData(Me.MB_FEETYPE.SelectedValue) Then
                iRowCount = mbMEMREVList.COUNT_INCOME_1(iB_YYYYMMDD, iE_YYYYMMDD)
            ElseIf isPeriod AndAlso Utility.isValidateData(Me.MB_MEMTYP.SelectedValue) AndAlso Utility.isValidateData(Me.MB_FEETYPE.SelectedValue) Then
                iRowCount = mbMEMREVList.COUNT_INCOME_2(iB_YYYYMMDD, iE_YYYYMMDD, Me.MB_MEMTYP.SelectedValue, Me.MB_FEETYPE.SelectedValue)
            ElseIf isPeriod AndAlso Not Utility.isValidateData(Me.MB_MEMTYP.SelectedValue) AndAlso Utility.isValidateData(Me.MB_FEETYPE.SelectedValue) Then
                iRowCount = mbMEMREVList.COUNT_INCOME_3(iB_YYYYMMDD, iE_YYYYMMDD, Me.MB_FEETYPE.SelectedValue)
            ElseIf Not isPeriod AndAlso Utility.isValidateData(Me.MB_MEMTYP.SelectedValue) AndAlso Utility.isValidateData(Me.MB_FEETYPE.SelectedValue) Then
                iRowCount = mbMEMREVList.COUNT_INCOME_4(Me.MB_MEMTYP.SelectedValue, Me.MB_FEETYPE.SelectedValue)
            ElseIf Not isPeriod AndAlso Utility.isValidateData(Me.MB_MEMTYP.SelectedValue) AndAlso Not Utility.isValidateData(Me.MB_FEETYPE.SelectedValue) Then
                iRowCount = mbMEMREVList.COUNT_INCOME_5(Me.MB_MEMTYP.SelectedValue)
            ElseIf Not isPeriod AndAlso Not Utility.isValidateData(Me.MB_MEMTYP.SelectedValue) AndAlso Utility.isValidateData(Me.MB_FEETYPE.SelectedValue) Then
                iRowCount = mbMEMREVList.COUNT_INCOME_6(Me.MB_FEETYPE.SelectedValue)
            ElseIf Not isPeriod AndAlso Not Utility.isValidateData(Me.MB_MEMTYP.SelectedValue) AndAlso Not Utility.isValidateData(Me.MB_FEETYPE.SelectedValue) Then
                iRowCount = mbMEMREVList.COUNT_INCOME_7()
            End If

            '目前在第?頁
            Me.lblCurrentPage.Text = iPageNUM.ToString

            '每頁?筆
            Me.lblPageSize.Text = Me.m_PageSize.ToString

            '共?筆
            Me.lblTotleSize.Text = iRowCount.ToString

            Me.DDL_Page.Items.Clear()

            Dim iTopNum As Integer = iRowCount \ Me.m_PageSize
            If (iRowCount Mod Me.m_PageSize) > 0 Then
                iTopNum += 1
            End If
            Me.lblTotalPage.Text = iTopNum.ToString

            For i As Integer = 1 To iTopNum
                Dim objListItem As New ListItem(i.ToString, i.ToString)
                Me.DDL_Page.Items.Add(objListItem)
            Next

            Me.DDL_Page.SelectedIndex = -1
            If Not IsNothing(Me.DDL_Page.Items.FindByValue(iPageNUM.ToString)) Then
                Me.DDL_Page.Items.FindByValue(iPageNUM.ToString).Selected = True
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Sub initCHGPage_5(ByVal iPageNUM As Integer)
        Try
            Dim iRowCount As Decimal = 0

            Dim isPeriod As Boolean = False
            If IsNumeric(Me.TXT_B_YYY.Text) AndAlso IsNumeric(Me.TXT_E_YYY.Text) Then
                isPeriod = True
            End If

            Dim mbMEMFEEList As New MB_MEMFEEList(Me.m_DBManager)
            If Not isPeriod AndAlso Not Utility.isValidateData(Me.DDL_MB_AREA.SelectedValue) AndAlso Not Utility.isValidateData(Me.DDL_MB_LEADER.SelectedValue) Then
                iRowCount = mbMEMFEEList.COUNT_1()
            ElseIf isPeriod AndAlso Not Utility.isValidateData(Me.DDL_MB_AREA.SelectedValue) AndAlso Not Utility.isValidateData(Me.DDL_MB_LEADER.SelectedValue) Then
                iRowCount = mbMEMFEEList.COUNT_2(CDec(Me.TXT_B_YYY.Text), CDec(Me.TXT_E_YYY.Text))
            ElseIf isPeriod AndAlso Utility.isValidateData(Me.DDL_MB_AREA.SelectedValue) AndAlso Not Utility.isValidateData(Me.DDL_MB_LEADER.SelectedValue) Then
                iRowCount = mbMEMFEEList.COUNT_3(CDec(Me.TXT_B_YYY.Text), CDec(Me.TXT_E_YYY.Text), Me.DDL_MB_AREA.SelectedValue)
            ElseIf isPeriod AndAlso Utility.isValidateData(Me.DDL_MB_AREA.SelectedValue) AndAlso Utility.isValidateData(Me.DDL_MB_LEADER.SelectedValue) Then
                iRowCount = mbMEMFEEList.COUNT_4(CDec(Me.TXT_B_YYY.Text), CDec(Me.TXT_E_YYY.Text), Me.DDL_MB_AREA.SelectedValue, Me.DDL_MB_LEADER.SelectedValue)
            ElseIf Not isPeriod AndAlso Utility.isValidateData(Me.DDL_MB_AREA.SelectedValue) AndAlso Utility.isValidateData(Me.DDL_MB_LEADER.SelectedValue) Then
                iRowCount = mbMEMFEEList.COUNT_5(Me.DDL_MB_AREA.SelectedValue, Me.DDL_MB_LEADER.SelectedValue)
            ElseIf Not isPeriod AndAlso Utility.isValidateData(Me.DDL_MB_AREA.SelectedValue) AndAlso Not Utility.isValidateData(Me.DDL_MB_LEADER.SelectedValue) Then
                iRowCount = mbMEMFEEList.COUNT_6(Me.DDL_MB_AREA.SelectedValue)
            ElseIf Not isPeriod AndAlso Not Utility.isValidateData(Me.DDL_MB_AREA.SelectedValue) AndAlso Utility.isValidateData(Me.DDL_MB_LEADER.SelectedValue) Then
                iRowCount = mbMEMFEEList.COUNT_7(Me.DDL_MB_LEADER.SelectedValue)
            End If

            '目前在第?頁
            Me.lblCurrentPage.Text = iPageNUM.ToString

            '每頁?筆
            Me.lblPageSize.Text = Me.m_PageSize.ToString

            '共?筆
            Me.lblTotleSize.Text = iRowCount.ToString

            Me.DDL_Page.Items.Clear()

            Dim iTopNum As Integer = iRowCount \ Me.m_PageSize
            If (iRowCount Mod Me.m_PageSize) > 0 Then
                iTopNum += 1
            End If
            Me.lblTotalPage.Text = iTopNum.ToString

            For i As Integer = 1 To iTopNum
                Dim objListItem As New ListItem(i.ToString, i.ToString)
                Me.DDL_Page.Items.Add(objListItem)
            Next

            Me.DDL_Page.SelectedIndex = -1
            If Not IsNothing(Me.DDL_Page.Items.FindByValue(iPageNUM.ToString)) Then
                Me.DDL_Page.Items.FindByValue(iPageNUM.ToString).Selected = True
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Sub initCHGPage_6(ByVal iPageNUM As Integer)
        Try
            Dim iRowCount As Decimal = 0

            Dim mbVIPList As New MB_VIPList(Me.m_DBManager)
            If Not Utility.isValidateData(Me.DDL_MB_AREA.SelectedValue) AndAlso Not Utility.isValidateData(Me.DDL_MB_LEADER.SelectedValue) Then
                iRowCount = mbVIPList.COUNT_1()
            ElseIf Utility.isValidateData(Me.DDL_MB_AREA.SelectedValue) AndAlso Not Utility.isValidateData(Me.DDL_MB_LEADER.SelectedValue) Then
                iRowCount = mbVIPList.COUNT_2(Me.DDL_MB_AREA.SelectedValue)
            ElseIf Utility.isValidateData(Me.DDL_MB_AREA.SelectedValue) AndAlso Utility.isValidateData(Me.DDL_MB_LEADER.SelectedValue) Then
                iRowCount = mbVIPList.COUNT_3(Me.DDL_MB_AREA.SelectedValue, Me.DDL_MB_LEADER.SelectedValue)
            ElseIf Utility.isValidateData(Me.DDL_MB_AREA.SelectedValue) AndAlso Not Utility.isValidateData(Me.DDL_MB_LEADER.SelectedValue) Then
                iRowCount = mbVIPList.COUNT_4(Me.DDL_MB_AREA.SelectedValue)
            End If

            '目前在第?頁
            Me.lblCurrentPage.Text = iPageNUM.ToString

            '每頁?筆
            Me.lblPageSize.Text = Me.m_PageSize.ToString

            '共?筆
            Me.lblTotleSize.Text = iRowCount.ToString

            Me.DDL_Page.Items.Clear()

            Dim iTopNum As Integer = iRowCount \ Me.m_PageSize
            If (iRowCount Mod Me.m_PageSize) > 0 Then
                iTopNum += 1
            End If
            Me.lblTotalPage.Text = iTopNum.ToString

            For i As Integer = 1 To iTopNum
                Dim objListItem As New ListItem(i.ToString, i.ToString)
                Me.DDL_Page.Items.Add(objListItem)
            Next

            Me.DDL_Page.SelectedIndex = -1
            If Not IsNothing(Me.DDL_Page.Items.FindByValue(iPageNUM.ToString)) Then
                Me.DDL_Page.Items.FindByValue(iPageNUM.ToString).Selected = True
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub

    '下一頁
    Sub btnLinkNext_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnLinkNext.Click, btnImgNext.Click
        Try
            Dim iRowCount As Integer = CDec(Me.lblTotleSize.Text)
            Dim iTopNum As Integer = iRowCount \ Me.m_PageSize
            If (iRowCount Mod Me.m_PageSize) > 0 Then
                iTopNum += 1
            End If

            Dim iPageNUM As Integer = CDec(Me.lblCurrentPage.Text) + 1
            If iPageNUM >= iTopNum Then
                iPageNUM = iTopNum
            End If

            Select Case Me.DDL_TYPE.SelectedValue
                Case "1"
                    Me.Bind_TYPE_1(iPageNUM)
                Case "2"
                    Me.Bind_TYPE_2(iPageNUM)
                Case "3"
                    Me.Bind_TYPE_3(iPageNUM)
                Case "4"
                    Me.Bind_TYPE_4(iPageNUM)
                Case "5"
                    Me.Bind_TYPE_5(iPageNUM)
                Case "6"
                    Me.Bind_TYPE_6(iPageNUM)
            End Select
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    '上一頁
    Sub btnLinkPrev_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnLinkPrev.Click, btnImgPrev.Click
        Try
            Dim iPageNUM As Integer = CDec(Me.lblCurrentPage.Text) - 1
            If iPageNUM <= 0 Then
                iPageNUM = 1
            End If

            Select Case Me.DDL_TYPE.SelectedValue
                Case "1"
                    Me.Bind_TYPE_1(iPageNUM)
                Case "2"
                    Me.Bind_TYPE_2(iPageNUM)
                Case "3"
                    Me.Bind_TYPE_3(iPageNUM)
                Case "4"
                    Me.Bind_TYPE_4(iPageNUM)
                Case "5"
                    Me.Bind_TYPE_5(iPageNUM)
                Case "6"
                    Me.Bind_TYPE_6(iPageNUM)
            End Select
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    '第一頁
    Sub btnLinkFirst_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnLinkFirst.Click, btnImgFirst.Click
        Try
            Select Case Me.DDL_TYPE.SelectedValue
                Case "1"
                    Me.Bind_TYPE_1(1)
                Case "2"
                    Me.Bind_TYPE_2(1)
                Case "3"
                    Me.Bind_TYPE_3(1)
                Case "4"
                    Me.Bind_TYPE_4(1)
                Case "5"
                    Me.Bind_TYPE_5(1)
                Case "6"
                    Me.Bind_TYPE_6(1)
            End Select
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    '最後一頁
    Sub btnLinkLast_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnLinkLast.Click, btnImgLast.Click
        Try
            Dim iRowCount As Integer = CDec(Me.lblTotleSize.Text)
            Dim iTopNum As Integer = iRowCount \ Me.m_PageSize
            If (iRowCount Mod Me.m_PageSize) > 0 Then
                iTopNum += 1
            End If

            Select Case Me.DDL_TYPE.SelectedValue
                Case "1"
                    Me.Bind_TYPE_1(iTopNum)
                Case "2"
                    Me.Bind_TYPE_2(iTopNum)
                Case "3"
                    Me.Bind_TYPE_3(iTopNum)
                Case "4"
                    Me.Bind_TYPE_4(iTopNum)
                Case "5"
                    Me.Bind_TYPE_5(iTopNum)
                Case "6"
                    Me.Bind_TYPE_6(iTopNum)
            End Select
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    '下拉換頁
    Sub DDL_Page_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DDL_Page.SelectedIndexChanged
        Try
            Select Case Me.DDL_TYPE.SelectedValue
                Case "1"
                    Me.Bind_TYPE_1(CDec(Me.DDL_Page.SelectedValue))
                Case "2"
                    Me.Bind_TYPE_2(CDec(Me.DDL_Page.SelectedValue))
                Case "3"
                    Me.Bind_TYPE_3(CDec(Me.DDL_Page.SelectedValue))
                Case "4"
                    Me.Bind_TYPE_4(CDec(Me.DDL_Page.SelectedValue))
                Case "5"
                    Me.Bind_TYPE_5(CDec(Me.DDL_Page.SelectedValue))
                Case "5"
                    Me.Bind_TYPE_6(CDec(Me.DDL_Page.SelectedValue))
            End Select
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    Sub ClearPageValue()
        Try
            Me.lblCurrentPage.Text = "1"
            Me.lblTotalPage.Text = String.Empty
            Me.lblPageSize.Text = String.Empty
            Me.lblTotleSize.Text = String.Empty
            Me.DDL_Page.Items.Clear()
        Catch ex As Exception
            Throw
        End Try
    End Sub
#End Region

#Region "Utility"
    Function ConvertWDate(ByVal sYYY As String, ByVal sMM As String, ByVal sDD As String) As Object
        Try
            sYYY = Trim(sYYY)
            sMM = Trim(sMM)
            sDD = Trim(sDD)

            If IsNumeric(sYYY) AndAlso IsNumeric(sMM) AndAlso IsNumeric(sDD) Then
                Dim sYYYY As String = String.Empty
                sYYYY = CDec(sYYY).ToString

                Dim sDateStr As String = String.Empty
                sDateStr = sYYYY & "/" & sMM & "/" & sDD

                If IsDate(sDateStr) Then
                    Return CDate(sDateStr)
                Else
                    Return Convert.DBNull
                End If
            Else
                Return Convert.DBNull
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function

    Function isValidDate(ByVal sMsg As String, ByVal sYYY As String, ByVal sMM As String, ByVal sDD As String) As String
        Try
            sYYY = Trim(sYYY)
            sMM = Trim(sMM)
            sDD = Trim(sDD)

            If IsNumeric(sYYY) OrElse IsNumeric(sMM) OrElse IsNumeric(sDD) Then
                Dim sYYYY As String = String.Empty
                sYYYY = CDec(sYYY).ToString

                Dim sDateStr As String = String.Empty
                sDateStr = sYYYY & "/" & sMM & "/" & sDD

                If IsDate(sDateStr) Then
                    Return String.Empty
                Else
                    Return sMsg & "日期格式錯誤"
                End If
            End If

            Return String.Empty
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function getMB_MEMSEQ(ByVal iMB_MEMSEQ As Object, ByVal sMB_AREA As Object) As String
        Try
            If IsNumeric(iMB_MEMSEQ) AndAlso Utility.isValidateData(sMB_AREA) Then
                Return sMB_AREA & "-" & com.Azion.EloanUtility.StrUtility.FillZero(CDec(iMB_MEMSEQ), 7)
            End If
            Return String.Empty
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function getMB_FAMILY(ByVal sMB_FAMILY As Object) As String
        Try
            If Utility.isValidateData(sMB_FAMILY) Then
                If IsNothing(Me.m_DT_UpCode28) Then
                    Dim apCODEList As New AP_CODEList(Me.m_DBManager)
                    apCODEList.loadByUpCode(Me.m_sUpCode28)
                    Me.m_DT_UpCode28 = apCODEList.getCurrentDataSet.Tables(0)
                End If

                Dim ROW_SELECT() As DataRow = Nothing
                ROW_SELECT = Me.m_DT_UpCode28.Select("VALUE='" & sMB_FAMILY & "'")
                If Not IsNothing(ROW_SELECT) AndAlso ROW_SELECT.Length > 0 Then
                    Return ROW_SELECT(0)("TEXT").ToString
                End If
            End If

            Return String.Empty
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function getMB_AREA(ByVal sMB_AREA As Object) As String
        Try
            If Utility.isValidateData(sMB_AREA) Then
                If IsNothing(Me.m_DT_UpCode7) Then
                    Dim apCODEList As New AP_CODEList(Me.m_DBManager)
                    apCODEList.loadByUpCode(Me.m_sUpCode7)
                    Me.m_DT_UpCode7 = apCODEList.getCurrentDataSet.Tables(0)
                End If

                Dim ROW_SELECT() As DataRow = Nothing
                ROW_SELECT = Me.m_DT_UpCode7.Select("VALUE='" & sMB_AREA & "'")
                If Not IsNothing(ROW_SELECT) AndAlso ROW_SELECT.Length > 0 Then
                    Return ROW_SELECT(0)("TEXT").ToString
                End If
            End If

            Return String.Empty
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function getMB_ITEMID(ByVal sMB_ITEMID As Object) As String
        Try
            If Utility.isValidateData(sMB_ITEMID) Then
                If IsNothing(Me.DT_UPCODE15) Then
                    Dim apCODEList As New AP_CODEList(Me.m_DBManager)
                    apCODEList.loadByUpCode(Me.m_sUpcode15)
                    Me.DT_UPCODE15 = apCODEList.getCurrentDataSet.Tables(0)
                End If

                Dim ROW_Select() As DataRow = Me.DT_UPCODE15.Select("VALUE='" & sMB_ITEMID & "'")
                If Not IsNothing(ROW_Select) AndAlso ROW_Select.Length > 0 Then
                    Return ROW_Select(0)("TEXT").ToString()
                End If
            End If

            Return String.Empty
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function get_DECIMAL_DATE(ByVal MB_TX_DATE As Object) As String
        Try
            If Utility.isDecimalDate(MB_TX_DATE) Then
                Return Utility.DateTransfer(Left(MB_TX_DATE.ToString, 4) & "/" & MB_TX_DATE.ToString.Substring(4, 2) & "/" & Right(MB_TX_DATE.ToString, 2))
            End If

            Return String.Empty
        Catch ex As Exception
            Throw
        End Try
    End Function

    Function getMB_FEETYPE(ByVal sMB_FEETYPE As Object) As String
        Try
            If Utility.isValidateData(sMB_FEETYPE) Then
                If IsNothing(Me.m_DT_UpCode23) Then
                    Dim apCODEList As New AP_CODEList(Me.m_DBManager)
                    apCODEList.loadByUpCode(Me.m_sUpcode23)
                    Me.m_DT_UpCode23 = apCODEList.getCurrentDataSet.Tables(0)
                End If

                Dim ROW_SELECT() As DataRow = Me.m_DT_UpCode23.Select("VALUE='" & sMB_FEETYPE & "'")
                If Not IsNothing(ROW_SELECT) AndAlso ROW_SELECT.Length > 0 Then
                    Return ROW_SELECT(0)("TEXT")
                End If
            End If

            Return String.Empty
        Catch ex As Exception
            Throw
        End Try
    End Function

    'Public Function getMB_CHKFLAG(ByVal sMB_CHKFLAG As Object) As String
    '    Try
    '        If Utility.isValidateData(sMB_CHKFLAG) Then
    '            If sMB_CHKFLAG = "1" Then
    '                Return "正取"
    '            ElseIf sMB_CHKFLAG = "2" Then
    '                Return "備取"
    '            End If
    '        End If

    '        Return String.Empty
    '    Catch ex As Exception
    '        Throw
    '    End Try
    'End Function

    Public Function getMB_CHKFLAG(ByVal iMB_SEQ As Object, ByVal iMB_MEMSEQ As Object, ByVal sMB_APV As Object, ByVal sMB_FWMK As Object, ByVal MB_CDATE As Object, ByVal MB_CREDATETIME As Object, ByVal MB_APVDATETIME As Object) As String
        Try
            If sMB_APV.ToString = "2" Then
                '不須核准
                If Utility.isValidateData(sMB_FWMK) AndAlso sMB_FWMK = "0" Then
                    Return "審核中<BR/>" & Utility.FMT_MBSC_DATE(MB_CREDATETIME)
                ElseIf Utility.isValidateData(sMB_FWMK) AndAlso sMB_FWMK = "3" Then
                    Return "取消報名" & "<BR/>" & Utility.FMT_MBSC_DATE(MB_CDATE)
                ElseIf Utility.isValidateData(sMB_FWMK) AndAlso sMB_FWMK = "4" Then
                    Dim sSECCancel As String = String.Empty
                    sSECCancel = "第二次取消報名"
                    If IsDate(MB_CDATE.ToString) Then
                        sSECCancel &= "<BR/>" & Utility.FMT_MBSC_DATE(MB_CDATE)
                    End If
                    Return sSECCancel
                Else
                    If IsNumeric(iMB_SEQ) AndAlso IsNumeric(iMB_MEMSEQ) Then
                        Dim sb As New System.Text.StringBuilder
                        Dim MB_CLASSList As New MB_CLASSList(Me.m_DBManager)
                        MB_CLASSList.setSQLCondition(" ORDER BY MB_BATCH ")
                        MB_CLASSList.LoadByMB_SEQ(iMB_SEQ)
                        For Each ROW As DataRow In MB_CLASSList.getCurrentDataSet.Tables(0).Rows
                            Dim DT_SEQ As DataTable = Nothing
                            Dim DT_FULL As DataTable = Nothing
                            Try
                                Dim sBATCH As String = String.Empty
                                If ROW("MB_BATCH") > 0 Then
                                    sBATCH = "第" & ROW("MB_BATCH").ToString & "梯次"
                                End If
                                Dim iMB_FULL As Decimal = 0
                                iMB_FULL = Utility.CheckNumNull(ROW("MB_FULL"))
                                Dim iMB_WAIT As Decimal = 0
                                iMB_WAIT = Utility.CheckNumNull(ROW("MB_WAIT"))

                                Dim mbMEMCLASSList As New MB_MEMCLASSList(Me.m_DBManager)
                                mbMEMCLASSList.loadByMB_SEQ(iMB_SEQ, ROW("MB_BATCH"))
                                Dim DV_SEQ As New DataView(mbMEMCLASSList.getCurrentDataSet.Tables(0), "ISNULL(MB_FWMK,' ') NOT IN ('3','4','5') AND ISNULL(MB_ELECT,' ')='1' AND ISNULL(MB_RESP,' ')<>'N'", "MB_CREDATETIME", DataViewRowState.CurrentRows)
                                DT_SEQ = DV_SEQ.ToTable
                                DT_FULL = DT_SEQ.Clone
                                Dim iFULL As Decimal = 0
                                For i As Integer = 0 To DT_SEQ.Rows.Count - 1
                                    iFULL += 1

                                    If iFULL <= iMB_FULL Then
                                        DT_FULL.ImportRow(DT_SEQ.Rows(i))
                                    End If
                                Next
                                Dim ROW_FULL() As DataRow = Nothing
                                ROW_FULL = DT_FULL.Select("MB_MEMSEQ=" & iMB_MEMSEQ)
                                If Not IsNothing(ROW_FULL) AndAlso ROW_FULL.Length > 0 Then
                                    sb.Append(sBATCH).Append("正取").Append("<BR/>")
                                Else
                                    sb.Append(sBATCH).Append("備取").Append("<BR/>")
                                End If
                            Catch ex As Exception
                                Throw
                            Finally
                                If Not IsNothing(DT_SEQ) Then
                                    DT_SEQ.Dispose()
                                End If

                                If Not IsNothing(DT_FULL) Then
                                    DT_FULL.Dispose()
                                End If
                            End Try
                        Next

                        Dim sMSG As String = String.Empty
                        sMSG = sb.ToString
                        If Utility.isValidateData(sMSG) Then
                            Dim sMB_CREDATETIME As String = String.Empty
                            If IsDate(MB_CREDATETIME.ToString) Then
                                sMSG &= Utility.FMT_MBSC_DATE(MB_CREDATETIME)
                            End If
                            Return sMSG
                        End If
                    End If
                End If
            Else
                '需核准
                If Utility.isValidateData(sMB_FWMK) AndAlso sMB_FWMK = "3" Then
                    Return "取消報名" & "<BR/>" & Utility.FMT_MBSC_DATE(MB_CDATE)
                ElseIf Utility.isValidateData(sMB_FWMK) AndAlso sMB_FWMK = "4" Then
                    Dim sSECCancel As String = String.Empty
                    sSECCancel = "第二次取消報名"
                    If IsDate(MB_CDATE.ToString) Then
                        sSECCancel &= "<BR/>" & Utility.FMT_MBSC_DATE(MB_CDATE)
                    End If
                    Return sSECCancel
                Else
                    Dim MB_MEMBATCHList As New MB_MEMBATCHList(Me.m_DBManager)
                    MB_MEMBATCHList.LoadBySEQ(iMB_MEMSEQ, iMB_SEQ)
                    Dim ROW_1() As DataRow = Nothing
                    ROW_1 = MB_MEMBATCHList.getCurrentDataSet.Tables(0).Select("ISNULL(MB_ELECT,' ')='1'", "MB_BATCH")
                    If Not IsNothing(ROW_1) AndAlso ROW_1.Length > 0 Then
                        Dim sMSG As String = String.Empty
                        For i As Integer = 0 To UBound(ROW_1)
                            Dim sBR As String = String.Empty
                            If i < UBound(ROW_1) Then
                                sBR = "<BR/>"
                            End If
                            Dim sINFO As String = String.Empty
                            Select Case ROW_1(i)("MB_CHKFLAG").ToString
                                Case "1"
                                    sINFO = "錄取"
                                Case "2"
                                    sINFO = "不准"
                                Case Else
                                    sINFO = "審核中"
                            End Select
                            sMSG &= "<span style='color:red;font-weight:bold'>" & ROW_1(i)("MB_BATCH").ToString & "</SPAN>" & sINFO & sBR
                        Next
                        Return sMSG
                    End If
                End If
            End If
            Return String.Empty
        Catch ex As Exception
            Throw
        End Try
    End Function

    Protected Function getFEEAmt(ByVal iAMT As Object) As String
        Try
            If IsNumeric(iAMT) Then
                Return Format(CDec(iAMT), "#,##0")
            Else
                Return "0"
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function
#End Region

#Region "REPEATER EVENTS"
    Private Sub RP_1_ItemCommand(source As Object, e As Web.UI.WebControls.RepeaterCommandEventArgs) Handles RP_1.ItemCommand
        Dim sURL As String = String.Empty
        sURL = com.Azion.EloanUtility.UIUtility.getRootPath & "/MNT/MBMnt_Member_01_v01.aspx?OPTYPE=A" & _
               "&MB_MEMSEQ=" & e.CommandArgument
        com.Azion.EloanUtility.UIUtility.showModalDialog(sURL)
    End Sub

    Private Sub RP_2_ItemDataBound(sender As Object, e As RepeaterItemEventArgs) Handles RP_2.ItemDataBound
        Dim DT_DTL As DataTable = Nothing
        Try
            If e.Item.ItemType = ListItemType.Item OrElse e.Item.ItemType = ListItemType.AlternatingItem Then
                Dim RP_2_DTL As Repeater = e.Item.FindControl("RP_2_DTL")

                Dim mbMEMREVList As New MB_MEMREVList(Me.m_DBManager)

                Dim D_BDay As Object = Convert.DBNull
                D_BDay = Me.ConvertWDate(Me.TXT_BY.Text, Me.TXT_BM.Text, Me.TXT_BD.Text)
                Dim D_EDay As Object = Convert.DBNull
                D_EDay = Me.ConvertWDate(Me.TXT_EY.Text, Me.TXT_EM.Text, Me.TXT_ED.Text)

                Dim isPeriod As Boolean = False
                If IsDate(D_BDay) AndAlso IsDate(D_EDay) Then
                    isPeriod = True
                End If

                Dim HID_MB_MEMSEQ As HtmlInputHidden = e.Item.FindControl("HID_MB_MEMSEQ")

                If Not isPeriod AndAlso Not Utility.isValidateData(Me.DDL_MB_AREA.SelectedValue) AndAlso Not Utility.isValidateData(Me.DDL_MB_LEADER.SelectedValue) Then
                    mbMEMREVList.DTL_1(CDec(HID_MB_MEMSEQ.Value))
                ElseIf isPeriod AndAlso Not Utility.isValidateData(Me.DDL_MB_AREA.SelectedValue) AndAlso Not Utility.isValidateData(Me.DDL_MB_LEADER.SelectedValue) Then
                    mbMEMREVList.DTL_2(CDec(HID_MB_MEMSEQ.Value), D_BDay, D_EDay)
                ElseIf isPeriod AndAlso Utility.isValidateData(Me.DDL_MB_AREA.SelectedValue) AndAlso Not Utility.isValidateData(Me.DDL_MB_LEADER.SelectedValue) Then
                    mbMEMREVList.DTL_2(CDec(HID_MB_MEMSEQ.Value), D_BDay, D_EDay)
                ElseIf isPeriod AndAlso Not Utility.isValidateData(Me.DDL_MB_AREA.SelectedValue) AndAlso Utility.isValidateData(Me.DDL_MB_LEADER.SelectedValue) Then
                    mbMEMREVList.DTL_2(CDec(HID_MB_MEMSEQ.Value), D_BDay, D_EDay)
                ElseIf isPeriod AndAlso Utility.isValidateData(Me.DDL_MB_AREA.SelectedValue) AndAlso Utility.isValidateData(Me.DDL_MB_LEADER.SelectedValue) Then
                    mbMEMREVList.DTL_2(CDec(HID_MB_MEMSEQ.Value), D_BDay, D_EDay)
                ElseIf Not isPeriod AndAlso Utility.isValidateData(Me.DDL_MB_AREA.SelectedValue) AndAlso Not Utility.isValidateData(Me.DDL_MB_LEADER.SelectedValue) Then
                    mbMEMREVList.DTL_1(CDec(HID_MB_MEMSEQ.Value))
                ElseIf Not isPeriod AndAlso Utility.isValidateData(Me.DDL_MB_AREA.SelectedValue) AndAlso Utility.isValidateData(Me.DDL_MB_LEADER.SelectedValue) Then
                    mbMEMREVList.DTL_1(CDec(HID_MB_MEMSEQ.Value))
                End If

                DT_DTL = mbMEMREVList.getCurrentDataSet.Tables(0)
                DT_DTL.Columns.Add("CNT", Type.GetType("System.Decimal"))

                Dim sqlRow As DataRow = DT_DTL.NewRow
                sqlRow("MB_TX_DATE") = Convert.DBNull
                sqlRow("MB_TOTFEE") = DT_DTL.Compute("SUM(MB_TOTFEE)", String.Empty)
                sqlRow("CNT") = DT_DTL.Rows.Count
                DT_DTL.Rows.Add(sqlRow)

                RP_2_DTL.DataSource = DT_DTL
                RP_2_DTL.DataBind()
            End If
        Catch ex As Exception
            Throw
        Finally
            If Not IsNothing(DT_DTL) Then
                DT_DTL.Dispose()
            End If
        End Try
    End Sub

    Sub RP_2_DTL_OnItemDataBound(ByVal Sender As Object, ByVal e As RepeaterItemEventArgs)
        Try
            If e.Item.ItemType = ListItemType.Item OrElse e.Item.ItemType = ListItemType.AlternatingItem Then
                Dim DRV As DataRowView = CType(e.Item.DataItem, DataRowView)
                Dim LTL_MB_TOTFEE As Literal = e.Item.FindControl("LTL_MB_TOTFEE")
                LTL_MB_TOTFEE.Text = Utility.FormatDec(DRV("MB_TOTFEE"), "#,##0")

                If IsDBNull(DRV("MB_TX_DATE")) AndAlso IsNumeric(DRV("CNT")) Then
                    Dim LTL_MB_TX_DATE As Literal = e.Item.FindControl("LTL_MB_TX_DATE")
                    LTL_MB_TX_DATE.Text = "小計"

                    Dim TD_1 As HtmlTableCell = e.Item.FindControl("TD_1")
                    TD_1.ColSpan = 6
                    TD_1.Attributes("Class") = "th1c_b"
                    Dim TD_2 As HtmlTableCell = e.Item.FindControl("TD_2")
                    TD_2.Visible = False
                    Dim TD_3 As HtmlTableCell = e.Item.FindControl("TD_3")
                    TD_3.Visible = False
                    Dim TD_4 As HtmlTableCell = e.Item.FindControl("TD_4")
                    TD_4.Visible = False
                    Dim TD_5 As HtmlTableCell = e.Item.FindControl("TD_5")
                    TD_5.Visible = False
                    Dim TD_6 As HtmlTableCell = e.Item.FindControl("TD_6")
                    TD_6.Visible = False
                    Dim TD_7 As HtmlTableCell = e.Item.FindControl("TD_7")
                    TD_7.Attributes("Class") = "th1r_b"
                    Dim TD_8 As HtmlTableCell = e.Item.FindControl("TD_8")
                    TD_8.Attributes("Class") = "th1r_b"
                End If
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Private Sub RP_4_ItemDataBound(sender As Object, e As Web.UI.WebControls.RepeaterItemEventArgs) Handles RP_4.ItemDataBound
        Try
            If e.Item.ItemType = ListItemType.Item OrElse e.Item.ItemType = ListItemType.AlternatingItem Then
                Dim DRV As DataRowView = CType(e.Item.DataItem, DataRowView)
                '區別
                Dim LTL_MB_AREA As Literal = e.Item.FindControl("LTL_MB_AREA")
                LTL_MB_AREA.Text = Me.getMB_AREA(DRV("MB_AREA").ToString)

                '日期
                Dim LTL_MB_TX_DATE As Literal = e.Item.FindControl("LTL_MB_TX_DATE")
                LTL_MB_TX_DATE.Text = Me.get_DECIMAL_DATE(DRV("MB_TX_DATE"))

                '會員編號
                Dim LTL_MB_MEMSEQ As Literal = e.Item.FindControl("LTL_MB_MEMSEQ")
                LTL_MB_MEMSEQ.Text = Me.getMB_MEMSEQ(DRV("MB_MEMSEQ"), DRV("MB_AREA"))

                '會費起訖日
                Dim LTL_MB_MEMFEE_SYMEYM As Literal = e.Item.FindControl("LTL_MB_MEMFEE_SYMEYM")
                Dim iMB_MEMFEE_SY As Decimal = 0
                If IsNumeric(DRV("MB_MEMFEE_SY")) Then
                    iMB_MEMFEE_SY = DRV("MB_MEMFEE_SY")
                End If
                Dim iMB_MEMFEE_SM As Decimal = 0
                If IsNumeric(DRV("MB_MEMFEE_SM")) Then
                    iMB_MEMFEE_SM = DRV("MB_MEMFEE_SM")
                End If
                Dim sSYM As String = iMB_MEMFEE_SY.ToString & "/" & iMB_MEMFEE_SM.ToString & "/1"
                Dim sBDate As String = String.Empty
                If IsDate(sSYM) Then
                    sBDate = Utility.FillZero(iMB_MEMFEE_SY, 4) & Utility.FillZero(iMB_MEMFEE_SM, 2)
                End If

                Dim iMB_MEMFEE_EY As Decimal = 0
                If IsNumeric(DRV("MB_MEMFEE_EY")) Then
                    iMB_MEMFEE_EY = DRV("MB_MEMFEE_EY")
                End If
                Dim iMB_MEMFEE_EM As Decimal = 0
                If IsNumeric(DRV("MB_MEMFEE_EM")) Then
                    iMB_MEMFEE_EM = DRV("MB_MEMFEE_EM")
                End If
                Dim sEYM As String = String.Empty
                sEYM = iMB_MEMFEE_EY.ToString & "/" & iMB_MEMFEE_EM.ToString & "/1"
                Dim sEDate As String = String.Empty
                If IsDate(sEYM) Then
                    sEDate = Utility.FillZero(iMB_MEMFEE_EY, 4) & Utility.FillZero(iMB_MEMFEE_EM, 2)
                End If
                LTL_MB_MEMFEE_SYMEYM.Text = sBDate & "~" & sEDate

                '會員種類
                Dim LTL_MB_MEMTYP As Literal = e.Item.FindControl("LTL_MB_MEMTYP")
                LTL_MB_MEMTYP.Text = Me.getMB_FAMILY(DRV("MB_MEMTYP"))

                '繳款方式
                Dim LTL_MB_FEETYPE As Literal = e.Item.FindControl("LTL_MB_FEETYPE")
                LTL_MB_FEETYPE.Text = Me.getMB_FEETYPE(DRV("MB_FEETYPE"))

                '每月金額
                Dim LTL_MB_MONFEE As Literal = e.Item.FindControl("LTL_MB_MONFEE")
                If IsNumeric(DRV("MB_MONFEE")) Then
                    LTL_MB_MONFEE.Text = Format(DRV("MB_MONFEE"), "#,##0")
                End If

                '繳款金額
                Dim LTL_MB_TOTFEE As Literal = e.Item.FindControl("LTL_MB_TOTFEE")
                If IsNumeric(DRV("MB_TOTFEE")) Then
                    LTL_MB_TOTFEE.Text = Format(DRV("MB_TOTFEE"), "#,##0")
                End If

                '護僧基金
                Dim LTL_FUND1 As Literal = e.Item.FindControl("LTL_FUND1")
                If DRV("FUND1").ToString = "1" Then
                    LTL_FUND1.Text = "V"
                End If

                '建設基金
                Dim LTL_FUND2 As Literal = e.Item.FindControl("LTL_FUND2")
                If DRV("FUND2").ToString = "1" Then
                    LTL_FUND2.Text = "V"
                End If

                '弘法基金
                Dim LTL_FUND3 As Literal = e.Item.FindControl("LTL_FUND3")
                If DRV("FUND3").ToString = "1" Then
                    LTL_FUND3.Text = "V"
                End If

                '列印收據
                Dim LTL_MB_PRINT As Literal = e.Item.FindControl("LTL_MB_PRINT")
                Select Case DRV("MB_PRINT").ToString
                    Case "1"
                        LTL_MB_PRINT.Text = "要列印"
                    Case "0"
                        LTL_MB_PRINT.Text = "不列印"
                    Case "Y"
                        LTL_MB_PRINT.Text = "已經列印"
                    Case Else
                        LTL_MB_PRINT.Text = "不列印"
                End Select

                '收據編號
                Dim LTL_MB_RECV As Literal = e.Item.FindControl("LTL_MB_RECV")
                LTL_MB_RECV.Text = DRV("MB_RECV_Y").ToString & DRV("MB_RECV_N").ToString
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Private Sub RP_6_ItemDataBound(sender As Object, e As RepeaterItemEventArgs) Handles RP_6.ItemDataBound
        Dim DT_MEMREV As DataTable = Nothing
        Try
            If e.Item.ItemType = ListItemType.Item OrElse e.Item.ItemType = ListItemType.AlternatingItem Then
                Dim LTL_MB_MEMSEQ As Literal = e.Item.FindControl("LTL_MB_MEMSEQ")

                Dim RP_6_DTL As Repeater = e.Item.FindControl("RP_6_DTL")

                Dim mbMEMREVList As New MB_MEMREVList(Me.m_DBManager)
                mbMEMREVList.LoadBySEQ(CDec(LTL_MB_MEMSEQ.Text))
                DT_MEMREV = mbMEMREVList.getCurrentDataSet.Tables(0)
                If IsNothing(DT_MEMREV) OrElse DT_MEMREV.Rows.Count = 0 Then
                    Dim objItem As RepeaterItem = RP_6_DTL.NamingContainer
                    Dim LTL_MB_MEMSEQ_P As Literal = objItem.FindControl("LTL_MB_MEMSEQ")
                    Dim LTL_MB_AREA_P As Literal = objItem.FindControl("LTL_MB_AREA")
                    Dim LTL_MB_LEADER_P As Literal = objItem.FindControl("LTL_MB_LEADER")
                    Dim LTL_MB_NAME_P As Literal = objItem.FindControl("LTL_MB_NAME")

                    Dim ROW_NEW As DataRow = DT_MEMREV.NewRow
                    ROW_NEW("MB_MEMSEQ") = LTL_MB_MEMSEQ_P.Text
                    ROW_NEW("MB_AREA") = LTL_MB_AREA_P.Text
                    ROW_NEW("MB_LEADER") = LTL_MB_LEADER_P.Text
                    ROW_NEW("MB_NAME") = LTL_MB_NAME_P.Text

                    DT_MEMREV.Rows.Add(ROW_NEW)
                End If

                Dim sqlRow As DataRow = DT_MEMREV.NewRow
                sqlRow("MB_MEMSEQ") = "-1"
                sqlRow("MB_TOTFEE") = Utility.CheckNumNull(DT_MEMREV.Compute("SUM(MB_TOTFEE)", String.Empty))
                DT_MEMREV.Rows.Add(sqlRow)

                RP_6_DTL.DataSource = DT_MEMREV
                RP_6_DTL.DataBind()
            End If
        Catch ex As Exception
            Throw
        Finally
            If Not IsNothing(DT_MEMREV) Then
                DT_MEMREV.Dispose()
            End If
        End Try
    End Sub

    Public Sub RP_6_DTL_OnItemDataBound(sender As Object, e As RepeaterItemEventArgs)
        Try
            If e.Item.ItemType = ListItemType.Item OrElse e.Item.ItemType = ListItemType.AlternatingItem Then
                Dim DRV As DataRowView = CType(e.Item.DataItem, DataRowView)
                '區別
                Dim LTL_MB_AREA As Literal = e.Item.FindControl("LTL_MB_AREA")
                LTL_MB_AREA.Text = Me.getMB_AREA(DRV("MB_AREA").ToString)

                '會員編號
                Dim LTL_MB_MEMSEQ As Literal = e.Item.FindControl("LTL_MB_MEMSEQ")
                If DRV("MB_MEMSEQ") = "-1" Then
                    LTL_MB_MEMSEQ.Text = "小計"
                End If

                '繳款日期
                Dim LTL_MB_TX_DATE As Literal = e.Item.FindControl("LTL_MB_TX_DATE")
                Dim sMB_TX_DATE As String = String.Empty
                sMB_TX_DATE = DRV("MB_TX_DATE").ToString
                If IsDate(sMB_TX_DATE) Then
                    LTL_MB_TX_DATE.Text = Utility.DateTransfer(CDate(sMB_TX_DATE))
                Else
                    LTL_MB_TX_DATE.Text = String.Empty
                End If

                '繳款金額
                Dim LTL_MB_TOTFEE As Literal = e.Item.FindControl("LTL_MB_TOTFEE")
                LTL_MB_TOTFEE.Text = Utility.FormatDec(DRV("MB_TOTFEE"), "#,##0")
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub
#End Region

    Private Sub RP_3_ItemCommand(source As Object, e As RepeaterCommandEventArgs) Handles RP_3.ItemCommand
        Try
            If e.CommandName = "SIGN" Then
                Dim LTL_MB_SEQ As Literal = e.Item.FindControl("LTL_MB_SEQ")
                Dim sURL As String = String.Empty
                Dim HID_MB_BATCH As HtmlInputHidden = e.Item.FindControl("HID_MB_BATCH")
                sURL = com.Azion.EloanUtility.UIUtility.getRootPath & "/MNT/MBMnt_Sign_01_v01.aspx?MBSEQ=" & LTL_MB_SEQ.Text & "&MB_MEMSEQ=" & e.CommandArgument &
                       "&OPTYPE=QRY&MB_BATCH=" & HID_MB_BATCH.Value

                com.Azion.EloanUtility.UIUtility.showOpen(sURL)
            ElseIf e.CommandName = "5Y" Then
                Dim LTL_MB_SEQ As Literal = e.Item.FindControl("LTL_MB_SEQ")
                Dim MB_MEMCLASS As New MB_MEMCLASS(Me.m_DBManager)
                If MB_MEMCLASS.LoadByPK(e.CommandArgument, LTL_MB_SEQ.Text) Then
                    'mail 錄取通知
                    Dim mbMEMBER As New MB_MEMBER(Me.m_DBManager)
                    If mbMEMBER.loadByPK(e.CommandArgument) AndAlso Utility.isValidateData(mbMEMBER.getAttribute("MB_EMAIL")) Then
                        Dim MB_CLASSList As New MB_CLASSList(Me.m_DBManager)
                        MB_CLASSList.LoadByMB_SEQ(LTL_MB_SEQ.Text)
                        If MB_CLASSList.getCurrentDataSet.Tables(0).Rows.Count > 0 Then
                            Dim ROW_CLASS As DataRow = MB_CLASSList.getCurrentDataSet.Tables(0).Rows(0)

                            Dim sMailTos() As String = {Trim(mbMEMBER.getString("MB_EMAIL"))}

                            Dim sMailSub As String = String.Empty
                            sMailSub = ROW_CLASS("MB_CLASS_NAME").ToString & " 錄取通知 請務必回覆是否出席"

                            Dim sMailBody As String = String.Empty
                            Dim sb As New System.Text.StringBuilder
                            sb.Append("<DIV style='text-align:center;font-size:24pt;color:red' >").Append(ROW_CLASS("MB_CLASS_NAME").ToString).Append("</DIV>")
                            Dim sForm As String = String.Empty
                            'If isWait Then
                            '    sForm = "備取通知"
                            'Else
                            '    sForm = "錄取通知"
                            'End If
                            sForm = "錄取通知"
                            sb.Append("<DIV style='text-align:center;font-size:24pt;color:black;font-weight:bold;text-decoration:underline' >").Append(sForm).Append("</DIV>")
                            sb.Append("<DIV style='font-size:12pt;color:black'>").Append("親愛的").Append(mbMEMBER.getString("MB_NAME")).Append("吉祥：").Append("</DIV>")

                            sb.Append("<DIV style='font-size:12pt;color:black;'>").Append("　　  歡迎您報名參加")
                            sb.Append("<span style='color:black'>")
                            sb.Append(ROW_CLASS("MB_CLASS_NAME").ToString)
                            sb.Append("</span>")
                            sb.Append("！此通知函，乃透過系統發送。為使您在課程期間能順利進行，並獲得最大收穫，請務必閱讀下列注意事項：").Append("</DIV>")
                            sb.Append("<ol type='1' style='font-size:12pt;color:black;' >")
                            Dim sMB_SDATE As String = String.Empty
                            If Utility.isValidateData(ROW_CLASS("MB_SDATE")) Then
                                sMB_SDATE = CDate(ROW_CLASS("MB_SDATE").ToString).Year & "年" & CDate(ROW_CLASS("MB_SDATE").ToString).Month & "月" &
                                    CDate(ROW_CLASS("MB_SDATE").ToString).Day & "日"
                            End If
                            Dim sREGTIME As String = String.Empty
                            sREGTIME = Left(Utility.FillZero(ROW_CLASS("REGTIME"), 4), 2) & Right(Utility.FillZero(ROW_CLASS("REGTIME"), 4), 2)
                            sb.Append("<li>")
                            sb.Append("<span style='font-size:12pt;'>").Append("當您收到確認後，").Append("</span>")
                            sb.Append("<span style='color:black;font-weight:bold;font-size:14pt'>").Append("請於開課五日前，按後面連結確定出席，").Append("</span>")
                            Dim sC_URL As String = String.Empty
                            sC_URL = "http://mbscnn.org/MNT/MBMnt_RESP_01_v01.aspx?MB_MEMSEQ=" & e.CommandArgument & "&MB_SEQ=" & ROW_CLASS("MB_SEQ").ToString & "&MB_BATCH=" & ROW_CLASS("MB_BATCH").ToString &
                             "&OPTYPE=Y"
                            sb.Append(Me.getEMail_BT(sC_URL, "確定出席"))
                            sb.Append("<span style='font-weight:bold;font-size:12pt;'>").Append("，若決定不出席請到課程報名網頁點""取消報名""").Append("</span>")
                            sb.Append(Me.getEMail_BT("http://mbscnn.org", "回MBSC首頁"))
                            sb.Append("<span style='font-weight:bold;font-size:12pt;'>").Append("開課前五日內若有變動，請與電話告知聯絡人").Append("</span>")
                            sb.Append("，以利增補候補學員，感謝您。")
                            sb.Append("</li>")

                            sb.Append("<li>").Append("本課程開始日期時間：").Append("<span style='color:black'>").Append(sMB_SDATE).Append("</span>")
                            sb.Append("　報到時間：").Append("<span style='color:black'>").Append(sREGTIME).Append("</span>")
                            sb.Append("</li>")
                            sb.Append("<li>")
                            sb.Append(" 課程地點：").Append("<span style='color:black'>MBSC").Append(ROW_CLASS("MB_PLACE").ToString).Append(" / ")
                            sb.Append(ROW_CLASS("CLASS_PLACE").ToString).Append("<BR/>").Append(ROW_CLASS("TRAFFIC_DESC").ToString)
                            sb.Append("</li>")

                            sb.Append("<li>")
                            sb.Append("<span style='color:black'>聯絡電話：</span>").Append(ROW_CLASS("CONTEL").ToString).Append("　　")
                            sb.Append("<span style='color:black'>聯絡人：").Append(ROW_CLASS("CONTACT").ToString).Append("</span>")
                            sb.Append("</li>")

                            Dim sMB_PREC_MEMO As String = String.Empty
                            sMB_PREC_MEMO = ROW_CLASS("MB_PREC_MEMO").ToString
                            If Utility.isValidateData(sMB_PREC_MEMO) Then
                                sMB_PREC_MEMO = ReplaceToBR(sMB_PREC_MEMO)
                                sb.Append("<li>")
                                sb.Append("<span style='color:black'>注意事項說明：</span>").Append("</span>")
                                sb.Append("<span style='color:black'><BR/>")
                                sb.Append(sMB_PREC_MEMO).Append("</span>")
                                sb.Append("</li>")
                            End If

                            sb.Append("</ol>")


                            sb.Append("<div style='color:black;font-size:12pt;'>")
                            sb.Append("　祝您").Append("<BR/>")
                            sb.Append("　　　學習愉快、收穫滿滿").Append("<BR/>").Append("<BR/>")
                            sb.Append("　　　　　　　　　 MBSC 台北教育中心 敬邀合十")
                            sb.Append("</div>")

                            sMailBody = sb.ToString
                            com.Azion.EloanUtility.NetUtility.GMail_Send(sMailTos, Nothing, sMailSub, sMailBody, True, Nothing, False)
                        End If
                    End If

                    MB_MEMCLASS.setAttribute("MB_FWMK", "1")
                    MB_MEMCLASS.setAttribute("MB_APVDATETIME", Now)
                    MB_MEMCLASS.save()
                End If

                Dim iPage As Decimal = 1
                If IsNumeric(Me.lblCurrentPage.Text) Then
                    iPage = CDec(Me.lblCurrentPage.Text)
                End If
                Me.Bind_TYPE_3(iPage)
            ElseIf e.CommandName = "5N" Then
                Dim LTL_MB_SEQ As Literal = e.Item.FindControl("LTL_MB_SEQ")
                Dim MB_MEMCLASS As New MB_MEMCLASS(Me.m_DBManager)
                If MB_MEMCLASS.LoadByPK(e.CommandArgument, LTL_MB_SEQ.Text) Then
                    MB_MEMCLASS.setAttribute("MB_FWMK", "5")
                    MB_MEMCLASS.setAttribute("MB_APVDATETIME", Now)
                    MB_MEMCLASS.save()
                End If

                Dim iPage As Decimal = 1
                If IsNumeric(Me.lblCurrentPage.Text) Then
                    iPage = CDec(Me.lblCurrentPage.Text)
                End If
                Me.Bind_TYPE_3(iPage)
            ElseIf e.CommandName = "5C" Then
                Dim LTL_MB_SEQ As Literal = e.Item.FindControl("LTL_MB_SEQ")
                Dim MB_MEMCLASS As New MB_MEMCLASS(Me.m_DBManager)
                If MB_MEMCLASS.LoadByPK(e.CommandArgument, LTL_MB_SEQ.Text) Then
                    MB_MEMCLASS.setAttribute("MB_FWMK", "0")
                    MB_MEMCLASS.setAttribute("MB_APVDATETIME", Now)
                    MB_MEMCLASS.save()
                End If

                Dim iPage As Decimal = 1
                If IsNumeric(Me.lblCurrentPage.Text) Then
                    iPage = CDec(Me.lblCurrentPage.Text)
                End If
                Me.Bind_TYPE_3(iPage)
            End If
        Catch ex As System.Threading.ThreadAbortException
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

    Function getEMail_BT(ByVal sURL As String, ByVal sTEXT As String) As String
        Dim sBT As String = String.Empty
        sBT = "<div><!--[if mso]>" &
              "  <v:roundrect xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:w=""urn:schemas-microsoft-com:office:word"" href=""" & sURL &
              """ style=""height:40px;v-text-anchor:middle;width:200px;"" arcsize=""10%"" strokecolor=""#1e3650"" fillcolor=""#5BC0DE""> " &
              "    <w:anchorlock/> " &
              "    <center style=""color:#ffffff;font-family:sans-serif;font-size:20pt;font-weight:bold;"">" & sTEXT & "</center>" &
              "  </v:roundrect> " &
              "<![endif]--><a href=""" & sURL & """style=""background-color:#5BC0DE;border:1px solid #1e3650;border-radius:4px;color:#ffffff;display:inline-block;font-family:sans-serif;font-size:20pt;font-weight:bold;line-height:40px;text-align:center;text-decoration:none;width:200px;-webkit-text-size-adjust:none;mso-hide:all;"">" &
              sTEXT & "</a></div>"

        Return sBT
    End Function

    Function ReplaceToBR(ByVal sInStr As Object) As String
        Try
            If Utility.isValidateData(sInStr) Then
                sInStr = sInStr.Replace(Chr(13) + Chr(10), "<BR/>")
                sInStr = sInStr.Replace(Chr(13), "<BR/>")
                sInStr = sInStr.Replace(Chr(10), "<BR/>")

                Return sInStr
            End If

            Return String.Empty
        Catch ex As Exception
            Throw
        End Try
    End Function

    Private Sub RP_3_ItemDataBound(sender As Object, e As Web.UI.WebControls.RepeaterItemEventArgs) Handles RP_3.ItemDataBound
        Try
            If e.Item.ItemType = ListItemType.Item OrElse e.Item.ItemType = ListItemType.AlternatingItem Then
                Dim DRV As DataRowView = CType(e.Item.DataItem, DataRowView)
                Dim MB_CLASS_NAME As Literal = e.Item.FindControl("MB_CLASS_NAME")
                MB_CLASS_NAME.Text = DRV("MB_CLASS_NAME").ToString

                If IsNumeric(DRV("MB_BATCH")) AndAlso CDec(DRV("MB_BATCH")) > 0 Then
                    MB_CLASS_NAME.Text &= "　第" & DRV("MB_BATCH").ToString & "梯次"
                End If

                '性別
                Dim LTL_MB_SEX As Literal = e.Item.FindControl("LTL_MB_SEX")
                If DRV("MB_SEX").ToString = "1" Then
                    LTL_MB_SEX.Text = "男"
                ElseIf DRV("MB_SEX").ToString = "2" Then
                    LTL_MB_SEX.Text = "女"
                Else
                    LTL_MB_SEX.Text = String.Empty
                End If

                '回信出席
                'Dim LTL_MB_RESP As Literal = e.Item.FindControl("LTL_MB_RESP")
                'If DRV("MB_RESP").ToString = "Y" Then
                '    LTL_MB_RESP.Text = "是"
                'ElseIf DRV("MB_RESP").ToString = "N" Then
                '    LTL_MB_RESP.Text = "否"
                'Else
                '    LTL_MB_RESP.Text = String.Empty
                'End If

                Dim MB_MEMBATCHList As New MB_MEMBATCHList(Me.m_DBManager)
                Dim HID_MB_MEMSEQ As HtmlInputHidden = e.Item.FindControl("HID_MB_MEMSEQ")
                Dim HID_MB_SEQ As HtmlInputHidden = e.Item.FindControl("HID_MB_SEQ")
                MB_MEMBATCHList.LoadBySEQ(HID_MB_MEMSEQ.Value, HID_MB_SEQ.Value)
                '回信出席
                Dim RP_MB_RESP As Repeater = e.Item.FindControl("RP_MB_RESP")
                RP_MB_RESP.DataSource = MB_MEMBATCHList.getCurrentDataSet.Tables(0).Select("ISNULL(MB_ELECT,' ')='1'")
                RP_MB_RESP.DataBind()

                '未出席
                Dim MB_FWMK As CheckBox = e.Item.FindControl("MB_FWMK")
                MB_FWMK.Checked = False
                'If DRV("MB_FWMK").ToString = "1" OrElse Not Utility.isValidateData(DRV("MB_FWMK")) Then
                '    If DRV("MB_RESP").ToString = "N" Then
                '        MB_FWMK.Visible = False
                '    Else
                '        MB_FWMK.Visible = True
                '    End If
                'Else
                '    If DRV("MB_FWMK").ToString = "2" Then
                '        MB_FWMK.Visible = True
                '        MB_FWMK.Checked = True
                '    Else
                '        MB_FWMK.Visible = False
                '    End If
                'End If
                If DRV("MB_APV") = "2" Then
                    '不須核准
                    Dim ROW_MB_RESP() As DataRow = Nothing
                    ROW_MB_RESP = MB_MEMBATCHList.getCurrentDataSet.Tables(0).Select("ISNULL(MB_ELECT,' ')='1' AND ISNULL(MB_RESP, ' ')<>'N'")
                    If DRV("MB_FWMK").ToString = "1" OrElse Not Utility.isValidateData(DRV("MB_FWMK")) Then
                        If Not IsNothing(ROW_MB_RESP) AndAlso ROW_MB_RESP.Length > 0 Then
                            MB_FWMK.Visible = True
                        Else
                            MB_FWMK.Visible = False
                        End If
                    ElseIf DRV("MB_FWMK").ToString = "2" Then
                        MB_FWMK.Checked = True
                        MB_FWMK.Visible = True
                    Else
                        MB_FWMK.Visible = False
                    End If
                Else
                    '須核准
                    If DRV("MB_FWMK").ToString = "3" OrElse DRV("MB_FWMK").ToString = "4" Then
                        MB_FWMK.Visible = False
                    Else
                        Dim ROW_CHKFLAG_1() As DataRow = Nothing
                        ROW_CHKFLAG_1 = MB_MEMBATCHList.getCurrentDataSet.Tables(0).Select("ISNULL(MB_ELECT,' ')='1' AND ISNULL(MB_CHKFLAG,' ')='1' AND ISNULL(MB_RESP, ' ')<>'N'")
                        If Not IsNothing(ROW_CHKFLAG_1) AndAlso ROW_CHKFLAG_1.Length > 0 Then
                            MB_FWMK.Visible = True
                        Else
                            Dim ROW_CHKFLAG_3() As DataRow = Nothing
                            ROW_CHKFLAG_3 = MB_MEMBATCHList.getCurrentDataSet.Tables(0).Select("ISNULL(MB_ELECT,' ')='1' AND ISNULL(MB_CHKFLAG,' ')='3' ")
                            If Not IsNothing(ROW_CHKFLAG_3) AndAlso ROW_CHKFLAG_3.Length > 0 Then
                                MB_FWMK.Visible = True
                                MB_FWMK.Checked = True
                            Else
                                MB_FWMK.Visible = False
                            End If
                        End If
                    End If
                End If

                Dim TD_MB_APV As HtmlTableCell = e.Item.FindControl("TD_MB_APV")
                If DRV("MB_APV").ToString = "1" OrElse DRV("MB_APV").ToString = "3" Then
                    If DRV("MB_FWMK").ToString = "3" Then
                        TD_MB_APV.Visible = False
                    Else
                        TD_MB_APV.Visible = True

                        'Dim btnMB_FWMK_5Y As Button = e.Item.FindControl("btnMB_FWMK_5Y")
                        'Dim btnMB_FWMK_5N As Button = e.Item.FindControl("btnMB_FWMK_5N")
                        'Dim btnMB_FWMK_5C As Button = e.Item.FindControl("btnMB_FWMK_5C")
                        'Dim LTL_MB_APV As Literal = e.Item.FindControl("LTL_MB_APV")
                        'Select Case DRV("MB_FWMK").ToString
                        '    Case "0"
                        '        btnMB_FWMK_5Y.Visible = True
                        '        btnMB_FWMK_5N.Visible = True
                        '    Case "1", "2"
                        '        LTL_MB_APV.Text = "錄取"
                        '        LTL_MB_APV.Visible = True
                        '    Case "3", "4"
                        '        LTL_MB_APV.Text = "取消報名"
                        '        LTL_MB_APV.Visible = True
                        '    Case "5"
                        '        btnMB_FWMK_5C.Visible = True
                        'End Select
                        Dim RP_MB_MEMBATCH As Repeater = e.Item.FindControl("RP_MB_MEMBATCH")
                        RP_MB_MEMBATCH.DataSource = MB_MEMBATCHList.getCurrentDataSet.Tables(0).Select("ISNULL(MB_ELECT,' ')='1'")
                        RP_MB_MEMBATCH.DataBind()
                    End If
                Else
                    TD_MB_APV.Visible = False
                End If
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Sub RP_MB_MEMBATCH_OnItemDataBound(ByVal Sender As Object, ByVal e As RepeaterItemEventArgs)
        If e.Item.ItemType = ListItemType.Item OrElse e.Item.ItemType = ListItemType.AlternatingItem Then
            Dim DRV As DataRow = CType(e.Item.DataItem, DataRow)
            Dim LTL_MB_APV As Literal = e.Item.FindControl("LTL_MB_APV")
            If Not Utility.isValidateData(DRV("MB_CHKFLAG")) Then
                Dim btnMB_FWMK_5Y As Button = e.Item.FindControl("btnMB_FWMK_5Y")
                btnMB_FWMK_5Y.Visible = True
                Dim btnMB_FWMK_5N As Button = e.Item.FindControl("btnMB_FWMK_5N")
                btnMB_FWMK_5N.Visible = True
            ElseIf DRV("MB_CHKFLAG").ToString = "1" Then
                LTL_MB_APV.Text = String.Empty
                LTL_MB_APV.Visible = True
                Dim LTL_MB_BATCH As Label = e.Item.FindControl("LTL_MB_BATCH")
                LTL_MB_BATCH.Visible = False
            ElseIf DRV("MB_CHKFLAG").ToString = "2" Then
                Dim btnMB_FWMK_5C As Button = e.Item.FindControl("btnMB_FWMK_5C")
                btnMB_FWMK_5C.Visible = True
            End If
        End If
    End Sub

    Sub RP_MB_MEMBATCH_OnItemCommand(ByVal Sender As Object, ByVal e As RepeaterCommandEventArgs)
        Try
            Dim LTL_MB_BATCH As Label = e.Item.FindControl("LTL_MB_BATCH")
            Dim HID_MB_MEMSEQ As HtmlInputHidden = e.Item.FindControl("HID_MB_MEMSEQ")
            Dim HID_MB_SEQ As HtmlInputHidden = e.Item.FindControl("HID_MB_SEQ")

            If e.CommandName = "5Y" Then
                Dim MB_MEMCLASS As New MB_MEMCLASS(Me.m_DBManager)
                If MB_MEMCLASS.LoadByPK(HID_MB_MEMSEQ.Value, HID_MB_SEQ.Value) Then
                    'mail 錄取通知
                    Dim mbMEMBER As New MB_MEMBER(Me.m_DBManager)
                    If mbMEMBER.loadByPK(e.CommandArgument) AndAlso Utility.isValidateData(mbMEMBER.getAttribute("MB_EMAIL")) Then
                        'Dim MB_CLASSList As New MB_CLASSList(Me.m_DBManager)
                        'MB_CLASSList.LoadByMB_SEQ(HID_MB_SEQ.Value)
                        Dim MB_CLASS As New MB_CLASS(Me.m_DBManager)
                        If MB_CLASS.loadByPK(HID_MB_SEQ.Value, LTL_MB_BATCH.Text) Then
                            'Dim ROW_CLASS As DataRow = MB_CLASSList.getCurrentDataSet.Tables(0).Rows(0)

                            Dim sMailTos() As String = {Trim(mbMEMBER.getString("MB_EMAIL"))}

                            Dim sMailSub As String = String.Empty
                            Dim sBATCH As String = String.Empty
                            If IsNumeric(LTL_MB_BATCH.Text) AndAlso CDec(LTL_MB_BATCH.Text) > 0 Then
                                sBATCH = "第" & LTL_MB_BATCH.Text & "梯次"
                            End If
                            sMailSub = MB_CLASS.getString("MB_CLASS_NAME") & sBATCH & " 錄取通知 請務必回覆是否出席"

                            Dim sMailBody As String = String.Empty
                            Dim sb As New System.Text.StringBuilder
                            sb.Append("<DIV style='text-align:center;font-size:24pt;color:red' >").Append(MB_CLASS.getString("MB_CLASS_NAME")).Append(sBATCH).Append("</DIV>")
                            Dim sForm As String = String.Empty
                            'If isWait Then
                            '    sForm = "備取通知"
                            'Else
                            '    sForm = "錄取通知"
                            'End If
                            sForm = "錄取通知"
                            sb.Append("<DIV style='text-align:center;font-size:24pt;color:black;font-weight:bold;text-decoration:underline' >").Append(sForm).Append("</DIV>")
                            sb.Append("<DIV style='font-size:12pt;color:black'>").Append("親愛的").Append(mbMEMBER.getString("MB_NAME")).Append("吉祥：").Append("</DIV>")

                            sb.Append("<DIV style='font-size:12pt;color:black;'>").Append("　　  歡迎您報名參加")
                            sb.Append("<span style='color:black'>")
                            sb.Append(MB_CLASS.getString("MB_CLASS_NAME"))
                            sb.Append("</span>")
                            sb.Append("！此通知函，乃透過系統發送。為使您在課程期間能順利進行，並獲得最大收穫，請務必閱讀下列注意事項：").Append("</DIV>")
                            sb.Append("<ol type='1' style='font-size:12pt;color:black;' >")
                            Dim sMB_SDATE As String = String.Empty
                            If Utility.isValidateData(MB_CLASS.getAttribute("MB_SDATE")) Then
                                sMB_SDATE = CDate(MB_CLASS.getAttribute("MB_SDATE").ToString).Year & "年" & CDate(MB_CLASS.getAttribute("MB_SDATE").ToString).Month & "月" &
                                    CDate(MB_CLASS.getAttribute("MB_SDATE").ToString).Day & "日"
                            End If
                            Dim sREGTIME As String = String.Empty
                            sREGTIME = Left(Utility.FillZero(MB_CLASS.getString("REGTIME"), 4), 2) & Right(Utility.FillZero(MB_CLASS.getString("REGTIME"), 4), 2)
                            sb.Append("<li>")
                            sb.Append("<span style='font-size:12pt;'>").Append("當您收到確認後，").Append("</span>")
                            sb.Append("<span style='color:black;font-weight:bold;font-size:14pt'>").Append("請於開課五日前，按後面連結確定出席，").Append("</span>")
                            Dim sC_URL As String = String.Empty
                            sC_URL = "http://mbscnn.org/MNT/MBMnt_RESP_01_v01.aspx?MB_MEMSEQ=" & e.CommandArgument & "&MB_SEQ=" & MB_CLASS.getString("MB_SEQ") & "&MB_BATCH=" & MB_CLASS.getString("MB_BATCH") &
                             "&OPTYPE=Y"
                            sb.Append(Me.getEMail_BT(sC_URL, "確定出席"))
                            sb.Append("<span style='font-weight:bold;font-size:12pt;'>").Append("，若決定不出席請到課程報名網頁點""取消報名""").Append("</span>")
                            sb.Append(Me.getEMail_BT("http://mbscnn.org", "回MBSC首頁"))
                            sb.Append("<span style='font-weight:bold;font-size:12pt;'>").Append("開課前五日內若有變動，請與電話告知聯絡人").Append("</span>")
                            sb.Append("，以利增補候補學員，感謝您。")
                            sb.Append("</li>")

                            sb.Append("<li>").Append("本課程開始日期時間：").Append("<span style='color:black'>").Append(sMB_SDATE).Append("</span>")
                            sb.Append("　報到時間：").Append("<span style='color:black'>").Append(sREGTIME).Append("</span>")
                            sb.Append("</li>")
                            sb.Append("<li>")
                            sb.Append(" 課程地點：").Append("<span style='color:black'>MBSC").Append(MB_CLASS.getString("MB_PLACE")).Append(" / ")
                            sb.Append(MB_CLASS.getString("CLASS_PLACE")).Append("<BR/>").Append(MB_CLASS.getString("TRAFFIC_DESC"))
                            sb.Append("</li>")

                            sb.Append("<li>")
                            sb.Append("<span style='color:black'>聯絡電話：</span>").Append(MB_CLASS.getString("CONTEL")).Append("　　")
                            sb.Append("<span style='color:black'>聯絡人：").Append(MB_CLASS.getString("CONTACT")).Append("</span>")
                            sb.Append("</li>")

                            Dim sMB_PREC_MEMO As String = String.Empty
                            sMB_PREC_MEMO = MB_CLASS.getString("MB_PREC_MEMO")
                            If Utility.isValidateData(sMB_PREC_MEMO) Then
                                sMB_PREC_MEMO = ReplaceToBR(sMB_PREC_MEMO)
                                sb.Append("<li>")
                                sb.Append("<span style='color:black'>注意事項說明：</span>").Append("</span>")
                                sb.Append("<span style='color:black'><BR/>")
                                sb.Append(sMB_PREC_MEMO).Append("</span>")
                                sb.Append("</li>")
                            End If

                            sb.Append("</ol>")


                            sb.Append("<div style='color:black;font-size:12pt;'>")
                            sb.Append("　祝您").Append("<BR/>")
                            sb.Append("　　　學習愉快、收穫滿滿").Append("<BR/>").Append("<BR/>")
                            sb.Append("　　　　　　　　　 MBSC 台北教育中心 敬邀合十")
                            sb.Append("</div>")

                            sMailBody = sb.ToString
                            com.Azion.EloanUtility.NetUtility.GMail_Send(sMailTos, Nothing, sMailSub, sMailBody, True, Nothing, False)
                        End If
                    End If

                    Dim MB_MEMBATCH As New MB_MEMBATCH(Me.m_DBManager)
                    If MB_MEMBATCH.LoadByPK(HID_MB_MEMSEQ.Value, HID_MB_SEQ.Value, LTL_MB_BATCH.Text) Then
                        MB_MEMBATCH.setAttribute("MB_CHKFLAG", "1")
                        MB_MEMBATCH.setAttribute("MB_APVDATE", Now)
                        MB_MEMBATCH.save()
                    End If
                End If

                Dim iPage As Decimal = 1
                If IsNumeric(Me.lblCurrentPage.Text) Then
                    iPage = CDec(Me.lblCurrentPage.Text)
                End If
                Me.Bind_TYPE_3(iPage)
            ElseIf e.CommandName = "5N" Then
                Dim MB_MEMBATCH As New MB_MEMBATCH(Me.m_DBManager)
                If MB_MEMBATCH.LoadByPK(HID_MB_MEMSEQ.Value, HID_MB_SEQ.Value, LTL_MB_BATCH.Text) Then
                    MB_MEMBATCH.setAttribute("MB_CHKFLAG", "2")
                    MB_MEMBATCH.setAttribute("MB_APVDATE", Now)
                    MB_MEMBATCH.save()
                End If

                Dim iPage As Decimal = 1
                If IsNumeric(Me.lblCurrentPage.Text) Then
                    iPage = CDec(Me.lblCurrentPage.Text)
                End If
                Me.Bind_TYPE_3(iPage)
            ElseIf e.CommandName = "5C" Then
                Dim MB_MEMBATCH As New MB_MEMBATCH(Me.m_DBManager)
                If MB_MEMBATCH.LoadByPK(HID_MB_MEMSEQ.Value, HID_MB_SEQ.Value, LTL_MB_BATCH.Text) Then
                    MB_MEMBATCH.setAttribute("MB_CHKFLAG", Nothing)
                    MB_MEMBATCH.setAttribute("MB_APVDATE", Nothing)
                    MB_MEMBATCH.save()
                End If

                Dim iPage As Decimal = 1
                If IsNumeric(Me.lblCurrentPage.Text) Then
                    iPage = CDec(Me.lblCurrentPage.Text)
                End If
                Me.Bind_TYPE_3(iPage)
            End If
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

    Sub RP_MB_RESP_OnItemDataBound(ByVal Sender As Object, ByVal e As RepeaterItemEventArgs)
        If e.Item.ItemType = ListItemType.Item OrElse e.Item.ItemType = ListItemType.AlternatingItem Then
            Dim DRV As DataRow = CType(e.Item.DataItem, DataRow)
            Dim LTL_MB_RESP As Literal = e.Item.FindControl("LTL_MB_RESP")
            If DRV("MB_RESP").ToString = "Y" Then
                LTL_MB_RESP.Text = "是"
            ElseIf DRV("MB_RESP").ToString = "N" Then
                LTL_MB_RESP.Text = "否"
            Else
                LTL_MB_RESP.Text = String.Empty
                Dim LTL_MB_BATCH As Label = e.Item.FindControl("LTL_MB_BATCH")
                LTL_MB_BATCH.Visible = False
            End If
        End If
    End Sub

    Private Sub btnMB_FWMK_Click(sender As Object, e As EventArgs) Handles btnMB_FWMK.Click
        Try
            Try
                Me.m_DBManager.beginTran()
                For i As Integer = 0 To Me.RP_3.Items.Count - 1
                    Dim MB_FWMK As CheckBox = Me.RP_3.Items(i).FindControl("MB_FWMK")
                    If MB_FWMK.Visible = True AndAlso MB_FWMK.Checked Then
                        Dim HID_MB_BATCH As HtmlInputHidden = Me.RP_3.Items(i).FindControl("HID_MB_BATCH")
                        Dim HID_MB_MEMSEQ As HtmlInputHidden = Me.RP_3.Items(i).FindControl("HID_MB_MEMSEQ")
                        Dim HID_MB_SEQ As HtmlInputHidden = Me.RP_3.Items(i).FindControl("HID_MB_SEQ")
                        Dim HID_MB_APV As HtmlInputHidden = Me.RP_3.Items(i).FindControl("HID_MB_APV")
                        If HID_MB_APV.Value = "2" Then
                            Dim MB_MEMCLASS As New MB_MEMCLASS(Me.m_DBManager)
                            If MB_MEMCLASS.LoadByPK(HID_MB_MEMSEQ.Value, HID_MB_SEQ.Value) Then
                                MB_MEMCLASS.setAttribute("MB_FWMK", "2")
                                MB_MEMCLASS.save()
                            End If
                        Else
                            Dim MB_MEMBATCHList As New MB_MEMBATCHList(Me.m_DBManager)
                            MB_MEMBATCHList.LoadBySEQ(HID_MB_MEMSEQ.Value, HID_MB_SEQ.Value)
                            Dim ROW_1() As DataRow = MB_MEMBATCHList.getCurrentDataSet.Tables(0).Select("ISNULL(MB_ELECT, ' ')='1'")
                            If Not IsNothing(ROW_1) AndAlso ROW_1.Length > 0 Then
                                For Each ROW As DataRow In ROW_1
                                    Dim MB_MEMBATCH As New MB_MEMBATCH(Me.m_DBManager)
                                    If MB_MEMBATCH.LoadByPK(ROW("MB_MEMSEQ"), ROW("MB_SEQ"), ROW("MB_BATCH")) Then
                                        MB_MEMBATCH.setAttribute("MB_CHKFLAG", "3")
                                        MB_MEMBATCH.save()
                                    End If
                                Next
                            End If
                        End If
                    End If
                Next
                Me.m_DBManager.commit()
            Catch ex As Exception
                Me.m_DBManager.Rollback()
                Throw
            End Try

            Me.Bind_TYPE_3(1)
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub
End Class