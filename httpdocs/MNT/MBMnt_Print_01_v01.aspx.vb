Imports MBSC.MB_OP
Imports com.Azion.NET.VB
Imports MBSC.UICtl

Public Class MBMnt_Print_01_v01
    Inherits System.Web.UI.Page

    Dim m_DBManager As com.Azion.NET.VB.DatabaseManager

    Dim m_sUpCode7 As String = String.Empty
    Dim DT_UPCODE_7 As DataTable

    Dim m_sUpcode15 As String = String.Empty
    Dim DT_UPCODE_15 As DataTable

    Dim m_sUpCode28 As String = String.Empty
    Dim DT_UPCODE_28 As DataTable

    Dim m_sUpcode23 As String = String.Empty
    Dim DT_UPCODE_23 As DataTable

    Dim m_sUpcode31 As String = String.Empty
    Dim DT_UPCODE_31 As DataTable

#Region "Page Events"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            '所屬區代碼
            Me.m_sUpCode7 = com.Azion.EloanUtility.CodeList.getAppSettings("UpCode7")

            '功德項目
            Me.m_sUpcode15 = com.Azion.EloanUtility.CodeList.getAppSettings("UpCode15")

            '會員類別
            Me.m_sUpCode28 = com.Azion.EloanUtility.CodeList.getAppSettings("UpCode28")

            '繳款方式
            Me.m_sUpcode23 = com.Azion.EloanUtility.CodeList.getAppSettings("UpCode23")

            '付款方式
            Me.m_sUpcode31 = com.Azion.EloanUtility.CodeList.getAppSettings("UpCode31")

            Me.m_DBManager = UIShareFun.getDataBaseManager

            If Not Page.IsPostBack Then
                Me.Bind_MB_AREA()
            End If
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    Private Sub MBMnt_Print_01_v01_Unload(sender As Object, e As System.EventArgs) Handles Me.Unload
        UIShareFun.releaseConnection(Me.m_DBManager)
    End Sub
#End Region

#Region "DDL"
    Sub Bind_MB_AREA()
        Dim DT_MB_AREA As DataTable = Nothing
        Try
            Me.MB_AREA_SUB_2.Items.Clear()

            Dim apCODEList As New AP_CODEList(Me.m_DBManager)
            apCODEList.loadByUpCode(m_sUpCode7)
            DT_MB_AREA = apCODEList.getCurrentDataSet.Tables(0)

            If Utility.isValidateData(Session("AREA")) Then
                Dim ROW_OTH_AREA() As DataRow = Nothing
                ROW_OTH_AREA = DT_MB_AREA.Select("VALUE<>'" & Session("AREA") & "'")
                If Not IsNothing(ROW_OTH_AREA) AndAlso ROW_OTH_AREA.Length > 0 Then
                    For i As Integer = 0 To UBound(ROW_OTH_AREA)
                        DT_MB_AREA.Rows.Remove(ROW_OTH_AREA(i))
                    Next
                End If
            End If

            Me.MB_AREA_SUB_2.DataSource = DT_MB_AREA
            Me.MB_AREA_SUB_2.DataTextField = "TEXT"
            Me.MB_AREA_SUB_2.DataValueField = "VALUE"
            Me.MB_AREA_SUB_2.DataBind()

            If Me.MB_AREA_SUB_2.Items.Count > 1 Then
                Me.MB_AREA_SUB_2.Items.Insert(0, New ListItem("請選擇", ""))
            End If
        Catch ex As Exception
            Throw
        Finally
            If Not IsNothing(DT_MB_AREA) Then
                DT_MB_AREA.Dispose()
            End If
        End Try
    End Sub
#End Region

#Region "Button Events"
    Sub bt_Qry_SUB_2_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles bt_Qry_SUB_2.Click
        Try
            If Not Utility.isValidateData(Me.RBL_Choose.SelectedValue) Then
                com.Azion.EloanUtility.UIUtility.alert("請先選擇操作項目!\n\n(列印/回收/補印)")
                Return
            End If

            If Me.RP_QRY_SUB_2_1.Checked = False AndAlso RP_QRY_SUB_2_2.Checked = False AndAlso Me.RP_QRY_SUB_2_3.Checked = False Then
                com.Azion.EloanUtility.UIUtility.alert("請選擇查詢方式!\n\n(所屬區/繳款年月/姓名)")
                Return
            End If

            If Me.RP_QRY_SUB_2_1.Checked AndAlso Not Utility.isValidateData(Me.MB_AREA_SUB_2.SelectedValue) Then
                com.Azion.EloanUtility.UIUtility.alert("請選擇所屬區!")
                Return
            ElseIf Me.RP_QRY_SUB_2_2.Checked Then
                If Not IsNumeric(Me.txt_QRY_SUB_2_BEGYYY.Text) Then
                    com.Azion.EloanUtility.UIUtility.alert("請輸入繳款起年!")

                    Return
                End If

                If Not IsNumeric(Me.txt_QRY_SUB_2_BEGMM.Text) Then
                    com.Azion.EloanUtility.UIUtility.alert("請輸入繳款起月!")

                    Return
                End If

                If Not IsDate(CDec(Me.txt_QRY_SUB_2_BEGYYY.Text) & "/" & Me.txt_QRY_SUB_2_BEGMM.Text & "/1") Then
                    com.Azion.EloanUtility.UIUtility.alert("繳款年月起日輸入錯誤!")

                    Return
                End If

                Dim D_BEG_DATE As Date = New Date(CDec(Me.txt_QRY_SUB_2_BEGYYY.Text), CDec(Me.txt_QRY_SUB_2_BEGMM.Text), 1)
                If D_BEG_DATE > Now Then
                    com.Azion.EloanUtility.UIUtility.alert("繳款年月起日不可大於系統日!")

                    Return
                End If

                If Not IsNumeric(Me.txt_QRY_SUB_2_ENDYYY.Text) Then
                    com.Azion.EloanUtility.UIUtility.alert("請輸入繳款迄年!")

                    Return
                End If

                If Not IsNumeric(Me.txt_QRY_SUB_2_ENDMM.Text) Then
                    com.Azion.EloanUtility.UIUtility.alert("請輸入繳款迄月!")

                    Return
                End If

                If Not IsDate(CDec(Me.txt_QRY_SUB_2_ENDYYY.Text) & "/" & Me.txt_QRY_SUB_2_ENDMM.Text & "/1") Then
                    com.Azion.EloanUtility.UIUtility.alert("繳款年月迄日輸入錯誤!")

                    Return
                End If

                Dim D_END_DATE As Date = New Date(CDec(Me.txt_QRY_SUB_2_ENDYYY.Text), CDec(Me.txt_QRY_SUB_2_ENDMM.Text), 1)
                If D_END_DATE > Now Then
                    com.Azion.EloanUtility.UIUtility.alert("繳款年月迄日不可大於系統日!")

                    Return
                End If

                Dim D_BEG As Date = New Date(CDec(Me.txt_QRY_SUB_2_BEGYYY.Text), CDec(Me.txt_QRY_SUB_2_BEGMM.Text), 1)
                Dim D_END As Date = New Date(CDec(Me.txt_QRY_SUB_2_ENDYYY.Text), CDec(Me.txt_QRY_SUB_2_ENDMM.Text), 1)
                If D_END < D_BEG Then
                    com.Azion.EloanUtility.UIUtility.alert("繳款年月迄日不可小於繳款年月起日!")

                    Return
                End If
            ElseIf Me.RP_QRY_SUB_2_3.Checked AndAlso Not Utility.isValidateData(Trim(Me.txt_QRY_MB_NAME.Text)) Then
                com.Azion.EloanUtility.UIUtility.alert("請輸入姓名!")
                Return
            End If

            Me.RP_MB_MEMREV.DataSource = Nothing
            Me.RP_MB_MEMREV.DataBind()

            Dim mbMEMREVList As New MB_MEMREVList(Me.m_DBManager)
            Select Case Me.RBL_Choose.SelectedValue
                Case "0"
                    '列印收據
                    If Me.RP_QRY_SUB_2_1.Checked Then
                        mbMEMREVList.LoadByMB_AREAForPrint(Me.MB_AREA_SUB_2.SelectedValue)
                        If mbMEMREVList.getCurrentDataSet.Tables(0).Rows.Count > 0 Then
                            Me.RP_MB_MEMREV.DataSource = mbMEMREVList.getCurrentDataSet.Tables(0)
                            Me.RP_MB_MEMREV.DataBind()

                            Me.PLH_QRY.Visible = False
                            Me.PLH_MB_MEMREV_LIST.Visible = True
                        Else
                            com.Azion.EloanUtility.UIUtility.alert("該區無未列印資料")
                            Return
                        End If
                    ElseIf Me.RP_QRY_SUB_2_2.Checked Then
                        Dim iBEG As Decimal = (CDec(Me.txt_QRY_SUB_2_BEGYYY.Text)) * 100 + CDec(Me.txt_QRY_SUB_2_BEGMM.Text)
                        Dim iEND As Decimal = (CDec(Me.txt_QRY_SUB_2_ENDYYY.Text)) * 100 + CDec(Me.txt_QRY_SUB_2_ENDMM.Text)

                        mbMEMREVList.LoadByMB_TX_DATEForPrint(iBEG, iEND, Session("AREA"))
                        If mbMEMREVList.getCurrentDataSet.Tables(0).Rows.Count > 0 Then
                            Me.RP_MB_MEMREV.DataSource = mbMEMREVList.getCurrentDataSet.Tables(0)
                            Me.RP_MB_MEMREV.DataBind()

                            Me.PLH_QRY.Visible = False
                            Me.PLH_MB_MEMREV_LIST.Visible = True
                        Else
                            com.Azion.EloanUtility.UIUtility.alert("該繳款年月無未列印資料")
                            Return
                        End If
                    ElseIf Me.RP_QRY_SUB_2_3.Checked Then
                        mbMEMREVList.LoadByMB_NAMEForPrint(Trim(Me.txt_QRY_MB_NAME.Text), Session("AREA"))
                        If mbMEMREVList.getCurrentDataSet.Tables(0).Rows.Count > 0 Then
                            Me.RP_MB_MEMREV.DataSource = mbMEMREVList.getCurrentDataSet.Tables(0)
                            Me.RP_MB_MEMREV.DataBind()

                            Me.PLH_QRY.Visible = False
                            Me.PLH_MB_MEMREV_LIST.Visible = True
                        Else
                            com.Azion.EloanUtility.UIUtility.alert("無未列印資料")
                            Return
                        End If
                    End If
                Case "1", "2"
                    '回收收據
                    '補印收據
                    If Me.RP_QRY_SUB_2_1.Checked Then
                        mbMEMREVList.LoadByMB_AREAForAPPEND(Me.MB_AREA_SUB_2.SelectedValue)
                        If mbMEMREVList.getCurrentDataSet.Tables(0).Rows.Count > 0 Then
                            Me.RP_MB_MEMREV.DataSource = mbMEMREVList.getCurrentDataSet.Tables(0)
                            Me.RP_MB_MEMREV.DataBind()

                            Me.PLH_QRY.Visible = False
                            Me.PLH_MB_MEMREV_LIST.Visible = True
                        Else
                            com.Azion.EloanUtility.UIUtility.alert("該區無資料")
                            Return
                        End If
                    ElseIf Me.RP_QRY_SUB_2_2.Checked Then
                        Dim iBEG As Decimal = (CDec(Me.txt_QRY_SUB_2_BEGYYY.Text)) * 100 + CDec(Me.txt_QRY_SUB_2_BEGMM.Text)
                        Dim iEND As Decimal = (CDec(Me.txt_QRY_SUB_2_ENDYYY.Text)) * 100 + CDec(Me.txt_QRY_SUB_2_ENDMM.Text)

                        mbMEMREVList.LoadByMB_TX_DATEForAPPEND(iBEG, iEND, Session("AREA"))
                        If mbMEMREVList.getCurrentDataSet.Tables(0).Rows.Count > 0 Then
                            Me.RP_MB_MEMREV.DataSource = mbMEMREVList.getCurrentDataSet.Tables(0)
                            Me.RP_MB_MEMREV.DataBind()

                            Me.PLH_QRY.Visible = False
                            Me.PLH_MB_MEMREV_LIST.Visible = True
                        Else
                            com.Azion.EloanUtility.UIUtility.alert("該繳款年月無資料")
                            Return
                        End If
                    ElseIf Me.RP_QRY_SUB_2_3.Checked Then
                        mbMEMREVList.LoadByMB_NAMEForAPPEND(Trim(Me.txt_QRY_MB_NAME.Text), Session("AREA"))
                        If mbMEMREVList.getCurrentDataSet.Tables(0).Rows.Count > 0 Then
                            Me.RP_MB_MEMREV.DataSource = mbMEMREVList.getCurrentDataSet.Tables(0)
                            Me.RP_MB_MEMREV.DataBind()

                            Me.PLH_QRY.Visible = False
                            Me.PLH_MB_MEMREV_LIST.Visible = True
                        Else
                            com.Azion.EloanUtility.UIUtility.alert("無資料")
                            Return
                        End If
                    End If
            End Select
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    Sub bt_Back_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles bt_Back.Click
        Try
            Me.PLH_QRY.Visible = True

            Me.PLH_MB_MEMREV_LIST.Visible = False

            Me.RP_MB_MEMREV.DataSource = Nothing
            Me.RP_MB_MEMREV.DataBind()
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    Sub bt_Cancel_Data_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles bt_Cancel_Data.Click
        Try
            Me.ClearPLH_MB_MEMREV()

            Me.PLH_MB_MEMREV.Visible = False

            Me.PLH_MB_MEMREV_LIST.Visible = True
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    '列印收據
    Sub bt_Print_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles bt_Print.Click
        Try
            Try
                Me.m_DBManager.beginTran()

                '傳票號碼???
                Dim iMB_YYYY As Decimal = 0
                iMB_YYYY = Now.Year

                Dim sProcName As String = String.Empty
                sProcName = LCase(com.Azion.NET.VB.Properties.getSchemaName) & ".MB_GET_MB_GETNO"
                Dim inParaAL As New ArrayList
                Dim outParaAL As New ArrayList
                inParaAL.Add(iMB_YYYY)
                inParaAL.Add("RE")

                outParaAL.Add(6)

                Dim HT_Return As Hashtable = com.Azion.NET.VB.DBObject.getProcedureData(Me.m_DBManager, sProcName, inParaAL, outParaAL)
                Dim iMB_VOUCHER_N As Decimal = HT_Return.Item("@IMAXID")

                Dim mbMEMREV As New MB_MEMREV(Me.m_DBManager)
                If mbMEMREV.loadByPK(CDec(Me.HID_MB_MEMSEQ.Value), CDec(Me.HID_MB_SEQNO.Value)) Then
                    '列印收據 Y:已經列印
                    mbMEMREV.setAttribute("MB_PRINT", "Y")

                    '收據號碼-西元年
                    mbMEMREV.setAttribute("MB_RECV_Y", iMB_YYYY)

                    '收據號碼
                    mbMEMREV.setAttribute("MB_RECV_N", iMB_VOUCHER_N)

                    mbMEMREV.save()
                End If

                Me.m_DBManager.commit()
            Catch ex As Exception
                Me.m_DBManager.Rollback()
                Throw
            End Try

            Dim sURL As String = String.Empty
            Dim sPrintAspx As String = String.Empty
            'sPrintAspx = "MBMnt_Print_01_01_v01.aspx"
            sPrintAspx = "MBMnt_Print_01_01_v02.aspx"
            sURL = com.Azion.EloanUtility.UIUtility.getRootPath & "/MNT/" & sPrintAspx & "?MB_MEMSEQ=" & Me.HID_MB_MEMSEQ.Value & _
                   "&MB_SEQNO=" & Me.HID_MB_SEQNO.Value

            com.Azion.EloanUtility.UIUtility.showOpen(sURL)

            Me.PLH_MB_MEMREV.Visible = False

            Me.ClearPLH_MB_MEMREV()

            Me.bt_Qry_SUB_2_Click(Sender, e)
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    '回收收據
    Sub bt_Return_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles bt_Return.Click
        Try
            Dim mbMEMREV As New MB_MEMREV(Me.m_DBManager)
            If mbMEMREV.loadByPK(CDec(Me.HID_MB_MEMSEQ.Value), CDec(Me.HID_MB_SEQNO.Value)) Then
                '回收收據 Y:回收收據
                mbMEMREV.setAttribute("MB_REBREV", "Y")

                mbMEMREV.save()
            End If

            Me.PLH_MB_MEMREV.Visible = False

            Me.ClearPLH_MB_MEMREV()

            Me.bt_Qry_SUB_2_Click(Sender, e)
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    '補印收據
    Sub bt_Append_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles bt_Append.Click
        Try
            Dim mbMEMREV As New MB_MEMREV(Me.m_DBManager)
            If mbMEMREV.loadByPK(CDec(Me.HID_MB_MEMSEQ.Value), CDec(Me.HID_MB_SEQNO.Value)) Then
                '收據補發次數
                mbMEMREV.setAttribute("MB_REBREV", mbMEMREV.getDecimal("MB_REISU") + 1)

                mbMEMREV.save()
            End If

            Dim sURL As String = String.Empty
            sURL = com.Azion.EloanUtility.UIUtility.getRootPath & "/MNT/MBMnt_Print_01_01_v01.aspx?MB_MEMSEQ=" & Me.HID_MB_MEMSEQ.Value & _
                   "&MB_SEQNO=" & Me.HID_MB_SEQNO.Value

            com.Azion.EloanUtility.UIUtility.showOpen(sURL)

            Me.PLH_MB_MEMREV.Visible = False

            Me.ClearPLH_MB_MEMREV()

            Me.bt_Qry_SUB_2_Click(Sender, e)
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub
#End Region

#Region "RP_MB_MEMREV"
    Sub RP_MB_MEMREV_ItemCommand(ByVal Sender As Object, ByVal e As RepeaterCommandEventArgs) Handles RP_MB_MEMREV.ItemCommand
        Try
            Dim iMB_MEMSEQ As Decimal = e.CommandArgument

            Dim HID_MB_SEQNO As HtmlInputHidden = e.Item.FindControl("HID_MB_SEQNO")
            Dim iMB_SEQNO As Decimal = CDec(HID_MB_SEQNO.Value)

            Dim mbMEMBER As New MB_MEMBER(Me.m_DBManager)
            mbMEMBER.loadByPK(iMB_MEMSEQ)

            Dim mbMEMREV As New MB_MEMREV(Me.m_DBManager)
            mbMEMREV.loadByPK(iMB_MEMSEQ, iMB_SEQNO)

            '法名/姓名
            Me.MB_NAME.Text = Trim(mbMEMBER.getString("MB_NAME"))

            '會員編號
            Me.MB_MEMSEQ.Text = Me.getMB_MEMSEQ(iMB_MEMSEQ, mbMEMBER.getString("MB_AREA"))
            Me.HID_MB_MEMSEQ.Value = iMB_MEMSEQ
            Me.HID_MB_SEQNO.Value = iMB_SEQNO

            '通訊地址
            Me.MB_ADDR.Text = mbMEMBER.getString("MB_CITY") & mbMEMBER.getString("MB_VLG") & mbMEMBER.getString("MB_ADDR")

            '繳款日期
            If Utility.isDecimalDate(mbMEMREV.getAttribute("MB_TX_DATE")) Then
                Me.MB_TX_DATE.Text = Me.get_DECIMAL_DATE(mbMEMREV.getAttribute("MB_TX_DATE"))
            Else
                Me.MB_TX_DATE.Text = String.Empty
            End If

            '功德項目
            Me.MB_ITEMID.Text = Me.get_MB_ITEMID_TEXT(mbMEMREV.getString("MB_ITEMID"))
            Me.HID_MB_ITEMID.Value = mbMEMREV.getString("MB_ITEMID")
            If mbMEMREV.getString("MB_ITEMID") = "A" Then
                Me.PLH_MB_MEMTYP.Visible = True
                Me.MB_MEMTYP.Text = Me.get_MB_MEMTYP_TEXT(mbMEMREV.getString("MB_MEMTYP"))
            Else
                Me.PLH_MB_MEMTYP.Visible = False
                Me.MB_MEMTYP.Text = String.Empty
            End If

            '繳款方式
            Me.MB_FEETYPE.Text = Me.get_MB_FEETYPE_TEXT(mbMEMREV.getString("MB_FEETYPE"))

            '繳款金額
            If IsNumeric(mbMEMREV.getAttribute("MB_TOTFEE")) Then
                Me.MB_TOTFEE.Text = Utility.FormatDec(mbMEMREV.getAttribute("MB_TOTFEE"), "#,##0")
            Else
                Me.MB_TOTFEE.Text = String.Empty
            End If
            '每月金額
            If mbMEMREV.getString("MB_FEETYPE") <> "D" Then
                Me.PLH_MB_TOTFEE_MM.Visible = True
                Dim sNOTE As String = String.Empty
                sNOTE = Me.get_MB_FEETYPE_NOTE(mbMEMREV.getString("MB_FEETYPE"))
                If IsNumeric(sNOTE) Then
                    Me.MB_TOTFEE_MM.Text = Utility.FormatDec(sNOTE, "#,##0")
                Else
                    Me.MB_TOTFEE_MM.Text = String.Empty
                End If
            Else
                Me.PLH_MB_TOTFEE_MM.Visible = False
                Me.MB_TOTFEE_MM.Text = String.Empty
            End If

            '收據捐款名稱
            Me.MB_RECNAME.Text = mbMEMREV.getString("MB_RECNAME")

            '所屬區
            Me.MB_AREA.Text = Me.getMB_AREA(mbMEMBER.getString("MB_AREA"))

            '所屬區委員
            Me.MB_LEADER.Text = mbMEMBER.getString("MB_LEADER")

            '會費期間
            If mbMEMREV.getString("MB_ITEMID") = "A" Then
                Me.MB_DESC.Visible = False
                Me.MB_DESC.Text = String.Empty

                Me.MB_MEMFEE.Visible = True
                Me.MB_MEMFEE.Text = (mbMEMREV.getDecimal("MB_MEMFEE_SY")) & "年" & mbMEMREV.getString("MB_MEMFEE_SM") & "月至" & _
                                    (mbMEMREV.getDecimal("MB_MEMFEE_EY")) & "年" & mbMEMREV.getString("MB_MEMFEE_EM") & "月"

                Me.HID_MB_MEMFEE_SY.Value = mbMEMREV.getDecimal("MB_MEMFEE_SY")
            Else
                Me.MB_MEMFEE.Visible = False
                Me.MB_MEMFEE.Text = String.Empty

                Me.MB_DESC.Visible = True
                Me.MB_DESC.Text = mbMEMREV.getString("MB_DESC")
            End If

            '付款方式
            Me.MB_PAY_TYPE.Text = Me.get_MB_PAY_TYPE_TEXT(mbMEMREV.getString("MB_PAY_TYPE"))

            '票據到期日
            If Utility.isDecimalDate(mbMEMREV.getAttribute("NOTE_DUE_DATE")) Then
                Me.NOTE_DUE_DATE.Text = Me.get_DECIMAL_DATE(mbMEMREV.getAttribute("NOTE_DUE_DATE"))
            Else
                Me.NOTE_DUE_DATE.Text = String.Empty
            End If

            '票據號碼
            Me.NOTE_NO.Text = mbMEMREV.getString("NOTE_NO")

            '發票行銀行
            Me.NOTE_BANK.Text = Trim(mbMEMREV.getString("NOTE_BANK"))

            '發票行分行
            Me.NOTE_BR.Text = Trim(mbMEMREV.getString("NOTE_BR"))

            '發票人
            Me.NOTE_HOLDER.Text = Trim(mbMEMREV.getString("NOTE_HOLDER"))

            '票據金額
            Me.NOTE_AMT.Text = Utility.FormatDec(mbMEMREV.getDecimal("NOTE_AMT"), "#,##0")

            '補印次數
            Me.MB_REISU.Text = mbMEMREV.getString("MB_REISU")

            Me.PLH_MB_MEMREV.Visible = True

            Me.PLH_MB_MEMREV_LIST.Visible = False

            If Me.RBL_Choose.SelectedValue = "0" Then
                '列印收據
                Me.bt_Print.Visible = True
            ElseIf Me.RBL_Choose.SelectedValue = "1" Then
                '回收收據
                Me.bt_Return.Visible = True
            ElseIf Me.RBL_Choose.SelectedValue = "2" Then
                '補印收據
                Me.bt_Append.Visible = True
            End If
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    Sub ClearPLH_MB_MEMREV()
        Try
            '法名/姓名
            Me.MB_NAME.Text = String.Empty

            '會員編號
            Me.MB_MEMSEQ.Text = String.Empty
            Me.HID_MB_MEMSEQ.Value = String.Empty
            Me.HID_MB_SEQNO.Value = String.Empty
            Me.HID_MB_ITEMID.Value = String.Empty
            Me.HID_MB_MEMFEE_SY.Value = String.Empty
            Me.HID_MB_MEMFEE_EY.Value = String.Empty

            '通訊地址
            Me.MB_ADDR.Text = String.Empty

            '繳款日期
            Me.MB_TX_DATE.Text = String.Empty

            '功德項目
            Me.MB_ITEMID.Text = String.Empty
            '會員類別
            Me.PLH_MB_MEMTYP.Visible = False
            Me.MB_MEMTYP.Text = String.Empty

            '繳款方式
            Me.MB_FEETYPE.Text = String.Empty

            '繳款金額
            Me.MB_TOTFEE.Text = String.Empty
            '每月金額
            Me.PLH_MB_TOTFEE_MM.Visible = False
            Me.MB_TOTFEE_MM.Text = String.Empty

            '收據捐款名稱
            Me.MB_RECNAME.Text = String.Empty

            '所屬區
            Me.MB_AREA.Text = String.Empty

            '委員
            Me.MB_LEADER.Text = String.Empty

            '會費期間
            Me.MB_MEMFEE.Text = String.Empty
            Me.MB_DESC.Visible = False

            '付款方式
            Me.MB_PAY_TYPE.Text = String.Empty

            '票據到期日
            Me.NOTE_DUE_DATE.Text = String.Empty

            '票據號碼
            Me.NOTE_NO.Text = String.Empty

            '發票行
            Me.NOTE_BANK.Text = String.Empty

            '分行
            Me.NOTE_BR.Text = String.Empty

            '發票人
            Me.NOTE_HOLDER.Text = String.Empty

            '票據金額
            Me.NOTE_AMT.Text = String.Empty

            '補印次數
            Me.MB_REISU.Text = String.Empty

            '列印收據
            Me.bt_Print.Visible = False

            '回收收據
            Me.bt_Return.Visible = False

            '補印收據
            Me.bt_Append.Visible = False
        Catch ex As Exception
            Throw
        End Try
    End Sub
#End Region

#Region "Utility"
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

    Public Function get_MB_PAY_TYPE_TEXT(ByVal sMB_PAY_TYPE As Object) As String
        Try
            If Utility.isValidateData(sMB_PAY_TYPE) Then
                If IsNothing(Me.DT_UPCODE_31) Then
                    Dim apCODEList As New AP_CODEList(Me.m_DBManager)
                    apCODEList.loadByUpCode(Me.m_sUpcode31)
                    Me.DT_UPCODE_31 = apCODEList.getCurrentDataSet.Tables(0)
                End If

                Dim ROW_SELECT() As DataRow = Me.DT_UPCODE_31.Select("VALUE='" & sMB_PAY_TYPE & "'")
                If Not IsNothing(ROW_SELECT) AndAlso ROW_SELECT.Length > 0 Then
                    Return ROW_SELECT(0)("TEXT").ToString()
                End If
            End If

            Return String.Empty
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function get_MB_FEETYPE_NOTE(ByVal sMB_FEETYPE As Object) As String
        Try
            If Utility.isValidateData(sMB_FEETYPE) Then
                If IsNothing(Me.DT_UPCODE_23) Then
                    Dim apCODEList As New AP_CODEList(Me.m_DBManager)
                    apCODEList.loadByUpCode(Me.m_sUpcode23)
                    Me.DT_UPCODE_23 = apCODEList.getCurrentDataSet.Tables(0)
                End If

                Dim ROW_SELECT() As DataRow = Me.DT_UPCODE_23.Select("VALUE='" & sMB_FEETYPE & "'")
                If Not IsNothing(ROW_SELECT) AndAlso ROW_SELECT.Length > 0 Then
                    Return ROW_SELECT(0)("NOTE").ToString()
                End If
            End If

            Return String.Empty
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function get_MB_MEMTYP_TEXT(ByVal sMB_MEMTYP As Object) As String
        Try
            If Utility.isValidateData(sMB_MEMTYP) Then
                If IsNothing(Me.DT_UPCODE_28) Then
                    Dim apCODEList As New AP_CODEList(Me.m_DBManager)
                    apCODEList.loadByUpCode(Me.m_sUpCode28)
                    Me.DT_UPCODE_28 = apCODEList.getCurrentDataSet.Tables(0)
                End If

                Dim ROW_SELECT() As DataRow = Me.DT_UPCODE_28.Select("VALUE='" & sMB_MEMTYP & "'")
                If Not IsNothing(ROW_SELECT) AndAlso ROW_SELECT.Length > 0 Then
                    Return ROW_SELECT(0)("TEXT").ToString()
                End If
            End If

            Return String.Empty
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function get_MB_FEETYPE_TEXT(ByVal sMB_FEETYPE As Object) As String
        Try
            If Utility.isValidateData(sMB_FEETYPE) Then
                If IsNothing(Me.DT_UPCODE_23) Then
                    Dim apCODEList As New AP_CODEList(Me.m_DBManager)
                    apCODEList.loadByUpCode(Me.m_sUpcode23)
                    Me.DT_UPCODE_23 = apCODEList.getCurrentDataSet.Tables(0)
                End If

                Dim ROW_SELECT() As DataRow = Me.DT_UPCODE_23.Select("VALUE='" & sMB_FEETYPE & "'")
                If Not IsNothing(ROW_SELECT) AndAlso ROW_SELECT.Length > 0 Then
                    Return ROW_SELECT(0)("TEXT").ToString()
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

    Public Function get_MB_ITEMID_TEXT(ByVal sMB_ITEMID As Object) As String
        Try
            If Utility.isValidateData(sMB_ITEMID) Then
                If IsNothing(Me.DT_UPCODE_15) Then
                    Dim apCODEList As New AP_CODEList(Me.m_DBManager)
                    apCODEList.loadByUpCode(Me.m_sUpcode15)
                    Me.DT_UPCODE_15 = apCODEList.getCurrentDataSet.Tables(0)
                End If

                Dim ROW_SELECT() As DataRow = Me.DT_UPCODE_15.Select("VALUE='" & sMB_ITEMID & "'")
                If Not IsNothing(ROW_SELECT) AndAlso ROW_SELECT.Length > 0 Then
                    Return ROW_SELECT(0)("TEXT").ToString()
                End If
            End If

            Return String.Empty
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function get_MB_TOTFEE_TEXT(ByVal iMB_TOTFEE As Object) As String
        Try
            If IsNumeric(iMB_TOTFEE) Then
                Return Utility.FormatDec(iMB_TOTFEE, "#,##0")
            End If

            Return String.Empty
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function getMB_AREA(ByVal sMB_AREA As Object) As String
        Try
            If Utility.isValidateData(sMB_AREA) Then
                If IsNothing(Me.DT_UPCODE_7) Then
                    Dim apCODEList As New AP_CODEList(Me.m_DBManager)
                    apCODEList.loadByUpCode(Me.m_sUpCode7)
                    Me.DT_UPCODE_7 = apCODEList.getCurrentDataSet.Tables(0)
                End If

                Dim ROW_SELECT() As DataRow = Me.DT_UPCODE_7.Select("VALUE='" & sMB_AREA & "'")
                If Not IsNothing(ROW_SELECT) AndAlso ROW_SELECT.Length > 0 Then
                    Return ROW_SELECT(0)("TEXT").ToString()
                End If
            End If

            Return String.Empty
        Catch ex As Exception
            Throw
        End Try
    End Function

#End Region

End Class