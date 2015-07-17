Imports MBSC.MB_OP
Imports com.Azion.NET.VB
Imports MBSC.UICtl

Public Class MBMnt_MEMFEE_01_v01
    Inherits System.Web.UI.Page

    Dim m_DBManager As com.Azion.NET.VB.DatabaseManager
    Dim m_sUpCode7 As String = String.Empty
    Dim DT_UPCODE_7 As DataTable

    Dim m_sUpcode15 As String = String.Empty
    Dim DT_UPCODE_15 As DataTable

    Dim m_sUpcode23 As String = String.Empty
    Dim DT_UPCODE_23 As DataTable

    Dim m_sUpcode31 As String = String.Empty
    Dim DT_UPCODE_31 As DataTable

    Dim m_sUpCode28 As String = String.Empty
    Dim DT_UPCODE_28 As DataTable
    Dim m_sUpcode272 As String = String.Empty
    Dim DT_UPCODE15 As DataTable = Nothing
    Dim m_sUpcode276 As String = String.Empty

#Region "Page Events"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            '所屬區代碼
            Me.m_sUpCode7 = com.Azion.EloanUtility.CodeList.getAppSettings("UpCode7")

            '功德項目
            Me.m_sUpcode15 = com.Azion.EloanUtility.CodeList.getAppSettings("UpCode15")

            '繳款方式
            Me.m_sUpcode23 = com.Azion.EloanUtility.CodeList.getAppSettings("UpCode23")

            '付款方式
            Me.m_sUpcode31 = com.Azion.EloanUtility.CodeList.getAppSettings("UpCode31")

            '會員類別
            Me.m_sUpCode28 = com.Azion.EloanUtility.CodeList.getAppSettings("UpCode28")

            '給收據方式
            Me.m_sUpcode272 = com.Azion.EloanUtility.CodeList.getAppSettings("UpCode272")

            '專案名稱
            Me.m_sUpcode276 = com.Azion.EloanUtility.CodeList.getAppSettings("UpCode276")

            Me.m_DBManager = UIShareFun.getDataBaseManager

            If Not Page.IsPostBack Then
                Me.Bind_MB_AREA()

                Me.Bind_DDL_MB_AREA()

                Me.Bind_MB_SEND_PRINT()

                Me.Bind_MB_ITEMID()

                Me.Bind_MB_FAMILY()

                Me.Bind_MB_FEETYPE("0")

                Me.Bind_MB_PAY_TYPE()

                '專案名稱
                Me.Bind_PROJCODE()
            End If
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    Private Sub MBMnt_MEMFEE_01_v01_Unload(sender As Object, e As System.EventArgs) Handles Me.Unload
        UIShareFun.releaseConnection(Me.m_DBManager)
    End Sub
#End Region

#Region "QRY Button Events"
    Sub bt_Qry_SEL_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles bt_Qry_SEL.Click
        Try
            If Me.RB_QRY_1.Checked = False AndAlso Me.RB_QRY_1_2.Checked = False AndAlso Me.RB_QRY_2.Checked = False Then
                com.Azion.EloanUtility.UIUtility.alert("請選擇作業方式(覆核修正，過帳/列印傳票，或已入帳)")
                Return
            End If

            Me.ClearQRYSubItem()

            If Me.RB_QRY_1.Checked Then
                Me.PLH_QRY_SEL_SUB_1.Visible = True
                '覆核修正
                Me.btnModify.Visible = True
                Me.LTL_TITLE.Text = "繳款覆核修正(繳款作業第2關)"
            ElseIf Me.RB_QRY_1_2.Checked Then
                Me.PLH_QRY_SEL_SUB_1.Visible = True
                '列印傳票
                Me.btnPrint.Visible = True
                '過帳
                Me.bt_Approve_Data.Visible = True
                Me.LTL_TITLE.Text = "過帳/列印傳票(繳款作業第3關)"
            ElseIf Me.RB_QRY_2.Checked Then
                Me.PLH_QRY_SEL_SUB_2.Visible = True
                '沖帳
                Me.bt_Acct_Data.Visible = True
                Me.LTL_TITLE.Text = "已入帳(可沖帳或查詢已入帳)"
            End If

            Me.PLH_QRY_SEL.Visible = False
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    Sub ClearQRYSubItem()
        Try
            '未入帳
            Me.PLH_QRY_SEL_SUB_1.Visible = False

            '所屬區
            Me.RP_QRY_SUB_1_1.Checked = False
            Me.MB_AREA_SUB_1.SelectedIndex = -1

            '繳款年月
            Me.RP_QRY_SUB_1_2.Checked = False
            Me.txt_QRY_SUB_1_YYY.Text = String.Empty
            Me.txt_QRY_SUB_1_MM.Text = String.Empty

            '已入帳
            Me.PLH_QRY_SEL_SUB_2.Visible = False

            '所屬區
            Me.RP_QRY_SUB_2_1.Checked = False
            Me.MB_AREA_SUB_2.SelectedIndex = -1

            '繳款年月
            Me.RP_QRY_SUB_2_2.Checked = False
            Me.txt_QRY_SUB_2_BEGYYY.Text = String.Empty
            Me.txt_QRY_SUB_2_BEGMM.Text = String.Empty
            Me.txt_QRY_SUB_2_ENDYYY.Text = String.Empty
            Me.txt_QRY_SUB_2_ENDMM.Text = String.Empty

            '姓名
            Me.RP_QRY_SUB_2_3.Checked = False
            Me.txt_QRY_MB_NAME.Text = String.Empty

            '覆核修正
            Me.btnModify.Visible = False
            '列印傳票
            Me.btnPrint.Visible = False
            '過帳
            Me.bt_Approve_Data.Visible = False
            '沖帳
            Me.bt_Acct_Data.Visible = False

            Me.LTL_TITLE.Text = "繳款覆核修正/入帳/沖帳作業"
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Sub bt_Qry_SUB_1_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles bt_Qry_SUB_1.Click
        Try
            If RP_QRY_SUB_1_1.Checked = False AndAlso RP_QRY_SUB_1_2.Checked = False Then
                com.Azion.EloanUtility.UIUtility.alert("請選擇所屬區或繳款年月!")

                Return
            End If

            If RP_QRY_SUB_1_1.Checked AndAlso Not Utility.isValidateData(Me.MB_AREA_SUB_1.SelectedValue) Then
                com.Azion.EloanUtility.UIUtility.alert("請選擇所屬區!")

                Return
            End If

            If RP_QRY_SUB_1_2.Checked Then
                If Not IsNumeric(Me.txt_QRY_SUB_1_YYY.Text) Then
                    com.Azion.EloanUtility.UIUtility.alert("請輸入繳款年!")

                    Return
                End If

                If Not IsNumeric(Me.txt_QRY_SUB_1_MM.Text) Then
                    com.Azion.EloanUtility.UIUtility.alert("請輸入繳款月!")

                    Return
                End If

                If Not IsDate(CDec(Me.txt_QRY_SUB_1_YYY.Text) & "/" & Me.txt_QRY_SUB_1_MM.Text & "/1") Then
                    com.Azion.EloanUtility.UIUtility.alert("繳款年月輸入錯誤!")

                    Return
                End If

                Dim D_QRY_DATE As Date = New Date(CDec(Me.txt_QRY_SUB_1_YYY.Text), CDec(Me.txt_QRY_SUB_1_MM.Text), 1)
                If D_QRY_DATE > Now Then
                    com.Azion.EloanUtility.UIUtility.alert("繳款年月不可大於系統日!")

                    Return
                End If
            End If

            If Me.RP_QRY_SUB_1_1.Checked Then
                Dim mbMEMREVList As New MB_MEMREVList(Me.m_DBManager)
                If Me.RB_QRY_1.Checked Then
                    mbMEMREVList.loadByMB_AREA_1(Me.MB_AREA_SUB_1.SelectedValue)
                ElseIf Me.RB_QRY_1_2.Checked Then
                    mbMEMREVList.loadByMB_AREA_2(Me.MB_AREA_SUB_1.SelectedValue)
                ElseIf Me.RB_QRY_2.Checked Then
                    mbMEMREVList.loadByMB_AREA_3(Me.MB_AREA_SUB_1.SelectedValue)
                End If

                If mbMEMREVList.getCurrentDataSet.Tables(0).Rows.Count = 0 Then
                    If Me.RB_QRY_1.Checked Then
                        com.Azion.EloanUtility.UIUtility.alert(Me.MB_AREA_SUB_1.SelectedItem.Text & "該區無未覆核更正資料!")
                    ElseIf Me.RB_QRY_1_2.Checked Then
                        com.Azion.EloanUtility.UIUtility.alert(Me.MB_AREA_SUB_1.SelectedItem.Text & "該區無未入帳資料!")
                    ElseIf Me.RB_QRY_2.Checked Then
                        com.Azion.EloanUtility.UIUtility.alert(Me.MB_AREA_SUB_1.SelectedItem.Text & "該區無已入帳資料!")
                    End If

                    Return
                Else
                    Me.RP_MB_MEMREV.DataSource = mbMEMREVList.getCurrentDataSet.Tables(0)
                    Me.RP_MB_MEMREV.DataBind()
                End If
            ElseIf Me.RP_QRY_SUB_1_2.Checked Then
                Dim mbMEMREVList As New MB_MEMREVList(Me.m_DBManager)
                If Me.RB_QRY_1.Checked Then
                    mbMEMREVList.loadByMB_TX_DATE_1(Utility.FillZero(CDec(Me.txt_QRY_SUB_1_YYY.Text), 4) & Utility.FillZero(Me.txt_QRY_SUB_1_MM.Text, 2), Session("AREA"))
                ElseIf Me.RB_QRY_1_2.Checked Then
                    mbMEMREVList.loadByMB_TX_DATE_2(Utility.FillZero(CDec(Me.txt_QRY_SUB_1_YYY.Text), 4) & Utility.FillZero(Me.txt_QRY_SUB_1_MM.Text, 2), Session("AREA"))
                ElseIf Me.RB_QRY_2.Checked Then
                    Dim sBYYYYMM As String = String.Empty
                    sBYYYYMM = Utility.FillZero(CDec(Me.txt_QRY_SUB_2_BEGYYY.Text), 4) & Utility.FillZero(Me.txt_QRY_SUB_2_BEGMM.Text, 2)
                    Dim sEYYYYMM As String = String.Empty
                    sEYYYYMM = Utility.FillZero(CDec(Me.txt_QRY_SUB_2_ENDYYY.Text), 4) & Utility.FillZero(Me.txt_QRY_SUB_2_ENDMM.Text, 2)
                    mbMEMREVList.loadByMB_TX_DATE_3(sBYYYYMM, sEYYYYMM, Session("AREA"))
                End If

                If mbMEMREVList.getCurrentDataSet.Tables(0).Rows.Count = 0 Then
                    If Me.RB_QRY_1.Checked Then
                        com.Azion.EloanUtility.UIUtility.alert("該繳款年月" & "\n" & "西元" & Me.txt_QRY_SUB_1_YYY.Text & "年" & Me.txt_QRY_SUB_1_MM.Text & "月" & "\n" & "無未覆核更正資料!")
                    ElseIf Me.RB_QRY_1_2.Checked Then
                        com.Azion.EloanUtility.UIUtility.alert("該繳款年月" & "\n" & "西元" & Me.txt_QRY_SUB_1_YYY.Text & "年" & Me.txt_QRY_SUB_1_MM.Text & "月" & "\n" & "無未入帳資料!")
                    ElseIf Me.RB_QRY_2.Checked Then
                        Dim sBYYYYMM As String = String.Empty
                        sBYYYYMM = Utility.FillZero(CDec(Me.txt_QRY_SUB_2_BEGYYY.Text), 4) & Utility.FillZero(Me.txt_QRY_SUB_2_BEGMM.Text, 2)
                        Dim sEYYYYMM As String = String.Empty
                        sEYYYYMM = Utility.FillZero(CDec(Me.txt_QRY_SUB_2_ENDYYY.Text), 4) & Utility.FillZero(Me.txt_QRY_SUB_2_ENDMM.Text, 2)

                        com.Azion.EloanUtility.UIUtility.alert("該繳款年月" & "\n" & "西元" & Me.txt_QRY_SUB_2_BEGYYY.Text & "年" & Me.txt_QRY_SUB_2_BEGMM.Text & "月~西元" & Me.txt_QRY_SUB_2_ENDYYY.Text & "年" & Me.txt_QRY_SUB_2_ENDMM.Text & "月" & "\n" & "無已入帳資料!")
                    End If

                    Return
                Else
                    Me.RP_MB_MEMREV.DataSource = mbMEMREVList.getCurrentDataSet.Tables(0)
                    Me.RP_MB_MEMREV.DataBind()
                End If
            End If

            Me.PLH_MB_MEMREV.Visible = True
            Me.PLH_QRY.Visible = False
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    Sub bt_Qry_SUB_2_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles bt_Qry_SUB_2.Click
        Try
            If RP_QRY_SUB_2_1.Checked = False AndAlso RP_QRY_SUB_2_2.Checked = False AndAlso RP_QRY_SUB_2_3.Checked = False Then
                com.Azion.EloanUtility.UIUtility.alert("請選擇所屬區或繳款年月或姓名!")

                Return
            End If

            If RP_QRY_SUB_2_1.Checked AndAlso Not Utility.isValidateData(Me.MB_AREA_SUB_2.SelectedValue) Then
                com.Azion.EloanUtility.UIUtility.alert("請選擇所屬區!")

                Return
            End If

            If RP_QRY_SUB_2_2.Checked Then
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
            End If

            If Me.RP_QRY_SUB_2_3.Checked Then
                If Not Utility.isValidateData(Trim(Me.txt_QRY_MB_NAME.Text)) Then
                    com.Azion.EloanUtility.UIUtility.alert("請輸入姓名!")

                    Return
                End If
            End If

            If Me.RP_QRY_SUB_2_1.Checked Then
                Dim mbMEMREVList As New MB_MEMREVList(Me.m_DBManager)
                mbMEMREVList.loadByMB_AREA_3(Me.MB_AREA_SUB_2.SelectedValue)

                If mbMEMREVList.getCurrentDataSet.Tables(0).Rows.Count = 0 Then
                    com.Azion.EloanUtility.UIUtility.alert(Me.MB_AREA_SUB_2.SelectedItem.Text & "無已入帳資料!")

                    Return
                Else
                    Me.RP_MB_MEMREV.DataSource = mbMEMREVList.getCurrentDataSet.Tables(0)
                    Me.RP_MB_MEMREV.DataBind()
                End If
            ElseIf Me.RP_QRY_SUB_2_2.Checked Then
                Dim mbMEMREVList As New MB_MEMREVList(Me.m_DBManager)
                Dim iBEG As Decimal = (CDec(Me.txt_QRY_SUB_2_BEGYYY.Text)) * 100 + CDec(Me.txt_QRY_SUB_2_BEGMM.Text)
                Dim iEND As Decimal = (CDec(Me.txt_QRY_SUB_2_ENDYYY.Text)) * 100 + CDec(Me.txt_QRY_SUB_2_ENDMM.Text)
                mbMEMREVList.loadByMB_TX_DATE_3(iBEG, iEND, Session("AREA"))

                If mbMEMREVList.getCurrentDataSet.Tables(0).Rows.Count = 0 Then
                    Dim sMsg As String = String.Empty
                    sMsg = "西元" & Me.txt_QRY_SUB_2_BEGYYY.Text & "年" & Me.txt_QRY_SUB_2_BEGMM.Text & "月" & _
                           "至" & "西元" & Me.txt_QRY_SUB_2_ENDYYY.Text & "年" & Me.txt_QRY_SUB_2_ENDMM.Text & "月"
                    com.Azion.EloanUtility.UIUtility.alert("該繳款年月" & "\n" & sMsg & "\n" & "無已入帳資料!")

                    Return
                Else
                    Me.RP_MB_MEMREV.DataSource = mbMEMREVList.getCurrentDataSet.Tables(0)
                    Me.RP_MB_MEMREV.DataBind()
                End If
            ElseIf Me.RP_QRY_SUB_2_3.Checked Then
                Dim mbMEMREVList As New MB_MEMREVList(Me.m_DBManager)
                mbMEMREVList.loadByMB_NAME_VOUCHER_Y(Trim(Me.txt_QRY_MB_NAME.Text))

                If mbMEMREVList.getCurrentDataSet.Tables(0).Rows.Count = 0 Then
                    com.Azion.EloanUtility.UIUtility.alert("無已入帳資料!")

                    Return
                Else
                    Me.RP_MB_MEMREV.DataSource = mbMEMREVList.getCurrentDataSet.Tables(0)
                    Me.RP_MB_MEMREV.DataBind()
                End If
            End If

            Me.PLH_MB_MEMREV.Visible = True
            Me.PLH_QRY.Visible = False

            Me.TD_PRINT.Visible = True
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    Sub RP_MB_MEMREV_ItemDataBound(ByVal Sender As Object, ByVal e As RepeaterItemEventArgs) Handles RP_MB_MEMREV.ItemDataBound
        Try
            If e.Item.ItemType = ListItemType.Item OrElse e.Item.ItemType = ListItemType.AlternatingItem Then
                Dim TD_PRINT As HtmlTableCell = e.Item.FindControl("TD_PRINT")

                If Me.RB_QRY_2.Checked Then
                    TD_PRINT.Visible = True
                Else
                    TD_PRINT.Visible = False
                End If

                Dim DRV As DataRowView = CType(e.Item.DataItem, DataRowView)

                'Dim IMG_EDIT As ImageButton = e.Item.FindControl("IMG_EDIT")
                'If Not Utility.isValidateData(DRV("MB_PRINT")) Then
                '    IMG_EDIT.Visible = True
                'ElseIf (Utility.isValidateData(DRV("MB_PRINT")) AndAlso DRV("MB_PRINT") = "Y") AndAlso _
                '    (Utility.isValidateData(DRV("MB_REBREV")) AndAlso DRV("MB_REBREV") = "Y") Then
                '    IMG_EDIT.Visible = True
                'Else
                '    IMG_EDIT.Visible = False
                'End If
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Sub bt_Cancel_SUB_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles bt_Cancel_SUB_1.Click, bt_Cancel_SUB_2.Click
        Try
            Me.ClearQRYSubItem()

            Me.PLH_QRY_SEL.Visible = True
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub
#End Region

#Region "RP_MB_MEMREV"

    Sub bt_Back_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles bt_Back.Click
        Try
            Me.ClearMB_MEMREV()

            Me.ClearMB_MEMREV_DATA()

            Me.ClearQRYSubItem()

            Me.PLH_QRY_SEL.Visible = True
            Me.PLH_QRY.Visible = True
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    Sub ClearMB_MEMREV()
        Try
            Me.PLH_MB_MEMREV.Visible = False

            Me.TD_PRINT.Visible = False
            Me.RP_MB_MEMREV.DataSource = Nothing
            Me.RP_MB_MEMREV.DataBind()
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Sub ClearEDIT()
        Try
            '功德項目
            Me.DDL_MB_ITEMID.SelectedIndex = -1
            Me.PLH_MB_MEMTYP.Visible = False
            Me.PLH_PROJCODE.Visible = False
            Me.PROJCODE.SelectedIndex = -1

            '繳款方式
            Me.RBL_MB_FEETYPE.SelectedIndex = -1

            '繳款金額
            Me.TXT_MB_TOTFEE.Text = String.Empty
            Me.TXT_MB_TOTFEE_MM.Text = String.Empty
            Me.PLH_MB_TOTFEE_MM.Visible = False

            '所屬區
            Me.DDL_MB_AREA.SelectedIndex = -1
            '委員
            Me.DDL_MB_LEADER.SelectedIndex = -1

            '是否開立收據
            Me.MB_PRINT.SelectedIndex = -1

            '給收據方式
            Me.MB_SEND_PRINT.SelectedIndex = -1
            Me.LTL_MB_SEND_PRINT.Visible = False
            Me.MB_SEND_PRINT.Visible = False
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Sub RP_MB_MEMREV_ItemCommand(ByVal Sender As Object, ByVal e As RepeaterCommandEventArgs) Handles RP_MB_MEMREV.ItemCommand
        Try
            Me.ClearEDIT()

            If Me.RB_QRY_1.Checked Then
                Me.SH_MODIFY(False, True)
            Else
                Me.SH_MODIFY(True, False)
            End If

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
                Me.MB_TX_DATE_YYY.Text = CDec(Left(mbMEMREV.getString("MB_TX_DATE"), 4))
                Me.MB_TX_DATE_MM.Text = mbMEMREV.getString("MB_TX_DATE").Substring(4, 2)
                Me.MB_TX_DATE_DD.Text = Right(mbMEMREV.getString("MB_TX_DATE"), 2)
            Else
                Me.MB_TX_DATE.Text = String.Empty
                Me.MB_TX_DATE_YYY.Text = String.Empty
                Me.MB_TX_DATE_MM.Text = String.Empty
                Me.MB_TX_DATE_DD.Text = String.Empty
            End If

            '功德項目
            Me.MB_ITEMID.Text = Me.get_MB_ITEMID_TEXT(mbMEMREV.getString("MB_ITEMID"))
            Me.DDL_MB_ITEMID.SelectedIndex = -1
            If Not IsNothing(Me.DDL_MB_ITEMID.Items.FindByValue(mbMEMREV.getString("MB_ITEMID"))) Then
                Me.DDL_MB_ITEMID.Items.FindByValue(mbMEMREV.getString("MB_ITEMID")).Selected = True
            End If
            Me.HID_MB_ITEMID.Value = mbMEMREV.getString("MB_ITEMID")
            Me.RBL_MB_MEMTYP.SelectedIndex = -1
            If mbMEMREV.getString("MB_ITEMID") = "A" Then
                Me.PLH_MB_MEMTYP.Visible = True
                Me.MB_MEMTYP.Text = Me.get_MB_MEMTYP_TEXT(mbMEMREV.getString("MB_MEMTYP"))
                If Not IsNothing(Me.RBL_MB_MEMTYP.Items.FindByValue(mbMEMREV.getString("MB_MEMTYP"))) Then
                    Me.RBL_MB_MEMTYP.Items.FindByValue(mbMEMREV.getString("MB_MEMTYP")).Selected = True
                End If
                If Me.RBL_MB_MEMTYP.SelectedValue = "1" Then
                    Me.PLH_MB_TOTFEE_MM.Visible = True
                    Me.Bind_MB_FEETYPE("1")
                Else
                    Me.PLH_MB_TOTFEE_MM.Visible = False
                    Me.Bind_MB_FEETYPE("2")
                End If

                Me.PLH_PROJCODE.Visible = False
                Me.PROJCODE.SelectedIndex = -1
            Else
                Me.PLH_MB_MEMTYP.Visible = False
                Me.MB_MEMTYP.Text = String.Empty
                Me.PLH_MB_TOTFEE_MM.Visible = False

                If mbMEMREV.getString("MB_ITEMID") = "B" Then
                    Me.PLH_PROJCODE.Visible = True
                    If Not IsNothing(Me.PROJCODE.Items.FindByValue(mbMEMREV.getString("PROJCODE"))) Then
                        Me.PROJCODE.Items.FindByValue(mbMEMREV.getString("PROJCODE")).Selected = True
                    End If
                Else
                    Me.PLH_PROJCODE.Visible = False
                    Me.PROJCODE.SelectedIndex = -1
                End If
            End If

            '繳款方式
            Me.MB_FEETYPE.Text = Me.get_MB_FEETYPE_TEXT(mbMEMREV.getString("MB_FEETYPE"))
            Me.RBL_MB_FEETYPE.SelectedIndex = -1
            If Not IsNothing(Me.RBL_MB_FEETYPE.Items.FindByValue(mbMEMREV.getString("MB_FEETYPE"))) Then
                Me.RBL_MB_FEETYPE.Items.FindByValue(mbMEMREV.getString("MB_FEETYPE")).Selected = True
            End If

            Me.TXT_MB_TOTFEE.Attributes.Remove("readonly")
            Select Case mbMEMREV.getString("MB_FEETYPE")
                Case "A", "B", "C", "F"
                    Me.TXT_MB_TOTFEE.Attributes.Add("readonly", "true")
            End Select

            '繳款金額
            If IsNumeric(mbMEMREV.getAttribute("MB_TOTFEE")) Then
                Me.MB_TOTFEE.Text = Utility.FormatDec(mbMEMREV.getAttribute("MB_TOTFEE"), "#,##0")
                Me.TXT_MB_TOTFEE.Text = Utility.FormatDec(mbMEMREV.getAttribute("MB_TOTFEE"), "#,##0")
            Else
                Me.MB_TOTFEE.Text = String.Empty
            End If

            '每月金額
            'If mbMEMREV.getString("MB_FEETYPE") <> "D" Then
            '    Me.PLH_MB_TOTFEE_MM.Visible = True
            '    Dim sNOTE As String = String.Empty
            '    sNOTE = Me.get_MB_FEETYPE_NOTE(mbMEMREV.getString("MB_FEETYPE"))
            '    If IsNumeric(sNOTE) Then
            '        Me.MB_TOTFEE_MM.Text = Utility.FormatDec(sNOTE, "#,##0")
            '        Me.TXT_MB_TOTFEE_MM.Text = Utility.FormatDec(sNOTE, "#,##0")
            '    Else
            '        Me.MB_TOTFEE_MM.Text = String.Empty
            '    End If
            'Else
            '    Me.PLH_MB_TOTFEE_MM.Visible = False
            '    Me.MB_TOTFEE_MM.Text = String.Empty
            '    Me.TXT_MB_TOTFEE_MM.Text = String.Empty
            'End If
            If mbMEMREV.getString("MB_ITEMID") = "A" Then
                If mbMEMREV.getString("MB_MEMTYP") = "1" Then
                    Me.PLH_MB_TOTFEE_MM.Visible = True
                    If Not IsNothing(mbMEMREV.getAttribute("MB_MONFEE")) Then
                        Me.TXT_MB_TOTFEE_MM.Text = Utility.FormatDec(mbMEMREV.getAttribute("MB_MONFEE"), "#,##0")
                    Else
                        Me.TXT_MB_TOTFEE_MM.Text = String.Empty
                    End If
                Else
                    Me.PLH_MB_TOTFEE_MM.Visible = False
                    Me.TXT_MB_TOTFEE_MM.Text = String.Empty
                End If
            Else
                Me.PLH_MB_TOTFEE_MM.Visible = False
                Me.TXT_MB_TOTFEE_MM.Text = String.Empty
            End If

            '繳款分配項目
            If mbMEMREV.getString("FUND1") = "1" Then
                Me.FUND1.Checked = True
            Else
                Me.FUND2.Checked = False
            End If
            If mbMEMREV.getString("FUND2") = "1" Then
                Me.FUND2.Checked = True
            Else
                Me.FUND2.Checked = False
            End If
            If mbMEMREV.getString("FUND3") = "1" Then
                Me.FUND3.Checked = True
            Else
                Me.FUND3.Checked = False
            End If

            '收據捐款名稱
            Me.MB_RECNAME.Text = mbMEMREV.getString("MB_RECNAME")
            Me.TXT_MB_RECNAME.Text = mbMEMREV.getString("MB_RECNAME")

            '所屬區
            Me.MB_AREA.Text = Me.getMB_AREA(mbMEMBER.getString("MB_AREA"))
            Me.DDL_MB_AREA.SelectedIndex = -1
            If Not IsNothing(Me.DDL_MB_AREA.Items.FindByValue(mbMEMBER.getString("MB_AREA"))) Then
                Me.DDL_MB_AREA.Items.FindByValue(mbMEMBER.getString("MB_AREA")).Selected = True
            End If

            '所屬區委員
            Me.MB_LEADER.Text = mbMEMBER.getString("MB_LEADER")
            Me.ReBind_MB_LEADER(mbMEMBER.getString("MB_AREA"), mbMEMBER.getString("MB_LEADER"))

            '會費期間
            If mbMEMREV.getString("MB_ITEMID") = "A" Then
                Me.MB_DESC.Visible = False
                Me.MB_DESC.Text = String.Empty

                Me.MB_MEMFEE.Visible = True
                Me.MB_MEMFEE.Text = (mbMEMREV.getDecimal("MB_MEMFEE_SY")) & "年" & mbMEMREV.getString("MB_MEMFEE_SM") & "月至" & _
                                    (mbMEMREV.getDecimal("MB_MEMFEE_EY")) & "年" & mbMEMREV.getString("MB_MEMFEE_EM") & "月"

                Me.HID_MB_MEMFEE_SY.Value = mbMEMREV.getDecimal("MB_MEMFEE_SY")

                Me.MB_MEMFEE_SY.Text = (mbMEMREV.getDecimal("MB_MEMFEE_SY"))
                Me.MB_MEMFEE_SM.Text = mbMEMREV.getString("MB_MEMFEE_SM")
                Me.MB_MEMFEE_EY.Text = (mbMEMREV.getDecimal("MB_MEMFEE_EY"))
                Me.MB_MEMFEE_EM.Text = mbMEMREV.getString("MB_MEMFEE_EM")
            Else
                Me.MB_MEMFEE.Visible = False
                Me.MB_MEMFEE.Text = String.Empty

                Me.MB_DESC.Visible = True
                Me.MB_DESC.Text = mbMEMREV.getString("MB_DESC")
            End If

            '付款方式
            Me.MB_PAY_TYPE.Text = Me.get_MB_PAY_TYPE_TEXT(mbMEMREV.getString("MB_PAY_TYPE"))
            Me.RBL_MB_PAY_TYPE.SelectedIndex = -1
            If Not IsNothing(Me.RBL_MB_PAY_TYPE.Items.FindByValue(mbMEMREV.getString("MB_PAY_TYPE"))) Then
                Me.RBL_MB_PAY_TYPE.Items.FindByValue(mbMEMREV.getString("MB_PAY_TYPE")).Selected = True
            End If

            '是否開立收據
            Me.MB_PRINT.SelectedIndex = -1
            If Not IsNothing(Me.MB_PRINT.Items.FindByValue(mbMEMREV.getString("MB_PRINT"))) Then
                Me.MB_PRINT.Items.FindByValue(mbMEMREV.getString("MB_PRINT")).Selected = True
            End If

            '給收據方式
            Me.MB_SEND_PRINT.SelectedIndex = -1
            If Not IsNothing(Me.MB_SEND_PRINT.Items.FindByValue(mbMEMREV.getString("MB_SEND_PRINT"))) Then
                Me.MB_SEND_PRINT.Items.FindByValue(mbMEMREV.getString("MB_SEND_PRINT")).Selected = True
            End If

            If Utility.isValidateData(Me.MB_PRINT.SelectedValue) Then
                If Me.MB_PRINT.SelectedValue = "0" Then
                    Me.LTL_MB_SEND_PRINT.Visible = False
                    Me.MB_SEND_PRINT.Visible = False
                Else
                    Me.LTL_MB_SEND_PRINT.Visible = True
                    Me.MB_SEND_PRINT.Visible = True
                End If
            Else
                Me.LTL_MB_SEND_PRINT.Visible = False
                Me.MB_SEND_PRINT.Visible = False
            End If

            If mbMEMREV.getString("MB_PAY_TYPE") = "N" Then
                Me.PLH_MB_PAY_TYPE.Visible = True

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
            Else
                Me.PLH_MB_PAY_TYPE.Visible = False

                Me.NOTE_DUE_DATE.Text = String.Empty

                Me.NOTE_NO.Text = String.Empty

                Me.NOTE_BANK.Text = String.Empty

                Me.NOTE_BR.Text = String.Empty

                Me.NOTE_HOLDER.Text = String.Empty
            End If

            Me.PLH_MB_MEMREV_DATA.Visible = True

            Me.PLH_MB_MEMREV_LIST.Visible = False
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    Sub bt_Cancel_Data_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles bt_Cancel_Data.Click
        Try
            Me.ClearMB_MEMREV_DATA()

            Me.PLH_MB_MEMREV_LIST.Visible = True
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    Sub ClearMB_MEMREV_DATA()
        Try
            Me.PLH_MB_MEMREV_DATA.Visible = False

            '法名/姓名
            Me.MB_NAME.Text = String.Empty

            '會員編號
            Me.MB_MEMSEQ.Text = String.Empty
            Me.HID_MB_MEMSEQ.Value = String.Empty
            Me.HID_MB_SEQNO.Value = String.Empty

            '通訊地址
            Me.MB_ADDR.Text = String.Empty

            '繳款日期
            Me.MB_TX_DATE.Text = String.Empty

            '功德項目
            Me.MB_ITEMID.Text = String.Empty
            Me.HID_MB_ITEMID.Value = String.Empty
            Me.PLH_MB_MEMTYP.Visible = False
            Me.PLH_PROJCODE.Visible = False
            Me.PROJCODE.SelectedIndex = -1
            Me.MB_MEMTYP.Text = String.Empty

            '繳款方式
            Me.MB_FEETYPE.Text = String.Empty

            '繳款金額
            Me.MB_TOTFEE.Text = String.Empty
            Me.PLH_MB_TOTFEE_MM.Visible = False
            '每月金額
            Me.MB_TOTFEE_MM.Text = String.Empty

            '收據捐款名稱
            Me.MB_RECNAME.Text = String.Empty

            '所屬區
            Me.MB_AREA.Text = String.Empty

            '所屬區委員
            Me.MB_LEADER.Text = String.Empty

            '會費期間
            Me.MB_MEMFEE.Text = String.Empty
            Me.HID_MB_MEMFEE_SY.Value = String.Empty
            Me.HID_MB_MEMFEE_EY.Value = String.Empty
            Me.MB_DESC.Text = String.Empty
            Me.MB_DESC.Visible = False

            '付款方式
            Me.MB_PAY_TYPE.Text = String.Empty

            '票據到期日
            Me.PLH_MB_PAY_TYPE.Visible = False
            Me.NOTE_DUE_DATE.Text = String.Empty

            '票據號碼
            Me.NOTE_NO.Text = String.Empty

            '發票行銀行
            Me.NOTE_BANK.Text = String.Empty

            '發票行分行
            Me.NOTE_BR.Text = String.Empty

            '發票人
            Me.NOTE_HOLDER.Text = String.Empty
        Catch ex As Exception
            Throw
        End Try
    End Sub
#End Region

#Region "DDL"
    Sub Bind_MB_AREA()
        Dim DT_MB_AREA As DataTable = Nothing
        Try
            Me.MB_AREA_SUB_1.Items.Clear()
            Me.MB_AREA_SUB_2.Items.Clear()

            Dim apCODEList As New AP_CODEList(Me.m_DBManager)
            apCODEList.loadByUpCode(m_sUpCode7)
            DT_MB_AREA = apCODEList.getCurrentDataSet.Tables(0)

            If Utility.isValidateData(Session("BRID")) AndAlso Session("BRID") <> "0" Then
                If Utility.isValidateData(Session("AREA")) Then
                    Dim ROW_OTH_AREA() As DataRow = Nothing
                    ROW_OTH_AREA = DT_MB_AREA.Select("VALUE<>'" & Session("AREA") & "'")
                    If Not IsNothing(ROW_OTH_AREA) AndAlso ROW_OTH_AREA.Length > 0 Then
                        For i As Integer = 0 To UBound(ROW_OTH_AREA)
                            DT_MB_AREA.Rows.Remove(ROW_OTH_AREA(i))
                        Next
                    End If
                End If
            End If

            Me.MB_AREA_SUB_1.DataSource = DT_MB_AREA
            Me.MB_AREA_SUB_1.DataTextField = "TEXT"
            Me.MB_AREA_SUB_1.DataValueField = "VALUE"
            Me.MB_AREA_SUB_1.DataBind()

            If Me.MB_AREA_SUB_1.Items.Count > 1 Then
                Me.MB_AREA_SUB_1.Items.Insert(0, New ListItem("請選擇", ""))
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

    Sub Bind_MB_SEND_PRINT()
        Try
            Me.MB_SEND_PRINT.Items.Clear()

            Dim apCODEList As New AP_CODEList(Me.m_DBManager)
            apCODEList.loadByUpCode(Me.m_sUpcode272)

            Me.MB_SEND_PRINT.DataSource = apCODEList.getCurrentDataSet.Tables(0)
            Me.MB_SEND_PRINT.DataTextField = "TEXT"
            Me.MB_SEND_PRINT.DataValueField = "VALUE"
            Me.MB_SEND_PRINT.DataBind()
            Me.MB_SEND_PRINT.Items.Insert(0, New ListItem("請選擇", ""))
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Sub Bind_MB_ITEMID()
        Try
            Me.DDL_MB_ITEMID.Items.Clear()

            Dim apCODEList As New AP_CODEList(Me.m_DBManager)
            apCODEList.loadByUpCode(Me.m_sUpcode15)

            Me.DDL_MB_ITEMID.DataSource = apCODEList.getCurrentDataSet.Tables(0)
            Me.DDL_MB_ITEMID.DataTextField = "TEXT"
            Me.DDL_MB_ITEMID.DataValueField = "VALUE"
            Me.DDL_MB_ITEMID.DataBind()

            Me.DDL_MB_ITEMID.Items.Insert(0, New ListItem("請選擇", ""))
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Sub Bind_MB_FAMILY()
        Try
            Me.RBL_MB_MEMTYP.Items.Clear()

            Dim apCODEList As New AP_CODEList(Me.m_DBManager)
            apCODEList.loadByUpCode(m_sUpCode28)

            Me.RBL_MB_MEMTYP.DataSource = apCODEList.getCurrentDataSet.Tables(0)
            Me.RBL_MB_MEMTYP.DataTextField = "TEXT"
            Me.RBL_MB_MEMTYP.DataValueField = "VALUE"
            Me.RBL_MB_MEMTYP.DataBind()
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Sub Bind_MB_FEETYPE(ByVal sMode As String)
        Dim DT_23 As DataTable = Nothing
        Try
            Me.RBL_MB_FEETYPE.Items.Clear()

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

            Me.RBL_MB_FEETYPE.DataSource = DT_23
            Me.RBL_MB_FEETYPE.DataTextField = "TEXT"
            Me.RBL_MB_FEETYPE.DataValueField = "VALUE"
            Me.RBL_MB_FEETYPE.DataBind()
        Catch ex As Exception
            Throw
        Finally
            If Not IsNothing(DT_23) Then
                DT_23.Dispose()
            End If
        End Try
    End Sub

    Sub Bind_MB_PAY_TYPE()
        Try
            Me.RBL_MB_PAY_TYPE.Items.Clear()

            Dim apCODEList As New AP_CODEList(Me.m_DBManager)
            apCODEList.loadByUpCode(Me.m_sUpcode31)

            Me.RBL_MB_PAY_TYPE.DataSource = apCODEList.getCurrentDataSet.Tables(0)
            Me.RBL_MB_PAY_TYPE.DataTextField = "TEXT"
            Me.RBL_MB_PAY_TYPE.DataValueField = "VALUE"
            Me.RBL_MB_PAY_TYPE.DataBind()
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Sub Bind_PROJCODE()
        Try
            Me.PROJCODE.Items.Clear()

            Dim apCODEList As New AP_CODEList(Me.m_DBManager)
            apCODEList.loadByUpCode(Me.m_sUpcode276)

            Me.PROJCODE.DataSource = apCODEList.getCurrentDataSet.Tables(0)
            Me.PROJCODE.DataTextField = "TEXT"
            Me.PROJCODE.DataValueField = "VALUE"
            Me.PROJCODE.DataBind()
            Me.PROJCODE.Items.Insert(0, New ListItem("請選擇", ""))
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Sub ReBind_MB_LEADER(ByVal sGA_AREA As String, ByVal sSelected As String)
        Try
            Dim apCODEList As New AP_CODEList(Me.m_DBManager)
            apCODEList.loadByUpCode(m_sUpCode7)
            Dim ROW_Select() As DataRow = apCODEList.getCurrentDataSet.Tables(0).Select("VALUE='" & sGA_AREA & "'")
            If Not IsNothing(ROW_Select) AndAlso ROW_Select.Length > 0 Then
                apCODEList.clear()
                apCODEList.loadByUpCode(ROW_Select(0)("CODEID"))

                Me.DDL_MB_LEADER.Items.Clear()
                Me.DDL_MB_LEADER.DataTextField = "TEXT"
                Me.DDL_MB_LEADER.DataValueField = "VALUE"
                Me.DDL_MB_LEADER.DataSource = apCODEList.getCurrentDataSet.Tables(0)
                Me.DDL_MB_LEADER.DataBind()

                Me.DDL_MB_LEADER.Items.Insert(0, New ListItem("請選擇", ""))

                Me.DDL_MB_LEADER.SelectedIndex = -1
                If Not IsNothing(Me.DDL_MB_LEADER.Items.FindByValue(sSelected)) Then
                    Me.DDL_MB_LEADER.Items.FindByValue(sSelected).Selected = True
                End If
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub
#End Region

#Region "Utility"
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

    Function get_MB_ITEMID_NOTE(ByVal sMB_ITEMID As Object) As String
        Try
            If Utility.isValidateData(sMB_ITEMID) Then
                If IsNothing(Me.DT_UPCODE_15) Then
                    Dim apCODEList As New AP_CODEList(Me.m_DBManager)
                    apCODEList.loadByUpCode(Me.m_sUpcode15)
                    Me.DT_UPCODE_15 = apCODEList.getCurrentDataSet.Tables(0)
                End If

                Dim ROW_SELECT() As DataRow = Me.DT_UPCODE_15.Select("VALUE='" & sMB_ITEMID & "'")
                If Not IsNothing(ROW_SELECT) AndAlso ROW_SELECT.Length > 0 Then
                    Return ROW_SELECT(0)("NOTE").ToString()
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

    Function get_Print_Desc(ByVal sMB_PRINT As Object, ByVal sMB_REBREV As Object) As String
        Try
            If Utility.isValidateData(sMB_PRINT) Then
                If sMB_PRINT = "Y" AndAlso (Utility.isValidateData(sMB_REBREV) AndAlso sMB_REBREV = "Y") Then
                    Return "已回收"
                ElseIf sMB_PRINT = "Y" AndAlso Not Utility.isValidateData(sMB_REBREV) Then
                    Return "已列印"
                End If
            Else
                Return "未列印"
            End If

            Return String.Empty
        Catch ex As Exception
            Throw
        End Try
    End Function
#End Region

#Region "SelectedIndexChanged"
    Private Sub DDL_MB_AREA_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DDL_MB_AREA.SelectedIndexChanged
        Try
            Me.ReBind_MB_LEADER(Me.DDL_MB_AREA.SelectedValue, String.Empty)
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

    Private Sub DDL_MB_ITEMID_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DDL_MB_ITEMID.SelectedIndexChanged
        Try
            If Me.DDL_MB_ITEMID.SelectedValue = "A" Then
                Me.PLH_MB_MEMTYP.Visible = True
                Me.PLH_PROJCODE.Visible = False
            ElseIf Me.DDL_MB_ITEMID.SelectedValue = "B" Then
                Me.PLH_MB_MEMTYP.Visible = False
                Me.PLH_PROJCODE.Visible = True
            Else
                Me.PLH_MB_MEMTYP.Visible = False
                Me.PLH_PROJCODE.Visible = False
            End If
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

    Private Sub RBL_MB_MEMTYP_SelectedIndexChanged(sender As Object, e As EventArgs) Handles RBL_MB_MEMTYP.SelectedIndexChanged
        Try
            Me.Bind_MB_FEETYPE(Me.RBL_MB_MEMTYP.SelectedValue)

            If Me.RBL_MB_MEMTYP.SelectedValue = "1" Then
                Me.PLH_MB_TOTFEE_MM.Visible = True
            Else
                Me.PLH_MB_TOTFEE_MM.Visible = False
            End If
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

    Private Sub RBL_MB_FEETYPE_SelectedIndexChanged(sender As Object, e As EventArgs) Handles RBL_MB_FEETYPE.SelectedIndexChanged
        Try
            If Not Utility.isValidateData(Me.DDL_MB_ITEMID.SelectedValue) Then
                com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "請選擇功德項目")
                com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請選擇功德項目")
                Return
            End If

            If Me.DDL_MB_ITEMID.SelectedValue = "A" Then
                If Not Utility.isValidateData(Me.RBL_MB_FEETYPE.SelectedValue) Then
                    com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "請選擇會員類別")
                    com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請選擇會員類別")
                    Return
                End If
            End If

            If Me.RBL_MB_FEETYPE.SelectedValue = "D" OrElse Me.RBL_MB_FEETYPE.SelectedValue = "E" Then
                Me.TXT_MB_TOTFEE.Attributes.Remove("readonly")
            Else
                Me.TXT_MB_TOTFEE.Attributes.Add("readonly", "true")
            End If

            If IsNumeric(Me.TXT_MB_TOTFEE_MM.Text) Then
                Select Case Me.RBL_MB_FEETYPE.SelectedValue
                    Case "A"
                        Me.TXT_MB_TOTFEE.Text = Me.TXT_MB_TOTFEE_MM.Text
                    Case "B"
                        Me.TXT_MB_TOTFEE.Text = CDec(Me.TXT_MB_TOTFEE_MM.Text) * 3
                    Case "C"
                        Me.TXT_MB_TOTFEE.Text = CDec(Me.TXT_MB_TOTFEE_MM.Text) * 12
                    Case "D", "E"
                        'MB_MEMREV.TOTFEE =繳款金額 (EDIT)
                    Case "F"
                        'MB_MEMREV.TOTFEE =100,000 (display)
                        Me.TXT_MB_TOTFEE.Text = "100,000"
                End Select
            Else
                Select Case Me.RBL_MB_FEETYPE.SelectedValue
                    Case "A"
                        Me.TXT_MB_TOTFEE.Text = String.Empty
                    Case "B"
                        Me.TXT_MB_TOTFEE.Text = String.Empty
                    Case "C"
                        Me.TXT_MB_TOTFEE.Text = String.Empty
                    Case "D", "E"
                        'MB_MEMREV.TOTFEE =繳款金額 (EDIT)
                    Case "F"
                        'MB_MEMREV.TOTFEE =100,000 (display)
                        Me.TXT_MB_TOTFEE.Text = "100,000"
                End Select
            End If
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

    Private Sub TXT_MB_TOTFEE_MM_TextChanged(sender As Object, e As EventArgs) Handles TXT_MB_TOTFEE_MM.TextChanged
        Try
            If IsNumeric(Me.TXT_MB_TOTFEE_MM.Text) Then
                Select Case Me.RBL_MB_FEETYPE.SelectedValue
                    Case "A"
                        Me.TXT_MB_TOTFEE.Text = Me.TXT_MB_TOTFEE_MM.Text
                    Case "B"
                        Me.TXT_MB_TOTFEE.Text = CDec(Me.TXT_MB_TOTFEE_MM.Text) * 3
                    Case "C"
                        Me.TXT_MB_TOTFEE.Text = CDec(Me.TXT_MB_TOTFEE_MM.Text) * 12
                    Case "D", "E"
                        'MB_MEMREV.TOTFEE =繳款金額 (EDIT)
                    Case "F"
                        'MB_MEMREV.TOTFEE =100,000 (display)
                        Me.TXT_MB_TOTFEE.Text = "100,000"
                End Select
            Else
                Select Case Me.RBL_MB_FEETYPE.SelectedValue
                    Case "A"
                        Me.TXT_MB_TOTFEE.Text = String.Empty
                    Case "B"
                        Me.TXT_MB_TOTFEE.Text = String.Empty
                    Case "C"
                        Me.TXT_MB_TOTFEE.Text = String.Empty
                    Case "D", "E"
                        'MB_MEMREV.TOTFEE =繳款金額 (EDIT)
                    Case "F"
                        'MB_MEMREV.TOTFEE =100,000 (display)
                        Me.TXT_MB_TOTFEE.Text = "100,000"
                End Select
            End If
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

    Private Sub MB_PRINT_SelectedIndexChanged(sender As Object, e As EventArgs) Handles MB_PRINT.SelectedIndexChanged
        Try
            If Me.MB_PRINT.SelectedValue = "1" Then
                Me.LTL_MB_SEND_PRINT.Visible = True
                Me.MB_SEND_PRINT.Visible = True
            Else
                Me.LTL_MB_SEND_PRINT.Visible = False
                Me.MB_SEND_PRINT.Visible = False
            End If
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub
#End Region

#Region "覆核修正"
    Private Sub btnModify_Click(sender As Object, e As EventArgs) Handles btnModify.Click
        Try
            Dim sTX_DATE_YYY As String = String.Empty
            If IsNumeric(Me.MB_TX_DATE_YYY.Text) Then
                sTX_DATE_YYY = CDec(Me.MB_TX_DATE_YYY.Text)
            End If

            Dim sTX_DATE_MM As String = String.Empty
            If IsNumeric(Me.MB_TX_DATE_MM.Text) Then
                sTX_DATE_MM = CDec(Me.MB_TX_DATE_MM.Text)
            End If

            Dim sTX_DATE_DD As String = String.Empty
            If IsNumeric(Me.MB_TX_DATE_DD.Text) Then
                sTX_DATE_DD = CDec(Me.MB_TX_DATE_DD.Text)
            End If

            If Not IsDate(sTX_DATE_YYY & "/" & sTX_DATE_MM & "/" & sTX_DATE_DD) Then
                com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "繳款日期錯誤或未完整輸入!")
                com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "繳款日期錯誤或未完整輸入!")
                Return
            End If

            If Not Utility.isValidateData(Me.DDL_MB_ITEMID.SelectedValue) Then
                com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "請選擇功德項目!")
                com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請選擇功德項目!")
                Return
            End If

            If Me.DDL_MB_ITEMID.SelectedValue = "A" Then
                If Not Utility.isValidateData(Me.RBL_MB_MEMTYP.SelectedValue) Then
                    com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "請選擇會員類別!")
                    com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請選擇會員類別!")
                    Return
                End If
            End If

            If Not Utility.isValidateData(Me.RBL_MB_FEETYPE.SelectedValue) Then
                com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "請選擇繳款方式!")
                com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請選擇繳款方式!")
                Return
            End If

            If Not IsNumeric(Me.TXT_MB_TOTFEE.Text) Then
                com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "請輸入繳款金額!")
                com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請輸入繳款金額!")
                Return
            End If

            If Not Utility.isValidateData(Me.TXT_MB_RECNAME.Text) Then
                com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "請輸入收據捐款名稱!")
                com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請輸入收據捐款名稱!")

                Return
            End If

            If Not Utility.isValidateData(Me.DDL_MB_AREA.SelectedValue) Then
                com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "請選擇所屬區!")
                com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請選擇所屬區!")
                Return
            End If

            If Not Utility.isValidateData(Me.DDL_MB_LEADER.SelectedValue) Then
                com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "請選擇委員!")
                com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請選擇委員!")
                Return
            End If

            If Me.DDL_MB_ITEMID.SelectedValue = "A" Then
                Dim sMB_MEMFEE_SY As String = String.Empty
                If IsNumeric(Me.MB_MEMFEE_SY.Text) Then
                    sMB_MEMFEE_SY = CDec(Me.MB_MEMFEE_SY.Text)
                End If
                Dim sMB_MEMFEE_SM As String = String.Empty
                If IsNumeric(Me.MB_MEMFEE_SM.Text) Then
                    sMB_MEMFEE_SM = CDec(Me.MB_MEMFEE_SM.Text)
                End If
                If Not IsDate(sMB_MEMFEE_SY & "/" & sMB_MEMFEE_SM & "/01") Then
                    com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "會費期間起日年月錯誤或未完整輸入!")
                    com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "會費期間起日年月錯誤或未完整輸入!")
                    Return
                End If

                Dim sMB_MEMFEE_EY As String = String.Empty
                If IsNumeric(Me.MB_MEMFEE_EY.Text) Then
                    sMB_MEMFEE_EY = CDec(Me.MB_MEMFEE_EY.Text)
                End If
                Dim sMB_MEMFEE_EM As String = String.Empty
                If IsNumeric(Me.MB_MEMFEE_EM.Text) Then
                    sMB_MEMFEE_EM = CDec(Me.MB_MEMFEE_EM.Text)
                End If
                If Not IsDate(sMB_MEMFEE_EY & "/" & sMB_MEMFEE_EM & "/01") Then
                    com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "會費期間迄日年月錯誤或未完整輸入!")
                    com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "會費期間迄日年月錯誤或未完整輸入!")
                    Return
                End If

                If CDate(sMB_MEMFEE_EY & "/" & sMB_MEMFEE_EM & "/01") < CDate(sMB_MEMFEE_SY & "/" & sMB_MEMFEE_SM & "/01") Then
                    com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "會費期間迄日年月不可小於起日年月!")
                    com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "會費期間迄日年月不可小於起日年月!")
                    Return
                End If
            End If

            If Not Utility.isValidateData(Me.MB_PRINT.SelectedValue) Then
                com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "請選擇是否開立收據!")
                com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請選擇是否開立收據!")
                Return
            End If

            If Me.MB_PRINT.SelectedValue = "1" Then
                If Not Utility.isValidateData(Me.MB_SEND_PRINT.SelectedValue) Then
                    com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "請選擇給收據方式!")
                    com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請選擇給收據方式!")
                    Return
                End If
            End If

            If Me.DDL_MB_ITEMID.SelectedValue = "B" Then
                If Not Utility.isValidateData(Me.PROJCODE.SelectedValue) Then
                    com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "請選擇專案名稱!")
                    com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請選擇專案名稱!")
                    Return
                End If
            End If

            Dim mbMEMREV As New MB_MEMREV(Me.m_DBManager)
            If mbMEMREV.loadByPK(Me.HID_MB_MEMSEQ.Value, Me.HID_MB_SEQNO.Value) Then
                '繳款日期
                Dim iMB_TX_DATE As Decimal = (CDec(Me.MB_TX_DATE_YYY.Text)) * 10000 + CDec(Me.MB_TX_DATE_MM.Text) * 100 + CDec(Me.MB_TX_DATE_DD.Text)
                mbMEMREV.setAttribute("MB_TX_DATE", iMB_TX_DATE)

                '功德項目
                mbMEMREV.setAttribute("MB_ITEMID", Me.DDL_MB_ITEMID.SelectedValue)

                '會員類別
                If Me.DDL_MB_ITEMID.SelectedValue = "A" Then
                    mbMEMREV.setAttribute("MB_MEMTYP", Me.RBL_MB_MEMTYP.SelectedValue)

                    mbMEMREV.setAttribute("PROJCODE", Nothing)
                Else
                    mbMEMREV.setAttribute("MB_MEMTYP", Nothing)

                    If Me.DDL_MB_ITEMID.SelectedValue = "B" Then
                        mbMEMREV.setAttribute("PROJCODE", Me.PROJCODE.SelectedValue)
                    Else
                        mbMEMREV.setAttribute("PROJCODE", Nothing)
                    End If
                End If

                '繳款方式
                mbMEMREV.setAttribute("MB_FEETYPE", Me.RBL_MB_FEETYPE.SelectedValue)

                '繳款金額
                If IsNumeric(Me.TXT_MB_TOTFEE.Text) Then
                    mbMEMREV.setAttribute("MB_TOTFEE", CDec(Me.TXT_MB_TOTFEE.Text))
                Else
                    mbMEMREV.setAttribute("MB_TOTFEE", Nothing)
                End If

                '每月金額
                Select Case Me.RBL_MB_FEETYPE.SelectedValue
                    Case "A", "B", "C"
                        If IsNumeric(Me.TXT_MB_TOTFEE_MM.Text) Then
                            mbMEMREV.setAttribute("MB_MONFEE", CDec(Me.TXT_MB_TOTFEE_MM.Text))
                        Else
                            mbMEMREV.setAttribute("MB_MONFEE", Nothing)
                        End If
                    Case Else
                        mbMEMREV.setAttribute("MB_MONFEE", Nothing)
                End Select

                '繳款分配項目
                If Me.FUND1.Checked Then
                    mbMEMREV.setAttribute("FUND1", "1")
                Else
                    mbMEMREV.setAttribute("FUND1", Nothing)
                End If
                If Me.FUND2.Checked Then
                    mbMEMREV.setAttribute("FUND2", "1")
                Else
                    mbMEMREV.setAttribute("FUND2", Nothing)
                End If
                If Me.FUND3.Checked Then
                    mbMEMREV.setAttribute("FUND3", "1")
                Else
                    mbMEMREV.setAttribute("FUND3", Nothing)
                End If

                '收據捐款名稱
                mbMEMREV.setAttribute("MB_RECNAME", Me.TXT_MB_RECNAME.Text)

                '所屬區
                ' mbMEMREV.setAttribute("MB_AREA", Me.DDL_MB_AREA.SelectedValue)

                '委員
                'mbMEMREV.setAttribute("MB_LEADER", Me.DDL_MB_LEADER.SelectedValue)
                Dim mbMEMBER As New MB_MEMBER(Me.m_DBManager)
                If mbMEMBER.loadByPK(Me.HID_MB_MEMSEQ.Value) Then
                    '所屬區
                    mbMEMBER.setAttribute("MB_AREA", Me.DDL_MB_AREA.SelectedValue)

                    '委員
                    mbMEMBER.setAttribute("MB_LEADER", Me.DDL_MB_LEADER.SelectedValue)

                    mbMEMBER.save()
                End If

                '會費期間 
                If IsNumeric(Me.MB_MEMFEE_SY.Text) AndAlso IsNumeric(Me.MB_MEMFEE_SM.Text) AndAlso IsNumeric(Me.MB_MEMFEE_EY.Text) AndAlso IsNumeric(Me.MB_MEMFEE_EM.Text) Then
                    '繳會費起年
                    mbMEMREV.setAttribute("MB_MEMFEE_SY", CDec(Me.MB_MEMFEE_SY.Text))
                    '繳會費起月
                    mbMEMREV.setAttribute("MB_MEMFEE_SM", CDec(Me.MB_MEMFEE_SM.Text))
                    '繳會費迄年
                    mbMEMREV.setAttribute("MB_MEMFEE_EY", CDec(Me.MB_MEMFEE_EY.Text))
                    '繳會費迄月
                    mbMEMREV.setAttribute("MB_MEMFEE_EM", CDec(Me.MB_MEMFEE_EM.Text))
                Else
                    '繳會費起年
                    mbMEMREV.setAttribute("MB_MEMFEE_SY", Nothing)
                    '繳會費起月
                    mbMEMREV.setAttribute("MB_MEMFEE_SM", Nothing)
                    '繳會費迄年
                    mbMEMREV.setAttribute("MB_MEMFEE_EY", Nothing)
                    '繳會費迄月
                    mbMEMREV.setAttribute("MB_MEMFEE_EM", Nothing)
                End If

                '是否開立收據
                mbMEMREV.setAttribute("MB_PRINT", Me.MB_PRINT.SelectedValue)

                '給收據方式
                If Me.MB_PRINT.SelectedValue = "1" Then
                    mbMEMREV.setAttribute("MB_SEND_PRINT", Me.MB_SEND_PRINT.SelectedValue)
                Else
                    mbMEMREV.setAttribute("MB_SEND_PRINT", Nothing)
                End If

                mbMEMREV.setAttribute("VRYUID", Session("UserId"))

                mbMEMREV.setAttribute("VRYDATE", Now)

                mbMEMREV.save()
            End If

            Me.bt_Qry_SUB_1_Click(sender, e)
            Me.PLH_MB_MEMREV_DATA.Visible = False
            Me.PLH_MB_MEMREV_LIST.Visible = True
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub
#End Region

#Region "過帳"
    Sub bt_Approve_Data_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles bt_Approve_Data.Click
        Try
            Try
                Me.m_DBManager.beginTran()

                Dim mbMEMREV As New MB_MEMREV(Me.m_DBManager)
                If mbMEMREV.loadByPK(CDec(Me.HID_MB_MEMSEQ.Value), CDec(Me.HID_MB_SEQNO.Value)) Then
                    '傳票號碼-西元年???
                    Dim iMB_VOUCHER_Y As Decimal = Now.Year
                    mbMEMREV.setAttribute("MB_VOUCHER_Y", iMB_VOUCHER_Y)

                    '傳票號碼???
                    Dim sProcName As String = String.Empty
                    sProcName = LCase(com.Azion.NET.VB.Properties.getSchemaName) & ".MB_GET_MB_GETNO"
                    Dim inParaAL As New ArrayList
                    Dim outParaAL As New ArrayList
                    inParaAL.Add(iMB_VOUCHER_Y)
                    inParaAL.Add("VC")

                    outParaAL.Add(6)

                    Dim HT_Return As Hashtable = com.Azion.NET.VB.DBObject.getProcedureData(Me.m_DBManager, sProcName, inParaAL, outParaAL)
                    Dim iMB_VOUCHER_N As Decimal = HT_Return.Item("@IMAXID")

                    mbMEMREV.setAttribute("MB_VOUCHER_N", iMB_VOUCHER_N)

                    '核准員工編號
                    mbMEMREV.setAttribute("APVUID", com.Azion.EloanUtility.UIUtility.getLoginUserID)

                    '核准日期時間
                    mbMEMREV.setAttribute("APVDATE", Now)

                    mbMEMREV.save()

                    If mbMEMREV.getString("MB_ITEMID") = "A" AndAlso mbMEMREV.getString("MB_MEMTYP") = "2" Then
                        Dim mbVIP As New MB_VIP(Me.m_DBManager)
                        If mbVIP.LoadByPK(CDec(Me.HID_MB_MEMSEQ.Value), "2") Then
                            'ACCAMT	decimal(6,0)		YES		0		select,insert,update,references	種子護法累計總繳金額
                            mbVIP.setAttribute("ACCAMT", mbVIP.getDecimal("ACCAMT") + mbMEMREV.getDecimal("MB_TOTFEE"))
                            'CHGUID	varchar(20)	utf8_general_ci	YES				select,insert,update,references	修改員工編號
                            mbVIP.setAttribute("CHGUID", Session("UserId"))
                            'CHGDATE	datetime		YES				select,insert,update,references	修改日期時間
                            mbVIP.setAttribute("CHGDATE", Now)
                            mbVIP.save()
                        End If
                    End If

                    '???
                    Me.INSERT_MB_ACCT_DTL(iMB_VOUCHER_Y, iMB_VOUCHER_N, mbMEMREV, "D", 1)
                    Me.INSERT_MB_ACCT_DTL(iMB_VOUCHER_Y, iMB_VOUCHER_N, mbMEMREV, "C", 1)

                    'MB_MEMFEE
                    If mbMEMREV.getString("MB_ITEMID") = "A" Then
                        Me.INSERT_MB_MEMFEE(CDec(HID_MB_MEMSEQ.Value), mbMEMREV.getDecimal("MB_MEMFEE_SY"), mbMEMREV)

                        If mbMEMREV.getDecimal("MB_MEMFEE_EY") > mbMEMREV.getDecimal("MB_MEMFEE_SY") Then
                            Me.INSERT_MB_MEMFEE(CDec(HID_MB_MEMSEQ.Value), mbMEMREV.getDecimal("MB_MEMFEE_EY"), mbMEMREV)
                        End If
                    End If
                End If

                Me.m_DBManager.commit()
            Catch ex As Exception
                Me.m_DBManager.Rollback()
                Throw
            End Try

            Dim isHaveNoData As Boolean = False

            'If Me.RP_QRY_SUB_1_1.Checked Then
            '    Dim mbMEMREVList As New MB_MEMREVList(Me.m_DBManager)
            '    mbMEMREVList.loadByMB_AREA_VOUCHER_Y_0(Me.MB_AREA_SUB_1.SelectedValue)

            '    If mbMEMREVList.getCurrentDataSet.Tables(0).Rows.Count = 0 Then
            '        com.Azion.EloanUtility.UIUtility.alert(Me.MB_AREA_SUB_1.SelectedItem.Text & "無未入帳資料!")

            '        isHaveNoData = True
            '    Else
            '        Me.RP_MB_MEMREV.DataSource = mbMEMREVList.getCurrentDataSet.Tables(0)
            '        Me.RP_MB_MEMREV.DataBind()
            '    End If
            'ElseIf Me.RP_QRY_SUB_1_2.Checked Then
            '    Dim mbMEMREVList As New MB_MEMREVList(Me.m_DBManager)
            '    mbMEMREVList.loadByMB_TX_DATE_VOUCHER_Y_0(Utility.FillZero(CDec(Me.txt_QRY_SUB_1_YYY.Text), 4) & Utility.FillZero(Me.txt_QRY_SUB_1_MM.Text, 2))

            '    If mbMEMREVList.getCurrentDataSet.Tables(0).Rows.Count = 0 Then
            '        com.Azion.EloanUtility.UIUtility.alert("該繳款年月" & "\n" & "西元" & Me.txt_QRY_SUB_1_YYY.Text & "年" & Me.txt_QRY_SUB_1_MM.Text & "月" & "\n" & "無未入帳資料!")

            '        isHaveNoData = True
            '    Else
            '        Me.RP_MB_MEMREV.DataSource = mbMEMREVList.getCurrentDataSet.Tables(0)
            '        Me.RP_MB_MEMREV.DataBind()
            '    End If
            'End If

            If Me.RP_QRY_SUB_1_1.Checked Then
                Dim mbMEMREVList As New MB_MEMREVList(Me.m_DBManager)
                If Me.RB_QRY_1.Checked Then
                    mbMEMREVList.loadByMB_AREA_1(Me.MB_AREA_SUB_1.SelectedValue)
                ElseIf Me.RB_QRY_1_2.Checked Then
                    mbMEMREVList.loadByMB_AREA_2(Me.MB_AREA_SUB_1.SelectedValue)
                ElseIf Me.RB_QRY_2.Checked Then
                    mbMEMREVList.loadByMB_AREA_3(Me.MB_AREA_SUB_1.SelectedValue)
                End If

                If mbMEMREVList.getCurrentDataSet.Tables(0).Rows.Count = 0 Then
                    If Me.RB_QRY_1.Checked Then
                        com.Azion.EloanUtility.UIUtility.alert(Me.MB_AREA_SUB_1.SelectedItem.Text & "該區無未覆核更正資料!")
                    ElseIf Me.RB_QRY_1_2.Checked Then
                        com.Azion.EloanUtility.UIUtility.alert(Me.MB_AREA_SUB_1.SelectedItem.Text & "該區無未入帳資料!")
                    ElseIf Me.RB_QRY_2.Checked Then
                        com.Azion.EloanUtility.UIUtility.alert(Me.MB_AREA_SUB_1.SelectedItem.Text & "該區無已入帳資料!")
                    End If
                Else
                    isHaveNoData = True
                    Me.RP_MB_MEMREV.DataSource = mbMEMREVList.getCurrentDataSet.Tables(0)
                    Me.RP_MB_MEMREV.DataBind()
                End If
            ElseIf Me.RP_QRY_SUB_1_2.Checked Then
                Dim mbMEMREVList As New MB_MEMREVList(Me.m_DBManager)
                If Me.RB_QRY_1.Checked Then
                    mbMEMREVList.loadByMB_TX_DATE_1(Utility.FillZero(CDec(Me.txt_QRY_SUB_1_YYY.Text), 4) & Utility.FillZero(Me.txt_QRY_SUB_1_MM.Text, 2), Session("AREA"))
                ElseIf Me.RB_QRY_1_2.Checked Then
                    mbMEMREVList.loadByMB_TX_DATE_2(Utility.FillZero(CDec(Me.txt_QRY_SUB_1_YYY.Text), 4) & Utility.FillZero(Me.txt_QRY_SUB_1_MM.Text, 2), Session("AREA"))
                ElseIf Me.RB_QRY_2.Checked Then
                    Dim sBYYYYMM As String = String.Empty
                    sBYYYYMM = Utility.FillZero(CDec(Me.txt_QRY_SUB_2_BEGYYY.Text), 4) & Utility.FillZero(Me.txt_QRY_SUB_2_BEGMM.Text, 2)
                    Dim sEYYYYMM As String = String.Empty
                    sEYYYYMM = Utility.FillZero(CDec(Me.txt_QRY_SUB_2_ENDYYY.Text), 4) & Utility.FillZero(Me.txt_QRY_SUB_2_ENDMM.Text, 2)
                    mbMEMREVList.loadByMB_TX_DATE_3(sBYYYYMM, sEYYYYMM, Session("AREA"))
                End If

                If mbMEMREVList.getCurrentDataSet.Tables(0).Rows.Count = 0 Then
                    If Me.RB_QRY_1.Checked Then
                        com.Azion.EloanUtility.UIUtility.alert("該繳款年月" & "\n" & "西元" & Me.txt_QRY_SUB_1_YYY.Text & "年" & Me.txt_QRY_SUB_1_MM.Text & "月" & "\n" & "無未覆核更正資料!")
                    ElseIf Me.RB_QRY_1_2.Checked Then
                        com.Azion.EloanUtility.UIUtility.alert("該繳款年月" & "\n" & "西元" & Me.txt_QRY_SUB_1_YYY.Text & "年" & Me.txt_QRY_SUB_1_MM.Text & "月" & "\n" & "無未入帳資料!")
                    ElseIf Me.RB_QRY_2.Checked Then
                        Dim sBYYYYMM As String = String.Empty
                        sBYYYYMM = Utility.FillZero(CDec(Me.txt_QRY_SUB_2_BEGYYY.Text), 4) & Utility.FillZero(Me.txt_QRY_SUB_2_BEGMM.Text, 2)
                        Dim sEYYYYMM As String = String.Empty
                        sEYYYYMM = Utility.FillZero(CDec(Me.txt_QRY_SUB_2_ENDYYY.Text), 4) & Utility.FillZero(Me.txt_QRY_SUB_2_ENDMM.Text, 2)

                        com.Azion.EloanUtility.UIUtility.alert("該繳款年月" & "\n" & "西元" & Me.txt_QRY_SUB_2_BEGYYY.Text & "年" & Me.txt_QRY_SUB_2_BEGMM.Text & "月~西元" & Me.txt_QRY_SUB_2_ENDYYY.Text & "年" & Me.txt_QRY_SUB_2_ENDMM.Text & "月" & "\n" & "無已入帳資料!")
                    End If
                Else
                    isHaveNoData = True
                    Me.RP_MB_MEMREV.DataSource = mbMEMREVList.getCurrentDataSet.Tables(0)
                    Me.RP_MB_MEMREV.DataBind()
                End If
            End If

            UIShareFun.showErrMsg(Me, "覆核成功!")

            Me.ClearMB_MEMREV_DATA()

            Me.PLH_MB_MEMREV_LIST.Visible = True

            If Not isHaveNoData Then
                Me.bt_Back_Click(Sender, e)
            End If
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    Sub INSERT_MB_MEMFEE(ByVal iMB_MEMSEQ As Decimal, ByVal iMB_YYYY As Decimal, ByVal mbMEMREV As MB_MEMREV)
        Try
            Dim mbMEMFEE As New MB_MEMFEE(Me.m_DBManager)
            mbMEMFEE.LoadByPK(iMB_MEMSEQ, iMB_YYYY)

            '會員編號
            mbMEMFEE.setAttribute("MB_MEMSEQ", iMB_MEMSEQ)

            '繳會費年度
            mbMEMFEE.setAttribute("MB_YYYY", iMB_YYYY)

            Me.setMAMT(mbMEMFEE, mbMEMREV, "1")
            Me.setMAMT(mbMEMFEE, mbMEMREV, "2")
            Me.setMAMT(mbMEMFEE, mbMEMREV, "3")
            Me.setMAMT(mbMEMFEE, mbMEMREV, "4")
            Me.setMAMT(mbMEMFEE, mbMEMREV, "5")
            Me.setMAMT(mbMEMFEE, mbMEMREV, "6")
            Me.setMAMT(mbMEMFEE, mbMEMREV, "7")
            Me.setMAMT(mbMEMFEE, mbMEMREV, "8")
            Me.setMAMT(mbMEMFEE, mbMEMREV, "9")
            Me.setMAMT(mbMEMFEE, mbMEMREV, "10")
            Me.setMAMT(mbMEMFEE, mbMEMREV, "11")
            Me.setMAMT(mbMEMFEE, mbMEMREV, "12")

            '核准員工編號
            mbMEMFEE.setAttribute("APVUID", com.Azion.EloanUtility.UIUtility.getLoginUserID)

            '核准日期時間
            mbMEMFEE.setAttribute("APVDATE", Now)

            mbMEMFEE.save()
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Sub setMAMT(ByVal mbMEMFEE As MB_MEMFEE, ByVal mbMEMREV As MB_MEMREV, ByVal sMM As String)
        Try
            If mbMEMREV.getString("MB_MEMTYP") = "1" Then
                Dim sColName As String = String.Empty
                sColName = "MB_M" & sMM

                Dim iAMT As Decimal = 0
                iAMT = mbMEMFEE.getDecimal(sColName)

                Dim iS_YYYYMM As Decimal = 0
                iS_YYYYMM = mbMEMREV.getDecimal("MB_MEMFEE_SY") * 100 + mbMEMREV.getDecimal("MB_MEMFEE_SM")

                Dim iE_YYYYMM As Decimal = 0
                iE_YYYYMM = mbMEMREV.getDecimal("MB_MEMFEE_EY") * 100 + mbMEMREV.getDecimal("MB_MEMFEE_EM")

                Dim iFEE_YYYYMM As Decimal = 0
                iFEE_YYYYMM = mbMEMFEE.getDecimal("MB_YYYY") * 100 + CInt(sMM)

                If iFEE_YYYYMM >= iS_YYYYMM AndAlso iFEE_YYYYMM <= iE_YYYYMM Then
                    iAMT += mbMEMREV.getDecimal("MB_MONFEE")

                    mbMEMFEE.setAttribute(sColName, iAMT)
                End If
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Sub INSERT_MB_ACCT_DTL(ByVal iMB_VOUCHER_Y As Decimal, ByVal iMB_VOUCHER_N As Decimal, ByVal mbMEMREV As MB_MEMREV, ByVal sMB_CD As String, ByVal iPlus As Decimal)
        Try
            Dim mbACCT_DTL As New MB_ACCT_DTL(Me.m_DBManager)
            Dim iMB_SEQ As Decimal = 0
            iMB_SEQ = mbACCT_DTL.getMAX_MB_SEQ(iMB_VOUCHER_Y, iMB_VOUCHER_N) + 1
            mbACCT_DTL.LoadByPK(iMB_VOUCHER_Y, iMB_VOUCHER_N, iMB_SEQ)

            'MB_VOUCHER_Y	decimal(4,0)		NO	PRI			select,insert,update,references	傳票號碼-西元年
            mbACCT_DTL.setAttribute("MB_VOUCHER_Y", iMB_VOUCHER_Y)

            'MB_VOUCHER_N	decimal(6,0)		NO	PRI			select,insert,update,references	傳票號碼
            mbACCT_DTL.setAttribute("MB_VOUCHER_N", iMB_VOUCHER_N)

            'MB_SEQ	decimal(3,0)		NO				select,insert,update,references	序號
            '??
            mbACCT_DTL.setAttribute("MB_SEQ", iMB_SEQ)

            'MB_CD	char(1)	latin1_swedish_ci	NO				select,insert,update,references	借貸別 C: 借D: 貸
            mbACCT_DTL.setAttribute("MB_CD", sMB_CD)

            'MB_ACTNO	varchar(4)	utf8_general_ci	YES				select,insert,update,references	主科目
            Dim sNOTE As String = String.Empty
            sNOTE = Me.get_MB_ITEMID_NOTE(mbMEMREV.getAttribute("MB_ITEMID"))
            Dim sTMP_NOTE() As String = Split(sNOTE, ";")
            If sMB_CD = "D" Then
                '借
                If mbMEMREV.getString("MB_PAY_TYPE") = "C" Then
                    '現金
                    mbACCT_DTL.setAttribute("MB_ACTNO", "1101")
                Else
                    'D票據
                    '應收票據
                    mbACCT_DTL.setAttribute("MB_ACTNO", "1121")
                End If
            Else
                If Not IsNothing(sTMP_NOTE) AndAlso sTMP_NOTE.Length >= 1 Then
                    mbACCT_DTL.setAttribute("MB_ACTNO", sTMP_NOTE(0))
                Else
                    mbACCT_DTL.setAttribute("MB_ACTNO", Nothing)
                End If
            End If

            'MB_SUBACT	varchar(4)	utf8_general_ci	YES				select,insert,update,references	子科目(戶號) MBSC銀行戶號(AP_CODE)

            'MB_ACTNAME	varchar(50)	utf8_general_ci	YES				select,insert,update,references	科目名稱
            If Not IsNothing(sTMP_NOTE) AndAlso sTMP_NOTE.Length >= 2 Then
                mbACCT_DTL.setAttribute("MB_ACTNAME", sTMP_NOTE(1))
            Else
                mbACCT_DTL.setAttribute("MB_ACTNAME", Nothing)
            End If

            'MB_AMT	decimal(9,0)		YES				select,insert,update,references	金額
            mbACCT_DTL.setAttribute("MB_AMT", mbMEMREV.getAttribute("MB_TOTFEE") * iPlus)

            'MB_DESC	varchar(100)	utf8_general_ci	YES				select,insert,update,references	說明
            mbACCT_DTL.setAttribute("MB_DESC", Me.getMB_ITEMID(mbMEMREV.getAttribute("MB_ITEMID")) & mbMEMREV.getAttribute("MB_DESC"))

            'DELFLAG	char(1)	latin1_swedish_ci	YES				select,insert,update,references	刪除註記
            mbACCT_DTL.setAttribute("DELFLAG", Nothing)

            'MB_TX_DATE	decimal(8,0)		YES				select,insert,update,references	繳款日期 YYYYMMDD
            mbACCT_DTL.setAttribute("MB_TX_DATE", mbMEMREV.getAttribute("MB_TX_DATE"))

            'CHGUID	varchar(100)	utf8_general_ci	YES				select,insert,update,references	修改員工編號
            mbACCT_DTL.setAttribute("CHGUID", com.Azion.EloanUtility.UIUtility.getLoginUserID)

            'CHGDATE	datetime		YES				select,insert,update,references	修改日期時間
            mbACCT_DTL.setAttribute("CHGUID", Now)

            mbACCT_DTL.save()
        Catch ex As Exception
            Throw
        End Try
    End Sub

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
#End Region

#Region "沖帳"
    'Sub bt_Acct_Data_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles bt_Acct_Data.Click
    '    Try
    '        Try
    '            Me.m_DBManager.beginTran()

    '            Dim mbMEMREV As New MB_MEMREV(Me.m_DBManager)
    '            If mbMEMREV.loadByPK(CDec(Me.HID_MB_MEMSEQ.Value), CDec(Me.HID_MB_SEQNO.Value)) Then
    '                mbMEMREV.setAttribute("DELFLAG", "D")
    '                mbMEMREV.setAttribute("CHGUID", com.Azion.EloanUtility.UIUtility.getLoginUserID)
    '                mbMEMREV.setAttribute("CHGDATE", Now)
    '                mbMEMREV.save()

    '                'MB_MEMREV
    '                Dim mbMEMFEE As New MB_MEMFEE(Me.m_DBManager)
    '                If mbMEMREV.getString("MB_FEETYPE") = "D" Then
    '                    mbMEMFEE.LoadByPK(CDec(Me.HID_MB_MEMSEQ.Value), CDec(Left(mbMEMREV.getString("MB_TX_DATE"), 4)))
    '                Else
    '                    mbMEMFEE.LoadByPK(CDec(Me.HID_MB_MEMSEQ.Value), mbMEMREV.getDecimal("MB_MEMFEE_SY"))
    '                End If

    '                If mbMEMFEE.isLoaded Then
    '                    Dim iMM_AMT As Decimal = 0
    '                    If mbMEMREV.getString("MB_FEETYPE") = "A" OrElse mbMEMREV.getString("MB_FEETYPE") = "B" OrElse _
    '                        mbMEMREV.getString("MB_FEETYPE") = "C" Then
    '                        Dim apCODEList As New AP_CODEList(Me.m_DBManager)
    '                        apCODEList.LoadbyUPCODEandVALUE(Me.m_sUpcode23, mbMEMREV.getString("MB_FEETYPE"))
    '                        If apCODEList.size > 0 AndAlso IsNumeric(apCODEList.item(0).getAttribute("NOTE")) Then
    '                            iMM_AMT = CDec(apCODEList.item(0).getAttribute("NOTE"))
    '                        End If
    '                    Else
    '                        iMM_AMT = mbMEMREV.getDecimal("MB_TOTFEE")
    '                    End If

    '                    Select Case mbMEMREV.getString("MB_FEETYPE")
    '                        Case "A"
    '                            '月
    '                            Dim sCOLNAME As String = String.Empty
    '                            sCOLNAME = "MB_M" & mbMEMREV.getString("MB_MEMFEE_SM")
    '                            Dim iAMT As Decimal = mbMEMFEE.getDecimal(sCOLNAME) - iMM_AMT
    '                            If iAMT < 0 Then iAMT = 0
    '                            mbMEMFEE.setAttribute(sCOLNAME, iAMT)
    '                            mbMEMFEE.setAttribute("APVUID", com.Azion.EloanUtility.UIUtility.getLoginUserID)
    '                            mbMEMFEE.setAttribute("APVDATE", Now)
    '                            mbMEMFEE.save()
    '                        Case "B"
    '                            '季
    '                            Dim D_BEG As Date = New Date(mbMEMREV.getDecimal("MB_MEMFEE_SY"), mbMEMREV.getDecimal("MB_MEMFEE_SM"), 1)
    '                            Dim D_END As Date = D_BEG.AddMonths(2)
    '                            If D_BEG.Year = D_END.Year Then
    '                                For i As Integer = D_BEG.Month To D_END.Month
    '                                    Dim sCOLNAME As String = String.Empty
    '                                    sCOLNAME = "MB_M" & i.ToString
    '                                    Dim iAMT As Decimal = mbMEMFEE.getDecimal(sCOLNAME) - iMM_AMT
    '                                    If iAMT < 0 Then iAMT = 0
    '                                    mbMEMFEE.setAttribute(sCOLNAME, iAMT)
    '                                Next

    '                                mbMEMFEE.setAttribute("APVUID", com.Azion.EloanUtility.UIUtility.getLoginUserID)
    '                                mbMEMFEE.setAttribute("APVDATE", Now)
    '                                mbMEMFEE.save()
    '                            Else
    '                                For i As Integer = D_BEG.Month To 12
    '                                    Dim sCOLNAME As String = String.Empty
    '                                    sCOLNAME = "MB_M" & i.ToString
    '                                    Dim iAMT As Decimal = mbMEMFEE.getDecimal(sCOLNAME) - iMM_AMT
    '                                    If iAMT < 0 Then iAMT = 0
    '                                    mbMEMFEE.setAttribute(sCOLNAME, iAMT)
    '                                Next

    '                                mbMEMFEE.setAttribute("APVUID", com.Azion.EloanUtility.UIUtility.getLoginUserID)
    '                                mbMEMFEE.setAttribute("APVDATE", Now)
    '                                mbMEMFEE.save()

    '                                mbMEMFEE.clear()
    '                                If mbMEMFEE.LoadByPK(CDec(Me.HID_MB_MEMSEQ.Value), D_END.Year) Then
    '                                    For i As Integer = 1 To D_END.Month
    '                                        Dim sCOLNAME As String = String.Empty
    '                                        sCOLNAME = "MB_M" & i.ToString
    '                                        Dim iAMT As Decimal = mbMEMFEE.getDecimal(sCOLNAME) - iMM_AMT
    '                                        If iAMT < 0 Then iAMT = 0
    '                                        mbMEMFEE.setAttribute(sCOLNAME, iAMT)
    '                                    Next

    '                                    mbMEMFEE.setAttribute("APVUID", com.Azion.EloanUtility.UIUtility.getLoginUserID)
    '                                    mbMEMFEE.setAttribute("APVDATE", Now)
    '                                    mbMEMFEE.save()
    '                                End If
    '                            End If
    '                        Case "C"
    '                            '年
    '                            Dim D_BEG As Date = New Date(mbMEMREV.getDecimal("MB_MEMFEE_SY"), mbMEMREV.getDecimal("MB_MEMFEE_SM"), 1)
    '                            Dim D_END As Date = D_BEG.AddMonths(12)
    '                            If D_BEG.Year = D_END.Year Then
    '                                For i As Integer = D_BEG.Month To D_END.Month
    '                                    Dim sCOLNAME As String = String.Empty
    '                                    sCOLNAME = "MB_M" & i.ToString
    '                                    Dim iAMT As Decimal = mbMEMFEE.getDecimal(sCOLNAME) - iMM_AMT
    '                                    If iAMT < 0 Then iAMT = 0
    '                                    mbMEMFEE.setAttribute(sCOLNAME, iAMT)
    '                                Next

    '                                mbMEMFEE.setAttribute("APVUID", com.Azion.EloanUtility.UIUtility.getLoginUserID)
    '                                mbMEMFEE.setAttribute("APVDATE", Now)
    '                                mbMEMFEE.save()
    '                            Else
    '                                For i As Integer = D_BEG.Month To 12
    '                                    Dim sCOLNAME As String = String.Empty
    '                                    sCOLNAME = "MB_M" & i.ToString
    '                                    Dim iAMT As Decimal = mbMEMFEE.getDecimal(sCOLNAME) - iMM_AMT
    '                                    If iAMT < 0 Then iAMT = 0
    '                                    mbMEMFEE.setAttribute(sCOLNAME, iAMT)
    '                                Next

    '                                mbMEMFEE.setAttribute("APVUID", com.Azion.EloanUtility.UIUtility.getLoginUserID)
    '                                mbMEMFEE.setAttribute("APVDATE", Now)
    '                                mbMEMFEE.save()

    '                                mbMEMFEE.clear()
    '                                If mbMEMFEE.LoadByPK(CDec(Me.HID_MB_MEMSEQ.Value), D_END.Year) Then
    '                                    For i As Integer = 1 To D_END.Month
    '                                        Dim sCOLNAME As String = String.Empty
    '                                        sCOLNAME = "MB_M" & i.ToString
    '                                        Dim iAMT As Decimal = mbMEMFEE.getDecimal(sCOLNAME) - iMM_AMT
    '                                        If iAMT < 0 Then iAMT = 0
    '                                        mbMEMFEE.setAttribute(sCOLNAME, iAMT)
    '                                    Next

    '                                    mbMEMFEE.setAttribute("APVUID", com.Azion.EloanUtility.UIUtility.getLoginUserID)
    '                                    mbMEMFEE.setAttribute("APVDATE", Now)
    '                                    mbMEMFEE.save()
    '                                End If
    '                            End If
    '                        Case "D"
    '                            '隨緣
    '                            Dim sCOLNAME As String = String.Empty
    '                            sCOLNAME = "MB_M" & mbMEMREV.getString("MB_TX_DATE").Substring(4, 2)
    '                            Dim iAMT As Decimal = mbMEMFEE.getDecimal(sCOLNAME) - iMM_AMT
    '                            If iAMT < 0 Then iAMT = 0
    '                            mbMEMFEE.setAttribute(sCOLNAME, iAMT)
    '                            mbMEMFEE.setAttribute("APVUID", com.Azion.EloanUtility.UIUtility.getLoginUserID)
    '                            mbMEMFEE.setAttribute("APVDATE", Now)
    '                            mbMEMFEE.save()
    '                    End Select
    '                End If

    '                If mbMEMREV.getString("MB_PAY_TYPE") = "C" Then
    '                    'Insert into MB_ACCT_DTL  (1筆 D: 借1筆 貸, 借->貸,貸->借 )
    '                    'copy 撈出之MB_ACCT_DTL 
    '                    'MB_CD C->D , D->C
    '                Else
    '                    '支票
    '                    If mbMEMREV.getString("NOTE_CASH") = "Y" Then
    '                        '已兌現
    '                    Else
    '                        '未兌現
    '                    End If
    '                End If
    '            End If

    '            Me.m_DBManager.commit()
    '        Catch ex As Exception
    '            Me.m_DBManager.Rollback()
    '            Throw
    '        End Try

    '        Dim isHaveNoData As Boolean = False
    '        If Me.RP_QRY_SUB_2_1.Checked Then
    '            Dim mbMEMREVList As New MB_MEMREVList(Me.m_DBManager)
    '            mbMEMREVList.loadByMB_AREA_VOUCHER_Y(Me.MB_AREA_SUB_2.SelectedValue)

    '            If mbMEMREVList.getCurrentDataSet.Tables(0).Rows.Count = 0 Then
    '                com.Azion.EloanUtility.UIUtility.alert(Me.MB_AREA_SUB_2.SelectedItem.Text & "無已入帳資料!")

    '                isHaveNoData = True
    '            Else
    '                Me.RP_MB_MEMREV.DataSource = mbMEMREVList.getCurrentDataSet.Tables(0)
    '                Me.RP_MB_MEMREV.DataBind()
    '            End If
    '        ElseIf Me.RP_QRY_SUB_2_2.Checked Then
    '            Dim mbMEMREVList As New MB_MEMREVList(Me.m_DBManager)
    '            Dim iBEG As Decimal = (CDec(Me.txt_QRY_SUB_2_BEGYYY.Text)) * 100 + CDec(Me.txt_QRY_SUB_2_BEGMM.Text)
    '            Dim iEND As Decimal = (CDec(Me.txt_QRY_SUB_2_ENDYYY.Text)) * 100 + CDec(Me.txt_QRY_SUB_2_ENDMM.Text)
    '            mbMEMREVList.loadByMB_TX_DATE_VOUCHER_Y(iBEG, iEND)

    '            If mbMEMREVList.getCurrentDataSet.Tables(0).Rows.Count = 0 Then
    '                Dim sMsg As String = String.Empty
    '                sMsg = "西元" & Me.txt_QRY_SUB_2_BEGYYY.Text & "年" & Me.txt_QRY_SUB_2_BEGMM.Text & "月" & _
    '                       "至" & "西元" & Me.txt_QRY_SUB_2_ENDYYY.Text & "年" & Me.txt_QRY_SUB_2_ENDMM.Text & "月"
    '                com.Azion.EloanUtility.UIUtility.alert("該繳款年月" & "\n" & sMsg & "\n" & "無已入帳資料!")

    '                isHaveNoData = True
    '            Else
    '                Me.RP_MB_MEMREV.DataSource = mbMEMREVList.getCurrentDataSet.Tables(0)
    '                Me.RP_MB_MEMREV.DataBind()
    '            End If
    '        ElseIf Me.RP_QRY_SUB_2_3.Checked Then
    '            Dim mbMEMREVList As New MB_MEMREVList(Me.m_DBManager)
    '            mbMEMREVList.loadByMB_NAME_VOUCHER_Y(Trim(Me.txt_QRY_MB_NAME.Text))

    '            If mbMEMREVList.getCurrentDataSet.Tables(0).Rows.Count = 0 Then
    '                com.Azion.EloanUtility.UIUtility.alert("無已入帳資料!")

    '                isHaveNoData = True
    '            Else
    '                Me.RP_MB_MEMREV.DataSource = mbMEMREVList.getCurrentDataSet.Tables(0)
    '                Me.RP_MB_MEMREV.DataBind()
    '            End If
    '        End If

    '        UIShareFun.showErrMsg(Me, "沖帳成功!")

    '        Me.PLH_MB_MEMREV.Visible = True
    '        Me.PLH_QRY.Visible = False

    '        Me.TD_PRINT.Visible = True

    '        If isHaveNoData Then
    '            Me.bt_Back_Click(Sender, e)
    '        End If
    '    Catch ex As Exception
    '        UIShareFun.showErrMsg(Me, ex)
    '    End Try
    'End Sub

    Sub bt_Acct_Data_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles bt_Acct_Data.Click
        Try
            Try
                Me.m_DBManager.beginTran()

                Dim mbMEMREV As New MB_MEMREV(Me.m_DBManager)
                If mbMEMREV.loadByPK(CDec(Me.HID_MB_MEMSEQ.Value), CDec(Me.HID_MB_SEQNO.Value)) Then
                    mbMEMREV.setAttribute("DELFLAG", "D")
                    mbMEMREV.setAttribute("CHGUID", com.Azion.EloanUtility.UIUtility.getLoginUserID)
                    mbMEMREV.setAttribute("CHGDATE", Now)
                    mbMEMREV.save()

                    Dim mbMEMFEE As New MB_MEMFEE(Me.m_DBManager)
                    If mbMEMFEE.LoadByPK(CDec(Me.HID_MB_MEMSEQ.Value), mbMEMFEE.getDecimal("MB_MEMFEE_SY")) Then
                        mbMEMFEE.setAttribute("MB_M1", 0)
                        mbMEMFEE.setAttribute("MB_M2", 0)
                        mbMEMFEE.setAttribute("MB_M3", 0)
                        mbMEMFEE.setAttribute("MB_M4", 0)
                        mbMEMFEE.setAttribute("MB_M5", 0)
                        mbMEMFEE.setAttribute("MB_M6", 0)
                        mbMEMFEE.setAttribute("MB_M7", 0)
                        mbMEMFEE.setAttribute("MB_M8", 0)
                        mbMEMFEE.setAttribute("MB_M9", 0)
                        mbMEMFEE.setAttribute("MB_M10", 0)
                        mbMEMFEE.setAttribute("MB_M11", 0)
                        mbMEMFEE.setAttribute("MB_M12", 0)
                        mbMEMFEE.save()
                    End If

                    '???
                    If mbMEMREV.getString("MB_PAY_TYPE") = "C" Then
                        Me.COPY_MB_ACCT_DTL(mbMEMREV, False)
                    Else
                        '支票
                        If mbMEMREV.getString("NOTE_CASH") = "Y" Then
                            '已兌現
                            Me.COPY_MB_ACCT_DTL(mbMEMREV, True)
                        Else
                            '未兌現
                            Me.COPY_MB_ACCT_DTL(mbMEMREV, False)
                        End If
                    End If
                End If

                Me.m_DBManager.commit()
            Catch ex As Exception
                Me.m_DBManager.Rollback()
                Throw
            End Try

            Dim isHaveNoData As Boolean = False

            If Me.RP_QRY_SUB_2_1.Checked Then
                Dim mbMEMREVList As New MB_MEMREVList(Me.m_DBManager)
                mbMEMREVList.loadByMB_AREA_3(Me.MB_AREA_SUB_2.SelectedValue)

                If mbMEMREVList.getCurrentDataSet.Tables(0).Rows.Count = 0 Then
                    com.Azion.EloanUtility.UIUtility.alert(Me.MB_AREA_SUB_2.SelectedItem.Text & "無已入帳資料!")
                    isHaveNoData = True
                    Return
                Else
                    Me.RP_MB_MEMREV.DataSource = mbMEMREVList.getCurrentDataSet.Tables(0)
                    Me.RP_MB_MEMREV.DataBind()
                End If
            ElseIf Me.RP_QRY_SUB_2_2.Checked Then
                Dim mbMEMREVList As New MB_MEMREVList(Me.m_DBManager)
                Dim iBEG As Decimal = (CDec(Me.txt_QRY_SUB_2_BEGYYY.Text)) * 100 + CDec(Me.txt_QRY_SUB_2_BEGMM.Text)
                Dim iEND As Decimal = (CDec(Me.txt_QRY_SUB_2_ENDYYY.Text)) * 100 + CDec(Me.txt_QRY_SUB_2_ENDMM.Text)
                mbMEMREVList.loadByMB_TX_DATE_3(iBEG, iEND, Session("AREA"))

                If mbMEMREVList.getCurrentDataSet.Tables(0).Rows.Count = 0 Then
                    Dim sMsg As String = String.Empty
                    sMsg = "西元" & Me.txt_QRY_SUB_2_BEGYYY.Text & "年" & Me.txt_QRY_SUB_2_BEGMM.Text & "月" & _
                           "至" & "西元" & Me.txt_QRY_SUB_2_ENDYYY.Text & "年" & Me.txt_QRY_SUB_2_ENDMM.Text & "月"
                    com.Azion.EloanUtility.UIUtility.alert("該繳款年月" & "\n" & sMsg & "\n" & "無已入帳資料!")
                    isHaveNoData = True
                    Return
                Else
                    Me.RP_MB_MEMREV.DataSource = mbMEMREVList.getCurrentDataSet.Tables(0)
                    Me.RP_MB_MEMREV.DataBind()
                End If
            ElseIf Me.RP_QRY_SUB_2_3.Checked Then
                Dim mbMEMREVList As New MB_MEMREVList(Me.m_DBManager)
                mbMEMREVList.loadByMB_NAME_VOUCHER_Y(Trim(Me.txt_QRY_MB_NAME.Text))

                If mbMEMREVList.getCurrentDataSet.Tables(0).Rows.Count = 0 Then
                    com.Azion.EloanUtility.UIUtility.alert("無已入帳資料!")
                    isHaveNoData = True
                    Return
                Else
                    Me.RP_MB_MEMREV.DataSource = mbMEMREVList.getCurrentDataSet.Tables(0)
                    Me.RP_MB_MEMREV.DataBind()
                End If
            End If

            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "沖帳成功")

            Me.ClearMB_MEMREV_DATA()

            Me.PLH_MB_MEMREV_LIST.Visible = True

            If isHaveNoData Then
                Me.bt_Back_Click(Sender, e)
            End If
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

    Sub COPY_MB_ACCT_DTL(ByVal mbMEMREV As MB_MEMREV, ByVal isRevACTNO As Boolean)
        Try
            Dim mbACCT_DTLList As New MB_ACCT_DTLList(Me.m_DBManager)
            mbACCT_DTLList.LoadByVOUCHER(mbMEMREV.getDecimal("MB_VOUCHER_Y"), mbMEMREV.getDecimal("MB_VOUCHER_N"))

            Dim mbACCT_DTL As New MB_ACCT_DTL(Me.m_DBManager)
            Dim mbSEQ As New MB_ACCT_DTL(Me.m_DBManager)
            For i As Integer = 0 To mbACCT_DTLList.size - 1
                Dim objDTL As MB_ACCT_DTL = CType(mbACCT_DTLList.item(i), MB_ACCT_DTL)

                mbACCT_DTL.clear()

                '傳票號碼???
                Dim iMB_VOUCHER_Y As Decimal = Now.Year
                Dim sProcName As String = String.Empty
                sProcName = LCase(com.Azion.NET.VB.Properties.getSchemaName) & ".MB_GET_MB_GETNO"
                Dim inParaAL As New ArrayList
                Dim outParaAL As New ArrayList
                inParaAL.Add(iMB_VOUCHER_Y)
                inParaAL.Add("VC")
                outParaAL.Add(6)
                Dim HT_Return As Hashtable = com.Azion.NET.VB.DBObject.getProcedureData(Me.m_DBManager, sProcName, inParaAL, outParaAL)
                Dim iMB_VOUCHER_N As Decimal = HT_Return.Item("@IMAXID")

                'MB_VOUCHER_Y	decimal(4,0)		NO	PRI			select,insert,update,references	傳票號碼-西元年
                mbACCT_DTL.setAttribute("MB_VOUCHER_Y", iMB_VOUCHER_Y)

                'MB_VOUCHER_N	decimal(6,0)		NO	PRI			select,insert,update,references	傳票號碼
                mbACCT_DTL.setAttribute("MB_VOUCHER_N", iMB_VOUCHER_N)

                'MB_SEQ	decimal(3,0)		NO	PRI			select,insert,update,references	序號
                mbSEQ.clear()
                Dim iMB_SEQ As Integer = 0
                iMB_SEQ = mbSEQ.getMAX_MB_SEQ(iMB_VOUCHER_Y, iMB_VOUCHER_N) + 1
                mbACCT_DTL.setAttribute("MB_SEQ", iMB_SEQ)

                'MB_CD	char(1)	latin1_swedish_ci	NO				select,insert,update,references	借貸別 C: 借D: 貸
                '???
                'MB_CD -> C -> D , D -> C,MB_AMT * -1,MB_TX_DATE = system date )
                'If objDTL.getString("MB_CD") = "D" Then
                '    mbACCT_DTL.setAttribute("MB_CD", "C")
                'Else
                '    mbACCT_DTL.setAttribute("MB_CD", "D")
                'End If
                mbACCT_DTL.setAttribute("MB_CD", objDTL.getAttribute("MB_CD"))

                'MB_ACTNO	varchar(4)	utf8_general_ci	YES				select,insert,update,references	主科目
                If Not isRevACTNO Then
                    mbACCT_DTL.setAttribute("MB_ACTNO", objDTL.getAttribute("MB_ACTNO"))
                Else
                    Dim sNOTE As String = String.Empty
                    sNOTE = Me.get_MB_ITEMID_NOTE(mbMEMREV.getAttribute("MB_ITEMID"))
                    Dim sTMP_NOTE() As String = Split(sNOTE, ";")
                    If mbACCT_DTL.getString("MB_CD") = "D" Then
                        If Not IsNothing(sTMP_NOTE) AndAlso sTMP_NOTE.Length >= 1 Then
                            mbACCT_DTL.setAttribute("MB_ACTNO", sTMP_NOTE(0))
                        Else
                            mbACCT_DTL.setAttribute("MB_ACTNO", Nothing)
                        End If
                    Else
                        '借
                        If mbMEMREV.getString("MB_PAY_TYPE") = "C" Then
                            '現金
                            mbACCT_DTL.setAttribute("MB_ACTNO", "1101")
                        Else
                            'D票據
                            '應收票據
                            mbACCT_DTL.setAttribute("MB_ACTNO", "1121")
                        End If
                    End If
                End If

                'MB_SUBACT	varchar(4)	utf8_general_ci	YES				select,insert,update,references	子科目(戶號) MBSC銀行戶號(AP_CODE)
                mbACCT_DTL.setAttribute("MB_SUBACT", objDTL.getAttribute("MB_SUBACT"))

                'MB_AMT	decimal(9,0)		YES				select,insert,update,references	金額
                mbACCT_DTL.setAttribute("MB_AMT", objDTL.getDecimal("MB_AMT") * -1)

                'MB_DESC	varchar(100)	utf8_general_ci	YES				select,insert,update,references	說明
                mbACCT_DTL.setAttribute("MB_DESC", objDTL.getAttribute("MB_DESC"))

                'CHGUID	varchar(100)	utf8_general_ci	YES				select,insert,update,references	修改員工編號
                mbACCT_DTL.setAttribute("CHGUID", com.Azion.EloanUtility.UIUtility.getLoginUserID)

                'MB_CURCODE	varchar(3)	utf8_general_ci	YES				select,insert,update,references	幣別
                mbACCT_DTL.setAttribute("MB_CURCODE", objDTL.getAttribute("MB_CURCODE"))

                'MB_TX_DATE	decimal(8,0)		YES				select,insert,update,references	繳款日期 YYYYMMDD
                mbACCT_DTL.setAttribute("MB_TX_DATE", Now)

                'MB_ACTNAME	varchar(50)	utf8_general_ci	YES				select,insert,update,references	科目名稱
                mbACCT_DTL.setAttribute("MB_ACTNAME", objDTL.getAttribute("MB_ACTNAME"))

                'DELFLAG	char(1)	latin1_swedish_ci	YES				select,insert,update,references	刪除註記
                mbACCT_DTL.setAttribute("DELFLAG", objDTL.getAttribute("DELFLAG"))

                mbACCT_DTL.save()
            Next
        Catch ex As Exception
            Throw
        End Try
    End Sub
#End Region

    Sub SH_MODIFY(ByVal isLTL As Boolean, ByVal isEDIT As Boolean)
        Try
            '繳款日期
            Me.MB_TX_DATE.Visible = isLTL
            Me.PLH_MB_TX_DATE.Visible = isEDIT

            '功德項目 
            Me.MB_ITEMID.Visible = isLTL
            Me.DDL_MB_ITEMID.Visible = isEDIT

            '會員類別
            Me.MB_MEMTYP.Visible = isLTL
            Me.RBL_MB_MEMTYP.Visible = isEDIT

            '繳款方式
            Me.MB_FEETYPE.Visible = isLTL
            Me.RBL_MB_FEETYPE.Visible = isEDIT

            '繳款金額
            Me.MB_TOTFEE.Visible = isLTL
            Me.TXT_MB_TOTFEE.Visible = isEDIT
            Me.MB_TOTFEE_MM.Visible = isLTL
            Me.TXT_MB_TOTFEE_MM.Visible = isEDIT

            '收據捐款名稱
            Me.MB_RECNAME.Visible = isLTL
            Me.TXT_MB_RECNAME.Visible = isEDIT

            '所屬區
            Me.MB_AREA.Visible = isLTL
            Me.DDL_MB_AREA.Visible = isEDIT

            '委員
            Me.MB_LEADER.Visible = isLTL
            Me.DDL_MB_LEADER.Visible = isEDIT

            '會費期間 
            Me.PLH_SHOW_MB_MEMFEE.Visible = isLTL
            Me.PLH_MB_MEMFEE.Visible = isEDIT
        Catch ex As Exception
            Throw
        End Try
    End Sub

End Class