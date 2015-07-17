Imports MBSC.MB_OP
Imports com.Azion.NET.VB
Imports MBSC.UICtl

Public Class MBMnt_Member_01_v01
    Inherits System.Web.UI.Page

    Dim m_DBManager As com.Azion.NET.VB.DatabaseManager

    Dim m_sMB_VLG As String = String.Empty
    Dim m_sMB_VLG1 As String = String.Empty
    Dim m_sUpCode1 As String = String.Empty
    Private Shared m_sMB_LEADER As String = String.Empty
    Private Shared m_sUpCode7 As String = String.Empty
    Dim m_sUpCode28 As String = String.Empty

    Dim m_sUpcode15 As String = String.Empty
    Private Shared m_sUpcode23 As String = String.Empty
    Dim m_sUpcode31 As String = String.Empty
    Dim m_sUpcode272 As String = String.Empty
    Dim m_sUpcode276 As String = String.Empty
    Dim m_sUpcode76 As String = String.Empty

    Dim m_sOPTYPE As String = String.Empty
    Dim DT_UPCODE15 As DataTable

    Dim m_sMB_MEMSEQ As String = String.Empty

#Region "Page Events"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            'A法友會會員—入會申請單
            'B繳款作業
            Me.m_sOPTYPE = "" & Request.QueryString("OPTYPE")

            Me.m_sMB_MEMSEQ = "" & Request.QueryString("MB_MEMSEQ")

            'Test
            'Me.m_sOPTYPE = "B"

            Me.m_sMB_VLG = "" & Request.Form("MB_VLG")

            Me.m_sMB_VLG1 = "" & Request.Form("MB_VLG1")

            m_sMB_LEADER = "" & Request.Form("MB_LEADER")

            '學歷
            Me.m_sUpCode1 = com.Azion.EloanUtility.CodeList.getAppSettings("UpCode1")

            '所屬區代碼
            m_sUpCode7 = com.Azion.EloanUtility.CodeList.getAppSettings("UpCode7")

            '會員類別
            Me.m_sUpCode28 = com.Azion.EloanUtility.CodeList.getAppSettings("UpCode28")

            '功德項目
            Me.m_sUpcode15 = com.Azion.EloanUtility.CodeList.getAppSettings("UpCode15")

            '繳款方式
            m_sUpcode23 = com.Azion.EloanUtility.CodeList.getAppSettings("UpCode23")

            '付款方式
            Me.m_sUpcode31 = com.Azion.EloanUtility.CodeList.getAppSettings("UpCode31")

            '給收據方式
            Me.m_sUpcode272 = com.Azion.EloanUtility.CodeList.getAppSettings("UpCode272")

            '專案名稱
            Me.m_sUpcode276 = com.Azion.EloanUtility.CodeList.getAppSettings("UpCode276")

            '維護人員
            Me.m_sUpcode76 = com.Azion.EloanUtility.CodeList.getAppSettings("UpCode76")

            Me.m_DBManager = UIShareFun.getDataBaseManager

            If Not Page.IsPostBack Then
                If Not Me.is76User(com.Azion.EloanUtility.UIUtility.getLoginUserID) Then
                    '非維護人員關閉依手機，電話，姓名搜尋
                    Me.PLH_QRY.Visible = False

                    Dim mbMEMBERList As New MB_MEMBERList(Me.m_DBManager)
                    mbMEMBERList.loadByMB_EMAIL(com.Azion.EloanUtility.UIUtility.getLoginUserID)
                    If mbMEMBERList.getCurrentDataSet.Tables(0).Rows.Count > 0 Then
                        Me.PLH_MAIL_SAME.Visible = True
                        Me.RP_MAIL_SAME.DataSource = mbMEMBERList.getCurrentDataSet.Tables(0)
                        Me.RP_MAIL_SAME.DataBind()
                    Else
                        Me.PLH_MAIL_SAME.Visible = False
                        Me.PLH_MB_MEMBER.Visible = True
                        Me.MB_EMAIL.Text = com.Azion.EloanUtility.UIUtility.getLoginUserID
                        Me.HID_OTHSIGN.Value = "2"
                    End If
                End If

                If Me.m_sOPTYPE = "A" Then
                    Me.LTL_PROGTITLE.Text = "法友會會員—入會申請單"
                ElseIf Me.m_sOPTYPE = "B" Then
                    Me.LTL_PROGTITLE.Text = "繳款作業"

                    '繼續繳款作業
                    Me.bt_ProcMB_MEMREV.Visible = True

                    '功德項目
                    Me.Bind_MB_ITEMID()

                    '繳款方式
                    Me.Bind_MB_FEETYPE("0")

                    '付款方式
                    Me.Bind_MB_PAY_TYPE()

                    '給收據方式
                    Me.Bind_MB_SEND_PRINT()

                    '專案名稱
                    Me.Bind_PROJCODE()

                    'Me.MB_TOTFEE_MM.Attributes.Add("Readonly", "Readonly")
                    Me.MB_TOTFEE.Attributes.Add("Readonly", "Readonly")
                End If

                Me.Bind_DDL_City()

                Me.Bind_MB_EDU()

                Me.Bind_MB_AREA()

                Me.Bind_MB_FAMILY()

                If Me.m_sOPTYPE = "A" AndAlso IsNumeric(Me.m_sMB_MEMSEQ) Then
                    Me.HID_MB_MEMSEQ.Value = Me.m_sMB_MEMSEQ
                    Me.Bind_MB_MEMBER(CDec(Me.m_sMB_MEMSEQ))
                    Me.PLH_QRY.Visible = False
                    Me.PLH_BUTTON.Visible = False
                    Me.PLH_PAGETAB.Visible = False
                End If
            Else
                If Utility.isValidateData(Me.MB_CITY.SelectedValue) Then
                    Me.ReBind_DDL_Town_PostBack(Me.MB_VLG, Me.MB_CITY.SelectedItem.Text, Me.m_sMB_VLG)
                End If

                If Utility.isValidateData(Me.MB_CITY1.SelectedValue) Then
                    Me.ReBind_DDL_Town_PostBack(Me.MB_VLG1, Me.MB_CITY1.SelectedItem.Text, Me.m_sMB_VLG1)
                End If

                If Utility.isValidateData(Me.MB_AREA.SelectedValue) Then
                    Me.ReBind_MB_LEADER(Me.MB_AREA.SelectedValue, m_sMB_LEADER)
                End If
            End If
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    Function is76User(ByVal sUserId As String) As Boolean
        Try
            Dim apCODEList As New AP_CODEList(Me.m_DBManager)
            apCODEList.loadByUpCode(Me.m_sUpcode76)
            Dim ROW_USER() As DataRow = Nothing
            ROW_USER = apCODEList.getCurrentDataSet.Tables(0).Select("TEXT='" & sUserId & "'")
            If Not IsNothing(ROW_USER) AndAlso ROW_USER.Length > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function

    Private Sub MBMnt_Member_01_v01_Unload(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Unload
        Try
            UIShareFun.releaseConnection(Me.m_DBManager)
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub
#End Region

#Region "Button Events"
    Sub bt_Qry_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles bt_Qry.Click
        Dim DT_MB_MEMBER As DataTable = Nothing
        Try
            If Me.RB_QRY_1.Checked = False AndAlso Me.RB_QRY_2.Checked = False AndAlso Me.RB_QRY_3.Checked = False Then
                com.Azion.EloanUtility.UIUtility.alert("請選擇輸入方式!")

                Return
            End If

            If Me.RB_QRY_1.Checked AndAlso Not Utility.isValidateData(Trim(Me.txt_QRY_Mobile.Text)) Then
                com.Azion.EloanUtility.UIUtility.alert("請輸入手機號碼!")

                Return
            End If

            If Me.RB_QRY_2.Checked AndAlso Not Utility.isValidateData(Trim(Me.txt_QRY_Phone_Pre.Text)) Then
                com.Azion.EloanUtility.UIUtility.alert("請輸入電話區碼!")

                Return
            End If

            If Me.RB_QRY_2.Checked AndAlso Not Utility.isValidateData(Trim(Me.txt_QRY_Phone.Text)) Then
                com.Azion.EloanUtility.UIUtility.alert("請輸入電話號碼!")

                Return
            End If

            If Me.RB_QRY_3.Checked AndAlso Not Utility.isValidateData(Trim(Me.txt_QRY_MB_NAME.Text)) Then
                com.Azion.EloanUtility.UIUtility.alert("請輸入姓名!")

                Return
            End If

            Dim mbMEMBERList As New MB_MEMBERList(Me.m_DBManager)

            Dim sMB_TEL As String = String.Empty
            If Me.RB_QRY_1.Checked Then
                mbMEMBERList.loadByMB_MOBIL(Trim(Me.txt_QRY_Mobile.Text))
            ElseIf Me.RB_QRY_2.Checked Then
                If Utility.isValidateData(Trim(Me.txt_QRY_Phone_Pre.Text)) Then
                    sMB_TEL = Trim(Me.txt_QRY_Phone_Pre.Text) & "-" & Trim(Me.txt_QRY_Phone.Text)
                Else
                    sMB_TEL = Trim(Me.txt_QRY_Phone.Text)
                End If

                mbMEMBERList.loadByMB_TEL(sMB_TEL)
            ElseIf Me.RB_QRY_3.Checked Then
                mbMEMBERList.loadByMB_NAME(Trim(Me.txt_QRY_MB_NAME.Text))
                If mbMEMBERList.getCurrentDataSet.Tables(0).Rows.Count <> 1 Then
                    mbMEMBERList.clear()
                    mbMEMBERList.loadByMB_NAME_Like(Trim(Me.txt_QRY_MB_NAME.Text))
                End If
            End If

            DT_MB_MEMBER = mbMEMBERList.getCurrentDataSet.Tables(0)

            Me.Clear_MB_MEMBER()

            Me.Clear_MB_MEMREV()

            Me.PLH_QRY.Visible = False
            Me.PLH_QRY_SAME.Visible = False
            Me.PLH_MB_MEMBER.Visible = False
            Me.PLH_MB_MEMREV.Visible = False

            If DT_MB_MEMBER.Rows.Count = 0 Then
                If Me.RB_QRY_1.Checked Then
                    Me.MB_MOBIL.Text = Trim(Me.txt_QRY_Mobile.Text)
                ElseIf Me.RB_QRY_2.Checked Then
                    Me.MB_TEL_Pre.Text = Trim(Me.txt_QRY_Phone_Pre.Text)
                    Me.MB_TEL.Text = Trim(Me.txt_QRY_Phone.Text)
                ElseIf Me.RB_QRY_3.Checked Then
                    Me.MB_NAME.Text = Trim(Me.txt_QRY_MB_NAME.Text)
                End If

                Me.PLH_MB_MEMBER.Visible = True
            ElseIf DT_MB_MEMBER.Rows.Count = 1 Then
                Me.HID_MB_MEMSEQ.Value = com.Azion.EloanUtility.StrUtility.FillZero(DT_MB_MEMBER.Rows(0)("MB_MEMSEQ").ToString, 7)

                If Me.m_sOPTYPE = "A" Then
                    Me.Bind_MB_MEMBER(CDec(Me.HID_MB_MEMSEQ.Value))
                Else
                    Me.Bind_MB_MEMREV_Default(CDec(Me.HID_MB_MEMSEQ.Value))
                End If
            Else
                Me.PLH_QRY_SAME.Visible = True
                Me.RP_QRY_SAME.DataSource = DT_MB_MEMBER
                Me.RP_QRY_SAME.DataBind()
            End If
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        Finally
            If Not IsNothing(DT_MB_MEMBER) Then DT_MB_MEMBER.Dispose()
        End Try
    End Sub

    Sub bt_Qry_NO_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles bt_Qry_NO.Click
        Try
            Me.Clear_MB_MEMBER()

            If Me.RB_QRY_1.Checked Then
                Me.MB_MOBIL.Text = Trim(Me.txt_QRY_Mobile.Text)
            ElseIf Me.RB_QRY_2.Checked Then
                Me.MB_TEL_Pre.Text = Trim(Me.txt_QRY_Phone_Pre.Text)
                Me.MB_TEL.Text = Trim(Me.txt_QRY_Phone.Text)
            ElseIf Me.RB_QRY_3.Checked Then
                Me.MB_NAME.Text = Trim(Me.txt_QRY_MB_NAME.Text)
            End If

            Me.PLH_MB_MEMBER.Visible = True
            Me.PLH_QRY_SAME.Visible = False
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    Sub RP_QRY_SAME_ItemCommand(ByVal sender As Object, ByVal e As RepeaterCommandEventArgs) Handles RP_QRY_SAME.ItemCommand
        Try
            Me.PLH_QRY_SAME.Visible = False

            Me.PLH_MB_MEMBER.Visible = True

            Me.HID_MB_MEMSEQ.Value = com.Azion.EloanUtility.StrUtility.FillZero(e.CommandArgument, 7)

            If Me.m_sOPTYPE = "A" Then
                Me.Bind_MB_MEMBER(e.CommandArgument)
            Else
                Me.Bind_MB_MEMREV_Default(e.CommandArgument)
            End If
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    Sub Bind_MB_MEMBER(ByVal iMB_MEMSEQ As Decimal)
        Try
            Me.PLH_MB_MEMBER.Visible = True
            Me.PLH_MB_MEMREV.Visible = False

            Me.HID_MB_MEMSEQ.Value = com.Azion.EloanUtility.StrUtility.FillZero(iMB_MEMSEQ.ToString, 7)

            Dim mbMEMBER As New MB_MEMBER(Me.m_DBManager)
            If mbMEMBER.loadByPK(iMB_MEMSEQ) Then
                '法名/ 姓名
                Me.MB_NAME.Text = Trim(mbMEMBER.getString("MB_NAME"))

                '性別
                Me.MB_SEX.SelectedIndex = -1
                If Not IsNothing(Me.MB_SEX.Items.FindByValue(mbMEMBER.getString("MB_SEX"))) Then
                    Me.MB_SEX.Items.FindByValue(mbMEMBER.getString("MB_SEX")).Selected = True
                End If

                '出生年月日
                If Utility.isValidateData(mbMEMBER.getAttribute("MB_BIRTH")) AndAlso IsDate(mbMEMBER.getAttribute("MB_BIRTH").ToString()) AndAlso CDate(mbMEMBER.getAttribute("MB_BIRTH").ToString()).Year > 1911 Then
                    Me.MB_BIRTH_YYY.Text = CDate(mbMEMBER.getAttribute("MB_BIRTH").ToString()).Year
                    Me.MB_BIRTH_MM.Text = CDate(mbMEMBER.getAttribute("MB_BIRTH").ToString()).Month
                    Me.MB_BIRTH_DD.Text = CDate(mbMEMBER.getAttribute("MB_BIRTH").ToString()).Day
                Else
                    Me.MB_BIRTH_YYY.Text = String.Empty
                    Me.MB_BIRTH_MM.Text = String.Empty
                    Me.MB_BIRTH_DD.Text = String.Empty
                End If

                '身分別
                Me.MB_IDENTIFY.SelectedIndex = -1
                If Utility.isValidateData(mbMEMBER.getString("MB_IDENTIFY")) Then
                    Dim sMB_IDENTIFY() As String = Split(mbMEMBER.getString("MB_IDENTIFY"), ",")
                    If Not IsNothing(sMB_IDENTIFY) AndAlso sMB_IDENTIFY.Length > 0 Then
                        For i As Integer = 0 To UBound(sMB_IDENTIFY)
                            If Not IsNothing(Me.MB_IDENTIFY.Items.FindByValue(sMB_IDENTIFY(i))) Then
                                Me.MB_IDENTIFY.Items.FindByValue(sMB_IDENTIFY(i)).Selected = True
                            End If
                        Next
                    End If
                End If

                '手機
                Me.MB_MOBIL.Text = Trim(mbMEMBER.getString("MB_MOBIL"))

                '電話
                If Utility.isValidateData(Trim(mbMEMBER.getString("MB_TEL"))) Then
                    If InStr(Trim(mbMEMBER.getString("MB_TEL")), "-") > 0 Then
                        Dim sMB_TEL() As String = Split(Trim(mbMEMBER.getString("MB_TEL")), "-")
                        If Not IsNothing(sMB_TEL) AndAlso sMB_TEL.Length >= 1 Then
                            Me.MB_TEL_Pre.Text = Trim(sMB_TEL(0))
                            Me.MB_TEL.Text = Trim(sMB_TEL(1))
                        End If
                    Else
                        Me.MB_TEL_Pre.Text = String.Empty
                        Me.MB_TEL.Text = Trim(mbMEMBER.getString("MB_TEL"))
                    End If
                Else
                    Me.MB_TEL_Pre.Text = String.Empty
                    Me.MB_TEL.Text = String.Empty
                End If

                'e-mail
                Me.MB_EMAIL.Text = Trim(mbMEMBER.getString("MB_EMAIL"))

                '身分證字號
                Me.MB_ID.Text = Trim(mbMEMBER.getString("MB_ID"))

                '學歷
                Me.MB_EDU.SelectedIndex = -1
                If Not IsNothing(Me.MB_EDU.Items.FindByValue(mbMEMBER.getString("MB_EDU"))) Then
                    Me.MB_EDU.Items.FindByValue(mbMEMBER.getString("MB_EDU")).Selected = True
                End If

                '所屬區
                Me.MB_AREA.SelectedIndex = -1
                If Not IsNothing(Me.MB_AREA.Items.FindByValue(mbMEMBER.getString("MB_AREA"))) Then
                    Me.MB_AREA.Items.FindByValue(mbMEMBER.getString("MB_AREA")).Selected = True

                    '委員
                    Me.ReBind_MB_LEADER(mbMEMBER.getString("MB_AREA"), mbMEMBER.getString("MB_LEADER"))

                    Me.HID_MB_LEADER.Value = mbMEMBER.getString("MB_LEADER")
                End If

                '通訊地址
                Me.MB_CITY.SelectedIndex = -1
                If Not IsNothing(Me.MB_CITY.Items.FindByText(mbMEMBER.getString("MB_CITY"))) Then
                    Me.MB_CITY.Items.FindByText(mbMEMBER.getString("MB_CITY")).Selected = True

                    Me.ReBind_DDL_Town_PostBackByText(Me.MB_VLG, mbMEMBER.getString("MB_CITY"), mbMEMBER.getString("MB_VLG"))
                End If

                Me.MB_ADDR.Text = Trim(mbMEMBER.getString("MB_ADDR"))
                '通訊郵遞區號
                Me.TXT_MB_ZIP.Text = Trim(mbMEMBER.getString("MB_ZIP"))

                '戶籍地址
                Me.MB_CITY1.SelectedIndex = -1
                If Not IsNothing(Me.MB_CITY1.Items.FindByText(mbMEMBER.getString("MB_CITY1"))) Then
                    Me.MB_CITY1.Items.FindByText(mbMEMBER.getString("MB_CITY1")).Selected = True

                    Me.ReBind_DDL_Town_PostBackByText(Me.MB_VLG1, mbMEMBER.getString("MB_CITY1"), mbMEMBER.getString("MB_VLG1"))
                End If

                Me.MB_ADDR2.Text = Trim(mbMEMBER.getString("MB_ADDR1"))
                '戶籍郵遞區號
                Me.TXT_MB_ZIP1.Text = Trim(mbMEMBER.getString("MB_ZIP1"))

                '同上
                Me.MB_ADDR2_SAME.Checked = False
                If Utility.isValidateData(Me.MB_CITY.SelectedValue) AndAlso Utility.isValidateData(Me.MB_VLG.SelectedValue) AndAlso Utility.isValidateData(Me.MB_ADDR.Text) AndAlso _
                    Utility.isValidateData(Me.MB_CITY1.SelectedValue) AndAlso Utility.isValidateData(Me.MB_VLG1.SelectedValue) AndAlso Utility.isValidateData(Me.MB_ADDR2.Text) Then

                    If Me.MB_CITY.SelectedValue = Me.MB_CITY1.SelectedValue Then
                        If Me.MB_VLG.SelectedValue = Me.MB_VLG1.SelectedValue Then
                            If Me.MB_ADDR.Text = Me.MB_ADDR2.Text Then
                                Me.MB_ADDR2_SAME.Checked = True
                            End If
                        End If
                    End If

                End If

                '會員類別
                Me.MB_FAMILY.SelectedIndex = -1
                If Not IsNothing(Me.MB_FAMILY.Items.FindByValue(mbMEMBER.getString("MB_FAMILY"))) Then
                    Me.MB_FAMILY.Items.FindByValue(mbMEMBER.getString("MB_FAMILY")).Selected = True
                End If

                '會員編號
                Me.LTL_MB_MEMSEQ.Text = getMB_MEMSEQ(iMB_MEMSEQ, mbMEMBER.getString("MB_AREA"))

                '護持會員
                Dim mbVIPList As New MB_VIPList(Me.m_DBManager)
                mbVIPList.LoadByMB_MEMSEQ(iMB_MEMSEQ)
                For i As Integer = 0 To Me.VIP.Items.Count - 1
                    Dim sVIP As String = String.Empty
                    sVIP = Me.VIP.Items(i).Value
                    Dim ROW_VIP() As DataRow = mbVIPList.getCurrentDataSet.Tables(0).Select("VIP='" & sVIP & "'")
                    If Not IsNothing(ROW_VIP) AndAlso ROW_VIP.Length > 0 Then
                        Me.VIP.Items(i).Selected = True
                    End If
                Next

                '家族人員
                If Me.is76User(com.Azion.EloanUtility.UIUtility.getLoginUserID) Then
                    Me.TR_FAMILY.Visible = True

                    Dim mbFAMILYList As New MB_FAMILYList(Me.m_DBManager)
                    mbFAMILYList.LoadByMB_MEMSEQ(CDec(Me.HID_MB_MEMSEQ.Value))
                    Me.RP_MB_FAMILY.DataSource = mbFAMILYList.getCurrentDataSet.Tables(0)
                    Me.RP_MB_FAMILY.DataBind()
                End If
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Sub bt_Save_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles bt_Save.Click
        Try
            If Not Utility.isValidateData(Trim(Me.MB_NAME.Text)) Then
                com.Azion.EloanUtility.UIUtility.alert("請輸入法名/姓名!")
                Return
            End If

            If Not Utility.isValidateData(Me.MB_SEX.SelectedValue) Then
                com.Azion.EloanUtility.UIUtility.alert("請選擇性別!")
                Return
            End If

            If Not Utility.isValidateData(Trim(Me.MB_MOBIL.Text)) AndAlso Not Utility.isValidateData(Trim(Me.MB_TEL.Text)) Then
                com.Azion.EloanUtility.UIUtility.alert("請輸入手機或電話!")
                Return
            End If

            If Not Utility.isValidateData(Trim(Me.MB_CITY.SelectedValue)) Then
                com.Azion.EloanUtility.UIUtility.alert("請選擇通訊地址(市)!")
                Return
            End If

            If Not Utility.isValidateData(Trim(Me.MB_VLG.SelectedValue)) Then
                com.Azion.EloanUtility.UIUtility.alert("請選擇通訊地址(區)!")
                Return
            End If

            If Not Utility.isValidateData(Trim(Me.MB_ADDR.Text)) Then
                com.Azion.EloanUtility.UIUtility.alert("請輸入通訊地址!")
                Return
            End If

            If Not Utility.isValidateData(Trim(Me.MB_CITY1.SelectedValue)) Then
                com.Azion.EloanUtility.UIUtility.alert("請選擇戶籍地址(市)!")
                Return
            End If

            If Not Utility.isValidateData(Trim(Me.MB_VLG1.SelectedValue)) Then
                com.Azion.EloanUtility.UIUtility.alert("請選擇戶籍地址(區)!")
                Return
            End If

            If Not Utility.isValidateData(Trim(Me.MB_ADDR2.Text)) Then
                com.Azion.EloanUtility.UIUtility.alert("請輸入戶籍地址!")
                Return
            End If

            If Not Utility.isValidateData(Trim(Me.MB_AREA.SelectedValue)) Then
                com.Azion.EloanUtility.UIUtility.alert("請選擇所屬區!")
                Return
            End If

            If Not Utility.isValidateData(Trim(Me.MB_LEADER.SelectedValue)) Then
                com.Azion.EloanUtility.UIUtility.alert("請選擇所屬區委員!")
                Return
            End If

            '1040120
            '護持會員不檢核必勾選
            'If Not Utility.isValidateData(Trim(Me.VIP.SelectedValue)) Then
            '    com.Azion.EloanUtility.UIUtility.alert("請選擇護持會員!")
            '    Return
            'End If

            Dim sMB_IDENTIFY As String = String.Empty
            For i As Integer = 0 To Me.MB_IDENTIFY.Items.Count - 1
                If Me.MB_IDENTIFY.Items(i).Selected Then
                    sMB_IDENTIFY &= Me.MB_IDENTIFY.Items(i).Value & ","
                End If
            Next
            If Utility.isValidateData(sMB_IDENTIFY) Then
                sMB_IDENTIFY = Left(sMB_IDENTIFY, sMB_IDENTIFY.Length - 1)
            Else
                com.Azion.EloanUtility.UIUtility.alert("請選擇身分別!")
                Return
            End If

            If Not Utility.isValidateData(Trim(Me.MB_FAMILY.SelectedValue)) Then
                com.Azion.EloanUtility.UIUtility.alert("請選擇會員類別!")
                Return
            End If

            If Utility.isValidateData(Me.MB_BIRTH_YYY.Text) OrElse Utility.isValidateData(Me.MB_BIRTH_MM.Text) OrElse Utility.isValidateData(Me.MB_BIRTH_DD.Text) Then
                Dim sYYYY As String = String.Empty
                If IsNumeric(Me.MB_BIRTH_YYY.Text) Then
                    sYYYY = com.Azion.EloanUtility.StrUtility.FillZero(CDec(Me.MB_BIRTH_YYY.Text), 4)
                End If

                Dim sMM As String = String.Empty
                If IsNumeric(Me.MB_BIRTH_MM.Text) Then
                    sMM = com.Azion.EloanUtility.StrUtility.FillZero(CDec(Me.MB_BIRTH_MM.Text), 2)
                End If

                Dim sDD As String = String.Empty
                If IsNumeric(Me.MB_BIRTH_DD.Text) Then
                    sDD = com.Azion.EloanUtility.StrUtility.FillZero(CDec(Me.MB_BIRTH_DD.Text), 2)
                End If

                If Not IsDate(sYYYY & "/" & sMM & "/" & sDD) Then
                    com.Azion.EloanUtility.UIUtility.alert("出生年月日未完整輸入或輸入錯誤!")
                    Return
                End If
            End If

            If Not Utility.isValidateData(Me.HID_MB_MEMSEQ.Value) Then
                Dim MB_MEMBERList As New MB_MEMBERList(m_DBManager)
                If Utility.isValidateData(Me.MB_NAME.Text.Trim) AndAlso Utility.isValidateData(Me.MB_MOBIL.Text.Trim) Then
                    MB_MEMBERList.Load_MB_NAME_MB_MOBIL(Me.MB_NAME.Text.Trim, Me.MB_MOBIL.Text.Trim)
                    If MB_MEMBERList.getCurrentDataSet.Tables(0).Rows.Count > 0 Then
                        com.Azion.EloanUtility.UIUtility.showJSMsg("【" & Me.MB_NAME.Text.Trim & "】【" & Me.MB_MOBIL.Text.Trim & "】已經是會員了，請再確認看看")
                        com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "【" & Me.MB_NAME.Text.Trim & "】【" & Me.MB_MOBIL.Text.Trim & "】已經是會員了，請再確認看看")

                        Return
                    End If
                End If

                If Utility.isValidateData(Me.MB_NAME.Text.Trim) AndAlso Utility.isValidateData(Me.MB_TEL.Text.Trim) Then
                    MB_MEMBERList.clear()
                    MB_MEMBERList.Load_MB_NAME_MB_TEL(Me.MB_NAME.Text.Trim, Me.MB_TEL.Text.Trim)
                    If MB_MEMBERList.getCurrentDataSet.Tables(0).Rows.Count > 0 Then
                        com.Azion.EloanUtility.UIUtility.showJSMsg("【" & Me.MB_NAME.Text.Trim & "】【" & Me.MB_TEL.Text.Trim & "】已經是會員了，請再確認看看")
                        com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "【" & Me.MB_NAME.Text.Trim & "】【" & Me.MB_TEL.Text.Trim & "】已經是會員了，請再確認看看")

                        Return
                    End If
                End If

                If Utility.isValidateData(Me.MB_NAME.Text.Trim) AndAlso Utility.isValidateData(Me.MB_EMAIL.Text.Trim) Then
                    MB_MEMBERList.clear()
                    MB_MEMBERList.Load_MB_NAME_MB_EMAIL(Me.MB_NAME.Text.Trim, Me.MB_EMAIL.Text.Trim)
                    If MB_MEMBERList.getCurrentDataSet.Tables(0).Rows.Count > 0 Then
                        com.Azion.EloanUtility.UIUtility.showJSMsg("【" & Me.MB_NAME.Text.Trim & "】【" & Me.MB_EMAIL.Text.Trim & "】已經是會員了，請再確認看看")
                        com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "【" & Me.MB_NAME.Text.Trim & "】【" & Me.MB_EMAIL.Text.Trim & "】已經是會員了，請再確認看看")

                        Return
                    End If
                End If
            End If
            Try
                Me.m_DBManager.beginTran()

                If Me.HID_OTHSIGN.Value = "1" Then
                    Me.SAVE_MB_MEMBER_TEMP(sMB_IDENTIFY)
                Else
                    Me.SAVE_MB_MEMBER(sMB_IDENTIFY)

                    Dim mbFAMILY As New MB_FAMILY(Me.m_DBManager)
                    mbFAMILY.LoadByPK(CDec(Me.HID_MB_MEMSEQ.Value), CDec(Me.HID_MB_MEMSEQ.Value))
                    mbFAMILY.setAttribute("MB_MEMSEQ", CDec(Me.HID_MB_MEMSEQ.Value))
                    mbFAMILY.setAttribute("MB_FAMSEQ", CDec(Me.HID_MB_MEMSEQ.Value))
                    mbFAMILY.save()
                End If

                Me.m_DBManager.commit()

                UIShareFun.showErrMsg(Me, "儲存成功!")

                Me.LTL_MB_MEMSEQ.Text = Me.MB_AREA.SelectedValue & "-" & Me.HID_MB_MEMSEQ.Value

                If Me.HID_OTHSIGN.Value = "1" Then
                    com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "儲存成功!")
                    Me.bt_Back_Click(Sender, e)
                ElseIf Me.HID_OTHSIGN.Value = "2" Then
                    com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "儲存成功!")
                    Me.bt_Back_Click(Sender, e)
                    Dim mbMEMBERList As New MB_MEMBERList(Me.m_DBManager)
                    mbMEMBERList.loadByMB_EMAIL(com.Azion.EloanUtility.UIUtility.getLoginUserID)
                    If mbMEMBERList.getCurrentDataSet.Tables(0).Rows.Count > 0 Then
                        Me.PLH_MAIL_SAME.Visible = True
                        Me.RP_MAIL_SAME.DataSource = mbMEMBERList.getCurrentDataSet.Tables(0)
                        Me.RP_MAIL_SAME.DataBind()
                    End If
                End If
            Catch ex As Exception
                Me.m_DBManager.Rollback()
                Throw
            End Try
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    Sub SAVE_MB_MEMBER(ByVal sMB_IDENTIFY As String)
        Try
            If Not Utility.isValidateData(Me.HID_MB_MEMSEQ.Value) Then
                Dim sProcName As String = String.Empty
                sProcName = LCase(com.Azion.NET.VB.Properties.getSchemaName) & ".MB_GET_MAXID"
                Dim inParaAL As New ArrayList
                Dim outParaAL As New ArrayList
                inParaAL.Add("01")
                inParaAL.Add("1")

                outParaAL.Add(7)

                Dim HT_Return As Hashtable = com.Azion.NET.VB.DBObject.getProcedureData(Me.m_DBManager, sProcName, inParaAL, outParaAL)
                Dim iMAXID As Decimal = HT_Return.Item("@IMAXID")

                Me.HID_MB_MEMSEQ.Value = com.Azion.EloanUtility.StrUtility.FillZero(iMAXID, 7)
            End If

            Dim mbMEMBER As New MB_MEMBER(Me.m_DBManager)
            mbMEMBER.loadByPK(CDec(Me.HID_MB_MEMSEQ.Value))

            '建檔日期
            If Not mbMEMBER.isLoaded Then
                mbMEMBER.setAttribute("CREDATE", Now)
            End If

            '會員編號
            mbMEMBER.setAttribute("MB_MEMSEQ", CDec(Me.HID_MB_MEMSEQ.Value))

            '法名/ 姓名
            mbMEMBER.setAttribute("MB_NAME", Trim(Me.MB_NAME.Text))

            '性別
            If Utility.isValidateData(Me.MB_SEX.SelectedValue) Then
                mbMEMBER.setAttribute("MB_SEX", Me.MB_SEX.SelectedValue)
            Else
                mbMEMBER.setAttribute("MB_SEX", Nothing)
            End If

            '出生年月日
            Dim sYYYY As String = String.Empty
            If IsNumeric(Me.MB_BIRTH_YYY.Text) Then
                sYYYY = com.Azion.EloanUtility.StrUtility.FillZero(CDec(Me.MB_BIRTH_YYY.Text), 4)
            End If

            Dim sMM As String = String.Empty
            If IsNumeric(Me.MB_BIRTH_MM.Text) Then
                sMM = com.Azion.EloanUtility.StrUtility.FillZero(CDec(Me.MB_BIRTH_MM.Text), 2)
            End If

            Dim sDD As String = String.Empty
            If IsNumeric(Me.MB_BIRTH_DD.Text) Then
                sDD = com.Azion.EloanUtility.StrUtility.FillZero(CDec(Me.MB_BIRTH_DD.Text), 2)
            End If
            If IsDate(sYYYY & "/" & sMM & "/" & sDD) Then
                mbMEMBER.setAttribute("MB_BIRTH", CDate(sYYYY & "/" & sMM & "/" & sDD))
            Else
                mbMEMBER.setAttribute("MB_BIRTH", Nothing)
            End If

            '身分別
            mbMEMBER.setAttribute("MB_IDENTIFY", sMB_IDENTIFY)

            '手機
            If Utility.isValidateData(Trim(Me.MB_MOBIL.Text)) Then
                mbMEMBER.setAttribute("MB_MOBIL", Trim(Me.MB_MOBIL.Text))
            Else
                mbMEMBER.setAttribute("MB_MOBIL", Nothing)
            End If

            '電話
            If Utility.isValidateData(Trim(Me.MB_TEL_Pre.Text)) OrElse Utility.isValidateData(Trim(Me.MB_TEL.Text)) Then
                Dim sMB_TEL As String = String.Empty
                If Utility.isValidateData(Trim(Me.MB_TEL_Pre.Text)) Then
                    sMB_TEL = Trim(Me.MB_TEL_Pre.Text) & "-" & Trim(Me.MB_TEL.Text)
                Else
                    sMB_TEL = Trim(Me.MB_TEL.Text)
                End If
                mbMEMBER.setAttribute("MB_TEL", sMB_TEL)
            Else
                mbMEMBER.setAttribute("MB_TEL", Nothing)
            End If

            'e-mail
            If Utility.isValidateData(Trim(Me.MB_EMAIL.Text)) Then
                mbMEMBER.setAttribute("MB_EMAIL", Trim(Me.MB_EMAIL.Text))
            Else
                mbMEMBER.setAttribute("MB_EMAIL", Nothing)
            End If

            '身分證字號
            If Utility.isValidateData(Trim(Me.MB_ID.Text)) Then
                mbMEMBER.setAttribute("MB_ID", Trim(Me.MB_ID.Text))
            Else
                mbMEMBER.setAttribute("MB_ID", Nothing)
            End If

            '學歷
            If Utility.isValidateData(Trim(Me.MB_EDU.SelectedValue)) Then
                mbMEMBER.setAttribute("MB_EDU", Trim(Me.MB_EDU.SelectedValue))
            Else
                mbMEMBER.setAttribute("MB_EDU", Nothing)
            End If

            '所屬區
            mbMEMBER.setAttribute("MB_AREA", Me.MB_AREA.SelectedValue)

            '委員
            mbMEMBER.setAttribute("MB_LEADER", Me.MB_LEADER.SelectedValue)

            '通訊地址
            mbMEMBER.setAttribute("MB_CITY", Me.MB_CITY.SelectedItem.Text)
            mbMEMBER.setAttribute("MB_VLG", Me.MB_VLG.SelectedItem.Text)
            mbMEMBER.setAttribute("MB_ADDR", Trim(Me.MB_ADDR.Text))
            '通訊郵遞區號
            mbMEMBER.setAttribute("MB_ZIP", Trim(Me.TXT_MB_ZIP.Text))

            '戶籍地址
            If Utility.isValidateData(Me.MB_CITY1.SelectedValue) Then
                mbMEMBER.setAttribute("MB_CITY1", Me.MB_CITY1.SelectedItem.Text)
            Else
                mbMEMBER.setAttribute("MB_CITY1", Nothing)
            End If
            If Utility.isValidateData(Me.MB_VLG1.SelectedValue) Then
                mbMEMBER.setAttribute("MB_VLG1", Me.MB_VLG1.SelectedItem.Text)
            Else
                mbMEMBER.setAttribute("MB_VLG1", Nothing)
            End If
            If Utility.isValidateData(Trim(Me.MB_ADDR2.Text)) Then
                mbMEMBER.setAttribute("MB_ADDR1", Trim(Me.MB_ADDR2.Text))
            Else
                mbMEMBER.setAttribute("MB_ADDR1", Nothing)
            End If
            '戶籍郵遞區號
            mbMEMBER.setAttribute("MB_ZIP1", Trim(Me.TXT_MB_ZIP1.Text))

            '會員類別
            mbMEMBER.setAttribute("MB_FAMILY", Me.MB_FAMILY.SelectedValue)

            '修改員工編號
            mbMEMBER.setAttribute("CHGUID", com.Azion.EloanUtility.UIUtility.getLoginUserID)

            '修改日期
            mbMEMBER.setAttribute("CHGDATE", Now)

            mbMEMBER.save()

            Dim mbVIP As New MB_VIP(Me.m_DBManager)
            For i As Integer = 0 To Me.VIP.Items.Count - 1
                mbVIP.clear()
                mbVIP.LoadByPK(CDec(Me.HID_MB_MEMSEQ.Value), Me.VIP.Items(i).Value)
                If Me.VIP.Items(i).Selected Then
                    'MB_MEMSEQ	decimal(7,0)		NO	PRI	0		select,insert,update,references	會員編號
                    mbVIP.setAttribute("MB_MEMSEQ", CDec(Me.HID_MB_MEMSEQ.Value))
                    'VIP	char(1)	latin1_swedish_ci	NO	PRI			select,insert,update,references	護持會員;1長期護持會員;2種子護法
                    mbVIP.setAttribute("VIP", Me.VIP.Items(i).Value)
                    'ACCAMT	decimal(6,0)		YES		0		select,insert,update,references	種子護法累計總繳金額
                    'CHGUID	varchar(20)	utf8_general_ci	YES				select,insert,update,references	修改員工編號
                    mbVIP.setAttribute("CHGUID", Session("UserId"))
                    'CHGDATE	datetime		YES				select,insert,update,references	修改日期時間
                    mbVIP.setAttribute("CHGDATE", Now)
                    mbVIP.save()
                Else
                    If mbVIP.isLoaded Then
                        mbVIP.remove()
                    End If
                End If
            Next
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Sub SAVE_MB_MEMBER_TEMP(ByVal sMB_IDENTIFY As String)
        Try
            Dim iMB_DATE As Long = 0
            Dim D_NOW As Date = Now
            iMB_DATE = CType((D_NOW.Year), Long) * 10000000000000
            iMB_DATE += CType(D_NOW.Month, Long) * 100000000000
            iMB_DATE += CType(D_NOW.Day, Long) * 1000000000
            iMB_DATE += CType(D_NOW.Hour, Long) * 10000000
            iMB_DATE += CType(D_NOW.Minute, Long) * 100000
            iMB_DATE += CType(D_NOW.Second, Long) * 1000
            iMB_DATE += CType(D_NOW.Millisecond, Long)
            Dim mbMEMBER_TEMP As New MB_MEMBER_TEMP(Me.m_DBManager)
            mbMEMBER_TEMP.LoadByPK(iMB_DATE)

            '建檔日期
            If Not mbMEMBER_TEMP.isLoaded Then
                mbMEMBER_TEMP.setAttribute("CHGDATE", Now)
            End If

            '會員編號
            mbMEMBER_TEMP.setAttribute("MB_DATE", iMB_DATE)

            '法名/ 姓名
            mbMEMBER_TEMP.setAttribute("MB_NAME", Trim(Me.MB_NAME.Text))

            '性別
            If Utility.isValidateData(Me.MB_SEX.SelectedValue) Then
                mbMEMBER_TEMP.setAttribute("MB_SEX", Me.MB_SEX.SelectedValue)
            Else
                mbMEMBER_TEMP.setAttribute("MB_SEX", Nothing)
            End If

            '出生年月日
            Dim sYYYY As String = String.Empty
            If IsNumeric(Me.MB_BIRTH_YYY.Text) Then
                sYYYY = com.Azion.EloanUtility.StrUtility.FillZero(CDec(Me.MB_BIRTH_YYY.Text), 4)
            End If

            Dim sMM As String = String.Empty
            If IsNumeric(Me.MB_BIRTH_MM.Text) Then
                sMM = com.Azion.EloanUtility.StrUtility.FillZero(CDec(Me.MB_BIRTH_MM.Text), 2)
            End If

            Dim sDD As String = String.Empty
            If IsNumeric(Me.MB_BIRTH_DD.Text) Then
                sDD = com.Azion.EloanUtility.StrUtility.FillZero(CDec(Me.MB_BIRTH_DD.Text), 2)
            End If
            If IsDate(sYYYY & "/" & sMM & "/" & sDD) Then
                mbMEMBER_TEMP.setAttribute("MB_BIRTH", CDate(sYYYY & "/" & sMM & "/" & sDD))
            Else
                mbMEMBER_TEMP.setAttribute("MB_BIRTH", Nothing)
            End If

            '身分別
            mbMEMBER_TEMP.setAttribute("MB_IDENTIFY", sMB_IDENTIFY)

            '手機
            If Utility.isValidateData(Trim(Me.MB_MOBIL.Text)) Then
                mbMEMBER_TEMP.setAttribute("MB_MOBIL", Trim(Me.MB_MOBIL.Text))
            Else
                mbMEMBER_TEMP.setAttribute("MB_MOBIL", Nothing)
            End If

            '電話
            If Utility.isValidateData(Trim(Me.MB_TEL_Pre.Text)) OrElse Utility.isValidateData(Trim(Me.MB_TEL.Text)) Then
                Dim sMB_TEL As String = String.Empty
                If Utility.isValidateData(Trim(Me.MB_TEL_Pre.Text)) Then
                    sMB_TEL = Trim(Me.MB_TEL_Pre.Text) & "-" & Trim(Me.MB_TEL.Text)
                Else
                    sMB_TEL = Trim(Me.MB_TEL.Text)
                End If
                mbMEMBER_TEMP.setAttribute("MB_TEL", sMB_TEL)
            Else
                mbMEMBER_TEMP.setAttribute("MB_TEL", Nothing)
            End If

            'e-mail
            If Utility.isValidateData(Trim(Me.MB_EMAIL.Text)) Then
                mbMEMBER_TEMP.setAttribute("MB_EMAIL", Trim(Me.MB_EMAIL.Text))
            ElseIf Utility.isValidateData(com.Azion.EloanUtility.UIUtility.getLoginUserID) Then
                mbMEMBER_TEMP.setAttribute("MB_EMAIL", com.Azion.EloanUtility.UIUtility.getLoginUserID)
            Else
                mbMEMBER_TEMP.setAttribute("MB_EMAIL", Nothing)
            End If

            '身分證字號
            If Utility.isValidateData(Trim(Me.MB_ID.Text)) Then
                mbMEMBER_TEMP.setAttribute("MB_ID", Trim(Me.MB_ID.Text))
            Else
                mbMEMBER_TEMP.setAttribute("MB_ID", Nothing)
            End If

            '學歷
            If Utility.isValidateData(Trim(Me.MB_EDU.SelectedValue)) Then
                mbMEMBER_TEMP.setAttribute("MB_EDU", Trim(Me.MB_EDU.SelectedValue))
            Else
                mbMEMBER_TEMP.setAttribute("MB_EDU", Nothing)
            End If

            '所屬區
            mbMEMBER_TEMP.setAttribute("MB_AREA", Me.MB_AREA.SelectedValue)

            '委員
            mbMEMBER_TEMP.setAttribute("MB_LEADER", Me.MB_LEADER.SelectedValue)

            '通訊地址
            mbMEMBER_TEMP.setAttribute("MB_CITY", Me.MB_CITY.SelectedItem.Text)
            mbMEMBER_TEMP.setAttribute("MB_VLG", Me.MB_VLG.SelectedItem.Text)
            mbMEMBER_TEMP.setAttribute("MB_ADDR", Trim(Me.MB_ADDR.Text))
            '通訊郵遞區號
            mbMEMBER_TEMP.setAttribute("MB_ZIP", Trim(Me.TXT_MB_ZIP.Text))

            '戶籍地址
            If Utility.isValidateData(Me.MB_CITY1.SelectedValue) Then
                mbMEMBER_TEMP.setAttribute("MB_CITY1", Me.MB_CITY1.SelectedItem.Text)
            Else
                mbMEMBER_TEMP.setAttribute("MB_CITY1", Nothing)
            End If
            If Utility.isValidateData(Me.MB_VLG1.SelectedValue) Then
                mbMEMBER_TEMP.setAttribute("MB_VLG1", Me.MB_VLG1.SelectedItem.Text)
            Else
                mbMEMBER_TEMP.setAttribute("MB_VLG1", Nothing)
            End If
            If Utility.isValidateData(Trim(Me.MB_ADDR2.Text)) Then
                mbMEMBER_TEMP.setAttribute("MB_ADDR1", Trim(Me.MB_ADDR2.Text))
            Else
                mbMEMBER_TEMP.setAttribute("MB_ADDR1", Nothing)
            End If
            '戶籍郵遞區號
            mbMEMBER_TEMP.setAttribute("MB_ZIP1", Trim(Me.TXT_MB_ZIP1.Text))

            '會員類別
            mbMEMBER_TEMP.setAttribute("MB_FAMILY", Me.MB_FAMILY.SelectedValue)

            '修改員工編號
            mbMEMBER_TEMP.setAttribute("CHGUID", com.Azion.EloanUtility.UIUtility.getLoginUserID)

            '修改日期
            mbMEMBER_TEMP.setAttribute("CHGDATE", Now)

            mbMEMBER_TEMP.save()

            Dim mbVIP_TEMP As New MB_VIP_TEMP(Me.m_DBManager)
            For i As Integer = 0 To Me.VIP.Items.Count - 1
                mbVIP_TEMP.clear()
                mbVIP_TEMP.LoadByPK(iMB_DATE, Me.VIP.Items(i).Value)
                If Me.VIP.Items(i).Selected Then
                    'MB_MEMSEQ	decimal(7,0)		NO	PRI	0		select,insert,update,references	會員編號
                    mbVIP_TEMP.setAttribute("MB_DATE", iMB_DATE)
                    'VIP	char(1)	latin1_swedish_ci	NO	PRI			select,insert,update,references	護持會員;1長期護持會員;2種子護法
                    mbVIP_TEMP.setAttribute("VIP", Me.VIP.Items(i).Value)
                    'ACCAMT	decimal(6,0)		YES		0		select,insert,update,references	種子護法累計總繳金額
                    'CHGUID	varchar(20)	utf8_general_ci	YES				select,insert,update,references	修改員工編號
                    mbVIP_TEMP.setAttribute("CHGUID", Session("UserId"))
                    'CHGDATE	datetime		YES				select,insert,update,references	修改日期時間
                    mbVIP_TEMP.setAttribute("CHGDATE", Now)
                    mbVIP_TEMP.save()
                Else
                    If mbVIP_TEMP.isLoaded Then
                        mbVIP_TEMP.remove()
                    End If
                End If
            Next
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Sub bt_Back_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles bt_Back.Click
        Try
            Me.Clear_MB_MEMBER()

            Me.PLH_MB_MEMBER.Visible = False

            If Me.is76User(com.Azion.EloanUtility.UIUtility.getLoginUserID) Then
                Me.PLH_QRY.Visible = True

                Me.PLH_QRY_SAME.Visible = False
            Else
                Me.PLH_MAIL_SAME.Visible = True
            End If

            Me.PLH_F_CHOOSE.Visible = False
            Me.btnChoose.Visible = True
            Me.PLH_F_QRY_SAME.Visible = False
            Me.RP_F_QRY_SAME.DataSource = Nothing
            Me.RP_F_QRY_SAME.DataBind()
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    Sub Clear_MB_MEMBER()
        Try
            Me.RP_QRY_SAME.DataSource = Nothing
            Me.RP_QRY_SAME.DataBind()

            Me.HID_OTHSIGN.Value = String.Empty
            Me.HID_MB_MEMSEQ.Value = String.Empty
            Me.HID_MB_LEADER.Value = String.Empty
            Me.LTL_MB_MEMSEQ.Text = String.Empty

            '法名/姓名
            Me.MB_NAME.Text = String.Empty

            '性別
            Me.MB_SEX.SelectedIndex = -1

            '出生年
            Me.MB_BIRTH_YYY.Text = String.Empty

            '出生月
            Me.MB_BIRTH_MM.Text = String.Empty

            '出生日
            Me.MB_BIRTH_DD.Text = String.Empty

            '身分別
            Me.MB_IDENTIFY.ClearSelection()

            '手機
            Me.MB_MOBIL.Text = String.Empty

            '電話區碼
            Me.MB_TEL_Pre.Text = String.Empty

            '電話
            Me.MB_TEL.Text = String.Empty

            'e-mail
            Me.MB_EMAIL.Text = String.Empty

            '身分證字號
            Me.MB_ID.Text = String.Empty

            '學歷
            Me.MB_EDU.SelectedIndex = -1

            '所屬區
            Me.MB_AREA.SelectedIndex = -1

            '所屬區委員
            Me.MB_LEADER.SelectedIndex = -1
            Me.MB_LEADER.Items.Clear()

            '通訊地址市
            Me.MB_CITY.SelectedIndex = -1

            '通訊地址區
            Me.MB_VLG.SelectedIndex = -1
            Me.MB_VLG.Items.Clear()

            '通訊地址
            Me.MB_ADDR.Text = String.Empty

            '戶籍地址同上
            Me.MB_ADDR2_SAME.Checked = False

            '戶籍地址市
            Me.MB_CITY1.SelectedIndex = -1

            '戶籍地址區
            Me.MB_VLG1.SelectedIndex = -1
            Me.MB_VLG1.Items.Clear()

            '戶籍地址
            Me.MB_ADDR2.Text = String.Empty

            '會員類別
            Me.MB_FAMILY.SelectedIndex = -1

            '護持會員
            Me.VIP.SelectedIndex = -1

            '家族人員
            Me.TR_FAMILY.Visible = False
            Me.RP_MB_FAMILY.DataSource = Nothing
            Me.RP_MB_FAMILY.DataBind()
        Catch ex As Exception
            Throw
        End Try
    End Sub
#End Region

#Region "繳款作業"
    Sub Bind_MB_MEMREV_Default(ByVal iMB_MEMSEQ As Decimal)
        Try
            Me.PLH_MB_MEMREV.Visible = True
            Me.PLH_MB_MEMBER.Visible = False

            Dim mbMEMBER As New MB_MEMBER(Me.m_DBManager)
            If mbMEMBER.loadByPK(iMB_MEMSEQ) Then
                '法名/ 姓名
                Me.LTL_MEMREV_MB_NAME.Text = mbMEMBER.getString("MB_NAME")

                '會員編號
                Me.LTL_MEMREV_MB_MEMSEQ.Text = Me.getMB_MEMSEQ(iMB_MEMSEQ, mbMEMBER.getString("MB_AREA"))

                '通訊地址
                Me.LTL_MEMREV_MB_ADDR.Text = mbMEMBER.getString("MB_CITY") & mbMEMBER.getString("MB_VLG") & mbMEMBER.getString("MB_ADDR")

                '所屬區
                If Utility.isValidateData(mbMEMBER.getString("MB_AREA")) Then
                    Dim apAREATEAMList As New AP_AREATEAMList(Me.m_DBManager)
                    If apAREATEAMList.loadByGA_AREA(mbMEMBER.getString("MB_AREA")) > 0 Then
                        Me.LTL_MEMREV_MB_AREA.Text = apAREATEAMList.item(0).getString("A_AREA")
                    End If
                End If

                '委員
                Me.LTL_MEMREV_MB_LEADER.Text = mbMEMBER.getString("MB_LEADER")
            End If

            Me.MB_TX_DATE_YYY.Text = Now.Year
            Me.MB_TX_DATE_MM.Text = Now.Month
            Me.MB_TX_DATE_DD.Text = Now.Day

            Me.MB_RECNAME.Text = Me.LTL_MEMREV_MB_NAME.Text
            Me.CB_MB_RECNAME.Checked = True

            Me.Bind_RP_MB_MEMREV(iMB_MEMSEQ)
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Sub Bind_RP_MB_MEMREV(ByVal iMB_MEMSEQ As Decimal)
        Try
            Dim mbMEMREVList As New MB_MEMREVList(Me.m_DBManager)
            mbMEMREVList.setSQLCondition(" ORDER BY MB_SEQNO ")
            mbMEMREVList.loadByMB_MEMSEQ(iMB_MEMSEQ)
            Dim DV_MB_MEMREV As New DataView(mbMEMREVList.getCurrentDataSet.Tables(0), "VRYUID IS NULL", String.Empty, DataViewRowState.CurrentRows)
            Me.RP_MB_MEMREV.DataSource = DV_MB_MEMREV.ToTable
            Me.RP_MB_MEMREV.DataBind()
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Sub RP_MB_MEMREV_ItemDataBound(ByVal Sender As Object, ByVal e As RepeaterItemEventArgs) Handles RP_MB_MEMREV.ItemDataBound
        Try
            If e.Item.ItemType = ListItemType.Item OrElse e.Item.ItemType = ListItemType.AlternatingItem Then
                Dim LTL_MB_AREA As Literal = e.Item.FindControl("LTL_MB_AREA")
                LTL_MB_AREA.Text = Me.LTL_MEMREV_MB_AREA.Text

                Dim LTL_MB_LEADER As Literal = e.Item.FindControl("LTL_MB_LEADER")
                LTL_MB_LEADER.Text = Me.LTL_MEMREV_MB_LEADER.Text
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Sub Clear_MB_MEMREV(Optional ByVal iMode As Integer = 0)
        Try
            If iMode = 0 Then
                '法名/姓名
                Me.LTL_MEMREV_MB_NAME.Text = String.Empty

                '會員編號
                Me.LTL_MEMREV_MB_MEMSEQ.Text = String.Empty

                '通訊地址
                Me.LTL_MEMREV_MB_ADDR.Text = String.Empty

                '所屬區
                Me.LTL_MEMREV_MB_AREA.Text = String.Empty

                '委員
                Me.LTL_MEMREV_MB_LEADER.Text = String.Empty

                Me.RP_MB_MEMREV.DataSource = Nothing
                Me.RP_MB_MEMREV.DataBind()
            End If

            '繳款日期
            Me.MB_TX_DATE_YYY.Text = Now.Year
            Me.MB_TX_DATE_MM.Text = Now.Month
            Me.MB_TX_DATE_DD.Text = Now.Day

            '功德項目
            Me.MB_ITEMID.SelectedIndex = -1
            Me.PLH_MB_FAMILY.Style.Item("display") = "none"
            Me.RBL_MEMREV_MB_FAMILY.SelectedIndex = -1

            '繳款方式
            Me.MB_FEETYPE.SelectedIndex = -1
            Me.LBL_ONLY.Style.Item("display") = "none"

            '繳款金額
            Me.MB_TOTFEE.Text = String.Empty
            Me.MB_TOTFEE.Attributes.Add("Readonly", "Readonly")
            Me.PLH_MB_TOTFEE_MM.Style.Item("display") = "none"
            Me.MB_TOTFEE_MM.Text = String.Empty
            'Me.MB_TOTFEE_MM.Attributes.Add("Readonly", "Readonly")

            '收據捐款名稱
            Me.CB_MB_RECNAME.Checked = True
            Me.MB_RECNAME.Text = Me.LTL_MEMREV_MB_NAME.Text

            Me.LTL_MB_ITEMID_TYP.Text = "會費期間"

            Me.PLH_MB_MEMFEE.Style.Item("display") = ""
            Me.MB_MEMFEE_SY.Text = String.Empty
            Me.MB_MEMFEE_SM.Text = String.Empty
            Me.MB_MEMFEE_EY.Text = String.Empty
            Me.MB_MEMFEE_EM.Text = String.Empty

            Me.MB_DESC.Style.Item("display") = "none"
            Me.MB_DESC.Text = String.Empty

            '付款方式
            Me.MB_PAY_TYPE.SelectedIndex = -1

            '取　　消
            Me.bt_MB_MEMREV_CANCEL.Visible = False

            '確　　定
            Me.bt_MB_MEMREV_YES.Visible = False

            '新　　增
            Me.bt_MB_MEMREV_Add.Visible = True

            '繳費序號
            Me.HID_MB_SEQNO.Value = String.Empty

            Me.PLH_MB_PAY_TYPE_N.Style.Item("display") = "none"
            '票據到期日
            Me.NOTE_DUE_DATE_YYY.Text = String.Empty
            Me.NOTE_DUE_DATE_MM.Text = String.Empty
            Me.NOTE_DUE_DATE_DD.Text = String.Empty

            '票據號碼
            Me.NOTE_NO.Text = String.Empty

            '發票行
            Me.NOTE_BANK.Text = String.Empty

            '銀行
            Me.NOTE_BR.Text = String.Empty

            '發票人
            Me.NOTE_HOLDER.Text = String.Empty

            '票據金額
            Me.NOTE_AMT.Text = String.Empty

            '是否開立收據
            Me.RBL_MB_SEND_PRINT.SelectedIndex = -1
            Me.TD_PRINT.ColSpan = 3
            Me.PLH_MB_SEND_PRINT.Visible = False
            Me.DDL_MB_SEND_PRINT.SelectedIndex = -1

            '專案代號
            Me.PLH_PROJCODE.Style.Item("display") = "none"
            Me.PROJCODE.SelectedIndex = -1

            '累計繳款
            Me.PLH_VIP_ACCAMT.Style.Item("display") = "none"
            Me.ACCAMT.Text = String.Empty
            Me.HID_ACCAMT.Value = String.Empty
        Catch ex As Exception
            Throw
        End Try
    End Sub

    <System.Web.Services.WebMethod()> _
    Public Shared Function getMB_TOTFEE(ByVal sFEETYPE As String) As String
        Dim dbManager As DatabaseManager = UIShareFun.getDataBaseManager
        Try
            Dim apCODEList As New AP_CODEList(dbManager)
            apCODEList.LoadbyUPCODEandVALUE(m_sUpcode23, sFEETYPE)

            Dim sMB_TOTFEE_MM As String = String.Empty
            sMB_TOTFEE_MM = Utility.FormatDec(apCODEList.item(0).getDecimal("NOTE"), "#,##0")

            Dim sMB_TOTFEE As String = String.Empty
            Select Case sFEETYPE
                Case "A"
                    sMB_TOTFEE = Utility.FormatDec(apCODEList.item(0).getDecimal("NOTE"), "#,##0")
                Case "B"
                    sMB_TOTFEE = Utility.FormatDec(apCODEList.item(0).getDecimal("NOTE") * 3, "#,##0")
                Case "C"
                    sMB_TOTFEE = Utility.FormatDec(apCODEList.item(0).getDecimal("NOTE") * 12, "#,##0")
            End Select

            Dim tmpStr As New StringBuilder

            tmpStr.Append("{""MB_TOTFEE_MM"":""" & sMB_TOTFEE_MM & """,""MB_TOTFEE"":""" & sMB_TOTFEE & """}")

            Return tmpStr.ToString
        Catch ex As Exception
            Throw
        Finally
            UIShareFun.releaseConnection(dbManager)
        End Try
    End Function

#End Region

#Region "繳款作業Button Events"
    Sub bt_BackMain_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles bt_BackMain.Click
        Try
            Me.Clear_MB_MEMBER()
            Me.Clear_MB_MEMREV()

            Me.PLH_MB_MEMREV.Visible = False

            Me.PLH_QRY_SAME.Visible = False

            Me.PLH_QRY.Visible = True
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    Sub bt_MB_MEMREV_Add_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles bt_MB_MEMREV_Add.Click
        Try
            If Utility.isValidateData(Me.MB_ITEMID.SelectedValue) AndAlso Me.MB_ITEMID.SelectedValue = "A" Then
                Me.PLH_MB_FAMILY.Attributes("style") = "display:''"
            Else
                Me.PLH_MB_FAMILY.Attributes("style") = "display:'none'"
            End If

            If Not Utility.isValidateData(Me.RBL_MB_SEND_PRINT.SelectedValue) Then
                com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "請選擇是否開立收據")
                com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請選擇是否開立收據")
                Return
            End If
            If Me.RBL_MB_SEND_PRINT.SelectedValue = "1" Then
                If Not Utility.isValidateData(Me.DDL_MB_SEND_PRINT.SelectedValue) Then
                    com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "請選擇給收據方式")
                    com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請選擇給收據方式")
                    Return
                End If
            End If

            If Not CheckMB_MEMREV() Then
                Me.Recover_MEMREV_Form()
                Return
            End If
            Try
                Me.m_DBManager.beginTran()
                Me.Save_MB_MEMREV()
                Me.m_DBManager.commit()
            Catch ex As Exception
                Me.m_DBManager.Rollback()
                Throw
            End Try

            Me.Bind_RP_MB_MEMREV(CDec(Me.HID_MB_MEMSEQ.Value))

            Me.Clear_MB_MEMREV(1)
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    Sub Save_MB_MEMREV()
        Try
            Dim mbMEMREV As New MB_MEMREV(Me.m_DBManager)

            If Not IsNumeric(Me.HID_MB_SEQNO.Value) Then
                Me.HID_MB_SEQNO.Value = mbMEMREV.getMB_SEQNOPlus1(CDec(Me.HID_MB_MEMSEQ.Value))
            End If

            mbMEMREV.loadByPK(CDec(Me.HID_MB_MEMSEQ.Value), CDec(Me.HID_MB_SEQNO.Value))

            '會員編號
            mbMEMREV.setAttribute("MB_MEMSEQ", CDec(Me.HID_MB_MEMSEQ.Value))

            '繳費序號
            mbMEMREV.setAttribute("MB_SEQNO", CDec(Me.HID_MB_SEQNO.Value))

            '功德項目
            mbMEMREV.setAttribute("MB_ITEMID", Me.MB_ITEMID.SelectedValue)

            If Me.MB_ITEMID.SelectedValue = "A" Then
                '繳會費起年
                mbMEMREV.setAttribute("MB_MEMFEE_SY", CDec(Me.MB_MEMFEE_SY.Text))

                '繳會費起月
                mbMEMREV.setAttribute("MB_MEMFEE_SM", CDec(Me.MB_MEMFEE_SM.Text))

                '繳會費迄年
                mbMEMREV.setAttribute("MB_MEMFEE_EY", CDec(Me.MB_MEMFEE_EY.Text))

                '繳會費迄月
                mbMEMREV.setAttribute("MB_MEMFEE_EM", CDec(Me.MB_MEMFEE_EM.Text))

                '每月金額
                If IsNumeric(Me.MB_TOTFEE_MM.Text) Then
                    mbMEMREV.setAttribute("MB_MONFEE", CDec(Me.MB_TOTFEE_MM.Text))
                Else
                    mbMEMREV.setAttribute("MB_MONFEE", Nothing)
                End If

                '專案代碼
                mbMEMREV.setAttribute("PROJCODE", Nothing)

                '專案名稱
                mbMEMREV.setAttribute("PROJNAME", Nothing)
            Else
                '繳會費起年
                mbMEMREV.setAttribute("MB_MEMFEE_SY", Nothing)

                '繳會費起月
                mbMEMREV.setAttribute("MB_MEMFEE_SM", Nothing)

                '繳會費迄年
                mbMEMREV.setAttribute("MB_MEMFEE_EY", Nothing)

                '繳會費迄月
                mbMEMREV.setAttribute("MB_MEMFEE_EM", Nothing)

                '每月金額
                mbMEMREV.setAttribute("MB_MONFEE", Nothing)

                If Me.MB_ITEMID.SelectedValue = "B" Then
                    '專案代碼
                    mbMEMREV.setAttribute("PROJCODE", Me.PROJCODE.SelectedValue)

                    '專案名稱
                    mbMEMREV.setAttribute("PROJNAME", Me.PROJCODE.SelectedItem.Text)
                Else
                    '專案代碼
                    mbMEMREV.setAttribute("PROJCODE", Nothing)

                    '專案名稱
                    mbMEMREV.setAttribute("PROJNAME", Nothing)
                End If
            End If

            '繳款方式
            mbMEMREV.setAttribute("MB_FEETYPE", Me.MB_FEETYPE.SelectedValue)

            '繳款金額
            mbMEMREV.setAttribute("MB_TOTFEE", CDec(Me.MB_TOTFEE.Text))

            '會員類別
            If Me.MB_ITEMID.SelectedValue = "A" Then
                mbMEMREV.setAttribute("MB_MEMTYP", Me.RBL_MEMREV_MB_FAMILY.SelectedValue)
            Else
                mbMEMREV.setAttribute("MB_MEMTYP", Nothing)
            End If

            '繳款分配項目
            If Me.MB_ITEMID.SelectedValue = "A" Then
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
            Else
                mbMEMREV.setAttribute("FUND1", Nothing)
                mbMEMREV.setAttribute("FUND2", Nothing)
                mbMEMREV.setAttribute("FUND3", Nothing)
            End If

            '收據捐款名稱 同繳款人
            mbMEMREV.setAttribute("MB_RECNAME", Trim(Me.MB_RECNAME.Text))

            '說明
            If Me.MB_ITEMID.SelectedValue = "A" Then
                mbMEMREV.setAttribute("MB_DESC", Nothing)
            Else
                mbMEMREV.setAttribute("MB_DESC", Trim(Me.MB_DESC.Text))
            End If

            '付款方式 C:現金,N:票據,
            mbMEMREV.setAttribute("MB_PAY_TYPE", Me.MB_PAY_TYPE.SelectedValue)

            '繳款日期
            Dim iMB_TX_DATE As Decimal = (CDec(Me.MB_TX_DATE_YYY.Text)) * 10000 + CDec(Me.MB_TX_DATE_MM.Text) * 100 + CDec(Me.MB_TX_DATE_DD.Text)
            mbMEMREV.setAttribute("MB_TX_DATE", iMB_TX_DATE)

            If Me.MB_PAY_TYPE.SelectedValue = "N" Then
                '票據到期日
                Dim iNOTE_DUE_DATE As Decimal = (CDec(Me.NOTE_DUE_DATE_YYY.Text)) * 10000 + CDec(Me.NOTE_DUE_DATE_MM.Text) * 100 + CDec(Me.NOTE_DUE_DATE_DD.Text)
                mbMEMREV.setAttribute("NOTE_DUE_DATE", iNOTE_DUE_DATE)

                '票據號碼
                mbMEMREV.setAttribute("NOTE_NO", Trim(Me.NOTE_NO.Text))

                '發票銀行
                mbMEMREV.setAttribute("NOTE_BANK", Trim(Me.NOTE_BANK.Text))

                '發票分行
                mbMEMREV.setAttribute("NOTE_BR", Trim(Me.NOTE_BR.Text))

                '發票人
                mbMEMREV.setAttribute("NOTE_HOLDER", Trim(Me.NOTE_HOLDER.Text))

                '票據金額
                If IsNumeric(Me.NOTE_AMT.Text) Then
                    mbMEMREV.setAttribute("NOTE_AMT", CDec(Me.NOTE_AMT.Text))
                Else
                    mbMEMREV.setAttribute("NOTE_AMT", Nothing)
                End If
            Else
                '票據到期日
                mbMEMREV.setAttribute("NOTE_DUE_DATE", Nothing)

                '票據號碼
                mbMEMREV.setAttribute("NOTE_NO", Nothing)

                '發票銀行
                mbMEMREV.setAttribute("NOTE_BANK", Nothing)

                '發票分行
                mbMEMREV.setAttribute("NOTE_BR", Nothing)

                '發票人
                mbMEMREV.setAttribute("NOTE_HOLDER", Nothing)

                '票據金額
                mbMEMREV.setAttribute("NOTE_AMT", Nothing)
            End If

            '修改人員編號
            mbMEMREV.setAttribute("CHGUID", com.Azion.EloanUtility.UIUtility.getLoginUserID)

            '修改日期時間
            mbMEMREV.setAttribute("CHGDATE", Now)

            '是否開立收據
            mbMEMREV.setAttribute("MB_PRINT", Me.RBL_MB_SEND_PRINT.SelectedValue)
            If Me.RBL_MB_SEND_PRINT.SelectedValue = "1" Then
                mbMEMREV.setAttribute("MB_SEND_PRINT", Me.DDL_MB_SEND_PRINT.SelectedValue)
            Else
                mbMEMREV.setAttribute("MB_SEND_PRINT", Nothing)
            End If

            mbMEMREV.save()
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Function CheckMB_MEMREV() As Boolean
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
                com.Azion.EloanUtility.UIUtility.alert("繳款日期錯誤或未完整輸入!")
                Return False
            End If

            If Not Utility.isValidateData(Me.MB_ITEMID.SelectedValue) Then
                com.Azion.EloanUtility.UIUtility.alert("請選擇功德項目!")

                Return False
            End If

            If Me.MB_ITEMID.SelectedValue = "A" Then
                If Not Utility.isValidateData(Me.RBL_MEMREV_MB_FAMILY.SelectedValue) Then
                    com.Azion.EloanUtility.UIUtility.alert("請選擇會員類別!")

                    Return False
                End If

                If Not Utility.isValidateData(Me.MB_FEETYPE.SelectedValue) Then
                    com.Azion.EloanUtility.UIUtility.alert("請選擇繳款方式!")

                    Return False
                End If
            End If

            If Me.MB_FEETYPE.SelectedValue = "D" Then
                If Not IsNumeric(Me.MB_TOTFEE.Text) Then
                    com.Azion.EloanUtility.UIUtility.alert("請輸入繳款金額!")

                    Return False
                End If
            End If

            If Not Utility.isValidateData(Me.MB_RECNAME.Text) Then
                com.Azion.EloanUtility.UIUtility.alert("請輸入收據捐款名稱!")

                Return False
            End If

            If Me.MB_ITEMID.SelectedValue = "A" Then
                Dim sMB_MEMFEE_SY As String = String.Empty
                If IsNumeric(Me.MB_MEMFEE_SY.Text) Then
                    sMB_MEMFEE_SY = CDec(Me.MB_MEMFEE_SY.Text)
                End If
                Dim sMB_MEMFEE_SM As String = String.Empty
                If IsNumeric(Me.MB_MEMFEE_SM.Text) Then
                    sMB_MEMFEE_SM = CDec(Me.MB_MEMFEE_SM.Text)
                End If
                If Not IsDate(sMB_MEMFEE_SY & "/" & sMB_MEMFEE_SM & "/01") Then
                    com.Azion.EloanUtility.UIUtility.alert("會費期間起日年月錯誤或未完整輸入!")
                    Return False
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
                    com.Azion.EloanUtility.UIUtility.alert("會費期間迄日年月錯誤或未完整輸入!")
                    Return False
                End If

                If CDate(sMB_MEMFEE_EY & "/" & sMB_MEMFEE_EM & "/01") < CDate(sMB_MEMFEE_SY & "/" & sMB_MEMFEE_SM & "/01") Then
                    com.Azion.EloanUtility.UIUtility.alert("會費期間迄日年月不可小於起日年月!")
                    Return False
                End If
            ElseIf Me.MB_ITEMID.SelectedValue = "B" Then
                If Not Utility.isValidateData(Me.PROJCODE.SelectedValue) Then
                    com.Azion.EloanUtility.UIUtility.alert("請選擇專案名稱!")
                    Return False
                End If
            End If

            If Not Utility.isValidateData(Me.MB_PAY_TYPE.SelectedValue) Then
                com.Azion.EloanUtility.UIUtility.alert("請選擇付款方式!")

                Return False
            End If

            If Me.MB_PAY_TYPE.SelectedValue = "N" Then
                '票據到期日
                Dim sNOTE_DUE_DATE_YYY As String = String.Empty
                If IsNumeric(Me.NOTE_DUE_DATE_YYY.Text) Then
                    sNOTE_DUE_DATE_YYY = CDec(Me.NOTE_DUE_DATE_YYY.Text)
                End If

                Dim sNOTE_DUE_DATE_MM As String = String.Empty
                If IsNumeric(Me.NOTE_DUE_DATE_MM.Text) Then
                    sNOTE_DUE_DATE_MM = Me.NOTE_DUE_DATE_MM.Text
                End If

                Dim sNOTE_DUE_DATE_DD As String = String.Empty
                If IsNumeric(Me.NOTE_DUE_DATE_DD.Text) Then
                    sNOTE_DUE_DATE_DD = Me.NOTE_DUE_DATE_DD.Text
                End If

                If Not IsDate(sNOTE_DUE_DATE_YYY & "/" & sNOTE_DUE_DATE_MM & "/" & sNOTE_DUE_DATE_DD) Then
                    com.Azion.EloanUtility.UIUtility.alert("票據到期日錯誤或未完整輸入!")
                    Return False
                Else
                    Dim D_CHK_DATE As Date = Now.AddYears(-1).AddDays(10)
                    If CDate(sNOTE_DUE_DATE_YYY & "/" & sNOTE_DUE_DATE_MM & "/" & sNOTE_DUE_DATE_DD) <= D_CHK_DATE Then
                        com.Azion.EloanUtility.UIUtility.alert("票據到期日必需大於系統日前一年加10天!")
                        Return False
                    End If
                End If

                '票據號碼
                If Not Utility.isValidateData(Trim(Me.NOTE_NO.Text)) Then
                    com.Azion.EloanUtility.UIUtility.alert("請輸入票據號碼!")

                    Return False
                End If

                '發票行
                If Not Utility.isValidateData(Trim(Me.NOTE_BANK.Text)) Then
                    com.Azion.EloanUtility.UIUtility.alert("請輸入發票行!")

                    Return False
                End If

                '發票行分行
                If Not Utility.isValidateData(Trim(Me.NOTE_BR.Text)) Then
                    com.Azion.EloanUtility.UIUtility.alert("請輸入發票行分行!")

                    Return False
                End If

                '發票人
                If Not Utility.isValidateData(Trim(Me.NOTE_HOLDER.Text)) Then
                    com.Azion.EloanUtility.UIUtility.alert("請輸入發票人!")

                    Return False
                End If

                '票據金額
                If Not IsNumeric(Trim(Me.NOTE_AMT.Text)) Then
                    com.Azion.EloanUtility.UIUtility.alert("請輸入票據金額!")

                    Return False
                End If
            End If

            Return True
        Catch ex As Exception
            Throw
        End Try
    End Function

    Sub Recover_MEMREV_Form()
        Try
            If Me.MB_ITEMID.SelectedValue = "A" Then
                Me.PLH_MB_FAMILY.Style.Item("display") = ""
                Me.LBL_ONLY.Style.Item("display") = "none"
                Me.MB_FEETYPE.Style.Item("display") = ""
                Me.LTL_MB_ITEMID_TYP.Text = "會費期間"
                Me.PLH_MB_MEMFEE.Style.Item("display") = ""
                Me.MB_DESC.Style.Item("display") = "none"
            Else
                Me.PLH_MB_FAMILY.Style.Item("display") = "none"
                Me.LBL_ONLY.Style.Item("display") = ""
                Me.MB_FEETYPE.Style.Item("display") = "none"
                Me.LTL_MB_ITEMID_TYP.Text = "說明"
                Me.PLH_MB_MEMFEE.Style.Item("display") = "none"
                Me.MB_DESC.Style.Item("display") = ""
            End If

            If Me.MB_FEETYPE.SelectedValue = "D" Then
                Me.PLH_MB_TOTFEE_MM.Style.Item("display") = "none"
                Me.MB_TOTFEE.Attributes.Remove("Readonly")
            Else
                Me.PLH_MB_TOTFEE_MM.Style.Item("display") = ""
                Me.MB_TOTFEE.Attributes.Add("Readonly", "Readonly")
            End If

            If Me.MB_PAY_TYPE.SelectedValue = "N" Then
                Me.PLH_MB_PAY_TYPE_N.Style.Item("display") = ""
            Else
                Me.PLH_MB_PAY_TYPE_N.Style.Item("display") = "none"
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Sub RP_MB_MEMREV_ItemCommand(ByVal Sender As Object, ByVal e As RepeaterCommandEventArgs) Handles RP_MB_MEMREV.ItemCommand
        Try
            If UCase(e.CommandName) = "EDIT" Then
                Dim iMB_SEQNO As Decimal = e.CommandArgument

                Me.HID_MB_SEQNO.Value = iMB_SEQNO

                Dim mbMEMREV As New MB_MEMREV(Me.m_DBManager)
                If mbMEMREV.loadByPK(CDec(Me.HID_MB_MEMSEQ.Value), iMB_SEQNO) Then
                    '繳款日期
                    If Utility.isDecimalDate(mbMEMREV.getAttribute("MB_TX_DATE").ToString) Then
                        Me.MB_TX_DATE_YYY.Text = CDec(Left(mbMEMREV.getString("MB_TX_DATE"), 4))
                        Me.MB_TX_DATE_MM.Text = mbMEMREV.getString("MB_TX_DATE").Substring(4, 2)
                        Me.MB_TX_DATE_DD.Text = Right(mbMEMREV.getString("MB_TX_DATE"), 2)
                    End If

                    '功德項目
                    Me.MB_ITEMID.SelectedIndex = -1
                    If Not IsNothing(Me.MB_ITEMID.Items.FindByValue(mbMEMREV.getString("MB_ITEMID"))) Then
                        Me.MB_ITEMID.Items.FindByValue(mbMEMREV.getString("MB_ITEMID")).Selected = True
                    End If

                    '繳款方式
                    Me.MB_FEETYPE.SelectedIndex = -1
                    If Not IsNothing(Me.MB_FEETYPE.Items.FindByValue(mbMEMREV.getString("MB_FEETYPE"))) Then
                        Me.MB_FEETYPE.Items.FindByValue(mbMEMREV.getString("MB_FEETYPE")).Selected = True
                    End If

                    '繳款金額
                    If IsNumeric(mbMEMREV.getAttribute("MB_TOTFEE")) Then
                        Me.MB_TOTFEE.Text = Utility.FormatDec(mbMEMREV.getDecimal("MB_TOTFEE"), "#,##0")
                    Else
                        Me.MB_TOTFEE.Text = String.Empty
                    End If

                    If mbMEMREV.getString("MB_ITEMID") = "A" Then
                        '會員類別
                        Me.PLH_MB_FAMILY.Style.Item("display") = ""

                        Me.RBL_MEMREV_MB_FAMILY.SelectedIndex = -1
                        If Not IsNothing(Me.RBL_MEMREV_MB_FAMILY.Items.FindByValue(mbMEMREV.getString("MB_MEMTYP"))) Then
                            Me.RBL_MEMREV_MB_FAMILY.Items.FindByValue(mbMEMREV.getString("MB_MEMTYP")).Selected = True
                        End If

                        '繳款方式
                        Me.MB_FEETYPE.Style.Item("display") = ""
                        Me.LBL_ONLY.Style.Item("display") = "none"
                        Me.MB_FEETYPE.Attributes.Add("Readonly", "Readonly")

                        '每月金額???
                        'If mbMEMREV.getString("MB_FEETYPE") <> "D" Then
                        '    Me.PLH_MB_TOTFEE_MM.Style.Item("display") = ""
                        '    Dim apCODEList As New AP_CODEList(Me.m_DBManager)

                        '    apCODEList.LoadbyUPCODEandVALUE(m_sUpcode23, mbMEMREV.getString("MB_FEETYPE"))
                        '    If apCODEList.size > 0 AndAlso IsNumeric(apCODEList.item(0).getAttribute("NOTE")) Then
                        '        Me.MB_TOTFEE_MM.Text = Utility.FormatDec(apCODEList.item(0).getAttribute("NOTE"), "#,##0")
                        '    Else
                        '        Me.MB_TOTFEE_MM.Text = String.Empty
                        '    End If

                        '    Me.MB_TOTFEE.Attributes.Add("Readonly", "Readonly")
                        'Else
                        '    Me.PLH_MB_TOTFEE_MM.Style.Item("display") = "none"
                        '    Me.MB_TOTFEE_MM.Text = String.Empty

                        '    Me.MB_TOTFEE.Attributes.Remove("Readonly")
                        'End If
                        Me.PLH_MB_TOTFEE_MM.Style.Item("display") = ""
                        Me.MB_TOTFEE_MM.Text = Utility.FormatDec(mbMEMREV.getDecimal("MB_MONFEE"), "#,##0")

                        '會費期間
                        Me.LTL_MB_ITEMID_TYP.Text = "會費期間"
                        Me.PLH_MB_MEMFEE.Style.Item("display") = ""
                        Me.MB_MEMFEE_SY.Text = mbMEMREV.getDecimal("MB_MEMFEE_SY")
                        Me.MB_MEMFEE_SM.Text = mbMEMREV.getString("MB_MEMFEE_SM")
                        Me.MB_MEMFEE_EY.Text = mbMEMREV.getString("MB_MEMFEE_EY")
                        Me.MB_MEMFEE_EM.Text = mbMEMREV.getString("MB_MEMFEE_EM")
                        Me.MB_DESC.Text = String.Empty
                        Me.MB_DESC.Style.Item("display") = "none"

                        Me.TR_FUND.Attributes("style") = "display:''"

                        If mbMEMREV.getString("FUND1") = "1" Then
                            Me.FUND1.Checked = True
                        Else
                            Me.FUND1.Checked = False
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

                        Me.PLH_PROJCODE.Style.Item("display") = "none"
                        Me.PROJCODE.SelectedIndex = -1

                        If Me.RBL_MEMREV_MB_FAMILY.SelectedValue = "2" Then
                            Me.PLH_VIP_ACCAMT.Style.Item("display") = ""
                            Dim mbVIP As New MB_VIP(Me.m_DBManager)
                            mbVIP.LoadByPK(CDec(Me.HID_MB_MEMSEQ.Value), "2")
                            Me.HID_ACCAMT.Value = mbVIP.getDecimal("ACCAMT")
                            If IsNumeric(Me.MB_TOTFEE.Text) Then
                                Me.ACCAMT.Text = Utility.FormatDec(CDec(Me.MB_TOTFEE.Text) + mbVIP.getDecimal("ACCAMT"), "#,##0")
                            Else
                                Me.ACCAMT.Text = "0"
                            End If
                        Else
                            Me.PLH_VIP_ACCAMT.Style.Item("display") = "none"
                            Me.ACCAMT.Text = String.Empty
                            Me.HID_ACCAMT.Value = String.Empty
                        End If
                    Else
                        '會員類別
                        Me.PLH_MB_FAMILY.Style.Item("display") = "none"

                        Me.RBL_MEMREV_MB_FAMILY.SelectedIndex = -1

                        '繳款方式
                        Me.MB_FEETYPE.Style.Item("display") = "none"
                        Me.LBL_ONLY.Style.Item("display") = ""
                        Me.MB_TOTFEE.Attributes.Remove("Readonly")

                        '每月金額
                        Me.PLH_MB_TOTFEE_MM.Style.Item("display") = "none"
                        Me.MB_TOTFEE_MM.Text = String.Empty

                        '說明
                        Me.LTL_MB_ITEMID_TYP.Text = "說明"
                        Me.PLH_MB_MEMFEE.Style.Item("display") = "none"
                        Me.MB_MEMFEE_SY.Text = String.Empty
                        Me.MB_MEMFEE_SM.Text = String.Empty
                        Me.MB_MEMFEE_EY.Text = String.Empty
                        Me.MB_MEMFEE_EM.Text = String.Empty
                        Me.MB_DESC.Text = mbMEMREV.getString("MB_DESC")
                        Me.MB_DESC.Style.Item("display") = ""

                        Me.TR_FUND.Attributes("style") = "display:none"
                        Me.FUND1.Checked = False
                        Me.FUND2.Checked = False
                        Me.FUND3.Checked = False

                        If mbMEMREV.getString("MB_ITEMID") = "B" Then
                            Me.PLH_PROJCODE.Style.Item("display") = ""
                            Me.PROJCODE.SelectedIndex = -1
                            If Not IsNothing(Me.PROJCODE.Items.FindByValue(mbMEMREV.getString("PROJCODE"))) Then
                                Me.PROJCODE.Items.FindByValue(mbMEMREV.getString("PROJCODE")).Selected = True
                            End If
                        Else
                            Me.PLH_PROJCODE.Style.Item("display") = "none"
                            Me.PROJCODE.SelectedIndex = -1
                        End If

                        Me.PLH_VIP_ACCAMT.Style.Item("display") = "none"
                        Me.ACCAMT.Text = String.Empty
                        Me.HID_ACCAMT.Value = String.Empty
                    End If

                    '收據捐款名稱
                    Me.MB_RECNAME.Text = mbMEMREV.getString("MB_RECNAME")
                    If Me.MB_RECNAME.Text = Me.LTL_MEMREV_MB_NAME.Text Then
                        Me.CB_MB_RECNAME.Checked = True
                    Else
                        Me.CB_MB_RECNAME.Checked = False
                    End If

                    '付款方式 C:現金,N:票據,
                    Me.MB_PAY_TYPE.SelectedIndex = -1
                    If Not IsNothing(Me.MB_PAY_TYPE.Items.FindByValue(mbMEMREV.getString("MB_PAY_TYPE"))) Then
                        Me.MB_PAY_TYPE.Items.FindByValue(mbMEMREV.getString("MB_PAY_TYPE")).Selected = True
                    End If

                    Me.bt_MB_MEMREV_Add.Visible = False
                    Me.bt_MB_MEMREV_CANCEL.Visible = True
                    Me.bt_MB_MEMREV_YES.Visible = True
                    Me.PLH_Repeater.Visible = False

                    If Me.MB_PAY_TYPE.SelectedValue = "N" Then
                        Me.PLH_MB_PAY_TYPE_N.Style.Item("display") = ""

                        '票據到期日
                        If Utility.isDecimalDate(mbMEMREV.getAttribute("NOTE_DUE_DATE")) Then
                            Me.NOTE_DUE_DATE_YYY.Text = CDec(Left(mbMEMREV.getAttribute("NOTE_DUE_DATE"), 4))
                            Me.NOTE_DUE_DATE_MM.Text = CDec(mbMEMREV.getString("NOTE_DUE_DATE").Substring(4, 2))
                            Me.NOTE_DUE_DATE_DD.Text = CDec(Right(mbMEMREV.getString("NOTE_DUE_DATE"), 2))
                        Else
                            Me.NOTE_DUE_DATE_YYY.Text = String.Empty
                            Me.NOTE_DUE_DATE_MM.Text = String.Empty
                            Me.NOTE_DUE_DATE_DD.Text = String.Empty
                        End If

                        '票據號碼
                        Me.NOTE_NO.Text = Trim(mbMEMREV.getString("NOTE_NO"))

                        '發票行
                        Me.NOTE_BANK.Text = Trim(mbMEMREV.getString("NOTE_BANK"))

                        '發票行分行
                        Me.NOTE_BR.Text = Trim(mbMEMREV.getString("NOTE_BR"))

                        '發票人
                        Me.NOTE_HOLDER.Text = Trim(mbMEMREV.getString("NOTE_HOLDER"))

                        '票據金額
                        If IsNumeric(mbMEMREV.getAttribute("NOTE_AMT")) Then
                            Me.NOTE_AMT.Text = Utility.FormatDec(mbMEMREV.getAttribute("NOTE_AMT"), "#,##0")
                        Else
                            Me.NOTE_AMT.Text = String.Empty
                        End If
                    Else
                        Me.PLH_MB_PAY_TYPE_N.Style.Item("display") = "none"

                        '票據到期日
                        Me.NOTE_DUE_DATE_YYY.Text = String.Empty
                        Me.NOTE_DUE_DATE_MM.Text = String.Empty
                        Me.NOTE_DUE_DATE_DD.Text = String.Empty

                        '票據號碼
                        Me.NOTE_NO.Text = String.Empty

                        '發票行
                        Me.NOTE_BANK.Text = String.Empty

                        '發票行分行
                        Me.NOTE_BR.Text = String.Empty

                        '發票人
                        Me.NOTE_HOLDER.Text = String.Empty

                        '票據金額
                        Me.NOTE_AMT.Text = String.Empty
                    End If

                    '是否開立收據
                    Me.RBL_MB_SEND_PRINT.SelectedIndex = -1
                    If Not IsNothing(Me.RBL_MB_SEND_PRINT.Items.FindByValue(mbMEMREV.getString("MB_PRINT"))) Then
                        Me.RBL_MB_SEND_PRINT.Items.FindByValue(mbMEMREV.getString("MB_PRINT")).Selected = True
                    ElseIf mbMEMREV.getString("MB_PRINT") = "Y" Then
                        Me.RBL_MB_SEND_PRINT.Items.FindByValue("1").Selected = True
                    End If
                    Me.DDL_MB_SEND_PRINT.SelectedIndex = -1
                    If Me.RBL_MB_SEND_PRINT.SelectedValue = "1" Then
                        '給收據方式
                        If Not IsNothing(Me.DDL_MB_SEND_PRINT.Items.FindByValue(mbMEMREV.getString("MB_SEND_PRINT"))) Then
                            Me.DDL_MB_SEND_PRINT.Items.FindByValue(mbMEMREV.getString("MB_SEND_PRINT")).Selected = True
                        End If

                        Me.PLH_MB_SEND_PRINT.Visible = True
                        Me.TD_PRINT.ColSpan = 1
                    Else
                        Me.PLH_MB_SEND_PRINT.Visible = False
                        Me.TD_PRINT.ColSpan = 3
                    End If
                End If
            ElseIf UCase(e.CommandName) = "DELETE" Then
                Try
                    Me.m_DBManager.beginTran()

                    Dim iMB_SEQNO As Decimal = e.CommandArgument
                    Dim mbMEMREV As New MB_MEMREV(Me.m_DBManager)
                    If mbMEMREV.loadByPK(CDec(Me.HID_MB_MEMSEQ.Value), iMB_SEQNO) Then
                        mbMEMREV.remove()
                    End If

                    Me.m_DBManager.commit()
                Catch ex As Exception
                    Me.m_DBManager.Rollback()
                    Throw
                End Try

                Me.Bind_RP_MB_MEMREV(CDec(Me.HID_MB_MEMSEQ.Value))

                UIShareFun.showErrMsg(Me, "刪除成功!")
            End If
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    Sub bt_MB_MEMREV_CANCEL_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles bt_MB_MEMREV_CANCEL.Click
        Try
            Me.PLH_Repeater.Visible = True

            Me.bt_MB_MEMREV_CANCEL.Visible = False

            Me.bt_MB_MEMREV_YES.Visible = False

            Me.bt_MB_MEMREV_Add.Visible = True

            Me.Clear_MB_MEMREV(1)
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    Sub bt_MB_MEMREV_YES_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles bt_MB_MEMREV_YES.Click
        Try
            If Not CheckMB_MEMREV() Then
                Me.Recover_MEMREV_Form()
                Return
            End If
            Try
                Me.m_DBManager.beginTran()

                Me.Save_MB_MEMREV()

                Me.m_DBManager.commit()
            Catch ex As Exception
                Me.m_DBManager.Rollback()
                Throw
            End Try

            Me.PLH_Repeater.Visible = True

            Me.Bind_RP_MB_MEMREV(CDec(Me.HID_MB_MEMSEQ.Value))

            Me.bt_MB_MEMREV_CANCEL.Visible = False

            Me.bt_MB_MEMREV_YES.Visible = False

            Me.bt_MB_MEMREV_Add.Visible = True

            Me.Clear_MB_MEMREV(1)

            UIShareFun.showErrMsg(Me, "儲存成功")
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    Sub bt_ProcMB_MEMREV_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles bt_ProcMB_MEMREV.Click
        Try
            If Not IsNumeric(Me.HID_MB_MEMSEQ.Value) Then
                com.Azion.EloanUtility.UIUtility.alert("請先儲存會員資料!")

                Return
            End If

            Me.Clear_MB_MEMREV()

            Me.Bind_MB_MEMREV_Default(CDec(Me.HID_MB_MEMSEQ.Value))
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub
#End Region

#Region "DDL"
    Sub Bind_DDL_City()
        Try
            Me.MB_CITY.Items.Clear()

            Me.MB_CITY1.Items.Clear()

            Dim AP_ROADSECList As New AP_ROADSECList(Me.m_DBManager)

            AP_ROADSECList.loadCityNOT_D()

            Me.MB_CITY.DataTextField = "CITY"
            Me.MB_CITY.DataValueField = "CITY_ID"
            Me.MB_CITY.DataSource = AP_ROADSECList.getCurrentDataSet.Tables(0)
            Me.MB_CITY.DataBind()

            Me.MB_CITY1.DataTextField = "CITY"
            Me.MB_CITY1.DataValueField = "CITY_ID"
            Me.MB_CITY1.DataSource = AP_ROADSECList.getCurrentDataSet.Tables(0)
            Me.MB_CITY1.DataBind()

            Me.MB_CITY.Items.Insert(0, New ListItem("《縣/市》", ""))
            Me.MB_CITY1.Items.Insert(0, New ListItem("《縣/市》", ""))
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Sub Bind_MB_EDU()
        Try
            Me.MB_EDU.Items.Clear()

            Dim apCODEList As New AP_CODEList(Me.m_DBManager)
            apCODEList.loadByUpCode(Me.m_sUpCode1)

            Me.MB_EDU.DataSource = apCODEList.getCurrentDataSet.Tables(0)
            Me.MB_EDU.DataTextField = "TEXT"
            Me.MB_EDU.DataValueField = "VALUE"
            Me.MB_EDU.DataBind()

            Me.MB_EDU.Items.Insert(0, New ListItem("請選擇", ""))
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Sub Bind_MB_AREA()
        Try
            Me.MB_AREA.Items.Clear()

            Dim apCODEList As New AP_CODEList(Me.m_DBManager)
            apCODEList.loadByUpCode(m_sUpCode7)

            Me.MB_AREA.DataSource = apCODEList.getCurrentDataSet.Tables(0)
            Me.MB_AREA.DataTextField = "TEXT"
            Me.MB_AREA.DataValueField = "VALUE"
            Me.MB_AREA.DataBind()

            Me.MB_AREA.Items.Insert(0, New ListItem("請選擇", ""))
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Sub Bind_MB_FAMILY()
        Try
            Me.MB_FAMILY.Items.Clear()
            Me.RBL_MEMREV_MB_FAMILY.Items.Clear()

            Dim apCODEList As New AP_CODEList(Me.m_DBManager)
            apCODEList.loadByUpCode(m_sUpCode28)

            Me.MB_FAMILY.DataSource = apCODEList.getCurrentDataSet.Tables(0)
            Me.MB_FAMILY.DataTextField = "TEXT"
            Me.MB_FAMILY.DataValueField = "VALUE"
            Me.MB_FAMILY.DataBind()

            Me.RBL_MEMREV_MB_FAMILY.DataSource = apCODEList.getCurrentDataSet.Tables(0)
            Me.RBL_MEMREV_MB_FAMILY.DataTextField = "TEXT"
            Me.RBL_MEMREV_MB_FAMILY.DataValueField = "VALUE"
            Me.RBL_MEMREV_MB_FAMILY.DataBind()
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Sub Bind_MB_ITEMID()
        Try
            Me.MB_ITEMID.Items.Clear()

            Dim apCODEList As New AP_CODEList(Me.m_DBManager)
            apCODEList.loadByUpCode(Me.m_sUpcode15)

            Me.MB_ITEMID.DataSource = apCODEList.getCurrentDataSet.Tables(0)
            Me.MB_ITEMID.DataTextField = "TEXT"
            Me.MB_ITEMID.DataValueField = "VALUE"
            Me.MB_ITEMID.DataBind()

            Me.MB_ITEMID.Items.Insert(0, New ListItem("請選擇", ""))
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
            Me.MB_PAY_TYPE.Items.Clear()

            Dim apCODEList As New AP_CODEList(Me.m_DBManager)
            apCODEList.loadByUpCode(Me.m_sUpcode31)

            Me.MB_PAY_TYPE.DataSource = apCODEList.getCurrentDataSet.Tables(0)
            Me.MB_PAY_TYPE.DataTextField = "TEXT"
            Me.MB_PAY_TYPE.DataValueField = "VALUE"
            Me.MB_PAY_TYPE.DataBind()
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Sub Bind_MB_SEND_PRINT()
        Try
            Me.DDL_MB_SEND_PRINT.Items.Clear()

            Dim apCODEList As New AP_CODEList(Me.m_DBManager)
            apCODEList.loadByUpCode(Me.m_sUpcode272)

            Me.DDL_MB_SEND_PRINT.DataSource = apCODEList.getCurrentDataSet.Tables(0)
            Me.DDL_MB_SEND_PRINT.DataTextField = "TEXT"
            Me.DDL_MB_SEND_PRINT.DataValueField = "VALUE"
            Me.DDL_MB_SEND_PRINT.DataBind()
            Me.DDL_MB_SEND_PRINT.Items.Insert(0, New ListItem("請選擇", ""))
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

    Sub ReBind_DDL_Town_PostBack(ByVal DDL_VLG As DropDownList, ByVal sCity As String, ByVal sVLG As String)
        Try
            Dim AP_ROADSECList As New AP_ROADSECList(Me.m_DBManager)

            AP_ROADSECList.loadByCityNOT_D(sCity)

            DDL_VLG.DataSource = AP_ROADSECList.getCurrentDataSet.Tables(0)
            DDL_VLG.DataTextField = "AREA"
            DDL_VLG.DataValueField = "AREA_ID"
            DDL_VLG.DataBind()

            DDL_VLG.Items.Insert(0, New ListItem("《鄉鎮市區》", ""))

            DDL_VLG.SelectedIndex = -1
            If Not IsNothing(DDL_VLG.Items.FindByValue(sVLG)) Then
                DDL_VLG.Items.FindByValue(sVLG).Selected = True
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Sub ReBind_DDL_Town_PostBackByText(ByVal DDL_VLG As DropDownList, ByVal sCity As String, ByVal sVLG As String)
        Try
            Dim AP_ROADSECList As New AP_ROADSECList(Me.m_DBManager)

            AP_ROADSECList.loadByCityNOT_D(sCity)

            DDL_VLG.DataSource = AP_ROADSECList.getCurrentDataSet.Tables(0)
            DDL_VLG.DataTextField = "AREA"
            DDL_VLG.DataValueField = "AREA_ID"
            DDL_VLG.DataBind()

            DDL_VLG.Items.Insert(0, New ListItem("《鄉鎮市區》", ""))

            DDL_VLG.SelectedIndex = -1
            If Not IsNothing(DDL_VLG.Items.FindByText(sVLG)) Then
                DDL_VLG.Items.FindByText(sVLG).Selected = True
            End If
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

                Me.MB_LEADER.Items.Clear()
                Me.MB_LEADER.DataTextField = "TEXT"
                Me.MB_LEADER.DataValueField = "VALUE"
                Me.MB_LEADER.DataSource = apCODEList.getCurrentDataSet.Tables(0)
                Me.MB_LEADER.DataBind()

                Me.MB_LEADER.Items.Insert(0, New ListItem("請選擇", ""))

                Me.MB_LEADER.SelectedIndex = -1
                If Not IsNothing(Me.MB_LEADER.Items.FindByValue(sSelected)) Then
                    Me.MB_LEADER.Items.FindByValue(sSelected).Selected = True
                End If
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub

    '初始化鄉鎮市區
    <System.Web.Services.WebMethod()> _
    Public Shared Function Bind_DDL_Town(ByVal sCity As String) As String
        Dim dbManager As DatabaseManager = UIShareFun.getDataBaseManager
        Try
            Dim AP_ROADSECList As New AP_ROADSECList(dbManager)

            AP_ROADSECList.loadByCityNOT_D(sCity)

            Dim DT As DataTable = AP_ROADSECList.getCurrentDataSet.Tables(0)

            Dim tmpStr As New StringBuilder

            tmpStr.Append("{""AddrCode"":[")
            For i As Integer = 0 To DT.Rows.Count - 1
                If i <> DT.Rows.Count - 1 Then
                    tmpStr.Append(GetJsonStr(DT.Rows(i).Item("AREA").ToString.Trim, DT.Rows(i).Item("AREA_ID").ToString.Trim) & ",")
                Else
                    tmpStr.Append(GetJsonStr(DT.Rows(i).Item("AREA").ToString.Trim, DT.Rows(i).Item("AREA_ID").ToString.Trim))
                End If
            Next
            tmpStr.Append("]}")

            Return tmpStr.ToString
        Catch ex As Exception
            Throw
        Finally
            UIShareFun.releaseConnection(dbManager)
        End Try
    End Function

    Private Shared Function GetJsonStr(ByVal sAREA As String, ByVal sAREA_ID As String) As String
        Return "{""AREA"":""" & sAREA & """,""AREA_ID"":""" & sAREA_ID & """}"
    End Function

    '初始化所屬區
    <System.Web.Services.WebMethod()> _
    Public Shared Function Bind_MB_AREA(ByVal sCITY_ID As String) As String
        Dim dbManager As DatabaseManager = UIShareFun.getDataBaseManager
        Try
            Dim apAREATEAM As New AP_AREATEAM(dbManager)

            apAREATEAM.loadByPK(sCITY_ID)

            Dim tmpStr As New StringBuilder

            tmpStr.Append("{""GA_AREA"":")
            tmpStr.Append("{""GA_AREA"":""" & apAREATEAM.getString("GA_AREA") & """},")

            'tmpStr.Append("""MB_LEADER"":")
            'tmpStr.Append("{""MB_LEADER"":""" & m_sMB_LEADER & """},")

            Dim apCODEList As New AP_CODEList(dbManager)
            apCODEList.loadByUpCode(m_sUpCode7)
            Dim ROW_Select() As DataRow = apCODEList.getCurrentDataSet.Tables(0).Select("VALUE='" & apAREATEAM.getString("GA_AREA") & "'")
            tmpStr.Append("""MB_LEADER"":[")
            If Not IsNothing(ROW_Select) AndAlso ROW_Select.Length > 0 Then
                apCODEList.clear()
                apCODEList.loadByUpCode(ROW_Select(0)("CODEID"))
                For i As Integer = 0 To apCODEList.size - 1
                    If i = apCODEList.size - 1 Then
                        tmpStr.Append("{""MB_LEADER_TEXT"":""" & apCODEList.item(i).getString("TEXT") & """,""MB_LEADER_VALUE"":""" & apCODEList.item(i).getString("VALUE") & """}")
                    Else
                        tmpStr.Append("{""MB_LEADER_TEXT"":""" & apCODEList.item(i).getString("TEXT") & """,""MB_LEADER_VALUE"":""" & apCODEList.item(i).getString("VALUE") & """},")
                    End If
                Next
            End If
            tmpStr.Append("]}")

            Return tmpStr.ToString
        Catch ex As Exception
            Throw
        Finally
            UIShareFun.releaseConnection(dbManager)
        End Try
    End Function

    '身分證字號檢核
    <System.Web.Services.WebMethod()> _
    Public Shared Function CHK_MB_ID(ByVal sMB_ID As String) As String
        Try
            Dim sChPersonalID As String = UCase(com.Azion.EloanUtility.ValidateUtility.ChPersonalID1(Trim(sMB_ID)))
            Dim sChresidenceID As String = UCase(com.Azion.EloanUtility.ValidateUtility.ChresidenceID(Trim(sMB_ID)))
            Dim sChIDNTax As String = UCase(com.Azion.EloanUtility.ValidateUtility.ChIDNTax(Trim(sMB_ID)))
            Dim sChCompanyID As String = UCase(com.Azion.EloanUtility.ValidateUtility.ChCompanyID1(Trim(sMB_ID)))

            Dim tmpStr As New StringBuilder

            Dim sERRORMsg As String = "TRUE"

            If sChPersonalID <> "TRUE" AndAlso sChCompanyID <> "TRUE" Then
                If sChresidenceID <> "TRUE" AndAlso sChIDNTax <> "TRUE" Then
                    sERRORMsg = sChPersonalID & "\n" & sChCompanyID & "\n" & sChresidenceID & "\n" & sChIDNTax
                End If
            End If

            Dim sMB_SEX As String = String.Empty
            If sChPersonalID = "TRUE" Then
                sMB_SEX = Trim(sMB_ID).Substring(1, 1)
            End If

            tmpStr.Append("{""CHK_MB_ID"":""" & sERRORMsg & """,""MB_SEX"":""" & sMB_SEX & """}")

            Return tmpStr.ToString
        Catch ex As Exception
            Throw
        End Try
    End Function

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

    Public Function getROCDate(ByVal MB_TX_DATE As Object) As String
        Try
            If Utility.isValidateData(MB_TX_DATE) AndAlso MB_TX_DATE.ToString.Length = 8 Then
                Dim sMB_TX_DATE As String = MB_TX_DATE.ToString

                Return Utility.DateTransfer(CDate(Left(sMB_TX_DATE, 4) & "/" & sMB_TX_DATE.Substring(4, 2) & "/" & Right(sMB_TX_DATE, 2)))
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

    Public Function getMB_TOTFEE(ByVal iMB_TOTFEE As Object) As String
        Try
            If IsNumeric(iMB_TOTFEE) Then
                Return Utility.FormatDec(iMB_TOTFEE, "#,##0")
            End If
            Return String.Empty
        Catch ex As Exception
            Throw
        End Try
    End Function
#End Region

    Private Sub RBL_MB_SEND_PRINT_SelectedIndexChanged(sender As Object, e As EventArgs) Handles RBL_MB_SEND_PRINT.SelectedIndexChanged
        Try
            If Me.RBL_MB_SEND_PRINT.SelectedValue = "1" Then
                Me.PLH_MB_SEND_PRINT.Visible = True
                Me.TD_PRINT.ColSpan = 1
            Else
                Me.PLH_MB_SEND_PRINT.Visible = False
                Me.TD_PRINT.ColSpan = 3
            End If

            If MB_ITEMID.SelectedValue = "B" Then
                Me.PLH_PROJCODE.Style.Item("display") = ""
            Else
                Me.PLH_PROJCODE.Style.Item("display") = "none"
            End If
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

    Private Sub RBL_MEMREV_MB_FAMILY_SelectedIndexChanged(sender As Object, e As EventArgs) Handles RBL_MEMREV_MB_FAMILY.SelectedIndexChanged
        Try
            Me.PLH_MB_FAMILY.Attributes("style") = "display:''"
            Me.TR_FUND.Attributes("style") = "display:''"
            Me.PLH_PROJCODE.Attributes("style") = "display:none"
            Me.PROJCODE.SelectedIndex = -1
            If Me.RBL_MEMREV_MB_FAMILY.SelectedValue = "1" Then
                Me.Bind_MB_FEETYPE(1)

                Me.PLH_VIP_ACCAMT.Attributes("style") = "display:'none'"

                Me.ACCAMT.Text = String.Empty

                Me.HID_ACCAMT.Value = String.Empty
            ElseIf Me.RBL_MEMREV_MB_FAMILY.SelectedValue = "2" Then
                Me.Bind_MB_FEETYPE(2)

                Me.PLH_VIP_ACCAMT.Attributes("style") = "display:''"
            End If
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

#Region "e-Mail登入"
    Sub RB_CHOOSE_OnCheckedChanged(ByVal Sender As Object, ByVal e As System.EventArgs)
        Try
            Dim objItem As RepeaterItem = Sender.namingcontainer
            For i As Integer = 0 To Me.RP_MAIL_SAME.Items.Count - 1
                Dim RB_CHOOSE As RadioButton = Me.RP_MAIL_SAME.Items(i).FindControl("RB_CHOOSE")
                RB_CHOOSE.Checked = False
            Next
            Dim RB_C_CHOOSE As RadioButton = objItem.FindControl("RB_CHOOSE")
            RB_C_CHOOSE.Checked = True
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

    Private Sub btnOTHSign_Click(sender As Object, e As EventArgs) Handles btnOTHSign.Click
        Try
            Me.Clear_MB_MEMBER()

            Me.Clear_MB_MEMREV()

            Me.PLH_QRY.Visible = False
            Me.PLH_QRY_SAME.Visible = False
            Me.PLH_MB_MEMBER.Visible = False
            Me.PLH_MB_MEMREV.Visible = False

            Me.PLH_MB_MEMBER.Visible = True
            Me.MB_EMAIL.Text = com.Azion.EloanUtility.UIUtility.getLoginUserID

            Me.HID_OTHSIGN.Value = "1"

            Me.PLH_MAIL_SAME.Visible = False
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

    Private Sub btnModify_Click(sender As Object, e As EventArgs) Handles btnModify.Click
        Try
            Dim sMB_MEMSEQ As String = String.Empty
            For i As Integer = 0 To Me.RP_MAIL_SAME.Items.Count - 1
                Dim RB_CHOOSE As RadioButton = Me.RP_MAIL_SAME.Items(i).FindControl("RB_CHOOSE")
                If RB_CHOOSE.Checked Then
                    Dim HID_MB_MEMSEQ As HtmlInputHidden = Me.RP_MAIL_SAME.Items(i).FindControl("HID_MB_MEMSEQ")
                    sMB_MEMSEQ = HID_MB_MEMSEQ.Value
                End If
            Next

            If Not IsNumeric(sMB_MEMSEQ) Then
                com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "請點選想要修改的會員")
                com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請點選想要修改的會員")
                Return
            Else
                Me.Bind_MB_MEMBER(CDec(sMB_MEMSEQ))

                Me.PLH_MAIL_SAME.Visible = False
            End If
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

    Private Sub RP_MB_FAMILY_ItemCommand(source As Object, e As DataListCommandEventArgs) Handles RP_MB_FAMILY.ItemCommand
        Try
            If UCase(e.CommandName) = "DEL" Then
                Dim sMB_FAMSEQ As String = String.Empty
                sMB_FAMSEQ = e.CommandArgument

                Dim mbFAMILY As New MB_FAMILY(Me.m_DBManager)
                If mbFAMILY.LoadByPK(CDec(Me.HID_MB_MEMSEQ.Value), CDec(sMB_FAMSEQ)) Then
                    mbFAMILY.remove()
                End If

                Dim mbFAMILYList As New MB_FAMILYList(Me.m_DBManager)
                mbFAMILYList.LoadByMB_MEMSEQ(CDec(Me.HID_MB_MEMSEQ.Value))
                Me.RP_MB_FAMILY.DataSource = mbFAMILYList.getCurrentDataSet.Tables(0)
                Me.RP_MB_FAMILY.DataBind()
            End If
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

    Private Sub RP_MB_FAMILY_ItemDataBound(sender As Object, e As Web.UI.WebControls.DataListItemEventArgs) Handles RP_MB_FAMILY.ItemDataBound
        Try
            If e.Item.ItemType = ListItemType.Item OrElse e.Item.ItemType = ListItemType.AlternatingItem Then
                Dim DRV As DataRowView = CType(e.Item.DataItem, DataRowView)

                Dim mbMEMBER As New MB_MEMBER(Me.m_DBManager)
                If mbMEMBER.loadByPK(CDec(DRV("MB_FAMSEQ"))) Then
                    Dim LTL_MB_FAMSEQ As Literal = e.Item.FindControl("LTL_MB_FAMSEQ")

                    LTL_MB_FAMSEQ.Text = Me.getMB_MEMSEQ(mbMEMBER.getDecimal("MB_MEMSEQ"), mbMEMBER.getString("MB_AREA"))

                    Dim LTL_MB_NAME As Literal = e.Item.FindControl("LTL_MB_NAME")
                    LTL_MB_NAME.Text = mbMEMBER.getString("MB_NAME")
                End If

                Dim IMG_DEL As ImageButton = e.Item.FindControl("IMG_DEL")
                IMG_DEL.ImageUrl = com.Azion.EloanUtility.UIUtility.getRootPath & "/IMG/trashcan.gif"
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub

#End Region

    Private Sub btnChoose_Click(sender As Object, e As EventArgs) Handles btnChoose.Click
        Try
            If Not IsNumeric(Me.HID_MB_MEMSEQ.Value) Then
                com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "請先儲存產生會員編號")
                com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請先儲存產生會員編號")
                Return
            End If

            Me.PLH_F_CHOOSE.Visible = True
            Me.PLH_F_QRY_SAME.Visible = False
            Me.RP_F_QRY_SAME.DataSource = Nothing
            Me.RP_F_QRY_SAME.DataBind()

            Me.btnChoose.Visible = False
            Me.PLH_F_QRY_PARAS.Visible = True
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

    Private Sub btn_F_ReChoose_Click(sender As Object, e As EventArgs) Handles btn_F_ReChoose.Click
        Me.PLH_F_QRY_SAME.Visible = False
        Me.RP_F_QRY_SAME.DataSource = Nothing
        Me.RP_F_QRY_SAME.DataBind()
        Me.PLH_F_QRY_PARAS.Visible = True
    End Sub

    Private Sub btn_F_YES_Click(sender As Object, e As EventArgs) Handles btn_F_YES.Click
        Try
            Dim AL_MB_MEMSEQ As New ArrayList
            For i As Integer = 0 To Me.RP_F_QRY_SAME.Items.Count - 1
                Dim CKB_CHOOSE As CheckBox = Me.RP_F_QRY_SAME.Items(i).FindControl("CKB_CHOOSE")
                If CKB_CHOOSE.Checked Then
                    Dim HID_MB_MEMSEQ As HtmlInputHidden = Me.RP_F_QRY_SAME.Items(i).FindControl("HID_MB_MEMSEQ")
                    If Not AL_MB_MEMSEQ.Contains(HID_MB_MEMSEQ.Value) Then
                        AL_MB_MEMSEQ.Add(HID_MB_MEMSEQ.Value)
                    End If
                End If
            Next
            If Not IsNumeric(Me.HID_MB_MEMSEQ.Value) Then
                com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "請先儲存產生會員編號")
                com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請先儲存產生會員編號")
                Return
            End If
            If AL_MB_MEMSEQ.Count = 0 Then
                com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "請點選家族人員")
                com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請點選家族人員")
                Return
            End If

            Me.m_DBManager.beginTran()
            Try
                Dim mbFAMILY As New MB_FAMILY(Me.m_DBManager)
                For Each sMB_MEMSEQ As String In AL_MB_MEMSEQ
                    mbFAMILY.clear()
                    mbFAMILY.LoadByPK(CDec(Me.HID_MB_MEMSEQ.Value), CDec(sMB_MEMSEQ))
                    mbFAMILY.setAttribute("MB_MEMSEQ", CDec(Me.HID_MB_MEMSEQ.Value))
                    mbFAMILY.setAttribute("MB_FAMSEQ", CDec(sMB_MEMSEQ))
                    mbFAMILY.save()
                Next
                Me.m_DBManager.commit()
            Catch ex As Exception
                Me.m_DBManager.Rollback()
                Throw
            End Try

            Dim mbFAMILYList As New MB_FAMILYList(Me.m_DBManager)
            mbFAMILYList.LoadByMB_MEMSEQ(CDec(Me.HID_MB_MEMSEQ.Value))
            Me.RP_MB_FAMILY.DataSource = mbFAMILYList.getCurrentDataSet.Tables(0)
            Me.RP_MB_FAMILY.DataBind()

            Me.PLH_F_CHOOSE.Visible = False
            Me.PLH_F_QRY_SAME.Visible = False
            Me.RP_F_QRY_SAME.DataSource = Nothing
            Me.RP_F_QRY_SAME.DataBind()
            Me.btnChoose.Visible = True
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Private Sub bt_F_Qry_Click(sender As Object, e As EventArgs) Handles bt_F_Qry.Click
        Dim DT_MB_MEMBER As DataTable = Nothing
        Try
            If Me.RB_F_QRY_1.Checked = False AndAlso Me.RB_F_QRY_2.Checked = False AndAlso Me.RB_F_QRY_3.Checked = False Then
                com.Azion.EloanUtility.UIUtility.alert("請選擇輸入方式!")

                Return
            End If

            If Me.RB_F_QRY_1.Checked AndAlso Not Utility.isValidateData(Trim(Me.txt_F_QRY_Mobile.Text)) Then
                com.Azion.EloanUtility.UIUtility.alert("請輸入手機號碼!")

                Return
            End If

            If Me.RB_F_QRY_2.Checked AndAlso Not Utility.isValidateData(Trim(Me.txt_F_QRY_Phone_Pre.Text)) Then
                com.Azion.EloanUtility.UIUtility.alert("請輸入電話區碼!")

                Return
            End If

            If Me.RB_F_QRY_2.Checked AndAlso Not Utility.isValidateData(Trim(Me.txt_F_QRY_Phone.Text)) Then
                com.Azion.EloanUtility.UIUtility.alert("請輸入電話號碼!")

                Return
            End If

            If Me.RB_F_QRY_3.Checked AndAlso Not Utility.isValidateData(Trim(Me.txt_F_QRY_MB_NAME.Text)) Then
                com.Azion.EloanUtility.UIUtility.alert("請輸入姓名!")

                Return
            End If

            Dim mbMEMBERList As New MB_MEMBERList(Me.m_DBManager)

            Dim sMB_TEL As String = String.Empty
            If Me.RB_F_QRY_1.Checked Then
                mbMEMBERList.loadByMB_MOBIL(Trim(Me.txt_F_QRY_Mobile.Text))
            ElseIf Me.RB_F_QRY_2.Checked Then
                If Utility.isValidateData(Trim(Me.txt_F_QRY_Phone_Pre.Text)) Then
                    sMB_TEL = Trim(Me.txt_F_QRY_Phone_Pre.Text) & "-" & Trim(Me.txt_F_QRY_Phone.Text)
                Else
                    sMB_TEL = Trim(Me.txt_F_QRY_Phone.Text)
                End If

                mbMEMBERList.loadByMB_TEL(sMB_TEL)
            ElseIf Me.RB_F_QRY_3.Checked Then
                mbMEMBERList.loadByMB_NAME(Trim(Me.txt_F_QRY_MB_NAME.Text))
                If mbMEMBERList.getCurrentDataSet.Tables(0).Rows.Count <> 1 Then
                    mbMEMBERList.clear()
                    mbMEMBERList.loadByMB_NAME_Like(Trim(Me.txt_F_QRY_MB_NAME.Text))
                End If
            End If

            DT_MB_MEMBER = mbMEMBERList.getCurrentDataSet.Tables(0)

            Me.PLH_F_QRY_SAME.Visible = True
            Me.RP_F_QRY_SAME.DataSource = DT_MB_MEMBER
            Me.RP_F_QRY_SAME.DataBind()
            Me.PLH_F_QRY_PARAS.Visible = False
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        Finally
            If Not IsNothing(DT_MB_MEMBER) Then DT_MB_MEMBER.Dispose()
        End Try
    End Sub
End Class