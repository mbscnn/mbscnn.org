﻿Imports MBSC.MB_OP
Imports com.Azion.NET.VB
Imports MBSC.UICtl
Imports com.Azion.EloanUtility

Public Class MBMnt_Sign_01_v01
    Inherits System.Web.UI.Page

    Dim m_sUSERID As String = String.Empty
    Dim m_bTEST = False  '是否為測試模式
    Dim m_DBManager As com.Azion.NET.VB.DatabaseManager
    Dim m_sCLASS As String = String.Empty
    Dim m_sOPTYPE As String = String.Empty
    Dim m_sUpcode76 As String = String.Empty

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Page.MaintainScrollPositionOnPostBack = True

            Me.initdata()

            '維護人員
            Me.m_sUpcode76 = com.Azion.EloanUtility.CodeList.getAppSettings("UpCode76")

            Me.m_sOPTYPE = "" & Request.QueryString("OPTYPE")

            If Not Page.IsPostBack Then
                '取消報名沒有確定
                If UCase(Me.m_sOPTYPE) = "CANCEL" Then
                    Me.btn_Save.Visible = False

                    If Me.is76User(com.Azion.EloanUtility.UIUtility.getLoginUserID) Then
                        Me.tb_Page1.Visible = True
                        Me.tb_Page1_btn.Visible = True
                        Me.PLH_MAIL_SAME.Visible = False
                    Else
                        Me.tb_Page1.Visible = False
                        Me.tb_Page1_btn.Visible = False

                        Me.btnModify.Text = "取消報名"

                        Using m_DBManager As com.Azion.NET.VB.DatabaseManager = UIShareFun.getDataBaseManager
                            Dim mbMEMBERList As New MB_MEMBERList(m_DBManager)
                            mbMEMBERList.loadByMB_EMAIL_FAMILY(com.Azion.EloanUtility.UIUtility.getLoginUserID)
                            If mbMEMBERList.getCurrentDataSet.Tables(0).Rows.Count > 0 Then
                                Me.RP_MAIL_SAME.DataSource = mbMEMBERList.getCurrentDataSet.Tables(0)
                                Me.RP_MAIL_SAME.DataBind()
                                Me.PLH_MAIL_SAME.Visible = True
                            Else
                                Me.PLH_MAIL_SAME.Visible = False
                                Me.SHOW_FMT(m_DBManager, mbMEMBERList)
                            End If
                        End Using
                    End If
                ElseIf UCase(Me.m_sOPTYPE) = "QRY" Then
                    Dim sMB_MEMSEQ As String = String.Empty
                    sMB_MEMSEQ = "" & Request.QueryString("MB_MEMSEQ")
                    Using m_DBManager As com.Azion.NET.VB.DatabaseManager = UIShareFun.getDataBaseManager
                        Dim mbMEMBERList As New MB_MEMBERList(m_DBManager)
                        mbMEMBERList.loadByMB_MEMSEQ(CDec(sMB_MEMSEQ))

                        Me.SHOW_FMT(m_DBManager, mbMEMBERList)

                        com.Azion.EloanUtility.UIUtility.setControlRead(Me)
                    End Using
                Else
                    If Me.is76User(com.Azion.EloanUtility.UIUtility.getLoginUserID) Then
                        Me.tb_Page1.Visible = True
                        Me.tb_Page1_btn.Visible = True
                        Me.PLH_MAIL_SAME.Visible = False
                    Else
                        Me.tb_Page1.Visible = False
                        Me.tb_Page1_btn.Visible = False

                        Using m_DBManager As com.Azion.NET.VB.DatabaseManager = UIShareFun.getDataBaseManager
                            Dim mbMEMBERList As New MB_MEMBERList(m_DBManager)
                            mbMEMBERList.loadByMB_EMAIL_FAMILY(com.Azion.EloanUtility.UIUtility.getLoginUserID)
                            If mbMEMBERList.getCurrentDataSet.Tables(0).Rows.Count > 0 Then
                                Me.RP_MAIL_SAME.DataSource = mbMEMBERList.getCurrentDataSet.Tables(0)
                                Me.RP_MAIL_SAME.DataBind()
                                Me.PLH_MAIL_SAME.Visible = True
                            Else
                                Me.PLH_MAIL_SAME.Visible = False
                                Me.SHOW_FMT(m_DBManager, mbMEMBERList)
                            End If
                        End Using
                    End If
                End If
            End If
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    Private Sub btnModify_Click(sender As Object, e As EventArgs) Handles btnModify.Click
        Try
            Dim sMB_MEMSEQ As String = String.Empty
            Dim iCOUNT As Integer = 0
            For i As Integer = 0 To Me.RP_MAIL_SAME.Items.Count - 1
                iCOUNT += 1
                Dim RB_CHOOSE As RadioButton = Me.RP_MAIL_SAME.Items(i).FindControl("RB_CHOOSE")
                If RB_CHOOSE.Checked Then
                    Dim HID_MB_MEMSEQ As HtmlInputHidden = Me.RP_MAIL_SAME.Items(i).FindControl("HID_MB_MEMSEQ")
                    sMB_MEMSEQ = HID_MB_MEMSEQ.Value
                End If
            Next

            If iCOUNT > 0 AndAlso Not IsNumeric(sMB_MEMSEQ) Then
                com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "請點選想要修改的會員")
                com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請點選想要修改的會員")
                Return
            End If

            '開啟資料庫連線
            Using m_DBManager As com.Azion.NET.VB.DatabaseManager = UIShareFun.getDataBaseManager
                Dim MB_MEMBERList As New MB_MEMBERList(m_DBManager)
                MB_MEMBERList.loadByMB_MEMSEQ(sMB_MEMSEQ)

                Me.SHOW_FMT(m_DBManager, MB_MEMBERList)
            End Using

            Me.PLH_MAIL_SAME.Visible = False
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

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

    Sub initdata()
        Try
            '測試模式
            If m_bTEST Then
                m_sUSERID = "nick_boy1229@hotmail.com"
                m_sCLASS = "1"
                lbl_USERID.Text = m_sUSERID
                Exit Sub
            End If

            '正式模式
            If Utility.isValidateData(Session("USERID")) Then
                m_sUSERID = Session("USERID").ToString
                lbl_USERID.Text = Session("USERID").ToString
            Else
                lbl_USERID.Text = "使用者登入錯誤!"
                UIUtility.setControlRead(Me)
                Exit Sub
            End If

            m_sCLASS = "" & Request("MBSEQ")

            If m_sCLASS.Trim = "" Or Not IsNumeric(m_sCLASS.Trim) Then
                lbl_USERID.Text = "課程選擇錯誤!"
                UIUtility.setControlRead(Me)
                'ElseIf Not IsNumeric(Me.m_sMB_BATCH) Then
                '    lbl_USERID.Text = "梯次錯誤!"
                '    UIUtility.setControlRead(Me)
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub

    '開啟一號畫面
    Sub GoPage1()
        Try
            '開啟一號畫面
            tb_Page1.Visible = True
            tb_Page1_btn.Visible = True
            '關閉二三號畫面
            dgRpt_Page2.Visible = False
            tb_Page2_Button.Visible = False
            tb_Page3.Visible = False
            tb_Page3_Button.Visible = False
        Catch ex As Exception
            Throw
        End Try
    End Sub

    '開啟二號畫面
    Sub GoPage2(ByVal dt As DataTable)
        Try
            '關閉一號畫面
            tb_Page1.Visible = False
            tb_Page1_btn.Visible = False
            '開啟二號畫面
            dgRpt_Page2.Visible = True
            tb_Page2_Button.Visible = True
            '關閉三號畫面
            tb_Page3.Visible = False
            tb_Page3_Button.Visible = False

            Me.RP_Page2.DataSource = dt
            Me.RP_Page2.DataBind()
        Catch ex As Exception
            Throw
        End Try
    End Sub

    '開啟三、四號畫面
    Sub GoPage3(ByVal sMB_MEMSEQ As String)
        Try
            '關閉一號畫面
            tb_Page1.Visible = False
            tb_Page1_btn.Visible = False
            '關閉二號畫面
            dgRpt_Page2.Visible = False
            tb_Page2_Button.Visible = False
            '關閉三號畫面
            tb_Page3.Visible = True
            tb_Page3_Button.Visible = True

            '記憶是否有資料
            If sMB_MEMSEQ <> "" Then
                hid_HaveMember.Value = True
            Else
                hid_HaveMember.Value = False
            End If

            '初始化Page3
            init_Page3()

            'Session mail寫入
            txt_MB_EMAIL.Text = m_sUSERID
            Me.btn_Save.Attributes.Remove("onmousedown")

            Using m_DBManager As com.Azion.NET.VB.DatabaseManager = UIShareFun.getDataBaseManager
                Dim MB_CLASS As New MB_CLASS(m_DBManager)
                Dim MB_MEMCLASS As New MB_MEMCLASS(m_DBManager)

                '是否已經報名(已經報名則顯示取消報名Button)
                If Me.m_sOPTYPE = "CANCEL" Then
                    Dim sVadCANCEL As String = String.Empty
                    sVadCANCEL = Me.isVadCANCEL(m_DBManager, sMB_MEMSEQ)
                    If Not Utility.isValidateData(sVadCANCEL) Then
                        If MB_MEMCLASS.LoadByPK(sMB_MEMSEQ, m_sCLASS) Then
                            If MB_MEMCLASS.getString("MB_FWMK") = "4" Then
                                btn_Cancel.Visible = False
                                com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "您已經取消報名2次了")
                                Dim sURL As String = String.Empty
                                sURL = Request.ApplicationPath() & "/NewsList.aspx"
                                Server.Transfer(sURL)
                            ElseIf MB_MEMCLASS.getString("MB_FWMK") <> "3" AndAlso
                                Utility.isValidateData(MB_MEMCLASS.getAttribute("MB_CDATE")) Then
                                Me.btn_Cancel.Visible = True
                                Me.btn_Cancel.Attributes("onmousedown") = "if (confirm('您曾經取消過報名，若再次取消報名，將無法再報名了，確定嗎？')){this.click();}else{return false}"
                            Else
                                Me.btn_Cancel.Visible = True
                            End If
                        Else
                            com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "您尚未報名過此課程")
                            Dim sURL As String = String.Empty
                            sURL = Request.ApplicationPath() & "/NewsList.aspx"
                            Server.Transfer(sURL)
                        End If
                    Else
                        com.Azion.EloanUtility.UIUtility.showJSMsg(Me, sVadCANCEL)
                        Dim sURL As String = String.Empty
                        sURL = Request.ApplicationPath() & "/NewsList.aspx"
                        Server.Transfer(sURL)
                    End If
                Else
                    If MB_MEMCLASS.LoadByPK(sMB_MEMSEQ, m_sCLASS) Then
                        If MB_MEMCLASS.getString("MB_FWMK") = "3" Then
                            Me.btn_Save.Attributes("onmousedown") = "if (confirm('你已經取消報名,是否確定重新報名')){this.click();}else{return false}"
                        ElseIf MB_MEMCLASS.getString("MB_FWMK") = "4" Then
                            com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "你已經取消報名2次,不可再報名")
                            Dim sURL As String = String.Empty
                            sURL = Request.ApplicationPath() & "/NewsList.aspx"
                            Server.Transfer(sURL)
                        End If
                    End If
                End If

                '課程
                Dim MB_CLASSList As New MB_CLASSList(m_DBManager)
                MB_CLASSList.LoadByMB_SEQ(CDec(m_sCLASS))
                If MB_CLASSList.getCurrentDataSet.Tables(0).Rows.Count > 0 Then
                    Dim ROW As DataRow = MB_CLASSList.getCurrentDataSet.Tables(0).Rows(0)
                    lbl_Class.Text = ROW("MB_CLASS_NAME").ToString & "(課程編號：" &
                                     ROW("MB_SEQ").ToString & ")"

                    'If IsNumeric(Me.m_sMB_BATCH) AndAlso CDec(Me.m_sMB_BATCH) > 0 Then
                    '    lbl_Class.Text &= "(第" & Me.m_sMB_BATCH & "梯次)"
                    'End If
                Else
                    UIUtility.alert("課程編號傳入錯誤!")
                    UIUtility.setControlRead(Me)
                    Exit Sub
                End If
            End Using

            '將資料填入Page3
            If sMB_MEMSEQ.Trim <> "" Then
                Bind_Page3(sMB_MEMSEQ)
            End If
        Catch ex As System.Threading.ThreadAbortException
        Catch ex As Exception
            Throw
        End Try
    End Sub

    '將資料寫入Page3
    Sub Bind_Page3(ByVal sMB_MEMSEQ As String)
        Try
            Using m_DBManager As com.Azion.NET.VB.DatabaseManager = UIShareFun.getDataBaseManager
                Dim MB_MEMBER As New MB_MEMBER(m_DBManager)

                If MB_MEMBER.loadByPK(sMB_MEMSEQ) Then
                    With MB_MEMBER
                        '會員編號
                        lbl_MB_AREA.Text = .getString("MB_AREA")
                        lbl_MB_MEMSEQ.Text = StrUtility.FillZero(.getString("MB_MEMSEQ"), 7)
                        '法名
                        txt_MB_NAME.Text = .getString("MB_NAME")
                        '性別
                        If Not IsNothing(rbtList_MB_SEX.Items.FindByValue(.getString("MB_SEX"))) Then
                            rbtList_MB_SEX.Items.FindByValue(.getString("MB_SEX")).Selected = True
                        End If
                        '出生年月日
                        If Not IsNothing(.getAttribute("MB_BIRTH")) Then
                            txt_MB_BIRTH_YYY.Text = CDate(.getAttribute("MB_BIRTH").ToString).Year
                            txt_MB_BIRTH_MM.Text = CDate(.getAttribute("MB_BIRTH").ToString).Month
                            txt_MB_BIRTH_DD.Text = CDate(.getAttribute("MB_BIRTH").ToString).Day
                        End If
                        '出家眾
                        If .getString("MB_MONK").Trim = "1" Then
                            rbt_MB_MONK_Y.Checked = True
                        ElseIf .getString("MB_MONK").Trim = "0" Then
                            rbt_MB_MONK_N.Checked = True
                        End If

                        '出家眾區域 開啟/關閉==========================================================
                        If rbt_MB_MONK_Y.Checked = True Then
                            tr_MB_MONK_1.Attributes("Style") = "display:'block'"
                            tr_MB_MONK_2.Attributes("Style") = "display:'block'"
                            tr_MB_MONK_3.Attributes("Style") = "display:'block'"
                            tr_MB_MONK_4.Attributes("Style") = "display:'block'"
                        Else
                            tr_MB_MONK_1.Attributes("Style") = "display:'none'"
                            tr_MB_MONK_2.Attributes("Style") = "display:'none'"
                            tr_MB_MONK_3.Attributes("Style") = "display:'none'"
                            tr_MB_MONK_4.Attributes("Style") = "display:'none'"
                        End If
                        '法名
                        txt_MB_MONKNAME.Text = .getString("MB_MONKNAME")
                        '剃度/皈依恩師/戒師
                        txt_MB_MONKTECH.Text = .getString("MB_MONKTECH")
                        '傳承
                        If Not IsNothing(dd_MB_EDUTYPE.Items.FindByValue(.getString("MB_EDUTYPE"))) Then
                            dd_MB_EDUTYPE.Items.FindByValue(.getString("MB_EDUTYPE")).Selected = True
                        End If
                        '常住/親近道場
                        txt_MB_MONKPLACE.Text = .getString("MB_MONKPLACE")
                        '戒別
                        If Not IsNothing(dd_MB_MONKTYPE.Items.FindByValue(.getString("MB_MONKTYPE"))) Then
                            dd_MB_MONKTYPE.Items.FindByValue(.getString("MB_MONKTYPE")).Selected = True
                        End If
                        '受戒日期
                        If IsDate(.getString("MB_MONKDATE")) Then
                            Dim dDate As Date = CDate(.getString("MB_MONKDATE"))
                            txt_MB_MONKDATE_YYY.Text = dDate.Year
                            txt_MB_MONKDATE_MM.Text = dDate.Month
                            txt_MB_MONKDATE_DD.Text = dDate.Day
                        End If
                        '受戒地點
                        txt_MB_MONKPLACE1.Text = .getString("MB_MONKPLACE1")
                        '====================================================================================
                        '手機
                        txt_MB_MOBIL.Text = .getString("MB_MOBIL")
                        '電話
                        If .getString("MB_TEL").Length >= 2 Then
                            If InStr(.getString("MB_TEL"), "-") > 0 Then
                                Dim sTMP() As String = Nothing
                                sTMP = Split(.getString("MB_TEL"), "-")
                                If Not IsNothing(sTMP) AndAlso sTMP.Length > 0 Then
                                    Dim iTel As Integer = 0
                                    For Each sTel As String In sTMP
                                        If IsNumeric(sTel) Then
                                            If iTel = 0 Then
                                                txt_MB_TEL_ZIP.Text = sTel
                                            ElseIf iTel = 1 Then
                                                txt_MB_TEL.Text = sTel
                                            End If
                                            iTel += 1
                                        End If
                                    Next
                                End If
                            Else
                                txt_MB_TEL_ZIP.Text = Left(.getString("MB_TEL"), 2)
                                txt_MB_TEL.Text = Right(.getString("MB_TEL"), .getString("MB_TEL").Length - 2)
                            End If
                        End If
                        'e-mail
                        txt_MB_EMAIL.Text = IIf(.getString("MB_EMAIL").Trim <> "", .getString("MB_EMAIL"), txt_MB_EMAIL.Text)
                        '身分證字號
                        txt_MB_ID.Text = .getString("MB_ID")
                        '學歷
                        If Not IsNothing(dd_MB_EDU.Items.FindByValue(.getString("MB_EDU"))) Then
                            dd_MB_EDU.Items.FindByValue(.getString("MB_EDU")).Selected = True
                        End If
                        '介紹人/訊息來源
                        txt_MB_REFER.Text = .getString("MB_REFER")
                        '通訊地址
                        If Not IsNothing(dd_MB_CITY.Items.FindByText(.getString("MB_CITY"))) Then
                            dd_MB_CITY.Items.FindByText(.getString("MB_CITY")).Selected = True

                            Bind_ddl_VLG(dd_MB_CITY.Items.FindByText(.getString("MB_CITY")).Value, dd_MB_VLG)
                            If Not IsNothing(dd_MB_VLG.Items.FindByText(.getString("MB_VLG"))) Then
                                dd_MB_VLG.Items.FindByText(.getString("MB_VLG")).Selected = True
                            End If
                        End If
                        txt_MB_ADDR.Text = .getString("MB_ADDR")
                        '戶籍地址
                        If Not IsNothing(dd_MB_CITY1.Items.FindByText(.getString("MB_CITY1"))) Then
                            dd_MB_CITY1.Items.FindByText(.getString("MB_CITY1")).Selected = True

                            Bind_ddl_VLG(dd_MB_CITY1.Items.FindByText(.getString("MB_CITY1")).Value, dd_MB_VLG1)
                            If Not IsNothing(dd_MB_VLG1.Items.FindByText(.getString("MB_VLG1"))) Then
                                dd_MB_VLG1.Items.FindByText(.getString("MB_VLG1")).Selected = True
                            End If
                        End If
                        txt_MB_ADDR1.Text = .getString("MB_ADDR1")
                        '語言
                        Dim sMB_LANG() As String = .getString("MB_LANG").Split(",")

                        For i As Integer = 0 To sMB_LANG.Count - 1
                            For j As Integer = 0 To dl_MB_LANG.Items.Count - 1
                                Dim cb_MB_LANG As CheckBox = dl_MB_LANG.Items(i).FindControl("cb_MB_LANG")
                                Dim hid_MB_LANG As HiddenField = dl_MB_LANG.Items(i).FindControl("hid_MB_LANG")
                                Dim txt_MB_LANG As TextBox = dl_MB_LANG.Items(i).FindControl("txt_MB_LANG")

                                If hid_MB_LANG.Value.Trim = sMB_LANG(i) Then
                                    cb_MB_LANG.Checked = True

                                    If hid_MB_LANG.Value.Trim = "9" Then
                                        txt_MB_LANG.Visible = True
                                        txt_MB_LANG.Text = .getString("MB_OLANG")
                                    End If
                                End If
                            Next
                        Next
                        '專長
                        If Not IsNothing(dd_MB_SPECIAL.Items.FindByValue(.getString("MB_SPECIAL"))) Then
                            dd_MB_SPECIAL.Items.FindByValue(.getString("MB_SPECIAL")).Selected = True
                        End If
                        '職業
                        If Not IsNothing(dd_MB_JOB.Items.FindByValue(.getString("MB_PROFESSION"))) Then
                            dd_MB_JOB.Items.FindByValue(.getString("MB_PROFESSION")).Selected = True
                        End If
                        '職稱
                        If Not IsNothing(dd_MB_JOBTITLE.Items.FindByValue(.getString("MB_JOBTITLE"))) Then
                            dd_MB_JOBTITLE.Items.FindByValue(.getString("MB_JOBTITLE")).Selected = True
                        End If
                        '宗教信仰
                        If Not IsNothing(dd_MB_RELIGION.Items.FindByValue(.getString("MB_RELIGION"))) Then
                            dd_MB_RELIGION.Items.FindByValue(.getString("MB_RELIGION")).Selected = True
                        End If
                        '公司/學校系級
                        Me.SCHOOL.Text = .getString("SCHOOL")
                        '是否曾參加過本中心課程;Y:是,N:否
                        Me.JOINMBSC.SelectedIndex = -1
                        If Utility.isValidateData(.getAttribute("JOINMBSC")) Then
                            If Not IsNothing(Me.JOINMBSC.Items.FindByValue(.getAttribute("JOINMBSC"))) Then
                                Me.JOINMBSC.Items.FindByValue(.getAttribute("JOINMBSC")).Selected = True
                            End If
                        End If
                        '打鼾
                        If Not IsNothing(rbtList_MB_SNORE.Items.FindByValue(.getString("MB_SNORE"))) Then
                            rbtList_MB_SNORE.Items.FindByValue(.getString("MB_SNORE")).Selected = True
                        End If
                        '身心狀況
                        Dim sMB_SICK() As String = .getString("MB_SICK").Split(",")

                        For i As Integer = 0 To sMB_SICK.Count - 1
                            For j As Integer = 0 To dl_MB_SICK.Items.Count - 1
                                Dim cb_MB_SICK As CheckBox = dl_MB_SICK.Items(j).FindControl("cb_MB_SICK")
                                Dim hid_MB_SICK As HiddenField = dl_MB_SICK.Items(j).FindControl("hid_MB_SICK")
                                Dim txt_MB_SICK As TextBox = dl_MB_SICK.Items(j).FindControl("txt_MB_SICK")

                                If hid_MB_SICK.Value.Trim = sMB_SICK(i) Then
                                    cb_MB_SICK.Checked = True

                                    If cb_MB_SICK.Text.Trim = "過敏" Then
                                        txt_MB_SICK.Visible = True
                                        txt_MB_SICK.Text = .getString("MB_ALLERGY")
                                    ElseIf cb_MB_SICK.Text.Trim = "開刀" Then
                                        txt_MB_SICK.Visible = True
                                        txt_MB_SICK.Text = .getString("MB_OPERATE")
                                    ElseIf cb_MB_SICK.Text.Trim = "其他" Then
                                        txt_MB_SICK.Visible = True
                                        txt_MB_SICK.Text = .getString("MB_OSICK")
                                    End If
                                End If
                            Next
                        Next
                        '修持法門區
                        '您曾經修持過毗婆舍那禪法嗎
                        If .getString("MB_PIPOSHENA") = "0" Then
                            rbt_MB_PIPOSHENA_N.Checked = True
                        ElseIf .getString("MB_PIPOSHENA") = "1" Then
                            rbt_MB_PIPOSHENA_Y.Checked = True
                        End If
                        '指導老師
                        txt_MB_TEACH.Text = .getString("MB_TEACH")
                        '您目前的修持法門
                        txt_MB_FAMENNIAN.Text = .getString("MB_FAMENNIAN")
                        '您過去曾經參加過七天以上的禪修嗎
                        If .getString("MB_OVER7DAY") = "0" Then
                            rbt_MB_OVER7DAY_N.Checked = True
                        ElseIf .getString("MB_OVER7DAY") = "1" Then
                            rbt_MB_OVER7DAY_Y.Checked = True
                        End If
                        '地點
                        txt_MB_PLACE.Text = .getString("MB_PLACE")
                        '緊急聯絡人姓名
                        txt_MB_EMGCONT.Text = .getString("MB_EMGCONT")
                        '緊急聯絡人 手機
                        txt_MB_CONTMOBIL.Text = .getString("MB_CONTMOBIL")
                        '捐助物資/捐款
                        MB_AMTMEMO.Text = .getString("MB_AMTMEMO")
                        '是否初學者
                        Me.MB_BEGIN.SelectedIndex = -1
                        If Not IsNothing(Me.MB_BEGIN.Items.FindByValue(.getString("MB_BEGIN"))) Then
                            Me.MB_BEGIN.Items.FindByValue(.getString("MB_BEGIN")).Selected = True
                        End If
                        '每次禪坐時間
                        Me.MB_SITIME.Text = .getString("MB_SITIME")
                    End With
                    'Else
                    '    '取號
                    '    lbl_MB_MEMSEQ.Text = getVadID(m_DBManager)
                    '    MB_MEMBER.loadByPK(lbl_MB_MEMSEQ.Text)
                    '    MB_MEMBER.setAttribute("MB_MEMSEQ", lbl_MB_MEMSEQ.Text)
                    '    MB_MEMBER.setAttribute("MB_EMAIL", m_sUSERID)
                    '    MB_MEMBER.setAttribute("MB_MOBIL", IIf(IsNumeric(txt_Cel.Text), txt_Cel.Text, DBNull.Value))
                    '    MB_MEMBER.save()

                    '參加動機或目的
                    Dim mbMEMCLASS As New MB_MEMCLASS(m_DBManager)
                    If mbMEMCLASS.LoadByPK(sMB_MEMSEQ, Me.m_sCLASS) Then
                        '參加動機或目的
                        Me.MB_OBJECT.Text = mbMEMCLASS.getString("MB_OBJECT")
                    End If
                End If
            End Using
        Catch ex As Exception
            Throw
        End Try
    End Sub

    '取號
    Function getVadID(ByVal dbManager As com.Azion.NET.VB.DatabaseManager) As String
        Try
            Dim sProcName As String = String.Empty
            sProcName = LCase(com.Azion.NET.VB.Properties.getSchemaName) & ".MB_GET_MAXID"
            Dim inParaAL As New ArrayList
            Dim outParaAL As New ArrayList
            inParaAL.Add("01")
            inParaAL.Add("1")
            outParaAL.Add(7)
            Dim HT_Return As Hashtable = com.Azion.NET.VB.DBObject.getProcedureData(dbManager, sProcName, inParaAL, outParaAL)
            Dim iMAXID As Decimal = HT_Return.Item("@IMAXID")
            Return com.Azion.EloanUtility.StrUtility.FillZero(iMAXID, 7)
        Catch ex As Exception
            Throw
        Finally
        End Try
    End Function


    '初始化各List
    Sub init_Page3()
        Try
            Using m_DBManager As com.Azion.NET.VB.DatabaseManager = UIShareFun.getDataBaseManager
                Dim AP_CODEList As New AP_CODEList(m_DBManager)
                Dim AP_ROADSECList As New AP_ROADSECList(m_DBManager)

                '學歷
                AP_CODEList.getVALUEandTEXTbyUPCODE(CodeList.getAppSettings("UPCODE1"))
                init_ddl(dd_MB_EDU, AP_CODEList.getCurrentDataSet.Tables(0))
                '通訊地址 & 戶籍地址(城市部分先初始化)
                AP_ROADSECList.loadCityNOT_D(False)
                init_ddl(dd_MB_CITY, AP_ROADSECList.getCurrentDataSet.Tables(0), "CITY_ID", "CITY")
                Me.dd_MB_CITY.Items.Insert(1, New ListItem("非本國", "非本國"))
                init_ddl(dd_MB_CITY1, AP_ROADSECList.getCurrentDataSet.Tables(0), "CITY_ID", "CITY")
                Me.dd_MB_CITY1.Items.Insert(1, New ListItem("非本國", "非本國"))
                AP_ROADSECList.clear()
                '語言
                AP_CODEList.getCurrentDataSet.Clear()
                AP_CODEList.getVALUEandTEXTbyUPCODE(CodeList.getAppSettings("UPCODE95"))
                'init_ddl(dl_MB_LANG, AP_CODEList.getCurrentDataSet.Tables(0))
                dl_MB_LANG.DataSource = AP_CODEList.getCurrentDataSet.Tables(0)
                dl_MB_LANG.DataBind()
                '專長
                AP_CODEList.getCurrentDataSet.Clear()
                AP_CODEList.getVALUEandTEXTbyUPCODE(CodeList.getAppSettings("UPCODE164"))
                init_ddl(dd_MB_SPECIAL, AP_CODEList.getCurrentDataSet.Tables(0))
                '職業
                AP_CODEList.getCurrentDataSet.Clear()
                AP_CODEList.getVALUEandTEXTbyUPCODE(CodeList.getAppSettings("UPCODE103"))
                init_ddl(dd_MB_JOB, AP_CODEList.getCurrentDataSet.Tables(0))
                '職稱
                AP_CODEList.getCurrentDataSet.Clear()
                AP_CODEList.getVALUEandTEXTbyUPCODE(CodeList.getAppSettings("UPCODE152"))
                init_ddl(dd_MB_JOBTITLE, AP_CODEList.getCurrentDataSet.Tables(0))
                '宗教信仰
                AP_CODEList.getCurrentDataSet.Clear()
                AP_CODEList.getVALUEandTEXTbyUPCODE(CodeList.getAppSettings("UPCODE127"))
                init_ddl(dd_MB_RELIGION, AP_CODEList.getCurrentDataSet.Tables(0))
                '身心狀況
                AP_CODEList.getCurrentDataSet.Clear()
                AP_CODEList.getVALUEandTEXTbyUPCODE(CodeList.getAppSettings("UPCODE136"))
                'init_ddl(dl_MB_SICK, AP_CODEList.getCurrentDataSet.Tables(0))
                dl_MB_SICK.DataSource = AP_CODEList.getCurrentDataSet.Tables(0)
                dl_MB_SICK.DataBind()
                '傳承
                AP_CODEList.getCurrentDataSet.Clear()
                AP_CODEList.getVALUEandTEXTbyUPCODE(CodeList.getAppSettings("UPCODE165"))
                init_ddl(dd_MB_EDUTYPE, AP_CODEList.getCurrentDataSet.Tables(0))
                '戒別
                AP_CODEList.getCurrentDataSet.Clear()
                AP_CODEList.getVALUEandTEXTbyUPCODE(CodeList.getAppSettings("UPCODE170"))
                init_ddl(dd_MB_MONKTYPE, AP_CODEList.getCurrentDataSet.Tables(0))
                Clear_Page3()
            End Using
        Catch ex As Exception
            Throw
        End Try
    End Sub

    'Bind ddl
    Sub init_ddl(ByVal dd_cb As Object,
                 ByVal dt As DataTable,
                 Optional ByVal sVALUE As String = "VALUE",
                 Optional ByVal sTEXT As String = "TEXT",
                 Optional ByVal bFirstItem As Boolean = True)
        Try
            If TypeOf (dd_cb) Is DropDownList Then
                dd_cb = CType(dd_cb, DropDownList)
            ElseIf TypeOf (dd_cb) Is CheckBoxList Then
                dd_cb = CType(dd_cb, CheckBoxList)
                bFirstItem = False
            ElseIf TypeOf (dd_cb) Is DataList Then
                dd_cb = CType(dd_cb, DataList)
                bFirstItem = False
            End If

            If Not TypeOf (dd_cb) Is DataList Then
                dd_cb.DataValueField = sVALUE
                dd_cb.DataTextField = sTEXT
            End If
            
            dd_cb.DataSource = dt
            dd_cb.DataBind()

            If bFirstItem Then
                dd_cb.Items.Insert(0, "請選擇")
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub

    '初始化Page3
    Sub Clear_Page3()
        Try
            '初始化Control
            For Each ctl As Web.UI.Control In tb_Page3.Controls
                If TypeOf (ctl) Is RadioButton Then
                    CType(ctl, RadioButton).Checked = False
                ElseIf TypeOf (ctl) Is RadioButtonList Then
                    For i As Integer = 0 To CType(ctl, RadioButtonList).Items.Count - 1
                        CType(ctl, RadioButtonList).Items(i).Selected = False
                    Next
                ElseIf TypeOf (ctl) Is CheckBox Then
                    CType(ctl, CheckBox).Checked = False
                ElseIf TypeOf (ctl) Is TextBox Then
                    CType(ctl, TextBox).Text = ""
                ElseIf TypeOf (ctl) Is DropDownList Then
                    CType(ctl, DropDownList).SelectedIndex = -1
                End If
            Next
        Catch ex As Exception
            Throw
        End Try
    End Sub

    '語言DataList-ItemDataBound
    Sub dl_MB_LANG_ItemDataBound(sender As Object, e As System.Web.UI.WebControls.RepeaterItemEventArgs) Handles dl_MB_LANG.ItemDataBound
        Try
            If e.Item.ItemType = ListItemType.Item Or e.Item.ItemType = ListItemType.AlternatingItem Then
                Dim hid_MB_LANG As HiddenField = e.Item.FindControl("hid_MB_LANG")
                Dim cb_MB_LANG As CheckBox = e.Item.FindControl("cb_MB_LANG")
                Dim txt_MB_LANG As TextBox = e.Item.FindControl("txt_MB_LANG")

                If Not IsNothing(hid_MB_LANG) And Not IsNothing(cb_MB_LANG) And Not IsNothing(txt_MB_LANG) Then
                    If cb_MB_LANG.Text.Trim = "其他" Or hid_MB_LANG.Value = "9" Then
                        txt_MB_LANG.Visible = True
                    End If
                End If
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub

    '身心狀況DataList-ItemDataBound
    Sub dl_MB_SICK_ItemDataBound(sender As Object, e As System.Web.UI.WebControls.RepeaterItemEventArgs) Handles dl_MB_SICK.ItemDataBound
        Try
            If e.Item.ItemType = ListItemType.Item Or e.Item.ItemType = ListItemType.AlternatingItem Then
                Dim hid_MB_SICK As HiddenField = e.Item.FindControl("hid_MB_SICK")
                Dim cb_MB_SICK As CheckBox = e.Item.FindControl("cb_MB_SICK")
                Dim txt_MB_SICK As TextBox = e.Item.FindControl("txt_MB_SICK")

                If Not IsNothing(hid_MB_SICK) And Not IsNothing(cb_MB_SICK) And Not IsNothing(txt_MB_SICK) Then
                    If cb_MB_SICK.Text.Trim = "過敏" Or cb_MB_SICK.Text.Trim = "開刀" Or cb_MB_SICK.Text.Trim = "其他" Then
                        txt_MB_SICK.Visible = True
                    End If
                End If
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub

#Region "地址區事件"
    '通訊地址區城市下拉事件
    Protected Sub dd_MB_CITY_SelectedIndexChanged(sender As Object, e As EventArgs) Handles dd_MB_CITY.SelectedIndexChanged
        Try
            Bind_ddl_VLG(dd_MB_CITY.SelectedValue, dd_MB_VLG)

            com.Azion.EloanUtility.UIUtility.setObjFocus(Me, Me.dd_MB_CITY.ClientID)
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    '戶籍地址區城市下拉事件
    Protected Sub dd_MB_CITY1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles dd_MB_CITY1.SelectedIndexChanged
        Try
            Bind_ddl_VLG(dd_MB_CITY1.SelectedValue, dd_MB_VLG1)

            com.Azion.EloanUtility.UIUtility.setObjFocus(Me, Me.dd_MB_CITY1.ClientID)
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    '區下拉初始化
    Sub Bind_ddl_VLG(ByVal sCITY As String, ByVal dd As DropDownList)
        Try
            Using m_DBManager As com.Azion.NET.VB.DatabaseManager = UIShareFun.getDataBaseManager
                Dim AP_ROADSECList As New AP_ROADSECList(m_DBManager)

                If sCITY.Trim = "請選擇" Then
                    dd.Items.Clear()
                    dd.Items.Add("請選擇")
                ElseIf sCITY = "非本國" Then
                    dd.Items.Clear()
                    dd.Items.Insert(0, New ListItem("非本國", "非本國"))
                Else
                    AP_ROADSECList.loadByCityId(sCITY)
                    init_ddl(dd, AP_ROADSECList.getCurrentDataSet.Tables(0), "AREA_ID", "AREA")
                End If
            End Using
        Catch ex As Exception
            Throw
        End Try
    End Sub

    '同上CheckBox
    Protected Sub cb_Ditto_CheckedChanged(sender As Object, e As EventArgs) Handles cb_Ditto.CheckedChanged
        Try
            If cb_Ditto.Checked Then
                If Not IsNothing(dd_MB_CITY1.Items.FindByValue(dd_MB_CITY.SelectedValue)) Then
                    dd_MB_CITY1.SelectedValue = dd_MB_CITY.SelectedValue
                    Bind_ddl_VLG(dd_MB_CITY1.SelectedValue, dd_MB_VLG1)

                    If Not IsNothing(dd_MB_VLG1.Items.FindByValue(dd_MB_VLG.SelectedValue)) Then
                        dd_MB_VLG1.SelectedValue = dd_MB_VLG.SelectedValue
                    End If
                End If

                txt_MB_ADDR1.Text = txt_MB_ADDR.Text
            Else
                'dd_MB_CITY1.SelectedIndex = -1
                'Bind_ddl_VLG(dd_MB_CITY1.SelectedValue, dd_MB_VLG1)
            End If

            com.Azion.EloanUtility.UIUtility.setObjFocus(Me, Me.cb_Ditto.ClientID)
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

#End Region

#Region "Button Events"
    '一號畫面Button
    Protected Sub btn_Qry_Click(sender As Object, e As EventArgs) Handles btn_Qry.Click
        Try
            '檢核是否有勾選查詢方式
            If Not rbt_Cel.Checked And Not rbt_Tel.Checked Then
                UIUtility.alert("請選擇查詢方式!")
                Exit Sub
            End If

            '開啟資料庫連線
            Using m_DBManager As com.Azion.NET.VB.DatabaseManager = UIShareFun.getDataBaseManager
                Dim MB_MEMBERList As New MB_MEMBERList(m_DBManager)

                If rbt_Cel.Checked Then '點選手機=====================================================================================
                    '檢核手機是否輸入
                    If IsNumeric(txt_Cel.Text.Trim) And txt_Cel.Text.Trim.Length = 10 And Left(txt_Cel.Text.Trim, 2) = "09" Then
                        '依手機撈取資料
                        MB_MEMBERList.loadByMB_MOBIL(txt_Cel.Text.Trim)
                    Else
                        UIUtility.alert("請輸入正確的手機號碼以查詢資料!")
                        Exit Sub
                    End If
                ElseIf rbt_Tel.Checked Then '點選電話=================================================================================
                    '檢核電話是否輸入
                    If IsNumeric(txt_Tel_Zip.Text.Trim) And IsNumeric(txt_Tel.Text.Trim) And txt_Tel_Zip.Text.Trim.Length = 2 Then
                        '依電話撈取資料
                        MB_MEMBERList.loadByMB_TEL(txt_Tel_Zip.Text.Trim & "-" & txt_Tel.Text.Trim)
                    ElseIf Not IsNumeric(txt_Tel_Zip.Text) AndAlso IsNumeric(txt_Tel.Text) then
                        MB_MEMBERList.loadByMB_TEL(txt_Tel.Text)
                    Else
                        UIUtility.alert("請輸入正確的電話號碼以查詢資料!")
                        Exit Sub
                    End If
                End If

                Me.SHOW_FMT(m_DBManager, MB_MEMBERList)
            End Using

        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    Sub SHOW_FMT(ByVal dbManager As DatabaseManager, ByVal MB_MEMBERList As MB_MEMBERList)
        Try
            Dim MB_CLASSList As New MB_CLASSList(dbManager)
            MB_CLASSList.setSQLCondition(" ORDER BY MB_BATCH ")
            MB_CLASSList.LoadByMB_SEQ(CDec(Me.m_sCLASS))
            If MB_CLASSList.getCurrentDataSet.Tables(0).Rows.Count > 0 Then
                Dim ROW As DataRow = MB_CLASSList.getCurrentDataSet.Tables(0).Rows(0)
                Me.DealPageFMT(ROW("MB_APV").ToString)

                If ROW("MB_BEGIN").ToString = "Y" Then
                    Me.TR_G_16_T.Visible = True
                Else
                    Me.TR_G_16_T.Visible = False
                End If
            End If

            '開啟下一個畫面
            If MB_MEMBERList.getCurrentDataSet.Tables(0).Rows.Count > 1 Then
                'show 圖(二)
                GoPage2(MB_MEMBERList.getCurrentDataSet.Tables(0))
            Else
                '圖(三) or 圖(四) 視為同一事件
                If MB_MEMBERList.getCurrentDataSet.Tables(0).Rows.Count = 1 Then
                    GoPage3(MB_MEMBERList.getCurrentDataSet.Tables(0).Rows(0).Item("MB_MEMSEQ").ToString.Trim)
                Else
                    GoPage3("")
                End If
            End If

            Me.RP_BATCH.DataSource = MB_CLASSList.getCurrentDataSet.Tables(0)
            Me.RP_BATCH.DataBind()

            If MB_CLASSList.getCurrentDataSet.Tables(0).Rows.Count > 1 Then
                Me.TR_BATCH.Visible = True
            Else
                Me.TR_BATCH.Visible = False
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Sub DealPageFMT(ByVal sMB_APV As String)
        Try
            sMB_APV = UCase(sMB_APV)
            If sMB_APV = "2" Then
                'Show 黃色區
                '出生年月日
                'Me.TD_Y_1.ColSpan = 3
                Me.TD_Y_1.Attributes("class") = "col-md-10 text-left"
                '西元
                Me.TD_G_1_1.Visible = False
                '出家眾
                Me.TD_G_1_2.Visible = False

                '法名,剃度/皈依恩師/戒師
                Me.tr_MB_MONK_1.Visible = False
                '傳承,常住/親近道場
                Me.tr_MB_MONK_2.Visible = False
                '戒別,受戒日期,西元
                Me.tr_MB_MONK_3.Visible = False
                '受戒地點
                Me.tr_MB_MONK_4.Visible = False

                'E-mail
                'Me.TD_Y_2.ColSpan = 3
                Me.TD_Y_2.Attributes("class") = "col-md-10 text-left"
                '身分證字號
                Me.TD_G_2_1.Visible = False
                Me.TD_G_2_2.Visible = False

                '通訊地址
                Me.TR_G_3.Visible = False
                '戶籍地址
                Me.TR_G_4.Visible = False
                '語言
                Me.TR_G_5.Visible = False
                '專長,職業
                Me.TR_G_6.Visible = False

                '宗教信仰
                'Me.TD_Y_7.ColSpan = 3
                Me.TD_Y_7.Attributes("class") = "col-md-10 text-left"
                '打鼾
                Me.TD_G_7_1.Visible = False
                Me.TD_G_7_2.Visible = False

                '身心狀況
                Me.TR_G_8.Visible = False

                '修持法門
                '緊急聯絡人
                Me.TR_G_9.Visible = False
                Me.TR_G_10.Visible = False
                Me.TR_G_11.Visible = False
                Me.TR_G_12.Visible = False
                Me.TR_G_13.Visible = False
                Me.TR_G_14.Visible = False
                Me.TR_G_15_T.Visible = False
                Me.TR_G_15.Visible = False
            Else
                '黃色區+ 灰色區
                'show all so do nothing
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub

    '二號畫面Button
    Protected Sub btn_Confirm_Click(sender As Object, e As EventArgs) Handles btn_Confirm.Click
        Try
            For i As Integer = 0 To Me.RP_Page2.Items.Count - 1
                Dim rbData As RadioButton = Me.RP_Page2.Items(i).FindControl("rbData")
                Dim lbl_MB_MEMSEQ As Label = Me.RP_Page2.Items(i).FindControl("lbl_MB_MEMSEQ")

                If rbData.Checked Then
                    GoPage3(lbl_MB_MEMSEQ.Text.Trim)
                    Exit Sub
                End If
            Next

            UIUtility.alert("請選擇報名者!")
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    '最終確定Button
    Protected Sub btn_Save_Click(sender As Object, e As EventArgs) Handles btn_Save.Click
        Try
            Dim sMB_APV As String = String.Empty
            Dim sMB_BEGIN As String = String.Empty
            Using m_DBManager As com.Azion.NET.VB.DatabaseManager = UIShareFun.getDataBaseManager
                Dim MB_CLASSList As New MB_CLASSList(m_DBManager)
                MB_CLASSList.LoadByMB_SEQ(CDec(Me.m_sCLASS))
                If MB_CLASSList.getCurrentDataSet.Tables(0).Rows.Count > 0 Then
                    Dim ROW As DataRow = MB_CLASSList.getCurrentDataSet.Tables(0).Rows(0)
                    sMB_APV = ROW("MB_APV").ToString
                    sMB_BEGIN = ROW("MB_BEGIN").ToString
                End If
            End Using
            '===================================檢核===================================
            If Me.RP_BATCH.Items.Count > 1 Then
                Dim iBATCH As Decimal = 0
                For Each ITEM As RepeaterItem In Me.RP_BATCH.Items
                    Dim CB_BATCH As CheckBox = ITEM.FindControl("CB_BATCH")
                    If CB_BATCH.Checked Then
                        iBATCH += 1
                    End If
                Next
                If iBATCH = 0 Then
                    com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "參加梯次至少勾選一筆")
                    com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "參加梯次至少勾選一筆")
                    Return
                End If
            End If

            '法名/姓名
            If txt_MB_NAME.Text.Trim = "" Then
                UIUtility.alert("請輸入:法名/姓名")
                Exit Sub
            Else
                Dim pattern As String = "[\p{P}\p{S}-[._]]"
                Dim mx As MatchCollection = Regex.Matches(txt_MB_NAME.Text.Trim, pattern)
                If mx.Count > 0 Then
                    UIUtility.showJSMsg(Me, "法名/姓名不可輸入標點符號")
                    UIUtility.showErrMsg(Me, "法名/姓名不可輸入標點符號")
                    Exit Sub
                End If

                pattern = "[0-9]"
                Dim mxn As MatchCollection = Regex.Matches(txt_MB_NAME.Text.Trim, pattern)
                If mxn.Count > 0 Then
                    UIUtility.showJSMsg(Me, "法名/姓名不可輸入數字")
                    UIUtility.showErrMsg(Me, "法名/姓名不可輸入數字")
                    Exit Sub
                End If
            End If
            '性別
            If rbtList_MB_SEX.SelectedValue = "" Then
                UIUtility.alert("請選擇:性別")
                Exit Sub
            End If
            '出生年月日
            If IsNumeric(txt_MB_BIRTH_YYY.Text) Then
                If Not IsDate(txt_MB_BIRTH_YYY.Text & "/" & txt_MB_BIRTH_MM.Text & "/" & txt_MB_BIRTH_DD.Text) Then
                    UIUtility.alert("請輸入正確之出生年月日!")
                    Exit Sub
                End If
            Else
                UIUtility.alert("請輸入正確之出生年月日!")
                Exit Sub
            End If
            '出家眾
            If sMB_APV = "1" Then
                If Not rbt_MB_MONK_Y.Checked And Not rbt_MB_MONK_N.Checked Then
                    UIUtility.alert("請選擇是否為出家眾!")
                    Exit Sub
                End If
            End If
            '手機
            If Utility.isValidateData(txt_MB_MOBIL.Text) Then
                If Not IsNumeric(txt_MB_MOBIL.Text) Then
                    com.Azion.EloanUtility.UIUtility.alert("手機號碼應為10個數字【範例:0933123456】，請檢查")
                    com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "手機號碼應為10個數字【範例:0933123456】，請檢查")
                    Exit Sub
                ElseIf IsNumeric(txt_MB_MOBIL.Text) AndAlso txt_MB_MOBIL.Text.Trim.Length <> 10 Then
                    com.Azion.EloanUtility.UIUtility.alert("手機號碼應為10個數字【範例:0933123456】，請檢查")
                    com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "手機號碼應為10個數字【範例:0933123456】，請檢查")
                    Exit Sub
                End If
            ElseIf Not Utility.isValidateData(Me.txt_MB_MOBIL.Text) AndAlso Not Utility.isValidateData(Me.txt_MB_TEL.Text) Then
                com.Azion.EloanUtility.UIUtility.alert("請輸入:手機或電話")
                com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請輸入:手機或電話")
                Exit Sub
            End If
            '電話
            If Utility.isValidateData(Me.txt_MB_TEL.Text) Then
                If Not IsNumeric(txt_MB_TEL.Text) Then
                    com.Azion.EloanUtility.UIUtility.alert("電話號碼應為數字，請檢查")
                    com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "電話號碼應為數字，請檢查")
                    Exit Sub
                End If
            End If
            'If Not IsNumeric(txt_MB_MOBIL.Text) Or Not txt_MB_MOBIL.Text.Trim.Length = 10 Then
            '    UIUtility.alert("請輸入:手機或電話")
            '    Exit Sub
            'End If
            'e-mail
            If Not com.Azion.EloanUtility.ValidateUtility.isValidateData(txt_MB_EMAIL.Text.Trim) Then
                UIUtility.alert("請輸入:e-mail")
                Exit Sub
            End If
            '身分證字號
            If sMB_APV = "1" Then
                If Not com.Azion.EloanUtility.ValidateUtility.isValidateData(txt_MB_ID.Text.Trim) Then
                    UIUtility.alert("請輸入:身分證字號")
                    Exit Sub
                End If
            End If
            '通訊地址
            If sMB_APV = "1" Then
                If dd_MB_CITY.SelectedValue = "請選擇" Or dd_MB_VLG.SelectedValue = "請選擇" Or txt_MB_ADDR.Text.Trim = "" Then
                    UIUtility.alert("請輸入:通訊地址")
                    Exit Sub
                End If
            End If
            '1041201大惠師不要必填
            '宗教信仰
            'If dd_MB_RELIGION.SelectedValue = "請選擇" Then
            '    UIUtility.alert("請輸入:宗教信仰")
            '    Exit Sub
            'End If
            '身心狀況
            If Me.TR_G_8.Visible = True Then
                Dim iMB_SICK As Integer = 0
                For Each Item As RepeaterItem In Me.dl_MB_SICK.Items
                    Dim cb_MB_SICK As CheckBox = Item.FindControl("cb_MB_SICK")

                    If cb_MB_SICK.Checked Then
                        iMB_SICK += 1
                    End If
                Next
                If iMB_SICK = 0 Then
                    UIUtility.showJSMsg(Me, "請選擇身心狀況")
                    UIUtility.showErrMsg(Me, "請選擇身心狀況")
                    Exit Sub
                End If

                For i As Integer = 0 To dl_MB_SICK.Items.Count - 1
                    Dim cb_MB_SICK As CheckBox = dl_MB_SICK.Items(i).FindControl("cb_MB_SICK")

                    If cb_MB_SICK.Checked Then
                        Dim txt_MB_SICK As TextBox = dl_MB_SICK.Items(i).FindControl("txt_MB_SICK")
                        Dim sMB_SICK As String = String.Empty


                        If cb_MB_SICK.Text = "過敏" Then
                            sMB_SICK = txt_MB_SICK.Text
                        ElseIf cb_MB_SICK.Text = "開刀" Then
                            sMB_SICK = txt_MB_SICK.Text
                        ElseIf cb_MB_SICK.Text = "其他" Then
                            sMB_SICK = txt_MB_SICK.Text
                        End If

                        If Utility.isValidateData(sMB_SICK) Then
                            Dim iCnt As Decimal = 0
                            iCnt = System.Text.Encoding.GetEncoding("Big5").GetByteCount(sMB_SICK)
                            If iCnt > 50 Then
                                com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "身心狀況【" & cb_MB_SICK.Text & "】說明輸入最多50個字元，目前輸入共【" & iCnt & "】字元")
                                com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "身心狀況【" & cb_MB_SICK.Text & "】說明輸入最多50個字元，目前輸入共【" & iCnt & "】字元")
                                Return
                            End If
                        End If
                    End If
                Next
            End If

            '打鼾
            If sMB_APV = "1" Then
                If rbtList_MB_SNORE.SelectedValue.Trim = "" Then
                    UIUtility.alert("請輸入:打鼾")
                    Exit Sub
                End If
            End If
            '緊急聯絡人
            If sMB_APV = "1" Then
                If Not Utility.isValidateData(Trim(Me.txt_MB_EMGCONT.Text)) Then
                    UIUtility.alert("請輸入:緊急聯絡人")
                    Exit Sub
                End If

                If Not IsNumeric(Me.txt_MB_CONTMOBIL.Text) Then
                    UIUtility.alert("請輸入:緊急聯絡人電話/手機【手機或電話請輸入數字】")
                    Exit Sub
                End If
            End If

            '1041201大惠師不要必填
            '參加動機或目的
            'If Not Utility.isValidateData(Trim(Me.MB_OBJECT.Text)) Then
            '    UIUtility.alert("請輸入:參加動機或目的")
            '    Return
            'End If
            '是否曾參加過本中心課程
            If Not Utility.isValidateData(Me.JOINMBSC.SelectedValue) Then
                UIUtility.alert("請選擇:是否曾參加過本中心課程")
                Return
            End If

            '是否初學者
            If sMB_BEGIN = "Y" Then
                If Not Utility.isValidateData(Me.MB_BEGIN.SelectedValue) Then
                    UIUtility.alert("請選擇:是否初學者")
                    Return
                End If
                '每次禪坐時間
                If Me.MB_BEGIN.SelectedValue = "N" Then
                    If Not Utility.isValidateData(Me.MB_SITIME.Text) Then
                        UIUtility.alert("請輸入:每次禪坐時間")
                        Return
                    End If

                    If Not IsNumeric(Me.MB_SITIME.Text) Then
                        UIUtility.alert("每次禪坐時間請輸入數字")
                        Return
                    ElseIf CDec(Me.MB_SITIME.Text) > 999 Then
                        UIUtility.alert("每次禪坐時間不可大於999分鐘")
                        Return
                    End If
                End If
            End If

            '開始寫入資料===================================================================
            '法名
            Using m_DBManager As com.Azion.NET.VB.DatabaseManager = UIShareFun.getDataBaseManager
                Dim MB_MEMBER As New MB_MEMBER(m_DBManager)

                With MB_MEMBER
                    If Not IsNumeric(lbl_MB_MEMSEQ.Text.Trim) Then
                        '取號前先檢核姓名+手機,姓名+電話,姓名+eMail是否曾經報名過
                        '如果已經是會員了，就不可再取號，避免會員重複
                        Dim MB_MEMBERList As New MB_MEMBERList(m_DBManager)
                        If Utility.isValidateData(txt_MB_NAME.Text.Trim) AndAlso Utility.isValidateData(txt_MB_MOBIL.Text.Trim) Then
                            MB_MEMBERList.Load_MB_NAME_MB_MOBIL(txt_MB_NAME.Text.Trim, txt_MB_MOBIL.Text.Trim)
                            If MB_MEMBERList.getCurrentDataSet.Tables(0).Rows.Count > 0 Then
                                com.Azion.EloanUtility.UIUtility.showJSMsg("【" & txt_MB_NAME.Text.Trim & "】【" & txt_MB_MOBIL.Text.Trim & "】已經是會員了，請再確認看看\n如果要修改手機號碼，請由頁首下拉【會員系統->入會申請單】修改")
                                com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "【" & txt_MB_NAME.Text.Trim & "】【" & txt_MB_MOBIL.Text.Trim & "】已經是會員了，請再確認看看<BR/>如果要修改手機號碼，請由頁首下拉【會員系統->入會申請單】修改")

                                Return
                            End If
                        End If

                        If Utility.isValidateData(txt_MB_NAME.Text.Trim) AndAlso Utility.isValidateData(txt_MB_TEL.Text.Trim) Then
                            Dim sMB_TEL As String = String.Empty
                            If Utility.isValidateData(Trim(Me.txt_MB_TEL_ZIP.Text)) Then
                                sMB_TEL = Trim(Me.txt_MB_TEL_ZIP.Text) & "-" & Trim(Me.txt_MB_TEL.Text)
                            Else
                                sMB_TEL = Trim(Me.txt_MB_TEL.Text)
                            End If

                            MB_MEMBERList.clear()
                            MB_MEMBERList.Load_MB_NAME_MB_TEL(txt_MB_NAME.Text.Trim, sMB_TEL)
                            If MB_MEMBERList.getCurrentDataSet.Tables(0).Rows.Count > 0 Then
                                com.Azion.EloanUtility.UIUtility.showJSMsg("【" & txt_MB_NAME.Text.Trim & "】【" & sMB_TEL & "】已經是會員了，請再確認看看\n如果要修改電話號碼，請由頁首下拉【會員系統->入會申請單】修改")
                                com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "【" & txt_MB_NAME.Text.Trim & "】【" & sMB_TEL & "】已經是會員了，請再確認看看<BR/>如果要修改電話號碼，請由頁首下拉【會員系統->入會申請單】修改")

                                Return
                            End If
                        End If

                        If Utility.isValidateData(txt_MB_NAME.Text.Trim) AndAlso Utility.isValidateData(txt_MB_EMAIL.Text.Trim) Then
                            MB_MEMBERList.clear()
                            MB_MEMBERList.Load_MB_NAME_MB_EMAIL(txt_MB_NAME.Text.Trim, txt_MB_EMAIL.Text.Trim)
                            If MB_MEMBERList.getCurrentDataSet.Tables(0).Rows.Count > 0 Then
                                com.Azion.EloanUtility.UIUtility.showJSMsg("【" & txt_MB_NAME.Text.Trim & "】【" & txt_MB_EMAIL.Text.Trim & "】已經是會員了，請再確認看看\n如果要修改e-MAIL，請由頁首下拉【會員系統->入會申請單】修改")
                                com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "【" & txt_MB_NAME.Text.Trim & "】【" & txt_MB_EMAIL.Text.Trim & "】已經是會員了，請再確認看看<BR/>如果要修改e-MAIL，請由頁首下拉【會員系統->入會申請單】修改")

                                Return
                            End If
                        End If

                        lbl_MB_MEMSEQ.Text = UIShareFun.getVadID(m_DBManager)
                    End If

                    .loadByPK(CInt(lbl_MB_MEMSEQ.Text.Trim))
                    .setAttribute("MB_MEMSEQ", CInt(lbl_MB_MEMSEQ.Text.Trim))
                    .setAttribute("MB_IDENTIFY", "A")

                    '法名/姓名
                    .setAttribute("MB_NAME", txt_MB_NAME.Text.Trim)
                    '性別
                    If Utility.isValidateData(rbtList_MB_SEX.SelectedValue) Then
                        .setAttribute("MB_SEX", rbtList_MB_SEX.SelectedValue)
                    End If
                    '出生年月日
                    If IsNumeric(txt_MB_BIRTH_YYY.Text) AndAlso IsDate(CDate(txt_MB_BIRTH_YYY.Text & "/" & txt_MB_BIRTH_MM.Text & "/" & txt_MB_BIRTH_DD.Text)) Then
                        .setAttribute("MB_BIRTH", CDate(txt_MB_BIRTH_YYY.Text & "/" & txt_MB_BIRTH_MM.Text & "/" & txt_MB_BIRTH_DD.Text))
                    Else
                        .setAttribute("MB_BIRTH", Nothing)
                    End If
                    '出家眾
                    If rbt_MB_MONK_Y.Checked Then
                        .setAttribute("MB_MONK", "1")
                    ElseIf rbt_MB_MONK_N.Checked Then
                        .setAttribute("MB_MONK", "0")
                    End If
                    '法名
                    .setAttribute("MB_MONKNAME", txt_MB_MONKNAME.Text)
                    '剃度/皈依恩師/戒師
                    .setAttribute("MB_MONKTECH", txt_MB_MONKTECH.Text)
                    '傳承
                    If Utility.isValidateData(dd_MB_EDUTYPE.SelectedValue) Then
                        .setAttribute("MB_EDUTYPE", dd_MB_EDUTYPE.SelectedValue)
                    End If
                    '常住/親近道場
                    .setAttribute("MB_MONKPLACE", txt_MB_MONKPLACE.Text)
                    '戒別
                    If Utility.isValidateData(dd_MB_MONKTYPE.SelectedValue) Then
                        .setAttribute("MB_MONKTYPE", dd_MB_MONKTYPE.SelectedValue)
                    End If
                    '受戒日期
                    If IsNumeric(txt_MB_MONKDATE_YYY.Text) Then
                        If IsDate((txt_MB_MONKDATE_YYY.Text) & "/" & txt_MB_MONKDATE_MM.Text & "/" & txt_MB_MONKDATE_DD.Text) Then
                            .setAttribute("MB_MONKDATE", CDate((txt_MB_MONKDATE_YYY.Text) & "/" & txt_MB_MONKDATE_MM.Text & "/" & txt_MB_MONKDATE_DD.Text))
                        End If
                    End If
                    '受戒地點
                    .setAttribute("MB_MONKPLACE1", txt_MB_MONKPLACE1.Text)
                    '手機
                    .setAttribute("MB_MOBIL", txt_MB_MOBIL.Text)
                    '電話
                    If IsNumeric(txt_MB_TEL_ZIP.Text) And IsNumeric(txt_MB_TEL.Text) Then
                        .setAttribute("MB_TEL", txt_MB_TEL_ZIP.Text & "-" & txt_MB_TEL.Text)
                    ElseIf Not IsNumeric(txt_MB_TEL_ZIP.Text) AndAlso IsNumeric(txt_MB_TEL.Text) Then
                        .setAttribute("MB_TEL",txt_MB_TEL.Text)
                    End If
                    'E-mail
                    .setAttribute("MB_EMAIL", txt_MB_EMAIL.Text)
                    '身分證字號
                    .setAttribute("MB_ID", txt_MB_ID.Text)
                    '學歷
                    If Utility.isValidateData(dd_MB_EDU.SelectedValue) Then
                        .setAttribute("MB_EDU", dd_MB_EDU.SelectedValue)
                    End If
                    '介紹人/訊息來源
                    .setAttribute("MB_REFER", txt_MB_REFER.Text)
                    '通訊地址
                    If dd_MB_CITY.SelectedIndex > 0 Then
                        .setAttribute("MB_CITY", dd_MB_CITY.SelectedItem.Text)
                    End If
                    If dd_MB_VLG.SelectedIndex > 0 Then
                        .setAttribute("MB_VLG", dd_MB_VLG.SelectedItem.Text)
                    End If
                    If dd_MB_VLG.SelectedIndex > 0 Then
                        .setAttribute("MB_AREA", dd_MB_VLG.SelectedValue)
                    End If
                    .setAttribute("MB_ADDR", txt_MB_ADDR.Text)
                    '戶籍地址
                    If dd_MB_CITY1.SelectedIndex > 0 AndAlso dd_MB_VLG1.SelectedIndex > 0 AndAlso txt_MB_ADDR1.Text.Trim <> "" Then
                        .setAttribute("MB_CITY1", dd_MB_CITY1.SelectedItem.Text)
                        .setAttribute("MB_VLG1", dd_MB_VLG1.SelectedItem.Text)
                        .setAttribute("MB_ADDR1", txt_MB_ADDR1.Text)
                    End If
                    '語言
                    Dim sMB_LANG As String = String.Empty
                    For i As Integer = 0 To dl_MB_LANG.Items.Count - 1
                        Dim hid_MB_LANG As HiddenField = dl_MB_LANG.Items(i).FindControl("hid_MB_LANG")
                        Dim cb_MB_LANG As CheckBox = dl_MB_LANG.Items(i).FindControl("cb_MB_LANG")
                        Dim txt_MB_LANG As TextBox = dl_MB_LANG.Items(i).FindControl("txt_MB_LANG")

                        If cb_MB_LANG.Checked Then
                            sMB_LANG = sMB_LANG & hid_MB_LANG.Value & ","

                            If cb_MB_LANG.Text = "其他" Then
                                .setAttribute("MB_OLANG", txt_MB_LANG.Text)
                            End If
                        End If
                    Next
                    If sMB_LANG.Length > 1 Then
                        sMB_LANG = Left(sMB_LANG, sMB_LANG.Length - 1)
                    End If
                    .setAttribute("MB_LANG", sMB_LANG)
                    '專長
                    .setAttribute("MB_SPECIAL", dd_MB_SPECIAL.SelectedValue)
                    '職業
                    .setAttribute("MB_PROFESSION", dd_MB_JOB.SelectedValue)
                    '職稱
                    .setAttribute("MB_JOBTITLE", dd_MB_JOBTITLE.SelectedValue)
                    '宗教信仰
                    .setAttribute("MB_RELIGION", dd_MB_RELIGION.SelectedValue)
                    '打鼾
                    .setAttribute("MB_SNORE", rbtList_MB_SNORE.SelectedValue)
                    '身心狀況
                    Dim sMB_SICK As String = String.Empty
                    For i As Integer = 0 To dl_MB_SICK.Items.Count - 1
                        Dim hid_MB_SICK As HiddenField = dl_MB_SICK.Items(i).FindControl("hid_MB_SICK")
                        Dim cb_MB_SICK As CheckBox = dl_MB_SICK.Items(i).FindControl("cb_MB_SICK")
                        Dim txt_MB_SICK As TextBox = dl_MB_SICK.Items(i).FindControl("txt_MB_SICK")

                        If cb_MB_SICK.Checked Then
                            sMB_SICK = sMB_SICK & hid_MB_SICK.Value & ","

                            If cb_MB_SICK.Text = "過敏" Then
                                .setAttribute("MB_ALLERGY", txt_MB_SICK.Text)
                            ElseIf cb_MB_SICK.Text = "開刀" Then
                                .setAttribute("MB_OPERATE", txt_MB_SICK.Text)
                            ElseIf cb_MB_SICK.Text = "其他" Then
                                .setAttribute("MB_OSICK", txt_MB_SICK.Text)
                            End If
                        End If
                    Next
                    If sMB_SICK.Length > 1 Then
                        sMB_SICK = Left(sMB_SICK, sMB_SICK.Length - 1)
                    End If
                    .setAttribute("MB_SICK", sMB_SICK)
                    '您曾經修持過毗婆舍那禪法嗎
                    If rbt_MB_PIPOSHENA_Y.Checked Then
                        .setAttribute("MB_PIPOSHENA", "1")
                    ElseIf rbt_MB_PIPOSHENA_N.Checked Then
                        .setAttribute("MB_PIPOSHENA", "0")
                    End If
                    '指導老師
                    .setAttribute("MB_TEACH", txt_MB_TEACH.Text)
                    '您目前的修持法門
                    .setAttribute("MB_FAMENNIAN", txt_MB_FAMENNIAN.Text)
                    '您過去曾經參加過七天以上的禪修嗎
                    If rbt_MB_OVER7DAY_Y.Checked Then
                        .setAttribute("MB_OVER7DAY", "1")
                    ElseIf rbt_MB_OVER7DAY_N.Checked Then
                        .setAttribute("MB_OVER7DAY", "0")
                    End If
                    '地點
                    .setAttribute("MB_PLACE", txt_MB_PLACE.Text)
                    '緊急聯絡人 - 法名/姓名
                    .setAttribute("MB_EMGCONT", txt_MB_EMGCONT.Text)
                    '緊急聯絡人 - 電話/手機
                    .setAttribute("MB_CONTMOBIL", txt_MB_CONTMOBIL.Text)
                    '捐助物資/捐款
                    .setAttribute("MB_AMTMEMO", MB_AMTMEMO.Text)
                    '是否曾參加過本中心課程
                    .setAttribute("JOINMBSC", Me.JOINMBSC.SelectedValue)
                    '公司/學校系級
                    .setAttribute("SCHOOL", Me.SCHOOL.Text)
                    '變更日期
                    .setAttribute("CHGDATE", Now)
                    '變更人員
                    .setAttribute("CHGUID", m_sUSERID)
                    '是否初學者
                    .setAttribute("MB_BEGIN", Me.MB_BEGIN.SelectedValue)
                    '每次禪坐時間
                    If Me.MB_BEGIN.SelectedValue = "Y" Then
                        .setAttribute("MB_SITIME", 0)
                    Else
                        If IsNumeric(Me.MB_SITIME.Text) Then
                            .setAttribute("MB_SITIME", CInt(Me.MB_SITIME.Text))
                        Else
                            .setAttribute("MB_SITIME", 0)
                        End If
                    End If
                    .save()
                End With

                Dim MB_MEMMAP As New MB_MEMMAP(m_DBManager)
                MB_MEMMAP.LoadByPK(MB_MEMBER.getDecimal("MB_MEMSEQ"))
                Dim sMB_BKSEQ As String = String.Empty
                If Not MB_MEMMAP.isLoaded Then
                    'MB_BKSEQ	decimal(16,0)		YES				select,insert,update,references	台銀編號
                    '7106800010100723
                    '710680 + BBB + 01 + 5 碼序號(MB_MEMSEQ )不足錢補0
                    'BBB==> 區代碼 用AP_CODE.VALUE (請將區代碼 1==> 001其餘比照)
                    Dim sBBB As String = String.Empty
                    Select Case MB_MEMBER.getString("MB_AREA")
                        Case "A"
                            sBBB = "001"
                        Case "B"
                            sBBB = "002"
                        Case "C"
                            sBBB = "003"
                        Case "D"
                            sBBB = "004"
                        Case Else
                            sBBB = "001"
                    End Select
                    sMB_BKSEQ = "710680" & sBBB & "01" & Utility.FillZero(MB_MEMBER.getDecimal("MB_MEMSEQ"), 5)
                    MB_MEMMAP.setAttribute("MB_BKSEQ", sMB_BKSEQ)
                Else
                    sMB_BKSEQ = MB_MEMMAP.getString("MB_BKSEQ")
                End If
                'MB_NAME	varchar(50)	utf8_general_ci	YES				select,insert,update,references	會員名稱
                MB_MEMMAP.setAttribute("MB_NAME", Trim(MB_MEMBER.getString("MB_NAME")))
                'MB_MEMSEQ	decimal(7,0)		NO	PRI	0		select,insert,update,references	會員編號
                MB_MEMMAP.setAttribute("MB_MEMSEQ", MB_MEMBER.getDecimal("MB_MEMSEQ"))
                MB_MEMMAP.save()

                '課程檔寫入
                Dim MB_CLASS As New MB_CLASS(m_DBManager)
                Dim MB_MEMCLASS As New MB_MEMCLASS(m_DBManager)
                Dim MB_MEMCLASSList As New MB_MEMCLASSList(m_DBManager)

                Dim MB_CLASSList As New MB_CLASSList(m_DBManager)
                MB_CLASSList.LoadByMB_SEQ(CDec(Me.m_sCLASS))
                Dim ROW_CLASS As DataRow = Nothing
                If MB_CLASSList.getCurrentDataSet.Tables(0).Rows.Count > 0 Then
                    ROW_CLASS = MB_CLASSList.getCurrentDataSet.Tables(0).Rows(0)
                End If
                Dim AL_MB_BATCH As New ArrayList
                If Me.TR_BATCH.Visible Then
                    For Each ITEM As RepeaterItem In Me.RP_BATCH.Items
                        Dim CB_BATCH As CheckBox = ITEM.FindControl("CB_BATCH")
                        If CB_BATCH.Checked Then
                            Dim MB_BATCH As Literal = ITEM.FindControl("MB_BATCH")
                            MB_MEMCLASSList.clear()
                            MB_MEMCLASSList.setSQLCondition(" AND IFNULL(A.MB_FWMK,' ') NOT IN ('3','4', '5') AND IFNULL(B.MB_ELECT,' ')='1' AND IFNULL(B.MB_RESP,' ')<>'N' ")
                            MB_MEMCLASSList.loadByMB_SEQ(m_sCLASS, CDec(MB_BATCH.Text))

                            MB_CLASS.clear()
                            MB_CLASS.loadByPK(CDec(Me.m_sCLASS), CDec(MB_BATCH.Text))
                            If MB_MEMCLASSList.size < (MB_CLASS.getInt("MB_FULL") + MB_CLASS.getInt("MB_WAIT")) Then
                                AL_MB_BATCH.Add(CDec(MB_BATCH.Text))
                            Else
                                Dim ROW_CHK() As DataRow = MB_MEMCLASSList.getCurrentDataSet.Tables(0).Select("MB_MEMSEQ=" & MB_MEMBER.getDecimal("MB_MEMSEQ"))
                                If IsNothing(ROW_CHK) OrElse ROW_CHK.Length = 0 Then
                                    com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "梯次【" & MB_BATCH.Text & "】額滿,不可報名!請取消勾選")
                                    com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "梯次【" & MB_BATCH.Text & "】額滿,不可報名!請取消勾選")
                                    Return
                                Else
                                    AL_MB_BATCH.Add(CDec(MB_BATCH.Text))
                                End If
                            End If
                        End If
                    Next
                Else
                    If Not IsNothing(ROW_CLASS) Then
                        MB_MEMCLASSList.clear()
                        MB_MEMCLASSList.setSQLCondition(" AND IFNULL(A.MB_FWMK,' ') NOT IN ('3','4', '5') AND IFNULL(B.MB_ELECT,' ')='1' AND IFNULL(B.MB_RESP,' ')<>'N' ")
                        MB_MEMCLASSList.loadByMB_SEQ(m_sCLASS, ROW_CLASS("MB_BATCH"))

                        If MB_MEMCLASSList.size < (Utility.CheckNumNull(ROW_CLASS("MB_FULL")) + Utility.CheckNumNull(ROW_CLASS("MB_WAIT"))) Then
                            AL_MB_BATCH.Add(ROW_CLASS("MB_BATCH"))
                        Else
                            Dim ROW_CHK() As DataRow = MB_MEMCLASSList.getCurrentDataSet.Tables(0).Select("MB_MEMSEQ=" & MB_MEMBER.getDecimal("MB_MEMSEQ"))
                            If IsNothing(ROW_CHK) OrElse ROW_CHK.Length = 0 Then
                                com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "額滿,不可報名!")
                                com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "額滿,不可報名!")
                                Return
                            Else
                                AL_MB_BATCH.Add(ROW_CLASS("MB_BATCH"))
                            End If
                        End If
                    End If
                End If

                Dim MB_MEMBATCHList As New MB_MEMBATCHList(m_DBManager)
                MB_MEMBATCHList.DeleteBySEQ(lbl_MB_MEMSEQ.Text, m_sCLASS)
                If AL_MB_BATCH.Count > 0 Then
                    MB_MEMCLASS.clear()
                    MB_MEMCLASS.LoadByPK(lbl_MB_MEMSEQ.Text, m_sCLASS)
                    With MB_MEMCLASS
                        .setAttribute("MB_MEMSEQ", lbl_MB_MEMSEQ.Text)
                        .setAttribute("MB_SEQ", m_sCLASS)
                        '參加動機或目的
                        .setAttribute("MB_OBJECT", Me.MB_OBJECT.Text)

                        If Not MB_MEMCLASS.isLoaded Then
                            If sMB_APV = "2" Then
                                .setAttribute("MB_FWMK", "1")
                            Else
                                .setAttribute("MB_FWMK", "0")
                            End If

                            .setAttribute("MB_CREDATETIME", Now)
                            .setAttribute("MB_CHGDATE", Now)
                        Else
                            If MB_MEMCLASS.getString("MB_FWMK") = "3" Then
                                .setAttribute("MB_FWMK", "1")
                                .setAttribute("MB_CREDATETIME", Now)
                                .setAttribute("MB_CHGDATE", Now)
                            Else
                                .setAttribute("MB_CHGDATE", Now)
                            End If
                        End If

                        .save()
                    End With

                    For Each ITEM As RepeaterItem In Me.RP_BATCH.Items
                        Dim CB_BATCH As CheckBox = ITEM.FindControl("CB_BATCH")
                        Dim MB_BATCH As Literal = ITEM.FindControl("MB_BATCH")
                        Dim MB_MEMBATCH As New MB_MEMBATCH(m_DBManager)
                        MB_MEMBATCH.setAttribute("MB_MEMSEQ", lbl_MB_MEMSEQ.Text)
                        MB_MEMBATCH.setAttribute("MB_SEQ", m_sCLASS)
                        MB_MEMBATCH.setAttribute("MB_BATCH", MB_BATCH.Text)
                        'MB_ELECT	char(1)	latin1_swedish_ci	YES				select,insert,update,references	選取註記;1: 選取;null:未選
                        If AL_MB_BATCH.Contains(CDec(MB_BATCH.Text)) Then
                            MB_MEMBATCH.setAttribute("MB_ELECT", "1")
                        Else
                            MB_MEMBATCH.setAttribute("MB_ELECT", Nothing)
                        End If
                        'MB_CHKFLAG	char(1)	latin1_swedish_ci	YES				select,insert,update,references	是否核准;1:核准;2.不准;3.核准未出席
                        If sMB_APV = "2" Then
                            MB_MEMBATCH.setAttribute("MB_CHKFLAG", "1")
                        End If
                        'MB_RESP	char(1)	latin1_swedish_ci	YES				select,insert,update,references	回信是否出席;Y:出席;N:不出席
                        MB_MEMBATCH.save()
                    Next

                    '發e-Mail
                    If Not MB_MEMCLASS.isLoaded Then
                        Dim AL_WAIT As New ArrayList
                        Dim AL_RIGHT As New ArrayList
                        For i As Integer = 0 To AL_MB_BATCH.Count - 1
                            MB_MEMCLASSList.clear()
                            MB_MEMCLASSList.setSQLCondition(" AND IFNULL(A.MB_FWMK,' ') NOT IN ('3','4', '5') AND IFNULL(B.MB_ELECT,' ')='1' AND IFNULL(B.MB_RESP,' ')<>'N' ")
                            MB_MEMCLASSList.loadByMB_SEQ(m_sCLASS, AL_MB_BATCH.Item(i))

                            If Not MB_MEMCLASS.isLoaded Then
                                MB_CLASS.clear()
                                MB_CLASS.loadByPK(CDec(Me.m_sCLASS), AL_MB_BATCH.Item(i))

                                If MB_MEMCLASSList.size > MB_CLASS.getInt("MB_FULL") Then
                                    '備取
                                    AL_WAIT.Add(AL_MB_BATCH.Item(i))
                                Else
                                    '正取
                                    AL_RIGHT.Add(AL_MB_BATCH.Item(i))
                                End If
                            End If
                        Next

                        Try
                            Dim mbMEMBER As New MB_MEMBER(m_DBManager)
                            If mbMEMBER.loadByPK(lbl_MB_MEMSEQ.Text) AndAlso Utility.isValidateData(mbMEMBER.getAttribute("MB_EMAIL")) Then
                                Dim sMailTos() As String = {Trim(mbMEMBER.getString("MB_EMAIL"))}

                                Dim sWAIT_BATCH As String = String.Empty
                                If AL_WAIT.Count > 0 AndAlso ROW_CLASS("MB_APV").ToString <> "1" AndAlso Me.TR_BATCH.Visible Then
                                    For i As Integer = 0 To AL_WAIT.Count - 1
                                        sWAIT_BATCH &= AL_WAIT.Item(i) & "，"
                                    Next
                                    If Utility.isValidateData(sWAIT_BATCH) Then
                                        sWAIT_BATCH = "第" & Left(sWAIT_BATCH, sWAIT_BATCH.Length - 1) & "梯次"
                                    End If
                                End If

                                Dim sRIGHT_BATCH As String = String.Empty
                                If AL_RIGHT.Count > 0 AndAlso ROW_CLASS("MB_APV").ToString <> "1" AndAlso Me.TR_BATCH.Visible Then
                                    For i As Integer = 0 To AL_RIGHT.Count - 1
                                        sRIGHT_BATCH &= AL_RIGHT.Item(i) & "，"
                                    Next
                                    If Utility.isValidateData(sRIGHT_BATCH) Then
                                        sRIGHT_BATCH = "第" & Left(sRIGHT_BATCH, sRIGHT_BATCH.Length - 1) & "梯次"
                                    End If
                                End If

                                Dim sMailSub As String = String.Empty
                                If ROW_CLASS("MB_APV").ToString = "2" Then
                                    If AL_WAIT.Count > 0 Then
                                        sMailSub = ROW_CLASS("MB_CLASS_NAME").ToString & sWAIT_BATCH & " 備取通知"
                                        Dim sMailBody As String = String.Empty
                                        sMailBody = Me.getMailBody(m_DBManager, lbl_MB_MEMSEQ.Text, Trim(mbMEMBER.getString("MB_NAME")), ROW_CLASS, True, sWAIT_BATCH)
                                        com.Azion.EloanUtility.NetUtility.GMail_Send(sMailTos, Nothing, sMailSub, sMailBody, True, Nothing, False)
                                    End If

                                    If AL_RIGHT.Count > 0 Then
                                        sMailSub = ROW_CLASS("MB_CLASS_NAME").ToString & sRIGHT_BATCH & " 錄取通知 請務必回覆是否出席"
                                        Dim sMailBody As String = String.Empty
                                        sMailBody = Me.getMailBody(m_DBManager, lbl_MB_MEMSEQ.Text, Trim(mbMEMBER.getString("MB_NAME")), ROW_CLASS, False, sRIGHT_BATCH)
                                        com.Azion.EloanUtility.NetUtility.GMail_Send(sMailTos, Nothing, sMailSub, sMailBody, True, Nothing, False)
                                    End If
                                Else
                                    sMailSub = ROW_CLASS("MB_CLASS_NAME").ToString
                                    Dim sMailBody As String = String.Empty
                                    sMailBody = Me.getMailBody(m_DBManager, lbl_MB_MEMSEQ.Text, Trim(mbMEMBER.getString("MB_NAME")), ROW_CLASS, False, String.Empty)
                                    com.Azion.EloanUtility.NetUtility.GMail_Send(sMailTos, Nothing, sMailSub, sMailBody, True, Nothing, False)
                                End If

                                Dim sMSG As String = String.Empty
                                sMSG = "請至您的信箱收信\n如未收到\n請確認是否被系統自動分入垃圾信箱,\n若確定信箱內找不到信件,\n麻煩您致電中心法工確認"
                                UIUtility.alert(sMSG)
                            End If
                        Catch ex As Exception

                        End Try

                        'UIUtility.alert("報名儲存成功!")

                        UIShareFun.showJSMsgGoList(Me, "報名儲存成功!")
                    Else
                        'UIUtility.alert("會員資料已修改!")

                        UIShareFun.showJSMsgGoList(Me, "會員資料已修改!")
                    End If
                End If
            End Using
        Catch ex As System.Threading.ThreadAbortException
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    Function getMailBody(ByVal dbManager As DatabaseManager, ByVal iMB_MEMSEQ As Decimal, ByVal sAPPNAME As String, ByVal ROW_CLASS As DataRow, ByVal isWait As Boolean, ByVal sBatchAPN As String) As String
        Try
            Dim sb As New StringBuilder

            If ROW_CLASS("MB_APV").ToString = "1" Then
                '發原來報名成功通知函
                sb.Append("<P>")
                sb.Append("親愛的" & sAPPNAME & "吉祥：<BR/>")
                sb.Append("已收到您的線上報名【" & ROW_CLASS("MB_CLASS_NAME").ToString & "】。<BR/>")
                sb.Append("感謝您﹗若獲錄取將於開課前14日內寄發錄取通知。<BR/>")
                sb.Append("若要取消報名，請於").Append(Me.getMB_CDAYS(ROW_CLASS("MB_SDATE"), ROW_CLASS("MB_CDAYS"))).Append("前取消!若超過期限欲取消報名，請聯絡中心<BR/>")
                sb.Append("<SPAN STYLE='color:red'>此信函為系統所發出，請勿直接回覆。</SPAN><BR/>")
                sb.Append("&nbsp;&nbsp;&nbsp;MBSC佛陀原始正法中心 合十")
                sb.Append("</P>")
            ElseIf ROW_CLASS("MB_APV").ToString = "3" Then
                sb.Append("<P>")
                sb.Append("親愛的" & sAPPNAME & "吉祥：<BR/>")
                sb.Append("已收到您的線上報名【" & ROW_CLASS("MB_CLASS_NAME").ToString & "】。<BR/>")
                Dim MB_MEMCLASS As New MB_MEMCLASS(dbManager)
                MB_MEMCLASS.LoadByPK(iMB_MEMSEQ, ROW_CLASS("MB_SEQ"))
                Dim isEarly As Boolean = False
                If IsDate(ROW_CLASS("EARLYDATE").ToString) AndAlso IsNumeric(ROW_CLASS("FAVCHARGE")) AndAlso Utility.isValidateData(MB_MEMCLASS.getAttribute("MB_CREDATETIME")) Then
                    Dim D_EARLYDATE As Date = CDate(ROW_CLASS("EARLYDATE").ToString)
                    Dim D_MB_CREDATETIME As Date = CDate(MB_MEMCLASS.getAttribute("MB_CREDATETIME").ToString)
                    If D_MB_CREDATETIME <= D_EARLYDATE Then
                        isEarly = True
                    End If
                End If
                Dim sMB_CREDATETIME As String = String.Empty
                If Utility.isValidateData(MB_MEMCLASS.getAttribute("MB_CREDATETIME")) Then
                    sMB_CREDATETIME = CDate(MB_MEMCLASS.getAttribute("MB_CREDATETIME").ToString).Year - 1911 & "年" &
                                      CDate(MB_MEMCLASS.getAttribute("MB_CREDATETIME").ToString).Month & "月" &
                                      CDate(MB_MEMCLASS.getAttribute("MB_CREDATETIME").ToString).Day & "日"
                End If
                If isEarly Then
                    sb.Append("<DIV style='color:blue;font-weight:bolde;font-size:14pt'>").Append("本課程您於").Append(sMB_CREDATETIME).Append("報名，符合早鳥優惠價").Append(ROW_CLASS("FAVCHARGE").ToString).Append("元</DIV>")
                Else
                    sb.Append("<DIV style='color:blue;font-weight:bolde;font-size:14pt'>").Append("本課程您於").Append(sMB_CREDATETIME).Append("報名，課程費用").Append(ROW_CLASS("CHARGE").ToString).Append("元</DIV>")
                End If
                sb.Append("<DIV style='color:blue;font-weight:bolde;font-size:14pt'>").Append("請收到本通知於3日內完成匯款後，告知您的 姓名及存款帳號後5碼</DIV>")
                sb.Append("<BR/>若要取消報名，請於").Append(Me.getMB_CDAYS(ROW_CLASS("MB_SDATE"), ROW_CLASS("MB_CDAYS"))).Append("前取消!若超過期限欲取消報名，請聯絡中心<BR/>")
                sb.Append("<SPAN STYLE='color:red'>此信函為系統所發出，請勿直接回覆。</SPAN><BR/>")
                sb.Append("&nbsp;&nbsp;&nbsp;MBSC佛陀原始正法中心 合十")
                sb.Append("</P>")
            Else
                '錄取通知 (參照下表)  (為原來提醒信一再修一下黃色螢光部分)
                sb.Append("<DIV style='text-align:center;font-size:24pt;color:red' >").Append(ROW_CLASS("MB_CLASS_NAME").ToString).Append("</DIV>")
                Dim sForm As String = String.Empty
                If isWait Then
                    sForm = "備取通知"
                Else
                    sForm = "錄取通知"
                End If
                sb.Append("<DIV style='text-align:center;font-size:24pt;color:black;font-weight:bold;text-decoration:underline' >").Append(sForm).Append("</DIV>")
                sb.Append("<DIV style='font-size:12pt;color:black'>").Append("親愛的").Append(sAPPNAME).Append("吉祥：").Append("</DIV>")

                sb.Append("<DIV style='font-size:12pt;color:black;'>").Append("　　  歡迎您報名參加")
                sb.Append("<span style='color:black'>")
                sb.Append(ROW_CLASS("MB_CLASS_NAME").ToString)
                sb.Append("</span>")
                sb.Append("！此通知函，乃透過系統發送。為使您在課程期間能順利進行，並獲得最大收穫，請務必閱讀下列注意事項：").Append("</DIV>")

                If isWait Then
                    '備取
                    sb.Append("<ol type='1' style='font-size:12pt;color:black;font-weight:bold' >")

                    sb.Append("<li>")
                    sb.Append("本課程目前錄取名額已滿，您已是備取學員。")
                    sb.Append("</li>")

                    sb.Append("<li>")
                    sb.Append("為鼓勵您向上的精神，請珍惜學習的機會，屆時請準時來上課。")
                    sb.Append("</li>")

                    sb.Append("</ol>")
                Else
                    '正取
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
                    sC_URL = "http://mbscnn.org/MNT/MBMnt_RESP_01_v01.aspx?MB_MEMSEQ=" & iMB_MEMSEQ & "&MB_SEQ=" & ROW_CLASS("MB_SEQ").ToString & "&MB_BATCH=" & ROW_CLASS("MB_BATCH").ToString &
                             "&OPTYPE=Y"
                    'Dim sN_URL As String = String.Empty
                    'sN_URL = "http://mbscnn.org/MNT/MBMnt_RESP_01_v01.aspx?MB_MEMSEQ=" & iMB_MEMSEQ & "&MB_SEQ=" & MB_CLASS.getString("MB_SEQ") & "&MB_BATCH=" & MB_CLASS.getString("MB_BATCH") &
                    '         "&OPTYPE=N"
                    'sb.Append("<a style='color:#000040;font-size:20pt;font-weight:bold;' href='").Append(sC_URL).Append("' >確定出席</a>").Append("　　")
                    sb.Append(Me.getEMail_BT(sC_URL, "確定出席"))
                    'sb.Append("<a style='color:#000040;font-size:20pt;font-weight:bold;' href='").Append(sN_URL).Append("' >確定不出席</a>").Append("　　，")
                    'sb.Append(Me.getEMail_BT(sN_URL, "確定不出席")).Append("　　")

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

                    '1050119 AMY 錄取通知刪除 5,6,7,8項 新增 注意事項說明: MB_CLASS.FAVCHARGE
                    'sb.Append("<li>")
                    'sb.Append("<span style='color:black'>請穿著寬鬆衣褲；男女眾皆請勿穿著短褲。女眾請勿穿著貼身衣裙。</span>")
                    'sb.Append("</li>")

                    'sb.Append("<li>")
                    'sb.Append("請攜帶環保杯、筷子。")
                    'sb.Append("</li>")

                    'sb.Append("<li>")
                    'sb.Append("歡迎隨喜發心贊助場地或推廣教育課程費用。")
                    'sb.Append("</li>")

                    'sb.Append("<li>")
                    'sb.Append("可代訂素食便當（報到時登記即可，歡迎隨喜打齋）。")
                    'sb.Append("</li>")

                    Dim sMB_PREC_MEMO As String = String.Empty
                    sMB_PREC_MEMO = ROW_CLASS("MB_PREC_MEMO").ToString
                    If Utility.isValidateData(sMB_PREC_MEMO) Then
                        sMB_PREC_MEMO = ReplaceToBR(sMB_PREC_MEMO)
                        sb.Append("<li>")
                        sb.Append("<span style='color:red'>注意事項說明：</span>").Append("</span>")
                        sb.Append("<span style='color:red'><BR/>")
                        sb.Append(sMB_PREC_MEMO).Append("</span>")
                        sb.Append("</li>")
                    End If

                    sb.Append("</ol>")
                End If

                sb.Append("<div style='color:black;font-size:12pt;'>")
                sb.Append("　祝您").Append("<BR/>")
                sb.Append("　　　學習愉快、收穫滿滿").Append("<BR/>").Append("<BR/>")
                sb.Append("　　　　　　　　　 MBSC佛陀原始正法中心 敬邀合十")
                sb.Append("</div>")
            End If

            Return sb.ToString
        Catch ex As Exception
            Throw
        End Try
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

    Function getMB_CDAYS(ByVal MB_SDATE As Object, ByVal MB_CDAYS As Object) As String
        Try
            If Utility.isValidateData(MB_SDATE) AndAlso IsDate(MB_SDATE.ToString) AndAlso IsNumeric(MB_CDAYS) Then
                Dim sCDay As String = String.Empty
                sCDay = Utility.DateTransfer(DateAdd(DateInterval.Day, MB_CDAYS * -1, CDate(MB_SDATE.ToString)))

                Return "請於" & sCDay & "前取消報名"
            End If

            Return String.Empty
        Catch ex As Exception
            Throw
        End Try
    End Function

#End Region

    '取消報名
    Protected Sub btn_Cancel_Click(sender As Object, e As EventArgs) Handles btn_Cancel.Click
        Try
            Using m_DBManager = DatabaseManager.getInstance
                Dim MB_MEMCLASS As New MB_MEMCLASS(m_DBManager)

                If MB_MEMCLASS.LoadByPK(lbl_MB_MEMSEQ.Text, m_sCLASS) Then
                    If MB_MEMCLASS.getString("MB_FWMK") = "3" OrElse MB_MEMCLASS.getString("MB_FWMK") = "4" Then
                        com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "你已經取消報名")
                        com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "你已經取消報名")
                        Return
                    ElseIf MB_MEMCLASS.getString("MB_FWMK") = "1" OrElse MB_MEMCLASS.getString("MB_FWMK") = "0" OrElse Not Utility.isValidateData(MB_MEMCLASS.getAttribute("MB_FWMK")) Then
                        If Not Utility.isValidateData(MB_MEMCLASS.getAttribute("MB_CDATE")) Then
                            MB_MEMCLASS.setAttribute("MB_FWMK", "3")
                        Else
                            MB_MEMCLASS.setAttribute("MB_FWMK", "4")
                        End If

                        MB_MEMCLASS.setAttribute("MB_CDATE", Now)
                        MB_MEMCLASS.setAttribute("MB_CHGDATE", Now)

                        MB_MEMCLASS.save()

                        com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "取消報名成功")
                        com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "取消報名成功")
                    End If

                    Me.btn_Cancel.Visible = False
                End If
            End Using
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    'Function isVadCANCEL(ByVal dbManager As DatabaseManager) As String
    '    If Me.m_sOPTYPE = "CANCEL" AndAlso IsNumeric(Me.m_sCLASS) Then
    '        Try
    '            Dim mbCLASS As New MB_CLASS(dbManager)
    '            If mbCLASS.loadByPK(Me.m_sCLASS, Me.m_sMB_BATCH) Then
    '                Dim D_MB_SDATE As Object = Convert.DBNull
    '                'If Utility.isValidateData(mbCLASS.getAttribute("MB_SDATE")) AndAlso IsDate(mbCLASS.getAttribute("MB_SDATE").ToString) Then
    '                '    D_MB_SDATE = CDate(mbCLASS.getAttribute("MB_SDATE").ToString)
    '                'Else
    '                '    D_MB_SDATE = New Date(1911, 1, 1)
    '                'End If

    '                'If UCase(mbCLASS.getString("MB_APV")) = "1" Then
    '                '    If D_MB_SDATE <= DateAdd(DateInterval.Day, 1, Now) Then
    '                '        Return False
    '                '    Else
    '                '        Return True
    '                '    End If
    '                'Else
    '                '    If D_MB_SDATE <= DateAdd(DateInterval.Day, 3, Now) Then
    '                '        Return False
    '                '    Else
    '                '        Return True
    '                '    End If
    '                'End If

    '                Dim D_MB_CDAYS As Object = Convert.DBNull
    '                If Utility.isValidateData(mbCLASS.getAttribute("MB_SDATE")) AndAlso IsDate(mbCLASS.getAttribute("MB_CDAYS").ToString) Then
    '                    D_MB_SDATE = CDate(mbCLASS.getAttribute("MB_SDATE").ToString)

    '                    Dim iMB_CDAYS As Decimal = 0
    '                    iMB_CDAYS = mbCLASS.getDecimal("MB_CDAYS")
    '                    iMB_CDAYS = iMB_CDAYS * -1

    '                    If iMB_CDAYS < 0 AndAlso DateAdd(DateInterval.Day, iMB_CDAYS, D_MB_SDATE) <= Now Then
    '                        Return "本課程開課前" & mbCLASS.getAttribute("MB_CDAYS") & "天,不可取消，請聯絡中心"
    '                    End If
    '                End If
    '            End If
    '        Catch ex As Exception
    '            Throw
    '        End Try

    '        Return String.Empty
    '    End If

    '    Return "不可取消報名"
    'End Function

    Function isVadCANCEL(ByVal dbManager As DatabaseManager, ByVal iMB_MEMSEQ As Decimal) As String
        If Me.m_sOPTYPE = "CANCEL" AndAlso IsNumeric(Me.m_sCLASS) Then
            Try
                Dim sb As New System.Text.StringBuilder
                Dim MB_MEMBATCHList As New MB_MEMBATCHList(dbManager)
                MB_MEMBATCHList.setSQLCondition(" AND IFNULL(MB_ELECT,' ') = '1' ")
                MB_MEMBATCHList.LoadBySEQ(iMB_MEMSEQ, Me.m_sCLASS)
                Dim MB_CLASS As New MB_CLASS(dbManager)
                For Each ROW As DataRow In MB_MEMBATCHList.getCurrentDataSet.Tables(0).Rows
                    MB_CLASS.clear()
                    If MB_CLASS.loadByPK(Me.m_sCLASS, ROW("MB_BATCH")) Then
                        Dim D_MB_SDATE As Object = Convert.DBNull

                        Dim D_MB_CDAYS As Object = Convert.DBNull
                        If Utility.isValidateData(MB_CLASS.getAttribute("MB_SDATE")) AndAlso IsDate(MB_CLASS.getAttribute("MB_CDAYS").ToString) Then
                            D_MB_SDATE = CDate(MB_CLASS.getAttribute("MB_SDATE").ToString)

                            Dim iMB_CDAYS As Decimal = 0
                            iMB_CDAYS = MB_CLASS.getDecimal("MB_CDAYS")
                            iMB_CDAYS = iMB_CDAYS * -1

                            If iMB_CDAYS < 0 AndAlso DateAdd(DateInterval.Day, iMB_CDAYS, D_MB_SDATE) <= Now Then
                                Dim sBatch As String = String.Empty
                                If MB_CLASS.getInt("MB_BATCH") > 0 Then
                                    sBatch = "第" & MB_CLASS.getString("MB_BATCH") & "梯次"
                                End If

                                sb.Append("本課程").Append(sBatch).Append("開課前").Append(MB_CLASS.getAttribute("MB_CDAYS")).Append("天,不可取消，請聯絡中心").Append("\n")
                            End If
                        End If
                    End If
                Next

                If sb.Length > 0 Then
                    Return sb.ToString
                End If
            Catch ex As Exception
                Throw
            End Try

            Return String.Empty
        End If

        Return "不可取消報名"
    End Function

    Private Sub RP_BATCH_ItemDataBound(sender As Object, e As RepeaterItemEventArgs) Handles RP_BATCH.ItemDataBound
        If e.Item.ItemType = ListItemType.Item OrElse e.Item.ItemType = ListItemType.AlternatingItem Then
            Dim DRV As DataRowView = CType(e.Item.DataItem, DataRowView)

            If IsDate(DRV("MB_SDATE").ToString()) AndAlso IsDate(DRV("MB_EDATE").ToString) Then
                Dim D_MB_SDATE As Date = CDate(DRV("MB_SDATE").ToString())
                Dim D_MB_EDATE As Date = CDate(DRV("MB_EDATE").ToString())
                Dim MB_SEDATE As Literal = e.Item.FindControl("MB_SEDATE")
                MB_SEDATE.Text = D_MB_SDATE.Year & "/" & Utility.FillZero(D_MB_SDATE.Month, 2) & "/" & Utility.FillZero(D_MB_SDATE.Day, 2) &
                                 "至" &
                                 D_MB_EDATE.Year & "/" & Utility.FillZero(D_MB_EDATE.Month, 2) & "/" & Utility.FillZero(D_MB_EDATE.Day, 2)
            End If

            Dim sMB_MEMSEQ As String = String.Empty
            sMB_MEMSEQ = Me.lbl_MB_MEMSEQ.Text
            If IsNumeric(sMB_MEMSEQ) Then
                Using m_DBManager As DatabaseManager = DatabaseManager.getInstance
                    Dim MB_MEMBATCH As New MB_MEMBATCH(m_DBManager)
                    If MB_MEMBATCH.LoadByPK(sMB_MEMSEQ, Me.m_sCLASS, DRV("MB_BATCH")) Then
                        Dim CB_BATCH As CheckBox = e.Item.FindControl("CB_BATCH")

                        If MB_MEMBATCH.getString("MB_CHKFLAG") = "1" Then
                            CB_BATCH.Enabled = False
                        End If

                        If MB_MEMBATCH.getString("MB_ELECT") = "1" Then
                            CB_BATCH.Checked = True
                        End If

                        Dim MB_CHKFLAG As Literal = e.Item.FindControl("MB_CHKFLAG")
                        Select Case MB_MEMBATCH.getString("MB_CHKFLAG")
                            Case "1", "3"
                                MB_CHKFLAG.Text = "錄取"
                            Case "2"
                                MB_CHKFLAG.Text = "不錄取"
                        End Select
                    End If
                End Using
            End If
        End If
    End Sub
End Class