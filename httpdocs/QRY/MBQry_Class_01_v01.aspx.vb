Imports com.Azion.NET.VB
Imports MBSC.MB_OP
Imports MBSC.UICtl

Public Class MBQry_Class_01_v01
    Inherits System.Web.UI.Page

    Dim m_DBManager As DatabaseManager = Nothing

    Dim m_sPLACE As String = String.Empty

    Dim m_sUpcode76 As String = String.Empty

#Region "Page Events"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.m_DBManager = UIShareFun.getDataBaseManager

            Me.m_sPLACE = "" & Request.QueryString("PLACE")

            '維護人員
            Me.m_sUpcode76 = com.Azion.EloanUtility.CodeList.getAppSettings("UpCode76")

            If Not Page.IsPostBack Then
                Me.IMG_SET.ImageUrl = com.Azion.EloanUtility.UIUtility.getRootPath & "/ckfinder/userfiles/images/Settlement.png"

                Me.IMG_PRE.ImageUrl = com.Azion.EloanUtility.UIUtility.getRootPath & "/ckfinder/userfiles/images/Design.png"

                Me.IMG_RUN.ImageUrl = com.Azion.EloanUtility.UIUtility.getRootPath & "/ckfinder/userfiles/images/Run.png"

                Me.IMG_HIS.ImageUrl = com.Azion.EloanUtility.UIUtility.getRootPath & "/ckfinder/userfiles/images/His.png"

                '可受理報名課程
                Me.Bind_DDL_MB_PLACE(Me.DDL_MB_PLACE_1)
                Me.Bind_DDL_Date(Me.DDL_MB_SDATE_SY_1, "1")
                Me.Bind_DDL_Date(Me.DDL_MB_SDATE_SM_1, "2")
                Me.Bind_DDL_Date(Me.DDL_MB_SDATE_EY_1, "1")
                Me.Bind_DDL_Date(Me.DDL_MB_SDATE_EM_1, "2")
                Dim mbCLASSList As New MB_CLASSList(Me.m_DBManager)
                If Utility.isValidateData(Me.m_sPLACE) Then
                    mbCLASSList.LoadAPLY_MB_PLACE(Me.m_sPLACE)
                    Me.PLH_APLY_1_1.Visible = False
                    'Me.TD_1_1.Visible = False
                    'Me.TD_1_2.Visible = False
                    'Me.TD_1_3.Visible = False
                    'Me.TD_1_4.Visible = False
                    'Me.TD_1_5.Visible = False
                    'Me.TD_1_6.Visible = False
                    'Me.TD_1_7.Visible = False
                    Me.PLH_APLY_1_2.Visible = False
                Else
                    mbCLASSList.LoadAPLY()
                End If
                If Me.is76User(Session("UserId")) Then
                    Me.TD_1_7.Visible = True
                Else
                    Me.TD_1_7.Visible = False
                End If
                Me.RP_CLASS_1.DataSource = mbCLASSList.getCurrentDataSet.Tables(0)
                Me.RP_CLASS_1.DataBind()

                '預告課程
                Me.Bind_DDL_MB_PLACE(Me.DDL_MB_PLACE_2)
                Me.Bind_DDL_Date(Me.DDL_MB_SDATE_SY_2, "1")
                Me.Bind_DDL_Date(Me.DDL_MB_SDATE_SM_2, "2")
                Me.Bind_DDL_Date(Me.DDL_MB_SDATE_EY_2, "1")
                Me.Bind_DDL_Date(Me.DDL_MB_SDATE_EM_2, "2")
                mbCLASSList.clear()
                If Utility.isValidateData(Me.m_sPLACE) Then
                    mbCLASSList.LoadAPLY_MB_PLACE_2(Me.m_sPLACE)
                    Me.PLH_APLY_2_1.Visible = False
                    'Me.TD_2_1.Visible = False
                    'Me.TD_2_2.Visible = False
                    'Me.TD_2_3.Visible = False
                    'Me.TD_2_4.Visible = False
                    'Me.TD_2_5.Visible = False
                    Me.PLH_APLY_2_2.Visible = False
                Else
                    mbCLASSList.LoadAPLY_2()
                End If

                Me.RP_CLASS_2.DataSource = mbCLASSList.getCurrentDataSet.Tables(0)
                Me.RP_CLASS_2.DataBind()

                '進行中課程
                Me.Bind_DDL_MB_PLACE(Me.DDL_MB_PLACE_4)
                Me.Bind_DDL_Date(Me.DDL_MB_SDATE_SY_4, "1")
                Me.Bind_DDL_Date(Me.DDL_MB_SDATE_SM_4, "2")
                Me.Bind_DDL_Date(Me.DDL_MB_SDATE_EY_4, "1")
                Me.Bind_DDL_Date(Me.DDL_MB_SDATE_EM_4, "2")
                mbCLASSList.clear()
                If Utility.isValidateData(Me.m_sPLACE) Then
                    mbCLASSList.LoadAPLY_MB_PLACE_4(Me.m_sPLACE)
                    Me.PLH_APLY_3_1.Visible = False
                    'Me.TD_3_1.Visible = False
                    'Me.TD_3_2.Visible = False
                    'Me.TD_3_3.Visible = False
                    'Me.TD_3_4.Visible = False
                    'Me.TD_3_5.Visible = False
                    Me.PLH_APLY_3_2.Visible = False
                Else
                    mbCLASSList.LoadAPLY_4()
                End If
                Me.RP_CLASS_4.DataSource = mbCLASSList.getCurrentDataSet.Tables(0)
                Me.RP_CLASS_4.DataBind()

                '截止報名尚未開課
                If Me.is76User(Session("UserId")) Then
                    Me.TD_5_6.Visible = True
                Else
                    Me.TD_5_6.Visible = False
                End If
                Me.Bind_DDL_MB_PLACE(Me.DDL_MB_PLACE_5)
                Me.Bind_DDL_Date(Me.DDL_MB_SDATE_SY_5, "1")
                Me.Bind_DDL_Date(Me.DDL_MB_SDATE_SM_5, "2")
                Me.Bind_DDL_Date(Me.DDL_MB_SDATE_EY_5, "1")
                Me.Bind_DDL_Date(Me.DDL_MB_SDATE_EM_5, "2")
                mbCLASSList.clear()
                If Utility.isValidateData(Me.m_sPLACE) Then
                    mbCLASSList.LoadAPLY_MB_PLACE_5(Me.m_sPLACE)
                    Me.PLH_APLY_5_1.Visible = False
                    'Me.TD_3_1.Visible = False
                    'Me.TD_3_2.Visible = False
                    'Me.TD_3_3.Visible = False
                    'Me.TD_3_4.Visible = False
                    'Me.TD_3_5.Visible = False
                    Me.PLH_APLY_5_2.Visible = False
                Else
                    mbCLASSList.LoadAPLY_5()
                End If
                Me.RP_CLASS_5.DataSource = mbCLASSList.getCurrentDataSet.Tables(0)
                Me.RP_CLASS_5.DataBind()

                '關閉課程(已完成之課程)
                Me.Bind_DDL_MB_PLACE(Me.DDL_MB_PLACE_3)
                Me.Bind_DDL_Date(Me.DDL_MB_SDATE_SY_3, "1")
                Me.Bind_DDL_Date(Me.DDL_MB_SDATE_SM_3, "2")
                Me.Bind_DDL_Date(Me.DDL_MB_SDATE_EY_3, "1")
                Me.Bind_DDL_Date(Me.DDL_MB_SDATE_EM_3, "2")
                mbCLASSList.clear()
                If Utility.isValidateData(Me.m_sPLACE) Then
                    mbCLASSList.LoadAPLY_MB_PLACE_3(Me.m_sPLACE)
                Else
                    mbCLASSList.LoadAPLY_3()
                End If
                Me.RP_CLASS_3.DataSource = mbCLASSList.getCurrentDataSet.Tables(0)
                Me.RP_CLASS_3.DataBind()

                If Utility.isValidateData(Me.m_sPLACE) Then
                    Me.PLH_4.Visible = False
                End If
            End If
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

    Private Sub MBQry_Class_01_v01_Unload(sender As Object, e As EventArgs) Handles Me.Unload
        UIShareFun.releaseConnection(Me.m_DBManager)
    End Sub

    Sub Bind_DDL_MB_PLACE(ByVal DDL As DropDownList)
        Try
            Dim AP_CODEList As New AP_CODEList(Me.m_DBManager)
            AP_CODEList.loadByUpCode(177)
            DDL.DataSource = AP_CODEList.getCurrentDataSet.Tables(0)
            DDL.DataTextField = "TEXT"
            DDL.DataValueField = "VALUE"
            DDL.DataBind()
            DDL.Items.Insert(0, New ListItem("全部", ""))
        Catch ex As Exception
            Throw
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
#End Region

#Region "可受理報名課程"
    Private Sub RP_CLASS_1_ItemCommand(source As Object, e As RepeaterCommandEventArgs) Handles RP_CLASS_1.ItemCommand, RP_CLASS_2.ItemCommand, RP_CLASS_3.ItemCommand, RP_CLASS_4.ItemCommand, RP_CLASS_5.ItemCommand
        Try
            Dim LTL_MB_SEQ As Literal = e.Item.FindControl("LTL_MB_SEQ")
            Dim MB_BATCH As HtmlInputHidden = e.Item.FindControl("MB_BATCH")
            If UCase(e.CommandName) = "CONTENT" Then
                Dim mbNEWSList As New MB_NEWSList(Me.m_DBManager)
                mbNEWSList.setSQLCondition(" ORDER BY CRETIME DESC ")
                mbNEWSList.LoadbyMB_SEQ(CDec(LTL_MB_SEQ.Text))
                If mbNEWSList.getCurrentDataSet.Tables(0).Rows.Count > 0 Then
                    Dim iCRETIME As Decimal = 0
                    iCRETIME = mbNEWSList.getCurrentDataSet.Tables(0).Rows(0)("CRETIME")

                    Dim sURL As String = String.Empty
                    sURL = com.Azion.EloanUtility.UIUtility.getRootPath & "/NewsList.aspx?CRETIME=" & iCRETIME
                    com.Azion.EloanUtility.UIUtility.showOpen(sURL, "")
                Else
                    com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "查無課程內容")
                    com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "查無課程內容")
                    Return
                End If
            ElseIf UCase(e.CommandName) = "SIGNUP" Then
                If Not com.Azion.EloanUtility.ValidateUtility.isValidateData(Session("USERID")) Then
                    com.Azion.EloanUtility.UIUtility.alert("請先登入或成為會員")
                    com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請先登入或成為會員")
                    Return
                End If

                Dim sURL As String = String.Empty
                sURL = com.Azion.EloanUtility.UIUtility.getRootPath & "/MNT/MBMnt_Sign_01_v01.aspx?MBSEQ=" & LTL_MB_SEQ.Text & "&MB_BATCH=" & MB_BATCH.Value

                'com.Azion.EloanUtility.UIUtility.showModalDialog(sURL)
                Response.Redirect(sURL)
            ElseIf UCase(e.CommandName) = "CANCEL" Then
                If Not com.Azion.EloanUtility.ValidateUtility.isValidateData(Session("USERID")) Then
                    com.Azion.EloanUtility.UIUtility.alert("請先登入或成為會員")
                    com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請先登入或成為會員")
                    Return
                End If

                Dim sURL As String = String.Empty
                sURL = com.Azion.EloanUtility.UIUtility.getRootPath & "/MNT/MBMnt_Sign_01_v01.aspx?MBSEQ=" & LTL_MB_SEQ.Text & "&OPTYPE=CANCEL" & "&MB_BATCH=" & MB_BATCH.Value

                Response.Redirect(sURL)
            ElseIf UCase(e.CommandName) = "MB_ALERT1" Then
                '提醒信一
                Me.StartSendMail(LTL_MB_SEQ.Text, MB_BATCH.Value, 1)

                Me.Bind_RP_CLASS_1()
            ElseIf UCase(e.CommandName) = "MB_ALERT2" Then
                '提醒信二
                Me.StartSendMail(LTL_MB_SEQ.Text, MB_BATCH.Value, 2)

                Me.Bind_RP_CLASS_1()
            End If
        Catch ex As System.Threading.ThreadAbortException
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

    Sub Bind_RP_CLASS_1()
        Dim mbCLASSList As New MB_CLASSList(Me.m_DBManager)
        If Utility.isValidateData(Me.m_sPLACE) Then
            mbCLASSList.LoadAPLY_MB_PLACE(Me.m_sPLACE)
            Me.PLH_APLY_1_1.Visible = False
            'Me.TD_1_1.Visible = False
            'Me.TD_1_2.Visible = False
            'Me.TD_1_3.Visible = False
            'Me.TD_1_4.Visible = False
            'Me.TD_1_5.Visible = False
            'Me.TD_1_6.Visible = False
            'Me.TD_1_7.Visible = False
            Me.PLH_APLY_1_2.Visible = False
        Else
            mbCLASSList.LoadAPLY()
        End If
        If Me.is76User(Session("UserId")) Then
            Me.TD_1_7.Visible = True
        Else
            Me.TD_1_7.Visible = False
        End If
        Me.RP_CLASS_1.DataSource = mbCLASSList.getCurrentDataSet.Tables(0)
        Me.RP_CLASS_1.DataBind()
    End Sub

    Sub StartSendMail(ByVal iMB_SEQ As Decimal, ByVal iMB_BATCH As Decimal, ByVal iMode As Integer)
        Try
            Dim MB_CLASS As New MB_CLASS(Me.m_DBManager)
            If MB_CLASS.loadByPK(iMB_SEQ, iMB_BATCH) Then
                Me.SendMail(MB_CLASS, iMode)

                If iMode = 1 Then
                    MB_CLASS.setAttribute("MB_ALERT1", "Y")
                ElseIf iMode = 2 Then
                    MB_CLASS.setAttribute("MB_ALERT2", "Y")
                End If
                MB_CLASS.save()
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Sub SendMail(ByVal MB_CLASS As MB_CLASS, ByVal iMode As Integer)
        Dim DT_MAIL As DataTable = Nothing
        Try
            Dim MB_MEMCLASSList As New MB_MEMCLASSList(Me.m_DBManager)
            MB_MEMCLASSList.getMail_TO(MB_CLASS.getDecimal("MB_SEQ"), MB_CLASS.getDecimal("MB_BATCH"))
            DT_MAIL = MB_MEMCLASSList.getCurrentDataSet.Tables(0)
            If Not IsNothing(DT_MAIL) AndAlso DT_MAIL.Rows.Count > 0 Then
                Dim sMailSub As String = String.Empty

                For Each ROW As DataRow In DT_MAIL.Rows
                    If Utility.isValidateData(Trim(ROW("MB_EMAIL").ToString)) Then
                        Dim sMailBody As String = String.Empty
                        If iMode = 1 Then
                            sMailBody = Me.getMailBody_1(MB_CLASS, ROW("MB_MEMSEQ"), ROW("MB_NAME").ToString)
                        ElseIf iMode = 2 Then
                            sMailBody = Me.getMailBody_2(MB_CLASS, ROW("MB_MEMSEQ"), ROW("MB_NAME").ToString)
                        End If

                        If iMode = 1 Then
                            '台北中階精進二日禪  提醒通知函
                            sMailSub = MB_CLASS.getString("MB_CLASS_NAME") & " 提醒通知函"
                        ElseIf iMode = 2 Then
                            '台北中階精進二日禪  提醒通知函
                            sMailSub = MB_CLASS.getString("MB_CLASS_NAME") & " 提醒通知函"
                        End If
                        Try
                            com.Azion.EloanUtility.NetUtility.GMail_Send({Trim(ROW("MB_EMAIL").ToString)}, Nothing, sMailSub, sMailBody, True, Nothing, False)
                        Catch ex As Exception

                        End Try
                    End If
                Next

                Dim sMsg As String = String.Empty
                If iMode = 1 Then
                    sMsg = sMailSub & "(一)，已發送完畢"
                Else
                    sMsg = sMailSub & "(二)，已發送完畢"
                End If
                com.Azion.EloanUtility.UIUtility.showErrMsg(Me, sMailSub)
            End If
        Catch ex As Exception
            Throw
        Finally
            If Not IsNothing(DT_MAIL) Then
                DT_MAIL.Dispose()
            End If
        End Try
    End Sub

    Function getMailBody_1(ByVal MB_CLASS As MB_CLASS, ByVal iMB_MEMSEQ As Decimal, ByVal sMB_NAME As String) As String
        Try
            Dim sb As New System.Text.StringBuilder

            sb.Append("<DIV style='text-align:center;font-size:24pt;color:red' >").Append(MB_CLASS.getString("MB_CLASS_NAME")).Append("</DIV>")
            sb.Append("<DIV style='text-align:center;font-size:24pt;color:red;font-weight:bold;text-decoration:underline' >★ 提醒您 ★").Append("</DIV>")
            sb.Append("<DIV style='font-size:12pt;color:#7030A0'>").Append("仁者吉祥：").Append("</DIV>")

            sb.Append("<DIV style='font-size:12pt;color:#7030A0;font-weight:bold'>").Append("　　  您已報名參加 ")
            sb.Append("<span style='color:red'>")
            sb.Append(MB_CLASS.getString("MB_CLASS_NAME"))
            sb.Append("</span>")
            sb.Append("！此通知函，提醒您，別在忙碌中流逝，忘了學習或放棄初衷，期待與您相見。").Append("</DIV>")

            sb.Append("<ol type='1' style='font-size:12pt;color:#7030A0;font-weight:bold' >")
            Dim sMB_SDATE As String = String.Empty
            If Utility.isValidateData(MB_CLASS.getAttribute("MB_SDATE")) Then
                sMB_SDATE = CDate(MB_CLASS.getAttribute("MB_SDATE").ToString).Year & "年" & CDate(MB_CLASS.getAttribute("MB_SDATE").ToString).Month & "月" & _
                            CDate(MB_CLASS.getAttribute("MB_SDATE").ToString).Day & "日"
            End If
            Dim sREGTIME As String = String.Empty
            sREGTIME = Left(Utility.FillZero(MB_CLASS.getString("REGTIME"), 4), 2) & Right(Utility.FillZero(MB_CLASS.getString("REGTIME"), 4), 2)
            sb.Append("<li>").Append("本課程開始日期時間：").Append("<span style='color:red'>").Append(sMB_SDATE).Append("</span>")
            sb.Append("　報到時間：").Append("<span style='color:red'>").Append(sREGTIME).Append("</span>")
            sb.Append("</li>")
            sb.Append("<li>")
            sb.Append(" 課程地點：").Append("<span style='color:#F335CF'>MBSC").Append(MB_CLASS.getString("MB_PLACE")).Append(" / ")
            sb.Append(MB_CLASS.getString("CLASS_PLACE")).Append("<BR/>").Append(MB_CLASS.getString("TRAFFIC_DESC"))
            sb.Append("</li>")
            'sb.Append("<li>")
            'sb.Append("當您收到確認後，").Append("<span style='color:red;font-weight:bold;font-size:24pt'>").Append("請按確定出席/不出席，").Append("</span>")
            'Dim sC_URL As String = String.Empty
            'sC_URL = "http://mbscnn.org/MNT/MBMnt_RESP_01_v01.aspx?MB_MEMSEQ=" & iMB_MEMSEQ & "&MB_SEQ=" & MB_CLASS.getString("MB_SEQ") & "&MB_BATCH=" & MB_CLASS.getString("MB_BATCH") & _
            '         "&OPTYPE=Y"
            'Dim sN_URL As String = String.Empty
            'sN_URL = "http://mbscnn.org/MNT/MBMnt_RESP_01_v01.aspx?MB_MEMSEQ=" & iMB_MEMSEQ & "&MB_SEQ=" & MB_CLASS.getString("MB_SEQ") & "&MB_BATCH=" & MB_CLASS.getString("MB_BATCH") & _
            '         "&OPTYPE=N"
            'sb.Append("<a style='color:#000040;font-size:20pt;font-weight:bold;' href='").Append(sC_URL).Append("' >確定出席</a>").Append("　　")
            'sb.Append("<a style='color:#000040;font-size:20pt;font-weight:bold;' href='").Append(sN_URL).Append("' >確定不出席</a>").Append("　　，")
            'sb.Append("以利增補候補學員，感謝您。")
            'sb.Append("</li>")

            sb.Append("<li>")
            sb.Append("<span style='color:red'>聯絡電話：</span>").Append(MB_CLASS.getString("CONTEL")).Append("　　")
            sb.Append("<span style='color:red'>聯絡人：").Append(MB_CLASS.getString("CONTACT")).Append("</span>")
            sb.Append("</li>")

            sb.Append("<li>")
            sb.Append("<span style='color:red'>請穿著寬鬆衣褲；男女眾皆請勿穿著短褲。女眾請勿穿著貼身衣裙。</span>")
            sb.Append("</li>")

            sb.Append("<li>")
            sb.Append("請攜帶環保杯、筷子。")
            sb.Append("</li>")

            sb.Append("<li>")
            sb.Append("歡迎隨喜發心贊助場地或推廣教育課程費用。")
            sb.Append("</li>")

            sb.Append("<li>")
            sb.Append(" 可代訂素食便當（報到時登記即可，歡迎隨喜打齋）。")
            sb.Append("</li>")

            sb.Append("<li>")
            sb.Append(" 尚未回覆者，請盡快回覆。")
            sb.Append("</li>")

            sb.Append("</ol>")

            sb.Append("<div style='color:#C80896;font-size:14pt;font-weight:bold'>")
            sb.Append("　祝您").Append("<BR/>")
            sb.Append("　　　禪修愉快、收穫滿滿").Append("<BR/>").Append("<BR/>")
            sb.Append("　　　　　　　　　 MBSC 台北教育中心 敬邀合十")
            sb.Append("</div>")

            Return sb.ToString
        Catch ex As Exception
            Throw
        End Try
    End Function

    Function getMailBody_2(ByVal MB_CLASS As MB_CLASS, ByVal iMB_MEMSEQ As Decimal, ByVal sMB_NAME As String) As String
        Try
            Dim sb As New System.Text.StringBuilder

            sb.Append("<DIV style='text-align:center;font-size:24pt;color:red' >").Append(MB_CLASS.getString("MB_CLASS_NAME")).Append("</DIV>")
            sb.Append("<DIV style='text-align:center;font-size:24pt;color:red;font-weight:bold;text-decoration:underline' >★ 提醒您 ★").Append("</DIV>")
            sb.Append("<DIV style='font-size:12pt;color:#7030A0'>").Append("仁者吉祥：").Append("</DIV>")

            sb.Append("<DIV style='font-size:12pt;color:#7030A0;font-weight:bold'>").Append("　　  您已報名參加 ")
            sb.Append("<span style='color:red'>")
            sb.Append(MB_CLASS.getString("MB_CLASS_NAME"))
            sb.Append("</span>")
            sb.Append("！此通知函，提醒您，別在忙碌中流逝，忘了學習或放棄初衷，期待與您相見。").Append("</DIV>")

            sb.Append("<ol type='1' style='font-size:12pt;color:#7030A0;font-weight:bold' >")
            Dim sMB_SDATE As String = String.Empty
            If Utility.isValidateData(MB_CLASS.getAttribute("MB_SDATE")) Then
                sMB_SDATE = CDate(MB_CLASS.getAttribute("MB_SDATE").ToString).Year & "年" & CDate(MB_CLASS.getAttribute("MB_SDATE").ToString).Month & "月" & _
                            CDate(MB_CLASS.getAttribute("MB_SDATE").ToString).Day & "日"
            End If
            Dim sREGTIME As String = String.Empty
            sREGTIME = Left(Utility.FillZero(MB_CLASS.getString("REGTIME"), 4), 2) & Right(Utility.FillZero(MB_CLASS.getString("REGTIME"), 4), 2)
            sb.Append("<li>").Append("本課程開始日期時間：").Append("<span style='color:red'>").Append(sMB_SDATE).Append("</span>")
            sb.Append("　報到時間：").Append("<span style='color:red'>").Append(sREGTIME).Append("</span>")
            sb.Append("</li>")
            sb.Append("<li>")
            sb.Append(" 課程地點：").Append("<span style='color:#F335CF'>MBSC").Append(MB_CLASS.getString("MB_PLACE")).Append(" / ")
            sb.Append(MB_CLASS.getString("CLASS_PLACE")).Append("<BR/>").Append(MB_CLASS.getString("TRAFFIC_DESC"))
            sb.Append("</li>")
            'sb.Append("<li>")
            'sb.Append("當您收到確認後，").Append("<span style='color:red;font-weight:bold;font-size:24pt'>").Append("請按確定出席/不出席，").Append("</span>")
            'Dim sC_URL As String = String.Empty
            'sC_URL = "http://mbscnn.org/MNT/MBMnt_RESP_01_v01.aspx?MB_MEMSEQ=" & iMB_MEMSEQ & "&MB_SEQ=" & MB_CLASS.getString("MB_SEQ") & "&MB_BATCH=" & MB_CLASS.getString("MB_BATCH") & _
            '         "&OPTYPE=Y"
            'Dim sN_URL As String = String.Empty
            'sN_URL = "http://mbscnn.org/MNT/MBMnt_RESP_01_v01.aspx?MB_MEMSEQ=" & iMB_MEMSEQ & "&MB_SEQ=" & MB_CLASS.getString("MB_SEQ") & "&MB_BATCH=" & MB_CLASS.getString("MB_BATCH") & _
            '         "&OPTYPE=N"
            'sb.Append("<a style='color:#000040;font-size:20pt;font-weight:bold;' href='").Append(sC_URL).Append("' >確定出席</a>").Append("　　")
            'sb.Append("<a style='color:#000040;font-size:20pt;font-weight:bold;' href='").Append(sN_URL).Append("' >確定不出席</a>").Append("　　，")
            'sb.Append("以利增補候補學員，感謝您。")
            'sb.Append("</li>")

            sb.Append("<li>")
            sb.Append("<span style='color:red'>聯絡電話：</span>").Append(MB_CLASS.getString("CONTEL")).Append("　　")
            sb.Append("<span style='color:red'>聯絡人：").Append(MB_CLASS.getString("CONTACT")).Append("</span>")
            sb.Append("</li>")

            sb.Append("<li>")
            sb.Append("<span style='color:red'>請穿著寬鬆衣褲；男女眾皆請勿穿著短褲。女眾請勿穿著貼身衣裙。</span>")
            sb.Append("</li>")

            sb.Append("<li>")
            sb.Append("請攜帶環保杯、筷子。")
            sb.Append("</li>")

            sb.Append("<li>")
            sb.Append("歡迎隨喜發心贊助場地或推廣教育課程費用。")
            sb.Append("</li>")

            sb.Append("<li>")
            sb.Append(" 可代訂素食便當（報到時登記即可，歡迎隨喜打齋）。")
            sb.Append("</li>")

            sb.Append("</ol>")

            sb.Append("<div style='color:#C80896;font-size:14pt;font-weight:bold'>")
            sb.Append("　祝您").Append("<BR/>")
            sb.Append("　　　禪修愉快、收穫滿滿").Append("<BR/>").Append("<BR/>")
            sb.Append("　　　　　　　　　 MBSC 台北教育中心 敬邀合十")
            sb.Append("</div>")

            Return sb.ToString
        Catch ex As Exception
            Throw
        End Try
    End Function

    Private Sub RP_CLASS_1_ItemDataBound(sender As Object, e As Web.UI.WebControls.RepeaterItemEventArgs) Handles RP_CLASS_1.ItemDataBound
        Try
            If e.Item.ItemType = ListItemType.Item OrElse e.Item.ItemType = ListItemType.AlternatingItem Then
                Dim DRV As DataRowView = CType(e.Item.DataItem, DataRowView)

                Dim sMB_SDATE As String = String.Empty
                If IsDate(DRV("MB_SDATE").ToString) Then
                    'sMB_SDATE = Utility.DateTransfer(CDate(DRV("MB_SDATE").ToString))
                    sMB_SDATE = MBSC_DATE(DRV("MB_SDATE"))
                End If
                Dim sMB_EWEEK As String = String.Empty
                If IsDate(DRV("MB_EDATE").ToString) Then
                    'sMB_EWEEK = Utility.DateTransfer(CDate(DRV("MB_EDATE").ToString))
                    sMB_EWEEK = MBSC_DATE(DRV("MB_EDATE"))
                End If
                Dim LTL_MB_SEDATE As Literal = e.Item.FindControl("LTL_MB_SEDATE")
                LTL_MB_SEDATE.Text = sMB_SDATE & "~" & sMB_EWEEK

                Dim sMB_SAPLY As String = String.Empty
                If IsDate(DRV("MB_SAPLY").ToString) Then
                    'sMB_SAPLY = Utility.DateTransfer(CDate(DRV("MB_SAPLY").ToString))
                    sMB_SAPLY = MBSC_DATE(DRV("MB_SAPLY"))
                End If
                Dim sMB_EAPLY As String = String.Empty
                If IsDate(DRV("MB_EAPLY").ToString) Then
                    'sMB_EAPLY = Utility.DateTransfer(CDate(DRV("MB_EAPLY").ToString))
                    sMB_EAPLY = MBSC_DATE(DRV("MB_EAPLY"))
                End If
                Dim LTL_MB_SEAPLY As Literal = e.Item.FindControl("LTL_MB_SEAPLY")
                LTL_MB_SEAPLY.Text = sMB_SAPLY & "~" & sMB_EAPLY

                If Me.is76User(Session("UserId")) Then
                    Dim btnMB_ALERT1 As Button = e.Item.FindControl("btnMB_ALERT1")
                    Dim btnMB_ALERT2 As Button = e.Item.FindControl("btnMB_ALERT2")

                    If CDate(DRV("MB_SDATE").ToString).AddDays(-1 * DRV("MB_ALERT1_DAY")) <= Now Then
                        If Not Utility.isValidateData(DRV("MB_ALERT1")) Then
                            btnMB_ALERT1.Visible = True
                        Else
                            btnMB_ALERT1.Enabled = False
                            btnMB_ALERT1.Text = "已經發第一封提醒信"

                            'If CDate(DRV("MB_SDATE").ToString).AddDays(-1 * DRV("MB_ALERT2_DAY")) <= Now Then
                            '    btnMB_ALERT2.Visible = True
                            'Else
                            '    btnMB_ALERT2.Visible = False
                            'End If

                            If CDate(DRV("MB_SDATE").ToString) > Now Then
                                btnMB_ALERT2.Visible = True
                            Else
                                btnMB_ALERT2.Visible = False
                            End If
                        End If
                    Else
                        btnMB_ALERT1.Visible = False
                        btnMB_ALERT2.Visible = False
                    End If
                Else
                    Dim TD_1_7 As HtmlTableCell = e.Item.FindControl("TD_1_7")
                    TD_1_7.Visible = False
                End If

                If Utility.isValidateData(Me.m_sPLACE) Then
                    Dim TD_1_1 As HtmlTableCell = e.Item.FindControl("TD_1_1")
                    TD_1_1.Visible = False

                    Dim TD_1_2 As HtmlTableCell = e.Item.FindControl("TD_1_2")
                    TD_1_2.Visible = False

                    Dim TD_1_3 As HtmlTableCell = e.Item.FindControl("TD_1_3")
                    TD_1_3.Visible = False

                    Dim TD_1_4 As HtmlTableCell = e.Item.FindControl("TD_1_4")
                    TD_1_4.Visible = False

                    Dim TD_1_5 As HtmlTableCell = e.Item.FindControl("TD_1_5")
                    TD_1_5.Visible = False

                    Dim TD_1_6 As HtmlTableCell = e.Item.FindControl("TD_1_6")
                    TD_1_6.Visible = False

                    Dim TD_1_7 As HtmlTableCell = e.Item.FindControl("TD_1_7")
                    TD_1_7.Visible = False
                End If

                If DRV("MB_YES").ToString = "N" Then
                    Dim btnSIGNUP As Button = e.Item.FindControl("btnSIGNUP")
                    btnSIGNUP.Visible = False
                    Dim btnCANCEL As Button = e.Item.FindControl("btnCANCEL")
                    btnCANCEL.Visible = False
                    Dim LTL_APLY As Label = e.Item.FindControl("LTL_APLY")
                    LTL_APLY.Visible = True

                    Dim btnMB_ALERT1 As Button = e.Item.FindControl("btnMB_ALERT1")
                    btnMB_ALERT1.Visible = False
                    Dim btnMB_ALERT2 As Button = e.Item.FindControl("btnMB_ALERT2")
                    btnMB_ALERT2.Visible = False
                End If

                If DRV("MB_ALERT2").ToString = "Y" Then
                    Dim LTL_MB_ALERT2 As Label = e.Item.FindControl("LTL_MB_ALERT2")
                    LTL_MB_ALERT2.Text = "<BR/>已發過"
                End If
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Private Sub btnQRY1_Click(sender As Object, e As EventArgs) Handles btnQRY1.Click
        Try
            Dim sMB_PLACE As String = String.Empty
            If Utility.isValidateData(Me.DDL_MB_PLACE_1.SelectedValue) Then
                sMB_PLACE = Me.DDL_MB_PLACE_1.SelectedItem.Text
            End If

            Dim sMB_SAPLY_S As String = String.Empty
            If IsNumeric(Me.DDL_MB_SDATE_SY_1.SelectedValue) AndAlso IsNumeric(Me.DDL_MB_SDATE_SM_1.SelectedValue) Then
                sMB_SAPLY_S = CDec(Me.DDL_MB_SDATE_SY_1.SelectedValue) & "/" & CDec(Me.DDL_MB_SDATE_SM_1.SelectedValue) & "/1"
            End If

            Dim sMB_SAPLY_E As String = String.Empty
            If IsNumeric(Me.DDL_MB_SDATE_EY_1.SelectedValue) AndAlso IsNumeric(Me.DDL_MB_SDATE_EM_1.SelectedValue) Then
                Dim sTMP As String = String.Empty
                sTMP = CDec(Me.DDL_MB_SDATE_EY_1.SelectedValue) & "/" & CDec(Me.DDL_MB_SDATE_EM_1.SelectedValue) & "/1"
                sMB_SAPLY_E = DateAdd(DateInterval.Day, -1, (DateAdd(DateInterval.Month, 1, CDate(sTMP)))).ToString
            End If

            Dim isVadDate As Boolean = False
            If IsDate(sMB_SAPLY_S) AndAlso IsDate(sMB_SAPLY_E) Then
                isVadDate = True
            End If

            Dim mbCLASSList As New MB_CLASSList(Me.m_DBManager)
            If Not Utility.isValidateData(sMB_PLACE) AndAlso isVadDate = False Then
                mbCLASSList.LoadAPLY()
            ElseIf Utility.isValidateData(sMB_PLACE) AndAlso isVadDate = False Then
                mbCLASSList.LoadAPLY_MB_PLACE(sMB_PLACE)
            ElseIf Utility.isValidateData(sMB_PLACE) AndAlso isVadDate = True Then
                mbCLASSList.LoadAPLY_PLACE_SDATE(sMB_PLACE, CDate(sMB_SAPLY_S), CDate(sMB_SAPLY_E))
            ElseIf Not Utility.isValidateData(sMB_PLACE) AndAlso isVadDate = True Then
                mbCLASSList.LoadAPLY_SDATE(CDate(sMB_SAPLY_S), CDate(sMB_SAPLY_E))
            End If

            Me.RP_CLASS_1.DataSource = mbCLASSList.getCurrentDataSet.Tables(0)
            Me.RP_CLASS_1.DataBind()
        Catch ex As Exception
            Throw
        End Try
    End Sub

#End Region

#Region "預告課程"
    Private Sub RP_CLASS_2_ItemDataBound(sender As Object, e As Web.UI.WebControls.RepeaterItemEventArgs) Handles RP_CLASS_2.ItemDataBound
        Try
            If e.Item.ItemType = ListItemType.Item OrElse e.Item.ItemType = ListItemType.AlternatingItem Then
                Dim DRV As DataRowView = CType(e.Item.DataItem, DataRowView)

                Dim sMB_SDATE As String = String.Empty
                If IsDate(DRV("MB_SDATE").ToString) Then
                    'sMB_SDATE = Utility.DateTransfer(CDate(DRV("MB_SDATE").ToString))
                    sMB_SDATE = MBSC_DATE(DRV("MB_SDATE"))
                End If
                Dim sMB_EWEEK As String = String.Empty
                If IsDate(DRV("MB_EDATE").ToString) Then
                    'sMB_EWEEK = Utility.DateTransfer(CDate(DRV("MB_EDATE").ToString))
                    sMB_EWEEK = MBSC_DATE(DRV("MB_EDATE"))
                End If
                Dim LTL_MB_SEDATE As Literal = e.Item.FindControl("LTL_MB_SEDATE")
                LTL_MB_SEDATE.Text = sMB_SDATE & "~" & sMB_EWEEK

                Dim sMB_SAPLY As String = String.Empty
                If IsDate(DRV("MB_SAPLY").ToString) Then
                    'sMB_SAPLY = Utility.DateTransfer(CDate(DRV("MB_SAPLY").ToString))
                    sMB_SAPLY = MBSC_DATE(DRV("MB_SAPLY"))
                End If
                Dim sMB_EAPLY As String = String.Empty
                If IsDate(DRV("MB_EAPLY").ToString) Then
                    'sMB_EAPLY = Utility.DateTransfer(CDate(DRV("MB_EAPLY").ToString))
                    sMB_EAPLY = MBSC_DATE(DRV("MB_EAPLY"))
                End If
                Dim LTL_MB_SEAPLY As Literal = e.Item.FindControl("LTL_MB_SEAPLY")
                LTL_MB_SEAPLY.Text = sMB_SAPLY & "~" & sMB_EAPLY

                If Utility.isValidateData(Me.m_sPLACE) Then
                    Dim TD_2_1 As HtmlTableCell = e.Item.FindControl("TD_2_1")
                    TD_2_1.Visible = False

                    Dim TD_2_2 As HtmlTableCell = e.Item.FindControl("TD_2_2")
                    TD_2_2.Visible = False

                    Dim TD_2_3 As HtmlTableCell = e.Item.FindControl("TD_2_3")
                    TD_2_3.Visible = False

                    Dim TD_2_4 As HtmlTableCell = e.Item.FindControl("TD_2_4")
                    TD_2_4.Visible = False

                    Dim TD_2_5 As HtmlTableCell = e.Item.FindControl("TD_2_5")
                    TD_2_5.Visible = False
                End If
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Private Sub btnQRY2_Click(sender As Object, e As EventArgs) Handles btnQRY2.Click
        Try
            Dim sMB_PLACE As String = String.Empty
            If Utility.isValidateData(Me.DDL_MB_PLACE_1.SelectedValue) Then
                sMB_PLACE = Me.DDL_MB_PLACE_1.SelectedItem.Text
            End If

            Dim sMB_SAPLY_S As String = String.Empty
            If IsNumeric(Me.DDL_MB_SDATE_SY_1.SelectedValue) AndAlso IsNumeric(Me.DDL_MB_SDATE_SM_1.SelectedValue) Then
                sMB_SAPLY_S = CDec(Me.DDL_MB_SDATE_SY_1.SelectedValue) & "/" & CDec(Me.DDL_MB_SDATE_SM_1.SelectedValue) & "/1"
            End If

            Dim sMB_SAPLY_E As String = String.Empty
            If IsNumeric(Me.DDL_MB_SDATE_EY_1.SelectedValue) AndAlso IsNumeric(Me.DDL_MB_SDATE_EM_1.SelectedValue) Then
                Dim sTMP As String = String.Empty
                sTMP = CDec(Me.DDL_MB_SDATE_EY_1.SelectedValue) & "/" & CDec(Me.DDL_MB_SDATE_EM_1.SelectedValue) & "/1"
                sMB_SAPLY_E = DateAdd(DateInterval.Day, -1, (DateAdd(DateInterval.Month, 1, CDate(sTMP)))).ToString
            End If

            Dim isVadDate As Boolean = False
            If IsDate(sMB_SAPLY_S) AndAlso IsDate(sMB_SAPLY_E) Then
                isVadDate = True
            End If

            Dim mbCLASSList As New MB_CLASSList(Me.m_DBManager)
            If Not Utility.isValidateData(sMB_PLACE) AndAlso isVadDate = False Then
                mbCLASSList.LoadAPLY_2()
            ElseIf Utility.isValidateData(sMB_PLACE) AndAlso isVadDate = False Then
                mbCLASSList.LoadAPLY_MB_PLACE_2(sMB_PLACE)
            ElseIf Utility.isValidateData(sMB_PLACE) AndAlso isVadDate = True Then
                mbCLASSList.LoadAPLY_PLACE_SDATE_2(sMB_PLACE, CDate(sMB_SAPLY_S), CDate(sMB_SAPLY_E))
            ElseIf Not Utility.isValidateData(sMB_PLACE) AndAlso isVadDate = True Then
                mbCLASSList.LoadAPLY_SDATE_2(CDate(sMB_SAPLY_S), CDate(sMB_SAPLY_E))
            End If

            Me.RP_CLASS_2.DataSource = mbCLASSList.getCurrentDataSet.Tables(0)
            Me.RP_CLASS_2.DataBind()
        Catch ex As Exception
            Throw
        End Try
    End Sub
#End Region

#Region "進行中課程"
    Private Sub RP_CLASS_4_ItemDataBound(sender As Object, e As Web.UI.WebControls.RepeaterItemEventArgs) Handles RP_CLASS_4.ItemDataBound
        Try
            If e.Item.ItemType = ListItemType.Item OrElse e.Item.ItemType = ListItemType.AlternatingItem Then
                Dim DRV As DataRowView = CType(e.Item.DataItem, DataRowView)

                Dim sMB_SDATE As String = String.Empty
                If IsDate(DRV("MB_SDATE").ToString) Then
                    'sMB_SDATE = Utility.DateTransfer(CDate(DRV("MB_SDATE").ToString))
                    sMB_SDATE = MBSC_DATE(DRV("MB_SDATE"))
                End If
                Dim sMB_EWEEK As String = String.Empty
                If IsDate(DRV("MB_EDATE").ToString) Then
                    'sMB_EWEEK = Utility.DateTransfer(CDate(DRV("MB_EDATE").ToString))
                    sMB_EWEEK = MBSC_DATE(DRV("MB_EDATE"))
                End If
                Dim LTL_MB_SEDATE As Literal = e.Item.FindControl("LTL_MB_SEDATE")
                LTL_MB_SEDATE.Text = sMB_SDATE & "~" & sMB_EWEEK

                Dim sMB_SAPLY As String = String.Empty
                If IsDate(DRV("MB_SAPLY").ToString) Then
                    'sMB_SAPLY = Utility.DateTransfer(CDate(DRV("MB_SAPLY").ToString))
                    sMB_SAPLY = MBSC_DATE(DRV("MB_SAPLY"))
                End If
                Dim sMB_EAPLY As String = String.Empty
                If IsDate(DRV("MB_EAPLY").ToString) Then
                    'sMB_EAPLY = Utility.DateTransfer(CDate(DRV("MB_EAPLY").ToString))
                    sMB_EAPLY = MBSC_DATE(DRV("MB_EAPLY"))
                End If
                Dim LTL_MB_SEAPLY As Literal = e.Item.FindControl("LTL_MB_SEAPLY")
                LTL_MB_SEAPLY.Text = sMB_SAPLY & "~" & sMB_EAPLY

                If Utility.isValidateData(Me.m_sPLACE) Then
                    Dim TD_3_1 As HtmlTableCell = e.Item.FindControl("TD_3_1")
                    TD_3_1.Visible = False

                    Dim TD_3_2 As HtmlTableCell = e.Item.FindControl("TD_3_2")
                    TD_3_2.Visible = False

                    Dim TD_3_3 As HtmlTableCell = e.Item.FindControl("TD_3_3")
                    TD_3_3.Visible = False

                    Dim TD_3_4 As HtmlTableCell = e.Item.FindControl("TD_3_4")
                    TD_3_4.Visible = False

                    Dim TD_3_5 As HtmlTableCell = e.Item.FindControl("TD_3_5")
                    TD_3_5.Visible = False
                End If
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Private Sub btnQRY4_Click(sender As Object, e As EventArgs) Handles btnQRY4.Click
        Try
            Dim sMB_PLACE As String = String.Empty
            If Utility.isValidateData(Me.DDL_MB_PLACE_4.SelectedValue) Then
                sMB_PLACE = Me.DDL_MB_PLACE_4.SelectedItem.Text
            End If

            Dim sMB_SAPLY_S As String = String.Empty
            If IsNumeric(Me.DDL_MB_SDATE_SY_4.SelectedValue) AndAlso IsNumeric(Me.DDL_MB_SDATE_SM_4.SelectedValue) Then
                sMB_SAPLY_S = CDec(Me.DDL_MB_SDATE_SY_4.SelectedValue) & "/" & CDec(Me.DDL_MB_SDATE_SM_4.SelectedValue) & "/1"
            End If

            Dim sMB_SAPLY_E As String = String.Empty
            If IsNumeric(Me.DDL_MB_SDATE_EY_4.SelectedValue) AndAlso IsNumeric(Me.DDL_MB_SDATE_EM_4.SelectedValue) Then
                Dim sTMP As String = String.Empty
                sTMP = CDec(Me.DDL_MB_SDATE_EY_4.SelectedValue) & "/" & CDec(Me.DDL_MB_SDATE_EM_4.SelectedValue) & "/1"
                sMB_SAPLY_E = DateAdd(DateInterval.Day, -1, (DateAdd(DateInterval.Month, 1, CDate(sTMP)))).ToString
            End If

            Dim isVadDate As Boolean = False
            If IsDate(sMB_SAPLY_S) AndAlso IsDate(sMB_SAPLY_E) Then
                isVadDate = True
            End If

            Dim mbCLASSList As New MB_CLASSList(Me.m_DBManager)
            If Not Utility.isValidateData(sMB_PLACE) AndAlso isVadDate = False Then
                mbCLASSList.LoadAPLY_4()
            ElseIf Utility.isValidateData(sMB_PLACE) AndAlso isVadDate = False Then
                mbCLASSList.LoadAPLY_MB_PLACE_4(sMB_PLACE)
            ElseIf Utility.isValidateData(sMB_PLACE) AndAlso isVadDate = True Then
                mbCLASSList.LoadAPLY_PLACE_SDATE_4(sMB_PLACE, CDate(sMB_SAPLY_S), CDate(sMB_SAPLY_E))
            ElseIf Not Utility.isValidateData(sMB_PLACE) AndAlso isVadDate = True Then
                mbCLASSList.LoadAPLY_SDATE_4(CDate(sMB_SAPLY_S), CDate(sMB_SAPLY_E))
            End If

            Me.RP_CLASS_4.DataSource = mbCLASSList.getCurrentDataSet.Tables(0)
            Me.RP_CLASS_4.DataBind()
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Private Sub btnQRY5_Click(sender As Object, e As EventArgs) Handles btnQRY5.Click
        Try
            Dim sMB_PLACE As String = String.Empty
            If Utility.isValidateData(Me.DDL_MB_PLACE_5.SelectedValue) Then
                sMB_PLACE = Me.DDL_MB_PLACE_5.SelectedItem.Text
            End If

            Dim sMB_SAPLY_S As String = String.Empty
            If IsNumeric(Me.DDL_MB_SDATE_SY_5.SelectedValue) AndAlso IsNumeric(Me.DDL_MB_SDATE_SM_5.SelectedValue) Then
                sMB_SAPLY_S = CDec(Me.DDL_MB_SDATE_SY_5.SelectedValue) & "/" & CDec(Me.DDL_MB_SDATE_SM_5.SelectedValue) & "/1"
            End If

            Dim sMB_SAPLY_E As String = String.Empty
            If IsNumeric(Me.DDL_MB_SDATE_EY_5.SelectedValue) AndAlso IsNumeric(Me.DDL_MB_SDATE_EM_5.SelectedValue) Then
                Dim sTMP As String = String.Empty
                sTMP = CDec(Me.DDL_MB_SDATE_EY_5.SelectedValue) & "/" & CDec(Me.DDL_MB_SDATE_EM_5.SelectedValue) & "/1"
                sMB_SAPLY_E = DateAdd(DateInterval.Day, -1, (DateAdd(DateInterval.Month, 1, CDate(sTMP)))).ToString
            End If

            Dim isVadDate As Boolean = False
            If IsDate(sMB_SAPLY_S) AndAlso IsDate(sMB_SAPLY_E) Then
                isVadDate = True
            End If

            Dim mbCLASSList As New MB_CLASSList(Me.m_DBManager)
            If Not Utility.isValidateData(sMB_PLACE) AndAlso isVadDate = False Then
                mbCLASSList.LoadAPLY_5()
            ElseIf Utility.isValidateData(sMB_PLACE) AndAlso isVadDate = False Then
                mbCLASSList.LoadAPLY_MB_PLACE_5(sMB_PLACE)
            ElseIf Utility.isValidateData(sMB_PLACE) AndAlso isVadDate = True Then
                mbCLASSList.LoadAPLY_PLACE_SDATE_5(sMB_PLACE, CDate(sMB_SAPLY_S), CDate(sMB_SAPLY_E))
            ElseIf Not Utility.isValidateData(sMB_PLACE) AndAlso isVadDate = True Then
                mbCLASSList.LoadAPLY_SDATE_5(CDate(sMB_SAPLY_S), CDate(sMB_SAPLY_E))
            End If

            Me.RP_CLASS_4.DataSource = mbCLASSList.getCurrentDataSet.Tables(0)
            Me.RP_CLASS_4.DataBind()
        Catch ex As Exception
            Throw
        End Try
    End Sub
#End Region

#Region "關閉課程(已完成之課程)"
    Private Sub RP_CLASS_3_ItemDataBound(sender As Object, e As Web.UI.WebControls.RepeaterItemEventArgs) Handles RP_CLASS_3.ItemDataBound
        Try
            If e.Item.ItemType = ListItemType.Item OrElse e.Item.ItemType = ListItemType.AlternatingItem Then
                Dim DRV As DataRowView = CType(e.Item.DataItem, DataRowView)

                Dim sMB_SDATE As String = String.Empty
                If IsDate(DRV("MB_SDATE").ToString) Then
                    'sMB_SDATE = Utility.DateTransfer(CDate(DRV("MB_SDATE").ToString))
                    sMB_SDATE = MBSC_DATE(DRV("MB_SDATE"))
                End If
                Dim sMB_EWEEK As String = String.Empty
                If IsDate(DRV("MB_EDATE").ToString) Then
                    'sMB_EWEEK = Utility.DateTransfer(CDate(DRV("MB_EDATE").ToString))
                    sMB_EWEEK = MBSC_DATE(DRV("MB_EDATE"))
                End If
                Dim LTL_MB_SEDATE As Literal = e.Item.FindControl("LTL_MB_SEDATE")
                LTL_MB_SEDATE.Text = sMB_SDATE & "~" & sMB_EWEEK

                Dim sMB_SAPLY As String = String.Empty
                If IsDate(DRV("MB_SAPLY").ToString) Then
                    'sMB_SAPLY = Utility.DateTransfer(CDate(DRV("MB_SAPLY").ToString))
                    sMB_SAPLY = MBSC_DATE(DRV("MB_SAPLY"))
                End If
                Dim sMB_EAPLY As String = String.Empty
                If IsDate(DRV("MB_EAPLY").ToString) Then
                    'sMB_EAPLY = Utility.DateTransfer(CDate(DRV("MB_EAPLY").ToString))
                    sMB_EAPLY = MBSC_DATE(DRV("MB_EAPLY"))
                End If
                Dim LTL_MB_SEAPLY As Literal = e.Item.FindControl("LTL_MB_SEAPLY")
                LTL_MB_SEAPLY.Text = sMB_SAPLY & "~" & sMB_EAPLY
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Private Sub btnQRY3_Click(sender As Object, e As EventArgs) Handles btnQRY3.Click
        Try
            Dim sMB_PLACE As String = String.Empty
            If Utility.isValidateData(Me.DDL_MB_PLACE_3.SelectedValue) Then
                sMB_PLACE = Me.DDL_MB_PLACE_3.SelectedItem.Text
            End If

            Dim sMB_SAPLY_S As String = String.Empty
            If IsNumeric(Me.DDL_MB_SDATE_SY_3.SelectedValue) AndAlso IsNumeric(Me.DDL_MB_SDATE_SM_3.SelectedValue) Then
                sMB_SAPLY_S = CDec(Me.DDL_MB_SDATE_SY_3.SelectedValue) & "/" & CDec(Me.DDL_MB_SDATE_SM_3.SelectedValue) & "/1"
            End If

            Dim sMB_SAPLY_E As String = String.Empty
            If IsNumeric(Me.DDL_MB_SDATE_EY_3.SelectedValue) AndAlso IsNumeric(Me.DDL_MB_SDATE_EM_3.SelectedValue) Then
                Dim sTMP As String = String.Empty
                sTMP = CDec(Me.DDL_MB_SDATE_EY_3.SelectedValue) & "/" & CDec(Me.DDL_MB_SDATE_EM_3.SelectedValue) & "/1"
                sMB_SAPLY_E = DateAdd(DateInterval.Day, -1, (DateAdd(DateInterval.Month, 1, CDate(sTMP)))).ToString
            End If

            Dim isVadDate As Boolean = False
            If IsDate(sMB_SAPLY_S) AndAlso IsDate(sMB_SAPLY_E) Then
                isVadDate = True
            End If

            Dim mbCLASSList As New MB_CLASSList(Me.m_DBManager)
            If Not Utility.isValidateData(sMB_PLACE) AndAlso isVadDate = False Then
                mbCLASSList.LoadAPLY_3()
            ElseIf Utility.isValidateData(sMB_PLACE) AndAlso isVadDate = False Then
                mbCLASSList.LoadAPLY_MB_PLACE_3(sMB_PLACE)
            ElseIf Utility.isValidateData(sMB_PLACE) AndAlso isVadDate = True Then
                mbCLASSList.LoadAPLY_PLACE_SDATE_3(sMB_PLACE, CDate(sMB_SAPLY_S), CDate(sMB_SAPLY_E))
            ElseIf Not Utility.isValidateData(sMB_PLACE) AndAlso isVadDate = True Then
                mbCLASSList.LoadAPLY_SDATE_3(CDate(sMB_SAPLY_S), CDate(sMB_SAPLY_E))
            End If

            Me.RP_CLASS_3.DataSource = mbCLASSList.getCurrentDataSet.Tables(0)
            Me.RP_CLASS_3.DataBind()
        Catch ex As Exception
            Throw
        End Try
    End Sub
#End Region

    Function MBSC_DATE(ByVal sDate As Object) As String
        Try
            If Utility.isValidateData(sDate) AndAlso IsDate(sDate.ToString) Then
                Dim D_Date As Date = CDate(sDate.ToString)

                Return Utility.FillZero(D_Date.Year, 4) & "/" & Utility.FillZero(D_Date.Month, 2) & "/" & Utility.FillZero(D_Date.Day, 2)
            End If

            Return String.Empty
        Catch ex As Exception
            Throw
        End Try
    End Function

    Function getMB_CLASS_NAME(ByVal MB_CLASS_NAME As Object, ByVal MB_BATCH As Object) As String
        Try
            If Utility.isValidateData(MB_CLASS_NAME) Then
                If IsNumeric(MB_BATCH) AndAlso CDec(MB_BATCH) > 0 Then
                    Return MB_CLASS_NAME & "　第" & MB_BATCH & "梯次"
                Else
                    Return MB_CLASS_NAME
                End If
            End If

            Return String.Empty
        Catch ex As Exception
            Throw
        End Try
    End Function

    Private Sub RP_CLASS_5_ItemDataBound(sender As Object, e As Web.UI.WebControls.RepeaterItemEventArgs) Handles RP_CLASS_5.ItemDataBound
        Try
            If e.Item.ItemType = ListItemType.Item OrElse e.Item.ItemType = ListItemType.AlternatingItem Then
                Dim DRV As DataRowView = CType(e.Item.DataItem, DataRowView)

                Dim sMB_SDATE As String = String.Empty
                If IsDate(DRV("MB_SDATE").ToString) Then
                    'sMB_SDATE = Utility.DateTransfer(CDate(DRV("MB_SDATE").ToString))
                    sMB_SDATE = MBSC_DATE(DRV("MB_SDATE"))
                End If
                Dim sMB_EWEEK As String = String.Empty
                If IsDate(DRV("MB_EDATE").ToString) Then
                    'sMB_EWEEK = Utility.DateTransfer(CDate(DRV("MB_EDATE").ToString))
                    sMB_EWEEK = MBSC_DATE(DRV("MB_EDATE"))
                End If
                Dim LTL_MB_SEDATE As Literal = e.Item.FindControl("LTL_MB_SEDATE")
                LTL_MB_SEDATE.Text = sMB_SDATE & "~" & sMB_EWEEK

                'Dim sMB_SAPLY As String = String.Empty
                'If IsDate(DRV("MB_SAPLY").ToString) Then
                '    'sMB_SAPLY = Utility.DateTransfer(CDate(DRV("MB_SAPLY").ToString))
                '    sMB_SAPLY = MBSC_DATE(DRV("MB_SAPLY"))
                'End If
                'Dim sMB_EAPLY As String = String.Empty
                'If IsDate(DRV("MB_EAPLY").ToString) Then
                '    'sMB_EAPLY = Utility.DateTransfer(CDate(DRV("MB_EAPLY").ToString))
                '    sMB_EAPLY = MBSC_DATE(DRV("MB_EAPLY"))
                'End If
                'Dim LTL_MB_SEDATE As Literal = e.Item.FindControl("LTL_MB_SEDATE")
                'LTL_MB_SEDATE.Text = sMB_SAPLY & "~" & sMB_EAPLY

                If Me.is76User(Session("UserId")) Then
                    Dim btnMB_ALERT1 As Button = e.Item.FindControl("btnMB_ALERT1")
                    Dim btnMB_ALERT2 As Button = e.Item.FindControl("btnMB_ALERT2")

                    If CDate(DRV("MB_SDATE").ToString).AddDays(-1 * DRV("MB_ALERT1_DAY")) <= Now Then
                        If Not Utility.isValidateData(DRV("MB_ALERT1")) Then
                            btnMB_ALERT1.Visible = True
                        Else
                            btnMB_ALERT1.Enabled = False
                            btnMB_ALERT1.Text = "已經發第一封提醒信"

                            'If CDate(DRV("MB_SDATE").ToString).AddDays(-1 * DRV("MB_ALERT2_DAY")) <= Now Then
                            '    btnMB_ALERT2.Visible = True
                            'Else
                            '    btnMB_ALERT2.Visible = False
                            'End If

                            If CDate(DRV("MB_SDATE").ToString) > Now Then
                                btnMB_ALERT2.Visible = True
                            Else
                                btnMB_ALERT2.Visible = False
                            End If
                        End If
                    Else
                        btnMB_ALERT1.Visible = False
                        btnMB_ALERT2.Visible = False
                    End If
                Else
                    Dim TD_5_6 As HtmlTableCell = e.Item.FindControl("TD_5_6")
                    TD_5_6.Visible = False
                End If

                If Utility.isValidateData(Me.m_sPLACE) Then
                    Dim TD_5_1 As HtmlTableCell = e.Item.FindControl("TD_5_1")
                    TD_5_1.Visible = False

                    Dim TD_5_2 As HtmlTableCell = e.Item.FindControl("TD_5_2")
                    TD_5_2.Visible = False

                    Dim TD_5_3 As HtmlTableCell = e.Item.FindControl("TD_5_3")
                    TD_5_3.Visible = False

                    Dim TD_5_4 As HtmlTableCell = e.Item.FindControl("TD_5_4")
                    TD_5_4.Visible = False

                    Dim TD_5_5 As HtmlTableCell = e.Item.FindControl("TD_5_5")
                    TD_5_5.Visible = False

                    Dim TD_5_6 As HtmlTableCell = e.Item.FindControl("TD_5_6")
                    TD_5_6.Visible = False
                End If

                If DRV("MB_ALERT2").ToString = "Y" Then
                    Dim LTL_MB_ALERT2 As Label = e.Item.FindControl("LTL_MB_ALERT2")
                    LTL_MB_ALERT2.Text = "<BR/>已發過"
                End If
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub

End Class