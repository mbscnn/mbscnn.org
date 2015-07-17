Imports MBSC.MB_OP
Public Class PageTab
    Inherits System.Web.UI.UserControl

    'Dim m_DBManager As com.Azion.NET.VB.DatabaseManager

    'Dim m_iUPCODE_LV As Integer = 37

    'Dim m_iUPCODE_78 As Integer = 78

    'Dim m_iLEVEL As Integer = 0

    'Dim m_iUPCODE_Admin As Integer = 76

    'Dim m_sCODEID As String = String.Empty

#Region "Page Events"
    'Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    '    Try
    '        If IsNumeric(Session("admin")) OrElse com.Azion.EloanUtility.ValidateUtility.isValidateData(Session("USERID")) Then
    '            Me.PLH_LOGIN.Visible = False
    '            Me.PLH_LOGOUT.Visible = True
    '        Else
    '            Me.PLH_LOGIN.Visible = True
    '            Me.PLH_LOGOUT.Visible = False
    '        End If

    '        Me.m_DBManager = UIShareFun.getDataBaseManager

    '        If IsNumeric(Session("admin")) Then
    '            Me.m_iLEVEL = CInt(Session("admin"))
    '        End If

    '        Me.m_sCODEID = "" & Request.QueryString("CLASS")

    '        '會員系統
    '        Me.m_iUPCODE_78 = com.Azion.EloanUtility.CodeList.getAppSettings("UPCODE78")

    '        If Not Page.IsPostBack Then
    '            If com.Azion.EloanUtility.ValidateUtility.isValidateData(Session("USERID")) Then
    '                Dim mbACCT As New MB_ACCT(Me.m_DBManager)
    '                If mbACCT.LoadByPK(Session("USERID")) Then
    '                    Me.LTL_MB_NAME.Text = mbACCT.getString("MB_NAME")
    '                End If
    '            End If

    '            'Me.Bind_RP_Tab()

    '            'Me.IMGTop.ImageUrl = com.Azion.EloanUtility.UIUtility.getRootPath & "/img/mbscbanner02.jpg"

    '            If Not com.Azion.EloanUtility.ValidateUtility.isValidateData(Me.m_sCODEID) Then
    '                Dim apCODEList As New AP_CODEList(Me.m_DBManager)
    '                apCODEList.setSQLCondition(" ORDER BY SORTNO ")
    '                apCODEList.loadByUpCode(Me.m_iUPCODE_LV)
    '                Dim ROW_LV1() As DataRow = apCODEList.getCurrentDataSet.Tables(0).Select("DISABLED='1'")
    '                If Not IsNothing(ROW_LV1) AndAlso ROW_LV1.Length > 0 Then
    '                    Me.m_sCODEID = ROW_LV1(0)("CODEID")
    '                End If
    '            End If
    '            Dim apCODE As New AP_CODE(Me.m_DBManager)
    '            If apCODE.loadByPK(Me.m_sCODEID) Then
    '                Me.LTL_TAB_TITLE.Text = apCODE.getString("TEXT")
    '                Me.LTL_TAB_TITLE_HID.Text = apCODE.getString("TEXT")

    '                'If Utility.isValidateData(apCODE.getString("TMPURL")) Then
    '                '    Dim sScript As String = String.Empty
    '                '    sScript = "<script language='javascript' >" & vbCrLf
    '                '    sScript &= "window.open('" & apCODE.getString("TMPURL") & "',null,'height=' + window.screen.height + ', width=' + window.screen.width + ', top=0, left=0, toolbar=yes, menubar=yes, scrollbars=no, resizable=no,location=n o, status=no');" & vbCrLf
    '                '    sScript &= "</" & "script>"
    '                '    Response.Write(sScript)
    '                'End If
    '            End If
    '        End If
    '    Catch ex As Exception
    '        com.Azion.EloanUtility.UIUtility.showErrMsg(Me.Page, ex)
    '    End Try
    'End Sub

    'Private Sub Page_Unload(sender As Object, e As System.EventArgs) Handles Me.Unload
    '    UIShareFun.releaseConnection(Me.m_DBManager)
    'End Sub
#End Region

    'Sub Bind_RP_Tab()
    '    Try
    '        Dim apCODEList As New AP_CODEList(Me.m_DBManager)
    '        apCODEList.setSQLCondition(" ORDER BY SORTNO ")
    '        apCODEList.loadByUpCode(Me.m_iUPCODE_LV)

    '        Dim ROW_LV1() As DataRow = apCODEList.getCurrentDataSet.Tables(0).Select("DISABLED='1' AND LEVEL<=" & Me.m_iLEVEL)
    '        Me.RP_Tab.DataSource = ROW_LV1
    '        Me.RP_Tab.DataBind()
    '    Catch ex As Exception
    '        Throw
    '    End Try
    'End Sub

    'Private Sub RP_Tab_ItemDataBound(sender As Object, e As System.Web.UI.WebControls.RepeaterItemEventArgs) Handles RP_Tab.ItemDataBound
    '    Try
    '        If e.Item.ItemType = ListItemType.Item OrElse e.Item.ItemType = ListItemType.AlternatingItem Then
    '            Dim DRV As DataRow = CType(e.Item.DataItem, DataRow)

    '            Dim apCODEList As New AP_CODEList(Me.m_DBManager)
    '            apCODEList.setSQLCondition(" ORDER BY SORTNO ")
    '            Dim ROW_LV2() As DataRow = Nothing
    '            If DRV("CODEID") = Me.m_iUPCODE_78 AndAlso Utility.isValidateData(Session("UserId")) Then
    '                apCODEList.Load_TAB_AUTH(DRV("CODEID"), Session("UserId"))
    '                ROW_LV2 = apCODEList.getCurrentDataSet.Tables(0).Select("DISABLED='1'")
    '            Else
    '                apCODEList.loadByUpCode(DRV("CODEID"))
    '                ROW_LV2 = apCODEList.getCurrentDataSet.Tables(0).Select("DISABLED='1' AND LEVEL<=" & Me.m_iLEVEL)
    '            End If

    '            Dim RP_Tab_LV2 As Repeater = e.Item.FindControl("RP_Tab_LV2")
    '            RP_Tab_LV2.DataSource = ROW_LV2
    '            RP_Tab_LV2.DataBind()

    '            Dim HPL_LV1 As HyperLink = e.Item.FindControl("HPL_LV1")
    '            If Utility.isValidateData(DRV("NOTE")) Then
    '                If DRV("NOTE").ToString = "NewsList.aspx" Then
    '                    HPL_LV1.NavigateUrl = com.Azion.EloanUtility.UIUtility.getRootPath & "/" & DRV("NOTE")
    '                Else
    '                    If InStr(DRV("NOTE"), "?") > 0 Then
    '                        HPL_LV1.NavigateUrl = com.Azion.EloanUtility.UIUtility.getRootPath & "/" & DRV("NOTE") & "&CLASS=" & DRV("CODEID")
    '                    Else
    '                        HPL_LV1.NavigateUrl = com.Azion.EloanUtility.UIUtility.getRootPath & "/" & DRV("NOTE") & "?CLASS=" & DRV("CODEID")
    '                    End If
    '                End If
    '            Else
    '                HPL_LV1.Attributes("href") = "#"
    '            End If
    '        End If
    '    Catch ex As Exception
    '        com.Azion.EloanUtility.UIUtility.showErrMsg(Me.Page, ex)
    '    End Try
    'End Sub

    'Sub RP_Tab_LV2_ItemDataBound(sender As Object, e As System.Web.UI.WebControls.RepeaterItemEventArgs)
    '    Try
    '        If e.Item.ItemType = ListItemType.Item OrElse e.Item.ItemType = ListItemType.AlternatingItem Then
    '            Dim DRV As DataRow = CType(e.Item.DataItem, DataRow)

    '            Dim HPL_LV2 As HyperLink = e.Item.FindControl("HPL_LV2")
    '            'If com.Azion.EloanUtility.ValidateUtility.isValidateData(DRV("TMPURL")) Then
    '            '    Dim sScript As String = String.Empty
    '            '    'sScript = "window.showModalDialog('" & DRV("TMPURL") & "',self,'dialogWidth:' + window.screen.width + 'px; dialogHeight:' + window.screen.height + 'px; center:yes;scroll:1;status:0;help:0;resizable:0');"
    '            '    sScript = "window.open('" & DRV("TMPURL") & "',null,'height=' + window.screen.height + ', width=' + window.screen.width + ', top=0, left=0, toolbar=yes, menubar=yes, scrollbars=no, resizable=no,location=n o, status=no');"
    '            '    HPL_LV2.Attributes("onmousedown") = sScript
    '            'End If
    '            If Utility.isValidateData(DRV("NOTE").ToString) Then
    '                If UCase(Left(DRV("NOTE").ToString, 4)) = "HTTP" Then
    '                    'HPL_LV2.Attributes("href") = "#"
    '                    'HPL_LV2.Attributes("onclick") = "window.open('" & DRV("NOTE").ToString & "',null,'height=' + window.screen.height + ', width=' + window.screen.width + ', top=0, left=0, toolbar=yes, menubar=yes, scrollbars=no, resizable=no,location=n o, status=no');" & vbCrLf
    '                Else
    '                    If InStr(DRV("NOTE"), "?") > 0 Then
    '                        HPL_LV2.NavigateUrl = com.Azion.EloanUtility.UIUtility.getRootPath & "/" & DRV("NOTE") & "&CLASS=" & DRV("CODEID")
    '                    Else
    '                        HPL_LV2.NavigateUrl = com.Azion.EloanUtility.UIUtility.getRootPath & "/" & DRV("NOTE") & "?CLASS=" & DRV("CODEID")
    '                    End If
    '                End If
    '            Else
    '                HPL_LV2.Attributes("href") = "#"
    '            End If
    '        End If
    '    Catch ex As Exception
    '        com.Azion.EloanUtility.UIUtility.showErrMsg(Me.Page, ex)
    '    End Try
    'End Sub

    'Private Sub btLogin_Click(sender As Object, e As System.EventArgs) Handles btLogin.Click
    '    Try
    '        If Not Utility.isValidateData(Trim(Me.txt_UserId.Text)) Then
    '            com.Azion.EloanUtility.UIUtility.alert("請輸入e-Mail")
    '            Return
    '        End If

    '        If Trim(Me.txt_UserId.Text) = "帳號(e-Mail)" Then
    '            com.Azion.EloanUtility.UIUtility.alert("請輸入e-Mail")
    '            Return
    '        End If

    '        If Not Utility.isValidateData(Trim(Me.txt_Password.Text)) Then
    '            com.Azion.EloanUtility.UIUtility.alert("請輸入密碼")
    '            Return
    '        End If

    '        Dim apCODEList As New AP_CODEList(Me.m_DBManager)
    '        apCODEList.loadByUpCode(Me.m_iUPCODE_Admin)
    '        Dim sFILTER As String = String.Empty
    '        sFILTER = "VALUE='" & Trim(Me.txt_Password.Text) & "' AND TEXT='" & Trim(Me.txt_UserId.Text) & "'"
    '        Dim ROW_SELECT() As DataRow = apCODEList.getCurrentDataSet.Tables(0).Select(sFILTER)

    '        If Not IsNothing(ROW_SELECT) AndAlso ROW_SELECT.Length > 0 Then
    '            Session("admin") = ROW_SELECT(0)("LEVEL")

    '            Session("USERID") = Trim(Me.txt_UserId.Text)

    '            Response.Redirect(Request.AppRelativeCurrentExecutionFilePath & Request.Url.Query)
    '        Else
    '            Dim mbACCT As New MB_ACCT(Me.m_DBManager)
    '            If mbACCT.LoadByPK(Trim(Me.txt_UserId.Text)) Then
    '                If mbACCT.getString("MB_APV") = "Y" Then
    '                    If Trim(Me.txt_Password.Text) = mbACCT.getString("MB_PSW") Then
    '                        Session("admin") = "1"

    '                        Session("USERID") = mbACCT.getString("MB_ACCT")

    '                        Response.Redirect(Request.AppRelativeCurrentExecutionFilePath & Request.Url.Query)
    '                    Else
    '                        com.Azion.EloanUtility.UIUtility.alert("密碼輸入錯誤")
    '                    End If
    '                Else
    '                    com.Azion.EloanUtility.UIUtility.alert("帳號(e-Mail)尚未啟用，請到您的信箱點集本中心驗證信連結啟用")
    '                End If
    '            Else
    '                com.Azion.EloanUtility.UIUtility.alert("帳號(e-Mail)不存在，請先註冊為會員")
    '            End If
    '        End If
    '    Catch ex As System.Threading.ThreadAbortException
    '    Catch ex As Exception
    '        com.Azion.EloanUtility.UIUtility.showErrMsg(Me.Page, ex)
    '    End Try
    'End Sub

    'Private Sub btLogout_Click(sender As Object, e As System.EventArgs) Handles btLogout.Click
    '    Session("admin") = Nothing

    '    Session("USERID") = Nothing

    '    'Response.Redirect(Request.AppRelativeCurrentExecutionFilePath & Request.Url.Query)
    '    Response.Redirect(com.Azion.EloanUtility.UIUtility.getRootPath & "/NewsList.aspx")
    'End Sub

    'Private Sub btSign_Click(sender As Object, e As System.EventArgs) Handles btSign.Click
    '    Session("ValidateNumber") = Nothing

    '    Dim sURL As String = String.Empty

    '    sURL = com.Azion.EloanUtility.UIUtility.getRootPath & "/MNT/MBMnt_Reg_01_v01.aspx"

    '    Response.Redirect(sURL)
    'End Sub

    'Private Sub btnReMail_Click(sender As Object, e As EventArgs) Handles btnReMail.Click
    '    Try
    '        If Not Utility.isValidateData(Trim(Me.txt_UserId.Text)) Then
    '            com.Azion.EloanUtility.UIUtility.showJSMsg(Me.Page, "請輸入e-Mail")
    '            com.Azion.EloanUtility.UIUtility.showErrMsg(Me.Page, "請輸入e-Mail")
    '            Return
    '        End If
    '        If Not Utility.isValidateData(Trim(Me.txt_Password.Text)) Then
    '            com.Azion.EloanUtility.UIUtility.showJSMsg(Me.Page, "請輸入密碼")
    '            com.Azion.EloanUtility.UIUtility.showErrMsg(Me.Page, "請輸入密碼")
    '            Return
    '        End If
    '        Dim mbACCT As New MB_ACCT(Me.m_DBManager)
    '        If Not mbACCT.LoadByPK(Trim(Me.txt_UserId.Text)) Then
    '            com.Azion.EloanUtility.UIUtility.showJSMsg(Me.Page, "您尚未註冊成為會員")
    '            com.Azion.EloanUtility.UIUtility.showErrMsg(Me.Page, "您尚未註冊成為會員")
    '            Return
    '        Else
    '            If Trim(Me.txt_Password.Text) <> mbACCT.getString("MB_PSW") Then
    '                com.Azion.EloanUtility.UIUtility.showJSMsg(Me.Page, "密碼輸入錯誤")
    '                com.Azion.EloanUtility.UIUtility.showErrMsg(Me.Page, "密碼輸入錯誤")
    '                Return
    '            Else
    '                Dim sMailTos() As String = {Trim(Me.txt_UserId.Text)}

    '                Dim sMailSub As String = String.Empty
    '                sMailSub = com.Azion.EloanUtility.FileUtility.getAppSettings("MAILSUB") & Now.ToShortDateString & " " & Now.ToShortTimeString

    '                Dim sMB_APVID As String = mbACCT.getString("MB_APVID")

    '                Dim sMailBody As String = String.Empty

    '                sMailBody = UIShareFun.getMailBody(sMB_APVID, Trim(Me.txt_UserId.Text), Trim(mbACCT.getString("MB_NAME")))

    '                If Not com.Azion.EloanUtility.NetUtility.GMail_Send(sMailTos, Nothing, sMailSub, sMailBody, True, Nothing, False) Then
    '                    com.Azion.EloanUtility.UIUtility.alert("驗證信發送失敗，請確認您的e-Mail網址是否正確或稍後再試")
    '                    com.Azion.EloanUtility.UIUtility.showErrMsg(Me.Page, "驗證信發送失敗，請確認您的e-Mail網址是否正確或稍後再試")
    '                Else
    '                    com.Azion.EloanUtility.UIUtility.alert("系統已發送驗證至您的e-Mail，請到您的信箱完成帳號啟用【若找不到驗證信，請試試垃圾郵件資料匣】")
    '                    com.Azion.EloanUtility.UIUtility.showErrMsg(Me.Page, "系統已發送驗證至您的e-Mail，請到您的信箱完成帳號啟用【若找不到驗證信，請試試垃圾郵件資料匣】")
    '                End If
    '            End If
    '        End If
    '    Catch ex As Exception
    '        com.Azion.EloanUtility.UIUtility.showErrMsg(Me.Page, ex)
    '    End Try
    'End Sub
End Class