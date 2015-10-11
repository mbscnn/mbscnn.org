Imports MBSC.MB_OP
Imports MBSC.UICtl

Public Class MBSignIn_01_v01
    Inherits System.Web.UI.Page

    Dim m_DBManager As com.Azion.NET.VB.DatabaseManager

    Dim m_iUPCODE_Admin As Integer = 76


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.m_DBManager = UIShareFun.getDataBaseManager

            If Not Page.IsPostBack Then
                Me.IMG_Vad.Src = com.Azion.EloanUtility.UIUtility.getRootPath & "/Module/ValidateNumber.ashx"

                '點圖片重新整理
                Me.IMG_Vad.Attributes("onclick") = "this.src='" & _
                                                   com.Azion.EloanUtility.UIUtility.getRootPath & "/Module/ValidateNumber.ashx?" & _
                                                  "'+Math.random();"
            End If
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

    Private Sub MBSignIn_01_v01_Unload(sender As Object, e As EventArgs) Handles Me.Unload
        UIShareFun.releaseConnection(Me.m_DBManager)
    End Sub


    Private Sub btLogin_Click(sender As Object, e As System.EventArgs) Handles btLogin.Click
        Try
            If Not Utility.isValidateData(Trim(Me.txt_UserId.Text)) Then
                com.Azion.EloanUtility.UIUtility.alert("請輸入e-Mail")
                Return
            End If

            If Trim(Me.txt_UserId.Text) = "帳號(e-Mail)" Then
                com.Azion.EloanUtility.UIUtility.alert("請輸入e-Mail")
                Return
            End If

            If Not Utility.isValidateData(Trim(Me.txt_Password.Text)) Then
                com.Azion.EloanUtility.UIUtility.alert("請輸入密碼")
                Return
            End If

            Dim apCODEList As New AP_CODEList(Me.m_DBManager)
            apCODEList.loadByUpCode(Me.m_iUPCODE_Admin)
            Dim sFILTER As String = String.Empty
            sFILTER = "VALUE='" & Trim(Me.txt_Password.Text) & "' AND TEXT='" & Trim(Me.txt_UserId.Text) & "'"
            Dim ROW_SELECT() As DataRow = apCODEList.getCurrentDataSet.Tables(0).Select(sFILTER)

            If Not IsNothing(ROW_SELECT) AndAlso ROW_SELECT.Length > 0 Then
                Session("admin") = ROW_SELECT(0)("LEVEL")

                Session("USERID") = Trim(Me.txt_UserId.Text)

                Session("BRID") = ROW_SELECT(0)("BRID").ToString

                Dim DT_177 As DataTable = Nothing
                Try
                    apCODEList.clear()
                    apCODEList.loadByUpCode("177")
                    DT_177 = apCODEList.getCurrentDataSet.Tables(0)

                    Dim ROW_AREA() As DataRow = Nothing
                    ROW_AREA = DT_177.Select("VALUE='" & ROW_SELECT(0)("BRID").ToString & "'")
                    If Not IsNothing(ROW_AREA) AndAlso ROW_AREA.Length > 0 Then
                        Session("AREA") = ROW_AREA(0)("NOTE")
                    End If
                Catch ex As Exception
                    Throw
                Finally
                    If Not IsNothing(DT_177) Then
                        DT_177.Dispose()
                    End If
                End Try

                'Response.Redirect(Request.AppRelativeCurrentExecutionFilePath & Request.Url.Query)
                Response.Redirect(com.Azion.EloanUtility.UIUtility.getRootPath & "/NewsList.aspx")
            Else
                Dim mbACCT As New MB_ACCT(Me.m_DBManager)
                If mbACCT.LoadByPK(Trim(Me.txt_UserId.Text)) Then
                    If mbACCT.getString("MB_APV") = "Y" Then
                        If Trim(Me.txt_Password.Text) = mbACCT.getString("MB_PSW") Then
                            Session("admin") = "1"

                            Session("USERID") = mbACCT.getString("MB_ACCT")

                            'Response.Redirect(Request.AppRelativeCurrentExecutionFilePath & Request.Url.Query)
                            Response.Redirect(com.Azion.EloanUtility.UIUtility.getRootPath & "/NewsList.aspx")
                        Else
                            com.Azion.EloanUtility.UIUtility.alert("密碼輸入錯誤")
                        End If
                    Else
                        com.Azion.EloanUtility.UIUtility.alert("帳號(e-Mail)尚未啟用，請到您的信箱點集本中心驗證信連結啟用")
                    End If
                Else
                    com.Azion.EloanUtility.UIUtility.alert("帳號(e-Mail)不存在，請先註冊為會員")
                End If
            End If
        Catch ex As System.Threading.ThreadAbortException
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me.Page, ex)
        End Try
    End Sub

    Private Sub btSign_Click(sender As Object, e As System.EventArgs) Handles btSign.Click
        Session("ValidateNumber") = Nothing

        Dim sURL As String = String.Empty

        sURL = com.Azion.EloanUtility.UIUtility.getRootPath & "/MNT/MBMnt_Reg_01_v01.aspx"

        Response.Redirect(sURL)
    End Sub

    Private Sub btnReMail_Click(sender As Object, e As EventArgs) Handles btnReMail.Click
        Try
            If Not Utility.isValidateData(Trim(Me.txt_UserId.Text)) Then
                com.Azion.EloanUtility.UIUtility.showJSMsg(Me.Page, "請輸入e-Mail")
                com.Azion.EloanUtility.UIUtility.showErrMsg(Me.Page, "請輸入e-Mail")
                Return
            End If
            If Not Utility.isValidateData(Trim(Me.txt_Password.Text)) Then
                com.Azion.EloanUtility.UIUtility.showJSMsg(Me.Page, "請輸入密碼")
                com.Azion.EloanUtility.UIUtility.showErrMsg(Me.Page, "請輸入密碼")
                Return
            End If
            Dim mbACCT As New MB_ACCT(Me.m_DBManager)
            If Not mbACCT.LoadByPK(Trim(Me.txt_UserId.Text)) Then
                com.Azion.EloanUtility.UIUtility.showJSMsg(Me.Page, "您尚未註冊成為會員")
                com.Azion.EloanUtility.UIUtility.showErrMsg(Me.Page, "您尚未註冊成為會員")
                Return
            Else
                If Trim(Me.txt_Password.Text) <> mbACCT.getString("MB_PSW") Then
                    com.Azion.EloanUtility.UIUtility.showJSMsg(Me.Page, "密碼輸入錯誤")
                    com.Azion.EloanUtility.UIUtility.showErrMsg(Me.Page, "密碼輸入錯誤")
                    Return
                Else
                    Dim sMailTos() As String = {Trim(Me.txt_UserId.Text)}

                    Dim sMailSub As String = String.Empty
                    sMailSub = com.Azion.EloanUtility.FileUtility.getAppSettings("MAILSUB") & Now.ToShortDateString & " " & Now.ToShortTimeString

                    Dim sMB_APVID As String = mbACCT.getString("MB_APVID")

                    Dim sMailBody As String = String.Empty

                    sMailBody = UIShareFun.getMailBody(sMB_APVID, Trim(Me.txt_UserId.Text), Trim(mbACCT.getString("MB_NAME")))

                    If Not com.Azion.EloanUtility.NetUtility.GMail_Send(sMailTos, Nothing, sMailSub, sMailBody, True, Nothing, False) Then
                        com.Azion.EloanUtility.UIUtility.alert("驗證信發送失敗，請確認您的e-Mail網址是否正確或稍後再試")
                        com.Azion.EloanUtility.UIUtility.showErrMsg(Me.Page, "驗證信發送失敗，請確認您的e-Mail網址是否正確或稍後再試")
                    Else
                        com.Azion.EloanUtility.UIUtility.alert("系統已發送驗證至您的e-Mail，請到您的信箱完成帳號啟用【若找不到驗證信，請試試垃圾郵件資料匣】")
                        com.Azion.EloanUtility.UIUtility.showErrMsg(Me.Page, "系統已發送驗證至您的e-Mail，請到您的信箱完成帳號啟用【若找不到驗證信，請試試垃圾郵件資料匣】")
                    End If
                End If
            End If
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me.Page, ex)
        End Try
    End Sub

    Private Sub btnFPass_Click(sender As Object, e As EventArgs) Handles btnFPass.Click
        Try
            If Not Utility.isValidateData(Trim(Me.txt_UserId.Text)) Then
                com.Azion.EloanUtility.UIUtility.alert("請輸入e-Mail")
                com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請輸入e-Mail")
                Return
            End If

            If Trim(Me.txt_UserId.Text) = "帳號(e-Mail)" Then
                com.Azion.EloanUtility.UIUtility.alert("請輸入e-Mail")
                com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請輸入e-Mail")
                Return
            End If

            Dim mbACCT As New MB_ACCT(Me.m_DBManager)
            If mbACCT.LoadByPK(Trim(Me.txt_UserId.Text)) Then
                If mbACCT.getString("MB_APV") <> "Y" Then
                    com.Azion.EloanUtility.UIUtility.alert("帳號(e-Mail)尚未啟用，請到您的信箱點集本中心驗證信連結啟用")
                    com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "帳號(e-Mail)尚未啟用，請到您的信箱點集本中心驗證信連結啟用")
                    Return
                End If

                Me.PLH_NUMBER.Visible = True
            Else
                com.Azion.EloanUtility.UIUtility.alert("帳號(e-Mail)不存在，請先註冊為會員")
                com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "帳號(e-Mail)不存在，請先註冊為會員")
                Return
            End If
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me.Page, ex)
        End Try
    End Sub

    <System.Web.Services.WebMethod()> _
    Public Shared Function ValidINPUT(ByVal sNUMBER As String) As String
        Dim sNUMBERErr As String = String.Empty
        If HttpContext.Current.Session Is Nothing OrElse HttpContext.Current.Session("ValidateNumber") = Nothing OrElse HttpContext.Current.Session("ValidateNumber") = "" Then
            sNUMBERErr = "驗證碼逾時，請重新整理"
        Else
            If sNUMBER <> HttpContext.Current.Session("ValidateNumber") Then
                sNUMBERErr = "驗證碼錯誤"
            End If
        End If

        Return "{""NUMBER"":""" & sNUMBERErr & """}"
    End Function

    Private Sub btnReSetPass_Click(sender As Object, e As EventArgs) Handles btnReSetPass.Click
        Dim dbManager As com.Azion.NET.VB.DatabaseManager = com.Azion.NET.VB.DatabaseManager.getInstance
        Try
            Dim sMailTos() As String = {Trim(Me.txt_UserId.Text)}

            Dim sMailSub As String = String.Empty
            sMailSub = "MBSC會員重設密碼"

            Dim sMailBody As String = String.Empty
            Dim MB_ACCT As New MB_ACCT(dbManager)
            If MB_ACCT.LoadByPK(Trim(Me.txt_UserId.Text)) Then
                Dim sMB_PASSVID As String = Me.getPASSID(dbManager)

                Dim sb As New StringBuilder
                sb.Append("<P>")
                sb.Append("親愛的" & MB_ACCT.getString("MB_NAME") & "您好：<BR/>")
                sb.Append("這封信是由佛陀原始正法中心發送的。<BR/>")
                sb.Append("非常感謝您註冊成為我們的會員。<BR/>")
                'sb.Append("您的原密碼為<span style='color:red;font-weight:bold;font-size:14pt'>" & MB_ACCT.getString("MB_PSW") & "</span><BR/>")
                sb.Append("您需點擊下面的連結，啟動重設密碼程序：<BR/>")
                sb.Append("<a href='http://www.mbscnn.org/mnt/MBMnt_Reg_01_v01.aspx?MOD=RESET&UID=" & sMB_PASSVID & "'>http://www.mbscnn.org/mnt/MBMnt_Reg_01_v01.aspx?MOD=RESET&UID=" & sMB_PASSVID & "</a>")
                'sb.Append("<a href='http://localhost/mnt/MBMnt_Reg_01_v01.aspx?MOD=RESET&UID=" & sMB_PASSVID & "'>http://localhost/mnt/MBMnt_Reg_01_v01.aspx?MOD=RESET&UID=" & sMB_PASSVID & "</a>")
                sb.Append("<BR/>")
                sb.Append("<span style='color:red'>(如果上面不是連結形式，請將該位址貼到瀏覽器網址欄)</span>")
                sb.Append("<BR/>")
                sb.Append("感謝您的訪問，祝順心愉快！")
                sb.Append("<BR/>")
                sb.Append("此致")
                sb.Append("</P>")

                sMailBody = sb.ToString

                If com.Azion.EloanUtility.NetUtility.GMail_Send(sMailTos, Nothing, sMailSub, sMailBody, True, Nothing, False) Then
                    MB_ACCT.setAttribute("MB_PASSVID", sMB_PASSVID)
                    MB_ACCT.save()

                    Me.PLH_NUMBER.Visible = False

                    com.Azion.EloanUtility.UIUtility.alert("系統已發送『MBSC會員重設密碼』至您的e-Mail，請到您的信箱完成重設密碼程序【若找不到驗證信，請試試垃圾郵件資料匣】")
                    com.Azion.EloanUtility.UIUtility.showErrMsg(Me.Page, "系統已發送『MBSC會員重設密碼』至您的e-Mail，請到您的信箱完成重設密碼程序【若找不到驗證信，請試試垃圾郵件資料匣】")
                Else
                    com.Azion.EloanUtility.UIUtility.alert("『MBSC會員重設密碼』e-Mail發送失敗，請確認您的e-Mail網址是否正確或稍後再試")
                    com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "『MBSC會員重設密碼』e-Mail發送失敗，請確認您的e-Mail網址是否正確或稍後再試")
                End If
            Else
                com.Azion.EloanUtility.UIUtility.showErrMsg(Me.Page, "查無帳號【" & Trim(Me.txt_UserId.Text) & "】的會員資料")
            End If
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me.Page, ex)
        Finally
            dbManager.releaseConnection()
        End Try
    End Sub

    Function getPASSID(ByVal dbManager As com.Azion.NET.VB.DatabaseManager) As String
        Try
            Dim sProcName As String = String.Empty
            sProcName = LCase(com.Azion.NET.VB.Properties.getSchemaName) & ".MB_GET_MAXID"
            Dim inParaAL As New ArrayList
            Dim outParaAL As New ArrayList
            inParaAL.Add("01")
            inParaAL.Add("4")

            outParaAL.Add(7)

            Dim HT_Return As Hashtable = com.Azion.NET.VB.DBObject.getProcedureData(dbManager, sProcName, inParaAL, outParaAL)
            Dim iMAXID As Decimal = HT_Return.Item("@IMAXID")

            Return com.Azion.EloanUtility.StrUtility.FillZero(iMAXID, 7)
        Catch ex As Exception
            Throw
        Finally

        End Try
    End Function
End Class