Imports MBSC.MB_OP
Imports MBSC.UICtl

Public Class MBSignIn_01_v01
    Inherits System.Web.UI.Page

    Dim m_DBManager As com.Azion.NET.VB.DatabaseManager

    Dim m_iUPCODE_Admin As Integer = 76


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.m_DBManager = UIShareFun.getDataBaseManager
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

End Class