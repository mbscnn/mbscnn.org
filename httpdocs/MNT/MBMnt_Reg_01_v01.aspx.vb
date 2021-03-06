﻿Imports System.Net.Sockets
Imports System.IO
Imports MBSC.MB_OP

Public Class MBMnt_Reg_01_v01
    Inherits System.Web.UI.Page

    Dim m_sMOD As String = String.Empty

    Dim m_sUID As String = String.Empty

    Dim m_sMAIL As String = String.Empty

    Dim m_sMB_MEMSEQ As String = String.Empty

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.m_sMOD = "" & Request.QueryString("MOD")

            Me.m_sUID = "" & Request.QueryString("UID")

            Me.m_sMAIL = "" & Request.QueryString("MAIL")

            Me.m_sMB_MEMSEQ = "" & Request.QueryString("MB_MEMSEQ")

            If Not Page.IsPostBack Then
                If Me.m_sMOD = "APV" AndAlso com.Azion.EloanUtility.ValidateUtility.isValidateData(m_sUID) Then
                    Me.PLH_MB_ACCT.Visible = False

                    Dim dbManage As com.Azion.NET.VB.DatabaseManager = com.Azion.NET.VB.DatabaseManager.getInstance
                    Try
                        Dim mbACCT As New MB_ACCT(dbManage)
                        If mbACCT.LoadByMB_APVID(Me.m_sUID) Then
                            mbACCT.setAttribute("MB_APV", "Y")
                            mbACCT.save()

                            Session("admin") = "1"

                            Session("USERID") = mbACCT.getString("MB_ACCT")

                            Me.HID_USERID.Value = mbACCT.getString("MB_ACCT")

                            Me.PLH_APV.Visible = True

                            Dim sMB_MEMSEQ_URL As String = String.Empty
                            Dim MB_MEMBERList As New MB_MEMBERList(dbManage)
                            MB_MEMBERList.loadByMB_EMAIL(mbACCT.getString("MB_ACCT"))
                            If MB_MEMBERList.getCurrentDataSet.Tables(0).Rows.Count = 0 Then
                                MB_MEMBERList.clear
                                Dim iCNT As Integer = 0
                                If Utility.isValidateData(mbACCT.getAttribute("MB_NAME")) AndAlso Utility.isValidateData(mbACCT.getAttribute("MB_MOBIL")) Then
                                    MB_MEMBERList.Load_MB_NAME_MB_MOBIL(mbACCT.getAttribute("MB_NAME"), mbACCT.getAttribute("MB_MOBIL"))
                                    If MB_MEMBERList.getCurrentDataSet.Tables(0).Rows.Count > 0 Then
                                        'com.Azion.EloanUtility.UIUtility.showJSMsg("【" & Me.MB_NAME.Text.Trim & "】【" & Me.MB_MOBIL.Text.Trim & "】已經是會員了，請再確認看看")
                                        'com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "【" & Me.MB_NAME.Text.Trim & "】【" & Me.MB_MOBIL.Text.Trim & "】已經是會員了，請再確認看看")

                                        'Return

                                        iCNT+=1
                                    End If
                                End If

                                If Utility.isValidateData(mbACCT.getAttribute("MB_NAME")) AndAlso Utility.isValidateData(mbACCT.getAttribute("MB_TEL")) Then
                                    'Dim sMB_TEL As String = String.Empty
                                    'If Utility.isValidateData(Trim(Me.MB_TEL_Pre.Text)) Then
                                    '    sMB_TEL = Trim(Me.MB_TEL_Pre.Text) & "-" & Trim(Me.MB_TEL.Text)
                                    'Else
                                    '    sMB_TEL = Trim(Me.MB_TEL.Text)
                                    'End If

                                    MB_MEMBERList.clear()
                                    MB_MEMBERList.Load_MB_NAME_MB_TEL(mbACCT.getAttribute("MB_NAME"), mbACCT.getAttribute("MB_TEL"))
                                    If MB_MEMBERList.getCurrentDataSet.Tables(0).Rows.Count > 0 Then
                                        'com.Azion.EloanUtility.UIUtility.showJSMsg("【" & Me.MB_NAME.Text.Trim & "】【" & sMB_TEL & "】已經是會員了，請再確認看看")
                                        'com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "【" & Me.MB_NAME.Text.Trim & "】【" & sMB_TEL & "】已經是會員了，請再確認看看")

                                        'Return

                                        iCNT+=1
                                    End If
                                End If

                                If Utility.isValidateData(mbACCT.getAttribute("MB_NAME")) AndAlso Utility.isValidateData(mbACCT.getAttribute("MB_ACCT")) Then
                                    MB_MEMBERList.clear()
                                    MB_MEMBERList.Load_MB_NAME_MB_EMAIL(mbACCT.getAttribute("MB_NAME"), mbACCT.getAttribute("MB_ACCT"))
                                    If MB_MEMBERList.getCurrentDataSet.Tables(0).Rows.Count > 0 Then
                                        'com.Azion.EloanUtility.UIUtility.showJSMsg("【" & Me.MB_NAME.Text.Trim & "】【" & Me.MB_EMAIL.Text.Trim & "】已經是會員了，請再確認看看")
                                        'com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "【" & Me.MB_NAME.Text.Trim & "】【" & Me.MB_EMAIL.Text.Trim & "】已經是會員了，請再確認看看")

                                        'Return

                                        iCNT+=1
                                    End If
                                End If

                                If iCNT = 0 Then
                                    Dim sProcName As String = String.Empty
                                    sProcName = LCase(com.Azion.NET.VB.Properties.getSchemaName) & ".MB_GET_MAXID"
                                    Dim inParaAL As New ArrayList
                                    Dim outParaAL As New ArrayList
                                    inParaAL.Add("01")
                                    inParaAL.Add("1")

                                    outParaAL.Add(7)

                                    Dim HT_Return As Hashtable = com.Azion.NET.VB.DBObject.getProcedureData(dbManage, sProcName, inParaAL, outParaAL)
                                    Dim iMAXID As Decimal = HT_Return.Item("@IMAXID")

                                    Dim sMB_MEMSEQ As String = String.Empty
                                    sMB_MEMSEQ = com.Azion.EloanUtility.StrUtility.FillZero(iMAXID, 7)
                                    Dim MB_MEMBER As New MB_MEMBER(dbManage)
                                    MB_MEMBER.loadByPK(CDec(sMB_MEMSEQ))
                                    'MB_MEMSEQ	decimal(7,0)		NO	PRI	0		select,insert,update,references	會員編號
                                    MB_MEMBER.setAttribute("MB_MEMSEQ",CDec(sMB_MEMSEQ))
                                    'MB_NAME	varchar(50)	utf8_general_ci	NO	MUL			select,insert,update,references	姓名
                                    MB_MEMBER.setAttribute("MB_NAME",mbACCT.getAttribute("MB_NAME"))
                                    'MB_SEX	char(1)	latin1_swedish_ci	YES				select,insert,update,references	性別 1:男 2:女
                                    MB_MEMBER.setAttribute("MB_SEX",mbACCT.getAttribute("MB_SEX"))
                                    'MB_MOBIL	varchar(20)	utf8_general_ci	YES	MUL			select,insert,update,references	手機
                                    MB_MEMBER.setAttribute("MB_MOBIL",mbACCT.getAttribute("MB_MOBIL"))
                                    'MB_TEL	varchar(40)	utf8_general_ci	YES	MUL			select,insert,update,references	電話
                                    MB_MEMBER.setAttribute("MB_TEL",mbACCT.getAttribute("MB_TEL"))
                                    'MB_EMAIL	varchar(40)	utf8_general_ci	YES				select,insert,update,references	e-mail
                                    MB_MEMBER.setAttribute("MB_EMAIL",mbACCT.getAttribute("MB_ACCT"))
                                    'CHGUID	varchar(100)	utf8_general_ci	YES				select,insert,update,references	修改員工編號
                                    MB_MEMBER.setAttribute("CHGUID",mbACCT.getAttribute("MB_ACCT"))
                                    'CHGDATE	datetime		NO				select,insert,update,references	修改日期
                                    MB_MEMBER.setAttribute("CHGDATE",Now)
                                    MB_MEMBER.save

                                    'sMB_MEMSEQ_URL = "OPTYPE=A&MB_MEMSEQ=" & sMB_MEMSEQ
                                    sMB_MEMSEQ_URL = "MB_MEMSEQ=" & sMB_MEMSEQ
                                End If
                            Else
                                If Utility.isValidateData(mbACCT.getAttribute("MB_MEMSEQ")) Then
                                    Dim MB_MEMBER As New MB_MEMBER(dbManage)
                                    If MB_MEMBER.loadByPK(mbACCT.getAttribute("MB_MEMSEQ")) Then
                                        'MB_EMAIL	varchar(40)	utf8_general_ci	YES				select,insert,update,references	e-mail
                                        MB_MEMBER.setAttribute("MB_EMAIL",mbACCT.getAttribute("MB_ACCT"))
                                        'MB_MOBIL	varchar(20)	utf8_general_ci	YES	MUL			select,insert,update,references	手機
                                        MB_MEMBER.setAttribute("MB_MOBIL",mbACCT.getAttribute("MB_MOBIL"))
                                        'MB_TEL	varchar(40)	utf8_general_ci	YES	MUL			select,insert,update,references	電話
                                        MB_MEMBER.setAttribute("MB_TEL",mbACCT.getAttribute("MB_TEL"))
                                        'MB_SEX	char(1)	latin1_swedish_ci	NO				select,insert,update,references	性別
                                        MB_MEMBER.setAttribute("MB_SEX",mbACCT.getAttribute("MB_SEX"))
                                        'MB_NAME	varchar(20)	utf8_general_ci	NO				select,insert,update,references	姓名
                                        MB_MEMBER.setAttribute("MB_NAME",mbACCT.getAttribute("MB_NAME"))
                                        'CHGUID	varchar(100)	utf8_general_ci	YES				select,insert,update,references	修改員工編號
                                        MB_MEMBER.setAttribute("CHGUID",mbACCT.getAttribute("MB_ACCT"))
                                        'CHGDATE	datetime		NO				select,insert,update,references	修改日期
                                        MB_MEMBER.setAttribute("CHGDATE",Now)
                                        MB_MEMBER.save

                                        'sMB_MEMSEQ_URL = "OPTYPE=A&MB_MEMSEQ=" & mbACCT.getAttribute("MB_MEMSEQ")
                                        sMB_MEMSEQ_URL = "MB_MEMSEQ=" & mbACCT.getAttribute("MB_MEMSEQ")
                                    End If
                                End If
                            End If

                            '直接導頁到入會申請單
                            'Dim sURL As String = String.Empty
                            'sURL = com.Azion.EloanUtility.UIUtility.getRootPath &
                            '       "/MNT/MBMnt_Member_01_v01.aspx.aspx"

                            LTL_SCRIPT.Text = "<script language='javascript' >" & vbCrLf
                            LTL_SCRIPT.Text &= "alert('e-Mail驗證成功\n\n請記得填寫入會申請單\n\n【按確定後系統將連結入會申請單】');" & vbCrLf
                            If Utility.isValidateData(sMB_MEMSEQ_URL) Then
                                LTL_SCRIPT.Text &= "window.location.href='MBMnt_Member_01_v01.aspx?" & sMB_MEMSEQ_URL & "';" & vbCrLf
                            Else
                                LTL_SCRIPT.Text &= "window.location.href='MBMnt_Member_01_v01.aspx';" & vbCrLf
                            End If
                            
                            LTL_SCRIPT.Text &= "</" & "script>"

                            LTL_SCRIPT.Visible = True
                        Else
                            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "e-Mail驗證失敗，無法註冊為會員")
                        End If
                    Catch ex As System.Threading.ThreadAbortException
                    Catch ex As Exception
                        Throw
                    Finally
                        dbManage.releaseConnection()
                    End Try
                ElseIf Me.m_sMOD = "RESET" AndAlso com.Azion.EloanUtility.ValidateUtility.isValidateData(m_sUID) Then
                    Me.TXT_EMAIL.Attributes.Add("Readonly", "True")

                    Me.PLH_Send.Visible = False

                    Me.LTL_TITLE.Text = "重設密碼"

                    Dim dbManage As com.Azion.NET.VB.DatabaseManager = com.Azion.NET.VB.DatabaseManager.getInstance
                    Try
                        Dim mbACCT As New MB_ACCT(dbManage)
                        If mbACCT.LoadByMB_PASSVID(Me.m_sUID) Then
                            Me.HID_USERID.Value = mbACCT.getString("MB_ACCT")
                            Me.TXT_EMAIL.Text = mbACCT.getString("MB_ACCT")

                            Me.TXT_APPNAME.Text = mbACCT.getString("MB_NAME")

                            Me.RBL_MB_SEX.SelectedIndex = -1
                            If Not IsNothing(Me.RBL_MB_SEX.Items.FindByValue(mbACCT.getString("MB_SEX"))) Then
                                Me.RBL_MB_SEX.Items.FindByValue(mbACCT.getString("MB_SEX")).Selected = True
                            End If

                            Me.PLH_REPASS.Visible = True
                        Else
                            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "查無會員資料，無法重設密碼")
                        End If
                    Catch ex As Exception
                        Throw
                    Finally
                        dbManage.releaseConnection()
                    End Try
                Else
                    If IsNumeric(Me.m_sMB_MEMSEQ) Then
                        Using dbManager As com.Azion.NET.VB.DatabaseManager = UICtl.UIShareFun.getDataBaseManager
                            Dim MB_MEMBER As New MB_MEMBER(dbManager)
                            If MB_MEMBER.loadByPK(Me.m_sMB_MEMSEQ) Then
                                Me.TXT_EMAIL.Text = MB_MEMBER.getString("MB_EMAIL")
                                Me.TXT_APPNAME.Text = MB_MEMBER.getString("MB_NAME")
                                Me.RBL_MB_SEX.SelectedIndex = -1
                                If MB_MEMBER.getString("MB_SEX") = "1" Then
                                    Me.RBL_MB_SEX.Items.FindByValue("M").Selected = True
                                ElseIf MB_MEMBER.getString("MB_SEX") = "2" Then
                                    Me.RBL_MB_SEX.Items.FindByValue("F").Selected = True
                                End If
                                Me.MB_MOBIL.Text = MB_MEMBER.getString("MB_MOBIL")
                                If Utility.isValidateData(MB_MEMBER.getAttribute("MB_TEL")) Then
                                    Dim sTEL() as String = nothing
                                    sTEL = Split(MB_MEMBER.getAttribute("MB_TEL"),"-")
                                    If Not IsNothing(sTEL) AndAlso sTEL.Length>=2 Then
                                        Me.MB_TEL_Pre.Text = sTEL(0)
                                        Me.MB_TEL.Text = sTEL(1)
                                    Elseif Not IsNothing(sTEL) AndAlso sTEL.Length=1 then
                                        Me.MB_TEL.Text = sTEL(0)
                                    End If 
                                End If                                
                            End If
                        End Using
                    End If

                    Me.IMG_Vad.Src = com.Azion.EloanUtility.UIUtility.getRootPath & "/Module/ValidateNumber.ashx"

                    '點圖片重新整理
                    Me.IMG_Vad.Attributes("onclick") = "this.src='" & _
                                                       com.Azion.EloanUtility.UIUtility.getRootPath & "/Module/ValidateNumber.ashx?" & _
                                                      "'+Math.random();"
                End If
            End If
        Catch ex As System.Threading.ThreadAbortException
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

    <System.Web.Services.WebMethod()> _
    Public Shared Function ValidINPUT(ByVal sNUMBER As String, ByVal sEMAIL As String, ByVal sPASSWORD As String, ByVal sVARPASSWORD As String, BYval sMB_MOBIL As String,Byval sMB_TEL_Pre As String,Byval sMB_TEL As string) As String
        Try
            Dim sNUMBERErr As String = String.Empty
            If HttpContext.Current.Session Is Nothing OrElse HttpContext.Current.Session("ValidateNumber") = Nothing OrElse HttpContext.Current.Session("ValidateNumber") = "" Then
                sNUMBERErr = "驗證碼逾時，請重新整理"
            Else
                If sNUMBER <> HttpContext.Current.Session("ValidateNumber") Then
                    sNUMBERErr = "驗證碼錯誤"
                End If
            End If

            Dim sEMAILErr As String = String.Empty
            If Not com.Azion.EloanUtility.ValidateUtility.isEmail(Trim(sEMAIL)) Then
                sEMAILErr = "EMAIL錯誤"
            Else
                Dim dbManager As com.Azion.NET.VB.DatabaseManager = com.Azion.NET.VB.DatabaseManager.getInstance
                Try
                    Dim mbACCT As New MB_ACCT(dbManager)
                    If mbACCT.LoadByPK(Trim(sEMAIL)) Then
                        If mbACCT.getString("MB_APV") = "Y" Then
                            sEMAILErr = sEMAIL & "已於" & com.Azion.EloanUtility.StrUtility.ToYYYMMDD(mbACCT.getAttribute("MB_CRE_DATE")) &
                                        "註冊啟用為會員，請直接用會員登入"
                        Else
                            sEMAILErr = sEMAIL & "已於" & com.Azion.EloanUtility.StrUtility.ToYYYMMDD(mbACCT.getAttribute("MB_CRE_DATE")) &
                                        "註冊為會員，但尚未驗證啟用"
                        End If

                        If Utility.isValidateData(sMB_MOBIL) OrElse Utility.isValidateData(sMB_TEL) Then
                            If Utility.isValidateData(sMB_MOBIL) Then
                                mbACCT.setAttribute("MB_MOBIL", sMB_MOBIL)
                            End If

                            If Utility.isValidateData(sMB_TEL) Then
                                mbACCT.setAttribute("MB_TEL", sMB_TEL)
                            End If

                            mbACCT.save()
                        End If
                    Else
                        If Utility.isValidateData(sMB_MOBIL) Then
                            Dim MB_ACCTList As New MB_ACCTList(dbManager)
                            MB_ACCTList.Load_MB_MOBIL(sMB_MOBIL)
                            If MB_ACCTList.getCurrentDataSet.Tables(0).Rows.Count > 0 Then
                                For Each ROW As DataRow In MB_ACCTList.getCurrentDataSet.Tables(0).Rows
                                    sEMAILErr &= ROW("MB_ACCT").ToString & "，"
                                Next
                                If Utility.isValidateData(sEMAILErr) Then
                                    sEMAILErr = Left(sEMAILErr, sEMAILErr.Length - 1)
                                End If
                                sEMAILErr &= "已註冊成為會員,請直接用會員登入"
                            End If
                        ElseIf Utility.isValidateData(sMB_TEL) Then
                            If Utility.isValidateData(sMB_TEL_Pre) Then
                                sMB_TEL = sMB_TEL_Pre & "-" & sMB_TEL
                            End If

                            Dim MB_ACCTList As New MB_ACCTList(dbManager)
                            MB_ACCTList.Load_MB_TEL(sMB_TEL)
                            If MB_ACCTList.getCurrentDataSet.Tables(0).Rows.Count > 0 Then
                                For Each ROW As DataRow In MB_ACCTList.getCurrentDataSet.Tables(0).Rows
                                    sEMAILErr &= ROW("MB_ACCT").ToString & "，"
                                Next
                                If Utility.isValidateData(sEMAILErr) Then
                                    sEMAILErr = Left(sEMAILErr, sEMAILErr.Length - 1)
                                End If
                                sEMAILErr &= "已註冊成為會員,請直接用會員登入"
                            End If
                        End If
                    End If
                Catch ex As Exception
                    Throw
                Finally
                    dbManager.releaseConnection()
                End Try
            End If

            Dim sPASSVadErr As String = String.Empty
            If sPASSWORD <> sVARPASSWORD Then
                sPASSVadErr = "密碼與確認密碼不符合"
            End If

            If Trim(sPASSWORD).Length < 8 OrElse Trim(sPASSWORD).Length > 15 Then
                sPASSVadErr = "密碼最多15碼最少8碼"
            End If

            Return "{""NUMBER"":""" & sNUMBERErr & """,""EMAIL"":""" & sEMAILErr & """,""PASSWORD"":""" & sPASSVadErr & """}"
        Catch ex As Exception
            Throw
        End Try
    End Function

    <System.Web.Services.WebMethod()> _
    Public Shared Function RE_ValidINPUT(ByVal sNUMBER As String, ByVal sEMAIL As String, ByVal sPASSWORD As String, ByVal sVARPASSWORD As String) As String
        Try
            Dim sNUMBERErr As String = String.Empty
            If HttpContext.Current.Session Is Nothing OrElse HttpContext.Current.Session("ValidateNumber") = Nothing OrElse HttpContext.Current.Session("ValidateNumber") = "" Then
                sNUMBERErr = "驗證碼逾時，請重新整理"
            Else
                If sNUMBER <> HttpContext.Current.Session("ValidateNumber") Then
                    sNUMBERErr = "驗證碼錯誤"
                End If
            End If

            Dim sEMAILErr As String = String.Empty
            If Not com.Azion.EloanUtility.ValidateUtility.isEmail(Trim(sEMAIL)) Then
                sEMAILErr = "EMAIL錯誤"
            Else
                Dim dbManager As com.Azion.NET.VB.DatabaseManager = com.Azion.NET.VB.DatabaseManager.getInstance
                Try
                    Dim mbACCT As New MB_ACCT(dbManager)
                    If mbACCT.LoadByPK(Trim(sEMAIL)) Then
                        If mbACCT.getString("MB_APV") = "Y" Then
                            'sEMAILErr = sEMAIL & "已於" & com.Azion.EloanUtility.StrUtility.ToYYYMMDD(mbACCT.getAttribute("MB_CRE_DATE")) & _
                            '            "註冊啟用為會員"
                        Else
                            sEMAILErr = sEMAIL & "已於" & com.Azion.EloanUtility.StrUtility.ToYYYMMDD(mbACCT.getAttribute("MB_CRE_DATE")) & _
                                        "註冊為會員，但尚未驗證啟用"
                        End If
                    Else
                        sEMAILErr = sEMAIL & "尚未註冊為會員"
                    End If
                Catch ex As Exception
                    Throw
                Finally
                    dbManager.releaseConnection()
                End Try
            End If

            Dim sPASSVadErr As String = String.Empty
            If sPASSWORD <> sVARPASSWORD Then
                sPASSVadErr = "密碼與確認密碼不符合"
            End If

            If Trim(sPASSWORD).Length < 8 OrElse Trim(sPASSWORD).Length > 15 Then
                sPASSVadErr = "密碼最多15碼最少8碼"
            End If

            Return "{""NUMBER"":""" & sNUMBERErr & """,""EMAIL"":""" & sEMAILErr & """,""PASSWORD"":""" & sPASSVadErr & """}"
        Catch ex As Exception
            Throw
        End Try
    End Function

    Private Sub btSend_Click(sender As Object, e As EventArgs) Handles btSend.Click
        Dim dbManager As com.Azion.NET.VB.DatabaseManager = com.Azion.NET.VB.DatabaseManager.getInstance
        Try
            Dim sMailTos() As String = {Trim(Me.TXT_EMAIL.Text)}

            Dim sMailSub As String = String.Empty
            sMailSub = com.Azion.EloanUtility.FileUtility.getAppSettings("MAILSUB") & Now.ToShortDateString & " " & Now.ToShortTimeString

            Dim sMB_APVID As String = Me.getVadID(dbManager)

            Dim sMailBody As String = String.Empty

            sMailBody = MBSC.UICtl.UIShareFun.getMailBody(sMB_APVID, Trim(Me.TXT_EMAIL.Text), Trim(Me.TXT_APPNAME.Text),Me.m_sMB_MEMSEQ)

            If com.Azion.EloanUtility.NetUtility.GMail_Send(sMailTos, Nothing, sMailSub, sMailBody, True, Nothing, False) Then
                Dim mbACCT As New MB_ACCT(dbManager)
                mbACCT.LoadByPK(Trim(Me.TXT_EMAIL.Text))

                'MB_ACCT varchar(100) 帳號(Mail)
                mbACCT.setAttribute("MB_ACCT", Trim(Me.TXT_EMAIL.Text))

                'MB_PSW varchar(15) 密碼
                mbACCT.setAttribute("MB_PSW", Trim(Me.TXT_PASSWORD.Text))

                'MB_NAME varchar(20) 姓名
                mbACCT.setAttribute("MB_NAME", Trim(Me.TXT_APPNAME.Text))

                'MB_SEX char(1) 性別
                mbACCT.setAttribute("MB_SEX", Trim(Me.RBL_MB_SEX.SelectedValue))

                'MB_CRE_DATE datetime 註冊日
                mbACCT.setAttribute("MB_CRE_DATE", Now)

                'MB_IDTIFY char(1) 身分 1.管理者;2.一般
                mbACCT.setAttribute("MB_IDTIFY", "2")

                'MB_APVID varchar(7) 驗證用ID
                mbACCT.setAttribute("MB_APVID", sMB_APVID)

                'MB_MOBIL	varchar(20)	utf8_general_ci	YES				select,insert,update,references	手機
                If Utility.isValidateData(Me.MB_MOBIL.text) Then
                    mbACCT.setAttribute("MB_MOBIL", Me.MB_MOBIL.text)
                Else
                    mbACCT.setAttribute("MB_MOBIL", Nothing)
                End If

                'MB_TEL	varchar(40)	utf8_general_ci	YES				select,insert,update,references	電話
                If Utility.isValidateData(Me.MB_TEL_Pre.text) AndAlso Utility.isValidateData(Me.MB_TEL.text) Then
                    mbACCT.setAttribute("MB_TEL",Me.MB_TEL_Pre.Text & "-" & Me.MB_TEL.text)
                ElseIf Not Utility.isValidateData(Me.MB_TEL_Pre.text) AndAlso Utility.isValidateData(Me.MB_TEL.text) Then
                     mbACCT.setAttribute("MB_TEL", Me.MB_TEL.text)
                Else
                     mbACCT.setAttribute("MB_TEL", Nothing)
                End If

                'MB_MEMSEQ	decimal(7,0)		YES				select,insert,update,references	會員編號
                If IsNumeric(Me.m_sMB_MEMSEQ) Then
                    mbACCT.setAttribute("MB_MEMSEQ", CDec(Me.m_sMB_MEMSEQ))
                Else
                    mbACCT.setAttribute("MB_MEMSEQ", nothing)
                End If

                mbACCT.save()

                Dim MB_MEMBERList As New MB_MEMBERList(dbManager)
                MB_MEMBERList.Load_EXISTS(mbACCT.getAttribute("MB_NAME"),mbACCT.getAttribute("MB_ACCT"),mbACCT.getAttribute("MB_MOBIL"),mbACCT.getAttribute("MB_TEL"))
                If MB_MEMBERList.getCurrentDataSet.Tables(0).Rows.Count>0 Then
                    For i As Integer = 0 To MB_MEMBERList.size-1
                        Dim MB_MEMBER As MB_MEMBER = CType(MB_MEMBERList.item(i),MB_MEMBER)
                        'MB_NAME	varchar(50)	utf8_general_ci	NO	MUL			select,insert,update,references	姓名
                        MB_MEMBER.setAttribute("MB_NAME",mbACCT.getAttribute("MB_NAME"))
                        'MB_SEX	char(1)	latin1_swedish_ci	YES				select,insert,update,references	性別 1:男 2:女
                        MB_MEMBER.setAttribute("MB_SEX",mbACCT.getAttribute("MB_SEX"))
                        'MB_MOBIL	varchar(20)	utf8_general_ci	YES	MUL			select,insert,update,references	手機
                        MB_MEMBER.setAttribute("MB_MOBIL",mbACCT.getAttribute("MB_MOBIL"))
                        'MB_TEL	varchar(40)	utf8_general_ci	YES	MUL			select,insert,update,references	電話
                        MB_MEMBER.setAttribute("MB_TEL",mbACCT.getAttribute("MB_TEL"))
                        'MB_EMAIL	varchar(40)	utf8_general_ci	YES				select,insert,update,references	e-mail
                        MB_MEMBER.setAttribute("MB_EMAIL",mbACCT.getAttribute("MB_ACCT"))
                        MB_MEMBER.save
                        MB_MEMBER.save
                    Next
                Else
                    Dim sProcName As String = String.Empty
                    sProcName = LCase(com.Azion.NET.VB.Properties.getSchemaName) & ".MB_GET_MAXID"
                    Dim inParaAL As New ArrayList
                    Dim outParaAL As New ArrayList
                    inParaAL.Add("01")
                    inParaAL.Add("1")

                    outParaAL.Add(7)

                    Dim HT_Return As Hashtable = com.Azion.NET.VB.DBObject.getProcedureData(dbManager, sProcName, inParaAL, outParaAL)
                    Dim iMAXID As Decimal = HT_Return.Item("@IMAXID")

                    Dim sMB_MEMSEQ As String = String.Empty
                    sMB_MEMSEQ = com.Azion.EloanUtility.StrUtility.FillZero(iMAXID, 7)

                    Dim MB_MEMBER As New MB_MEMBER(dbManager)
                    MB_MEMBER.loadByPK(CDEC(sMB_MEMSEQ))
                    MB_MEMBER.setAttribute("MB_MEMSEQ",CDec(sMB_MEMSEQ))
                    'MB_NAME	varchar(50)	utf8_general_ci	NO	MUL			select,insert,update,references	姓名
                    MB_MEMBER.setAttribute("MB_NAME",mbACCT.getAttribute("MB_NAME"))
                    'MB_SEX	char(1)	latin1_swedish_ci	YES				select,insert,update,references	性別 1:男 2:女
                    MB_MEMBER.setAttribute("MB_SEX",mbACCT.getAttribute("MB_SEX"))
                    'MB_MOBIL	varchar(20)	utf8_general_ci	YES	MUL			select,insert,update,references	手機
                    MB_MEMBER.setAttribute("MB_MOBIL",mbACCT.getAttribute("MB_MOBIL"))
                    'MB_TEL	varchar(40)	utf8_general_ci	YES	MUL			select,insert,update,references	電話
                    MB_MEMBER.setAttribute("MB_TEL",mbACCT.getAttribute("MB_TEL"))
                    'MB_EMAIL	varchar(40)	utf8_general_ci	YES				select,insert,update,references	e-mail
                    MB_MEMBER.setAttribute("MB_EMAIL",mbACCT.getAttribute("MB_ACCT"))
                    MB_MEMBER.save
                End If

                Me.PLH_MAILSend.Visible = True
                Me.PLH_MB_ACCT.Visible = False
            Else
                com.Azion.EloanUtility.UIUtility.alert("驗證信發送失敗，請確認您的e-Mail網址是否正確或稍後再試")
                com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "驗證信發送失敗，請確認您的e-Mail網址是否正確或稍後再試")
            End If
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        Finally
            dbManager.releaseConnection()
        End Try
    End Sub

    Function getVadID(ByVal dbManager As com.Azion.NET.VB.DatabaseManager) As String
        Try
            Dim sProcName As String = String.Empty
            sProcName = LCase(com.Azion.NET.VB.Properties.getSchemaName) & ".MB_GET_MAXID"
            Dim inParaAL As New ArrayList
            Dim outParaAL As New ArrayList
            inParaAL.Add("01")
            inParaAL.Add("2")

            outParaAL.Add(7)

            Dim HT_Return As Hashtable = com.Azion.NET.VB.DBObject.getProcedureData(dbManager, sProcName, inParaAL, outParaAL)
            Dim iMAXID As Decimal = HT_Return.Item("@IMAXID")

            Return com.Azion.EloanUtility.StrUtility.FillZero(iMAXID, 7)
        Catch ex As Exception
            Throw
        Finally

        End Try
    End Function

    '#Region "CheckEmail"
    '    '1表示郵寄地址合法，2表示只有功能變數名稱正確（可能是用戶的郵件帳戶無效），3表示出現了錯誤，4表示無法找到郵件伺服器。
    '    Public Function CheckEmail(ByVal sEmail As String) As Long

    '        Dim oStream As NetworkStream = Nothing

    '        '發件人
    '        Dim sFrom As String = String.Empty

    '        '收件人
    '        Dim sTo As String = String.Empty

    '        '郵件伺服器的應答
    '        Dim sResponse As String = String.Empty

    '        '發件人的域名
    '        Dim Remote_Addr As String = String.Empty

    '        '郵件伺服器
    '        Dim mserver As String = String.Empty

    '        Dim sText As String() = Nothing

    '        sTo = "<" + sEmail + ">"

    '        ' 從郵件地址分離出帳戶名和域名
    '        sText = sEmail.Split(CType("@", Char))

    '        '查找該域的郵件伺服器
    '        mserver = GetMailServer(sText(1))

    '        'mserver為空值表明查找郵件伺服器失敗
    '        If mserver = "" Then
    '            Return 4
    '            Exit Function
    '        End If

    '        '發件人地址的域名必須合法
    '        Remote_Addr = "gmail.com"

    '        '盡可能延遲創建對象的時間"
    '        'sFrom = "<" + com.Azion.EloanUtility.FileUtility.getAppSettings("MAILFROM") + ">"
    '        sFrom = "　　" '盡可能延遲創建對象的時間

    '        Dim oConnection As New TcpClient()
    '        Try
    '            '超時時間
    '            oConnection.SendTimeout = 3000

    '            '連接SMTP埠
    '            oConnection.Connect(mserver, 25)

    '            '收集郵件伺服器的應答信息
    '            oStream = oConnection.GetStream()
    '            sResponse = GetData(oStream)
    '            sResponse = SendData(oStream, "HELO " & Remote_Addr & vbCrLf)
    '            'sResponse = SendData(oStream, "MAIL FROM: " & sFrom & vbCrLf)
    '            '如果對MAIL FROM指令有肯定的應答，
    '            '至少表明郵件地址的域名正确
    '            If ValidResponse(sResponse) Then
    '                'MAIL FROM不檢核
    '                '因可能被HOTMAIL BLOCK
    '                sResponse = SendData(oStream, "MAIL FROM: " & sFrom & vbCrLf)

    '                sResponse = SendData(oStream, "RCPT TO: " & sTo & vbCrLf)
    '                '如果對RCPT TO指令有肯定的應答
    '                '表明郵件伺服器已認可該地址
    '                If ValidResponse(sResponse) Then
    '                    Return 1 '郵件地址有效
    '                Else
    '                    Return 2 '隻有域名有效
    '                End If
    '            End If
    '            '結束與郵件伺服器的會話
    '            SendData(oStream, "QUIT" & vbCrLf)

    '            oConnection.Close()
    '            oStream = Nothing
    '        Catch
    '            Return 3 '錯誤！
    '        End Try
    '    End Function

    '    Private Function GetMailServer(ByVal sDomain As String) As String
    '        Dim info As New ProcessStartInfo()
    '        Dim ns As Process
    '        '调用Windows的nslookup命令，查找邮件服务器
    '        info.UseShellExecute = False
    '        info.RedirectStandardInput = True
    '        info.RedirectStandardOutput = True
    '        info.FileName = "nslookup"
    '        info.CreateNoWindow = True
    '        '查找类型为MX。关于nslookup的详细说明，请参见
    '        'Windows帮助
    '        info.Arguments = "-type=MX " + sDomain.ToUpper.Trim
    '        '启动一个进行执行Windows的nslookup命令()
    '        ns = Process.Start(info)
    '        Dim sout As StreamReader
    '        sout = ns.StandardOutput
    '        ' 利用正则表达式找出nslookup命令输出结果中的邮件服务器信息
    '        'Dim reg As Regex = New Regex("mail exchanger = (?[^\\\s]+)")
    '        'Dim reg As Regex = New Regex("mail exchanger = (?<mailServer>[^\\s]+)")
    '        'Dim reg As Regex = New Regex("mail exchanger = (?[^\s]+)")
    '        Dim reg As Regex = New Regex("mail exchanger = (?<mailServer>[^\s]+)")

    '        Dim mailserver As String = String.Empty
    '        Dim response As String = String.Empty
    '        Do While (sout.Peek() > -1)
    '            response = sout.ReadLine()
    '            Dim amatch As Match = reg.Match(response)
    '            If (amatch.Success) Then
    '                mailserver = amatch.Groups("mailServer").Value
    '                Exit Do
    '            End If
    '        Loop
    '        Return mailserver
    '    End Function

    '    '獲取伺服器應答數據，并将其轉換為String
    '    Private Function GetData(ByRef oStream As NetworkStream) As String
    '        Dim bResponse(1024) As Byte
    '        Dim sResponse As String = String.Empty

    '        Dim lenStream As Integer = oStream.Read(bResponse, 0, 1024)
    '        If lenStream > 0 Then
    '            sResponse = Encoding.ASCII.GetString(bResponse, 0, 1024)
    '        End If
    '        Return sResponse
    '    End Function

    '    '向郵件伺服器發送數據
    '    Private Function SendData(ByRef oStream As NetworkStream, ByVal sToSend As String) As String
    '        Dim sResponse As String
    '        '将String轉換成Byte數組
    '        Dim bArray() As Byte = Encoding.ASCII.GetBytes(sToSend.ToCharArray)
    '        '發送數據
    '        oStream.Write(bArray, 0, bArray.Length())
    '        sResponse = GetData(oStream)
    '        '返回應答
    '        Return sResponse
    '    End Function

    '    '伺服器是否返回肯定的回答？
    '    Private Function ValidResponse(ByVal sResult As String) As Boolean
    '        Dim bResult As Boolean
    '        Dim iFirst As Integer
    '        If sResult.Length > 1 Then
    '            iFirst = CType(sResult.Substring(0, 1), Integer)
    '            '如果伺服器返回應答的第一個字元小于'3'
    '            '我們認為伺服器已認可剛才的操作
    '            If iFirst < 3 Then bResult = True
    '        End If
    '        Return bResult
    '    End Function
    '#End Region

    Private Sub bntRePASS_Click(sender As Object, e As EventArgs) Handles bntRePASS.Click
        Dim dbManager As com.Azion.NET.VB.DatabaseManager = com.Azion.NET.VB.DatabaseManager.getInstance
        Try
            Dim MB_ACCT As New MB_ACCT(dbManager)
            If MB_ACCT.LoadByPK(Me.TXT_EMAIL.Text) Then
                'MB_PSW varchar(15) 密碼
                MB_ACCT.setAttribute("MB_PSW", Trim(Me.TXT_PASSWORD.Text))

                'MB_NAME varchar(20) 姓名
                MB_ACCT.setAttribute("MB_NAME", Trim(Me.TXT_APPNAME.Text))

                'MB_SEX char(1) 性別
                MB_ACCT.setAttribute("MB_SEX", Trim(Me.RBL_MB_SEX.SelectedValue))

                MB_ACCT.save()

                com.Azion.EloanUtility.UIUtility.alert("密碼已重設，請使用新密碼做會員登入")

                Dim sURL As String = String.Empty
                'sURL = com.Azion.EloanUtility.UIUtility.getRootPath & "/MNT/MBMnt_Sign_01_v01.aspx"
                sURL = com.Azion.EloanUtility.UIUtility.getRootPath & "/NewsList.aspx"

                Response.Redirect(sURL)
            Else
                com.Azion.EloanUtility.UIUtility.alert("查無會員資料，無法重設密碼")
                com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "查無會員資料，無法重設密碼")
            End If
        Catch ex As System.Threading.ThreadAbortException
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        Finally
            dbManager.releaseConnection()
        End Try
    End Sub
End Class