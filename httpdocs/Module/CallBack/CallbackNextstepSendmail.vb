Imports FLOW_OP
Imports com.Azion.EloanUtility
Imports com.Azion.NET.VB
Imports System.Net.Mail
Imports System.Net
Imports System.Threading
Imports System.IO



Public Class CallbackNextstepSendmail
    Implements IFlowSentNotification

    '流程送出後送EMAIL至USER

    Public Shared m_SYNC As New Object

    Public Overridable Sub Notify(ByVal dbManager As com.Azion.NET.VB.DatabaseManager,
                              ByVal infoNextSteps() As StepInfoItemExt) Implements IFlowSentNotification.Notify
        For Each infoNextStep As StepInfoItemExt In infoNextSteps
            Try
                Notify(dbManager, infoNextStep)
            Catch ex As Exception
            End Try
        Next
    End Sub



    ''' <summary>
    ''' Eloan保管袋通知
    ''' </summary>
    ''' <param name="dbManager"></param>
    ''' <param name="sCaseid"></param>
    ''' <param name="sNextUserId"></param>
    ''' <remarks></remarks>
    Public Sub Notify(ByVal dbManager As com.Azion.NET.VB.DatabaseManager,
                      ByVal sCaseid As String,
                      ByVal sNextUserId As String)


        Dim sContext As String
        sContext = "您於ELoan保管袋系統有一件待處理文件，案號：" & sCaseid

        Dim sii As New StepInfoItemExt
        Dim syUser As New FLOW_OP.TABLE.SY_USER(dbManager)

        sii.caseOwnerList = Nothing
        sii.caseOwner = syUser.GetUserIdName(sNextUserId)

        Notify(dbManager, sii, sContext)
    End Sub



    ''' <summary>
    ''' 流程送出後送EMAIL至USER
    ''' </summary>
    ''' <param name="dbManager"></param>
    ''' <param name="infoNextStep"></param>
    ''' <param name="sContext">若為空，則送出內容由程式組成</param>
    ''' <remarks></remarks>
    Public Overridable Sub Notify(ByVal dbManager As com.Azion.NET.VB.DatabaseManager,
                                  ByVal infoNextStep As StepInfoItemExt,
                                  Optional ByVal sContext As String = "")

        Dim syUser As FLOW_OP.TABLE.SY_USER = Nothing
        Dim syCaseId As FLOW_OP.TABLE.SY_CASEID = Nothing
        Dim syConfig As FLOW_OP.TABLE.SY_CONFIGURATION = Nothing
        Dim syAgent As BosBase
        Dim listMailedUser As New List(Of String)

        Dim mailMsg As New MailMessage
        Dim drCaseId As DataRow
        Dim sText As String
        Dim bInEniteProudctServer As Boolean = True
        'Dim sStepName As String

        Try
            syUser = New FLOW_OP.TABLE.SY_USER(dbManager)
            syCaseId = New FLOW_OP.TABLE.SY_CASEID(dbManager)
            syConfig = New FLOW_OP.TABLE.SY_CONFIGURATION(dbManager)
            syAgent = New BosBase("SY_USERROLEAGENT", dbManager)

            If Environment.MachineName <> syConfig.GetValue("SYSTEM_ENTIE_PRODUCT_WEBSERVER_NAME") Then
                bInEniteProudctServer = False
            End If

            'F消金的不用寄信
            If syUser.getSYCaseId.QueryString("[SY_FLOW_ID].[SYSID]", "FLOW_ID", infoNextStep.flowId) = "F" Then
                Return
            End If

            If String.IsNullOrEmpty(sContext) = True Then
                If IsNothing(infoNextStep.caseOwner) = False Then
                    sContext = "您於ELoan系統有一件待處理文件"
                Else
                    sContext = "您於ELoan系統有一件佇列文件待處理"
                End If

#If DEBUG Then
                FLOW_OP.Dbg.Assert(String.IsNullOrEmpty(infoNextStep.flowName) = False)
#End If

                'sStepName = syCaseId.QueryValue(Of String)("[SY_STEP_NO].[STEP_NAME]", infoNextStep.stepNo, "STEP_NO", infoNextStep.stepNo)

                sContext &= _
                    "，流程名稱：" & syCaseId.getSYFlowId.GetFlow_CName(infoNextStep.flowId) & _
                    "，案號：" & infoNextStep.caseId & _
                    "，流程步驟：" & infoNextStep.stepNo & _
                    "，步驟名稱：" & syCaseId.QueryString("[SY_STEP_NO].[STEP_NAME]", "STEP_NO", infoNextStep.stepNo)

                drCaseId = syCaseId.GetDataRow(Nothing, "CASEID", infoNextStep.caseId)
                If IsNothing(drCaseId) Then
                    Throw New SYException("案件不存在", SYMSG.SYFLOW_CASE_NOT_FOUND)
                End If

                'sText = BosBase.CDbType(Of String)(drCaseId("CPL_APL_ID"), "")
                'If String.IsNullOrEmpty(sText) = False Then
                '    sContext &= "，授信戶統編：" & sText
                'End If

                sText = BosBase.CDbType(Of String)(drCaseId("CPL_APL_NAM"), "")
                If String.IsNullOrEmpty(sText) = False Then
                    sContext &= "，授信戶名稱：" & sText
                End If

                sContext &= "。"
            End If



            mailMsg.Body = sContext
            mailMsg.Subject = sContext
            mailMsg.SubjectEncoding = System.Text.Encoding.UTF8
            mailMsg.From = New MailAddress(syConfig.GetValue("FLOW_MAILSENDEREMAIL"), syConfig.GetValue("FLOW_MAILSENDERNAME"))
            mailMsg.BodyEncoding = Encoding.UTF8
            mailMsg.IsBodyHtml = False
            mailMsg.Priority = MailPriority.Normal

            '設定收件人，非佇列，只有一人
            If IsNothing(infoNextStep.caseOwner) = False Then
                listMailedUser.Add(infoNextStep.caseOwner.userId)

                If bInEniteProudctServer Then
                    sText = syUser.GetEMail(infoNextStep.caseOwner.userId)
                Else
                    sText = syConfig.GetValue("FLOW_TEST_RECEIVER_EMAIL")
                End If

                mailMsg.To.Add(New MailAddress(sText, infoNextStep.caseOwner.userName))

            End If

            '設定收件人，佇列，有多個人
            If IsNothing(infoNextStep.caseOwnerList) = False Then
                For Each userIdName In infoNextStep.caseOwnerList
                    listMailedUser.Add(userIdName.userId)

                    If bInEniteProudctServer Then
                        sText = syUser.GetEMail(userIdName.userId)
                    Else
                        sText = syConfig.GetValue("FLOW_TEST_RECEIVER_EMAIL")
                    End If

                    If String.IsNullOrEmpty(sText) = False Then
                        mailMsg.To.Add(New MailAddress(sText, userIdName.userName))
                    End If
                Next
            End If

            '若有代理人，則需寄副本給代理人
            For Each Receiver As String In listMailedUser
                Dim drc As DataRowCollection

                drc = syAgent.GetDataRowCollection2( _
                    "select *" & vbCrLf & _
                    "  from SY_USERROLEAGENT" & vbCrLf & _
                    " where STAFFID = @STAFFID@" & vbCrLf & _
                    "   and STARTTIME <= GETDATE()" & vbCrLf & _
                    "   and ENDTIME >= GETDATE()" & vbCrLf & _
                    "   and CANCELTIME is NULL",
                    "STAFFID", Receiver)

                For Each drAgent As DataRow In drc
                    Dim AgentId, AgentEMail As String

                    AgentId = BosBase.CDbType(Of String)(drAgent("STAFFID"), "")
                    If String.IsNullOrEmpty(AgentId) = False Then
                        If bInEniteProudctServer Then
                            AgentEMail = syUser.GetEMail(AgentId)
                        Else
                            AgentEMail = syConfig.GetValue("FLOW_TEST_RECEIVER_EMAIL")
                        End If

                        If String.IsNullOrEmpty(AgentEMail) = False Then
                            mailMsg.CC.Add(New MailAddress(AgentEMail, AgentId))
                        End If
                    End If
                Next
            Next

            Dim params(5) As Object
            params(0) = mailMsg
            params(1) = syConfig.GetValue("FLOW_SMTPSERVER")
            params(2) = syConfig.GetValue("FLOW_MAILLOGINACCOUNT")
            params(3) = syConfig.GetValue("FLOW_MAILLOGINPASSWORD")
            params(4) = syConfig.GetValue("FLOW_SMTP_PORT")
            params(5) = IIf(syConfig.GetValue("FLOW_SMTP_IS_SSL") = "Y", True, False)

            Dim MailThread As New Thread(AddressOf Me.SendEMailThread)
            MailThread.Start(params)

        Catch ex As Exception
            Throw
        End Try

    End Sub



    ''' <summary>
    ''' 送EMAIL至USER
    ''' </summary>
    ''' <param name="dbManager"></param>
    ''' <param name="sStaffid"></param>
    ''' <param name="sSubject"></param>
    ''' <param name="sBody"></param>
    ''' <remarks></remarks>
    Public Overridable Sub EMail(ByVal dbManager As com.Azion.NET.VB.DatabaseManager,
                                  ByVal sStaffid As String,
                                  ByVal sSubject As String,
                                  ByVal sBody As String)
        Try

            Dim syUser As New FLOW_OP.TABLE.SY_USER(dbManager)
            Dim syConfig As New FLOW_OP.TABLE.SY_CONFIGURATION(dbManager)
            Dim mailMsg As New MailMessage
            Dim sText As String
            Dim bInEniteProudctServer As Boolean = True

            If Environment.MachineName <> syConfig.GetValue("SYSTEM_ENTIE_PRODUCT_WEBSERVER_NAME") Then
                bInEniteProudctServer = False
            End If

            mailMsg.Body = sBody
            mailMsg.Subject = sSubject
            mailMsg.SubjectEncoding = System.Text.Encoding.UTF8
            mailMsg.From = New MailAddress(syConfig.GetValue("FLOW_MAILSENDEREMAIL"), syConfig.GetValue("FLOW_MAILSENDERNAME"))
            mailMsg.BodyEncoding = Encoding.UTF8
            mailMsg.IsBodyHtml = False
            mailMsg.Priority = MailPriority.Normal

            '設定收件人，非佇列，只有一人
            If bInEniteProudctServer Then
                sText = syUser.GetEMail(sStaffid)
            Else
                sText = syConfig.GetValue("FLOW_TEST_RECEIVER_EMAIL")
            End If

            mailMsg.To.Add(New MailAddress(sText, syUser.GetUserIdName(sStaffid).userName))


            Dim params(5) As Object
            params(0) = mailMsg
            params(1) = syConfig.GetValue("FLOW_SMTPSERVER")
            params(2) = syConfig.GetValue("FLOW_MAILLOGINACCOUNT")
            params(3) = syConfig.GetValue("FLOW_MAILLOGINPASSWORD")
            params(4) = syConfig.GetValue("FLOW_SMTP_PORT")
            params(5) = IIf(syConfig.GetValue("FLOW_SMTP_IS_SSL") = "Y", True, False)

            Dim MailThread As New Thread(AddressOf Me.SendEMailThread)
            MailThread.Start(params)


        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' 送EMAIL至USER (多人)
    ''' </summary>
    ''' <param name="dbManager"></param>
    ''' <param name="sStaffid"></param>
    ''' <param name="sStaffidCC"></param>
    ''' <param name="sSubject"></param>
    ''' <param name="sBody"></param>
    ''' <remarks></remarks>
    Public Overridable Function MultiEMail(ByVal dbManager As com.Azion.NET.VB.DatabaseManager,
                                  ByVal sStaffid As ArrayList,
                                  ByVal sStaffidCC As ArrayList,
                                  ByVal sSubject As String,
                                  ByVal sBody As String) As String
        Try

            Dim syUser As New FLOW_OP.TABLE.SY_USER(dbManager)
            Dim syConfig As New FLOW_OP.TABLE.SY_CONFIGURATION(dbManager)
            Dim mailMsg As New MailMessage
            Dim sText As String
            Dim bInEniteProudctServer As Boolean = True
            Dim sEX As String = String.Empty

            If Environment.MachineName <> syConfig.GetValue("SYSTEM_ENTIE_PRODUCT_WEBSERVER_NAME") Then
                bInEniteProudctServer = False
            End If

            mailMsg.Body = sBody
            mailMsg.Subject = sSubject
            mailMsg.SubjectEncoding = System.Text.Encoding.UTF8
            mailMsg.From = New MailAddress(syConfig.GetValue("FLOW_MAILSENDEREMAIL"), syConfig.GetValue("FLOW_MAILSENDERNAME"))
            mailMsg.BodyEncoding = Encoding.UTF8
            mailMsg.IsBodyHtml = False
            mailMsg.Priority = MailPriority.Normal

            '正本
            For I As Integer = 0 To sStaffid.Count - 1
                '設定收件人，非佇列，只有一人
                If bInEniteProudctServer Then
                    sText = syUser.GetEMail(sStaffid(I))
                Else
                    sText = syConfig.GetValue("FLOW_TEST_RECEIVER_EMAIL")
                End If

                If sText.IndexOf("@") > 0 Then
                    mailMsg.To.Add(New MailAddress(sText, syUser.GetUserIdName(sStaffid(I)).userName))
                Else
                    sEX &= "Email錯誤(" & sStaffid(I) & "):" & sText & ";"
                End If

            Next

            '副本
            For J As Integer = 0 To sStaffidCC.Count - 1
                '設定收件人，非佇列，只有一人
                If bInEniteProudctServer Then
                    sText = syUser.GetEMail(sStaffidCC(J))
                Else
                    sText = syConfig.GetValue("FLOW_TEST_RECEIVER_EMAIL")
                End If

                If sText.IndexOf("@") > 0 Then
                    mailMsg.CC.Add(New MailAddress(sText, syUser.GetUserIdName(sStaffidCC(J)).userName))
                Else
                    sEX &= "Email錯誤(" & sStaffidCC(J) & "):" & sText & ";"
                End If

            Next

            Dim params(5) As Object
            params(0) = mailMsg
            params(1) = syConfig.GetValue("FLOW_SMTPSERVER")
            params(2) = syConfig.GetValue("FLOW_MAILLOGINACCOUNT")
            params(3) = syConfig.GetValue("FLOW_MAILLOGINPASSWORD")
            params(4) = syConfig.GetValue("FLOW_SMTP_PORT")
            params(5) = IIf(syConfig.GetValue("FLOW_SMTP_IS_SSL") = "Y", True, False)

            Dim MailThread As New Thread(AddressOf Me.SendEMailThread)
            MailThread.Start(params)

            Return sEX
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' 送EMAIL至USER (多人)
    ''' </summary>
    ''' <param name="dbManager"></param>
    ''' <param name="HT_STAFFID"></param>
    ''' <param name="HT_STAFFIDCC"></param>
    ''' <param name="sSubject"></param>
    ''' <param name="sBody"></param>
    ''' <remarks></remarks>
    Public Overridable Sub MultiEMailByHT(ByVal dbManager As com.Azion.NET.VB.DatabaseManager,
                                  ByVal HT_STAFFID As Hashtable,
                                  ByVal HT_STAFFIDCC As Hashtable,
                                  ByVal sSubject As String,
                                  ByVal sBody As String)
        Try

            Dim syUser As New FLOW_OP.TABLE.SY_USER(dbManager)
            Dim syConfig As New FLOW_OP.TABLE.SY_CONFIGURATION(dbManager)
            Dim mailMsg As New MailMessage
            'Dim bInEniteProudctServer As Boolean = True

            'If Environment.MachineName <> syConfig.GetValue("SYSTEM_ENTIE_PRODUCT_WEBSERVER_NAME") Then
            '    bInEniteProudctServer = False
            'End If

            mailMsg.Body = sBody
            mailMsg.Subject = sSubject
            mailMsg.SubjectEncoding = System.Text.Encoding.UTF8
            mailMsg.From = New MailAddress(syConfig.GetValue("FLOW_MAILSENDEREMAIL"), syConfig.GetValue("FLOW_MAILSENDERNAME"))
            mailMsg.BodyEncoding = Encoding.UTF8
            mailMsg.IsBodyHtml = False
            mailMsg.Priority = MailPriority.Normal

            '正本
            'For I As Integer = 0 To sStaffid.Count - 1
            '    '設定收件人，非佇列，只有一人
            '    If bInEniteProudctServer Then
            '        sText = syUser.GetEMail(sStaffid(I))
            '    Else
            '        sText = syConfig.GetValue("FLOW_TEST_RECEIVER_EMAIL")
            '    End If

            '    mailMsg.To.Add(New MailAddress(sText, syUser.GetUserIdName(sStaffid(I)).userName))
            'Next
            For Each sKey As String In HT_STAFFID.Keys
                mailMsg.To.Add(New MailAddress(sKey, HT_STAFFID.Item(sKey)))
            Next

            '副本
            'For J As Integer = 0 To sStaffidCC.Count - 1
            '    '設定收件人，非佇列，只有一人
            '    If bInEniteProudctServer Then
            '        sText = syUser.GetEMail(sStaffidCC(J))
            '    Else
            '        sText = syConfig.GetValue("FLOW_TEST_RECEIVER_EMAIL")
            '    End If

            '    mailMsg.CC.Add(New MailAddress(sText, syUser.GetUserIdName(sStaffidCC(J)).userName))
            'Next
            For Each sKey As String In HT_STAFFIDCC.Keys
                mailMsg.To.Add(New MailAddress(sKey, HT_STAFFIDCC.Item(sKey)))
            Next

            Dim params(5) As Object
            params(0) = mailMsg
            params(1) = syConfig.GetValue("FLOW_SMTPSERVER")
            params(2) = syConfig.GetValue("FLOW_MAILLOGINACCOUNT")
            params(3) = syConfig.GetValue("FLOW_MAILLOGINPASSWORD")
            params(4) = syConfig.GetValue("FLOW_SMTP_PORT")
            params(5) = IIf(syConfig.GetValue("FLOW_SMTP_IS_SSL") = "Y", True, False)

            Dim MailThread As New Thread(AddressOf Me.SendEMailThread)
            MailThread.Start(params)

        Catch ex As Exception
            Throw
        End Try
    End Sub

    Protected Sub SendEMailThread(ByVal data As Object)
        Try
            Dim params() As Object
            params = data

            sendMail(params(0),
                      params(1),
                      params(2),
                      params(3),
                      params(4),
                      params(5))
        Catch ex As Exception
        End Try
    End Sub


    Public Shared Sub sendMail(ByVal mailMsg As MailMessage,
                               ByVal sSMTP As String,
                               ByVal sSMTPAcc As String,
                               ByVal sSMTPPwd As String,
                               ByVal sSMTPPort As String,
                               Optional isSSL As Boolean = False)

        'SMTP Setting 
        Dim client As New SmtpClient()
        client.Host = sSMTP

        If com.Azion.EloanUtility.ValidateUtility.isValidateData(sSMTPPort) Then
            client.Port = CInt(sSMTPPort)
        End If

        If com.Azion.EloanUtility.ValidateUtility.isValidateData(sSMTPAcc) And com.Azion.EloanUtility.ValidateUtility.isValidateData(sSMTPPwd) Then
            client.Credentials = New NetworkCredential(sSMTPAcc, sSMTPPwd)
        End If

        client.EnableSsl = isSSL

        ' Send Mail 
        Try
            client.Send(mailMsg)
        Catch ex As Exception
#If _MAIL_LOG = 1 Then
            ShowMsgThread(ex.ToString)
#End If
            Throw ex

        End Try

        AddHandler client.SendCompleted, AddressOf client_SendCompleted

        ' Sent Compeleted Eevet 
    End Sub

    Protected Shared Sub client_SendCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.AsyncCompletedEventArgs)
        If e.[Error] IsNot Nothing Then
            'MessageBox.Show(e.[Error].ToString())
            Throw New Exception(e.[Error].ToString())
        End If
    End Sub

#If _MAIL_LOG = 1 Then

    Public Shared Sub ShowMsgThread(ByVal sMsg As String, Optional ByVal flgMust As Boolean = False)
        'If bDispMode Or flgMust Then
        Console.WriteLine("")
        Console.WriteLine("[" & Now & "]" & sMsg)
        'End If

        Try
            Monitor.Enter(m_SYNC)
            EventLog(sMsg, "EMailThread" & Format(Now(), "yyyyMMdd") & ".txt", FileMode.Append)
        Catch ex As Exception
        Finally
            Monitor.Exit(m_SYNC)
        End Try

    End Sub
#End If

    Private Shared Sub EventLog(sMsg As String, p2 As String, fileMode As FileMode)
        Throw New NotImplementedException
    End Sub


End Class
