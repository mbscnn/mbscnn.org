Imports com.Azion.EloanUtility
Imports com.Azion.NET.VB
Imports MBSC.MB_OP
Imports FLOW_OP

Public Class UIShareFun

#Region "DataBaseManager"

#Region "Property"

    Private Shared _sConnectionString As String = getConnectionString()
    Shared ReadOnly Property ConnectionString As String
        Get
            Return _sConnectionString
        End Get
    End Property

#End Region

    ''' <summary>
    ''' 取得連接字串
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function getConnectionString() As String
        Return System.Web.HttpContext.Current.Application.Item("BotDSN")
        '"Data Source=" & com.Azion.EloanUtility.FileUtility.getAppSettings("DBSource") & ";database=" & com.Azion.EloanUtility.FileUtility.getAppSettings("DataBase") & ";User ID=" & com.Azion.EloanUtility.FileUtility.getAppSettings("DBUserID") & ";Persist Security Info=True;Pooling=true;Min Pool Size=" & com.Azion.EloanUtility.FileUtility.getAppSettings("MinPool") & ";Max Pool Size=" & com.Azion.EloanUtility.FileUtility.getAppSettings("MaxPool") & ";Connection Lifetime=" & com.Azion.EloanUtility.FileUtility.getAppSettings("ConnectionLifetime") & ";Connect Timeout=" & com.Azion.EloanUtility.FileUtility.getAppSettings("ConnectTimeout") & ";Application Name=" & com.Azion.EloanUtility.FileUtility.getAppSettings("AppName") & ";Password=" & com.Azion.EloanUtility.EncryptUtility.Decrypto(com.Azion.EloanUtility.FileUtility.getAppSettings("DBPassword"))
    End Function

    ''' <summary>
    ''' 取得ELoan DatabaseManager Object
    ''' </summary>
    ''' <param name="bCreatedByUIBase">是否是由UIBase所建立的，UIBase所建立的Connection會自行釋放</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Shared Function getDataBaseManager(Optional ByVal bCreatedByUIBase As Boolean = False) As com.Azion.NET.VB.DatabaseManager
        Return com.Azion.NET.VB.DatabaseManager.getInstance(ConnectionString, 0, bCreatedByUIBase)
    End Function

    ''' <summary>
    ''' release DatabaseManager
    ''' </summary>
    ''' <param name="dbManager"></param>
    ''' <remarks></remarks>
    Shared Sub releaseConnection(ByRef dbManager As com.Azion.NET.VB.DatabaseManager)
        Try
            If Not IsNothing(dbManager) Then
                dbManager.releaseConnection()
            End If
        Catch ex As Exception
            Throw New com.Azion.NET.VB.BosException(ex, CType(ex.TargetSite, System.Reflection.MethodInfo))
        End Try
    End Sub

#End Region

#Region "錯誤資訊的開關 "
    Public Shared Sub showErrMsg(ByRef page As System.Web.UI.Page, ByVal obj As Object)
        Try
            'If UIShareFun.isTabPostBack(page) Then
            '    Dim pageTab As PageTab = CType(page.FindControl("PageTab"), PageTab)
            '    If TypeOf obj Is Exception Then
            '        pageTab.setException(CType(obj, Exception))
            '    Else
            '        pageTab.setException(New Exception(CType(obj, String)))
            '    End If
            'End If

            Dim lblErrorMsg As Label = CType(page.FindControl("lblErrorMsg"), System.Web.UI.WebControls.Label)
            Dim palErrorMsg As Panel = CType(page.FindControl("ErrorMsg"), System.Web.UI.WebControls.Panel)

            Dim sErrorMsg As String = String.Empty

            If TypeOf obj Is Exception Then
                Dim exception As Exception = CType(obj, Exception)
                Dim sb As New System.Text.StringBuilder

                Do
                    If Utility.isValidateData(exception.Message) Then
                        sb.Append(exception.Message.ToString & "。<br>")
                    End If

                    If Utility.isValidateData(exception.StackTrace) Then
                        sb.Append(exception.StackTrace.ToString & "。<br>")
                    End If

                    exception = exception.InnerException
                Loop Until (exception Is Nothing)

                'ENLogger.log.Debug(obj)
                sErrorMsg = sb.ToString
            Else
                sErrorMsg = CType(obj, String)
            End If

            If Not IsNothing(palErrorMsg) Then
                palErrorMsg.Visible = True
                lblErrorMsg.Visible = True
                lblErrorMsg.Text = sErrorMsg
            Else
                sErrorMsg = "<span style='color:red'>" & sErrorMsg & "</span>"

                HttpContext.Current.Response.Write(sErrorMsg)
            End If
        Catch ex As Exception
            Dim s As String = CType(obj, Exception).Message.ToString & "<br>" & CType(obj, Exception).StackTrace
        End Try
    End Sub

    Public Shared Sub closeErrMsg()
        Dim page As Page = CType(HttpContext.Current.Handler, Page)
        Try
            Dim lblErrorMsg As Label = CType(page.FindControl("lblErrorMsg"), System.Web.UI.WebControls.Label)
            Dim ErrorMsg As Panel = CType(page.FindControl("ErrorMsg"), System.Web.UI.WebControls.Panel)

            lblErrorMsg.Text = ""
            ErrorMsg.Visible = False
        Catch ex As Exception

        End Try
    End Sub

    Public Shared Function getEXInfo(ByVal obj As Object) As String
        Dim sErrorMsg As String = String.Empty

        If TypeOf obj Is Exception Then
            Dim exception As Exception = CType(obj, Exception)
            Dim sb As New System.Text.StringBuilder

            Do
                If Utility.isValidateData(exception.Message) Then
                    sb.Append(exception.Message.ToString & "。<br>")
                End If

                If Utility.isValidateData(exception.StackTrace) Then
                    sb.Append(exception.StackTrace.ToString & "。<br>")
                End If

                exception = exception.InnerException
            Loop Until (exception Is Nothing)

            sErrorMsg = sb.ToString
        Else
            sErrorMsg = CType(obj, String)
        End If

        Return sErrorMsg
    End Function
#End Region

#Region "Web"
    Public Shared Function getLoginUserid() As String
        Dim httpContext As System.Web.HttpContext = System.Web.HttpContext.Current
        Dim sUserId As String = ""

        If IsNothing(httpContext) Then Return sUserId

        If Utility.isValidateData(httpContext.Current.Session("userid")) Then
            sUserId = httpContext.Current.Session("userid")
        ElseIf Utility.isValidateData(httpContext.Current.Request("userid")) Then
            sUserId = httpContext.Current.Request("userid")
            httpContext.Current.Session("userid") = sUserId
        End If

        Return sUserId
    End Function

    Public Shared Function getLoginBrid() As String
        Dim httpContext As System.Web.HttpContext = System.Web.HttpContext.Current
        Dim sBrid As String = ""

        If IsNothing(httpContext) Then Return sBrid

        If Utility.isValidateData(httpContext.Current.Session("brid")) Then
            sBrid = httpContext.Current.Session("brid")
        ElseIf Utility.isValidateData(httpContext.Current.Request("brid")) Then
            sBrid = httpContext.Current.Request("brid")
            httpContext.Current.Session("brid") = sBrid
        End If

        Return sBrid
    End Function

    Public Shared Function getMailBody(ByVal sUID As String, ByVal sMail As String, ByVal sAPPNAME As String,Byval sMB_MEMSEQ As string) As String
        Try
            Dim sb As New StringBuilder
            sb.Append("<P>")
            sb.Append("親愛的" & sAPPNAME & "您好：<BR/>")
            sb.Append("這封信是由佛陀原始正法中心發送的。<BR/>")
            sb.Append("非常感謝您註冊成為我們的會員。<BR/>")
            sb.Append("您需點擊下面的連結啟用您的帳號：<BR/>")
            'sb.Append("<a href='http://www.mbscnn.org/mbsc/mnt/MBMnt_Reg_01_v01.aspx?MOD=APV&UID=" & sUID & "&MAIL=" & sMail & "'>http://www.mbscnn.org/mbsc/mnt/MBMnt_Reg_01_v01.aspx?MOD=APV&UID=" & sUID & "&MAIL=" & sMail & "</a>")
            'sb.Append("<a href='http://www.mbscnn.org/mbsc/mnt/MBMnt_Reg_01_v01.aspx?MOD=APV&UID=" & sUID & "'>http://www.mbscnn.org/mbsc/mnt/MBMnt_Reg_01_v01.aspx?MOD=APV&UID=" & sUID & "</a>")
            'sb.Append("http://www.mbscnn.org/mbsc/mnt/MBMnt_Reg_01_v01.aspx?MOD=APV&UID=" & sUID)
            Dim sMB_MEMSEQ_URL As String = String.Empty
            If IsNumeric(sMB_MEMSEQ) Then
                sMB_MEMSEQ_URL = "&MB_MEMSEQ=" & sMB_MEMSEQ
            End If
            sb.Append("<a href='http://www.mbscnn.org/mnt/MBMnt_Reg_01_v01.aspx?MOD=APV&UID=" & sUID & "'>http://www.mbscnn.org/mnt/MBMnt_Reg_01_v01.aspx?MOD=APV&UID=" & sUID & sMB_MEMSEQ_URL & "</a>")
            sb.Append("<BR/>")
            sb.Append("<span style='color:red'>(如果上面不是連結形式，請將該位址貼到瀏覽器網址欄)</span>")
            sb.Append("<BR/>")
            sb.Append("帳號啟用後，請您記得至會員系統填寫入會申請單！")
            sb.Append("<BR/>")
            sb.Append("感謝您的訪問，祝順心愉快！")
            sb.Append("<BR/>")
            sb.Append("此致")
            sb.Append("</P>")

            Return sb.ToString
        Catch ex As Exception
            Throw
        End Try
    End Function

#End Region

#Region "會員取號"
    Public Shared Function getVadID(ByVal dbManager As com.Azion.NET.VB.DatabaseManager) As String
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

#End Region

    Public Shared Sub downLoadFile(ByVal page As System.Web.UI.Page, ByVal sFileName As String, ByVal str As String)
        page.Response.ClearHeaders()
        page.Response.Clear()
        page.Response.Expires = 0
        page.Response.Buffer = True
        '原本有編碼的問題,出來的EXCEL內容會出現亂碼,是因為系統採"繁體中文BIG5"的關係,改成UTF-8就行了[原因不明,待查]
        page.Response.ContentEncoding = System.Text.Encoding.UTF8
        page.Response.AddHeader("Accept-Language", "zh-tw")
        page.Response.ContentType = "application/octet-stream; charset=iso-8859-1"  '"Application/octet-stream"
        'page.Response.ContentType = "Application/vnd.ms-excel" 'GAQryCase02的 System.Text.Encoding.UTF8.GetString(System.Text.Encoding.UTF8.GetBytes(sFileName))
        page.Response.AddHeader("content-disposition", "attachment; filename=" & Chr(34) & System.Web.HttpUtility.UrlEncode(sFileName, System.Text.Encoding.UTF8) & Chr(34))
        page.Response.Write(str)
        page.Response.End()
    End Sub

#Region "User"
    Shared Function getRoles(ByVal databaseManager As DatabaseManager, ByVal sStaffId As String) As DataTable
        Dim flowFacade As New FLOW_OP.FlowFacade(databaseManager)
        Return flowFacade.GetRolesByStaffid(sStaffId)
    End Function
#End Region

#Region "startFlow"
    ''' <summary>
    ''' 新流程
    ''' bAutoSendToNext=Tru，跳過第一關直接到第二關，第二關邏輯依FLOW_MAP定義
    ''' 假如bAutoSendToNext=True
    ''' 請指定第二關人員，假使第一關人員與第二關人員相同，也請傳入參數:StaffInfo.WorkingId
    ''' </summary>
    ''' <param name="databaseManager"></param>
    ''' <param name="sFlowName">流程名稱</param>
    ''' <param name="bAutoSendToNext">boolean,是否跳過第一關直接到第二關，依FLOW_MAP定義</param>
    ''' <param name="sNextUserid">String</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' created  by Titan 2012/06/27
    ''' </remarks>
    Shared Function startFlow(ByVal databaseManager As DatabaseManager,
                              ByVal sFlowName As String,
                              Optional ByVal bAutoSendToNext As Boolean = False,
                              Optional ByVal sNextUserid As String = Nothing,
                              Optional ByVal bSendToEnd As Boolean = False,
                              Optional ByVal sBrid As String = "",
                              Optional ByVal sApvCasId As String = "",
                              Optional ByVal sAppBrid As String = "",
                              Optional ByVal sCplAplId As String = "",
                              Optional ByVal sCplAplNam As String = "",
                              Optional ByVal bSendMail As Boolean = True) As FLOW_OP.StepInfo

        Dim flowFacade As New FLOW_OP.FlowFacade(databaseManager)

        If bSendToEnd Then
            Return flowFacade.StartFlow(sBrid, sFlowName, True, sNextUserid, bSendToEnd, sApvCasId, sAppBrid, sCplAplId, sCplAplNam, bSendMail)
        End If

        'If bAutoSendToNext Then
        '    'If sNextUserid Is Nothing OrElse sNextUserid = "" Then
        '    '    Throw New Exception("請指定第二關人員，假使第一關人員與第二關人員相同，也請傳入參數:StaffInfo.WorkingId")
        '    'End If
        '    Return flowFacade.StartFlow("", sFlowName, bAutoSendToNext, sNextUserid)
        'End If

        Return flowFacade.StartFlow(sBrid, sFlowName, bAutoSendToNext, sNextUserid, bSendToEnd, sApvCasId, sAppBrid, sCplAplId, sCplAplNam, bSendMail)
    End Function

    ''' <summary>
    ''' 直接起案並回傳new caseid
    ''' </summary>
    ''' <param name="databaseManager"></param>
    ''' <param name="subsysId">子系統代號ex:06,05,04</param>
    ''' <param name="sBrId">分行別</param>
    ''' <param name="_sFlowName">Flow Name</param>
    ''' <param name="_sFirstUserId">員工編號ex:S064933</param>
    ''' <param name="sCaseid">回傳  new Caseid</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    'Private Shared Function startFlow(ByVal databaseManager As DatabaseManager, ByVal subsysId As String, ByVal sBrId As String, _
    '                          ByVal _sFlowName As String, ByVal _sFirstUserId As String, ByRef sCaseid As String, Optional ByVal bAutoSendToNext As Boolean = False) As Boolean

    '    If Not com.Azion.EloanUtility.UIUtility.isNewAuthorize() Then
    '        sCaseid = IEFLOW.getNewCaseId(databaseManager, subsysId, sBrId)
    '        Return IEFLOW.startFlowByCaseId(databaseManager, sCaseid, _sFlowName, "", "", _sFirstUserId)
    '    End If

    '    Dim stepInfo As FLOW_OP.StepInfo = UIShareFun.startFlow(databaseManager, _sFlowName, bAutoSendToNext, _sFirstUserId)
    '    sCaseid = stepInfo.currentStepInfo.caseId
    '    Return True
    'End Function

#End Region

#Region "取得流程資訊"
    ''' <summary>    
    ''' 取得目前案件最新的流程步驟資訊
    ''' </summary>
    ''' <param name="databaseManager">DatabaseManager</param>
    ''' <param name="sCaseId"></param> 
    ''' <returns></returns>
    ''' <remarks></remarks>
    Shared Function getCurrentStepInfo(ByVal databaseManager As DatabaseManager, ByVal sCaseId As String, Optional ByVal nSubFlowSeq As Integer = 0) As FLOW_OP.StepInfoItemExt
        Dim flowFacade As New FLOW_OP.FlowFacade(databaseManager)

        Return (flowFacade.getELoanFlow().GetCurrentStepInfoOfCaseid(sCaseId, nSubFlowSeq))
    End Function
#End Region

#Region "送關元件"
    ''' <summary>
    ''' 送關元件
    ''' </summary>
    ''' <param name="databaseManager"></param>
    ''' <param name="sCaseId">String</param>
    ''' <param name="nSubFlowSeq">Integer</param>
    ''' <param name="sNextStep">String</param>
    ''' <param name="sNextUserid">String</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Shared Function pushFlow(ByVal databaseManager As DatabaseManager, ByVal sCaseId As String,
                             Optional ByVal nSubFlowSeq As Integer = 0,
                             Optional ByVal sNextStep As String = "",
                             Optional ByVal sNextUserid As String = "",
                             Optional ByVal bShowMsg As Boolean = True) As FLOW_OP.StepInfo

        Dim flowFacade As New FLOW_OP.FlowFacade(databaseManager)
        Dim stepInfo As FLOW_OP.StepInfo = flowFacade.SendFlow(sCaseId, nSubFlowSeq, sNextStep, sNextUserid)


        If bShowMsg Then
            If flowFacade.IsCaseCompleted(stepInfo.currentStepInfo.caseId) Then
                com.Azion.EloanUtility.UIUtility.alert("結案成功")
            Else
                Dim sMsg As String = "案號:[" & stepInfo.currentStepInfo.caseId & "]已送出"
                com.Azion.EloanUtility.UIUtility.alert(sMsg)
            End If
            Dim js As String = "<Script Language='JAVAScript'>" & vbCrLf
            js += "self.close();" & vbCrLf
            js += "if (opener!=undefined){if (opener.window!=null){opener.window.close();}}" & vbCrLf
            js += "if (parent.window!=null){parent.window.close();}" & vbCrLf
            js += "if (window.dialogArguments!=null){if (window.dialogArguments.window!=undefined){window.dialogArguments.close();}}" & vbCrLf
            js += "</Script>" & vbCrLf
            HttpContext.Current.Response.Write(js)
        End If

        Return stepInfo
    End Function


    Shared Function pushFlow(ByVal databaseManager As DatabaseManager, ByVal sCaseId As String,
                             ByVal nSubFlowSeq As Integer,
                             ByVal sNextStep As String,
                             ByVal sNextUserid As String,
                             ByVal nBraDepNo As Integer,
                             Optional ByVal bShowMsg As Boolean = True) As FLOW_OP.StepInfo

        Dim flowFacade As New FLOW_OP.FlowFacade(databaseManager)
        Dim stepInfo As FLOW_OP.StepInfo = flowFacade.SendFlow(sCaseId, nSubFlowSeq, sNextStep, sNextUserid, nBraDepNo)


        If bShowMsg Then
            If flowFacade.IsCaseCompleted(stepInfo.currentStepInfo.caseId) Then
                com.Azion.EloanUtility.UIUtility.alert("結案成功")
            Else
                Dim sMsg As String = "案號:[" & stepInfo.currentStepInfo.caseId & "]已送出"
                com.Azion.EloanUtility.UIUtility.alert(sMsg)
            End If
            Dim js As String = "<Script Language='JAVAScript'>" & vbCrLf
            js += "self.close();" & vbCrLf
            js += "if (opener!=undefined){if (opener.window!=null){opener.window.close();}}" & vbCrLf
            js += "if (parent.window!=null){parent.window.close();}" & vbCrLf
            js += "if (window.dialogArguments!=null){if (window.dialogArguments.window!=undefined){window.dialogArguments.close();}}" & vbCrLf
            js += "</Script>" & vbCrLf
            HttpContext.Current.Response.Write(js)
        End If

        Return stepInfo
    End Function
#End Region

#Region "退關元件(修正補充)"
    ''' <summary>
    ''' 退關元件(修正補充)
    ''' 退關後會顯示訊息 
    ''' </summary>
    ''' <param name="databaseManager">DatabaseManager</param>
    ''' <param name="sCaseId"></param>
    ''' <param name="nSubFlowSeq"></param>
    ''' <param name="sRevisionNotes">理由</param>
    ''' <param name="sDestStpNo">Optional 指定某一關，步驟代碼</param>
    ''' <param name="sNextUserid">Optional 指定某一user，員工編號ex:S064933</param>
    ''' <param name="iCount"></param>
    ''' <remarks></remarks>
    Shared Function rollBack(ByVal databaseManager As DatabaseManager, ByVal sCaseId As String,
                                 Optional ByVal nSubFlowSeq As Integer = 0,
                                 Optional ByVal sRevisionNotes As String = "",
                                 Optional ByVal sDestStpNo As String = "",
                                 Optional ByVal sNextUserid As String = "",
                                 Optional ByVal iCount As Integer = 0,
                                 Optional ByVal bShowMsg As Boolean = True) As FLOW_OP.StepInfo

        Dim stepInfo As FLOW_OP.StepInfo = Nothing

        'if forwarding step is more than one 

        Try
            If iCount <= 1 Then
                Dim flowFacade As New FLOW_OP.FlowFacade(databaseManager)
                stepInfo = flowFacade.JumpRollBack(sCaseId, nSubFlowSeq, sRevisionNotes, sDestStpNo, sNextUserid)
            Else
                If sDestStpNo = "" Then Throw New Exception("請選擇送回步驟!!")

                Dim sFunctionName As String = "rollBack_" & sDestStpNo
                Dim flowRollBack As New FlowRollBack(databaseManager, sCaseId)

                Dim method As System.Reflection.MethodInfo = flowRollBack.GetType.GetMethod(sFunctionName)
                If Not method Is Nothing Then
                    '未帶參數的
                    sNextUserid = method.Invoke(flowRollBack, Nothing)
                End If
                Dim flowFacade As New FLOW_OP.FlowFacade(databaseManager)
                stepInfo = flowFacade.JumpRollBack(sCaseId, nSubFlowSeq, sRevisionNotes, sDestStpNo, sNextUserid)
            End If

            If (bShowMsg) Then
                Dim sMsg As String = "案號" & stepInfo.nextStepInfo(0).caseId & "修正補充完成! \n"

                If IsNothing(stepInfo.nextStepInfo(0).caseOwner) = False Then
                    sMsg &= stepInfo.nextStepInfo(0).caseOwner.userId
                ElseIf stepInfo.nextStepInfo(0).caseOwnerList.Count > 0 Then
                    Dim nIndex As Integer = 0

                    For Each user As FLOW_OP.USER_ID_NAME In stepInfo.nextStepInfo(0).caseOwnerList
                        If nIndex > 0 Then
                            sMsg &= " ,"
                        End If
                        sMsg &= user.userId
                        nIndex = nIndex + 1
                    Next
                End If

                jumpRollBackMsg(sMsg)
            End If
        Catch ex As Exception
            Throw
        End Try

        Return stepInfo
    End Function

    Private Shared Sub jumpRollBackMsg(ByVal sMsg As String)

        com.Azion.EloanUtility.UIUtility.alert(sMsg)
        Dim js As String = "<Script Language='JAVAScript'>" & vbCrLf
        js += "self.close();" & vbCrLf
        js += "if (opener!=undefined){if (opener.window!=null){opener.window.close();}}" & vbCrLf
        js += "if (parent.window!=null){parent.window.close();}" & vbCrLf
        js += "if (window.dialogArguments!=null){if (window.dialogArguments.window!=undefined){window.dialogArguments.close();}}" & vbCrLf
        js += "</Script>" & vbCrLf

        HttpContext.Current.Response.Write(js)
    End Sub
#End Region


End Class
