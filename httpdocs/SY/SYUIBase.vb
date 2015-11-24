Option Explicit On
Option Strict On

Imports com.Azion.EloanUtility
Imports com.Azion.NET.VB
Imports MBSC.UICtl
Imports FLOW_OP
Imports System.Text.RegularExpressions

Public Class SYUIBase
    Inherits System.Web.UI.Page

    Protected m_bTesting As Boolean = False  '目前是否為測試模式
    Protected m_bDebug As Boolean = False '目前是否為Debug模式

    Protected m_bDisplayMode As Boolean '是否在可編輯模式 
    Protected m_bCheck As Boolean = False '目前是否為Check模式

    Protected m_sFlowName As String = String.Empty
    Protected m_sFlowId As Integer = 0 ' FlowId
    Protected m_sWatchStepNo As String = String.Empty '查詢時使用者想看的步驟，非版本版號，查詢時才會傳入
    Protected m_sCaseId As String = String.Empty ' 案件編號
    Protected m_sSubFlowSeq As String = String.Empty '子流程編號
    Protected m_sStepNo As String = String.Empty '專案目前步驟
    Protected m_sFourStepNo As String = String.Empty ' 專案目前步驟后四碼 

    Protected g_oUserInfo As StaffInfo  '取得登入者信息 

    Protected m_sLoginUserid As String = String.Empty '目前登入使用者，內部系統userid(7碼)(S0XXXXX)(S000819)
    Protected m_sLoginBrid As String = String.Empty '目前登入使用者的分行別(3碼)
    Protected m_sLoginTopDepNo As String = String.Empty '目前登入使用者的最上層部門序號
    Protected m_sLoginRoleId As String = String.Empty '目前登入使用者的角色 
    Protected m_sLoginEnTieUserid As String = String.Empty '目前登入使用者，安泰userid(5碼)(XXXXX)(00819)

    Protected m_sWorkingUserid As String = String.Empty '目前被代理使用者，內部系統userid(7碼)(S0XXXXX)(S000819)
    Protected m_sWorkingBrid As String = String.Empty '目前被代理使用者的分行別(3碼)
    Protected m_sWorkingTopDepNo As String = String.Empty '目前被代理使用者最上層部的部門序號
    Protected m_sWorkingRoleId As String = String.Empty '目前被代理使用者的角色 
    Protected m_sWorkingEnTieUserid As String = String.Empty '目前被代理使用者，安泰userid(5碼)(XXXXX)(00819)

    Protected m_sFuncCode As String = String.Empty '從左側進入時取得FuncCode 
    Protected m_sHoFlag As String = String.Empty '從左側進入時取得HoFlag總管理處(1) 區域中心(2) 單位(3) 外部單位(4) 海外單位(5) 系統功能(0) 
    Protected m_sSysId As String = String.Empty '從左側進入時取得系統別D(法金),F(消金),Z(系統管理),X(共同),..
    Protected m_sSubSysId As String = String.Empty '從左側進入時取得子系統別04(法金授信),05(消金授信),06(擔保品),00(共同),SY(系統管理)...

    Protected m_sTreeLevel As String = String.Empty
    ''' <summary>
    ''' 預設的DatabaseManager, 當網頁結束時會自行關閉m_dbManager內的Connection
    ''' 請使用getDatabaseManager()取得Instance of DatabaseManager
    ''' 儘量不直接使用m_dbManager.
    ''' </summary>
    ''' <remarks></remarks>
    Protected m_dbManager As DatabaseManager
    Protected m_EloanFlow As ELoanFlow

    Protected m_bDisableAllControlsInDisplayMode As Boolean = True

    Protected Shared export_UserInfo As ExportUserInfo
    Protected Shared export_FlowMail As CallbackNextstepSendmail


    ''' <summary>
    ''' Initial fundamental variables
    ''' </summary>
    ''' <remarks></remarks>
    Protected Sub InitParas()

        Response.Cache.SetCacheability(HttpCacheability.NoCache)

        m_sFlowName = com.Azion.EloanUtility.CodeList.getAppSettings(Me.GetType.BaseType.Name)
        m_sCaseId = UIUtility.getCaseID()
        m_sSubFlowSeq = UIUtility.getSubFlowSeq()
        m_sWorkingUserid = UIUtility.getWorkingUserId()       'will be changed to UIUtility.getWorkingUserId()  
        m_sLoginUserid = UIUtility.getLoginUserID()
        m_sLoginBrid = UIUtility.getLoginBrID()
        m_bTesting = UIUtility.isTesting()
        m_bDebug = UIUtility.isDebug()
        m_sStepNo = UIUtility.getStepno()
        m_sWatchStepNo = UIUtility.getWatchStepno
        m_bCheck = UIUtility.isCheck
        m_sSysId = com.Azion.EloanUtility.UIUtility.getSysId()
        m_sSubSysId = com.Azion.EloanUtility.UIUtility.getSubSysId()
        m_sHoFlag = com.Azion.EloanUtility.UIUtility.getHoFlag()
        m_sFuncCode = com.Azion.EloanUtility.UIUtility.getFuncCode()
        m_sTreeLevel = com.Azion.EloanUtility.CodeList.getAppSettings("treeLevel")

        If m_sWatchStepNo <> Nothing Then
            m_sStepNo = m_sWatchStepNo
            m_bDisplayMode = True
        End If

        ' Add by Avril 2012/04/18
        If Not Session("StaffInfo") Is Nothing AndAlso Session("StaffInfo").ToString() <> "" Then
            g_oUserInfo = CType(Session("StaffInfo"), StaffInfo)
        End If

        ' 如果Session不為空 則取得登陸信息
        If Not g_oUserInfo Is Nothing Then

            '目前登入使用者，內部系統userid(7碼)(S0XXXXX)(S000819)
            m_sLoginUserid = g_oUserInfo.LoginUserId

            '目前登入使用者，安泰userid(5碼)(XXXXX)(00819)
            m_sLoginEnTieUserid = Right(m_sLoginUserid, 5)

            '目前登入使用者的分行別(3碼)
            m_sLoginBrid = g_oUserInfo.LoginBrid

            '目前登入使用者最上層的部門序號
            m_sLoginTopDepNo = g_oUserInfo.LoginTopDepNo

            '目前登入使用者的角色
            m_sLoginRoleId = g_oUserInfo.LoginRoleId


            '目前代理使用者，內部系統userid(7碼)(S0XXXXX)(S000819)
            m_sWorkingUserid = g_oUserInfo.WorkingStaffid

            '目前代理使用者的分行別(3碼)
            m_sWorkingBrid = g_oUserInfo.WorkingBrid

            '目前代理使用者的部門序號
            m_sWorkingTopDepNo = g_oUserInfo.WorkingTopDepNo

            '目前代理使用者的角色
            m_sWorkingRoleId = g_oUserInfo.WorkingRoleId

            m_sWorkingEnTieUserid = Right(m_sWorkingUserid, 5)
            ''案件編號
            'm_sCaseId = "049251000000818"

            ' 專案目前步驟后四碼
            m_sFourStepNo = Right(m_sStepNo, 4)
        End If

        If String.IsNullOrEmpty(m_sCaseId) = False AndAlso String.IsNullOrEmpty(m_sFuncCode) = True Then
            m_sFuncCode = BosBase.CDbType(Of String)(BosBase.getNewBosBase("SY_CASEID", GetDatabaseManager).ExecuteScalar("select FUNCCODE from SY_CASEID where CASEID=@CASEID@", "CASEID", m_sCaseId), "")
        End If

        If IsNothing(export_UserInfo) Then
            export_UserInfo = New ExportUserInfo
            FLOW_OP.ELoanFlow.m_callbackUserInfo = export_UserInfo
        End If

        If IsNothing(export_FlowMail) Then
            export_FlowMail = New CallbackNextstepSendmail
            FLOW_OP.ELoanFlow.m_callbackFLowMail = export_FlowMail
        End If

    End Sub

    Protected Sub Page_PreInit(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.PreInit
        Page.MaintainScrollPositionOnPostBack = True

        If m_bTesting = False AndAlso Request("TESTMODE") = "1" Then
            m_bTesting = True
        End If

        m_bDisableAllControlsInDisplayMode = True

        InitParas()
        ' If Not Page.IsPostBack Then
        'Dim pageTab As New com.Azion.UITools.PageTitle(Me)
        ' End If
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        loadTempMode()
    End Sub

    Private Sub loadTempMode()

        'If Not IsPostBack And Not m_sFuncCode = String.Empty Then
        If Not IsPostBack AndAlso String.IsNullOrEmpty(m_sCaseId) Then
            ' 如果m_sStepNo 為空
            If m_sStepNo = String.Empty Then
                Dim syFlowStep As New AUTH_OP.SY_FLOWSTEP(GetDatabaseManager())
                ' 取出流程步驟編號
                If syFlowStep.loadRelSYTEMPINFO(m_sWorkingUserid, m_sWorkingTopDepNo, m_sFuncCode, "1") Then
                    m_sStepNo = syFlowStep.getString("STEP_NO")
                    m_sWatchStepNo = m_sStepNo
                    m_bDisplayMode = True '只要有起案，從左側列表進入，一定要設為true 

                    ViewState("m_sWatchStepNo") = m_sWatchStepNo
                    ViewState("m_bDisplayMode") = m_bDisplayMode

                End If
            End If
            'ElseIf IsPostBack And Not m_sFuncCode = String.Empty Then
        ElseIf IsPostBack AndAlso String.IsNullOrEmpty(m_sCaseId) Then
            If Not ViewState("m_sWatchStepNo") Is Nothing Then m_sWatchStepNo = CStr(ViewState("m_sWatchStepNo"))
            If Not ViewState("m_bDisplayMode") Is Nothing Then m_bDisplayMode = CBool(ViewState("m_bDisplayMode"))
            m_sStepNo = m_sWatchStepNo
        End If

        ' 取得所在區域的標識值
        If String.IsNullOrEmpty(m_sCaseId) = False Then
            Dim syfunctionCode As New AUTH_OP.SY_FUNCTION_CODE(GetDatabaseManager())
            ' 若能查詢出數據
            If syfunctionCode.getHoflag(m_sCaseId) Then
                m_sHoFlag = syfunctionCode.getString("HOFLAG")
            End If
        End If
    End Sub


    ''' <summary>
    ''' 發生於伺服器控制項從記憶體卸載時。
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub Page_Unload(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Unload
        Try
            FreeDatabaseManager()
        Catch ex As Exception
            Throw
        End Try
    End Sub 'close DB

    ''' <summary>
    ''' 在lblErrorMsg及ErrorMsg顯示錯誤訊息
    ''' </summary>
    ''' <param name="obj"></param>
    ''' <remarks></remarks>
    Protected Sub ShowErrMsg(ByVal obj As Object)
#If DEBUG Then
        Dim lblErrorMsg As Label = CType(Page.FindControl("lblErrorMsg"), System.Web.UI.WebControls.Label)
        Dim palErrorMsg As Panel = CType(Page.FindControl("ErrorMsg"), System.Web.UI.WebControls.Panel)

        Dbg.Assert(Not IsNothing(lblErrorMsg), "必須要有lblErrorMsg Label")
        Dbg.Assert(Not IsNothing(palErrorMsg), "必須要有ErrorMsg Panel")
#End If

        SYUIBase.showErrMsg(Me, obj)
    End Sub

    ''' <summary>
    ''' 當擲回未處理的例外狀況時發生。
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Friend Sub Page_Error(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Error
        Server.Transfer(com.Azion.EloanUtility.UIUtility.getRootPath & "/ErrorPage.aspx")
    End Sub

    ''' <summary>
    ''' 傳回DatabaseManager的物件
    ''' 如果已經有DatabaseManager，傳回已建立的DatabaseManager，
    ''' 如果沒有DatabaseManager，建立新Instance of DatabaseManager，再傳回 
    ''' </summary>
    ''' <value></value>
    ''' <returns>傳回DatabaseManager的物件</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property GetDatabaseManager() As DatabaseManager
        Get
            If IsNothing(m_dbManager) Then
                m_dbManager = MBSC.UICtl.UIShareFun.getDataBaseManager(True)
            End If

            Return m_dbManager
        End Get
    End Property


    ''' <summary>
    ''' 釋放預設的Database Manager
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub FreeDatabaseManager()
        If Not IsNothing(m_dbManager) Then
            m_dbManager.dispose()
            m_dbManager = Nothing
        End If
    End Sub

    'Protected Function GetEloanFlow() As ELoanFlow
    '    If IsNothing(m_ELoanFlow) Then
    '        m_ELoanFlow = New ELoanFlow(GetDatabaseManager)
    '    End If

    '    Return m_ELoanFlow
    'End Function

    'PreRender事件
    Private Sub Page_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.PreRender

        If m_bDisplayMode AndAlso m_bDisableAllControlsInDisplayMode Then
            ' com.Azion.EloanUtility.UIUtility.setControlRead(Me)
            Dim javaScript As New System.Text.StringBuilder()

            javaScript.AppendLine("<Script lang=""JavaScript"" src=""" & com.Azion.EloanUtility.UIUtility.getRootPath() & "/EN/js/ReadOnly.js""></Script>")
            javaScript.AppendLine("<script type='text/javascript'>")
            javaScript.Append("var sWatchStepNo =""" & m_sWatchStepNo & """;")
            javaScript.Append("var bDisplayMode =" & m_bDisplayMode.ToString.ToLower & ";")
            javaScript.AppendLine("window.document.body.onload = setControlReadOnly;")
            'javaScript.AppendLine("window.document.body.onunload = setControlReadOnly;")
            javaScript.AppendLine("</script>")

            Me.Page.ClientScript.RegisterClientScriptBlock(Me.GetType(), "BodyLoadUnloadScript", javaScript.ToString())
        End If
    End Sub


    Protected Sub CloseWindow()
        If String.IsNullOrEmpty(m_sStepNo) Then
            com.Azion.EloanUtility.UIUtility.goMainPage("")
        Else
            com.Azion.EloanUtility.UIUtility.closeWindow()
        End If
    End Sub


    Public Shared Function GetTagValueRecursive(ByVal sInputString As String, ByVal sTagString As String) As String
        Dim rgx As New Regex("<(?<tag>\w*)>(?<text>.*)</\k<tag>>", RegexOptions.IgnoreCase)
        Dim matches As MatchCollection = rgx.Matches(sInputString)

        For Each match As Match In matches
            Dim groups As GroupCollection = match.Groups

            '取得tag及text
            Dim tag = groups.Item("tag").Value
            Dim text = groups.Item("text").Value

            If String.Compare(tag, sTagString, True) = 0 Then
                Return text
            Else
                text = text.Trim
                If String.IsNullOrEmpty(text) = False Then
                    text = GetTagValueRecursive(text, sTagString)
                    If String.IsNullOrEmpty(text) = False Then
                        Return text
                    End If
                End If
            End If
        Next

        Return Nothing
    End Function


    Public Shared Sub showErrMsg(ByRef page As System.Web.UI.Page, ByVal obj As Object)
        Try

            Dim lblErrorMsg As Label = CType(page.FindControl("lblErrorMsg"), System.Web.UI.WebControls.Label)
            Dim palErrorMsg As Panel = CType(page.FindControl("ErrorMsg"), System.Web.UI.WebControls.Panel)

            Dim sErrorMsg As String = String.Empty

            If TypeOf obj Is Exception Then
                Dim exception As Exception = CType(obj, Exception)
                Dim sb As New System.Text.StringBuilder

                If com.Azion.EloanUtility.UIUtility.isDebug Then
                    Do
                        If ValidateUtility.isValidateData(exception.Message) Then
                            sb.Append(exception.Message.ToString & "<br><br>")
                        End If

                        If ValidateUtility.isValidateData(exception.StackTrace) Then
                            sb.Append(exception.StackTrace.ToString & "<br>")
                        End If

                        exception = exception.InnerException
                    Loop Until (exception Is Nothing)
                Else
                    sb.Append(exception.Message.ToString & "<br>")
                End If

                If TypeOf obj Is SYException Then

                    Dim syException As SYException = CType(obj, SYException)
                    sb.Append("<br>" & "錯誤碼：" & Hex(syException.Code) & "<br>")

                    If String.IsNullOrEmpty(syException.LastSQL) = False Then
                        sb.Append("<br>最後執行的SQL指令：<br>" & syException.LastSQL & "<br>")
                    End If
                Else
                    If String.IsNullOrEmpty(BosBase.GetGlobalLastSQL) = False Then
                        sb.Append("<br>最後執行的SQL指令：<br>" & BosBase.GetGlobalLastSQL & "<br>")
                    End If
                End If

                sErrorMsg = sb.ToString
            Else
                sErrorMsg = CType(obj, String)

                If String.IsNullOrEmpty(BosBase.GetGlobalLastSQL) = False Then
                    sErrorMsg &= "<br>最後執行的SQL指令：<br>" & BosBase.GetGlobalLastSQL & "<br>"
                End If
            End If

            sErrorMsg = sErrorMsg.Replace(vbCrLf, "<br>").Replace(" ", "&nbsp;")

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


    Protected Function DynamicSQL(ByVal sDynamicSQL As String, ByVal ParamArray objs() As String) As String
        Dim sb As New StringBuilder(sDynamicSQL)

        sb.Replace("#FLOWNAME#", m_sFlowName)
        sb.Replace("#CASEID#", m_sCaseId)
        sb.Replace("#WORKINGUSERID#", m_sWorkingUserid)
        sb.Replace("#LOGINUSERID#", m_sLoginUserid)
        sb.Replace("#WORKINGBRID#", m_sWorkingBrid)
        sb.Replace("#LOGINBRID#", m_sLoginBrid)
        sb.Replace("#STEPNO#", m_sStepNo)
        sb.Replace("#SYSID#", m_sSysId)
        sb.Replace("#SUBSYSID#", m_sSubSysId)
        sb.Replace("#HOFLAG#", m_sHoFlag)
        sb.Replace("#FUNCCODE#", m_sFuncCode)

        sb.Replace("#USERID#", m_sWorkingUserid)
        sb.Replace("#BRID#", m_sWorkingBrid)

        Dim arrayBosParameters As New BosParamsList

        Dbg.Assert((objs.Count Mod 2) = 0, "參數數目不正確")

        For i As Integer = 0 To UBound(objs, 1) Step 2
            sb.Replace(objs(i), objs(i + 1))
        Next i

        Return sb.ToString
    End Function

End Class

