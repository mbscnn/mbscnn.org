''' <summary>
''' 程式說明：切換身分
''' 建立者：Avril
''' 建立日期：2012-06-11
''' </summary>
Imports com.Azion.NET.VB
Imports AUTH_OP
Imports AUTH_OP.TABLE
Imports Eloan.EN_OP
Imports System.Xml
Imports System.IO
Imports FLOW_OP
Imports com.Azion.EloanUtility

Public Class SY_CHANGEUSER
    Inherits SYUIBase
     
    ''' <summary>
    ''' 頁面載入
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        ' 頁面第一次加載
        If Not IsPostBack Then
            initData()
        End If
    End Sub

    ''' <summary>
    ''' 資料加載
    ''' </summary>
    ''' <remarks></remarks>
    Sub initData()
        Dim syUserRoleAgentList As New SY_USERROLEAGENTList(GetDatabaseManager())
        Dim dtData As New DataTable
        Dim sBrId As String = String.Empty
        'Dim sBraDepNo As String = String.Empty
        Dim sStaffId As String = String.Empty
        Dim sUserName As String = String.Empty
        Dim sStartTime As String = String.Empty
        Dim sEndTime As String = String.Empty
        Dim sTitle As String = String.Empty
        Dim sStaffIdRequest As String = String.Empty
        Dim sStaffidParm As String = String.Empty
        Dim sAgentName As String = String.Empty

        sStaffIdRequest = Request("SStaffId")

        If sStaffIdRequest = "" Then

            ' 取得Session的值
            g_oUserInfo = Session("StaffInfo")

            ' 暫存登入者信息
            ViewState("StaffId") = g_oUserInfo.LoginUserId
            ViewState("UserName") = g_oUserInfo.LoginUserName
            'ViewState("BridList") = g_oUserInfo.BridList
            ViewState("BrId") = g_oUserInfo.LoginBrid
            'ViewState("Bra_DepNo") = g_oUserInfo.Bra_DepNo
            ViewState("BrCname") = g_oUserInfo.LoginBrCname
            ViewState("SysId") = g_oUserInfo.SysId

            'hidUserData.Value &= g_oUserInfo.WorkingBrid & "*" & g_oUserInfo.WorkingDepNo & "*" & g_oUserInfo.WorkingId & "*"
            'hidUserData.Value &= g_oUserInfo.WorkingBrid & "*" & g_oUserInfo.WorkingId & "*"
        End If
       
        If sStaffIdRequest <> "" Then
            sStaffidParm = sStaffIdRequest
        Else
            sStaffidParm = g_oUserInfo.LoginUserId
        End If

        ' 查詢相關的代理人資料
        If syUserRoleAgentList.loadDataByCon(sStaffidParm) Then
            dtData = syUserRoleAgentList.getCurrentDataSet.Tables(0)
        End If

        GridView1.DataSource = dtData
        GridView1.DataBind()
        '' 循環代理人資料，依次新增到radiobuttinlist中
        'If Not dtData Is Nothing Then
        '    If dtData.Rows.Count > 0 Then
        '        For i = 0 To dtData.Rows.Count - 1
        '            sBrId = dtData.Rows(i)("BRID").ToString()
        '            'sBraDepNo = dtData.Rows(i)("BRA_DEPNO").ToString()
        '            sStaffId = dtData.Rows(i)("STAFFID").ToString()
        '            sUserName = dtData.Rows(i)("Name").ToString()
        '            sAgentName = dtData.Rows(i)("AGENTNAME").ToString

        '            If Not dtData.Rows(i)("STARTTIME") Is Nothing Then
        '                If dtData.Rows(i)("STARTTIME").ToString() <> "" Then
        '                    sStartTime = Convert.ToDateTime(dtData.Rows(i)("STARTTIME").ToString()).AddYears(-1911).ToString("yyy/MM/dd HH:mm")
        '                End If
        '            End If

        '            If Not dtData.Rows(i)("ENDTIME") Is Nothing Then
        '                If dtData.Rows(i)("ENDTIME").ToString() <> "" Then
        '                    sEndTime = Convert.ToDateTime(dtData.Rows(i)("ENDTIME").ToString()).AddYears(-1911).ToString("yyy/MM/dd HH:mm")
        '                End If
        '            End If

        '            ' 自己
        '            If dtData.Rows(i)("STARTTIME").ToString() = "" Then
        '                'rdoListChangeUser.Items.Insert(i, New ListItem(sUserName, sBrId & "*" & sBraDepNo & "*" &
        '                '                                               sStaffId & "*" & sUserName))
        '                rdoListChangeUser.Items.Insert(i, New ListItem(sUserName, sBrId & "*" &
        '                                                               sStaffId & "*" & sUserName & "*" & "AgentN"))
        '                ViewState("StaffId") = sStaffId
        '                ViewState("BrId") = sBrId

        '                g_oUserInfo.UserName = sUserName
        '            Else
        '                rdoListChangeUser.Items.Insert(i, New ListItem(sUserName & sStartTime & "~" & sEndTime & ")",
        '                                                               dtData.Rows(i)("AGENT_BRID").ToString() & "*" & dtData.Rows(i)("AGENT_STAFFID").ToString() & "*" & sAgentName & "*" & "AgentY"))
        '            End If
        '        Next
        '    End If
        'End If

        '' 設置初始值默認選中
        'For m As Integer = 0 To rdoListChangeUser.Items.Count - 1
        '    Dim sTempData As String = String.Empty
        '    For i As Integer = 0 To rdoListChangeUser.Items(m).Value.Split("*").Count
        '        sTempData &= rdoListChangeUser.Items(m).Value.Split("*")(i) & "*"
        '        If i = 2 Then
        '            Exit For
        '        End If
        '    Next

        '    If hidUserData.Value <> "" Then
        '        If hidUserData.Value = sTempData Then
        '            rdoListChangeUser.Items(m).Selected = True
        '        End If
        '    End If
        'Next
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound

        If (e.Row.RowType = DataControlRowType.DataRow) Then
            Dim rb As RadioButton = e.Row.FindControl("rdbLoginStaffid")
            If Not rb Is Nothing Then
                Dim sLoginStaffId As String = e.Row.DataItem("LoginStaffId")
                Dim sLoginBrid As String = e.Row.DataItem("LoginBrId")
                Dim sLoginUserName As String = e.Row.DataItem("LoginUserName")

                Dim sWorkingStaffid As String = e.Row.DataItem("WorkingStaffid")
                Dim sWorkingBrid As String = e.Row.DataItem("WorkingBrid")
                Dim sWorkingUserName As String = e.Row.DataItem("WorkingUserName")
                rb.Attributes("CommandArgument") = sLoginStaffId & "/" & sLoginBrid & "/" & sLoginUserName & "/" & sWorkingStaffid & "/" & sWorkingBrid & "/" & sWorkingUserName
            End If

        End If
         

    End Sub

    Sub select_RowCommand(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim sCommandArgument As String = DirectCast(sender, System.Web.UI.WebControls.RadioButton).Attributes("CommandArgument")

        Dim sLoginStaffId As String = sCommandArgument.Split("/")(0)
        Dim sLoginBrid As String = sCommandArgument.Split("/")(1)
        Dim sLoginUserName As String = sCommandArgument.Split("/")(2)

        Dim sWorkingStaffid As String = sCommandArgument.Split("/")(3)
        Dim sWorkingBrid As String = sCommandArgument.Split("/")(4)
        Dim sWorkingUserName As String = sCommandArgument.Split("/")(5)

        Dim syLogin As New SY_LOGIN(GetDatabaseManager)

        If sWorkingStaffid = "" Then
            syLogin.getUserInfo(sLoginBrid, sLoginStaffId, sLoginUserName)
        Else
            syLogin.getUserInfo(sLoginBrid, sLoginStaffId, sLoginUserName, sWorkingBrid, sWorkingStaffid, sWorkingUserName)
        End If

        If Not Request.QueryString("SStaffId") Is Nothing Then
            Response.Write("<script>window.close();window.dialogArguments.location.href='SY_MAINFRAME.aspx';</script>")
        Else
            Response.Write("<script>window.close();</script>")
        End If
    End Sub

    ''' <summary>
    ''' 切換身分
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    'Protected Sub rdoListChangeUser_SelectedIndexChanged(sender As Object, e As EventArgs) Handles rdoListChangeUser.SelectedIndexChanged
    '    Dim syRelRoleUserList As New AUTH_OP.SY_REL_ROLE_USERList(GetDatabaseManager())
    '    Dim sySysIdList As New AUTH_OP.SY_SYSIDList(GetDatabaseManager())
    '    Dim funs As New SY_FUNCTION_CODEList(GetDatabaseManager())
    '    Dim sTempTable As New DataTable
    '    Dim sSysId As String = String.Empty
    '    Dim sSubsysid As String = String.Empty
    '    Dim sWorkingRoleId As String = String.Empty

    '    '是否為代理人Flag
    '    Dim sAgentFlag As String = String.Empty

    '    'sBrId & "*" & sStaffId & "*" & sUserName & "*" & "AgentN"
    '    '代理人, Owner, login user
    '    Dim sLoginStaffId As String = rdoListChangeUser.SelectedValue.Split("*")(1).ToString()
    '    Dim sLoginBrid As String = rdoListChangeUser.SelectedValue.Split("*")(0).ToString()
    '    Dim sLoginName As String = rdoListChangeUser.SelectedValue.Split("*")(2).ToString()

    '    '被代理人, Client
    '    Dim sWorkingId As String = String.Empty
    '    Dim sWorkingBrid As String = String.Empty
    '    Dim sWorkingName As String = String.Empty

    '    sAgentFlag = rdoListChangeUser.SelectedValue.Split("*")(3).ToString()

    '    '重新初始StaffInfo object
    '    Dim syLogin As New SY_LOGIN(GetDatabaseManager)
    '    ' 勾選資料並非代理人員
    '    If (sAgentFlag = "AgentN") Then 
    '        syLogin.getUserInfo(sLoginBrid, sLoginStaffId, sLoginName)
    '    Else
    '        '0# 1 #    2     #   3   #   4  #
    '        '1#904#資訊管理部#S003413#鍾志宏#
    '        sWorkingBrid = rdoListChangeUser.SelectedValue.Split("*")(2).Split("#")(1)
    '        sWorkingId = rdoListChangeUser.SelectedValue.Split("*")(2).Split("#")(3)
    '        sWorkingName = rdoListChangeUser.SelectedValue.Split("*")(2).Split("#")(4)
    '        syLogin.getUserInfo(sLoginBrid, sLoginStaffId, sLoginName, sWorkingBrid, sWorkingId, sWorkingName)
    '    End If 

    '   If Not Request.QueryString("SStaffId") Is Nothing Then
    '            Response.Write("<script>window.close();window.dialogArguments.location.href='SY_MAINFRAME.aspx';</script>")
    '            Session("StaffInfo") = g_oUserInfo
    '    Else
    '        Response.Write("<script>window.close();</script>")
    '        Session("StaffInfo") = g_oUserInfo
    '    End If


    'End Sub

End Class