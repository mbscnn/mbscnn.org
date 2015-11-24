Imports System.Threading.Tasks

''' <summary>
''' 程式說明：待處理工作
''' 建立者：Lake
''' 建立日期：2012-05-12
''' </summary>

Imports MBSC.UICtl
Imports AUTH_OP
Imports FLOW_OP

Public Class SY_CASELIST
    Inherits SYUIBase

    ' 每頁筆數
    Dim m_sPageSize As String = String.Empty
    Private m_sSortDirect As String = "DESC"
    Private m_sSortField As String = "STARTTIME" '排序欄位

#Region "PageLoad"
    ''' <summary>
    ''' 頁面加載
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>[Lake] 2012/05/12 Created</remarks>
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        ' 若頁面Session出現超時 則返回到登陸頁面
        monitorTimeOut()
        Response.AddHeader("cache-control", "private")
        Response.AddHeader("P3P", "CP='CAO PSA OUR'")
        Response.ExpiresAbsolute = Now.AddDays(-1)
        Response.Expires = 0 'prevents caching at the proxy server 

        ' 為分頁用戶控件指定關聯的GirdView和綁定數據的方法
        Me.PagerMenu1.SetTarget(Me.dgCase, New BindDataDelegate(AddressOf BindDataGrid))

        ' 初始化參數
        initParas()

        If Not IsPostBack Then
            ViewState("m_sSortField") = m_sSortField
            ViewState("m_sSortDirect") = m_sSortDirect
            ' 綁定案件資料
            BindDataGrid()
        End If
    End Sub
#End Region

    ''' <summary>
    '''  Session出現超時則需要重新登陸頁面
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub monitorTimeOut()
        Dim bTimeOut As Boolean = False

        ' 判斷Session是否為空
        If IsNothing(Session("StaffInfo")) Then
            bTimeOut = True
        End If

        ' 如果為空 則跳轉到登陸頁面重新登入
        If bTimeOut Then
            Dim sbTimeout As New System.Text.StringBuilder
            sbTimeout.Append("<SCRIPT LANGUAGE='JAVASCRIPT'>" & vbCrLf)
            sbTimeout.Append("alert('作業逾時,請重新登入!');" & vbCrLf)
            sbTimeout.Append("window.top.location = 'SY_DEFAULT.aspx'" & vbCrLf)
            sbTimeout.Append("</" & "SCRIPT>" & vbCrLf)
            Response.Write(HttpUtility.JavaScriptStringEncode(sbTimeout.ToString))
        End If
    End Sub

#Region "Function"

    ''' <summary>
    ''' 初始化參數
    ''' </summary>
    ''' <remarks>[Lake] 2012/05/12 Created</remarks>
    Sub initParas()

        ' 測試數據
        If m_bTesting OrElse Request("TESTMODE") = "1" Then
            m_bTesting = True
            Dim loginUser As New SY_LOGIN(GetDatabaseManager)
            loginUser.getUserInfo("912", "S000106")
        End If

        ' 取得每頁筆數
        m_sPageSize = com.Azion.EloanUtility.CodeList.getAppSettings("PageSize")
    End Sub

    ''' <summary>
    ''' 待處理案件數據綁定
    ''' </summary>
    ''' <remarks>[Lake] 2012/05/12 Created</remarks>
    Sub BindDataGrid()
        Try
            Dim flowFacade As New FlowFacade(GetDatabaseManager())
            Dim flowstep As New FLOW_OP.TABLE.SY_FLOWSTEP(GetDatabaseManager())


            Dim dt As DataTable
            Dim sStepNo As String

            'dt = flowFacade.GetPendingCaseList(m_sWorkingUserid, m_sWorkingBrid, 1, 9999999, "STARTTIME desc")
            dt = flowFacade.getSYFlowStep().GetCaselistByUseridStatus(
                m_sWorkingUserid,
                m_sWorkingBrid,
                1,
                1,
                999999,
                "STARTTIME desc",
                "FS.STEP_NO not like '0450%'"
                )

            Try
                Dim column As DataColumn
                column = dt.Columns.Add("PR_STEPNAME", System.Type.GetType("System.String"))
                column = dt.Columns.Add("PR_BRCNAME", System.Type.GetType("System.String"))
            Catch ex As Exception
            End Try

            For Each dr As DataRow In dt.Rows
                Try
                    Dim dr2 As DataRow
                    dr2 = flowstep.GetFormerFlowStepInfo(dr("CASEID"), dr("SUBFLOW_SEQ"))

                    dr.Item("PR_STEPNAME") = FLOW_OP.FlowFacade.CDbType(Of String)(dr2("STEP_NAME"), "")
                    dr.Item("PR_BRCNAME") = FLOW_OP.FlowFacade.CDbType(Of String)(dr2("BRCNAME"), "")
                Catch ex As Exception

                End Try
            Next

            ' 案件清單資料控件綁定數據
            Dim dv As DataView = dt.DefaultView
            dv.Sort = ViewState("m_sSortField") & " " & ViewState("m_sSortDirect")
            dgCase.DataSource = dv
            dgCase.DataBind()
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub
#End Region

#Region "Event"
    ''' <summary>
    ''' 點選案件編號
    ''' </summary>
    ''' <param name="source"></param>
    ''' <param name="e"></param>
    ''' <remarks>[Lake] 2012/05/12 Created</remarks>
    Protected Sub dgCase_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles dgCase.ItemCommand
        Try
            If e.Item.ItemType = ListItemType.Item OrElse e.Item.ItemType = ListItemType.AlternatingItem Then
                Dim syFlowMap As New AUTH_OP.SY_FLOW_MAP(GetDatabaseManager())
                Dim oCase As LinkButton = e.Item.FindControl("lnkbtnCaseId")
                Dim oFormUrl As HiddenField = e.Item.FindControl("hidFormUrl")
                Dim oStepNo As HiddenField = e.Item.FindControl("hidStepNo")
                Dim sUrl As String = String.Empty
                Dim oSubFlowSeq As HiddenField = e.Item.FindControl("hidSubFlowSeq")

                ' sUrl = oFormUrl.Value & "caseid=" & oCase.Text & "&stepno=" & oStepNo.Value '& "&MDISPLAYMODE=false"
                sUrl = com.Azion.EloanUtility.UIUtility.getRootPath & "/PageTab/FakeFrame.aspx?FinalPage=" & oFormUrl.Value & _
                        "&caseid=" & oCase.Text & _
                        "&stepno=" & oStepNo.Value & _
                        "&subflowseq=" & oSubFlowSeq.Value

                '并開啟SY_FLOW_MAP.FORMURL欄位所對應的頁面()
                Response.Write("<script language='javascript'>window.showModalDialog('" & sUrl & "','window','DialogHeight:900px;DialogWidth:1440px;scroll:yes;status:no');window.location=window.location;</script>")
            End If
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 案件資料綁定
    ''' 格式化【申請日期】、【收件日期】
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>[Lake] 2012/06/08 Created</remarks>
    Protected Sub dgCase_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles dgCase.ItemDataBound
        Try
            If e.Item.ItemType = ListItemType.AlternatingItem OrElse e.Item.ItemType = ListItemType.Item Then

                ' 申請日期
                Dim oCaseStTime As Label = e.Item.FindControl("lblCaseStTime")

                ' 收件日期
                Dim oStartTime As Label = e.Item.FindControl("lblStartTime")
                Dim iCaseStYear As Integer = oCaseStTime.Text.Substring(0, 4)
                Dim iStartYear As Integer = oStartTime.Text.Substring(0, 4)

                ' 格式化申請日期
                If iCaseStYear - 1911 > 0 Then
                    oCaseStTime.Text = (iCaseStYear - 1911).ToString & oCaseStTime.Text.Substring(4)
                ElseIf iCaseStYear - 1911 = 0 Then
                    oCaseStTime.Text = "00" & oCaseStTime.Text.Substring(4)
                Else
                    oCaseStTime.Text = "0" & (iCaseStYear - 1911).ToString & oCaseStTime.Text.Substring(4)
                End If

                ' 格式化收件日期
                If iStartYear - 1911 > 0 Then
                    oStartTime.Text = (iStartYear - 1911).ToString & oStartTime.Text.Substring(4)
                ElseIf iStartYear - 1911 = 0 Then
                    oStartTime.Text = "00" & oStartTime.Text.Substring(4)
                Else
                    oStartTime.Text = "0" & (iStartYear - 1911).ToString & oStartTime.Text.Substring(4)
                End If
            End If
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub
#End Region

    Protected Sub dgCase_SelectedIndexChanged(sender As Object, e As EventArgs) Handles dgCase.SelectedIndexChanged

    End Sub

    '下面是sort排序
    Sub sortCommand(ByVal Sender As Object, ByVal e As System.Web.UI.WebControls.DataGridSortCommandEventArgs)
        Try
            m_sSortField = e.SortExpression
            setSortCond(m_sSortField)
            '若是to_number大小排列則使用getdata
            BindDataGrid()
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Sub setSortCond(ByVal sSort As String)
        Try
            If sSort = viewstate("m_sSortField") Then
                If viewstate("m_sSortDirect") = "ASC" Then
                    setDirect("DESC")
                Else
                    setDirect("ASC")
                End If
            Else
                setDirect("ASC")
            End If
            viewstate("m_sSortField") = m_sSortField
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Private Sub setDirect(ByVal strDirect As String)
        Try
            m_sSortDirect = strDirect
            viewstate("m_sSortDirect") = strDirect
        Catch ex As Exception
            Throw
        End Try
    End Sub
End Class