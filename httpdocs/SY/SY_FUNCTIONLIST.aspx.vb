Imports com.Azion.EloanUtility

''' <summary>
''' 程式說明：權限主畫面
''' 建立者：Lake
''' 建立日期：2012-04-06
''' </summary>
Imports com.Azion.NET.VB
Imports AUTH_OP
Imports AUTH_OP.TABLE
Imports System.IO

Public Class SY_FUNCTIONLIST
    Inherits SYUIBase

#Region " Page Load "

    ''' <summary>
    ''' 頁面加載
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load, Me.Load
        Try
            ' 若頁面Session出現超時 則返回到登陸頁面
            monitorTimeOut()
            Response.AddHeader("cache-control", "private")
            Response.AddHeader("P3P", "CP='CAO PSA OUR'")
            Response.ExpiresAbsolute = Now.AddDays(-1)
            Response.Expires = 0 'prevents caching at the proxy server 

            If Not IsPostBack Then
                ' 顯示TITLE
                showUserInfo()

                '取得列表清單
                bindTree()
            End If
        Catch ex As Exception
            SYUIBase.ShowErrMsg(Me, ex)
        Finally
        End Try
    End Sub
#End Region

#Region "Function "

    ''' <summary>
    ''' 目錄樹綁定
    ''' </summary>
    ''' <remarks></remarks>
    Sub bindTree()
        Dim sTempTable As DataTable
        treeViewSys.Nodes.Clear()

        Dim node As New TreeNode
        node.Value = ""
        node.Text = "系統功能"

        Dim funs As New SY_FUNCTION_CODEList(GetDatabaseManager())
        sTempTable = funs.genSYSFunList()

        Dim removeRowList As New List(Of DataRow)

        '441或461沒有代理人選單
        Dim item As DataRow
        If m_sLoginBrid = "441" OrElse m_sLoginBrid = "461" OrElse m_sLoginUserid <> m_sWorkingUserid Then
            For Each item In sTempTable.Rows
                If item("FUNCCODE") = 6 OrElse item("FUNCCODE") = 7 Then
                    removeRowList.Add(item)
                End If
            Next
        End If

        For Each item In removeRowList
            sTempTable.Rows.Remove(item)
        Next


        addNodes(node, 0, sTempTable, "", "")
        treeViewSys.Nodes.Add(node)

        ' 綁定子節點
        bindChildNode()
    End Sub

    ''' <summary>
    '''  組HTML顯示左邊的目錄樹
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Sub bindChildNode()
        Dim Subsysid As String = ""
        Dim hashTable As New Hashtable

        Dim dtTable As New DataTable
        Dim sTempTable As New DataTable
        Dim sb As New System.Text.StringBuilder
        Dim syRelRoleUserList As New SY_REL_ROLE_USERList(GetDatabaseManager())
        Dim sTempBraDepNo As String = String.Empty

        ' 若參數不為空 則系統編號=參數值
        If m_sSysId <> Nothing Then
            Subsysid = m_sSysId
        Else
            Subsysid = g_oUserInfo.SysId
        End If

        ' 根據父節點查詢 若沒有資料 則使用子節點的部門看看是否有資料
        If syRelRoleUserList.loadByParas(g_oUserInfo.WorkingStaffid, g_oUserInfo.WorkingBrid) > 0 Then

            ' 循環集合，組字串
            For Each row As DataRow In syRelRoleUserList.getCurrentDataSet.Tables(0).Rows
                sTempBraDepNo &= "'" & row("BRA_DEPNO").ToString & "'" & ","
            Next
        End If

        If sTempBraDepNo.Length > 0 Then
            sTempBraDepNo = sTempBraDepNo.Substring(0, sTempBraDepNo.Length - 1)
        End If

        Try
            If Not Subsysid Is Nothing Then
                hashTable = g_oUserInfo.FuncList

                ' 若hashtable中有資料，則取得 然後賦值
                ' 若沒有資料 則查詢 并將資料新增到hashtable中
                If Not hashTable Is Nothing Then
                    If hashTable.ContainsKey(Subsysid) Then
                        For Each item As DictionaryEntry In hashTable
                            If Subsysid = item.Key Then
                                dtTable = item.Value
                            End If
                        Next item
                    End If
                End If

                If Not dtTable Is Nothing Then
                    If dtTable.Rows.Count = 0 Then
                        Dim syLogin As New SY_LOGIN(GetDatabaseManager)
                        Dim funs As New SY_FUNCTION_CODEList(GetDatabaseManager())
                        sTempTable = funs.genFunList(g_oUserInfo.WorkingStaffid, g_oUserInfo.WorkingTopDepNo, Subsysid)

                        If sTempTable Is Nothing Then
                            sTempTable = funs.genFunListList(g_oUserInfo.WorkingStaffid, sTempBraDepNo, Subsysid)
                        End If

                        ' 以key value的形式將系統資料存入到hashtable中
                        If Subsysid <> "" Then
                            If g_oUserInfo.FuncList.ContainsKey(Subsysid) Then
                                g_oUserInfo.FuncList.Remove(Subsysid)
                            End If

                            If Not sTempTable Is Nothing Then
                                g_oUserInfo.FuncList.Add(Subsysid, sTempTable)
                            End If
                        End If

                        If Not sTempTable Is Nothing Then
                            dtTable = sTempTable.DefaultView.ToTable(True)
                        End If
                    End If
                End If

                ' 循環添加父節點
                If (Not dtTable Is Nothing) AndAlso (dtTable.Rows.Count > 0) Then

                    Dim dt As DataTable = dtTable.DefaultView.ToTable(True, New String() {"SYSID", "SUBSYSID", "SUBSYSNAME"})

                    For Each row In dt.Rows
                        Dim rootNode As New TreeNode
                        rootNode.Text = HttpUtility.HtmlEncode(row("SUBSYSNAME").ToString())
                        rootNode.Value = row("SUBSYSID").ToString
                        rootNode.ToolTip = rootNode.Text
                        addNodes(rootNode, 0, dtTable, row("SUBSYSID").ToString, row("SYSID").ToString)

                        ' 若沒有子節點 則無需新增到左邊交易樹
                        If rootNode.ChildNodes.Count > 0 Then
                            treeViewSys.Nodes.Add(rootNode)
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Function addNodes(ByRef tNode As TreeNode, ByVal PId As Integer, ByVal dt As DataTable, ByVal sSubSysId As String, ByVal sSysId As String) As String
        If Not dt Is Nothing Then

            '定義DataRow承接DataTable篩選的結果
            Dim rows() As DataRow

            '定義篩選的條件
            Dim filterExpr As String
            If sSubSysId = "" Then
                filterExpr = "ParentId = " & PId
            Else
                filterExpr = "ParentId = " & PId & " AND SUBSYSID = '" & sSubSysId & "'"
            End If

            '資料篩選並把結果傳入Rows
            rows = dt.Select(filterExpr)
            '如果篩選結果有資料
            If rows.GetUpperBound(0) >= 0 Then

                Dim childNode As TreeNode
                Dim rc As String

                '逐筆取出篩選後資料
                For i As Integer = 0 To rows.GetLength(0) - 1
                    '放入相關變數中
                    Dim row As DataRow = rows(i)

                    '實體化新節點
                    childNode = New TreeNode
                    '設定節點各屬性
                    childNode.Text = HttpUtility.HtmlEncode(row("FUNCTIONNAME").ToString)
                    childNode.ToolTip = childNode.Text
                    If row("FUCURL").ToString() <> Nothing Then
                        childNode.Value = HttpUtility.HtmlEncode(row("FUNCCODE").ToString() & ";" & row("parentid").ToString() & ";" & row("HOFLAG").ToString() & ";" & row("DISABLED").ToString() & ";" & row("SORTCTRL").ToString()) 'row("FUCURL").ToString()

                        Dim sPath As String = row("FUCURL").ToString
                        If sPath.IndexOf("?") = -1 Then
                            If Not sPath.EndsWith(".aspx", StringComparison.OrdinalIgnoreCase) Then
                                sPath &= ".aspx"
                            End If
                            sPath &= "?"
                        End If

                        If Not sPath.Split("?")(1).ToString.Trim = "" Then
                            If Not sPath.EndsWith("&", StringComparison.OrdinalIgnoreCase) Then
                                sPath &= "&"
                            End If
                        End If

                        childNode.NavigateUrl = HttpUtility.UrlPathEncode(com.Azion.EloanUtility.UIUtility.getRootPath & "/" & sPath & "funccode=" & row("FUNCCODE").ToString() & "&hoflag=" & row("HOFLAG").ToString() & IIf(sSubSysId <> Nothing, "&subsysid=" & sSubSysId, "") & "&sysid=" & sSysId)
                        childNode.Target = "mainFrame"

                        If childNode.Text = "登出系統" Then 'javascript:window.close();
                            If sPath.IndexOf("javascript") <> -1 Then
                                childNode.NavigateUrl = HttpUtility.JavaScriptStringEncode(row("FUCURL").ToString)
                            End If
                            childNode.Target = "_top"
                        End If
              
                    End If

                    If row("FUCURL").ToString = "" Then

                        If sSubSysId = "" Then
                            filterExpr = "ParentId = " & row("FUNCCODE").ToString
                        Else
                            filterExpr = "ParentId = " & row("FUNCCODE").ToString & " AND SUBSYSID = '" & sSubSysId & "'"
                        End If

                        If dt.Select(filterExpr).Count > 0 Then
                            ' 這樣的資料不需要添加

                            '將節點加入Tree中
                            tNode.ChildNodes.Add(childNode)

                            '呼叫遞回取得子節點
                            rc = addNodes(childNode, row("FUNCCODE"), dt, sSubSysId, sSysId)
                        End If
                    Else
                        '將節點加入Tree中
                        tNode.ChildNodes.Add(childNode)

                        '呼叫遞回取得子節點
                        rc = addNodes(childNode, row("FUNCCODE"), dt, sSubSysId, sSysId)
                    End If
                Next
            End If
        End If
    End Function

    ''' <summary>
    ''' 添加子項
    ''' </summary>
    ''' <param name="parentNode"></param>
    ''' <param name="dtTable"></param>
    ''' <remarks></remarks>
    'Sub addNode(ByVal parentNode As TreeNode, ByVal dtTable As DataTable)

    '    ' 循環添加子節點
    '    For i As Integer = 1 To dtTable.Rows.Count - 1
    '        If Trim(dtTable.Rows(i)("PARENT")) <> "0" Then
    '            Dim node As New TreeNode

    '            node.Text = dtTable.Rows(i)("FUNCTIONNAME").ToString()
    '            node.Value = dtTable.Rows(i)("FUCURL").ToString()

    '            ' 設置target
    '            node.Target = "mainFrame"
    '            node.NavigateUrl = com.Azion.EloanUtility.UIUtility.getRootPath & "/" & dtTable.Rows(i)("FUCURL").ToString()
    '            node.ImageUrl = "~/img/leftitem/left_icon.gif"

    '            ' 如果是最後一個節點 則調整圖片
    '            If i = dtTable.Rows.Count - 1 Then
    '                node.ImageUrl = "~/img/leftitem/left_endicon.gif"
    '            End If

    '            ' 添加到父節點中
    '            parentNode.ChildNodes.Add(node)
    '        End If
    '    Next
    'End Sub

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

    ''' <summary>
    ''' 設置Title
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub showUserInfo()
        Dim sTARGET_USERNAME As String = ""
        Dim sStaffid As String = g_oUserInfo.LoginUserId
        Dim sWorkingId As String = g_oUserInfo.WorkingStaffid
        Dim sUserName As String = g_oUserInfo.LoginUserName
        Dim sWorkingName As String = g_oUserInfo.WorkingName

        ' 若 AGENTUSERID 有值(且不等於 USERID), 則表示 有代理人
        If m_sWorkingUserid.Equals(m_sLoginUserid) Then

            If sUserName.Contains("(") Then
                sUserName = sUserName.Substring(sUserName.LastIndexOf("(") + 1, sUserName.LastIndexOf(")") - sUserName.LastIndexOf("(") - 1)
            End If

            sTARGET_USERNAME = HttpUtility.HtmlEncode(sUserName) & " 您好!"
        Else

            If sUserName.Contains("(") Then
                sUserName = sUserName.Substring(sUserName.LastIndexOf("(") + 1, sUserName.LastIndexOf(")") - sUserName.LastIndexOf("(") - 1)
            End If

            sTARGET_USERNAME = HttpUtility.HtmlEncode(sUserName) & " 您好!" & "<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color=red>(代理 " & HttpUtility.HtmlEncode(sWorkingName) & ")</font>"
        End If

        infomation.InnerHtml = sTARGET_USERNAME
    End Sub
#End Region
End Class