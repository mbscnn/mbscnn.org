Imports com.Azion.UITools

''' <summary>
''' 程式說明：行內網登入
''' 建立者：Zack 
''' 建立日期：2012-06-07 
''' </summary>

Imports com.Azion.NET.VB
Imports FLOW_OP

Public Class SY_WEBAUTH
    Inherits SY_LOGIN

    ''' <summary>
    ''' 頁面加載事件
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            ' 實例化
            'Dim syLogin As New SY_LOGIN()
            Dim sUserID As String = String.Empty
            Dim sUserName As String = String.Empty

            ' 取得使用者編號了，使用者姓名
            GetCurrentUserInfo(sUserID, sUserName)
            'sUserID = "S000022"

            ' 檢核用戶編號是否存在
            If Not checkExsitUserId(sUserID) Then
                com.Azion.EloanUtility.UIUtility.alert("使用者員工編號不正確，無法登入此系統！")
                Return
            End If

            Dim dt As DataTable = MBSC.UICtl.UIShareFun.getRoles(Me.GetDatabaseManager(), sUserID)
            Dim dtTable As DataTable = dt.DefaultView.ToTable(True, New String() {"BRID"})


            ' 查看登入者的組織資料,若存在多個部門，彈出“身份切換”頁面，用登入者選擇登入人員
            'If dtTable.Rows.Count > 1 Then
            '    Dim sURL As String = "SY_CHANGEUSER.aspx?SStaffId=" & sUserID
            '    com.Azion.EloanUtility.UIUtility.showModalDialog(sURL, "500px", "300px")
            'ElseIf dtTable.Rows.Count = 1 Then
            '    getUserInfo(dtTable.Rows(0)("BRA_DEPNO").ToString(), sUserID, dtTable.Rows(0)("USERNAME").ToString())
            '    com.Azion.EloanUtility.UIUtility.Redirect("SY_MAINFRAME.aspx")
            'End If

            Dim sURL As String

            If dtTable.Rows.Count > 1 Then
                sURL = "SY_CHANGEUSER.aspx?SStaffId=" & sUserID
                com.Azion.EloanUtility.UIUtility.showModalDialog(sURL, "300x800")
            ElseIf dtTable.Rows.Count = 1 Then
                getUserInfo(dtTable.Rows(0)("BRID").ToString(), sUserID, dt.Rows(0)("USERNAME").ToString())
                com.Azion.EloanUtility.UIUtility.Redirect("SY_MAINFRAME.aspx")
            Else
                com.Azion.EloanUtility.UIUtility.alert("使用者沒有權限可以使用本系統!!")
                Dim js As String = "<Script Language='JAVAScript'>" & vbCrLf
                js += "self.close();" & vbCrLf
                js += "</Script>" & vbCrLf
                HttpContext.Current.Response.Write(js)

            End If

            '' 取得登入者相關信息
            'getUserInfo(sUserID)

            '' 進入功能頁面
            'com.Azion.EloanUtility.UIUtility.Redirect("SY_MAINFRAME.aspx")
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 取得使用者編號，姓名
    ''' </summary>
    ''' <param name="sUserID">用者編號</param>
    ''' <param name="sUserName">姓名</param>
    ''' <remarks></remarks>
    Sub GetCurrentUserInfo(ByRef sUserID As String, ByRef sUserName As String)
        Try
            ' 設置參數
            Dim sSessionID As String = Request("SessionID")

            '將Notes的SessionId寫入至Session內
            Dim currContext As System.Web.HttpContext = System.Web.HttpContext.Current
            currContext.Session("notesid") = sSessionID

            'If IsNothing(sSessionID) Then
            '    sSessionID = Session.SessionID
            'End If

            Dim sApSysNO As String = String.Empty

            'sSessionID = "002D5389773DADB148257A5B0009C5BF"

            Using dbManager As DatabaseManager = DatabaseManager.getInstance(Application("EmpDSN"))

                ' 聲明參數
                Dim sComdText As String = "sp_GetEmpIdBySessionId"
                Dim bosBase As New BosBase("Emf_SessionId", dbManager)
                Dim iDbCommand As IDbCommand = ProviderFactory.CreateCommand

                iDbCommand.Connection = dbManager.getConnection
                iDbCommand.CommandText = sComdText
                iDbCommand.CommandType = CommandType.StoredProcedure

                ' 設置存儲過程的參數
                Dim para As IDbDataParameter
                para = com.Azion.NET.VB.ProviderFactory.CreateDataParameter("SessionId", System.Data.DbType.String, sSessionID)
                para.Direction = ParameterDirection.Input
                iDbCommand.Parameters.Add(para)

                Dim iDataRead As IDataReader = iDbCommand.ExecuteReader

                ' 如果有查詢出數據
                If iDataRead.Read() Then
                    sUserID = "S0" & iDataRead("EmployNO")
                    sUserName = iDataRead("EmployName")
                Else
                    Throw New Exception("無法從行內網取得員工編號, [SessionId=" & sSessionID & "]")
                End If
            End Using
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub
End Class


