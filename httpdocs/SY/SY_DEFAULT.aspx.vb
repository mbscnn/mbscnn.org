''' <summary>
''' 程式說明：開發使用者登入畫面
''' 建立者：Lake
''' 建立日期：2012-04-06
''' </summary>

Imports com.Azion.EloanUtility
Imports com.Azion.UITools
Imports AUTH_OP

Public Class SY_DEFAULT
    Inherits SY_LOGIN

#Region "PageLogin"

    ''' <summary>
    ''' 頁面加載
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    ''' <history>
    ''' [Lake] 2012/04/06 Created
    ''' </history>
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    End Sub
#End Region

#Region "Event"
    ''' <summary>
    ''' 登入系統
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    ''' <history>
    ''' [Lake] 2012/04/06 Created
    ''' </history>
    Protected Sub btnLogin_Click(sender As Object, e As EventArgs) Handles btnLogin.Click
        Try

            Dim sURL As String = String.Empty
            Dim sUserId As String = "S0"

            'If UIShareFun.eLoanConnectionString.IndexOf("192.168.8.118") <> -1 Then
            '    com.Azion.NET.VB.DatabaseManager.AVOID_TABLE_LOCKING = True
            'Else
            '    com.Azion.NET.VB.DatabaseManager.AVOID_TABLE_LOCKING = False
            'End If
            '1040506測試財報現金流量表Timeout問題
            com.Azion.NET.VB.DatabaseManager.AVOID_TABLE_LOCKING = False

            ' 檢查是否有輸入用戶編號
            If Not IsNothing(txtStaffId.Text) AndAlso txtStaffId.Text.Length > 0 Then
                'Dim syLogin As New SY_Login()
                '可參考文件，SY_DEFAULT繼承SY_LOGIN
                '取消，因為造成DB Connection無法釋放

                sUserId &= txtStaffId.Text

                ' 檢核用戶輸入編號是否存在
                If checkExsitUserId(sUserId) Then

                    Dim dt As DataTable = MBSC.UICtl.UIShareFun.getRoles(Me.GetDatabaseManager(), sUserId)
                    Dim dtTable As DataTable = dt.DefaultView.ToTable(True, New String() {"BRID"})

                    ' 查看登入者的組織資料,若存在多個部門，彈出“身份切換”頁面，用登入者選擇登入人員
                    If dtTable.Rows.Count > 1 Then
                        sURL = "SY_CHANGEUSER.aspx?SStaffId=" & sUserId
                        com.Azion.EloanUtility.UIUtility.showModalDialog(sURL, "300x800")
                    ElseIf dtTable.Rows.Count = 1 Then
                        getUserInfo(dtTable.Rows(0)("BRID").ToString(), sUserId, dt.Rows(0)("USERNAME").ToString())
                        com.Azion.EloanUtility.UIUtility.Redirect("SY_MAINFRAME.aspx")
                    End If
                Else
                    outError.InnerHtml = "使用者員工編號不正確，請重新輸入！"
                End If
            Else
                outError.InnerHtml = "請輸入使用者員工帳號！"
            End If
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub
#End Region
End Class