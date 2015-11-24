Imports com.Azion.EloanUtility

''' <summary>
''' 程式說明：個人資訊
''' 建立者：Lake
''' 建立日期：2012-05-07
''' </summary>

Imports AUTH_OP.TABLE
Imports AUTH_OP

Public Class SY_USERINFO
    Inherits SYUIBase

#Region "PageLoad"
    ''' <summary>
    ''' 頁面加載
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>[Lake] 2012/05/07 Created</remarks>
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If Not IsPostBack Then

            ' 初始化頁面數據
            initData()
        End If
    End Sub
#End Region

#Region "Function"
     

    ''' <summary>
    ''' 初始化頁面數據
    ''' </summary>
    ''' <remarks>[Lake] 2012/05/07 Created</remarks>
    Sub initData()
        Try

            Dim syUserList As New SY_USER(GetDatabaseManager())

            ' 取得個人資訊
            If syUserList.loadByPK(m_sWorkingUserid) Then
                lblJobTitle.Text = syUserList.getString("JOBTITLE")
                txtOfficeTel1.Text = syUserList.getString("OFFICETEL")
                txtEmail.Text = syUserList.getString("EMAIL")
            End If

            lblUserName.Text = g_oUserInfo.LoginUserName
            lblStaffId.Text = g_oUserInfo.LoginUserId
            lblBrcName.Text = g_oUserInfo.LoginBrCname
            
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub
#End Region

#Region "Event"
    ''' <summary>
    ''' 儲存變更
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>[Lake] 2012/05/07 Created</remarks>
    Protected Sub btnSaveChange_Click(sender As Object, e As EventArgs) Handles btnSaveChange.Click
        Try
            Dim syUser As New SY_USER(GetDatabaseManager())

            ' True,更新SY_USER表
            If syUser.loadByPK(m_sWorkingUserid) Then
                syUser.setAttribute("OFFICETEL", txtOfficeTel1.Text.Trim())
                syUser.setAttribute("EMAIL", txtEmail.Text.Trim())

                syUser.save()

                com.Azion.EloanUtility.UIUtility.closeWindow()
            End If
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 關閉視窗
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>[Lake] 2012/05/07 Created</remarks>
    Protected Sub btnCloseWin_Click(sender As Object, e As EventArgs) Handles btnCloseWin.Click
        Try
            com.Azion.EloanUtility.UIUtility.closeWindow()
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub
#End Region
End Class