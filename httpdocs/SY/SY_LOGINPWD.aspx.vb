Imports com.Azion.UITools

''' <summary>
''' 程式說明：外部使用者登入畫面
''' 建立者：Zack 
''' 建立日期：2012-06-05
''' </summary>

Imports AUTH_OP.TABLE
Imports com.Azion.EloanUtility
Imports AUTH_OP
Imports com.Azion.NET.VB

Public Class SY_LOGINPWD
    Inherits SY_LOGIN

    ' 聲明參數
    Dim m_sHoflag As String = String.Empty  ' 屬性的標示值
    Dim m_sMaxTimes_WrongPassword As String = String.Empty  ' 密碼可輸入的錯誤次數
    Dim m_sPasswordExpirationDays As String = String.Empty  ' 密碼過期天數
    Dim m_sMinPasswordLength As String = String.Empty   ' 所輸入密碼的最小長度
    Dim m_sMaxPasswordLength As String = String.Empty  ' 所輸入密碼的最大長度

#Region "Page_load"

    ''' <summary>
    ''' 頁面初始化事件
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            ' 初始化頁面參數
            initParas()

            If Not IsPostBack Then

                ' 初始化頁面數據
                initData()
            End If

            ' 設置頁面控件的初始狀態
            If m_bDisplayMode Then

                ' 只讀模式
                com.Azion.EloanUtility.UIUtility.setControlRead(Me)
            End If
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub
#End Region

#Region "Function"

    ''' <summary>
    ''' 初始化參數
    ''' </summary>
    ''' <remarks></remarks>
    Sub initParas()

        ' 取得參數
        If Request.QueryString("hoflag") <> Nothing Then
            m_sHoflag = Request.QueryString("hoflag")
        End If

        m_sMaxTimes_WrongPassword = com.Azion.EloanUtility.CodeList.getAppSettings("MaxTimes_WrongPassword")
        m_sPasswordExpirationDays = com.Azion.EloanUtility.CodeList.getAppSettings("PasswordExpirationDays")
        m_sMinPasswordLength = com.Azion.EloanUtility.CodeList.getAppSettings("MinPasswordLength")
        m_sMaxPasswordLength = com.Azion.EloanUtility.CodeList.getAppSettings("MaxPasswordLength")

        ' 測試數據
        If m_bTesting Then
            m_sWorkingUserid = "S000001"
            m_bDisplayMode = False
            m_sHoflag = "1"
        End If
    End Sub

    ''' <summary>
    ''' 初始化頁面
    ''' </summary>
    ''' <remarks></remarks>
    Sub initData()
        Try
            ' 如果沒有傳入參數,顯示登入畫面
            If m_sHoflag = "" Then
                divloginPassWord.Style.Value = "display:block"
                divEditTwoPassWord.Style.Value = "display:none"
                divEditThridPassWord.Style.Value = "display:none"

                ' 設置初始焦點
                txtStaffId.Focus()
            Else
                Dim syUserPassword As New SY_USERPASSWORD(GetDatabaseManager())

                divEditThridPassWord.Style.Value = "display:block"
                divloginPassWord.Style.Value = "display:none"
                divEditTwoPassWord.Style.Value = "display:none"

                ' 獲得當前登錄者的編號
                If m_sHoflag = "4" Then
                    txtEditThridStaffId.Text = m_sWorkingUserid
                Else
                    txtEditThridStaffId.Text = ""
                End If

                ' 取得當前人的Mail
                If syUserPassword.loadByPK(m_sWorkingUserid) Then
                    txtMail.Text = syUserPassword.getAttribute("MAILADDR")
                End If

                ' 設置焦點
                txtEditThridStaffId.Focus()

                If m_sHoflag = "4" Then

                    ' 使用者員工編號不可編輯
                    txtEditThridStaffId.Enabled = False

                    ' 設施初始焦點
                    txtEditThridOPWD.Focus()
                ElseIf m_sHoflag = "3" Or m_sHoflag = "5" Then
                    com.Azion.EloanUtility.UIUtility.alert("無使用此交易權限！")

                    ' 頁面不可編輯
                    com.Azion.EloanUtility.UIUtility.setControlRead(Me)
                End If

                If m_sHoflag = "1" Then
                    txtEditThridOPWD.Enabled = False
                End If


            End If
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 計算字串的MD5值
    ''' </summary>
    ''' <param name="input">要計算MD5的原始字串</param>
    ''' <returns>傳回MD5值，以字串的形式傳回</returns>
    ''' <remarks></remarks>
    Public Function CreateMD5Hash(ByVal input As String) As String
        ' 創建MD5加密對象
        Dim md5 As System.Security.Cryptography.MD5 = System.Security.Cryptography.MD5.Create()
        Dim inputBytes As Byte() = System.Text.Encoding.ASCII.GetBytes(input)
        Dim hashBytes As Byte() = md5.ComputeHash(inputBytes)

        Dim sb As New StringBuilder()
        For i As Integer = 0 To hashBytes.Length - 1
            sb.Append(hashBytes(i).ToString("X2"))
        Next

        Return sb.ToString()
    End Function

    ''' <summary>
    ''' 檢查是否違反密碼原則
    ''' </summary>
    ''' <param name="sPassWord">輸入的新密碼</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function VerifyPasswordPolicy(ByVal sPassWord As String) As String
        Dim chrPassWord As Char()
        chrPassWord = sPassWord.ToCharArray()

        ' 檢核密碼長度
        If sPassWord.Length() < m_sMinPasswordLength OrElse sPassWord.Length() > m_sMaxPasswordLength Then
            Return "提示密碼不符合規則，請重新輸入！"
        End If

        ' 檢核密碼組合規則
        For nIndex As Integer = 0 To chrPassWord.Length - 2
            ' 當前字符，下一個字符
            Dim cCurr, cNext As Char

            cCurr = chrPassWord(nIndex)
            cNext = chrPassWord(nIndex + 1)

            '是否是數字
            If cCurr >= "0" AndAlso cCurr <= "9" AndAlso cNext >= "0" AndAlso cNext <= "9" Then

                '相連二數字不可相同
                If cCurr = cNext Then
                    Return "相連兩個數字不可相同！"
                End If

                '密碼數字部分不可逐次遞增或遞減
                If Asc(cCurr) = Asc(cNext) + 1 OrElse Asc(cCurr) = Asc(cNext) - 1 Then
                    Return "密碼數字部份不可逐次遞增或遞減！"
                End If
            End If
        Next

        Return ""
    End Function
#End Region

#Region "Event"

    ''' <summary>
    ''' 點擊"登入系統"按鈕事件
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnLogin_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnLogin.Click
        Try
            ' 實例化
            Dim syUserPassWord As New SY_USERPASSWORD(GetDatabaseManager())

            Dim sPassword As String = String.Empty

            ' 密碼是否正確
            Dim bCorrectPassword As Boolean = False
            Dim sUserId As String = "S0"

            sUserId &= txtStaffId.Text

            ' 如果當前用戶編號不存在
            If Not checkExsitUserId(sUserId) Then
                lblSystemMsg.Text = "使用者員工編號不正確，請重新輸入！"

                Return
            End If

            If syUserPassWord.loadByPK(sUserId) Then

                ' 使用者是否失效，(登入時輸入錯誤的密碼次數五次或超過五次)
                If Convert.ToInt32(syUserPassWord.getAttribute("PWD_WRONG_TIMES").ToString()) >= Convert.ToInt32(m_sMaxTimes_WrongPassword) Then
                    lblSystemMsg.Text = "帳號已失效！"

                    Return
                End If

                ' 檢查密碼是否跟資料庫儲存的相同
                sPassword = CreateMD5Hash(txtPWD.Text.Trim)
                If String.Compare(sPassword, syUserPassWord.getAttribute("PASSWORD")) = 0 Then

                    ' 密碼相同
                    bCorrectPassword = True

                    ' 重置錯誤登入次數
                    syUserPassWord.updatePWD_WRONG_TIMES(sUserId, 0)
                Else

                    ' 密碼不相同
                    bCorrectPassword = False

                    ' 登入次數+1
                    Dim iwrongTimes As Integer = Convert.ToInt32(syUserPassWord.getAttribute("PWD_WRONG_TIMES").ToString()) + 1

                    ' 修改登錄次數
                    syUserPassWord.updatePWD_WRONG_TIMES(sUserId, iwrongTimes)

                    lblSystemMsg.Text = "密碼輸入錯誤，請重新輸入！"

                    Return
                End If

                ' 使用者密碼是否過期
                If bCorrectPassword = True AndAlso _
                    DateDiff(DateInterval.Day, Convert.ToDateTime(syUserPassWord.getAttribute("PWD_CHANGE_DATE").ToString()), DateAndTime.Today) > m_sPasswordExpirationDays Then

                    divEditTwoPassWord.Style.Value = "display:block"
                    divEditThridPassWord.Style.Value = "display:none"
                    divloginPassWord.Style.Value = "display:none"

                    ' 設置頁面加載時的使用者，且不可編輯
                    txtEditTwoStaffId.Text = txtStaffId.Text
                    txtEditTwoStaffId.Enabled = False
                Else

                    Dim dt As DataTable = MBSC.UICtl.UIShareFun.getRoles(Me.GetDatabaseManager(), sUserId)
                    Dim dtTable As DataTable = dt.DefaultView.ToTable(True, New String() {"BRID"})
                    Dim sURL As String

                    ' 查看登入者的組織資料,若存在多個部門，彈出“身份切換”頁面，用登入者選擇登入人員
                    'If dtTable.Rows.Count > 1 Then
                    '    sURL = "SY_CHANGEUSER.aspx?SStaffId=" & sUserId
                    '    com.Azion.EloanUtility.UIUtility.showModalDialog(sURL, "500px", "300px")
                    'ElseIf dtTable.Rows.Count = 1 Then
                    '    com.Azion.EloanUtility.UIUtility.Redirect("SY_MAINFRAME.aspx")
                    '    getUserInfo(dtTable.Rows(0)("BRA_DEPNO").ToString(), sUserId, dtTable.Rows(0)("USERNAME").ToString())
                    'End If

                    ' 查看登入者的組織資料,若存在多個部門，彈出“身份切換”頁面，用登入者選擇登入人員
                    If dtTable.Rows.Count > 1 Then
                        sURL = "SY_CHANGEUSER.aspx?SStaffId=" & sUserId
                        com.Azion.EloanUtility.UIUtility.showModalDialog(sURL, "300x800")
                    ElseIf dtTable.Rows.Count = 1 Then
                        getUserInfo(dtTable.Rows(0)("BRID").ToString(), sUserId, dt.Rows(0)("USERNAME").ToString())
                        com.Azion.EloanUtility.UIUtility.Redirect("SY_MAINFRAME.aspx")
                    End If


                    '' 取得登入者相關信息
                    'getUserInfo(txtStaffId.Text.Trim, "")

                    '' 開啟主頁面
                    'com.Azion.EloanUtility.UIUtility.Redirect("SY_MAINFRAME.aspx")
                End If
            End If
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 點擊修改密碼頁面的"存儲"按鈕事件
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnEidtTwoOK_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnEidtTwoOK.Click
        Try
            ' 實例化
            Dim syUserPassWord As New SY_USERPASSWORD(GetDatabaseManager())

            ' 經MD5加密后的密碼
            Dim sMD5PWD As String = String.Empty
            Dim sUserId As String = "S0"
            sUserId &= txtStaffId.Text.Trim

            ' 對密碼進行MD5加密
            sMD5PWD = CreateMD5Hash(txtEditTwoOPWD.Text.Trim)

            ' 如果能查詢到人員 信息
            If syUserPassWord.loadByPK(txtEditTwoStaffId.Text.Trim) Then

                ' 如果輸入的舊密碼錯誤
                If String.Compare(sMD5PWD, syUserPassWord.getAttribute("PASSWORD")) <> 0 Then
                    lblEditTwoSystemMsg.Text = "舊密碼輸入錯誤，請重新輸入！"

                    Return
                End If
            End If

            ' 如果輸入的新密碼和確認密碼不相同
            If String.Compare(txtEditTwoNewPWD.Text.Trim, txtEditTwoCheckPWD.Text.Trim) <> 0 Then
                lblEditTwoSystemMsg.Text = "密碼不一致！"

                Return
            End If

            ' 如果輸入的密碼符合規則
            If VerifyPasswordPolicy(txtEditTwoNewPWD.Text.Trim) = "" Then

                ' 對新密碼進行MD5加密
                Dim sNewPWD As String = CreateMD5Hash(txtEditTwoNewPWD.Text.Trim)

                ' 更新密碼
                syUserPassWord.updatePWDDATETIME(sUserId, sNewPWD)

                Dim dt As DataTable = MBSC.UICtl.UIShareFun.getRoles(Me.GetDatabaseManager(), sUserId)
                Dim dtTable As DataTable = dt.DefaultView.ToTable(True, New String() {"BRID"})

                ' 查看登入者的組織資料,若存在多個部門，彈出“身份切換”頁面，用登入者選擇登入人員

                Dim sURL As String
                If dtTable.Rows.Count > 1 Then
                    sURL = "SY_CHANGEUSER.aspx?SStaffId=" & sUserId
                    com.Azion.EloanUtility.UIUtility.showModalDialog(sURL, "300x800")
                ElseIf dtTable.Rows.Count = 1 Then
                    getUserInfo(dtTable.Rows(0)("BRID").ToString(), sUserId, dt.Rows(0)("USERNAME").ToString())
                    com.Azion.EloanUtility.UIUtility.Redirect("SY_MAINFRAME.aspx")
                End If

                '' 取得登入者相關信息
                'syLogin.getUserInfo(txtEditTwoStaffId.Text.Trim)

                '' 跳轉頁面
                'com.Azion.EloanUtility.UIUtility.Redirect("SY_MAINFRAME.aspx")
            Else
                lblEditTwoSystemMsg.Text = VerifyPasswordPolicy(txtEditTwoNewPWD.Text.Trim)
            End If
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 點擊"資訊"頁面的"存儲"按鈕事件
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnEditThrid_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnEditThrid.Click
        Try

            ' 實例化
            Dim syUserPassWord As New SY_USERPASSWORD(GetDatabaseManager())
            Dim sMd5Password As String = String.Empty
            Dim sStaffid As String = "S0" & txtEditThridStaffId.Text.Trim

            ' 檢核員工編號是否存在
            If m_sHoflag = "1" Or m_sHoflag = "2" Then

                ' 若沒有當前員工信息
                If Not syUserPassWord.loadByPK(sStaffid) Then

                    com.Azion.EloanUtility.UIUtility.alert("使用者員工編號不正確，請重新輸入！")

                    Return
                End If
            End If

            ' 如果有輸入新密碼和舊密碼
            If txtEditThridNewPWD.Text <> "" And txtEditThridCheckPWD.Text <> "" Then

                ' 如果新密碼和確認密碼不相同
                If txtEditThridNewPWD.Text.Trim <> txtEditThridCheckPWD.Text.Trim Then

                    com.Azion.EloanUtility.UIUtility.alert("密碼不一致！")

                    Return
                End If

                ' 如果修改的不是自己的密碼或者修改其他人的密碼
                If (sStaffid = m_sWorkingUserid AndAlso "1,2,4".Contains(m_sHoflag)) OrElse m_sHoflag = "1" Then

                    If m_sHoflag <> "1" AndAlso m_sHoflag <> "2" Then
                        ' 如果沒有找到當前人員信息
                        If syUserPassWord.loadByPK(sStaffid) Then
                            sMd5Password = CreateMD5Hash(txtEditThridOPWD.Text.Trim)

                            ' 如果輸入的舊密碼不正確
                            If String.Compare(syUserPassWord.getAttribute("PASSWORD"), sMd5Password) <> 0 Then

                                com.Azion.EloanUtility.UIUtility.alert("舊密碼輸入錯誤，請重新輸入！")

                                Return
                            End If
                        End If
                    End If

                    ' 如果新密碼符合規則
                    If VerifyPasswordPolicy(txtEditThridNewPWD.Text.Trim) = "" Then

                        ' 對新密碼進行MD5加密
                        Dim sNewPassWord As String = CreateMD5Hash(txtEditThridNewPWD.Text.Trim)

                        If Not syUserPassWord.loadByPK(sStaffid) Then
                            syUserPassWord.setAttribute("STAFFID", sStaffid)
                        End If

                        syUserPassWord.setAttribute("PASSWORD", sNewPassWord)
                        syUserPassWord.setAttribute("MAILADDR", txtMail.Text.Trim)
                        syUserPassWord.setAttribute("PWD_CHANGE_DATE", Date.Now)
                        syUserPassWord.setAttribute("PWD_WRONG_TIMES", 0)

                        syUserPassWord.save()

                        ' 提示信息
                        com.Azion.EloanUtility.UIUtility.alert("儲存成功！")

                        ' 清空Mail欄位
                        txtMail.Text = ""
                    Else

                        ' 提示信息
                        com.Azion.EloanUtility.UIUtility.alert(VerifyPasswordPolicy(txtEditThridNewPWD.Text.Trim))
                    End If
                End If
            Else

                ' 修改Mail欄位
                If Not syUserPassWord.loadByPK(sStaffid) Then
                    syUserPassWord.setAttribute("STAFFID", sStaffid)
                End If

                syUserPassWord.setAttribute("MAILADDR", txtMail.Text.Trim)
                syUserPassWord.save()

                ' 提示信息
                com.Azion.EloanUtility.UIUtility.alert("儲存成功！")
            End If
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 人員編號文本框里的值變化時的事件
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub txtEditThridStaffId_TextChanged(ByVal sender As Object, ByVal e As EventArgs) Handles txtEditThridStaffId.TextChanged
        Try
            Dim syUserPassword As New SY_USERPASSWORD(GetDatabaseManager())

            ' 取得當前人的Mail,若能查詢出資料
            If syUserPassword.loadByPK("S0" & txtEditThridStaffId.Text.Trim) Then
                txtMail.Text = syUserPassword.getAttribute("MAILADDR")
            Else
                txtMail.Text = ""
            End If
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub
#End Region
End Class