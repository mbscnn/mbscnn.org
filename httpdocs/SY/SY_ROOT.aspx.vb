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

Public Class SY_ROOT
    Inherits SYUIBase

    Protected m_sMinPasswordLength As String = String.Empty   ' 所輸入密碼的最小長度
    Protected m_sMaxPasswordLength As String = String.Empty  ' 所輸入密碼的最大長度

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

                ' 添加Onclick事件
                treeViewAdministrators.Attributes.Add("onclick", "postBackByObject();")
            End If

            ' 設置頁面控件的初始狀態
            If m_bDisplayMode Then

                ' 只讀模式
                com.Azion.EloanUtility.UIUtility.setControlRead(Me)
            End If
        Catch ex As Exception
            SYUIBase.ShowErrMsg(Me, ex)
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

        'm_sMaxTimes_WrongPassword = com.Azion.EloanUtility.CodeList.getAppSettings("MaxTimes_WrongPassword")
        'm_sPasswordExpirationDays = com.Azion.EloanUtility.CodeList.getAppSettings("PasswordExpirationDays")
        m_sMinPasswordLength = com.Azion.EloanUtility.CodeList.getAppSettings("MinPasswordLength")
        m_sMaxPasswordLength = com.Azion.EloanUtility.CodeList.getAppSettings("MaxPasswordLength")

        ' 測試數據
        If m_bTesting Then
            m_sWorkingUserid = "S000001"
            m_bDisplayMode = False
            m_sHoflag = ""
        End If
    End Sub

    ''' <summary>
    ''' 初始化頁面
    ''' </summary>
    ''' <remarks></remarks>
    Sub initData()
        Try
            divloginPassWord.Style.Value = "display:block"
            divEditTwoPassWord.Style.Value = "display:none"
            divEditThridPassWord.Style.Value = "display:none"

            ' 如果沒有傳入參數,顯示登入畫面
            'If m_sHoflag = "" Then
            '    divloginPassWord.Style.Value = "display:block"
            '    divEditTwoPassWord.Style.Value = "display:none"
            '    divEditThridPassWord.Style.Value = "display:none"

            '    ' 設置初始焦點
            '    'txtStaffId.Focus()
            'Else
            '    Dim syUserPassword As New SY_USERPASSWORD(GetDatabaseManager())

            '    divEditThridPassWord.Style.Value = "display:block"
            '    divloginPassWord.Style.Value = "display:none"
            '    divEditTwoPassWord.Style.Value = "display:none"

            '    ' 獲得當前登錄者的編號
            '    'If m_sHoflag = "4" Then
            '    '    txtEditThridStaffId.Text = m_sWorkingUserid
            '    'Else
            '    '    txtEditThridStaffId.Text = ""
            '    'End If

            '    ' 取得當前人的Mail
            '    'If syUserPassword.loadByPK(m_sWorkingUserid) Then
            '    '    txtMail.Text = syUserPassword.getAttribute("MAILADDR")
            '    'End If

            '    '' 設置焦點
            '    'txtEditThridStaffId.Focus()

            '    'If m_sHoflag = "4" Then

            '    '    ' 使用者員工編號不可編輯
            '    '    txtEditThridStaffId.Enabled = False

            '    '    ' 設施初始焦點
            '    '    txtEditThridOPWD.Focus()
            '    'ElseIf m_sHoflag = "3" Or m_sHoflag = "5" Then
            '    '    com.Azion.EloanUtility.UIUtility.alert("無使用此交易權限！")

            '    '    ' 頁面不可編輯
            '    '    com.Azion.EloanUtility.UIUtility.setControlRead(Me)
            '    'End If

            '    If m_sHoflag = "1" Then
            '        txtEditThridOPWD.Enabled = False
            '    End If

            InitAdministratorsTree()

            'End If
        Catch ex As Exception
            SYUIBase.ShowErrMsg(Me, ex)
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
            Dim syRoot As New BosBase("SY_ROOT", GetDatabaseManager())

            If txtPWD1.Text.Length = 0 OrElse txtPWD2.Text.Length = 0 Then
                lblSystemMsg.Text = "登入需要輸入密碼！"
                Return
            End If

            '輸入的密碼
            Dim sPwd1 As String = CreateMD5Hash(txtPWD1.Text.Trim)
            Dim sPwd2 As String = CreateMD5Hash(txtPWD2.Text.Trim)

            '存在DB內的密碼
            Dim sdbPwd1 As String = CreateMD5Hash("P@ssw0rd")
            Dim sdbPwd2 As String = CreateMD5Hash("P@ssw0rd")

            '取得存在DB內的密碼
            Dim dr As DataRow = syRoot.GetDataRow(Nothing, "ROOTNAME", "ROOT")
            If IsNothing(dr) = False Then
                sdbPwd1 = BosBase.CDbType(Of String)(dr("PASSWORD1"), sdbPwd1)
                sdbPwd2 = BosBase.CDbType(Of String)(dr("PASSWORD2"), sdbPwd2)
            End If

            '如果不相同，要求重新輸入
            If sdbPwd1 = sPwd1 AndAlso sdbPwd2 = sPwd2 Then
                divEditTwoPassWord.Style.Value = "display:block"
                divEditThridPassWord.Style.Value = "display:none"
                divloginPassWord.Style.Value = "display:none"
                ResetAllControls()
            Else
                lblSystemMsg.Text = "密碼輸入錯誤，請重新輸入！"
                Return
            End If

            InitAdministratorsTree()

        Catch ex As Exception
            SYUIBase.ShowErrMsg(Me, ex)
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

            Dim syRelRoleUser As New FLOW_OP.TABLE.SY_REL_ROLE_USER(GetDatabaseManager())

            GetDatabaseManager.beginTran()


            '刪除所有應用系統管理人員
            syRelRoleUser.Delete("ROLEID", 1)


            For Each itemBranch As TreeNode In treeViewAdministrators.Nodes(0).ChildNodes
                For Each itemUser As TreeNode In itemBranch.ChildNodes
                    If itemUser.Checked Then
                        syRelRoleUser.ExecuteNonQuery( _
                            "insert into SY_REL_ROLE_USER(BRA_DEPNO, STAFFID, ROLEID)" & vbCrLf & _
                            "  select BRA_DEPNO, STAFFID, 1 AS ROLEID" & vbCrLf & _
                            "    from SY_REL_BRANCH_USER" & vbCrLf & _
                            "   where STAFFID = @STAFFID@" & vbCrLf,
                             "STAFFID", itemUser.Value)
                    End If
                Next
            Next

            GetDatabaseManager.commit()
            com.Azion.EloanUtility.UIUtility.alert("應用系統管理人員變更成功!")

            divEditThridPassWord.Style.Value = "display:none"
            divloginPassWord.Style.Value = "display:block"
            divEditTwoPassWord.Style.Value = "display:none"
            ResetAllControls()

        Catch ex As Exception
            GetDatabaseManager.Rollback()
            SYUIBase.ShowErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 點擊"資訊"頁面的"存儲"按鈕事件
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub btnEditThrid_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnEditThrid.Click

        Dim syRoot As BosBase

        Try

            ' 實例化
            syRoot = New BosBase("SY_ROOT", GetDatabaseManager())
            'Dim sMd5Password As String = String.Empty

            If txtEditThridOPWD.Text.Length = 0 Then
                lblEditThirdSystemMsg.Text = "變更密碼需要輸入原密碼！"
                txtEditThridOPWD.Focus()
                Return
            End If

            If txtEditThridNewPWD.Text.Length = 0 Then
                lblEditThirdSystemMsg.Text = "變更密碼需要輸入新密碼！"
                txtEditThridNewPWD.Focus()
                Return
            End If

            If txtEditThridCheckPWD.Text.Length = 0 Then
                lblEditThirdSystemMsg.Text = "變更密碼需要輸入確認密碼！"
                txtEditThridCheckPWD.Focus()
                Return
            End If

            ' 如果新密碼和確認密碼不相同
            If txtEditThridNewPWD.Text <> txtEditThridCheckPWD.Text Then
                lblEditThirdSystemMsg.Text = "新密碼及確認密碼不一致！"
                txtEditThridCheckPWD.Focus()
                Return
            End If

            '如果新密碼不符合密碼原則
            Dim sPwdPolicy As String = VerifyPasswordPolicy(txtEditThridNewPWD.Text)
            If String.IsNullOrEmpty(sPwdPolicy) = False Then
                lblEditThirdSystemMsg.Text = sPwdPolicy
                txtEditThridNewPWD.Focus()
                Return
            End If

            '輸入的密碼
            Dim sPwdNew As String = CreateMD5Hash(txtEditThridNewPWD.Text)
            Dim sPwdOld As String = CreateMD5Hash(txtEditThridOPWD.Text)

            '存在DB內的密碼
            Dim sdbPwd As String = CreateMD5Hash("P@ssw0rd")
            sdbPwd = CreateMD5Hash("1")

            '取得存在DB內的密碼
            Dim dr As DataRow = syRoot.GetDataRow(Nothing, "ROOTNAME", "ROOT")
            If IsNothing(dr) = False Then
                If RbChangePwd1.Checked Then
                    sdbPwd = BosBase.CDbType(Of String)(dr("PASSWORD1"), sdbPwd)
                Else
                    sdbPwd = BosBase.CDbType(Of String)(dr("PASSWORD2"), sdbPwd)
                End If
            End If

            If sdbPwd <> sPwdOld Then
                lblEditThirdSystemMsg.Text = "密碼輸入錯誤，請重新輸入！"
                txtEditThridOPWD.Focus()
                Return
            End If



            If RbChangePwd1.Checked Then
                syRoot.InsertUpdate("ROOTNAME", "ROOT",
                                     "PASSWORD1", sPwdNew)
            Else
                syRoot.InsertUpdate("ROOTNAME", "ROOT",
                                     "PASSWORD2", sPwdNew)
            End If

            '' 提示信息
            com.Azion.EloanUtility.UIUtility.alert("密碼變更存成功!")
            lblEditThirdSystemMsg.Text = "密碼變更存成功!"

        Catch ex As Exception
            SYUIBase.ShowErrMsg(Me, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 人員編號文本框里的值變化時的事件
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    'Protected Sub txtEditThridStaffId_TextChanged(ByVal sender As Object, ByVal e As EventArgs) Handles txtEditThridStaffId.TextChanged
    '    Try
    '        Dim syUserPassword As New SY_USERPASSWORD(GetDatabaseManager())

    '        ' 取得當前人的Mail,若能查詢出資料
    '        'If syUserPassword.loadByPK("S0" & txtEditThridStaffId.Text.Trim) Then
    '        '    txtMail.Text = syUserPassword.getAttribute("MAILADDR")
    '        'Else
    '        '    txtMail.Text = ""
    '        'End If
    '    Catch ex As Exception
    '        SYUIBase.ShowErrMsg(Me, ex)
    '    End Try
    'End Sub
#End Region

    Private Function checkExsitUserId(sUserId As String) As Boolean
        Throw New NotImplementedException
    End Function

    Protected Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Dim strScript As String
        strScript = "<Script Language='JAVAScript'> window.close();</Script>"
        Response.Write(strScript)
    End Sub

    Protected Sub btnChangePwd_Click(sender As Object, e As EventArgs) Handles btnChangePwd.Click
        divEditThridPassWord.Style.Value = "display:block"
        divloginPassWord.Style.Value = "display:none"
        divEditTwoPassWord.Style.Value = "display:none"
        ResetAllControls()

    End Sub


    Protected Sub btnCancelThrid_Click(sender As Object, e As EventArgs) Handles btnCance3.Click
        divEditThridPassWord.Style.Value = "display:none"
        divloginPassWord.Style.Value = "display:block"
        divEditTwoPassWord.Style.Value = "display:none"
        ResetAllControls()
    End Sub

    Protected Sub ResetAllControls()
        txtPWD1.Text = ""
        txtPWD2.Text = ""
        lblSystemMsg.Text = ""

        RbChangePwd1.Checked = True
        txtEditThridOPWD.Text = ""
        txtEditThridNewPWD.Text = ""
        txtEditThridCheckPWD.Text = ""
        lblEditThirdSystemMsg.Text = ""
    End Sub

    Protected Sub InitAdministratorsTree()
        Try
            treeViewAdministrators.Nodes.Clear()

            Dim sUpCode2366 As String = com.Azion.EloanUtility.CodeList.getAppSettings("UpCode2366")
            Dim apCodelist As New FLOW_OP.TABLE.SY_TABLEBASE("AP_CODE", GetDatabaseManager())
            Dim drRootData As DataRow

            drRootData = apCodelist.GetDataRow(Nothing, "CODEID", sUpCode2366)

            Dim rootNode As TreeNode = New TreeNode()
            rootNode.Value = 0
            rootNode.Text = "安泰銀行"

            If IsNothing(drRootData) = False Then
                ' root節點
                rootNode.Value = BosBase.CDbType(Of Integer)(drRootData("VALUE"), 0)
                rootNode.Text = BosBase.CDbType(Of String)(drRootData("TEXT"), "")
            End If

            rootNode.SelectAction = TreeNodeSelectAction.None
            rootNode.Checked = False
            rootNode.ShowCheckBox = False
            rootNode.Expand()
            treeViewAdministrators.Nodes.Add(rootNode)

            Dim syBranchUser As New FLOW_OP.TABLE.SY_REL_BRANCH_USER(GetDatabaseManager())
            Dim drcBranchUser As DataRowCollection

            drcBranchUser = syBranchUser.GetDataRowCollection( _
                "select distinct BR.BRID, BR.BRCNAME, UR.STAFFID, UR.USERNAME, RRU.ROLEID" & vbCrLf & _
                "  from SY_USER UR" & vbCrLf & _
                " inner join SY_REL_BRANCH_USER RBU" & vbCrLf & _
                "    on UR.STAFFID = RBU.STAFFID" & vbCrLf & _
                "   and UR.STATUS = '0'" & vbCrLf & _
                " inner join SY_BRANCH BR" & vbCrLf & _
                "    on BR.BRA_DEPNO = RBU.BRA_DEPNO" & vbCrLf & _
                "  left join SY_REL_ROLE_USER RRU" & vbCrLf & _
                "    on RRU.STAFFID = RBU.STAFFID" & vbCrLf & _
                "   and RRU.BRA_DEPNO = RBU.BRA_DEPNO" & vbCrLf & _
                "   and RRU.ROLEID = 1" & vbCrLf & _
                " where BR.PARENT = 0" & vbCrLf & _
                "   and BR.HOFLAG in ('1', '2', '3')" & vbCrLf & _
                " order by BR.BRID, UR.STAFFID" & vbCrLf, Nothing)

            Dim drCurrentBranch As DataRow = Nothing
            Dim nodeCurrentBranch As TreeNode
            Dim nodeUser As TreeNode

            '加入所有分行內的所有USER
            For Each drBranchUser As DataRow In drcBranchUser
                '加入分行NODE
                If IsNothing(drCurrentBranch) = True OrElse
                    BosBase.CDbType(Of String)(drCurrentBranch("BRID"), "") <> BosBase.CDbType(Of String)(drBranchUser("BRID"), "") Then
                    drCurrentBranch = drBranchUser

                    nodeCurrentBranch = New TreeNode

                    nodeCurrentBranch.Value = BosBase.CDbType(Of String)(drCurrentBranch("BRID"), "")
                    nodeCurrentBranch.Text = BosBase.CDbType(Of String)(drCurrentBranch("BRCNAME"), "") & " (" & nodeCurrentBranch.Value & ")"
                    nodeCurrentBranch.Checked = False
                    nodeCurrentBranch.ShowCheckBox = False
                    nodeCurrentBranch.Collapse()
                    nodeCurrentBranch.SelectAction = TreeNodeSelectAction.None
                    treeViewAdministrators.Nodes(0).ChildNodes.Add(nodeCurrentBranch)
                End If

                '加入USER的NODE
                nodeUser = New TreeNode
                nodeUser.Value = BosBase.CDbType(Of String)(drBranchUser("STAFFID"), "")
                nodeUser.Text = BosBase.CDbType(Of String)(drBranchUser("USERNAME"), "") & " (" & nodeUser.Value & ")"

                If BosBase.CDbType(Of Integer)(drBranchUser("ROLEID"), 0) <> 0 Then
                    nodeUser.Checked = True
                    nodeCurrentBranch.Expand()
                End If

                nodeCurrentBranch.ChildNodes.Add(nodeUser)
            Next
        Catch ex As Exception
            SYUIBase.ShowErrMsg(Me, ex)
        End Try
    End Sub


    Protected Sub treeViewAdministrators_TreeNodeCheckChanged(sender As Object, e As System.Web.UI.WebControls.TreeNodeEventArgs) Handles treeViewAdministrators.TreeNodeCheckChanged
        Try

            Dim nodeUser As TreeNode = e.Node
            Dim syRoleUser As New FLOW_OP.TABLE.SY_REL_ROLE_USER(GetDatabaseManager)
            Dim nCount As Integer

            If IsNothing(nodeUser) Then
                Return
            End If

            nCount = syRoleUser.GetCount("STAFFID", nodeUser.Value, "ROLEID", 1)

            If (nodeUser.Checked = True And nCount = 0) OrElse
                (nodeUser.Checked = False And nCount <> 0) Then

                nodeUser.Text = nodeUser.Text.Replace("<font color='red'>", "").Replace("</font>", "")
                nodeUser.Text = String.Format("<font color='red'>{0}</font>", nodeUser.Text)

            Else
                nodeUser.Text = nodeUser.Text.Replace("<font color='red'>", "").Replace("</font>", "")
            End If

        Catch ex As Exception
            SYUIBase.ShowErrMsg(Me, ex)
        End Try
    End Sub

    Protected Sub btnCancel2_Click(sender As Object, e As EventArgs) Handles Cancel2.Click
        divEditThridPassWord.Style.Value = "display:none"
        divloginPassWord.Style.Value = "display:block"
        divEditTwoPassWord.Style.Value = "display:none"
        ResetAllControls()
    End Sub
End Class