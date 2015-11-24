Imports com.Azion.EloanUtility

''' <summary>
''' 程式說明：開發使用者登入畫面
''' 建立者：Lake
''' 建立日期：2012-04-06
''' </summary>

Imports com.Azion.UITools
Imports AUTH_OP
Imports AUTH_OP.TABLE

Public Class SY_LOGIN
    Inherits SYUIBase

    Public Sub New()
        MyBase.new()
    End Sub

    Public Sub New(ByVal dbManager As com.Azion.NET.VB.DatabaseManager)
        MyBase.new()
        m_dbManager = dbManager
    End Sub



#Region "濟南昱勝添加"
#Region "Lake Function"

    ''' <summary>
    ''' 檢核用戶輸入編號是否存在
    ''' </summary>
    ''' <param name="sStaffId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' <history>
    ''' [Lake] 2012/04/06 Created
    '''</history>
    Function checkExsitUserId(ByVal sStaffId As String) As Boolean
        Try
            ' 實例化用戶類
            Dim syUser As New SY_USER(GetDatabaseManager())

            ' 檢核用戶輸入編號是否存在
            Return syUser.checkExsitUserId(sStaffId)
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' 查詢登入者相關信息
    ''' </summary>
    ''' <param name="sStaffId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' <history>
    ''' [Lake] 2012/04/06 Created
    '''</history>
    Function getUserInfo(ByVal sLoginBrId As String, ByVal sLoginStaffId As String, Optional ByVal sLoginName As String = "", Optional ByVal sWorkingBrId As String = "", Optional ByVal sWorkingStaffId As String = "", Optional ByVal sWorkingName As String = "") As Boolean ', Optional sBraDepno As Integer = 0, Optional sBrId As String = ""
        Try
            ' 實例化用戶資料類
            Dim staffInfo As New StaffInfo
            Dim sRoleId As String = ""
            Dim sWorkingRoleId As String = ""
            Dim sSysId As String = ""

            Session.Remove("StaffInfo")

            Dim funs As New SY_FUNCTION_CODEList(GetDatabaseManager())
            Dim syUser As New SY_USER(GetDatabaseManager())
            Dim sySyss As New SY_SYSIDList(GetDatabaseManager())
            'Dim syRelBranchUserList As New SY_REL_BRANCH_USERList(GetDatabaseManager())

            Dim syRelRoleUserList As New SY_REL_ROLE_USERList(GetDatabaseManager())

            Dim sTempTable As New DataTable


            'staffInfo.SessionID = sStaffId
            ' 取得user基本信息
            If syUser.checkExsitUserId(sLoginStaffId) Then
                '代理人, Owner, login user
                staffInfo.LoginUserId = sLoginStaffId
                staffInfo.LoginUserName = sLoginName
                staffInfo.LoginBrid = sLoginBrId

                staffInfo.WorkingStaffid = staffInfo.LoginUserId
                staffInfo.WorkingName = staffInfo.LoginUserName
                staffInfo.WorkingBrid = staffInfo.LoginBrid
                '被代理人, Client, 
                If sWorkingBrId <> Nothing Then
                    staffInfo.WorkingStaffid = sWorkingStaffId
                    staffInfo.WorkingBrid = sWorkingBrId
                    staffInfo.WorkingName = sWorkingName
                End If

            End If

            ' 取得登入者對應部門及單位
            Dim dt As DataTable = MBSC.UICtl.UIShareFun.getRoles(Me.GetDatabaseManager(), sLoginStaffId)
            Dim dtDepNo As DataTable = dt.DefaultView.ToTable(True, New String() {"BRID", "BRA_DEPNO", "BRCNAME"}) '取得登入者"所有的"DepNo
            Dim dtBrId As DataTable = dt.DefaultView.ToTable(True, New String() {"BRID"}) '取得登入者的"所有的"BrId

            Dim rowCurrentUserDepNo() As DataRow = dtDepNo.Select("brid='" & sLoginBrId & "'") '取得登入者欲登入的部門代碼
            Dim rowCurrentUserTopDepNo() As DataRow = dt.DefaultView.ToTable(True, New String() {"BRID", "BRA_DEPNO", "BRCNAME", "PARENTID"}).Select("brid='" & sLoginBrId & "' and parentid=0") '取得登入者欲登入的最上層部門代碼,一定是<=1

            If rowCurrentUserTopDepNo.Length >= 1 Then
                staffInfo.LoginTopDepNo = rowCurrentUserTopDepNo(0)("BRA_DEPNO").ToString
            Else
                Dim syBranch As New SY_BRANCH(Me.GetDatabaseManager())
                If syBranch.getFirstLevelBraDepNo(sLoginBrId) Then
                    staffInfo.LoginTopDepNo = syBranch.getString("BRA_DEPNO")
                End If
            End If

            staffInfo.WorkingTopDepNo = staffInfo.LoginTopDepNo

            If rowCurrentUserDepNo.Count > 0 Then
                staffInfo.LoginBrid = sLoginBrId
                staffInfo.LoginBrCname = rowCurrentUserDepNo(0)("BRCNAME").ToString
                staffInfo.WorkingBrCname = staffInfo.LoginBrCname
            Else
                Throw New Exception("使用者無權限")
            End If
             

            ' 取得角色
            If syRelRoleUserList.loadByParas(staffInfo.LoginUserId, staffInfo.LoginBrid) > 0 Then

                ' 循環集合，組字串 
                For Each row As DataRow In syRelRoleUserList.getCurrentDataSet.Tables(0).Rows
                    sRoleId &= row("ROLEID").ToString & ","
                    'sTempBraDepNo &= "'" & row("BRA_DEPNO").ToString & "'" & ","
                Next

                staffInfo.LoginRoleId = sRoleId.Substring(0, sRoleId.Length - 1)
                staffInfo.WorkingRoleId = staffInfo.LoginRoleId
            End If

            If staffInfo.LoginUserId <> staffInfo.WorkingStaffid Then
                syRelRoleUserList.clear()
                If syRelRoleUserList.loadByParas(staffInfo.WorkingStaffid, staffInfo.WorkingBrid) > 0 Then
                    ' 循環集合，組字串
                    For Each row As DataRow In syRelRoleUserList.getCurrentDataSet.Tables(0).Rows
                        sWorkingRoleId &= row("ROLEID").ToString & ","
                        'sTempBraDepNo &= "'" & row("BRA_DEPNO").ToString & "'" & ","
                    Next
                    staffInfo.WorkingRoleId = sWorkingRoleId.Substring(0, sWorkingRoleId.Length - 1)
                End If


                Dim syBranch As New SY_BRANCH(Me.GetDatabaseManager())
                If syBranch.getFirstLevelBraDepNo(sWorkingBrId) Then
                    staffInfo.WorkingTopDepNo = syBranch.getString("BRA_DEPNO")
                    staffInfo.WorkingBrCname = syBranch.getString("BRCNAMW")
                End If
            End If


            ' 取得第一個系統功能ID及名稱
            If Not staffInfo.LoginRoleId Is Nothing Then

                If sySyss.loadByWorkingRoleId(staffInfo.WorkingRoleId) > 0 Then
                    For Each sySys As SY_SYSID In sySyss
                        sSysId = sySys.getAttribute("SYSID").ToString()

                        ' 根據登入者, 角色添加相關的交易
                        funs.clear()

                        If staffInfo.FuncList.Count = 0 Then
                            sTempTable = funs.genFunList(staffInfo.WorkingStaffid, staffInfo.WorkingBrid, sSysId)

                            '' 根據父節點查詢 若沒有資料 則使用子節點的部門看看是否有資料
                            'If sTempBraDepNo.Length > 0 Then
                            '    sTempBraDepNo = sTempBraDepNo.Substring(0, sTempBraDepNo.Length - 1)
                            'End If

                            If sTempTable Is Nothing Then
                                ' sTempTable = funs.genFunListList(staffInfo.WorkingStaffid, sTempBraDepNo, sSysId)
                                Dim sss
                                sss = ""
                            End If

                            ' 以key value的形式將系統資料存入到hashtable中
                            staffInfo.FuncList.Add(sSysId, sTempTable)

                            ' todo:暫時存取第一筆資料
                            staffInfo.SysId = sSysId
                        End If
                    Next
                End If
            End If

            ' 用戶基本資料保存到Session中
            Session("StaffInfo") = staffInfo

        Catch ex As Exception
            Throw
            'Finally
            '    '輸出UserInfo的Funciton至FLOW_OP
            '    If (IsNothing(export_UserInfo)) Then
            '        export_UserInfo = New ExportUserInfo
            '        FLOW_OP.ELoanFlow.m_callbackUserInfo = export_UserInfo
            '    End If
        End Try
    End Function

    ''' <summary>
    ''' 根據“登錄人員編號”，“用戶角色”，“系統功能編號”取得交易資料
    ''' </summary>
    ''' <param name="sStaffid">登錄人員編號</param>
    ''' <param name="sRoleIdList">用戶角色</param>
    ''' <param name="sSysId">系統功能編號</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' <history>
    ''' [Lake] 2012/04/11 Created
    ''' </history>
    'Sub getFuncionList(ByVal sStaffid As String, ByVal sRoleIdList As String, ByVal sSysId As String)
    '    Try
    '        Using databaseManager As com.Azion.NET.VB.DatabaseManager = GetDatabaseManager()

    '            Dim syTempFunctionList As New TEMPFUNCTIONLIST(databaseManager)
    '            Dim syTempFunctionListList As New TEMPFUNCTIONLISTList(databaseManager)
    '            Dim sCount As Integer = 0

    '            ' 獲取用戶交易資料One
    '            sCount = syTempFunctionList.insertTempOne(sStaffid, sRoleIdList, sSysId)

    '            ' 獲取用戶交易資料Two
    '            sCount += syTempFunctionList.insertTempTwo(sStaffid, sRoleIdList, sSysId)

    '            If sCount > 1 Then

    '                ' 獲取用戶交易資料Three
    '                sCount += syTempFunctionList.insertTempThree(sStaffid, sRoleIdList)
    '            End If
    '        End Using
    '    Catch ex As Exception
    '        SYUIBase.showErrMsg(Me, ex)
    '    End Try
    'End Sub

    ''' <summary>
    '''  查詢登入者的部門信息
    ''' </summary>
    ''' <param name="sStaffId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function GetBranchCount(ByVal sStaffId As String) As DataTable
        Dim syRelBranchUserList As New AUTH_OP.SY_REL_BRANCH_USERList(GetDatabaseManager())
        Dim dtData As New DataTable

        If syRelBranchUserList.GetBranchCount(sStaffId) Then
            dtData = syRelBranchUserList.getCurrentDataSet.Tables(0)
        End If

        Return dtData
    End Function
#End Region
#End Region

End Class
