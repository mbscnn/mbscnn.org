Imports com.Azion.NET.VB

Public Class SY_ROLEList
    Inherits BosList

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.New("SY_ROLE", dbManager)
    End Sub

    Overrides Function newBos() As BosBase
        Dim bos As SY_ROLE = New SY_ROLE(MyBase.getDatabaseManager)
        Return bos
    End Function


#Region "濟南昱勝添加"
#Region "Avril Function"

    ''' <summary>
    ''' 查詢類型不為系統定義的資料
    ''' </summary>
    ''' <param name="sParent">父節點</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Add by Avril 2012/04/18
    ''' </remarks>
    Function loadDataByType(ByVal sParent As String) As Boolean
        Try
            Dim strSql = "SELECT " & _
                         "    * FROM SY_ROLE " & _
                         " WHERE " & _
                         "    ROLETYPE<>1 " & _
                         " AND PARENT = " & ProviderFactory.PositionPara & "PARENT"

            Dim para(0) As IDbDataParameter
            para(0) = ProviderFactory.CreateDataParameter("PARENT", sParent)

            Return MyBase.loadBySQL(strSql, para)
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 根據父節點和角色編號查詢資料
    ''' </summary>
    ''' <param name="sParent">父節點</param>
    ''' <param name="sRoleId">角色編號</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Add by 2012/07/13
    ''' </remarks>
    Function loadDataByCon(ByVal sParent As String, ByVal sRoleId As String) As Boolean
        Try
            Dim strSql = "SELECT " & _
                         "    * FROM SY_ROLE " & _
                         " WHERE " & _
                         "    ROLETYPE<>1 " & _
                         " AND PARENT = " & ProviderFactory.PositionPara & "PARENT" & _
                         " AND ROLEID = " & ProviderFactory.PositionPara & "ROLEID"

            Dim para(1) As IDbDataParameter
            para(0) = ProviderFactory.CreateDataParameter("PARENT", sParent)
            para(1) = ProviderFactory.CreateDataParameter("ROLEID", sRoleId)

            Return MyBase.loadBySQL(strSql, para)
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function


    ''' <summary>
    ''' 根據角色編號查詢資料
    ''' </summary>
    ''' <param name="sRoleId">角色編號</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Avril 2012/05/10 Add
    ''' </remarks>
    Function loadDataByRoleId(ByVal sRoleId As String) As Boolean
        Try
            Dim strSql As String = "SELECT  " & _
                                   "   *" & _
                                   " FROM  SY_ROLE " & _
                                   " WHERE ROLEID IN( " & sRoleId & ") "

            Return MyBase.loadBySQL(strSql)
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 依這個員工所登錄的DepNo取得所有的角色名稱
    ''' </summary>
    ''' <param name="sStaffid">Sxxxxxx</param>
    ''' <param name="iDepNo">部門編號</param>
    ''' <param name="sSysId">D、F、Z</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Add by Avril 2012/06/06
    ''' </remarks>
    Function genFunList(ByVal sStaffid As String, ByVal sSubSysId As String, ByVal sSysId As String) As DataTable
        Try
            Dim strSql As String = "  ;WITH rc(ROLENAME,parentid,ROLEID,DISABLED,SerialStr,Level) AS" & _
                                    " (SELECT ROLENAME," & _
                                    "         parent," & _
                                    "         ROLEID," & _
                                    "         DISABLED," & _
                                    "         CONVERT(nvarchar(128), RIGHT(REPLICATE('0', 5) + CAST(a.ROLEID as VARCHAR), 5)),0 Level" & _
                                    "    FROM SY_ROLE a" & _
                                    "   WHERE parent = 0" & _
                                    "  UNION ALL" & _
                                    "  SELECT p.ROLENAME," & _
                                    "         p.parent," & _
                                    "         p.ROLEID," & _
                                    "         p.DISABLED," & _
                                    "         CONVERT(nvarchar(128), RIGHT(REPLICATE('0', 5) + CAST(SerialStr as VARCHAR), 5) + '-' + CONVERT(nvarchar(128), RIGHT(REPLICATE('0', 5) + CAST(p.ROLEID as VARCHAR), 5))), Level+1 Level" & _
                                    "    FROM SY_ROLE P, rc" & _
                                    "   WHERE rc.ROLEID = P.PARENT" & _
                                    "     and p.PARENT > 0)" & _
                                    "select distinct Level,rc.ROLENAME ROLENAME, rc.parentid, a.SYSID, a.SUBSYSID, b.ROLEID,  " & ProviderFactory.PositionPara & "STAFFID STAFFID, null CheckFlag ,SUBSYSNAME, SerialStr " & _
                                    "  from SY_ROLESUITSYS a, SY_ROLESUITSYS b, rc,SY_SUBSYSID c " & _
                                    " where a.SYSID=" & ProviderFactory.PositionPara & "SYSID" & _
                                    " and a.SUBSYSID=" & ProviderFactory.PositionPara & "SUBSYSID " & _
                                    " and rc.ROLEID = b.ROLEID" & _
                                    " and a.ROLEID = b.ROLEID and a.SUBSYSID=c.SUBSYSID   and b.SUBSYSID=c.SUBSYSID " & _
                                    " and rc.DISABLED=0 " & _
                                    " order by  SerialStr OPTION(MAXRECURSION 5)"

            Dim para(2) As IDbDataParameter
            para(0) = ProviderFactory.CreateDataParameter("STAFFID", sStaffid)
            para(1) = ProviderFactory.CreateDataParameter("SUBSYSID", sSubSysId)
            para(2) = ProviderFactory.CreateDataParameter("SYSID", sSysId)

            If Me.loadBySQLOnlyDs(strSql, para) > 0 Then
                Return Me.getCurrentDataSet.Tables(0)
            End If
            Return Nothing
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function


    ''' <summary>
    ''' 依這個員工所登錄的DepNo取得所有的角色名稱
    ''' </summary>
    ''' <param name="sStaffid">登入者</param>
    ''' <param name="iDepNo">部門編號</param>
    ''' <param name="sSysId">系統編號</param>
    ''' <param name="sRoleList">角色編號</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Add by Avril 2012/06/07
    ''' </remarks>
    Function genFunListCheck(ByVal sStaffid As String, ByVal iDepNo As Integer, ByVal sSysId As String, ByVal sRoleList As String, ByVal sSysIdList As String, ByVal sSubSysId As String, ByVal sTSubSysId As String) As DataTable
        Try
            Dim strSql As String = "  ;WITH rc(ROLENAME,parentid,ROLEID,DISABLED,SerialStr,Level) AS" & _
                                    " (SELECT ROLENAME," & _
                                    "         parent," & _
                                    "         ROLEID," & _
                                    "         DISABLED," & _
                                    "         CONVERT(nvarchar(128), RIGHT(REPLICATE('0', 5) + CAST(a.ROLEID as VARCHAR), 5)),0 Level" & _
                                    "    FROM SY_ROLE a" & _
                                    "   WHERE parent = 0" & _
                                    "  UNION ALL" & _
                                    "  SELECT p.ROLENAME," & _
                                    "         p.parent," & _
                                    "         p.ROLEID," & _
                                    "         p.DISABLED," & _
                                    "         CONVERT(nvarchar(128), RIGHT(REPLICATE('0', 5) + CAST(SerialStr as VARCHAR), 5) + '-' + CONVERT(nvarchar(128), RIGHT(REPLICATE('0', 5) + CAST(p.ROLEID as VARCHAR), 5))), Level+1 Level" & _
                                    "    FROM SY_ROLE P, rc" & _
                                    "   WHERE rc.ROLEID = P.PARENT" & _
                                    "     and p.PARENT > 0)" & _
                                    "select Level,rc.ROLENAME ROLENAME, rc.parentid, a.SYSID, a.SUBSYSID, b.ROLEID,  " & ProviderFactory.PositionPara & "STAFFID STAFFID, null CheckFlag ,SUBSYSNAME" & _
                                    "  from SY_ROLESUITSYS a, SY_ROLESUITSYS b, rc,SY_SUBSYSID c " & _
                                    " where a.SYSID=" & ProviderFactory.PositionPara & "SYSID" & _
                                    " and rc.ROLEID = b.ROLEID" & _
                                    " and a.ROLEID = b.ROLEID and a.SUBSYSID=c.SUBSYSID   and b.SUBSYSID=c.SUBSYSID " & _
                                    " and rc.DISABLED=0 " & _
                                    " and  a.SUBSYSID=" & ProviderFactory.PositionPara & "SUBSYSID" & _
                                    " and a.ROLEID IN ( " & sRoleList & ") " & _
                                    " order by  SerialStr OPTION(MAXRECURSION 5)"

            '"         where m.STAFFID = " & ProviderFactory.PositionPara & "STAFFID" & _
            '                    "         and m.BRA_DEPNO=" & ProviderFactory.PositionPara & "BRA_DEPNO" & _

            Dim para(2) As IDbDataParameter
            para(0) = ProviderFactory.CreateDataParameter("STAFFID", sStaffid)
            'para(1) = ProviderFactory.CreateDataParameter("BRA_DEPNO", iDepNo)
            para(1) = ProviderFactory.CreateDataParameter("SYSID", sSysId)
            para(2) = ProviderFactory.CreateDataParameter("SUBSYSID", sTSubSysId)

            If Me.loadBySQLOnlyDs(strSql, para) > 0 Then
                Return Me.getCurrentDataSet.Tables(0)
            End If
            Return Nothing
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

#End Region

#Region "Zack Function"

    ''' <summary>
    ''' 根據分行單位代碼，角色類別查詢角色信息
    ''' </summary>
    ''' <param name="sBraDepno">單位代碼</param>
    ''' <param name="sHoFlag">單位標示</param>
    ''' <param name="sRoleid">角色編號</param>
    ''' <param name="sFlag">是否查詢ListBox中所需的角色</param>
    ''' <param name="sStaffid">人員編號</param>
    ''' <param name="sDynamicSQL">過濾條件</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Zack 2012-05-16 Create
    ''' Zack 2012-06-26 Update
    ''' </remarks>
    Public Function loadByBraDepnoStaffid(ByVal sFlag As String,
                                          ByVal sHoFlag As String,
                                          ByVal sRoleid As String,
                                          ByVal sBraDepno As String,
                                          ByVal sStaffid As String,
                                          ByVal sRight As String,
                                          Optional ByVal sDynamicSQL As String = Nothing) As Integer
        Try
            Dim sSql As String = String.Empty

            ' 單位安控主管（6）
            If sHoFlag = "1" And sRoleid.Contains(",3,") Then

                If sRoleid.Length = 3 Then
                    sSql = "SELECT ROLEID AS LValue,CONVERT(VARCHAR(4),ROLEID) + '  ' +  ROLENAME AS LText " & _
                           "  FROM SY_ROLE " & _
                           " WHERE ROLEID= 5"
                ElseIf sRoleid.Contains(",5,") Then
                    sSql = "SELECT ROLEID AS LValue,CONVERT(VARCHAR(4),ROLEID) + '  ' + ROLENAME AS LText " & _
                           " FROM SY_ROLE" & _
                           " WHERE DISABLED = 0" & _
                                " AND((ROLEID NOT IN(SELECT roleid FROM SY_ROLESUITBRANCH WHERE BRA_DEPNO !=" & ProviderFactory.PositionPara & "BRA_DEPNO)" & _
                                " AND ROLETYPE=2) OR roleid=6 OR roleid=5)"
                Else
                    sSql = "SELECT ROLEID AS LValue,CONVERT(VARCHAR(4),ROLEID) + '  ' + ROLENAME AS LText " & _
                           " FROM SY_ROLE" & _
                           " WHERE DISABLED = 0" & _
                                 " AND((ROLEID NOT IN(SELECT roleid FROM SY_ROLESUITBRANCH WHERE BRA_DEPNO !=" & ProviderFactory.PositionPara & "BRA_DEPNO)" & _
                                 " AND ROLETYPE=2) OR roleid=5)"
                End If
            ElseIf sHoFlag = "1" And (sRoleid.Contains(",5,") And sRoleid.Length > 3) Then

                sSql = "SELECT ROLEID AS LValue,CONVERT(VARCHAR(4),ROLEID) + '  ' + ROLENAME AS LText " & _
                       " FROM SY_ROLE" & _
                       " WHERE DISABLED = 0" & _
                           " AND((ROLEID NOT IN(SELECT roleid FROM SY_ROLESUITBRANCH WHERE BRA_DEPNO !=" & ProviderFactory.PositionPara & "BRA_DEPNO)" & _
                           " AND ROLETYPE=2) OR roleid=6)"

            Else

                sSql = "SELECT ROLEID AS LValue, CONVERT(VARCHAR(4),ROLEID) + '  ' + ROLENAME AS LText" & _
                      " FROM SY_ROLE " & _
                      " WHERE DISABLED = 0 " & _
                          " AND ROLEID NOT IN (SELECT roleid FROM SY_ROLESUITBRANCH WHERE BRA_DEPNO !=" & ProviderFactory.PositionPara & "BRA_DEPNO)" & _
                          " AND ROLETYPE = '2'"
            End If

            Dim para(0) As IDbDataParameter
            para(0) = ProviderFactory.CreateDataParameter("BRA_DEPNO", sBraDepno)

            ' 若不是查詢ListBox中所需的角色
            If sFlag <> "N" Then

                If sRight = "N" Then
                    sSql = sSql & " AND SY_ROLE.Roleid NOT IN (SELECT roleid FROM SY_REL_ROLE_USER " & _
                          " WHERE BRA_DEPNO=" & ProviderFactory.PositionPara & "BRA_DEPNO" & _
                          " AND staffid=" & ProviderFactory.PositionPara & "staffid)"
                Else
                    sSql = sSql & " AND SY_ROLE.Roleid IN (SELECT roleid FROM SY_REL_ROLE_USER " & _
                           " WHERE BRA_DEPNO=" & ProviderFactory.PositionPara & "BRA_DEPNO" & _
                           " AND staffid=" & ProviderFactory.PositionPara & "staffid)"
                End If

                ReDim Preserve para(1)
                para(1) = ProviderFactory.CreateDataParameter("staffid", sStaffid)
            End If

            If String.IsNullOrEmpty(sDynamicSQL) = False Then
                sSql = "Select * from (" & sSql & ") ORI where " & sDynamicSQL
            End If


            Return MyBase.loadBySQL(sSql, para)

        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 根據角色編號，分行代號查詢角色信息
    ''' </summary>
    ''' <param name="sBraDepno">分行代號</param>
    ''' <param name="sRoleID">角色編號</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Zack 2012-05-18 Create
    ''' Zack 2012-07-12 Update
    ''' </remarks>
    Public Function loadByPeopleEdit(ByVal sBraDepno As String, ByVal sUinRoleID As String, ByVal sHoFlag As String, ByVal sRoleID As String) As Integer
        Try
            Dim sSql As String = String.Empty

            ' 單位安控主管（6）
            If sHoFlag = "1" And sRoleID.Contains(",3,") Then

                If sRoleID.Length = 3 Then
                    sSql = "SELECT ROLEID AS LValue,CONVERT(VARCHAR(4),ROLEID) + '  ' +  ROLENAME AS LText FROM SY_ROLE WHERE ROLEID= 5"
                ElseIf sRoleID.Contains(",5,") Then
                    sSql = "SELECT ROLEID AS LValue,CONVERT(VARCHAR(4),ROLEID) + '  ' + ROLENAME AS LText " & _
                           " FROM SY_ROLE" & _
                           " WHERE DISABLED = 0" & _
                                " AND((ROLEID NOT IN(SELECT roleid FROM SY_ROLESUITBRANCH WHERE BRA_DEPNO !=" & ProviderFactory.PositionPara & "BRA_DEPNO)" & _
                                " AND ROLETYPE=2) OR roleid=6 OR roleid=5)"
                Else
                    sSql = "SELECT ROLEID AS LValue,CONVERT(VARCHAR(4),ROLEID) + '  ' + ROLENAME AS LText " & _
                           " FROM SY_ROLE" & _
                           " WHERE DISABLED = 0" & _
                                 " AND((ROLEID NOT IN(SELECT roleid FROM SY_ROLESUITBRANCH WHERE BRA_DEPNO !=" & ProviderFactory.PositionPara & "BRA_DEPNO)" & _
                                 " AND ROLETYPE=2) OR roleid=5)"
                End If
            ElseIf sHoFlag = "1" And (sRoleID.Contains(",5,") And sRoleID.Length > 3) Then

                sSql = "SELECT ROLEID AS LValue,CONVERT(VARCHAR(4),ROLEID) + '  ' + ROLENAME AS LText " & _
                       " FROM SY_ROLE" & _
                       " WHERE DISABLED = 0" & _
                           " AND((ROLEID NOT IN(SELECT roleid FROM SY_ROLESUITBRANCH WHERE BRA_DEPNO !=" & ProviderFactory.PositionPara & "BRA_DEPNO)" & _
                           " AND ROLETYPE=2) OR roleid=6)"

            Else

                sSql = "SELECT ROLEID AS LValue, CONVERT(VARCHAR(4),ROLEID) + '  ' + ROLENAME AS LText" & _
                      " FROM SY_ROLE " & _
                      " WHERE DISABLED = 0 " & _
                          " AND ROLEID NOT IN (SELECT roleid FROM SY_ROLESUITBRANCH WHERE BRA_DEPNO !=" & ProviderFactory.PositionPara & "BRA_DEPNO)" & _
                          " AND ROLETYPE = '2'"
            End If

            Dim para(0) As IDbDataParameter
            para(0) = ProviderFactory.CreateDataParameter("BRA_DEPNO", sBraDepno)

            ' 若有角色編號
            If sUinRoleID <> "" Then
                sSql = sSql & " AND SY_ROLE.ROLEID NOT IN " & _
                        " (" + sUinRoleID + ")"
            End If

            Return MyBase.loadBySQL(sSql, para)
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region
#End Region

#Region "Titan Function"
    Function genAllRoleList() As DataTable
        Try
            Dim strSql As String = ";WITH rc(RoleId,RoleName,Disabled,RoleType,PARENTID,SerialNum ,SerialStr,Level) AS" & _
                                    "(SELECT RoleId,RoleName,Disabled,RoleType,PARENT," & _
                                    "     CAST(row_number() over (order by RoleId) as int) SerialNum ," & _
                                    "     CAST(row_number() over (order by RoleId) as varchar(255)) SerialStr," & _
                                    "     0 Level" & _
                                    "    FROM SY_ROLE a" & _
                                    "   WHERE PARENT = 0" & _
                                    "  UNION ALL" & _
                                    "  SELECT p.RoleId," & _
                                    "         p.RoleName," & _
                                    "         p.Disabled," & _
                                    "         p.RoleType," & _
                                    "         p.PARENT," & _
                                    "         CAST(rc.SerialNum + 1 as int)," & _
                                    "         CAST(RTrim(rc.SerialStr) + '-' + CAST(row_number() over (order by p.RoleId) as varchar(10)) as varchar(255))," & _
                                    "              rc.Level + 1" & _
                                    "    FROM SY_ROLE P, rc" & _
                                    "   WHERE rc.RoleId = P.PARENT" & _
                                    "     and p.PARENT > 0)" & _
                                    "select  *" & _
                                    "  from rc" & _
                                    " order by roleid,SerialStr  OPTION(MAXRECURSION 5)"



            If Me.loadBySQLOnlyDs(strSql) > 0 Then
                Return Me.getCurrentDataSet.Tables(0)
            End If
            Return Nothing
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region
End Class
