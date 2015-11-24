Imports com.Azion.NET.VB
Imports System.Text

Public Class SY_REL_ROLE_USERList
    : Inherits BosList

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.new("SY_REL_ROLE_USER", dbManager)
    End Sub

    Overrides Function newBos() As BosBase
        Return New SY_REL_ROLE_USER(MyBase.getDatabaseManager)
    End Function

#Region "濟南昱勝添加"
#Region "Lake Function"

    ''' <summary>
    ''' TODO:SY_REL_ROLE_USER裱中沒有BRID欄位,
    ''' 故暫屏蔽BRID條件
    ''' 待確認
    ''' 取得角色
    ''' </summary>
    ''' <param name="sStaffId"></param>
    ''' <param name="sBrid"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' <history>
    ''' [Lake] 2012/04/06 Created
    ''' Modify  by Avril 目前查詢條件由bra_depno 改為brid
    '''</history>
    Function loadByParas(ByVal sStaffId As String, ByVal sBrid As String) As Integer
        Try
            Dim strSql As String = "SELECT" _
                                 & "   SY_REL_ROLE_USER.ROLEID,SY_REL_ROLE_USER.BRA_DEPNO  " _
                                 & "FROM" _
                                 & "   SY_REL_ROLE_USER " _
                                 & " LEFT JOIN SY_BRANCH " _
                                 & " ON SY_REL_ROLE_USER.BRA_DEPNO = SY_BRANCH.BRA_DEPNO " _
                                 & " LEFT JOIN SY_ROLE ON SY_REL_ROLE_USER.ROLEID = SY_ROLE.ROLEID " _
                                 & "WHERE" _
                                 & "       STAFFID = " & ProviderFactory.PositionPara & "STAFFID " _
                                 & " AND SY_BRANCH.BRID  = " & ProviderFactory.PositionPara & "BRID" _
                                 & " AND SY_ROLE.DISABLED ='0'  AND SY_BRANCH.DISABLED='0'  "

            Dim para(1) As IDataParameter
            para(0) = ProviderFactory.CreateDataParameter("STAFFID", sStaffId)
            para(1) = ProviderFactory.CreateDataParameter("BRID", sBrid)

            Return MyBase.loadBySQLOnlyDs(strSql, para)

        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 根據當前登錄人員所在部門及其下屬部門查詢【人員角色清單】
    ''' </summary>
    ''' <param name="sBraDepNo"></param>
    ''' <returns></returns>
    ''' <remarks>[Lake] 2012/06/05 Created</remarks>
    Public Function getStaffRoleList(ByVal sBraDepNo As String) As Boolean
        Try
            Dim sb As New StringBuilder()
            sb.Append("SELECT ")
            sb.Append("     '[' + RIGHT(SY_USER.STAFFID,5) + '] ' + SY_USER.USERNAME AS STAFF ")
            sb.Append("    ,SY_ROLE.ROLENAME ")
            sb.Append("FROM ")
            sb.Append("     SY_REL_ROLE_USER JOIN SY_ROLE ON SY_REL_ROLE_USER.ROLEID= SY_ROLE.ROLEID ")
            sb.Append("     JOIN SY_USER ON SY_REL_ROLE_USER.STAFFID = SY_USER.STAFFID ")
            sb.Append("WHERE ")
            sb.Append("     SY_USER.STATUS=0 AND SY_ROLE.DISABLED=0 ")
            sb.Append("     AND SY_REL_ROLE_USER.BRA_DEPNO = '" & sBraDepNo & "' ")
            sb.Append("GROUP BY SY_USER.STAFFID,'[' + RIGHT(SY_USER.STAFFID,5) + '] ' + SY_USER.USERNAME,SY_ROLE.ROLENAME")

            Return Me.loadBySQL(sb.ToString())
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 根據當前登錄人員所在部門及其下屬部門查詢【角色人員清單】
    ''' </summary>
    ''' <param name="sBraDepNo"></param>
    ''' <returns></returns>
    ''' <remarks>[Lake] 2012/06/05 Created</remarks>
    Public Function getRoleStaffList(ByVal sBraDepNo As String) As Boolean
        Try
            Dim sb As New StringBuilder()
            sb.Append("SELECT ")
            sb.Append("     '[' + RIGHT(SY_USER.STAFFID,5) + '] ' + SY_USER.USERNAME AS STAFF ")
            sb.Append("    ,SY_ROLE.ROLENAME ")
            sb.Append("FROM ")
            sb.Append("     SY_REL_ROLE_USER JOIN SY_ROLE ON SY_REL_ROLE_USER.ROLEID= SY_ROLE.ROLEID ")
            sb.Append("     JOIN SY_USER ON SY_REL_ROLE_USER.STAFFID = SY_USER.STAFFID ")
            sb.Append("WHERE ")
            sb.Append("     SY_USER.STATUS=0 AND SY_ROLE.DISABLED=0 ")
            sb.Append("     AND SY_REL_ROLE_USER.BRA_DEPNO = '" & sBraDepNo & "' ")
            sb.Append("GROUP BY SY_ROLE.ROLEID,'[' + RIGHT(SY_USER.STAFFID,5) + '] ' + SY_USER.USERNAME,SY_ROLE.ROLENAME")

            Return Me.loadBySQL(sb.ToString())
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function


    ''' <summary>
    ''' 查詢待刪除人員是否已分配角色
    ''' </summary>
    ''' <param name="sStaffId"></param>
    ''' <returns></returns>
    ''' <remarks>[Lake] 2012/06/15 Created</remarks>
    Public Function getByToDeleteStaffId(ByVal sBraDepNo As String, ByVal sStaffId As String) As Boolean
        Try
            Dim sSql As String = "SELECT * FROM SY_REL_ROLE_USER WHERE BRA_DEPNO =" & ProviderFactory.PositionPara & "BRADEPNO " & _
                    "AND STAFFID IN('" & sStaffId & "')"

            Dim paras As IDbDataParameter = ProviderFactory.CreateDataParameter("BRADEPNO", sBraDepNo)

            Return Me.loadBySQL(sSql, paras) > 0
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 根據部門代碼和人員編號
    ''' </summary>
    ''' <param name="sBraDepNo">部門代碼</param>
    ''' <param name="sStaffId">人員編號</param>
    ''' <returns></returns>
    ''' <remarks>[Lake] 2012/06/15 Created</remarks>
    Public Function deleteByBraDepNoAndStaffId(ByVal sBraDepNo As String, ByVal sStaffId As String) As Boolean
        Try
            Dim sSql As String = "delete from sy_rel_role_user where bra_depno in('" & sBraDepNo & "');" & _
                                 "delete from sy_rel_role_user where staffid in('" & sStaffId & "')"

            If Me.getDatabaseManager.isTransaction Then
                Return DBObject.ExecuteNonQuery(Me.getDatabaseManager.getTransaction, CommandType.Text, sSql) > 0
            Else
                Return DBObject.ExecuteNonQuery(Me.getDatabaseManager.getConnection, CommandType.Text, sSql) > 0
            End If
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
    ''' 根據部門編號，角色編號查詢信息
    ''' </summary>
    ''' <param name="sBRA_DEPNO">部門編號</param>
    ''' <param name="sROLEID">角色編號</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Zack 2012-05-15 Create
    ''' </remarks>
    Public Function loadByBRIDRoleID(ByVal sBRA_DEPNO As String, ByVal sROLEID As String) As Integer
        Try
            Dim strSql As String = "SELECT * FROM SY_REL_ROLE_USER " & _
                                   " Join SY_ROLESUITBRANCH ON SY_REL_ROLE_USER.ROLEID = SY_ROLESUITBRANCH.ROLEID" & _
                                   " WHERE SY_ROLESUITBRANCH.BRA_DEPNO =" & ProviderFactory.PositionPara & "BRA_DEPNO" & _
                                   " and SY_REL_ROLE_USER.ROLEID=" & ProviderFactory.PositionPara & "ROLEID"

            Dim para(1) As IDataParameter
            para(0) = ProviderFactory.CreateDataParameter("BRA_DEPNO", sBRA_DEPNO)
            para(1) = ProviderFactory.CreateDataParameter("ROLEID", sROLEID)

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
    ''' 根據人員編號，分行號查詢人員信息
    ''' </summary>
    ''' <param name="sSTAFFID">人員編號</param>
    ''' <param name="sBraDepno">分行號</param>
    ''' <param name="sROLEID">角色編號</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Zack 2012-06-01 Create
    ''' </remarks>
    Public Function loadByStaffidDepnoROLEID(ByVal sStaffid As String, ByVal sBraDepno As String, ByVal sRoleID As String) As Integer
        Try
            Dim sSQL As String = "SELECT " & _
                                    " SY_REL_ROLE_USER.* " & _
                                 " FROM SY_REL_ROLE_USER" & _
                                    " JOIN SY_REL_BRANCH_USER ON SY_REL_ROLE_USER.STAFFID = SY_REL_BRANCH_USER. STAFFID" & _
                                 " WHERE SY_REL_ROLE_USER.STAFFID =" & ProviderFactory.PositionPara & "STAFFID" & _
                                     " AND SY_REL_ROLE_USER.BRA_DEPNO =" & ProviderFactory.PositionPara & "BRA_DEPNO" & _
                                     " AND ROLEID =" & ProviderFactory.PositionPara & "ROLEID"

            Dim para(2) As IDataParameter
            para(0) = ProviderFactory.CreateDataParameter("STAFFID", sStaffid)
            para(1) = ProviderFactory.CreateDataParameter("BRA_DEPNO", sBraDepno)
            para(2) = ProviderFactory.CreateDataParameter("ROLEID", sRoleID)

            Return MyBase.loadBySQL(sSQL, para)
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 根據角色編號查詢人員信息
    ''' </summary>
    ''' <param name="sROLEID">角色編號</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Zack 2012-06-04 Create
    ''' </remarks>
    Public Function loadByRoleID(ByVal sROLEID As String) As Integer
        Try
            Dim sSQL As String = "SELECT '(' +SY_BRANCH.BRCNAME +') '+ SY_USER.STAFFID + ' ' + SY_USER.USERNAME  AS LstName" & _
                                " ,CONVERT(varchar(10), SY_BRANCH.BRA_DEPNO) + ';' + SY_USER.STAFFID AS LstValue" & _
                                " FROM SY_REL_ROLE_USER  join SY_USER on SY_REL_ROLE_USER.STAFFID =SY_USER.STAFFID" & _
                                " INNER JOIN SY_ROLESUITBRANCH ON SY_REL_ROLE_USER.ROLEID = SY_ROLESUITBRANCH.ROLEID" & _
                                " INNER JOIN SY_BRANCH  ON SY_BRANCH.BRA_DEPNO = SY_ROLESUITBRANCH.BRA_DEPNO " & _
                                " WHERE SY_USER.STATUS = 0 And SY_BRANCH.DISABLED = 0" & _
                                " AND SY_REL_ROLE_USER.ROLEID=" & ProviderFactory.PositionPara & "ROLEID"

            Dim para(0) As IDataParameter
            para(0) = ProviderFactory.CreateDataParameter("ROLEID", sROLEID)

            Return MyBase.loadBySQL(sSQL, para)
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

#Region "Avril Function"
    ''' <summary>
    ''' 根據登入者和部門查詢角色編號資料
    ''' </summary>
    ''' <param name="sStaffid">登入者</param>
    ''' <param name="sBraDepNo">部門</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Add by Avril 2012/06/11
    ''' </remarks>
    Public Function getRoleIdByCon(ByVal sStaffid As String, ByVal sBraDepNo As String) As Boolean
        Try
            Dim strSql As String = "SELECT * FROM SY_REL_ROLE_USER " & _
                                   " WHERE STAFFID = " & ProviderFactory.PositionPara & "STAFFID" & _
                                   " AND BRA_DEPNO = " & ProviderFactory.PositionPara & "BRA_DEPNO"

            Dim para(1) As IDataParameter
            para(0) = ProviderFactory.CreateDataParameter("STAFFID", sStaffid)
            para(1) = ProviderFactory.CreateDataParameter("BRA_DEPNO", sBraDepNo)

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
    ''' 根據角色編號查詢人員資料
    ''' </summary>
    ''' <param name="sRoleId">角色編號</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Add by Avril 2012/06/26
    ''' </remarks>
    Public Function loadDataByRoleId(ByVal sRoleId As String) As Integer
        Try
            Dim strSql As String = "SELECT * FROM SY_REL_ROLE_USER " & _
                                   " WHERE ROLEID = " & ProviderFactory.PositionPara & "ROLEID"

            Dim para(0) As IDataParameter
            para(0) = ProviderFactory.CreateDataParameter("ROLEID", sRoleId)

            Return MyBase.loadBySQL(strSql, para)
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
End Class
