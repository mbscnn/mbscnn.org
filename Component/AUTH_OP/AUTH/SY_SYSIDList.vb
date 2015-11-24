Imports com.Azion.NET.VB
Imports AUTH_OP.TABLE

Public Class SY_SYSIDList
    : Inherits BosList

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.new("SY_SYSID", dbManager)
    End Sub

    Overrides Function newBos() As BosBase
        Return New SY_SYSID(MyBase.getDatabaseManager)
    End Function

#Region "濟南昱勝添加"
#Region "Lake Function"

    ''' <summary>
    ''' 根據登錄人員角色取得系統功能
    ''' </summary>
    ''' <param name="sWorkingRoleId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' <history>
    ''' [Lake] 2012/04/06 Created
    '''</history>
    Function loadByWorkingRoleId(ByVal sWorkingRoleId As String) As Integer
        Try
            Dim strSql As String = "SELECT " & _
                                   "   DISTINCT sys.SYSNAME AS SYSNAME, sys.SYSID AS SYSID " & _
                                   "FROM " & _
                                   "   SY_SYSID sys LEFT OUTER JOIN SY_ROLESUITSYS syr " & _
                                   "ON " & _
                                   "   sys.SYSID = syr.SYSID " & _
                                   "WHERE " & _
                                   "        syr.ROLEID IN(" & sWorkingRoleId & ") " & _
                                   " and exists (select * from SY_REL_ROLE_FUNCTION a left join sy_function_code b on a.funccode=b.funccode where a.ROLEID=syr.ROLEID and a.SUBSYSID=syr.SUBSYSID and a.SYSID=syr.SYSID and b.disabled='0' and  isnull(b.fucurl,'')<>'' ) " & _
                                   "    AND sys.DISABLED=0 " & _
                                   "ORDER BY sys.SYSID"

            Return MyBase.loadBySQL(strSql)
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

#Region "Avril Function"

    ''' <summary>
    ''' 查詢可用的系統資料
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' Add by Avril 2012/04/18
    ''' </remarks>
    Function loadUseData() As Boolean
        Try
            Dim strSql = "SELECT " & _
                         "    * FROM SY_SYSID " & _
                         " WHERE " & _
                         "     DISABLED=0 "

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
    ''' 查詢某一系統的資料
    ''' </summary>
    ''' <param name="sSysIdList">系統編號</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function loadDataBySysId(ByVal sSysIdList As String) As Boolean
        Try
            Dim strSql = "SELECT " & _
                         "    * FROM SY_SYSID " & _
                         " WHERE " & _
                         "     DISABLED=0 " & _
                         " AND SYSID IN (" & sSysIdList & ")"


            Return MyBase.loadBySQL(strSql)
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
#End Region
End Class
