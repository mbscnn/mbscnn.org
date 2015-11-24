Imports com.Azion.NET.VB

Public Class SY_REL_ROLE_FUNCTIONList
    : Inherits BosList

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.new("SY_REL_ROLE_FUNCTION", dbManager)
    End Sub

    Overrides Function newBos() As BosBase
        Return New SY_REL_ROLE_FUNCTION(MyBase.getDatabaseManager)
    End Function

#Region "濟南昱勝添加"
#Region "Avril Function"

    ''' <summary>
    ''' 根據角色編號查詢角色維護檔
    ''' </summary>
    ''' <param name="sRoleId">角色編號</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Add by Avril 2012/04/27
    ''' </remarks>
    Function loadDataByRoleId(ByVal sRoleId As String, ByVal sSysId As String) As Boolean
        Try
            Dim strSql As String = "SELECT" _
                                 & "   * " _
                                 & "FROM" _
                                 & "   SY_REL_ROLE_FUNCTION " _
                                 & "WHERE" _
                                 & "       ROLEID = " & ProviderFactory.PositionPara & "ROLEID " _
                                 & " AND SYSID = " & ProviderFactory.PositionPara & "SYSID"

            Dim para(1) As IDataParameter
            para(0) = ProviderFactory.CreateDataParameter("ROLEID", sRoleId)
            para(1) = ProviderFactory.CreateDataParameter("SYSID", sSysId)

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
    ''' <param name="sRoleId"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Add by Avril 2012/07/06
    ''' </remarks>
    Function loadAllDataByRoleId(ByVal sRoleId As String, ByVal sSysId As String, ByVal sSubSysId As String) As Boolean
        Try
            Dim strSql As String = "SELECT" _
                                 & "   * " _
                                 & "FROM" _
                                 & "   SY_REL_ROLE_FUNCTION " _
                                 & "WHERE" _
                                 & "       ROLEID = " & ProviderFactory.PositionPara & "ROLEID " _
                                 & "   AND SYSID = " & ProviderFactory.PositionPara & "SYSID" _
                                 & "  AND SUBSYSID = " & ProviderFactory.PositionPara & "SUBSYSID"


            Dim para(2) As IDataParameter
            para(0) = ProviderFactory.CreateDataParameter("ROLEID", sRoleId)
            para(1) = ProviderFactory.CreateDataParameter("SYSID", sSysId)
            para(2) = ProviderFactory.CreateDataParameter("SUBSYSID", sSubSysId)

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
    ''' 根據角色集合，系統，子系統查詢資料
    ''' </summary>
    ''' <param name="sRoleIdList">角色集合</param>
    ''' <param name="sSysId">系統</param>
    ''' <param name="sSubSysId">子系統</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Add by Avril 2012/07/06
    ''' </remarks>
    Function loadAddDataByCon(ByVal sRoleIdList As String, ByVal sSysId As String, ByVal sSubSysId As String) As Boolean
        Try
            Dim strSql As String = "SELECT" _
                                 & "   * " _
                                 & "FROM" _
                                 & "   SY_REL_ROLE_FUNCTION " _
                                 & "WHERE" _
                                 & "  ROLEID  IN( " & sRoleIdList & ") " _
                                 & "   AND SYSID = " & ProviderFactory.PositionPara & "SYSID" _
                                 & "  AND SUBSYSID = " & ProviderFactory.PositionPara & "SUBSYSID"


            Dim para(1) As IDataParameter
            para(0) = ProviderFactory.CreateDataParameter("SYSID", sSysId)
            para(1) = ProviderFactory.CreateDataParameter("SUBSYSID", sSubSysId)

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
    ''' 根據主鍵進行查詢
    ''' </summary>
    ''' <param name="sRoleId">角色</param>
    ''' <param name="sSysId">系統編號</param>
    ''' <param name="sSubSysId">子系統編號</param>
    ''' <param name="sFunccode">交易項目</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function loadDataByPK(ByVal sRoleId As String, ByVal sSysId As String, ByVal sSubSysId As String, ByVal sFunccode As String) As Boolean
        Try
            Dim strSql As String = "SELECT" _
                                 & "   * " _
                                 & "FROM" _
                                 & "   SY_REL_ROLE_FUNCTION " _
                                 & "WHERE" _
                                 & "  ROLEID  = " & ProviderFactory.PositionPara & "ROLEID" _
                                 & "   AND SYSID = " & ProviderFactory.PositionPara & "SYSID" _
                                 & "  AND SUBSYSID = " & ProviderFactory.PositionPara & "SUBSYSID" _
                                 & " AND FUNCCODE  = " & ProviderFactory.PositionPara & "FUNCCODE"


            Dim para(3) As IDbDataParameter
            para(0) = ProviderFactory.CreateDataParameter("ROLEID", sRoleId)
            para(1) = ProviderFactory.CreateDataParameter("SYSID", sSysId)
            para(2) = ProviderFactory.CreateDataParameter("SUBSYSID", sSubSysId)
            para(3) = ProviderFactory.CreateDataParameter("FUNCCODE", sFunccode)

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
