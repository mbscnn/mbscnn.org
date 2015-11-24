Imports com.Azion.NET.VB
Public Class SY_ROLESUITSYSList
    : Inherits BosList

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.new("SY_ROLESUITSYS", dbManager)
    End Sub

    Overrides Function newBos() As BosBase
        Return New SY_ROLESUITSYS(MyBase.getDatabaseManager)
    End Function

#Region "濟南昱勝添加"
#Region "Avril Function"
    ''' <summary>
    ''' 根據角色編號查詢角色維護檔
    ''' </summary>
    ''' <param name="sRoleId">角色編號</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' create by Avril
    ''' modify by titan 2012/06/21
    ''' </remarks>
    Function loadDataByRoleId(ByVal sRoleId As String) As Integer
        Try
            Dim strSql As String = "select a.roleid,a.parent,a.rolename,a.disabled,a.roletype,b.sysid,b.subsysid " & _
                                    "from sy_role a " & _
                                    "left join " & _
                                    "sy_rolesuitsys b " & _
                                    "on a.roleid = b.roleid " & _
                                    "where a.roleid=" & ProviderFactory.PositionPara & "ROLEID order by SYSID,SUBSYSID"

            Dim paras(0) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("ROLEID", sRoleId)

            Return MyBase.loadBySQLOnlyDs(strSql, paras)
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 根據角色編號刪除資料
    ''' </summary>
    ''' <param name="sRoleId">角色編號</param>
    ''' <remarks>
    ''' Add by Avril 2012/04/18
    ''' </remarks>
    Sub delSyData(ByVal sRoleId As String)
        Try
            Dim sSQL As String = "DELETE FROM  SY_ROLESUITSYS  WHERE ROLEID=" & ProviderFactory.PositionPara & "ROLEID"

            Dim paras(0) As System.Data.IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("ROLEID", sRoleId)


            If Me.getDatabaseManager.isTransaction Then
                DBObject.ExecuteNonQuery(Me.getDatabaseManager.getTransaction, CommandType.Text, sSQL, paras)
            Else
                DBObject.ExecuteNonQuery(Me.getDatabaseManager.getConnection, CommandType.Text, sSQL, paras)
            End If
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ''' <summary>
    ''' 根據案號查詢角色歷史檔中的資料
    ''' </summary>
    ''' <param name="sCaseId">案號</param>
    ''' <param name="sRoleId">角色編號</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Add by Avril 2012/05/16
    ''' </remarks>
    Function loadDataByCaseId(ByVal sCaseId As String, ByVal sRoleId As String)
        Try
            Dim strSql As String = "SELECT  " & _
                                   "   Select SY_ROLESUITSYS.* from SY_ROLESUITSYS " & _
                                   "   join sy_flowstep  on sy_flowstep .step_no= SY_ROLESUITSYS. step_no  " & _
                                   "  and sy_flowstep .subflow_seq= SY_ROLESUITSYS. subflow_seq " & _
                                   "   and sy_flowstep.subflow_count= SY_ROLESUITSYS. subflow_count " & _
                                   " where sy_flowstep.status=1   " & _
                                       " and sy_flowstep.caseid = " & ProviderFactory.PositionPara & "CASEID"
            If String.IsNullOrEmpty(sRoleId) Then
                Dim para(0) As IDbDataParameter
                para(0) = ProviderFactory.CreateDataParameter("CASEID", sCaseId)

                Return MyBase.loadBySQL(strSql, para)
            Else
                strSql = strSql & " and sy_role_his.roleid = " & ProviderFactory.PositionPara & "ROLEID"
                Dim para(1) As IDbDataParameter
                para(0) = ProviderFactory.CreateDataParameter("CASEID", sCaseId)
                para(0) = ProviderFactory.CreateDataParameter("ROLEID", sRoleId)

                Return MyBase.loadBySQL(strSql, para)
            End If
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 根據角色編號查詢角色維護檔
    ''' </summary>
    ''' <param name="sRoleId">角色編號</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' create by Avril 2012/07/24
    ''' </remarks>
    Function loadDataByCon(ByVal sRoleId As String, ByVal sSysId As String) As Integer
        Try
            Dim strSql As String = "SELECT * FROM SY_ROLESUITSYS " & _
                                    "WHERE  ROLEID =" & ProviderFactory.PositionPara & "ROLEID" & _
                                    " AND SYSID = " & ProviderFactory.PositionPara & "SYSID"

            Dim paras(1) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("ROLEID", sRoleId)
            paras(1) = ProviderFactory.CreateDataParameter("SYSID", sSysId)

            Return MyBase.loadBySQLOnlyDs(strSql, paras)
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
