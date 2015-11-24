Imports com.Azion.NET.VB
Public Class SY_ROLE_HISList
    : Inherits BosList

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.new("SY_ROLE_HIS", dbManager)
    End Sub

    Overrides Function newBos() As BosBase
        Return New SY_ROLE_HIS(MyBase.getDatabaseManager)
    End Function


#Region "濟南昱勝添加"
#Region "Avril Function"

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
                                   "    sy_role_his.* from sy_role_his " & _
                                   "   join sy_flowstep  on sy_flowstep .step_no= sy_role_his. step_no  " & _
                                   "  and sy_flowstep .subflow_seq= sy_role_his. subflow_seq " & _
                                   "   and sy_flowstep.subflow_count= sy_role_his. subflow_count " & _
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
#End Region
#End Region
End Class
