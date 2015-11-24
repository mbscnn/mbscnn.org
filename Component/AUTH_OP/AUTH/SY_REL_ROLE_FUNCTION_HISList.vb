Imports com.Azion.NET.VB
Public Class SY_REL_ROLE_FUNCTION_HISList
    : Inherits BosList

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.new("SY_REL_ROLE_FUNCTION_HIS", dbManager)
    End Sub

    Overrides Function newBos() As BosBase
        Return New SY_REL_ROLE_FUNCTION_HIS(MyBase.getDatabaseManager)
    End Function

#Region "濟南昱勝添加"
#Region "Avril Function"

    ''' <summary>
    ''' 根據角色編號，交易編號，系統，子系統查詢資料
    ''' </summary>
    ''' <param name="sRoleId">角色編號</param>
    ''' <param name="sFuncCode">交易編號</param>
    ''' <param name="sSysId">系統</param>
    ''' <param name="sSubSysId">子系統</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Add by Avril 2012/06/08
    ''' </remarks>
    Function loadDataByCon(ByVal sRoleId As String, ByVal sFuncCode As String, ByVal sSysId As String, ByVal sSubSysId As String) As Integer
        Try
            Dim strSql As String = "SELECT" _
                                 & "   * " _
                                 & "FROM" _
                                 & "   SY_REL_ROLE_FUNCTION_HIS " _
                                 & "WHERE" _
                                 & "       ROLEID = " & ProviderFactory.PositionPara & "ROLEID" _
                                 & "     AND FUNCCODE = " & ProviderFactory.PositionPara & "FUNCCODE" _
                                 & "     AND SYSID = " & ProviderFactory.PositionPara & "SYSID" _
                                 & "     AND SUBSYSID = " & ProviderFactory.PositionPara & "SUBSYSID"

            Dim para(3) As IDataParameter
            para(0) = ProviderFactory.CreateDataParameter("ROLEID", sRoleId)
            para(1) = ProviderFactory.CreateDataParameter("FUNCCODE", sFuncCode)
            para(2) = ProviderFactory.CreateDataParameter("SYSID", sSysId)
            para(3) = ProviderFactory.CreateDataParameter("SUBSYSID", sSubSysId)

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
