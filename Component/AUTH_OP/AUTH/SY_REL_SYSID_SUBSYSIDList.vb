Imports com.Azion.NET.VB
Imports AUTH_OP.TABLE
Public Class SY_REL_SYSID_SUBSYSIDList
    : Inherits BosList

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.new("SY_REL_SYSID_SUBSYSID", dbManager)
    End Sub

    Overrides Function newBos() As BosBase
        Return New SY_REL_SYSID_SUBSYSID(MyBase.getDatabaseManager)
    End Function


#Region "濟南昱勝添加"
#Region "Avril Function"
    ''' <summary>
    ''' 根據角色和系統代碼查詢相關的子系統資料
    ''' </summary>
    ''' <param name="sRoleId">角色</param>
    ''' <param name="sSysId">系統代碼</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Add by Avril 2012/04/11
    ''' </remarks>
    Function loadSubSysById(ByVal sRoleId As String, ByVal sSysId As String) As Boolean
        Try
            Dim strSql As String = "SELECT  DISTINCT" & _
                                   "    SY_SUBSYSID .SUBSYSNAME," & _
                                   "    SY_REL_SYSID_SUBSYSID. SYSID, SY_REL_SYSID_SUBSYSID. SUBSYSID " & _
                                   " FROM SY_REL_SYSID_SUBSYSID JOIN SY_SUBSYSID " & _
                                   "   ON SY_REL_SYSID_SUBSYSID. SUBSYSID = SY_SUBSYSID. SUBSYSID " & _
                                   " JOIN SY_ROLESUITSYS   " & _
                                   "  ON SY_REL_SYSID_SUBSYSID. SYSID = SY_ROLESUITSYS. SYSID   " & _
                                   "  AND SY_REL_SYSID_SUBSYSID. SUBSYSID = SY_ROLESUITSYS. SUBSYSID  " & _
                                   " WHERE SY_SUBSYSID .DISABLED=0 " & _
                                       " and SY_REL_SYSID_SUBSYSID. SYSID = " & ProviderFactory.PositionPara & "SYSID" & _
                                    "     and     SY_ROLESUITSYS.ROLEID IN(" & sRoleId & ")"
            Dim para(0) As IDbDataParameter
            para(0) = ProviderFactory.CreateDataParameter("SYSID", sSysId)

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
