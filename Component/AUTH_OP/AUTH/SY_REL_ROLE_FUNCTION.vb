Imports com.Azion.NET.VB

Public Class SY_REL_ROLE_FUNCTION
    Inherits BosBase

    ''' <summary>
    ''' Constructor
    ''' </summary>
    ''' <param name="dbManager">DatabaseManager</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal dbManager As com.Azion.NET.VB.DatabaseManager)
        MyBase.new("SY_REL_ROLE_FUNCTION", dbManager)
    End Sub

    Sub New()
        MyBase.New()
        Me.setPrimaryKeys()
    End Sub

#Region "濟南昱勝添加"
#Region "Avril Function"

    ''' <summary>
    ''' 根據角色編號與功能編號查詢資料
    ''' </summary>
    ''' <param name="sRoleId">角色編號</param>
    ''' <param name="sFuncCode">功能編號</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Add by Avril 2012/05/05
    ''' </remarks>
    Function loadByPK(ByVal sRoleId As String, ByVal sSubSysId As String, ByVal sSysIdD As String, ByVal sFuncCode As String) As Boolean
        Try
            Dim para(3) As IDbDataParameter
            para(0) = ProviderFactory.CreateDataParameter("ROLEID", sRoleId)
            para(1) = ProviderFactory.CreateDataParameter("SUBSYSID", sSubSysId)
            para(2) = ProviderFactory.CreateDataParameter("SYSID", sSysIdD)
            para(3) = ProviderFactory.CreateDataParameter("FUNCCODE", sFuncCode)

            Return MyBase.loadBySQL(para)
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

 
    ''' <summary>
    ''' 依角色ID、子系統ID、系統ID查詢交易
    ''' </summary>
    ''' <param name="iRoleId">角色編號</param>
    ''' <param name="sSubSysId">子系統04,05,06,00,SY..</param>
    ''' <param name="sSysId">系統D,F,X,SY..</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function loadFunction(ByVal iRoleId As Integer, ByVal sSubSysId As String, ByVal sSysId As String) As Boolean
        Dim para(2) As IDataParameter
        para(0) = ProviderFactory.CreateDataParameter("ROLEID", iRoleId)
        para(1) = ProviderFactory.CreateDataParameter("SUBSYSID", sSubSysId)
        para(2) = ProviderFactory.CreateDataParameter("SYSID", sSysId)

        Return MyBase.loadBySQL(para)
    End Function



    ''' <summary>
    ''' PK-FK
    ''' 1.刪除沒有在Temp裡的角色與交易的關聯(SY_REL_ROLE_FUNCTION) 
    ''' </summary>
    ''' <param name="iRoleId">角色編號</param>
    ''' <param name="sSubSysId">子系統04,05,06,00,SY...</param>
    ''' <param name="sSysIdD">系統D,F,X,SY...</param>
    ''' <remarks>
    ''' titan by titan 2012/06/24
    ''' </remarks>
    Sub delRoleTask(ByVal iRoleId As Integer, ByVal sSubSysId As String, ByVal sSysIdD As String)
        Try
            '刪除沒有在Temp裡的角色與交易的關聯(SY_REL_ROLE_FUNCTION)
            Dim sSQLDelRoleTask As String = "DELETE " & _
                                "FROM " & _
                                "   SY_REL_ROLE_FUNCTION" & _
                                " WHERE  ROLEID = " & ProviderFactory.PositionPara & "ROLEID and SUBSYSID = " & ProviderFactory.PositionPara & "SUBSYSID and SYSID = " & ProviderFactory.PositionPara & "SYSID"
 
            Dim para(2) As IDbDataParameter
            para(0) = ProviderFactory.CreateDataParameter("ROLEID", iRoleId)
            para(1) = ProviderFactory.CreateDataParameter("SUBSYSID", sSubSysId)
            para(2) = ProviderFactory.CreateDataParameter("SYSID", sSysIdD)

            DBObject.ExecuteNonQuery(Me.getDatabaseManager, CommandType.Text, sSQLDelRoleTask, para)
            
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region
End Class
