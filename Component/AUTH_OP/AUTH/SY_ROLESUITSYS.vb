Option Explicit On
Option Strict On

Imports com.Azion.NET.VB


Public Class SY_ROLESUITSYS
    Inherits BosBase


    ''' <summary>
    ''' Constructor
    ''' </summary>
    ''' <param name="dbManager">DatabaseManager</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal dbManager As com.Azion.NET.VB.DatabaseManager)
        MyBase.new("SY_ROLESUITSYS", dbManager)
    End Sub


#Region "濟南昱勝添加"
#Region "Avril Function"

    ''' <summary>
    ''' 根據主鍵查詢資料
    ''' </summary>
    ''' <param name="sRoleId">角色編號</param>
    ''' <param name="sSubSysId">子系統編號</param>
    ''' <param name="sSysId">系統編號</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function LoadByPK(ByVal sRoleId As String, ByVal sSubSysId As String, ByVal sSysId As String) As Boolean
        Try
            Dim paras(2) As IDbDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("ROLEID", sRoleId)

            paras(1) = ProviderFactory.CreateDataParameter("SUBSYSID", sSubSysId)

            paras(2) = ProviderFactory.CreateDataParameter("SYSID", sSysId)

            Return Me.loadBySQL(paras)
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 根據RoleId刪除資料
    ''' </summary>
    ''' <param name="sRoleId">角色編號</param>
    ''' <remarks>
    ''' Add by Avril 2012/04/26
    ''' </remarks>
    Sub delDataByRoleId(ByVal sRoleId As String)
        Try
            Dim strSql As String = "DELETE " & _
                                   "FROM " & _
                                   "   SY_ROLESUITSYS" & _
                                   " WHERE SY_ROLESUITSYS.ROLEID = " & ProviderFactory.PositionPara & "ROLEID"

            Dim para(0) As IDbDataParameter
            para(0) = ProviderFactory.CreateDataParameter("ROLEID", sRoleId)

            If Me.getDatabaseManager.isTransaction Then
                DBObject.ExecuteNonQuery(Me.getDatabaseManager.getTransaction, CommandType.Text, strSql, para)
            Else
                DBObject.ExecuteNonQuery(Me.getDatabaseManager.getConnection, CommandType.Text, strSql, para)
            End If
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region
#End Region


    ''' <summary>
    ''' PK-FK
    ''' 刪除角色與子系統、系統的關聯(SY_ROLESUITSYS)
    ''' </summary>
    ''' <param name="iRoleId">角色編號</param>
    ''' <param name="sSubSysId">子系統04,05,06,00,SY...</param>
    ''' <param name="sSysIdD">系統D,F,X,SY...</param>
    ''' <remarks>
    ''' titan by titan 2012/06/24
    ''' </remarks>
    Sub delRoleSys(ByVal iRoleId As Integer, ByVal sSubSysId As String, ByVal sSysIdD As String)
        Try
            
            '刪除角色與子系統、系統的關聯(SY_ROLESUITSYS)
            Dim sSQLDelRoleSys As String = "DELETE " & _
                                   "FROM " & _
                                   "   SY_ROLESUITSYS" & _
                                    " WHERE  ROLEID = " & ProviderFactory.PositionPara & "ROLEID and SUBSYSID = " & ProviderFactory.PositionPara & "SUBSYSID and SYSID = " & ProviderFactory.PositionPara & "SYSID"


            Dim para(2) As IDbDataParameter
            para(0) = ProviderFactory.CreateDataParameter("ROLEID", iRoleId)
            para(1) = ProviderFactory.CreateDataParameter("SUBSYSID", sSubSysId)
            para(2) = ProviderFactory.CreateDataParameter("SYSID", sSysIdD)

            DBObject.ExecuteNonQuery(Me.getDatabaseManager, CommandType.Text, sSQLDelRoleSys, para)

        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
End Class

