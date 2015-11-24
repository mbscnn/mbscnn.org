Option Explicit On
Option Strict On

Imports com.Azion.NET.VB


Public Class SY_ROLE
    Inherits BosBase


    ''' <summary>
    ''' Constructor
    ''' </summary>
    ''' <param name="dbManager">DatabaseManager</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal dbManager As com.Azion.NET.VB.DatabaseManager)
        MyBase.new("SY_ROLE", dbManager)
    End Sub

    Sub New()
        MyBase.New()
        Me.setPrimaryKeys()
    End Sub

#Region "濟南昱勝添加"
#Region "Avril Function"

    ''' <summary>
    ''' 根據主鍵查找資料
    ''' </summary>
    ''' <param name="sRoleId">角色編號</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Add by Avril 2012/04/18
    ''' </remarks>
    Function loadByPK(ByVal sRoleId As String) As Boolean
        Try
            Dim para(0) As IDbDataParameter
            para(0) = ProviderFactory.CreateDataParameter("ROLEID", sRoleId)

            Return MyBase.loadBySQL(para)
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
    ''' <param name="sRoleNo">角色編號</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Add by Avril 2012/04/24
    ''' </remarks>
    Function loadByRoleNo(ByVal sRoleNo As String) As Boolean
        Try
            Dim sSQL As String = "SELECT *  FROM SY_ROLE WHERE ROLENO = " & ProviderFactory.PositionPara & "ROLENO "

            Dim paras(0) As System.Data.IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("ROLENO", sRoleNo)

            Return MyBase.loadBySQL(sSQL, paras)
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 取得最大的RoleId
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    '''  Add by Avril 2012/04/24
    ''' </remarks>
    Function getMaxRoleId() As Integer
        Try
            If Not IsNothing(Me.getDatabaseManager) Then
                Dim syRole As New SY_ROLE(Me.getDatabaseManager)
                Dim sSQL As String = "SELECT MAX(ROLEID) ROLEID  FROM SY_ROLE WHERE ROLETYPE<>'1'  "

                If (syRole.loadBySQL(sSQL)) Then

                    ' 如果“Roleid”不為Nothing
                    If Not syRole.getAttribute("ROLEID") Is Nothing Then
                        If Convert.ToInt32(syRole.getAttribute("ROLEID").ToString()) > 0 Then
                            Return Convert.ToInt32(syRole.getAttribute("ROLEID").ToString())
                        Else
                            Return 1
                        End If
                    Else
                        Return 1
                    End If
                End If
            End If

            Return 1
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 根據資料查詢角色名稱是否存在
    ''' </summary>
    ''' <param name="sRoleName">角色名稱</param>
    ''' <param name="sRoleId">角色編號</param>
    ''' <param name="sRoleIdList">臨時表中的角色編號集合</param>
    ''' <returns></returns>
    ''' <remarks>
    '''  Add by Avril 2012/05/10
    ''' </remarks>
    Function loadDataByCon(ByVal sRoleName As String, ByVal sRoleId As String, ByVal sRoleIdList As String) As Boolean
        Try
            Dim sSQL As String = "SELECT *  FROM SY_ROLE WHERE ROLENAME = " & ProviderFactory.PositionPara & "ROLENAME " & _
                                " AND ROLEID <> " & ProviderFactory.PositionPara & "ROLEID "

            If sRoleIdList.Length > 0 Then
                sSQL = sSQL & " AND ROLEID  IN( " & sRoleIdList & ") "
            End If

            Dim paras(1) As System.Data.IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("ROLENAME", sRoleName)
            paras(1) = ProviderFactory.CreateDataParameter("ROLEID", sRoleId)

            Return MyBase.loadBySQL(sSQL, paras)
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 根據RoleName查詢角色檔
    ''' </summary>
    ''' <param name="sRoleName">角色名稱</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Add by Avril 2012/05/17
    ''' </remarks>
    Function loadByRoleName(ByVal sRoleName As String) As Boolean
        Try
            Dim para(0) As IDbDataParameter
            para(0) = ProviderFactory.CreateDataParameter("ROLENAME", sRoleName)

            Return MyBase.loadBySQL(para)
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function


    ''' <summary>
    ''' 根據功能名稱查詢資料
    ''' </summary>
    ''' <param name="sRoleName">角色名稱</param>
    ''' <param name="sRoleId">角色編號</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Add by Avril 2012/06/19
    ''' </remarks>
    Function loadByName(ByVal sRoleName As String, ByVal sRoleId As String) As Boolean
        Try
            Dim sSQL As String = "SELECT *  FROM SY_ROLE WHERE ROLENAME = " & ProviderFactory.PositionPara & "ROLENAME " & _
                                    " AND ROLEID <> " & ProviderFactory.PositionPara & "ROLEID "

            Dim paras(1) As System.Data.IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("ROLENAME", sRoleName)
            paras(1) = ProviderFactory.CreateDataParameter("ROLEID", sRoleId)

            Return MyBase.loadBySQL(sSQL, paras)
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 根據角色編號 更新相關的狀態欄位
    ''' </summary>
    ''' <param name="sRoleIdList"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function updateDisabledByRoleIdList(ByVal sRoleIdList As String, ByVal sDisabled As String) As Boolean
        Try
            Dim sSql As String = String.Empty

            sSql = "UPDATE SY_ROLE SET DISABLED = " & ProviderFactory.PositionPara & "DISABLED " & _
                "WHERE ROLEID IN  ( " & sRoleIdList & ") "

            Dim paras(1) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("DISABLED", sDisabled)

            If Me.getDatabaseManager.isTransaction Then
                Return DBObject.ExecuteNonQuery(Me.getDatabaseManager.getTransaction, CommandType.Text, sSql, paras) > 0
            Else
                Return DBObject.ExecuteNonQuery(Me.getDatabaseManager.getConnection, CommandType.Text, sSql, paras) > 0
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

#Region "Titan function"
    ''' <summary>
    ''' PK-FK
    ''' 1.刪除沒有在Temp裡的角色與交易的關聯(SY_REL_ROLE_FUNCTION)
    ''' 2.刪除角色與子系統、系統的關聯(SY_ROLESUITSYS)
    ''' 3.刪除角色
    ''' </summary>
    ''' <param name="iRoleId">角色編號</param>
    ''' <param name="sSubSysId">子系統04,05,06,00,SY...</param>
    ''' <param name="sSysIdD">系統D,F,X,SY...</param> 
    ''' <remarks>
    ''' created by titan 2012/06/24
    ''' </remarks>
    Sub delData(ByVal iRoleId As Integer, ByVal sSubSysId As String, ByVal sSysIdD As String, Optional ByVal bDelRole As Boolean = False)
        Try
            Dim funcs As New SY_ROLESUITSYS(Me.getDatabaseManager)
            funcs.delRoleSys(iRoleId, sSubSysId, sSysIdD)

            '3刪除角色
            Dim sSQLDelRole As String = "DELETE " & _
                                  "FROM " & _
                                  "   SY_ROLE" & _
                                   " WHERE  ROLEID = " & ProviderFactory.PositionPara & "ROLEID"


            Dim para(0) As IDbDataParameter
            para(0) = ProviderFactory.CreateDataParameter("ROLEID", iRoleId)

            If bDelRole Then
                DBObject.ExecuteNonQuery(Me.getDatabaseManager, CommandType.Text, sSQLDelRole, para)
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

End Class
