Imports com.Azion.NET.VB


Public Class SY_SYSID
    Inherits BosBase

    ''' <summary>
    ''' Constructor
    ''' </summary>
    ''' <param name="dbManager">DatabaseManager</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal dbManager As com.Azion.NET.VB.DatabaseManager)
        MyBase.new("SY_SYSID", dbManager)
    End Sub

#Region "濟南昱勝添加"
#Region "Lake Function"

    Sub New()
        MyBase.New()
        Me.setPrimaryKeys()
    End Sub

    ''' <summary>
    ''' 根據角色查詢系統功能
    ''' </summary>
    ''' <param name="sRoleId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' <history>
    ''' [Lake] 2012/04/11 Created
    '''</history>
    Function loadByRoleId(ByVal sRoleId As String) As Boolean
        Try
            Dim strSql = "SELECT " & _
                         "    DISTINCT TOP 1 sys.SYSNAME AS SYSNAME " & _
                         "   ,sys.SYSID AS SYSID " & _
                         "FROM " & _
                         "   SY_SYSID sys LEFT OUTER JOIN SY_ROLESUITSYS syr " & _
                         "ON " & _
                         "   sys.SYSID = syr.SYSID " & _
                         "WHERE " & _
                         "        syr.ROLEID IN(" & sRoleId & ")" & _
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

    ''' <summary>
    ''' 根據系統編號查詢資料
    ''' </summary>
    ''' <param name="sSysId">系統編號</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Add by Avril 2012/04/26
    ''' </remarks>
    Function loadByPK(ByVal sSysId As String) As Boolean
        Try
            Dim para(0) As IDbDataParameter
            para(0) = ProviderFactory.CreateDataParameter("SYSID", sSysId)

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
End Class

