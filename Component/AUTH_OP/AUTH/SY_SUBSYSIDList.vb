Imports com.Azion.NET.VB
Public Class SY_SUBSYSIDList
    : Inherits BosList

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.new("SY_SUBSYSID", dbManager)
    End Sub

    Overrides Function newBos() As BosBase
        Return New SY_SUBSYSID(MyBase.getDatabaseManager)
    End Function


#Region "濟南昱勝添加"
#Region "Avril Function"

    ''' <summary>
    ''' 查詢子系統
    ''' </summary>
    ''' <param name="sSysId">系統編號</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Add by Avril 2012/04/18
    ''' </remarks>
    Function loadSubSys(ByRef sSysId As String) As Boolean
        Try
            Dim strSql = "SELECT " & _
                         "    SY_SUBSYSID .*  from SY_SUBSYSID join SY_REL_SYSID_SUBSYSID " & _
                         " On SY_SUBSYSID. SUBSYSID = SY_REL_SYSID_SUBSYSID. SUBSYSID " & _
                         "    Where DISABLED=0  AND SY_REL_SYSID_SUBSYSID.SYSID=  " & ProviderFactory.PositionPara & "SYSID"

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

    ''' <summary>
    ''' 查詢子系統
    ''' </summary>
    ''' <param name="sSysId">系統編號</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Add by Avril 2012/04/18
    ''' </remarks>
    Function loadSubSysList(ByRef sSysId As String, ByVal sSubSysId As String) As Boolean
        Try
            Dim strSql = "SELECT " & _
                         "    SY_SUBSYSID .*  from SY_SUBSYSID join SY_REL_SYSID_SUBSYSID " & _
                         " On SY_SUBSYSID. SUBSYSID = SY_REL_SYSID_SUBSYSID. SUBSYSID " & _
                         "    Where DISABLED=0  AND SY_REL_SYSID_SUBSYSID.SYSID=  " & ProviderFactory.PositionPara & "SYSID" & _
                         " AND SY_REL_SYSID_SUBSYSID.SUBSYSID IN (" & sSubSysId & ")"

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

#Region "Titan Function"
    ''' <summary>
    ''' 查詢子系統
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' Add by titan 2012/06/20
    ''' </remarks>
    Function loadAllSubSys() As Integer
        Try
            Dim strSql = "SELECT " & _
                         "    SY_SUBSYSID.*, SY_REL_SYSID_SUBSYSID.SYSID,  " & _
                         "    case SY_SUBSYSID.SUBSYSID " & _
                         "    when '06' then " & _
                         "         case SY_REL_SYSID_SUBSYSID.SYSID" & _
                         "         when 'D' then '法金' + SY_SUBSYSID.SHORTSSNAME" & _
                         "         when 'F' then '消金' + SY_SUBSYSID.SHORTSSNAME" & _
                         "         End" & _
                         "    else SY_SUBSYSID.SHORTSSNAME" & _
                         "    end NewShortNM" & _
                         " from SY_SUBSYSID join SY_REL_SYSID_SUBSYSID " & _
                         " On SY_SUBSYSID. SUBSYSID = SY_REL_SYSID_SUBSYSID. SUBSYSID " & _
                         "    Where DISABLED=0 "
            Return MyBase.loadBySQLOnlyDs(strSql)
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region
End Class
