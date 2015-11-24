Imports com.Azion.NET.VB
Imports System.Text

Public Class SY_REL_ROLE_USER_HISList
    : Inherits BosList

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.new("SY_REL_ROLE_USER_HIS", dbManager)
    End Sub

    Overrides Function newBos() As BosBase
        Return New SY_REL_ROLE_USER_HIS(MyBase.getDatabaseManager)
    End Function

#Region "濟南昱勝添加"
#Region "Lake Function"
    ''' <summary>
    ''' 根據當前登錄人員所在部門及下屬部門查詢SY_FLOWINCIDENT表中ENDTIME
    ''' </summary>
    ''' <param name="sBraDepNo"></param>
    ''' <returns></returns>
    ''' <remarks>[Lake] 2012/06/05 Created</remarks>
    Public Function getEndTimeList(ByVal sBraDepNo As String) As Boolean
        Try
            Dim sb As New StringBuilder()
            sb.Append("SELECT ")
            sb.Append("    DISTINCT YEAR(SY_FLOWINCIDENT.ENDTIME) AS ENDYEAR ")
            sb.Append("FROM ")
            sb.Append("    SY_REL_ROLE_USER_HIS ")
            sb.Append("    JOIN SY_FLOWINCIDENT on SY_FLOWINCIDENT.CASEID= SY_REL_ROLE_USER_HIS.CASEID ")
            sb.Append("WHERE ")
            sb.Append("       SY_REL_ROLE_USER_HIS.BRA_DEPNO ='" & sBraDepNo & "' ")
            sb.Append("   AND SY_REL_ROLE_USER_HIS.OPERATION <> 'N' ")
            sb.Append("   AND SY_REL_ROLE_USER_HIS.APPROVED = 'Y'")

            Return Me.loadBySQL(sb.ToString())
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
