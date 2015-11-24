Imports com.Azion.NET.VB

Public Class SY_REL_SYSID_SUBSYSID
    Inherits BosBase

    Sub New()
        MyBase.New()
        Me.setPrimaryKeys()
    End Sub

    ''' <summary>
    ''' Constructor
    ''' </summary>
    ''' <param name="dbManager">DatabaseManager</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal dbManager As com.Azion.NET.VB.DatabaseManager)
        MyBase.new("SY_REL_SYSID_SUBSYSID", dbManager)
    End Sub

#Region "濟南昱勝添加"
#Region "Avril Function"

    ''' <summary>
    ''' 根據子系統編號查詢資料
    ''' </summary>
    ''' <param name="sSubSysId">子系統編號</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Add by Avril 2012/05/10
    ''' </remarks>
    Function loadBySubSysId(ByVal sSubSysId As String) As Boolean
        Try

            Dim syRole As New SY_ROLE(Me.getDatabaseManager)
            Dim sSQL As String = "SELECT *  FROM SY_REL_SYSID_SUBSYSID WHERE SUBSYSID = " & ProviderFactory.PositionPara & "SUBSYSID"

            Dim paras(0) As System.Data.IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("SUBSYSID", sSubSysId)

            Return MyBase.loadBySQL(sSQL, paras)
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


