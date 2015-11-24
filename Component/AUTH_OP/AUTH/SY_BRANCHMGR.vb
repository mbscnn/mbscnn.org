Option Explicit On
Option Strict On

Imports com.Azion.NET.VB
Public Class SY_BRANCHMGR
    Inherits BosBase

    ''' <summary>
    ''' Constructor
    ''' </summary>
    ''' <param name="dbManager">DatabaseManager</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal dbManager As com.Azion.NET.VB.DatabaseManager)
        MyBase.new("SY_BRANCHMGR", dbManager)
    End Sub

#Region "Titan Function"
    ''' <summary>
    ''' 根據主鍵查找資料
    ''' </summary>
    ''' <param name="iBraDepNo">單位編號</param>
    ''' <param name="iMgrBraDepNo">管理單位編號</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function loadByPK(ByVal iBraDepNo As Integer, ByVal iMgrBraDepNo As Integer) As Boolean
        Try
            Dim para(1) As IDbDataParameter
            para(0) = ProviderFactory.CreateDataParameter("BRA_DEPNO", iBraDepNo)
            para(1) = ProviderFactory.CreateDataParameter("MGR_BRADEPNO", iMgrBraDepNo)

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

End Class
