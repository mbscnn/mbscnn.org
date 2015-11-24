Imports com.Azion.NET.VB

Public Class SY_CASEID
    Inherits BosBase

    Public Sub New(ByVal dbManager As com.Azion.NET.VB.DatabaseManager)
        MyBase.new("SY_CASEID", dbManager)
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
    ''' <param name="sCaseId">案件編號</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Add by Avril 2012/05/11
    ''' </remarks>
    Function loadByPK(ByVal sCaseId As String) As Boolean
        Try
            Dim para(0) As IDbDataParameter
            para(0) = ProviderFactory.CreateDataParameter("CASEID", sCaseId)

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
