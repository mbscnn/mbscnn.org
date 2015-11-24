Imports com.Azion.NET.VB

Public Class SY_BRANCHMGRList
    Inherits BosList

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.New("SY_BRANCHMGR", dbManager)
    End Sub

    Overrides Function newBos() As BosBase
        Dim bos As SY_BRANCHMGR = New SY_BRANCHMGR(MyBase.getDatabaseManager)
        Return bos
    End Function

    Function loadMgDepNo(ByVal iBraDepNo As Integer) As Integer
        Dim paras As IDbDataParameter = ProviderFactory.CreateDataParameter("BRA_DEPNO", iBraDepNo)
        Return MyBase.loadBySQL(paras)
    End Function


#Region "濟南昱勝添加"
#Region "Lake Function"
    ''' <summary>
    ''' 根據傳入單位代碼，
    ''' 查詢管理單位為此單位資料（不包含自己管理自己的資料）
    ''' </summary>
    ''' <param name="sBraDepNo">單位代碼</param>
    ''' <returns></returns>
    ''' <remarks>[Lake] 2012/07/19 Created</remarks>
    Public Function loadBymgrBraDepNo(ByVal sBraDepNo As String) As Boolean
        Try
            Dim sSql As String = "SELECT * FROM SY_BRANCHMGR WHERE MGR_BRADEPNO = " & ProviderFactory.PositionPara & "MGRBRADEPNO " & _
                "AND BRA_DEPNO <> " & ProviderFactory.PositionPara & "BRADEPNO"

            Dim paras(1) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("MGRBRADEPNO", sBraDepNo)
            paras(1) = ProviderFactory.CreateDataParameter("BRADEPNO", sBraDepNo)

            Return Me.loadBySQL(sSql, paras) > 0
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