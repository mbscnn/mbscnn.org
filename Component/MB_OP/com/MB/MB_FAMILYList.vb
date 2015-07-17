Imports com.Azion.NET.VB

Public Class MB_FAMILYList
    Inherits BosList

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.New("MB_FAMILY", dbManager)
    End Sub

    Overrides Function newBos() As BosBase
        Dim bos As MB_FAMILY = New MB_FAMILY(MyBase.getDatabaseManager)
        Return bos
    End Function

#Region "Ted Function"
    ''' <summary>
    ''' 根據"會員編號"取得資料
    ''' </summary>
    ''' <param name="iMB_MEMSEQ">會員編號</param>
    ''' <returns>Integer 取得筆數</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Function LoadByMB_MEMSEQ(ByVal iMB_MEMSEQ As Decimal) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT * FROM MB_FAMILY WHERE MB_MEMSEQ = " & ProviderFactory.PositionPara & "MB_MEMSEQ " & _
                     "  AND MB_FAMSEQ <> " & ProviderFactory.PositionPara & "MB_FAMSEQ "

            Dim paras(1) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("MB_MEMSEQ", iMB_MEMSEQ)
            paras(1) = ProviderFactory.CreateDataParameter("MB_FAMSEQ", iMB_MEMSEQ)

            Return Me.loadBySQLOnlyDs(sqlStr, paras)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function
#End Region

End Class
