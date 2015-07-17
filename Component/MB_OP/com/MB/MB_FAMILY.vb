Imports com.Azion.NET.VB

Public Class MB_FAMILY
    Inherits BosBase

    Sub New()
        MyBase.New()
        Me.setPrimaryKeys()
    End Sub

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.New("MB_FAMILY", dbManager)
    End Sub

#Region "Ted Function"
    ''' <summary>
    ''' 根據"會員編號","家庭成員會員編號"取得資料
    ''' </summary>
    ''' <param name="iMB_MEMSEQ">會員編號</param>
    ''' <param name="iMB_FAMSEQ">家庭成員會員編號</param>
    ''' <returns>Boolean 是否有取得</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Function LoadByPK(ByVal iMB_MEMSEQ As Decimal, ByVal iMB_FAMSEQ As Decimal) As Boolean
        Try
            Dim paras(1) As IDbDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("MB_MEMSEQ", iMB_MEMSEQ)

            paras(1) = ProviderFactory.CreateDataParameter("MB_FAMSEQ", iMB_FAMSEQ)

            Return Me.loadBySQL(paras)
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
