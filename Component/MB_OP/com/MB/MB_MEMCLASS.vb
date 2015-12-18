Imports com.Azion.NET.VB

Public Class MB_MEMCLASS
    Inherits BosBase

    Sub New()
        MyBase.New()
        Me.setPrimaryKeys()
    End Sub

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.New("MB_MEMCLASS", dbManager)
    End Sub

#Region "Ted Function"
    ''' <summary>
    ''' 根據 "學員編號" , "課程序號" 取得資料
    ''' </summary>
    ''' <param name="sMB_MEMSEQ">學員編號</param>
    ''' <param name="sMB_SEQ">課程序號</param>
    ''' <returns>Boolean 是否有取得</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Function LoadByPK(ByVal sMB_MEMSEQ As String, ByVal sMB_SEQ As String) As Boolean
        Try
            Dim paras(1) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("MB_MEMSEQ", sMB_MEMSEQ)
            paras(1) = ProviderFactory.CreateDataParameter("MB_SEQ", sMB_SEQ)
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
