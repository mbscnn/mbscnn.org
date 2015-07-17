Imports com.Azion.NET.VB

Public Class MB_CLASS_M
    Inherits BosBase

    Sub New()
        MyBase.New()
        Me.setPrimaryKeys()
    End Sub

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.New("MB_CLASS_M", dbManager)
    End Sub

#Region "Ted Function"
    ''' <summary>
    ''' 根據"課程序號"取得資料
    ''' </summary>
    ''' <param name="iMB_SEQ">MB_SEQ課程序號</param>
    ''' <returns>Boolean 是否有取得</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Function LoadByPK(ByVal iMB_SEQ As Decimal) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT * FROM MB_CLASS_M WHERE MB_SEQ = " & ProviderFactory.PositionPara & "MB_SEQ "

            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_SEQ", iMB_SEQ)

            Return Me.loadBySQL(sqlStr, para)
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
