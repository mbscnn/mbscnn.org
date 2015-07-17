Imports com.Azion.NET.VB

Public Class MB_MEMBER_TEMP
    Inherits BosBase

    Sub New()
        MyBase.New()
        Me.setPrimaryKeys()
    End Sub

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.New("MB_MEMBER_TEMP", dbManager)
    End Sub

#Region "Ted Function"
    ''' <summary>
    ''' 根據"YYYMMDDHHMISS000"取得資料
    ''' </summary>
    ''' <param name="iYYYMMDDHHMISS000">YYYMMDDHHMISS000</param>
    ''' <returns>Boolean 是否有取得</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Function LoadByPK(ByVal iYYYMMDDHHMISS000 As Long) As Boolean
        Try
            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_DATE", iYYYMMDDHHMISS000)

            Return Me.loadBySQL(para)
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
