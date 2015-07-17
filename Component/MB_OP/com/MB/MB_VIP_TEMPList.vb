Imports com.Azion.NET.VB

Public Class MB_VIP_TEMPList
    Inherits BosList

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.New("MB_VIP_TEMP", dbManager)
    End Sub

    Overrides Function newBos() As BosBase
        Dim bos As MB_VIP_TEMP = New MB_VIP_TEMP(MyBase.getDatabaseManager)
        Return bos
    End Function

    ''' <summary>
    ''' 根據"會員編號"取得資料
    ''' </summary>
    ''' <param name="iMB_DATE">會員編號</param>
    ''' <returns>Boolean 是否有取得</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Function LoadByMB_DATE(ByVal iMB_DATE As Long) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT * FROM MB_VIP_TEMP WHERE MB_DATE = " & ProviderFactory.PositionPara & "MB_DATE "

            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_DATE", iMB_DATE)

            Return Me.loadBySQL(sqlStr, para)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function
End Class
