Imports com.Azion.NET.VB
Public Class AP_POST
    Inherits BosBase

    Sub New()
        MyBase.New()
        Me.setPrimaryKeys()
    End Sub

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.new("AP_POST", dbManager)
    End Sub
#Region "CHRIS"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 依PK帶出該行資料
    ''' </summary>
    ''' <returns>Boolean</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[chris]	2011/8/30	Created
    ''' </history>
    Function LoadByPK(ByVal sCITY As String, ByVal sAREA As String, ByVal sROAD As String) As Boolean
        Try
            If Utility.isValidateData(sCITY) AndAlso Utility.isValidateData(sAREA) Then
                Dim paras(2) As System.Data.IDbDataParameter
                paras(0) = ProviderFactory.CreateDataParameter("CITY", sCITY)
                paras(1) = ProviderFactory.CreateDataParameter("AREA", sAREA)
                paras(2) = ProviderFactory.CreateDataParameter("ROAD", sROAD)

                Return MyBase.loadBySQL(paras)
            End If
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
