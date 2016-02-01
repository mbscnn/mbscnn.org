Imports com.Azion.NET.VB

Public Class MB_YENTList
    Inherits BosList

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.New("MB_YENT", dbManager)
    End Sub

    Overrides Function newBos() As BosBase
        Dim bos As MB_YENT = New MB_YENT(MyBase.getDatabaseManager)
        Return bos
    End Function

#Region "Ted Function"
    Public Function Load_SDATE(ByVal D_S_SDATE As Date, ByVal D_E_SDATE As Date)
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT * FROM MB_YENT " &
                     "  WHERE SDATE >= " & ProviderFactory.PositionPara & "S_SDATE " &
                     "      AND SDATE <= " & ProviderFactory.PositionPara & "E_SDATE " &
                     "  ORDER BY SDATE,EDATE"

            Dim paras(1) As IDbDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("S_SDATE", D_S_SDATE)
            paras(1) = ProviderFactory.CreateDataParameter("E_SDATE", D_E_SDATE)

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
