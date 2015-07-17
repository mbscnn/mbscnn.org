
Imports com.Azion.NET.VB

Public Class SUBSYS_CODEList
    : Inherits BosList

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.New("SUBSYS_CODE", dbManager)
    End Sub

    Overrides Function newBos() As BosBase
        Dim bos As SUBSYS_CODE = New SUBSYS_CODE(MyBase.getDatabaseManager)
        Return bos
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' loadDemo 絛ㄒ
    ''' </summary>
    ''' <param name="_sPara1">把计1</param>
    ''' <param name="_iPara2">把计2</param>
    ''' <param name="_sPara3">把计3</param>
    ''' <returns>Integer 肚掸计</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Rechar] 20080728 Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Function loadDemo(ByVal _sPara1 As String, ByVal _iPara2 As Integer, ByVal _sPara3 As String) As Integer
        Try
            If ((Not IsNothing(_sPara1)) AndAlso (_sPara1.Trim.Length > 0)) Then
                Dim para(2) As System.Data.IDbDataParameter
                para(0) = ProviderFactory.CreateDataParameter("FILED_1", _sPara1)
                para(1) = ProviderFactory.CreateDataParameter("FILED_2", _iPara2)
                para(2) = ProviderFactory.CreateDataParameter("FILED_3", _sPara3)
                Return MyBase.loadBySQL(para)
            End If

            Return 0
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try

    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' loadDemo 絛ㄒ
    ''' </summary>
    ''' <returns>Integer 肚掸计</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Rechar] 20080728 Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Function loadValidDAO() As Integer
        Dim sbSql As New System.Text.StringBuilder
        Try
            sbSql.Append(" select * from SUBSYS_CODE where disabled=0 order by subsysid ")
            Return MyBase.loadBySQL(sbSql.ToString())

        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
        Return 0
    End Function
End Class
