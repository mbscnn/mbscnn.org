
Imports com.Azion.NET.VB

Public Class SUBSYS_CODE
    Inherits BosBase

    Sub New()
        MyBase.New()
        MyBase.setPrimaryKeys()
    End Sub

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.new("SUBSYS_CODE", dbManager)
    End Sub



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Demo
    ''' </summary>
    ''' <param name="_sPara1">把计1</param>
    ''' <param name="_iPara2">把计2</param>
    ''' <param name="_sPara3">把计3</param>
    ''' <returns>String 肚</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Rechar] 20080728 Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function getDemo(ByVal _sPara1 As String, ByVal _iPara2 As Integer, ByVal _sPara3 As String) As String
        Return "Demo"
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 眔╰参 戈
    ''' </summary>
    ''' <param name="_sSUBSYSID">╰参</param>
    ''' <returns>boolean 肚</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Rechar] 20080728 Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function loadBySubSysId(ByVal _sSUBSYSID As String) As Boolean
        Try
            Dim paras(0) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("SUBSYSID", _sSUBSYSID)
            Return MyBase.loadData(paras)

        Catch ex As Exception
            Throw ex
        End Try
        Return False
    End Function

End Class
