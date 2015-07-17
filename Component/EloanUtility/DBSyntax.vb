''' <summary>
''' 依DB Server的SQL語法不同，回傳不同的String
''' </summary>
''' <remarks>
''' 目前僅區分Oracel與MS_SQL
''' 方法名稱以Oracle為主
''' </remarks>

Public Class DBSyntax
    Private Shared eDBFlag As String = FileUtility.getAppSettings("Provider").ToUpper

    Private Structure DB
        Public Const ORACLE As String = "ODP"
        Public Const MS_SQL As String = "SQLCLIENT"
    End Structure


    Public Shared Function sysdate() As String
        Dim sResult As String = String.Empty

        Select Case eDBFlag
            Case DB.MS_SQL
                sResult = " getdate() "
            Case DB.ORACLE
                sResult = " sysdate "
        End Select

        Return sResult
    End Function

    Public Shared Function nvl(ByVal sVal1 As String, ByVal sVal2 As String) As String
        Dim sResult As String = String.Empty

        Select Case eDBFlag
            Case DB.ORACLE
                sResult = " nvl(" & sVal1 & "," & sVal2 & ") "
            Case DB.MS_SQL
                sResult = " isnull(" & sVal1 & "," & sVal2 & ") "
        End Select

        Return sResult
    End Function

    Public Shared Function len(ByVal sVal1 As String) As String
        Dim sResult As String = String.Empty

        Select Case eDBFlag
            Case DB.ORACLE
                sResult = " len(" & sVal1 & ") "
            Case DB.MS_SQL
                sResult = " length(" & sVal1 & ") "
        End Select

        Return sResult
    End Function

    Public Shared Function trim(ByVal sVal1 As String) As String
        Dim sResult As String = String.Empty


        Select Case eDBFlag
            Case DB.ORACLE
                sResult = " trim(" & sVal1 & ") "
            Case DB.MS_SQL
                sResult = " ltrim(rtrim(" & sVal1 & ")) "
        End Select

        Return sResult
    End Function
End Class
