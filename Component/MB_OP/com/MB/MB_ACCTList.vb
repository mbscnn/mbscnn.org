Imports com.Azion.NET.VB

Public Class MB_ACCTList
    Inherits BosList

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.New("MB_ACCT", dbManager)
    End Sub

    Overrides Function newBos() As BosBase
        Dim bos As MB_ACCT = New MB_ACCT(MyBase.getDatabaseManager)
        Return bos
    End Function

#Region "Ted Function"
    ''' <summary>
    ''' 根據"手機"取得資料
    ''' </summary>
    ''' <param name="sMB_MOBIL">手機</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Function Load_MB_MOBIL(Byval sMB_MOBIL As string) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT * FROM MB_ACCT " & _
                     "  WHERE MB_MOBIL = " & ProviderFactory.PositionPara & "MB_MOBIL "

            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_MOBIL",sMB_MOBIL)

            Return Me.loadBySQLOnlyDs(sqlStr,para)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' 根據"電話"取得資料
    ''' </summary>
    ''' <param name="sMB_TEL">電話</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Function Load_MB_TEL(Byval sMB_TEL As string) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT * FROM MB_ACCT " & _
                     "  WHERE MB_TEL = " & ProviderFactory.PositionPara & "MB_TEL "

            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_TEL",sMB_TEL)

            Return Me.loadBySQLOnlyDs(sqlStr,para)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Function Load_BATCH() As Integer
        Try
            Dim sqlStr As String = String.Empty

            'sqlStr = "SELECT * " & _
            '         "  FROM MB_ACCT A " & _
            '         " WHERE     LENGTH(TRIM(IFNULL(A.MB_NAME, ' '))) > 0 " & _
            '         "       AND (    NOT EXISTS " & _
            '         "                   (SELECT * " & _
            '         "                      FROM MB_MEMBER " & _
            '         "                     WHERE MB_EMAIL = A.MB_ACCT) " & _
            '         "            AND (   (    A.MB_MOBIL IS NOT NULL " & _
            '         "                     AND NOT EXISTS " & _
            '         "                            (SELECT * " & _
            '         "                               FROM MB_MEMBER " & _
            '         "                              WHERE MB_MOBIL = A.MB_MOBIL)) " & _
            '         "                 OR A.MB_MOBIL IS NULL) " & _
            '         "            AND (   (    A.MB_TEL IS NOT NULL " & _
            '         "                     AND NOT EXISTS " & _
            '         "                            (SELECT * " & _
            '         "                               FROM MB_MEMBER " & _
            '         "                              WHERE MB_TEL = A.MB_TEL)) " & _
            '         "                 OR A.MB_TEL IS NULL)) "

            sqlStr = "SELECT A.* " &
                     "  FROM MB_ACCT A " &
                     " WHERE NOT EXISTS " &
                     "          (SELECT * " &
                     "             FROM MB_MEMBER " &
                     "            WHERE MB_NAME = A.MB_NAME) "

            Return Me.loadBySQLOnlyDs(sqlStr)
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
