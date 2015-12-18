Imports com.Azion.NET.VB

Public Class MB_MEMCLASSList
    Inherits BosList

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.New("MB_MEMCLASS", dbManager)
    End Sub

    Overrides Function newBos() As BosBase
        Dim bos As MB_MEMCLASS = New MB_MEMCLASS(MyBase.getDatabaseManager)
        Return bos
    End Function

#Region "Grace Function"
    ''' <summary>
    ''' 根據[課程編號]取得資料
    ''' </summary>
    ''' <param name="iMB_SEQ">MB_SEQ會員編號</param>
    ''' <returns>Integer 取得筆數</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Function loadByMB_SEQ(ByVal iMB_SEQ As Decimal, ByVal iMB_BATCH As Decimal) As Integer
        Try
            Dim sqlStr As String = String.Empty
            sqlStr = "SELECT A.*, B.MB_BATCH " &
                     "  FROM MB_MEMCLASS A, MB_MEMBATCH B " &
                     " WHERE     A.MB_SEQ = " & ProviderFactory.PositionPara & "MB_SEQ " &
                     "       AND A.MB_MEMSEQ = B.MB_MEMSEQ " &
                     "       AND A.MB_SEQ = B.MB_SEQ " &
                     "       AND B.MB_BATCH = " & ProviderFactory.PositionPara & "MB_BATCH "

            Dim paras(1) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("MB_SEQ", iMB_SEQ)
            paras(1) = ProviderFactory.CreateDataParameter("MB_BATCH", iMB_BATCH)

            Return Me.loadBySQL(sqlStr, paras)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function
#End Region

#Region "Ted Function"
    ''' <summary>
    ''' 根據[課程編號][梯次]取得課程報名查詢資料
    ''' </summary>
    ''' <param name="iMB_SEQ">MB_SEQ課程編號</param>
    ''' <param name="iMB_BATCH">梯次</param>
    ''' <returns>Integer 取得筆數</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Function getMail_TO(ByVal iMB_SEQ As Decimal, ByVal iMB_BATCH As Decimal) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT B.MB_EMAIL, B.MB_NAME, A.*, C.MB_BATCH " &
                     "  FROM MB_MEMCLASS A, MB_MEMBER B, MB_MEMBATCH C " &
                     " WHERE     A.MB_SEQ = " & ProviderFactory.PositionPara & "MB_SEQ " &
                     "       AND C.MB_BATCH = " & ProviderFactory.PositionPara & "MB_BATCH " &
                     "       AND IFNULL(A.MB_FWMK, ' ') NOT IN ('3', '4', '5') " &
                     "       AND IFNULL(MB_RESP,' ')<>'N' " &
                     "       AND A.MB_MEMSEQ = B.MB_MEMSEQ " &
                     "       AND A.MB_MEMSEQ = C.MB_MEMSEQ " &
                     "       AND A.MB_SEQ = C.MB_SEQ " &
                     "       AND B.MB_MEMSEQ = C.MB_MEMSEQ "

            Dim paras(1) As IDbDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("MB_SEQ", iMB_SEQ)

            paras(1) = ProviderFactory.CreateDataParameter("MB_BATCH", iMB_BATCH)

            Return Me.loadBySQLOnlyDs(sqlStr, paras)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' 根據[課程編號]取得課程報名查詢資料
    ''' </summary>
    ''' <param name="iMB_SEQ">MB_SEQ會員編號</param>
    ''' <returns>Integer 取得筆數</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Function LoadQData(ByVal iStart As Decimal, ByVal iPageSize As Decimal, ByVal iMB_SEQ As Decimal) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT C.*, (SELECT MAX(MB_CLASS_NAME) FROM MB_CLASS WHERE MB_SEQ=C.MB_SEQ) AS MB_CLASS_NAME,  " &
                     "  (SELECT MAX(MB_APV) FROM MB_CLASS WHERE MB_SEQ=C.MB_SEQ) AS MB_APV " &
                     "  FROM (SELECT A.MB_SEQ, " &
                     "               A.MB_CHKFLAG, " &
                     "               A.MB_CREDATETIME, " &
                     "               A.MB_FWMK, " &
                     "               A.MB_SORTNO, " &
                     "               A.MB_CDATE," &
                     "               A.MB_RESP," &
                     "               A.MB_APVDATETIME," &
                     "               NULL AS MB_BATCH, " &
                     "               B.* " &
                     "          FROM MB_MEMCLASS A, MB_MEMBER B" &
                     "         WHERE A.MB_SEQ = " & ProviderFactory.PositionPara & "MB_SEQ AND A.MB_MEMSEQ = B.MB_MEMSEQ) C " &
                     " ORDER BY MB_CREDATETIME " &
                     "   LIMIT " & iStart & ", " & iPageSize

            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_SEQ", iMB_SEQ)

            Return Me.loadBySQLOnlyDs(sqlStr, para)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function LoadQEXCELData(ByVal iMB_SEQ As Decimal) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT C.*, (SELECT MAX(MB_CLASS_NAME) FROM MB_CLASS WHERE MB_SEQ = C.MB_SEQ) AS MB_CLASS_NAME " &
                     "  FROM (SELECT A.MB_SEQ, " &
                     "               A.MB_CHKFLAG, " &
                     "               A.MB_CREDATETIME, " &
                     "               A.MB_FWMK, " &
                     "               A.MB_SORTNO, " &
                     "               NULL AS MB_BATCH," &
                     "               A.MB_RESP," &
                     "               A.MB_CDATE," &
                     "               A.MB_OBJECT," &
                     "               B.* " &
                     "          FROM MB_MEMCLASS A, MB_MEMBER B " &
                     "         WHERE A.MB_SEQ = " & ProviderFactory.PositionPara & "MB_SEQ AND A.MB_MEMSEQ = B.MB_MEMSEQ) C " &
                     " ORDER BY MB_MEMSEQ "

            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_SEQ", iMB_SEQ)

            Return Me.loadBySQLOnlyDs(sqlStr, para)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function COUNT_MB_MEMCLASS(ByVal iMB_SEQ As Decimal) As Decimal
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT IFNULL(COUNT(*), 0) " & _
                     "  FROM MB_MEMCLASS " & _
                     " WHERE MB_SEQ = " & ProviderFactory.PositionPara & "MB_SEQ "

            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_SEQ", iMB_SEQ)

            If Me.getDatabaseManager.isTransaction Then
                Return DBObject.ExecuteScalar(Me.getDatabaseManager.getTransaction, CommandType.Text, sqlStr, para)
            Else
                Return DBObject.ExecuteScalar(Me.getDatabaseManager.getConnection, CommandType.Text, sqlStr, para)
            End If
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
