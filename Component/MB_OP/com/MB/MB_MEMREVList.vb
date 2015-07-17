Imports com.Azion.NET.VB

Public Class MB_MEMREVList
    Inherits BosList

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.New("MB_MEMREV", dbManager)
    End Sub

    Overrides Function newBos() As BosBase
        Dim bos As MB_MEMREV = New MB_MEMREV(MyBase.getDatabaseManager)
        Return bos
    End Function

#Region "Ted Function"
    ''' <summary>
    ''' 根據[會員編號]取得資料
    ''' </summary>
    ''' <param name="iMB_MEMSEQ">MB_MEMSEQ會員編號</param>
    ''' <returns>Integer 取得筆數</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Function loadByMB_MEMSEQ(ByVal iMB_MEMSEQ As Decimal) As Integer
        Try
            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_MEMSEQ", iMB_MEMSEQ)

            Return Me.loadBySQL(para)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' ----------------------------------------------------------------------------
    ''' <summary>
    ''' 根據[所屬區代碼]取得MB_VOUCHER_Y傳票號碼-西元年為NULL或0的資料
    ''' </summary>
    ''' <param name="sMB_AREA">MB_AREA所屬區代碼</param>
    ''' <returns>integer取得筆數</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Function loadByMB_AREA_VOUCHER_Y_0(ByVal sMB_AREA As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT A.*                               " & _
                     "       ,B.MB_NAME,B.MB_AREA,B.MB_LEADER  " & _
                     "  FROM MB_MEMREV A, MB_MEMBER B          " & _
                     " WHERE     IFNULL(A.MB_VOUCHER_Y, 0) = 0 " & _
                     "       AND A.MB_MEMSEQ = B.MB_MEMSEQ	   " & _
                     "       AND B.MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA "

            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_AREA", sMB_AREA)

            Return Me.loadBySQLOnlyDs(sqlStr, para)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' 根據[所屬區代碼]取得MB_VOUCHER_Y傳票號碼-西元年為NULL或0的資料
    ''' </summary>
    ''' <param name="sMB_AREA">MB_AREA所屬區代碼</param>
    ''' <returns>integer取得筆數</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Function loadByMB_AREA_1(ByVal sMB_AREA As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT A.*                               " & _
                     "       ,B.MB_NAME,B.MB_AREA,B.MB_LEADER  " & _
                     "  FROM MB_MEMREV A, MB_MEMBER B          " & _
                     " WHERE     A.VRYUID IS NULL           " & _
                     "       AND A.MB_MEMSEQ = B.MB_MEMSEQ	   " & _
                     "       AND B.MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA "

            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_AREA", sMB_AREA)

            Return Me.loadBySQLOnlyDs(sqlStr, para)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' 根據[所屬區代碼]取得MB_VOUCHER_Y傳票號碼-西元年為NULL或0的資料
    ''' </summary>
    ''' <param name="sMB_AREA">MB_AREA所屬區代碼</param>
    ''' <returns>integer取得筆數</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Function loadByMB_AREA_2(ByVal sMB_AREA As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT A.*                               " & _
                     "       ,B.MB_NAME,B.MB_AREA,B.MB_LEADER  " & _
                     "  FROM MB_MEMREV A, MB_MEMBER B          " & _
                     " WHERE  A.VRYUID IS NOT NULL          " & _
                     "       AND IFNULL(A.MB_VOUCHER_Y, 0) = 0 " & _
                     "       AND A.MB_MEMSEQ = B.MB_MEMSEQ	   " & _
                     "       AND B.MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA "

            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_AREA", sMB_AREA)

            Return Me.loadBySQLOnlyDs(sqlStr, para)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' 根據[所屬區代碼]取得MB_VOUCHER_Y傳票號碼-西元年為NULL或0的資料
    ''' </summary>
    ''' <param name="sMB_AREA">MB_AREA所屬區代碼</param>
    ''' <returns>integer取得筆數</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Function loadByMB_AREA_3(ByVal sMB_AREA As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT A.*                               " & _
                     "       ,B.MB_NAME,B.MB_AREA,B.MB_LEADER  " & _
                     "  FROM MB_MEMREV A, MB_MEMBER B          " & _
                     " WHERE  IFNULL(A.MB_VOUCHER_Y, 0) > 0 " & _
                     "       AND A.MB_MEMSEQ = B.MB_MEMSEQ	   " & _
                     "       AND B.MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA " & _
                     "       AND A.DELFLAG IS NULL " & _
                     "       AND IFNULL(A.NOTE_CASH,' ')<>'N'"

            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_AREA", sMB_AREA)

            Return Me.loadBySQLOnlyDs(sqlStr, para)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' ----------------------------------------------------------------------------
    ''' <summary>
    ''' 根據[所屬區代碼]取得MB_VOUCHER_Y傳票號碼-西元年不為0的資料
    ''' </summary>
    ''' <param name="sMB_AREA">MB_AREA所屬區代碼</param>
    ''' <returns>integer取得筆數</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Function loadByMB_AREA_VOUCHER_Y(ByVal sMB_AREA As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT a.*,                                " & _
                     "       B.MB_NAME,                          " & _
                     "       B.MB_AREA,                          " & _
                     "       B.MB_LEADER                         " & _
                     "  FROM MB_MEMREV a, MB_MEMBER b            " & _
                     " WHERE     IFNULL(a.MB_VOUCHER_Y, 0) > 0   " & _
                     "       AND a.MB_MEMSEQ = b.MB_MEMSEQ       " & _
                     "       AND b.MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA " & _
                     "       AND a.DELFLAG IS NULL               " & _
                     "       AND IFNULL(a.NOTE_CASH, ' ') <> 'N' "

            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_AREA", sMB_AREA)

            Return Me.loadBySQLOnlyDs(sqlStr, para)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' ----------------------------------------------------------------------------
    ''' <summary>
    ''' 根據[繳款日期]取得MB_VOUCHER_Y傳票號碼-西元年為NULL或0的資料
    ''' </summary>
    ''' <param name="sMB_TX_DATE">MB_TX_DATE繳款日期 </param>
    ''' <returns>integer取得筆數</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Function loadByMB_TX_DATE_VOUCHER_Y_0(ByVal sMB_TX_DATE As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT A.*,                              " & _
                     "       B.MB_NAME,                        " & _
                     "       B.MB_AREA,                        " & _
                     "       B.MB_LEADER                       " & _
                     "  FROM MB_MEMREV A, MB_MEMBER B          " & _
                     " WHERE     IFNULL(A.MB_VOUCHER_Y, 0) = 0 " & _
                     "       AND A.MB_MEMSEQ = B.MB_MEMSEQ     " & _
                     "       AND SUBSTR(A.MB_TX_DATE, 1, 6) =  " & ProviderFactory.PositionPara & "MB_TX_DATE "

            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_TX_DATE", sMB_TX_DATE)

            Return Me.loadBySQLOnlyDs(sqlStr, para)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' 根據[繳款日期]取得MB_VOUCHER_Y傳票號碼-西元年為NULL或0的資料
    ''' </summary>
    ''' <param name="sMB_TX_DATE">MB_TX_DATE繳款日期 </param>
    ''' <returns>integer取得筆數</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Function loadByMB_TX_DATE_1(ByVal sMB_TX_DATE As String, ByVal sMB_AREA As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT A.*,                              " & _
                     "       B.MB_NAME,                        " & _
                     "       B.MB_AREA,                        " & _
                     "       B.MB_LEADER                       " & _
                     "  FROM MB_MEMREV A, MB_MEMBER B          " & _
                     " WHERE VRYUID IS NULL                    " & _
                     "       AND A.MB_MEMSEQ = B.MB_MEMSEQ     " & _
                     "       AND SUBSTR(A.MB_TX_DATE, 1, 6) =  " & ProviderFactory.PositionPara & "MB_TX_DATE "

            If Utility.isValidateData(sMB_AREA) Then
                sqlStr &= "       AND B.MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA "

                Dim paras(1) As IDbDataParameter

                paras(0) = ProviderFactory.CreateDataParameter("MB_TX_DATE", sMB_TX_DATE)

                paras(1) = ProviderFactory.CreateDataParameter("MB_AREA", sMB_AREA)

                Return Me.loadBySQLOnlyDs(sqlStr, paras)
            Else
                Dim paras As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_TX_DATE", sMB_TX_DATE)

                Return Me.loadBySQLOnlyDs(sqlStr, paras)
            End If
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' 根據[繳款日期]取得MB_VOUCHER_Y傳票號碼-西元年為NULL或0的資料
    ''' </summary>
    ''' <param name="sMB_TX_DATE">MB_TX_DATE繳款日期 </param>
    ''' <returns>integer取得筆數</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Function loadByMB_TX_DATE_2(ByVal sMB_TX_DATE As String, ByVal sMB_AREA As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT A.*,                              " & _
                     "       B.MB_NAME,                        " & _
                     "       B.MB_AREA,                        " & _
                     "       B.MB_LEADER                       " & _
                     "  FROM MB_MEMREV A, MB_MEMBER B          " & _
                     " WHERE VRYUID IS NOT NULL             " & _
                     "       AND IFNULL(A.MB_VOUCHER_Y, 0) = 0 " & _
                     "       AND A.MB_MEMSEQ = B.MB_MEMSEQ     " & _
                     "       AND SUBSTR(A.MB_TX_DATE, 1, 6) =  " & ProviderFactory.PositionPara & "MB_TX_DATE "

            If Utility.isValidateData(sMB_AREA) Then
                sqlStr = "       AND B.MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA "

                Dim paras(1) As IDbDataParameter
                paras(0) = ProviderFactory.CreateDataParameter("MB_TX_DATE", sMB_TX_DATE)
                paras(1) = ProviderFactory.CreateDataParameter("MB_AREA", sMB_AREA)

                Return Me.loadBySQLOnlyDs(sqlStr, paras)
            Else
                Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_TX_DATE", sMB_TX_DATE)

                Return Me.loadBySQLOnlyDs(sqlStr, para)
            End If
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' 根據[繳款日期]取得MB_VOUCHER_Y傳票號碼-西元年為NULL或0的資料
    ''' </summary>
    ''' <param name="iB_MB_TX_DATE">MB_TX_DATE繳款日期 </param>
    ''' <param name="iE_MB_TX_DATE">MB_TX_DATE繳款日期 </param>
    ''' <returns>integer取得筆數</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Function loadByMB_TX_DATE_3(ByVal iB_MB_TX_DATE As Decimal, ByVal iE_MB_TX_DATE As Decimal, ByVal sMB_AREA As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT A.*,                              " & _
                     "       B.MB_NAME,                        " & _
                     "       B.MB_AREA,                        " & _
                     "       B.MB_LEADER                       " & _
                     "  FROM MB_MEMREV A, MB_MEMBER B          " & _
                     " WHERE IFNULL(A.MB_VOUCHER_Y, 0) > 0 " & _
                     "       AND A.MB_MEMSEQ = B.MB_MEMSEQ     " & _
                     "       AND CAST(SUBSTR(A.MB_TX_DATE, 1, 6) AS decimal) >=  " & ProviderFactory.PositionPara & "B_MB_TX_DATE " & _
                     "       AND CAST(SUBSTR(A.MB_TX_DATE, 1, 6) AS decimal) <=  " & ProviderFactory.PositionPara & "E_MB_TX_DATE " & _
                     "       AND A.DELFLAG IS NULL " & _
                     "       AND IFNULL(A.NOTE_CASH,' ')<>'N'"

            If Utility.isValidateData(sMB_AREA) Then
                sqlStr &= "       AND B.MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA "

                Dim paras(2) As IDbDataParameter
                paras(0) = ProviderFactory.CreateDataParameter("B_MB_TX_DATE", iB_MB_TX_DATE)
                paras(1) = ProviderFactory.CreateDataParameter("E_MB_TX_DATE", iE_MB_TX_DATE)
                paras(2) = ProviderFactory.CreateDataParameter("MB_AREA", sMB_AREA)

                Return Me.loadBySQLOnlyDs(sqlStr, paras)
            Else
                Dim paras(1) As IDbDataParameter
                paras(0) = ProviderFactory.CreateDataParameter("B_MB_TX_DATE", iB_MB_TX_DATE)
                paras(1) = ProviderFactory.CreateDataParameter("E_MB_TX_DATE", iE_MB_TX_DATE)

                Return Me.loadBySQLOnlyDs(sqlStr, paras)
            End If
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' ----------------------------------------------------------------------------
    ''' <summary>
    ''' 根據[繳款日期起日][繳款日期迄日]取得MB_VOUCHER_Y傳票號碼-西元年不為0的資料
    ''' </summary>
    ''' <param name="iMB_TX_DATE_BEG">MB_TX_DATE繳款日期起日</param>
    ''' <param name="iMB_TX_DATE_END">MB_TX_DATE繳款日期迄日</param>
    ''' <returns>integer取得筆數</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Function loadByMB_TX_DATE_VOUCHER_Y(ByVal iMB_TX_DATE_BEG As Decimal, ByVal iMB_TX_DATE_END As Decimal) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT a.*,                                                            " & _
                     "       B.MB_NAME,                                                      " & _
                     "       B.MB_AREA,                                                      " & _
                     "       B.MB_LEADER                                                     " & _
                     "  FROM MB_MEMREV a, MB_MEMBER b                                        " & _
                     " WHERE     IFNULL(a.MB_VOUCHER_Y, 0) > 0                               " & _
                     "       AND a.MB_MEMSEQ = b.MB_MEMSEQ                                   " & _
                     "       AND (    CAST(SUBSTR(A.MB_TX_DATE, 1, 6) AS decimal) >= " & ProviderFactory.PositionPara & "MB_TX_DATE_BEG  " & _
                     "            AND CAST(SUBSTR(A.MB_TX_DATE, 1, 6) AS decimal) <= " & ProviderFactory.PositionPara & "MB_TX_DATE_END ) " & _
                     "       AND a.DELFLAG IS NULL											 " & _
                     "       AND IFNULL(a.NOTE_CASH, ' ') <> 'N'                             "

            Dim paras(1) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("MB_TX_DATE_BEG", iMB_TX_DATE_BEG)
            paras(1) = ProviderFactory.CreateDataParameter("MB_TX_DATE_END", iMB_TX_DATE_END)

            Return Me.loadBySQLOnlyDs(sqlStr, paras)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' ----------------------------------------------------------------------------
    ''' <summary>
    ''' 根據[姓名]取得MB_VOUCHER_Y傳票號碼-西元年不為0的資料
    ''' </summary>
    ''' <param name="sMB_NAME">MB_NAME姓名</param>
    ''' <returns>integer取得筆數</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Function loadByMB_NAME_VOUCHER_Y(ByVal sMB_NAME As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT a.*,                                " & _
                     "       B.MB_NAME,                          " & _
                     "       B.MB_AREA,                          " & _
                     "       B.MB_LEADER                         " & _
                     "  FROM MB_MEMREV a, MB_MEMBER b            " & _
                     " WHERE     IFNULL(a.MB_VOUCHER_Y, 0) > 0   " & _
                     "       AND a.MB_MEMSEQ = b.MB_MEMSEQ       " & _
                     "       AND b.MB_NAME LIKE " & ProviderFactory.PositionPara & "MB_NAME " & _
                     "       AND a.DELFLAG IS NULL               " & _
                     "       AND IFNULL(a.NOTE_CASH, ' ') <> 'N' "

            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_NAME", sMB_NAME & "%")

            Return Me.loadBySQLOnlyDs(sqlStr, para)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' ----------------------------------------------------------------------------
    ''' <summary>
    ''' 根據[所屬區代碼]取得MB_MEMREV的列印收據資料
    ''' </summary>
    ''' <param name="sMB_AREA">MB_AREA所屬區代碼</param>
    ''' <returns>integer取得筆數</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Function LoadByMB_AREAForPrint(ByVal sMB_AREA As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT a.*, b.*" & _
                     "  FROM MB_MEMREV a, MB_MEMBER b" & _
                     " WHERE     (a.MB_VOUCHER_Y > 0)" & _
                     "       AND a.MB_MEMSEQ = b.MB_MEMSEQ" & _
                     "       AND b.MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA_1 " & _
                     "       AND a.DELFLAG IS NULL" & _
                     "       AND a.MB_PRINT <> 'Y'" & _
                     "       AND a.MB_PAY_TYPE = 'C'" & _
                     " UNION " & _
                     "SELECT a.*, b.*" & _
                     "  FROM MB_MEMREV a, MB_MEMBER b" & _
                     " WHERE     (a.MB_VOUCHER_Y > 0)" & _
                     "       AND a.MB_MEMSEQ = b.MB_MEMSEQ" & _
                     "       AND b.MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA_2 " & _
                     "       AND a.DELFLAG IS NULL" & _
                     "       AND a.MB_PRINT <> 'Y'" & _
                     "       AND a.MB_PAY_TYPE = 'N'" & _
                     "       AND a.NOTE_CASH = 'Y'"

            Dim paras(1) As IDbDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("MB_AREA_1", sMB_AREA)

            paras(1) = ProviderFactory.CreateDataParameter("MB_AREA_2", sMB_AREA)

            Return Me.loadBySQL(sqlStr, paras)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' ----------------------------------------------------------------------------
    ''' <summary>
    ''' 根據[繳款日期起日][繳款日期迄日]取得MB_MEMREV的列印收據資料
    ''' </summary>
    ''' <param name="iMB_TX_DATE_BEG">MB_TX_DATE繳款日期起日</param>
    ''' <param name="iMB_TX_DATE_END">MB_TX_DATE繳款日期迄日</param>
    ''' <returns>integer取得筆數</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Function LoadByMB_TX_DATEForPrint(ByVal iMB_TX_DATE_BEG As Decimal, ByVal iMB_TX_DATE_END As Decimal, ByVal sMB_AREA As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            If Utility.isValidateData(sMB_AREA) Then
                sqlStr = "SELECT a.*, b.*" & _
                         "  FROM MB_MEMREV a, MB_MEMBER b" & _
                         " WHERE IFNULL(a.MB_VOUCHER_Y, 0) > 0" & _
                         "   AND a.MB_MEMSEQ = b.MB_MEMSEQ" & _
                         "   AND (CAST(SUBSTR(A.MB_TX_DATE, 1, 6) AS decimal) >= " & ProviderFactory.PositionPara & "MB_TX_DATE_BEG_1 AND" & _
                         "        CAST(SUBSTR(A.MB_TX_DATE, 1, 6) AS decimal) <= " & ProviderFactory.PositionPara & "MB_TX_DATE_END_1 )" & _
                         "   AND a.DELFLAG IS NULL" & _
                         "   AND a.MB_PRINT <> 'Y' " & _
                         "   AND a.MB_PAY_TYPE = 'C'" & _
                         "   AND B.MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA_1 " & _
                         " union " & _
                         "SELECT a.*, b.*" & _
                         "  FROM MB_MEMREV a, MB_MEMBER b" & _
                         " WHERE IFNULL(a.MB_VOUCHER_Y, 0) > 0" & _
                         "   AND a.MB_MEMSEQ = b.MB_MEMSEQ" & _
                         "   AND (CAST(SUBSTR(A.MB_TX_DATE, 1, 6) AS decimal) >= " & ProviderFactory.PositionPara & "MB_TX_DATE_BEG_2 And " & _
                         "        CAST(SUBSTR(A.MB_TX_DATE, 1, 6) AS decimal) <= " & ProviderFactory.PositionPara & "MB_TX_DATE_END_2 )" & _
                         "   AND a.DELFLAG IS NULL" & _
                         "   AND a.MB_PRINT <> 'Y'" & _
                         "   AND a.MB_PAY_TYPE = 'N'" & _
                         "   AND a.NOTE_CASH = 'Y'" & _
                         "   AND B.MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA_2 "

                Dim paras(5) As IDbDataParameter

                paras(0) = ProviderFactory.CreateDataParameter("MB_TX_DATE_BEG_1", iMB_TX_DATE_BEG)

                paras(1) = ProviderFactory.CreateDataParameter("MB_TX_DATE_END_1", iMB_TX_DATE_END)

                paras(2) = ProviderFactory.CreateDataParameter("MB_AREA_1", sMB_AREA)

                paras(3) = ProviderFactory.CreateDataParameter("MB_TX_DATE_BEG_2", iMB_TX_DATE_BEG)

                paras(4) = ProviderFactory.CreateDataParameter("MB_TX_DATE_END_2", iMB_TX_DATE_END)

                paras(5) = ProviderFactory.CreateDataParameter("MB_AREA_2", sMB_AREA)

                Return Me.loadBySQL(sqlStr, paras)
            Else
                sqlStr = "SELECT a.*, b.*" & _
                         "  FROM MB_MEMREV a, MB_MEMBER b" & _
                         " WHERE IFNULL(a.MB_VOUCHER_Y, 0) > 0" & _
                         "   AND a.MB_MEMSEQ = b.MB_MEMSEQ" & _
                         "   AND (CAST(SUBSTR(A.MB_TX_DATE, 1, 6) AS decimal) >= " & ProviderFactory.PositionPara & "MB_TX_DATE_BEG_1 AND" & _
                         "        CAST(SUBSTR(A.MB_TX_DATE, 1, 6) AS decimal) <= " & ProviderFactory.PositionPara & "MB_TX_DATE_END_1 )" & _
                         "   AND a.DELFLAG IS NULL" & _
                         "   AND a.MB_PRINT <> 'Y' " & _
                         "   AND a.MB_PAY_TYPE = 'C'" & _
                         " union " & _
                         "SELECT a.*, b.*" & _
                         "  FROM MB_MEMREV a, MB_MEMBER b" & _
                         " WHERE IFNULL(a.MB_VOUCHER_Y, 0) > 0" & _
                         "   AND a.MB_MEMSEQ = b.MB_MEMSEQ" & _
                         "   AND (CAST(SUBSTR(A.MB_TX_DATE, 1, 6) AS decimal) >= " & ProviderFactory.PositionPara & "MB_TX_DATE_BEG_2 And " & _
                         "        CAST(SUBSTR(A.MB_TX_DATE, 1, 6) AS decimal) <= " & ProviderFactory.PositionPara & "MB_TX_DATE_END_2 )" & _
                         "   AND a.DELFLAG IS NULL" & _
                         "   AND a.MB_PRINT <> 'Y'" & _
                         "   AND a.MB_PAY_TYPE = 'N'" & _
                         "   AND a.NOTE_CASH = 'Y'"

                Dim paras(3) As IDbDataParameter

                paras(0) = ProviderFactory.CreateDataParameter("MB_TX_DATE_BEG_1", iMB_TX_DATE_BEG)

                paras(1) = ProviderFactory.CreateDataParameter("MB_TX_DATE_END_1", iMB_TX_DATE_END)

                paras(2) = ProviderFactory.CreateDataParameter("MB_TX_DATE_BEG_2", iMB_TX_DATE_BEG)

                paras(3) = ProviderFactory.CreateDataParameter("MB_TX_DATE_END_2", iMB_TX_DATE_END)

                Return Me.loadBySQL(sqlStr, paras)
            End If
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' ----------------------------------------------------------------------------
    ''' <summary>
    ''' 根據[姓名]取得MB_MEMREV的列印收據資料
    ''' </summary>
    ''' <param name="sMB_NAME">MB_NAME姓名</param>
    ''' <returns>integer取得筆數</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Function LoadByMB_NAMEForPrint(ByVal sMB_NAME As String, ByVal sMB_AREA As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            If Utility.isValidateData(sMB_AREA) Then
                sqlStr = "SELECT a.*, b.*" & _
                         "  FROM MB_MEMREV a, MB_MEMBER b" & _
                         " WHERE     (a.MB_VOUCHER_Y > 0)" & _
                         "       AND a.MB_MEMSEQ = b.MB_MEMSEQ" & _
                         "       AND b.MB_NAME LIKE " & ProviderFactory.PositionPara & "MB_NAME_1 " & _
                         "       AND a.DELFLAG IS NULL" & _
                         "       AND a.MB_PRINT <> 'Y'" & _
                         "       AND a.MB_PAY_TYPE = 'C'" & _
                         "       AND B.MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA_1 " & _
                         " UNION " & _
                         "SELECT a.*, b.*" & _
                         "  FROM MB_MEMREV a, MB_MEMBER b" & _
                         " WHERE     (a.MB_VOUCHER_Y > 0)" & _
                         "       AND a.MB_MEMSEQ = b.MB_MEMSEQ" & _
                         "       AND b.MB_NAME = " & ProviderFactory.PositionPara & "MB_NAME_2 " & _
                         "       AND a.DELFLAG IS NULL" & _
                         "       AND a.MB_PRINT <> 'Y'" & _
                         "       AND a.MB_PAY_TYPE = 'N'" & _
                         "       AND a.NOTE_CASH = 'Y'" & _
                         "       AND B.MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA_2 "

                Dim paras(1) As IDbDataParameter

                paras(0) = ProviderFactory.CreateDataParameter("MB_NAME_1", sMB_NAME & "%")

                paras(1) = ProviderFactory.CreateDataParameter("MB_AREA_1", sMB_AREA)

                paras(2) = ProviderFactory.CreateDataParameter("MB_NAME_2", sMB_NAME & "%")

                paras(3) = ProviderFactory.CreateDataParameter("MB_AREA_2", sMB_AREA)

                Return Me.loadBySQL(sqlStr, paras)
            Else
                sqlStr = "SELECT a.*, b.*" & _
                         "  FROM MB_MEMREV a, MB_MEMBER b" & _
                         " WHERE     (a.MB_VOUCHER_Y > 0)" & _
                         "       AND a.MB_MEMSEQ = b.MB_MEMSEQ" & _
                         "       AND b.MB_NAME LIKE " & ProviderFactory.PositionPara & "MB_NAME_1 " & _
                         "       AND a.DELFLAG IS NULL" & _
                         "       AND a.MB_PRINT <> 'Y'" & _
                         "       AND a.MB_PAY_TYPE = 'C'" & _
                         " UNION " & _
                         "SELECT a.*, b.*" & _
                         "  FROM MB_MEMREV a, MB_MEMBER b" & _
                         " WHERE     (a.MB_VOUCHER_Y > 0)" & _
                         "       AND a.MB_MEMSEQ = b.MB_MEMSEQ" & _
                         "       AND b.MB_NAME = " & ProviderFactory.PositionPara & "MB_NAME_2 " & _
                         "       AND a.DELFLAG IS NULL" & _
                         "       AND a.MB_PRINT <> 'Y'" & _
                         "       AND a.MB_PAY_TYPE = 'N'" & _
                         "       AND a.NOTE_CASH = 'Y'"

                Dim paras(1) As IDbDataParameter

                paras(0) = ProviderFactory.CreateDataParameter("MB_NAME_1", sMB_NAME & "%")

                paras(1) = ProviderFactory.CreateDataParameter("MB_NAME_2", sMB_NAME & "%")

                Return Me.loadBySQL(sqlStr, paras)
            End If
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' ----------------------------------------------------------------------------
    ''' <summary>
    ''' 根據[所屬區代碼]取得MB_MEMREV的回收收據，補印收據資料
    ''' </summary>
    ''' <param name="sMB_AREA">MB_AREA所屬區代碼</param>
    ''' <returns>integer取得筆數</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Function LoadByMB_AREAForAPPEND(ByVal sMB_AREA As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT a.*, b.*" & _
                     "  FROM MB_MEMREV a, MB_MEMBER b" & _
                     " WHERE     a.MB_MEMSEQ = b.MB_MEMSEQ" & _
                     "       AND b.MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA " & _
                     "       AND a.DELFLAG IS NULL" & _
                     "       AND a.NOTE_CASH <> 'N'" & _
                     "       AND a.MB_PRINT = 'Y'" & _
                     "       AND a.MB_REBREV IS NULL"

            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_AREA", sMB_AREA)

            Return Me.loadBySQL(sqlStr, para)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' ----------------------------------------------------------------------------
    ''' <summary>
    ''' 根據[繳款日期起日][繳款日期迄日]取得MB_MEMREV的回收收據，補印收據資料
    ''' </summary>
    ''' <param name="iMB_TX_DATE_BEG">MB_TX_DATE繳款日期起日</param>
    ''' <param name="iMB_TX_DATE_END">MB_TX_DATE繳款日期迄日</param>
    ''' <returns>integer取得筆數</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Function LoadByMB_TX_DATEForAPPEND(ByVal iMB_TX_DATE_BEG As Decimal, ByVal iMB_TX_DATE_END As Decimal, ByVal sMB_AREA As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT a.*, b.*" & _
                     "  FROM MB_MEMREV a, MB_MEMBER b" & _
                     " WHERE a.MB_MEMSEQ = b.MB_MEMSEQ" & _
                     "   AND (CAST(SUBSTR(A.MB_TX_DATE, 1, 6) AS decimal) >= " & ProviderFactory.PositionPara & "MB_TX_DATE_BEG_1 AND" & _
                     "        CAST(SUBSTR(A.MB_TX_DATE, 1, 6) AS decimal) <= " & ProviderFactory.PositionPara & "MB_TX_DATE_END_1 )" & _
                     "   AND a.DELFLAG IS NULL" & _
                     "   AND a.NOTE_CASH<>'N' " & _
                     "   AND a.MB_PRINT='Y'" & _
                     "   AND a.MB_REBREV IS NULL"

            If Utility.isValidateData(sMB_AREA) Then
                sqlStr &= "   AND B.MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA "

                Dim paras(2) As IDbDataParameter

                paras(0) = ProviderFactory.CreateDataParameter("MB_TX_DATE_BEG_1", iMB_TX_DATE_BEG)

                paras(1) = ProviderFactory.CreateDataParameter("MB_TX_DATE_END_1", iMB_TX_DATE_END)

                paras(2) = ProviderFactory.CreateDataParameter("MB_AREA", sMB_AREA)

                Return Me.loadBySQL(sqlStr, paras)
            Else
                Dim paras(1) As IDbDataParameter

                paras(0) = ProviderFactory.CreateDataParameter("MB_TX_DATE_BEG_1", iMB_TX_DATE_BEG)

                paras(1) = ProviderFactory.CreateDataParameter("MB_TX_DATE_END_1", iMB_TX_DATE_END)

                Return Me.loadBySQL(sqlStr, paras)
            End If
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' ----------------------------------------------------------------------------
    ''' <summary>
    ''' 根據[姓名]取得MB_MEMREV的回收收據，補印收據資料
    ''' </summary>
    ''' <param name="sMB_NAME">MB_NAME姓名</param>
    ''' <returns>integer取得筆數</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Function LoadByMB_NAMEForAPPEND(ByVal sMB_NAME As String, ByVal sMB_AREA As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT a.*, b.*" & _
                     "  FROM MB_MEMREV a, MB_MEMBER b" & _
                     " WHERE a.MB_MEMSEQ = b.MB_MEMSEQ" & _
                     "       AND b.MB_NAME LIKE " & ProviderFactory.PositionPara & "MB_NAME " & _
                     "       AND a.DELFLAG IS NULL" & _
                     "       AND a.NOTE_CASH<>'N' " & _
                     "       AND a.MB_PRINT='Y' " & _
                     "       AND a.MB_REBREV IS NULL "

            If Utility.isValidateData(sMB_AREA) Then
                sqlStr &= "       AND B.MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA "

                Dim paras(1) As IDbDataParameter
                paras(0) = ProviderFactory.CreateDataParameter("MB_NAME", sMB_NAME & "%")
                paras(1) = ProviderFactory.CreateDataParameter("MB_AREA", sMB_AREA)

                Return Me.loadBySQL(sqlStr, paras)
            Else
                Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_NAME", sMB_NAME & "%")

                Return Me.loadBySQL(sqlStr, para)
            End If
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

#Region "報表查詢"
    Public Function QRY_1(ByVal iSTART As Decimal, ByVal iPageSize As Decimal) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT DISTINCT(MB_MEMSEQ) FROM MB_MEMREV ORDER BY MB_TX_DATE LIMIT " & iSTART & "," & iPageSize

            Return Me.loadBySQLOnlyDs(sqlStr)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function QRY_2(ByVal iSTART As Decimal, ByVal iPageSize As Decimal, ByVal D_BDay As Date, ByVal D_EDay As Date) As Integer
        Try
            Dim sqlStr As String = String.Empty

            Dim iD_BDay As Decimal = 0
            iD_BDay = D_BDay.Year * 10000 + D_BDay.Month * 100 + D_BDay.Day

            Dim iD_EDay As Decimal = 0
            iD_EDay = D_EDay.Year * 10000 + D_EDay.Month * 100 + D_EDay.Day

            sqlStr = "SELECT DISTINCT(MB_MEMSEQ) FROM MB_MEMREV WHERE MB_TX_DATE BETWEEN " & ProviderFactory.PositionPara & "D_BDay " & _
                     "  AND " & ProviderFactory.PositionPara & "D_EDay ORDER BY MB_TX_DATE LIMIT " & iSTART & "," & iPageSize

            Dim paras(1) As IDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("D_BDay", iD_BDay)
            paras(1) = ProviderFactory.CreateDataParameter("D_EDay", iD_EDay)

            Return Me.loadBySQLOnlyDs(sqlStr, paras)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function QRY_3(ByVal iSTART As Decimal, ByVal iPageSize As Decimal, ByVal D_BDay As Date, ByVal D_EDay As Date, ByVal sMB_AREA As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            Dim iD_BDay As Decimal = 0
            iD_BDay = D_BDay.Year * 10000 + D_BDay.Month * 100 + D_BDay.Day

            Dim iD_EDay As Decimal = 0
            iD_EDay = D_EDay.Year * 10000 + D_EDay.Month * 100 + D_EDay.Day

            sqlStr = "SELECT DISTINCT (A.MB_MEMSEQ) " & _
                     "  FROM MB_MEMREV A " & _
                     " WHERE     MB_TX_DATE BETWEEN " & ProviderFactory.PositionPara & "D_BDay " & _
                     "      AND " & ProviderFactory.PositionPara & "D_EDay " & _
                     "      AND EXISTS " & _
                     "              (SELECT * " & _
                     "                 FROM MB_MEMBER " & _
                     "                WHERE MB_MEMSEQ = A.MB_MEMSEQ AND MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA)" & _
                     " ORDER BY MB_TX_DATE " & _
                     "   LIMIT " & iSTART & ", " & iPageSize

            Dim paras(2) As IDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("D_BDay", iD_BDay)
            paras(1) = ProviderFactory.CreateDataParameter("D_EDay", iD_EDay)
            paras(2) = ProviderFactory.CreateDataParameter("MB_AREA", sMB_AREA)

            Return Me.loadBySQLOnlyDs(sqlStr, paras)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function QRY_4(ByVal iSTART As Decimal, ByVal iPageSize As Decimal, ByVal D_BDay As Date, ByVal D_EDay As Date, ByVal sMB_LEADER As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            Dim iD_BDay As Decimal = 0
            iD_BDay = D_BDay.Year * 10000 + D_BDay.Month * 100 + D_BDay.Day

            Dim iD_EDay As Decimal = 0
            iD_EDay = D_EDay.Year * 10000 + D_EDay.Month * 100 + D_EDay.Day

            sqlStr = "SELECT DISTINCT (A.MB_MEMSEQ) " & _
                     "  FROM MB_MEMREV A " & _
                     " WHERE     MB_TX_DATE BETWEEN " & ProviderFactory.PositionPara & "D_BDay " & _
                     "      AND " & ProviderFactory.PositionPara & "D_EDay " & _
                     "      AND EXISTS " & _
                     "              (SELECT * " & _
                     "                 FROM MB_MEMBER " & _
                     "                WHERE MB_MEMSEQ = A.MB_MEMSEQ AND MB_LEADER = " & ProviderFactory.PositionPara & "MB_LEADER)" & _
                     " ORDER BY MB_TX_DATE " & _
                     "   LIMIT " & iSTART & ", " & iPageSize

            Dim paras(2) As IDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("D_BDay", iD_BDay)
            paras(1) = ProviderFactory.CreateDataParameter("D_EDay", iD_EDay)
            paras(2) = ProviderFactory.CreateDataParameter("MB_LEADER", sMB_LEADER)

            Return Me.loadBySQLOnlyDs(sqlStr, paras)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function QRY_5(ByVal iSTART As Decimal, ByVal iPageSize As Decimal, ByVal D_BDay As Date, ByVal D_EDay As Date, ByVal sMB_AREA As String, ByVal sMB_LEADER As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            Dim iD_BDay As Decimal = 0
            iD_BDay = D_BDay.Year * 10000 + D_BDay.Month * 100 + D_BDay.Day

            Dim iD_EDay As Decimal = 0
            iD_EDay = D_EDay.Year * 10000 + D_EDay.Month * 100 + D_EDay.Day

            sqlStr = "SELECT DISTINCT (A.MB_MEMSEQ) " & _
                     "  FROM MB_MEMREV A " & _
                     " WHERE     MB_TX_DATE BETWEEN " & ProviderFactory.PositionPara & "D_BDay " & _
                     "      AND " & ProviderFactory.PositionPara & "D_EDay " & _
                     "      AND EXISTS " & _
                     "              (SELECT * " & _
                     "                 FROM MB_MEMBER " & _
                     "                WHERE MB_MEMSEQ = A.MB_MEMSEQ AND MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA " & _
                     "                          AND MB_LEADER = " & ProviderFactory.PositionPara & "MB_LEADER)" & _
                     " ORDER BY MB_TX_DATE " & _
                     "   LIMIT " & iSTART & ", " & iPageSize

            Dim paras(3) As IDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("D_BDay", iD_BDay)
            paras(1) = ProviderFactory.CreateDataParameter("D_EDay", iD_EDay)
            paras(2) = ProviderFactory.CreateDataParameter("MB_AREA", sMB_AREA)
            paras(3) = ProviderFactory.CreateDataParameter("MB_LEADER", sMB_LEADER)

            Return Me.loadBySQLOnlyDs(sqlStr, paras)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function QRY_6(ByVal iSTART As Decimal, ByVal iPageSize As Decimal, ByVal sMB_AREA As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT DISTINCT (A.MB_MEMSEQ) FROM MB_MEMREV A " & _
                     "  WHERE EXISTS (SELECT * FROM MB_MEMBER WHERE MB_MEMSEQ = A.MB_MEMSEQ  " & _
                     "                 AND MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA ) " & _
                     "  ORDER BY MB_TX_DATE LIMIT " & iSTART & "," & iPageSize

            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_AREA", sMB_AREA)

            Return Me.loadBySQLOnlyDs(sqlStr, para)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function QRY_7(ByVal iSTART As Decimal, ByVal iPageSize As Decimal, ByVal sMB_AREA As String, ByVal sMB_LEADER As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT DISTINCT (A.MB_MEMSEQ) FROM MB_MEMREV A " & _
                     "  WHERE EXISTS (SELECT * FROM MB_MEMBER WHERE MB_MEMSEQ = A.MB_MEMSEQ  " & _
                     "                 AND MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA " & _
                     "                 AND MB_LEADER = " & ProviderFactory.PositionPara & "MB_LEADER ) " & _
                     "  ORDER BY MB_TX_DATE LIMIT " & iSTART & "," & iPageSize

            Dim paras(1) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("MB_AREA", sMB_AREA)
            paras(1) = ProviderFactory.CreateDataParameter("MB_LEADER", sMB_LEADER)

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

#Region "報表查詢(EXCEL)"
    Public Function QRY_EXCEL_1() As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT DISTINCT(MB_MEMSEQ) FROM MB_MEMREV ORDER BY MB_TX_DATE "

            Return Me.loadBySQLOnlyDs(sqlStr)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function QRY_EXCEL_2(ByVal D_BDay As Date, ByVal D_EDay As Date) As Integer
        Try
            Dim sqlStr As String = String.Empty

            Dim iD_BDay As Decimal = 0
            iD_BDay = D_BDay.Year * 10000 + D_BDay.Month * 100 + D_BDay.Day

            Dim iD_EDay As Decimal = 0
            iD_EDay = D_EDay.Year * 10000 + D_EDay.Month * 100 + D_EDay.Day

            sqlStr = "SELECT DISTINCT(MB_MEMSEQ) FROM MB_MEMREV WHERE MB_TX_DATE BETWEEN " & ProviderFactory.PositionPara & "D_BDay " & _
                     "  AND " & ProviderFactory.PositionPara & "D_EDay ORDER BY MB_TX_DATE "

            Dim paras(1) As IDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("D_BDay", iD_BDay)
            paras(1) = ProviderFactory.CreateDataParameter("D_EDay", iD_EDay)

            Return Me.loadBySQLOnlyDs(sqlStr, paras)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function QRY_EXCEL_3(ByVal D_BDay As Date, ByVal D_EDay As Date, ByVal sMB_AREA As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            Dim iD_BDay As Decimal = 0
            iD_BDay = D_BDay.Year * 10000 + D_BDay.Month * 100 + D_BDay.Day

            Dim iD_EDay As Decimal = 0
            iD_EDay = D_EDay.Year * 10000 + D_EDay.Month * 100 + D_EDay.Day

            sqlStr = "SELECT DISTINCT (A.MB_MEMSEQ) " & _
                     "  FROM MB_MEMREV A " & _
                     " WHERE     MB_TX_DATE BETWEEN " & ProviderFactory.PositionPara & "D_BDay " & _
                     "      AND " & ProviderFactory.PositionPara & "D_EDay " & _
                     "      AND EXISTS " & _
                     "              (SELECT * " & _
                     "                 FROM MB_MEMBER " & _
                     "                WHERE MB_MEMSEQ = A.MB_MEMSEQ AND MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA)" & _
                     " ORDER BY MB_TX_DATE "

            Dim paras(2) As IDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("D_BDay", iD_BDay)
            paras(1) = ProviderFactory.CreateDataParameter("D_EDay", iD_EDay)
            paras(2) = ProviderFactory.CreateDataParameter("MB_AREA", sMB_AREA)

            Return Me.loadBySQLOnlyDs(sqlStr, paras)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function QRY_EXCEL_4(ByVal D_BDay As Date, ByVal D_EDay As Date, ByVal sMB_LEADER As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            Dim iD_BDay As Decimal = 0
            iD_BDay = D_BDay.Year * 10000 + D_BDay.Month * 100 + D_BDay.Day

            Dim iD_EDay As Decimal = 0
            iD_EDay = D_EDay.Year * 10000 + D_EDay.Month * 100 + D_EDay.Day

            sqlStr = "SELECT DISTINCT (A.MB_MEMSEQ) " & _
                     "  FROM MB_MEMREV A " & _
                     " WHERE     MB_TX_DATE BETWEEN " & ProviderFactory.PositionPara & "D_BDay " & _
                     "      AND " & ProviderFactory.PositionPara & "D_EDay " & _
                     "      AND EXISTS " & _
                     "              (SELECT * " & _
                     "                 FROM MB_MEMBER " & _
                     "                WHERE MB_MEMSEQ = A.MB_MEMSEQ AND MB_LEADER = " & ProviderFactory.PositionPara & "MB_LEADER)" & _
                     " ORDER BY MB_TX_DATE "

            Dim paras(2) As IDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("D_BDay", iD_BDay)
            paras(1) = ProviderFactory.CreateDataParameter("D_EDay", iD_EDay)
            paras(2) = ProviderFactory.CreateDataParameter("MB_LEADER", sMB_LEADER)

            Return Me.loadBySQLOnlyDs(sqlStr, paras)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function QRY_EXCEL_5(ByVal D_BDay As Date, ByVal D_EDay As Date, ByVal sMB_AREA As String, ByVal sMB_LEADER As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            Dim iD_BDay As Decimal = 0
            iD_BDay = D_BDay.Year * 10000 + D_BDay.Month * 100 + D_BDay.Day

            Dim iD_EDay As Decimal = 0
            iD_EDay = D_EDay.Year * 10000 + D_EDay.Month * 100 + D_EDay.Day

            sqlStr = "SELECT DISTINCT (A.MB_MEMSEQ) " & _
                     "  FROM MB_MEMREV A " & _
                     " WHERE     MB_TX_DATE BETWEEN " & ProviderFactory.PositionPara & "D_BDay " & _
                     "      AND " & ProviderFactory.PositionPara & "D_EDay " & _
                     "      AND EXISTS " & _
                     "              (SELECT * " & _
                     "                 FROM MB_MEMBER " & _
                     "                WHERE MB_MEMSEQ = A.MB_MEMSEQ AND MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA " & _
                     "                          AND MB_LEADER = " & ProviderFactory.PositionPara & "MB_LEADER)" & _
                     " ORDER BY MB_TX_DATE "

            Dim paras(3) As IDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("D_BDay", iD_BDay)
            paras(1) = ProviderFactory.CreateDataParameter("D_EDay", iD_EDay)
            paras(2) = ProviderFactory.CreateDataParameter("MB_AREA", sMB_AREA)
            paras(3) = ProviderFactory.CreateDataParameter("MB_LEADER", sMB_LEADER)

            Return Me.loadBySQLOnlyDs(sqlStr, paras)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function QRY_EXCEL_6(ByVal sMB_AREA As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT DISTINCT (A.MB_MEMSEQ) FROM MB_MEMREV A " & _
                     "  WHERE EXISTS (SELECT * FROM MB_MEMBER WHERE MB_MEMSEQ = A.MB_MEMSEQ  " & _
                     "                 AND MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA ) " & _
                     "  ORDER BY MB_TX_DATE "

            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_AREA", sMB_AREA)

            Return Me.loadBySQLOnlyDs(sqlStr, para)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function QRY_EXCEL_7(ByVal sMB_AREA As String, ByVal sMB_LEADER As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT DISTINCT (A.MB_MEMSEQ) FROM MB_MEMREV A " & _
                     "  WHERE EXISTS (SELECT * FROM MB_MEMBER WHERE MB_MEMSEQ = A.MB_MEMSEQ  " & _
                     "                 AND MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA " & _
                     "                 AND MB_LEADER = " & ProviderFactory.PositionPara & "MB_LEADER ) " & _
                     "  ORDER BY MB_TX_DATE "

            Dim paras(1) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("MB_AREA", sMB_AREA)
            paras(1) = ProviderFactory.CreateDataParameter("MB_LEADER", sMB_LEADER)

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

#Region "報表統計"
    Public Function COUNT_1() As Decimal
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT IFNULL(COUNT(DISTINCT(MB_MEMSEQ)),0) FROM MB_MEMREV "

            If Me.getDatabaseManager.isTransaction Then
                Return DBObject.ExecuteScalar(Me.getDatabaseManager.getTransaction, CommandType.Text, sqlStr)
            Else
                Return DBObject.ExecuteScalar(Me.getDatabaseManager.getConnection, CommandType.Text, sqlStr)
            End If
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function COUNT_2(ByVal D_BDay As Date, ByVal D_EDay As Date) As Decimal
        Try
            Dim sqlStr As String = String.Empty

            Dim iD_BDay As Decimal = 0
            iD_BDay = D_BDay.Year * 10000 + D_BDay.Month * 100 + D_BDay.Day

            Dim iD_EDay As Decimal = 0
            iD_EDay = D_EDay.Year * 10000 + D_EDay.Month * 100 + D_EDay.Day

            sqlStr = "SELECT IFNULL(COUNT(DISTINCT(MB_MEMSEQ)),0) FROM MB_MEMREV WHERE MB_TX_DATE BETWEEN " & ProviderFactory.PositionPara & "D_BDay " & _
                     "  AND " & ProviderFactory.PositionPara & "D_EDay "

            Dim paras(1) As IDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("D_BDay", iD_BDay)
            paras(1) = ProviderFactory.CreateDataParameter("D_EDay", iD_EDay)

            If Me.getDatabaseManager.isTransaction Then
                Return DBObject.ExecuteScalar(Me.getDatabaseManager.getTransaction, CommandType.Text, sqlStr, paras)
            Else
                Return DBObject.ExecuteScalar(Me.getDatabaseManager.getConnection, CommandType.Text, sqlStr, paras)
            End If
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function COUNT_3(ByVal D_BDay As Date, ByVal D_EDay As Date, ByVal sMB_AREA As String) As Decimal
        Try
            Dim sqlStr As String = String.Empty

            Dim iD_BDay As Decimal = 0
            iD_BDay = D_BDay.Year * 10000 + D_BDay.Month * 100 + D_BDay.Day

            Dim iD_EDay As Decimal = 0
            iD_EDay = D_EDay.Year * 10000 + D_EDay.Month * 100 + D_EDay.Day

            sqlStr = "SELECT IFNULL(COUNT(DISTINCT (A.MB_MEMSEQ)),0) " & _
                     "  FROM MB_MEMREV A " & _
                     " WHERE     MB_TX_DATE BETWEEN " & ProviderFactory.PositionPara & "D_BDay " & _
                     "      AND " & ProviderFactory.PositionPara & "D_EDay " & _
                     "      AND EXISTS " & _
                     "              (SELECT * " & _
                     "                 FROM MB_MEMBER " & _
                     "                WHERE MB_MEMSEQ = A.MB_MEMSEQ AND MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA)"

            Dim paras(2) As IDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("D_BDay", iD_BDay)
            paras(1) = ProviderFactory.CreateDataParameter("D_EDay", iD_EDay)
            paras(2) = ProviderFactory.CreateDataParameter("MB_AREA", sMB_AREA)

            If Me.getDatabaseManager.isTransaction Then
                Return DBObject.ExecuteScalar(Me.getDatabaseManager.getTransaction, CommandType.Text, sqlStr, paras)
            Else
                Return DBObject.ExecuteScalar(Me.getDatabaseManager.getConnection, CommandType.Text, sqlStr, paras)
            End If
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function COUNT_4(ByVal D_BDay As Date, ByVal D_EDay As Date, ByVal sMB_LEADER As String) As Decimal
        Try
            Dim sqlStr As String = String.Empty

            Dim iD_BDay As Decimal = 0
            iD_BDay = D_BDay.Year * 10000 + D_BDay.Month * 100 + D_BDay.Day

            Dim iD_EDay As Decimal = 0
            iD_EDay = D_EDay.Year * 10000 + D_EDay.Month * 100 + D_EDay.Day

            sqlStr = "SELECT IFNULL(COUNT(DISTINCT (A.MB_MEMSEQ)),0) " & _
                     "  FROM MB_MEMREV A " & _
                     " WHERE     MB_TX_DATE BETWEEN " & ProviderFactory.PositionPara & "D_BDay " & _
                     "      AND " & ProviderFactory.PositionPara & "D_EDay " & _
                     "      AND EXISTS " & _
                     "              (SELECT * " & _
                     "                 FROM MB_MEMBER " & _
                     "                WHERE MB_MEMSEQ = A.MB_MEMSEQ AND MB_LEADER = " & ProviderFactory.PositionPara & "MB_LEADER)"

            Dim paras(2) As IDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("D_BDay", iD_BDay)
            paras(1) = ProviderFactory.CreateDataParameter("D_EDay", iD_EDay)
            paras(2) = ProviderFactory.CreateDataParameter("MB_LEADER", sMB_LEADER)

            If Me.getDatabaseManager.isTransaction Then
                Return DBObject.ExecuteScalar(Me.getDatabaseManager.getTransaction, CommandType.Text, sqlStr, paras)
            Else
                Return DBObject.ExecuteScalar(Me.getDatabaseManager.getConnection, CommandType.Text, sqlStr, paras)
            End If
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function COUNT_5(ByVal D_BDay As Date, ByVal D_EDay As Date, ByVal sMB_AREA As String, ByVal sMB_LEADER As String) As Decimal
        Try
            Dim sqlStr As String = String.Empty

            Dim iD_BDay As Decimal = 0
            iD_BDay = D_BDay.Year * 10000 + D_BDay.Month * 100 + D_BDay.Day

            Dim iD_EDay As Decimal = 0
            iD_EDay = D_EDay.Year * 10000 + D_EDay.Month * 100 + D_EDay.Day

            sqlStr = "SELECT IFNULL(COUNT(DISTINCT (A.MB_MEMSEQ)),0) " & _
                     "  FROM MB_MEMREV A " & _
                     " WHERE     MB_TX_DATE BETWEEN " & ProviderFactory.PositionPara & "D_BDay " & _
                     "      AND " & ProviderFactory.PositionPara & "D_EDay " & _
                     "      AND EXISTS " & _
                     "              (SELECT * " & _
                     "                 FROM MB_MEMBER " & _
                     "                WHERE MB_MEMSEQ = A.MB_MEMSEQ AND MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA " & _
                     "                          AND MB_LEADER = " & ProviderFactory.PositionPara & "MB_LEADER)"

            Dim paras(3) As IDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("D_BDay", iD_BDay)
            paras(1) = ProviderFactory.CreateDataParameter("D_EDay", iD_EDay)
            paras(2) = ProviderFactory.CreateDataParameter("MB_AREA", sMB_AREA)
            paras(3) = ProviderFactory.CreateDataParameter("MB_LEADER", sMB_LEADER)

            If Me.getDatabaseManager.isTransaction Then
                Return DBObject.ExecuteScalar(Me.getDatabaseManager.getTransaction, CommandType.Text, sqlStr, paras)
            Else
                Return DBObject.ExecuteScalar(Me.getDatabaseManager.getConnection, CommandType.Text, sqlStr, paras)
            End If
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function COUNT_6(ByVal sMB_AREA As String) As Decimal
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT IFNULL(COUNT(DISTINCT (A.MB_MEMSEQ)),0) FROM MB_MEMREV A " & _
                     "  WHERE EXISTS (SELECT * FROM MB_MEMBER WHERE MB_MEMSEQ = A.MB_MEMSEQ  " & _
                     "                 AND MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA ) "

            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_AREA", sMB_AREA)

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

    Public Function COUNT_7(ByVal sMB_AREA As String, ByVal sMB_LEADER As String) As Decimal
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT IFNULL(COUNT(DISTINCT (A.MB_MEMSEQ)),0) FROM MB_MEMREV A " & _
                     "  WHERE EXISTS (SELECT * FROM MB_MEMBER WHERE MB_MEMSEQ = A.MB_MEMSEQ  " & _
                     "                 AND MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA " & _
                     "                 AND MB_LEADER = " & ProviderFactory.PositionPara & "MB_LEADER ) "

            Dim paras(1) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("MB_AREA", sMB_AREA)
            paras(1) = ProviderFactory.CreateDataParameter("MB_LEADER", sMB_LEADER)

            If Me.getDatabaseManager.isTransaction Then
                Return DBObject.ExecuteScalar(Me.getDatabaseManager.getTransaction, CommandType.Text, sqlStr, paras)
            Else
                Return DBObject.ExecuteScalar(Me.getDatabaseManager.getConnection, CommandType.Text, sqlStr, paras)
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

#Region "報表明細"
    Public Function DTL_1(ByVal iMB_MEMSEQ As Decimal) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT A.MB_TX_DATE, A.MB_SEQNO, A.MB_TOTFEE, A.MB_ITEMID, B.* " & _
                     "  FROM MB_MEMREV A, MB_MEMBER B " & _
                     " WHERE A.MB_MEMSEQ = " & ProviderFactory.PositionPara & "MB_MEMSEQ AND A.MB_MEMSEQ = B.MB_MEMSEQ " & _
                     "  ORDER BY A.MB_TX_DATE, A.MB_SEQNO"

            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_MEMSEQ", iMB_MEMSEQ)

            Return Me.loadBySQLOnlyDs(sqlStr, para)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function DTL_2(ByVal iMB_MEMSEQ As Decimal, ByVal D_BDay As Date, ByVal D_EDay As Date) As Integer
        Try
            Dim sqlStr As String = String.Empty

            Dim iD_BDay As Decimal = 0
            iD_BDay = D_BDay.Year * 10000 + D_BDay.Month * 100 + D_BDay.Day

            Dim iD_EDay As Decimal = 0
            iD_EDay = D_EDay.Year * 10000 + D_EDay.Month * 100 + D_EDay.Day

            sqlStr = "SELECT A.MB_TX_DATE, A.MB_SEQNO, A.MB_TOTFEE, A.MB_ITEMID, B.* " & _
                     "  FROM MB_MEMREV A, MB_MEMBER B " & _
                     " WHERE     A.MB_MEMSEQ = " & ProviderFactory.PositionPara & "MB_MEMSEQ " & _
                     "       AND A.MB_TX_DATE BETWEEN " & ProviderFactory.PositionPara & "D_BDay " & _
                     "       AND " & ProviderFactory.PositionPara & "D_EDay " & _
                     "       AND A.MB_MEMSEQ = B.MB_MEMSEQ " & _
                     "  ORDER BY A.MB_TX_DATE, A.MB_SEQNO "

            Dim paras(2) As IDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("MB_MEMSEQ", iMB_MEMSEQ)
            paras(1) = ProviderFactory.CreateDataParameter("D_BDay", iD_BDay)
            paras(2) = ProviderFactory.CreateDataParameter("D_EDay", iD_EDay)

            Return Me.loadBySQLOnlyDs(sqlStr, paras)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function DTL_3(ByVal iMB_MEMSEQ As Decimal, ByVal D_BDay As Date, ByVal D_EDay As Date) As Integer
        Try
            Dim sqlStr As String = String.Empty

            Dim iD_BDay As Decimal = 0
            iD_BDay = D_BDay.Year * 10000 + D_BDay.Month * 100 + D_BDay.Day

            Dim iD_EDay As Decimal = 0
            iD_EDay = D_EDay.Year * 10000 + D_EDay.Month * 100 + D_EDay.Day

            sqlStr = "SELECT A.MB_TX_DATE, A.MB_SEQNO, A.MB_TOTFEE, A.MB_ITEMID, B.* " & _
                     "  FROM MB_MEMREV A, MB_MEMBER B " & _
                     " WHERE     A.MB_MEMSEQ = " & ProviderFactory.PositionPara & "MB_MEMSEQ " & _
                     "       AND A.MB_TX_DATE BETWEEN " & ProviderFactory.PositionPara & "D_BDay " & _
                     "       AND " & ProviderFactory.PositionPara & "D_EDay " & _
                     "       AND A.MB_MEMSEQ = B.MB_MEMSEQ " & _
                     "  ORDER BY A.MB_TX_DATE, A.MB_SEQNO "

            Dim paras(2) As IDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("MB_MEMSEQ", iMB_MEMSEQ)
            paras(1) = ProviderFactory.CreateDataParameter("D_BDay", iD_BDay)
            paras(2) = ProviderFactory.CreateDataParameter("D_EDay", iD_EDay)

            Return Me.loadBySQLOnlyDs(sqlStr, paras)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function DTL_4(ByVal iMB_MEMSEQ As Decimal, ByVal D_BDay As Date, ByVal D_EDay As Date) As Integer
        Try
            Dim sqlStr As String = String.Empty

            Dim iD_BDay As Decimal = 0
            iD_BDay = D_BDay.Year * 10000 + D_BDay.Month * 100 + D_BDay.Day

            Dim iD_EDay As Decimal = 0
            iD_EDay = D_EDay.Year * 10000 + D_EDay.Month * 100 + D_EDay.Day

            sqlStr = "SELECT A.MB_TX_DATE, A.MB_SEQNO, A.MB_TOTFEE, A.MB_ITEMID, B.* " & _
                     "  FROM MB_MEMREV A, MB_MEMBER B " & _
                     " WHERE     A.MB_MEMSEQ = " & ProviderFactory.PositionPara & "MB_MEMSEQ " & _
                     "       AND A.MB_TX_DATE BETWEEN " & ProviderFactory.PositionPara & "D_BDay " & _
                     "       AND " & ProviderFactory.PositionPara & "D_EDay " & _
                     "       AND A.MB_MEMSEQ = B.MB_MEMSEQ " & _
                     "  ORDER BY A.MB_TX_DATE, A.MB_SEQNO "

            Dim paras(2) As IDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("MB_MEMSEQ", iMB_MEMSEQ)
            paras(1) = ProviderFactory.CreateDataParameter("D_BDay", iD_BDay)
            paras(2) = ProviderFactory.CreateDataParameter("D_EDay", iD_EDay)

            Return Me.loadBySQLOnlyDs(sqlStr, paras)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function DTL_5(ByVal iMB_MEMSEQ As Decimal, ByVal D_BDay As Date, ByVal D_EDay As Date) As Integer
        Try
            Dim sqlStr As String = String.Empty

            Dim iD_BDay As Decimal = 0
            iD_BDay = D_BDay.Year * 10000 + D_BDay.Month * 100 + D_BDay.Day

            Dim iD_EDay As Decimal = 0
            iD_EDay = D_EDay.Year * 10000 + D_EDay.Month * 100 + D_EDay.Day

            sqlStr = "SELECT A.MB_TX_DATE, A.MB_SEQNO, A.MB_TOTFEE, A.MB_ITEMID, B.* " & _
                     "  FROM MB_MEMREV A, MB_MEMBER B " & _
                     " WHERE     A.MB_MEMSEQ = " & ProviderFactory.PositionPara & "MB_MEMSEQ " & _
                     "       AND A.MB_TX_DATE BETWEEN " & ProviderFactory.PositionPara & "D_BDay " & _
                     "       AND " & ProviderFactory.PositionPara & "D_EDay " & _
                     "       AND A.MB_MEMSEQ = B.MB_MEMSEQ " & _
                     "  ORDER BY A.MB_TX_DATE, A.MB_SEQNO "

            Dim paras(2) As IDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("MB_MEMSEQ", iMB_MEMSEQ)
            paras(1) = ProviderFactory.CreateDataParameter("D_BDay", iD_BDay)
            paras(2) = ProviderFactory.CreateDataParameter("D_EDay", iD_EDay)

            Return Me.loadBySQLOnlyDs(sqlStr, paras)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function DTL_6(ByVal iMB_MEMSEQ As Decimal) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT A.MB_TX_DATE, A.MB_SEQNO, A.MB_TOTFEE, A.MB_ITEMID, B.* " & _
                     "  FROM MB_MEMREV A, MB_MEMBER B " & _
                     " WHERE     A.MB_MEMSEQ = " & ProviderFactory.PositionPara & "MB_MEMSEQ " & _
                     "       AND A.MB_MEMSEQ = B.MB_MEMSEQ " & _
                     "  ORDER BY A.MB_TX_DATE, A.MB_SEQNO "

            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_MEMSEQ", iMB_MEMSEQ)

            Return Me.loadBySQLOnlyDs(sqlStr, para)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function DTL_7(ByVal iMB_MEMSEQ As Decimal) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT A.MB_TX_DATE, A.MB_SEQNO, A.MB_TOTFEE, A.MB_ITEMID, B.* " & _
                     "  FROM MB_MEMREV A, MB_MEMBER B " & _
                     " WHERE     A.MB_MEMSEQ = " & ProviderFactory.PositionPara & "MB_MEMSEQ " & _
                     "       AND A.MB_MEMSEQ = B.MB_MEMSEQ " & _
                     "  ORDER BY A.MB_TX_DATE, A.MB_SEQNO "

            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_MEMSEQ", iMB_MEMSEQ)

            Return Me.loadBySQLOnlyDs(sqlStr, para)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function
#End Region

#Region "報表查詢(會員收入查詢)"
    Public Function QRY_INCOME_1(ByVal iSTART As Decimal, ByVal iPageSize As Decimal, ByVal iB_MB_TX_DATE As Decimal, ByVal iE_MB_TX_DATE As Decimal) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT A.*, B.MB_AREA, B.MB_LEADER, B.MB_NAME " & _
                     "  FROM MB_MEMREV A, MB_MEMBER B " & _
                     " WHERE     A.MB_MEMSEQ = B.MB_MEMSEQ " & _
                     "       AND A.MB_TX_DATE >= " & ProviderFactory.PositionPara & "B_MB_TX_DATE " & _
                     "       AND A.MB_TX_DATE <= " & ProviderFactory.PositionPara & "E_MB_TX_DATE " & _
                     "  ORDER BY MB_TX_DATE, MB_MEMSEQ, MB_SEQNO LIMIT " & iSTART & "," & iPageSize

            Dim paras(1) As IDbDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("B_MB_TX_DATE", iB_MB_TX_DATE)

            paras(1) = ProviderFactory.CreateDataParameter("E_MB_TX_DATE", iE_MB_TX_DATE)

            Return Me.loadBySQLOnlyDs(sqlStr, paras)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function QRY_INCOME_2(ByVal iSTART As Decimal, ByVal iPageSize As Decimal, ByVal iB_MB_TX_DATE As Decimal, ByVal iE_MB_TX_DATE As Decimal, ByVal sMB_MEMTYP As String, ByVal sMB_FEETYPE As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT A.*, B.MB_AREA, B.MB_LEADER, B.MB_NAME " & _
                     "  FROM MB_MEMREV A, MB_MEMBER B " & _
                     " WHERE     A.MB_MEMSEQ = B.MB_MEMSEQ " & _
                     "       AND A.MB_TX_DATE >= " & ProviderFactory.PositionPara & "B_MB_TX_DATE " & _
                     "       AND A.MB_TX_DATE <= " & ProviderFactory.PositionPara & "E_MB_TX_DATE " & _
                     "       AND A.MB_MEMTYP = " & ProviderFactory.PositionPara & "MB_MEMTYP " & _
                     "       AND A.MB_FEETYPE = " & ProviderFactory.PositionPara & "MB_FEETYPE " & _
                     "  ORDER BY MB_TX_DATE, MB_MEMSEQ, MB_SEQNO LIMIT " & iSTART & "," & iPageSize

            Dim paras(3) As IDbDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("B_MB_TX_DATE", iB_MB_TX_DATE)

            paras(1) = ProviderFactory.CreateDataParameter("E_MB_TX_DATE", iE_MB_TX_DATE)

            paras(2) = ProviderFactory.CreateDataParameter("MB_MEMTYP", sMB_MEMTYP)

            paras(3) = ProviderFactory.CreateDataParameter("MB_FEETYPE", sMB_FEETYPE)

            Return Me.loadBySQLOnlyDs(sqlStr, paras)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function QRY_INCOME_3(ByVal iSTART As Decimal, ByVal iPageSize As Decimal, ByVal iB_MB_TX_DATE As Decimal, ByVal iE_MB_TX_DATE As Decimal, ByVal sMB_FEETYPE As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT A.*, B.MB_AREA, B.MB_LEADER, B.MB_NAME " & _
                     "  FROM MB_MEMREV A, MB_MEMBER B " & _
                     " WHERE     A.MB_MEMSEQ = B.MB_MEMSEQ " & _
                     "       AND A.MB_TX_DATE >= " & ProviderFactory.PositionPara & "B_MB_TX_DATE " & _
                     "       AND A.MB_TX_DATE <= " & ProviderFactory.PositionPara & "E_MB_TX_DATE " & _
                     "       AND A.MB_FEETYPE = " & ProviderFactory.PositionPara & "MB_FEETYPE " & _
                     "  ORDER BY MB_TX_DATE, MB_MEMSEQ, MB_SEQNO LIMIT " & iSTART & "," & iPageSize

            Dim paras(2) As IDbDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("B_MB_TX_DATE", iB_MB_TX_DATE)

            paras(1) = ProviderFactory.CreateDataParameter("E_MB_TX_DATE", iE_MB_TX_DATE)

            paras(2) = ProviderFactory.CreateDataParameter("MB_FEETYPE", sMB_FEETYPE)

            Return Me.loadBySQLOnlyDs(sqlStr, paras)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function QRY_INCOME_4(ByVal iSTART As Decimal, ByVal iPageSize As Decimal, ByVal sMB_MEMTYP As String, ByVal sMB_FEETYPE As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT A.*, B.MB_AREA, B.MB_LEADER, B.MB_NAME " & _
                     "  FROM MB_MEMREV A, MB_MEMBER B " & _
                     " WHERE     A.MB_MEMSEQ = B.MB_MEMSEQ " & _
                     "       AND A.MB_MEMTYP = " & ProviderFactory.PositionPara & "MB_MEMTYP " & _
                     "       AND A.MB_FEETYPE = " & ProviderFactory.PositionPara & "MB_FEETYPE " & _
                     "  ORDER BY MB_TX_DATE, MB_MEMSEQ, MB_SEQNO LIMIT " & iSTART & "," & iPageSize

            Dim paras(1) As IDbDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("MB_MEMTYP", sMB_MEMTYP)

            paras(1) = ProviderFactory.CreateDataParameter("MB_FEETYPE", sMB_FEETYPE)

            Return Me.loadBySQLOnlyDs(sqlStr, paras)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function QRY_INCOME_5(ByVal iSTART As Decimal, ByVal iPageSize As Decimal, ByVal sMB_MEMTYP As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT A.*, B.MB_AREA, B.MB_LEADER, B.MB_NAME " & _
                     "  FROM MB_MEMREV A, MB_MEMBER B " & _
                     " WHERE     A.MB_MEMSEQ = B.MB_MEMSEQ " & _
                     "       AND A.MB_MEMTYP = " & ProviderFactory.PositionPara & "MB_MEMTYP " & _
                     "  ORDER BY MB_TX_DATE, MB_MEMSEQ, MB_SEQNO LIMIT " & iSTART & "," & iPageSize

            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_MEMTYP", sMB_MEMTYP)

            Return Me.loadBySQLOnlyDs(sqlStr, para)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function QRY_INCOME_6(ByVal iSTART As Decimal, ByVal iPageSize As Decimal, ByVal sMB_FEETYPE As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT A.*, B.MB_AREA, B.MB_LEADER, B.MB_NAME " & _
                     "  FROM MB_MEMREV A, MB_MEMBER B " & _
                     " WHERE     A.MB_MEMSEQ = B.MB_MEMSEQ " & _
                     "       AND A.MB_FEETYPE = " & ProviderFactory.PositionPara & "MB_FEETYPE " & _
                     "  ORDER BY MB_TX_DATE, MB_MEMSEQ, MB_SEQNO LIMIT " & iSTART & "," & iPageSize

            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_FEETYPE", sMB_FEETYPE)

            Return Me.loadBySQLOnlyDs(sqlStr, para)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function QRY_INCOME_7(ByVal iSTART As Decimal, ByVal iPageSize As Decimal) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT A.*, B.MB_AREA, B.MB_LEADER, B.MB_NAME " & _
                     "  FROM MB_MEMREV A, MB_MEMBER B " & _
                     " WHERE     A.MB_MEMSEQ = B.MB_MEMSEQ " & _
                     "  ORDER BY MB_TX_DATE, MB_MEMSEQ, MB_SEQNO LIMIT " & iSTART & "," & iPageSize

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

#Region "報表查詢(會員收入查詢)EXCEL"
    Public Function QRY_EXCEL_INCOME_1(ByVal iB_MB_TX_DATE As Decimal, ByVal iE_MB_TX_DATE As Decimal) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT A.*, B.MB_AREA, B.MB_LEADER, B.MB_NAME " & _
                     "  FROM MB_MEMREV A, MB_MEMBER B " & _
                     " WHERE     A.MB_MEMSEQ = B.MB_MEMSEQ " & _
                     "       AND A.MB_TX_DATE >= " & ProviderFactory.PositionPara & "B_MB_TX_DATE " & _
                     "       AND A.MB_TX_DATE <= " & ProviderFactory.PositionPara & "E_MB_TX_DATE " & _
                     "  ORDER BY MB_TX_DATE, MB_MEMSEQ, MB_SEQNO "

            Dim paras(1) As IDbDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("B_MB_TX_DATE", iB_MB_TX_DATE)

            paras(1) = ProviderFactory.CreateDataParameter("E_MB_TX_DATE", iE_MB_TX_DATE)

            Return Me.loadBySQLOnlyDs(sqlStr, paras)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function QRY_EXCEL_INCOME_2(ByVal iB_MB_TX_DATE As Decimal, ByVal iE_MB_TX_DATE As Decimal, ByVal sMB_MEMTYP As String, ByVal sMB_FEETYPE As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT A.*, B.MB_AREA, B.MB_LEADER, B.MB_NAME " & _
                     "  FROM MB_MEMREV A, MB_MEMBER B " & _
                     " WHERE     A.MB_MEMSEQ = B.MB_MEMSEQ " & _
                     "       AND A.MB_TX_DATE >= " & ProviderFactory.PositionPara & "B_MB_TX_DATE " & _
                     "       AND A.MB_TX_DATE <= " & ProviderFactory.PositionPara & "E_MB_TX_DATE " & _
                     "       AND A.MB_MEMTYP = " & ProviderFactory.PositionPara & "MB_MEMTYP " & _
                     "       AND A.MB_FEETYPE = " & ProviderFactory.PositionPara & "MB_FEETYPE " & _
                     "  ORDER BY MB_TX_DATE, MB_MEMSEQ, MB_SEQNO "

            Dim paras(3) As IDbDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("B_MB_TX_DATE", iB_MB_TX_DATE)

            paras(1) = ProviderFactory.CreateDataParameter("E_MB_TX_DATE", iE_MB_TX_DATE)

            paras(2) = ProviderFactory.CreateDataParameter("MB_MEMTYP", sMB_MEMTYP)

            paras(3) = ProviderFactory.CreateDataParameter("MB_FEETYPE", sMB_FEETYPE)

            Return Me.loadBySQLOnlyDs(sqlStr, paras)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function QRY_EXCEL_INCOME_3(ByVal iB_MB_TX_DATE As Decimal, ByVal iE_MB_TX_DATE As Decimal, ByVal sMB_FEETYPE As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT A.*, B.MB_AREA, B.MB_LEADER, B.MB_NAME " & _
                     "  FROM MB_MEMREV A, MB_MEMBER B " & _
                     " WHERE     A.MB_MEMSEQ = B.MB_MEMSEQ " & _
                     "       AND A.MB_TX_DATE >= " & ProviderFactory.PositionPara & "B_MB_TX_DATE " & _
                     "       AND A.MB_TX_DATE <= " & ProviderFactory.PositionPara & "E_MB_TX_DATE " & _
                     "       AND A.MB_FEETYPE = " & ProviderFactory.PositionPara & "MB_FEETYPE " & _
                     "  ORDER BY MB_TX_DATE, MB_MEMSEQ, MB_SEQNO "

            Dim paras(2) As IDbDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("B_MB_TX_DATE", iB_MB_TX_DATE)

            paras(1) = ProviderFactory.CreateDataParameter("E_MB_TX_DATE", iE_MB_TX_DATE)

            paras(2) = ProviderFactory.CreateDataParameter("MB_FEETYPE", sMB_FEETYPE)

            Return Me.loadBySQLOnlyDs(sqlStr, paras)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function QRY_EXCEL_INCOME_4(ByVal sMB_MEMTYP As String, ByVal sMB_FEETYPE As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT A.*, B.MB_AREA, B.MB_LEADER, B.MB_NAME " & _
                     "  FROM MB_MEMREV A, MB_MEMBER B " & _
                     " WHERE     A.MB_MEMSEQ = B.MB_MEMSEQ " & _
                     "       AND A.MB_MEMTYP = " & ProviderFactory.PositionPara & "MB_MEMTYP " & _
                     "       AND A.MB_FEETYPE = " & ProviderFactory.PositionPara & "MB_FEETYPE " & _
                     "  ORDER BY MB_TX_DATE, MB_MEMSEQ, MB_SEQNO "

            Dim paras(1) As IDbDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("MB_MEMTYP", sMB_MEMTYP)

            paras(1) = ProviderFactory.CreateDataParameter("MB_FEETYPE", sMB_FEETYPE)

            Return Me.loadBySQLOnlyDs(sqlStr, paras)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function QRY_EXCEL_INCOME_5(ByVal sMB_MEMTYP As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT A.*, B.MB_AREA, B.MB_LEADER, B.MB_NAME " & _
                     "  FROM MB_MEMREV A, MB_MEMBER B " & _
                     " WHERE     A.MB_MEMSEQ = B.MB_MEMSEQ " & _
                     "       AND A.MB_MEMTYP = " & ProviderFactory.PositionPara & "MB_MEMTYP " & _
                     "  ORDER BY MB_TX_DATE, MB_MEMSEQ, MB_SEQNO "

            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_MEMTYP", sMB_MEMTYP)

            Return Me.loadBySQLOnlyDs(sqlStr, para)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function QRY_EXCEL_INCOME_6(ByVal sMB_FEETYPE As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT A.*, B.MB_AREA, B.MB_LEADER, B.MB_NAME " & _
                     "  FROM MB_MEMREV A, MB_MEMBER B " & _
                     " WHERE     A.MB_MEMSEQ = B.MB_MEMSEQ " & _
                     "       AND A.MB_FEETYPE = " & ProviderFactory.PositionPara & "MB_FEETYPE " & _
                     "  ORDER BY MB_TX_DATE, MB_MEMSEQ, MB_SEQNO "

            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_FEETYPE", sMB_FEETYPE)

            Return Me.loadBySQLOnlyDs(sqlStr, para)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function
#End Region


#Region "報表統計(會員收入查詢)"
    Public Function COUNT_INCOME_1(ByVal iB_YYYYMMDD As Decimal, ByVal iE_YYYYMMDD As Decimal) As Decimal
        Try
            Dim sqlStr As String = String.Empty

            'sqlStr = "SELECT IFNULL(COUNT(DISTINCT(MB_MEMSEQ)),0) FROM MB_MEMREV "

            sqlStr = "SELECT IFNULL(COUNT(*),0) " & _
                     "  FROM MB_MEMREV A, MB_MEMBER B " & _
                     " WHERE     A.MB_MEMSEQ = B.MB_MEMSEQ " & _
                     "       AND A.MB_TX_DATE >= " & ProviderFactory.PositionPara & "B_MB_TX_DATE " & _
                     "       AND A.MB_TX_DATE <= " & ProviderFactory.PositionPara & "E_MB_TX_DATE "

            Dim paras(1) As IDbDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("B_MB_TX_DATE", iB_YYYYMMDD)

            paras(1) = ProviderFactory.CreateDataParameter("E_MB_TX_DATE", iE_YYYYMMDD)

            If Me.getDatabaseManager.isTransaction Then
                Return DBObject.ExecuteScalar(Me.getDatabaseManager.getTransaction, CommandType.Text, sqlStr, paras)
            Else
                Return DBObject.ExecuteScalar(Me.getDatabaseManager.getConnection, CommandType.Text, sqlStr, paras)
            End If
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function COUNT_INCOME_2(ByVal iB_YYYYMMDD As Decimal, ByVal iE_YYYYMMDD As Decimal, ByVal sMB_MEMTYP As String, ByVal sMB_FEETYPE As String) As Decimal
        Try
            Dim sqlStr As String = String.Empty

            'sqlStr = "SELECT IFNULL(COUNT(DISTINCT(MB_MEMSEQ)),0) FROM MB_MEMREV "

            sqlStr = "SELECT IFNULL(COUNT(*),0) " & _
                     "  FROM MB_MEMREV A, MB_MEMBER B " & _
                     " WHERE     A.MB_MEMSEQ = B.MB_MEMSEQ " & _
                     "       AND A.MB_TX_DATE >= " & ProviderFactory.PositionPara & "B_MB_TX_DATE " & _
                     "       AND A.MB_TX_DATE <= " & ProviderFactory.PositionPara & "E_MB_TX_DATE " & _
                     "       AND A.MB_MEMTYP = " & ProviderFactory.PositionPara & "MB_MEMTYP " & _
                     "       AND A.MB_FEETYPE = " & ProviderFactory.PositionPara & "MB_FEETYPE "

            Dim paras(3) As IDbDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("B_MB_TX_DATE", iB_YYYYMMDD)

            paras(1) = ProviderFactory.CreateDataParameter("E_MB_TX_DATE", iE_YYYYMMDD)

            paras(2) = ProviderFactory.CreateDataParameter("MB_MEMTYP", sMB_MEMTYP)

            paras(3) = ProviderFactory.CreateDataParameter("MB_FEETYPE", sMB_FEETYPE)

            If Me.getDatabaseManager.isTransaction Then
                Return DBObject.ExecuteScalar(Me.getDatabaseManager.getTransaction, CommandType.Text, sqlStr, paras)
            Else
                Return DBObject.ExecuteScalar(Me.getDatabaseManager.getConnection, CommandType.Text, sqlStr, paras)
            End If
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function COUNT_INCOME_3(ByVal iB_YYYYMMDD As Decimal, ByVal iE_YYYYMMDD As Decimal, ByVal sMB_FEETYPE As String) As Decimal
        Try
            Dim sqlStr As String = String.Empty

            'sqlStr = "SELECT IFNULL(COUNT(DISTINCT(MB_MEMSEQ)),0) FROM MB_MEMREV "

            sqlStr = "SELECT IFNULL(COUNT(*),0) " & _
                     "  FROM MB_MEMREV A, MB_MEMBER B " & _
                     " WHERE     A.MB_MEMSEQ = B.MB_MEMSEQ " & _
                     "       AND A.MB_TX_DATE >= " & ProviderFactory.PositionPara & "B_MB_TX_DATE " & _
                     "       AND A.MB_TX_DATE <= " & ProviderFactory.PositionPara & "E_MB_TX_DATE " & _
                     "       AND A.MB_FEETYPE = " & ProviderFactory.PositionPara & "MB_FEETYPE "

            Dim paras(2) As IDbDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("B_MB_TX_DATE", iB_YYYYMMDD)

            paras(1) = ProviderFactory.CreateDataParameter("E_MB_TX_DATE", iE_YYYYMMDD)

            paras(2) = ProviderFactory.CreateDataParameter("MB_FEETYPE", sMB_FEETYPE)

            If Me.getDatabaseManager.isTransaction Then
                Return DBObject.ExecuteScalar(Me.getDatabaseManager.getTransaction, CommandType.Text, sqlStr, paras)
            Else
                Return DBObject.ExecuteScalar(Me.getDatabaseManager.getConnection, CommandType.Text, sqlStr, paras)
            End If
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function COUNT_INCOME_4(ByVal sMB_MEMTYP As String, ByVal sMB_FEETYPE As String) As Decimal
        Try
            Dim sqlStr As String = String.Empty

            'sqlStr = "SELECT IFNULL(COUNT(DISTINCT(MB_MEMSEQ)),0) FROM MB_MEMREV "

            sqlStr = "SELECT IFNULL(COUNT(*),0) " & _
                     "  FROM MB_MEMREV A, MB_MEMBER B " & _
                     " WHERE     A.MB_MEMSEQ = B.MB_MEMSEQ " & _
                     "       AND A.MB_MEMTYP = " & ProviderFactory.PositionPara & "MB_MEMTYP " & _
                     "       AND A.MB_FEETYPE = " & ProviderFactory.PositionPara & "MB_FEETYPE "

            Dim paras(1) As IDbDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("MB_MEMTYP", sMB_MEMTYP)

            paras(1) = ProviderFactory.CreateDataParameter("MB_FEETYPE", sMB_FEETYPE)

            If Me.getDatabaseManager.isTransaction Then
                Return DBObject.ExecuteScalar(Me.getDatabaseManager.getTransaction, CommandType.Text, sqlStr, paras)
            Else
                Return DBObject.ExecuteScalar(Me.getDatabaseManager.getConnection, CommandType.Text, sqlStr, paras)
            End If
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function COUNT_INCOME_5(ByVal sMB_MEMTYP As String) As Decimal
        Try
            Dim sqlStr As String = String.Empty

            'sqlStr = "SELECT IFNULL(COUNT(DISTINCT(MB_MEMSEQ)),0) FROM MB_MEMREV "

            sqlStr = "SELECT IFNULL(COUNT(*),0) " & _
                     "  FROM MB_MEMREV A, MB_MEMBER B " & _
                     " WHERE     A.MB_MEMSEQ = B.MB_MEMSEQ " & _
                     "       AND A.MB_MEMTYP = " & ProviderFactory.PositionPara & "MB_MEMTYP "

            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_MEMTYP", sMB_MEMTYP)

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

    Public Function COUNT_INCOME_6(ByVal sMB_FEETYPE As String) As Decimal
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT IFNULL(COUNT(*),0) " & _
                     "  FROM MB_MEMREV A, MB_MEMBER B " & _
                     " WHERE     A.MB_MEMSEQ = B.MB_MEMSEQ " & _
                     "       AND A.MB_FEETYPE = " & ProviderFactory.PositionPara & "MB_FEETYPE "

            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_FEETYPE", sMB_FEETYPE)

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

    Public Function COUNT_INCOME_7() As Decimal
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT IFNULL(COUNT(*),0) " & _
                     "  FROM MB_MEMREV A, MB_MEMBER B " & _
                     " WHERE     A.MB_MEMSEQ = B.MB_MEMSEQ "

            If Me.getDatabaseManager.isTransaction Then
                Return DBObject.ExecuteScalar(Me.getDatabaseManager.getTransaction, CommandType.Text, sqlStr)
            Else
                Return DBObject.ExecuteScalar(Me.getDatabaseManager.getConnection, CommandType.Text, sqlStr)
            End If
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' 根據"會員編號"取得種籽會員繳款資料
    ''' </summary>
    ''' <param name="iMB_MEMSEQ">MB_MEMSEQ會員編號</param>
    ''' <returns>Integer 取得筆數</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Function LoadBySEQ(ByVal iMB_MEMSEQ As Decimal) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT A.*, B.MB_AREA, B.MB_LEADER, B.MB_NAME " & _
                     "  FROM MB_MEMREV A, MB_MEMBER B" & _
                     " WHERE A.MB_MEMSEQ = 8 AND A.MB_ITEMID = 'A' AND A.MB_MEMTYP = '2' AND A.MB_MEMSEQ = B.MB_MEMSEQ " & _
                     " ORDER BY A.MB_MEMSEQ, A.MB_TX_DATE "

            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_MEMSEQ", iMB_MEMSEQ)

            Return Me.loadBySQLOnlyDs(sqlStr, para)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function
#End Region

#End Region

End Class
