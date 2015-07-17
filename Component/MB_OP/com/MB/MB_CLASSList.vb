Imports com.Azion.NET.VB

Public Class MB_CLASSList
    Inherits BosList

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.New("MB_CLASS", dbManager)
    End Sub

    Overrides Function newBos() As BosBase
        Dim bos As MB_CLASS = New MB_CLASS(MyBase.getDatabaseManager)
        Return bos
    End Function

#Region "Grace Function"
    ''' <summary>
    ''' 根據類別,地點及開課時間取得資料
    ''' </summary>
    ''' <param name="sTypeNO">類別</param>
    ''' <param name="sPlace">地點</param>
    ''' <param name="sSdate">開課起日</param>
    ''' <param name="sEdate">開課迄日</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Function getClassList(Optional ByVal sTypeNo As String = "", Optional ByVal sPlace As String = "", Optional ByVal sSdate As String = "", Optional ByVal sEdate As String = "") As Integer
        Try
            Me.setSQLCondition(" ORDER BY A.MB_SEQ, B.MB_BATCH ")

            Dim sSQL As String = String.Empty
            Dim sINSQL As String = String.Empty

            Dim AL_Paras As New ArrayList
            If sTypeNo <> "" Then
                sINSQL &= " AND A.MB_TYPE_NO=" & ProviderFactory.PositionPara & "MB_TYPE_NO"
                AL_Paras.Add(ProviderFactory.CreateDataParameter("MB_TYPE_NO", sTypeNo))
            End If

            If sPlace <> "" Then
                sINSQL &= " AND B.MB_PLACE=" & ProviderFactory.PositionPara & "MB_PLACE"
                AL_Paras.Add(ProviderFactory.CreateDataParameter("MB_PLACE", sPlace))
            End If

            If sSdate <> "" Then
                sINSQL &= " AND B.MB_SDATE>=" & ProviderFactory.PositionPara & "MB_SDATE"
                AL_Paras.Add(ProviderFactory.CreateDataParameter("MB_SDATE", sSdate))
            End If

            If sEdate <> "" Then
                sINSQL &= " AND B.MB_EDATE<=" & ProviderFactory.PositionPara & "MB_EDATE"
                AL_Paras.Add(ProviderFactory.CreateDataParameter("MB_EDATE", sEdate))
            End If

            sSQL = "SELECT A.MB_TYPE, A.MB_TYPE_NO ,B.* FROM MB_CLASS_M A, MB_CLASS B "

            If AL_Paras.Count > 0 Then
                'sSQL = sSQL & " WHERE A.MB_SEQ = B.MB_SEQ " & Right(sINSQL, sINSQL.Length - 4)
                sSQL = sSQL & " WHERE A.MB_SEQ = B.MB_SEQ " & sINSQL
                Dim paras(AL_Paras.Count - 1) As IDbDataParameter
                For i As Integer = 0 To AL_Paras.Count - 1
                    paras(i) = AL_Paras.Item(i)
                Next
                Return MyBase.loadBySQL(sSQL, paras)
            Else
                sSQL = sSQL & " WHERE A.MB_SEQ = B.MB_SEQ "
                Return MyBase.loadBySQL(sSQL)
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

#Region "Ted Function"
    ''' <summary>
    ''' 根據"課程起日","課程訖日"查詢，取得MB_CLASS的資料
    ''' </summary>
    ''' <param name="D_MB_SDATE">課程起日</param>
    ''' <param name="D_MB_EDATE">課程訖日</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Function LoadBySEDate(ByVal D_MB_SDATE As Date, ByVal D_MB_EDATE As Date) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT A.* FROM MB_CLASS A WHERE A.MB_SDATE >= " & ProviderFactory.PositionPara & "MB_SDATE " & _
                     "  AND A.MB_EDATE <= " & ProviderFactory.PositionPara & "MB_EDATE " & _
                     "      ORDER BY MB_EDATE "

            Dim paras(1) As IDbDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("MB_SDATE", D_MB_SDATE)

            paras(1) = ProviderFactory.CreateDataParameter("MB_EDATE", D_MB_EDATE)

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
    ''' 根據"課程序號"查詢，取得MB_CLASS的資料
    ''' </summary>
    ''' <param name="iMB_SEQ">課程序號</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Function LoadByMB_SEQ(ByVal iMB_SEQ As Decimal) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT * FROM MB_CLASS WHERE MB_SEQ = " & ProviderFactory.PositionPara & "MB_SEQ "

            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_SEQ", iMB_SEQ)

            Return Me.loadBySQL(sqlStr, para)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

#Region "可受理報名課程"
    ''' <summary>
    ''' 取得MB_CLASS今天在報名起訖日的資料
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Function LoadAPLY() As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT *" & _
                     "  FROM MB_CLASS" & _
                     " WHERE SYSDATE() >= MB_SAPLY" & _
                     "   AND SYSDATE() < MB_EAPLY"

            Return Me.loadBySQLOnlyDs(sqlStr)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' 取得MB_CLASS今天在報名起訖日的資料
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Function LoadAPLY_G() As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT *" & _
                     "  FROM MB_CLASS" & _
                     " WHERE SYSDATE() >= MB_SAPLY" & _
                     "   AND SYSDATE() < MB_EAPLY" & _
                     "      GROUP BY MB_SEQ ORDER BY MB_SAPLY "

            Return Me.loadBySQLOnlyDs(sqlStr)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function


    ''' <summary>
    ''' 根據傳入的課程地點取得MB_CLASS今天在報名起訖日的資料
    ''' </summary>
    ''' <param name="sMB_PLACE">課程地點</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Function LoadAPLY_MB_PLACE(ByVal sMB_PLACE As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT *" & _
                     "  FROM MB_CLASS" & _
                     " WHERE SYSDATE() >= MB_SAPLY" & _
                     "   AND SYSDATE() < MB_EAPLY" & _
                     "   AND MB_PLACE = " & ProviderFactory.PositionPara & "MB_PLACE "

            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_PLACE", sMB_PLACE)

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
    ''' 根據傳入的"課程地點","課程起日","課程起日之查詢訖日"取得MB_CLASS今天在報名起訖日的資料
    ''' </summary>
    ''' <param name="sMB_PLACE">課程地點</param>
    ''' <param name="D_S_SDATE">課程起日</param>
    ''' <param name="D_E_SDATE">課程起日之查詢訖日</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Function LoadAPLY_PLACE_SDATE(ByVal sMB_PLACE As String, ByVal D_S_SDATE As Date, ByVal D_E_SDATE As Date) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT *" & _
                     "  FROM MB_CLASS" & _
                     " WHERE SYSDATE() >= MB_SAPLY" & _
                     "   AND SYSDATE() < MB_EAPLY" & _
                     "   AND MB_PLACE = " & ProviderFactory.PositionPara & "MB_PLACE " & _
                     "   AND MB_SDATE BETWEEN " & ProviderFactory.PositionPara & "S_MB_SDATE " & _
                     "   AND " & ProviderFactory.PositionPara & "E_MB_SDATE "

            Dim paras(2) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("MB_PLACE", sMB_PLACE)
            paras(1) = ProviderFactory.CreateDataParameter("S_MB_SDATE", D_S_SDATE)
            paras(2) = ProviderFactory.CreateDataParameter("E_MB_SDATE", D_E_SDATE)

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
    ''' 根據傳入的課程起日取得MB_CLASS今天在報名起訖日的資料
    ''' </summary>
    ''' <param name="D_S_SDATE">課程起日</param>
    ''' <param name="D_E_SDATE">課程起日之查詢訖日</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Function LoadAPLY_SDATE(ByVal D_S_SDATE As Date, ByVal D_E_SDATE As Date) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT *" & _
                     "  FROM MB_CLASS" & _
                     " WHERE SYSDATE() >= MB_SAPLY" & _
                     "   AND SYSDATE() < MB_EAPLY" & _
                     "   AND MB_SDATE BETWEEN " & ProviderFactory.PositionPara & "S_MB_SDATE " & _
                     "   AND " & ProviderFactory.PositionPara & "E_MB_SDATE "

            Dim paras(1) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("S_MB_SDATE", D_S_SDATE)
            paras(1) = ProviderFactory.CreateDataParameter("E_MB_SDATE", D_E_SDATE)

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

#Region "預告課程"
    ''' <summary>
    ''' 取得MB_CLASS今天在報名起訖日的資料
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Function LoadAPLY_2() As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT *" & _
                     "  FROM MB_CLASS" & _
                     " WHERE MB_SAPLY > SYSDATE() " 

            Return Me.loadBySQLOnlyDs(sqlStr)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' 取得MB_CLASS今天在報名起訖日的資料
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Function LoadAPLY_2_G() As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT *" & _
                     "  FROM MB_CLASS" & _
                     " WHERE MB_SAPLY > SYSDATE() " & _
                     "  GROUP BY MB_SEQ ORDER BY MB_SAPLY "

            Return Me.loadBySQLOnlyDs(sqlStr)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' 根據傳入的課程地點取得MB_CLASS今天在報名起訖日的資料
    ''' </summary>
    ''' <param name="sMB_PLACE">課程地點</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Function LoadAPLY_MB_PLACE_2(ByVal sMB_PLACE As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT *" & _
                     "  FROM MB_CLASS" & _
                     " WHERE MB_SAPLY > SYSDATE()" & _
                     "   AND MB_PLACE = " & ProviderFactory.PositionPara & "MB_PLACE "

            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_PLACE", sMB_PLACE)

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
    ''' 根據傳入的"課程地點","課程起日","課程起日之查詢訖日"取得MB_CLASS今天在報名起訖日的資料
    ''' </summary>
    ''' <param name="sMB_PLACE">課程地點</param>
    ''' <param name="D_S_SDATE">課程起日</param>
    ''' <param name="D_E_SDATE">課程起日之查詢訖日</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Function LoadAPLY_PLACE_SDATE_2(ByVal sMB_PLACE As String, ByVal D_S_SDATE As Date, ByVal D_E_SDATE As Date) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT *" & _
                     "  FROM MB_CLASS" & _
                     " WHERE MB_SAPLY > SYSDATE()" & _
                     "   AND MB_PLACE = " & ProviderFactory.PositionPara & "MB_PLACE " & _
                     "   AND MB_SDATE BETWEEN " & ProviderFactory.PositionPara & "S_MB_SDATE " & _
                     "   AND " & ProviderFactory.PositionPara & "E_MB_SDATE "

            Dim paras(2) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("MB_PLACE", sMB_PLACE)
            paras(1) = ProviderFactory.CreateDataParameter("S_MB_SDATE", D_S_SDATE)
            paras(2) = ProviderFactory.CreateDataParameter("E_MB_SDATE", D_E_SDATE)

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
    ''' 根據傳入的課程起日取得MB_CLASS今天在報名起訖日的資料
    ''' </summary>
    ''' <param name="D_S_SDATE">課程起日</param>
    ''' <param name="D_E_SDATE">課程起日之查詢訖日</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Function LoadAPLY_SDATE_2(ByVal D_S_SDATE As Date, ByVal D_E_SDATE As Date) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT *" & _
                     "  FROM MB_CLASS" & _
                     " WHERE MB_SAPLY > SYSDATE()" & _
                     "   AND MB_SDATE BETWEEN " & ProviderFactory.PositionPara & "S_MB_SDATE " & _
                     "   AND " & ProviderFactory.PositionPara & "E_MB_SDATE "

            Dim paras(1) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("S_MB_SDATE", D_S_SDATE)
            paras(1) = ProviderFactory.CreateDataParameter("E_MB_SDATE", D_E_SDATE)

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

#Region "進行中課程"
    ''' <summary>
    ''' 取得MB_CLASS今天在報名起訖日的資料
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Function LoadAPLY_4() As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT *" & _
                     "  FROM MB_CLASS" & _
                     " WHERE SYSDATE() BETWEEN MB_SDATE AND MB_EDATE "

            Return Me.loadBySQLOnlyDs(sqlStr)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function LoadAPLY_5() As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT *" & _
                     "  FROM MB_CLASS" & _
                     " WHERE SYSDATE() < MB_SDATE AND SYSDATE() > MB_EAPLY "

            Return Me.loadBySQLOnlyDs(sqlStr)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' 取得MB_CLASS今天在報名起訖日的資料
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Function LoadAPLY_4_G() As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT *" & _
                     "  FROM MB_CLASS" & _
                     " WHERE SYSDATE() BETWEEN MB_SDATE AND MB_EDATE " & _
                     "  GROUP BY MB_SEQ ORDER BY MB_SDATE "

            Return Me.loadBySQLOnlyDs(sqlStr)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' 根據傳入的課程地點取得MB_CLASS今天在報名起訖日的資料
    ''' </summary>
    ''' <param name="sMB_PLACE">課程地點</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Function LoadAPLY_MB_PLACE_4(ByVal sMB_PLACE As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT *" & _
                     "  FROM MB_CLASS" & _
                     " WHERE (SYSDATE() BETWEEN MB_SDATE AND MB_EDATE) " & _
                     "   AND MB_PLACE = " & ProviderFactory.PositionPara & "MB_PLACE "

            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_PLACE", sMB_PLACE)

            Return Me.loadBySQLOnlyDs(sqlStr, para)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function LoadAPLY_MB_PLACE_5(ByVal sMB_PLACE As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT *" & _
                     "  FROM MB_CLASS" & _
                     " WHERE (SYSDATE() < MB_SDATE AND SYSDATE() > MB_EAPLY) " & _
                     "   AND MB_PLACE = " & ProviderFactory.PositionPara & "MB_PLACE "

            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_PLACE", sMB_PLACE)

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
    ''' 根據傳入的"課程地點","課程起日","課程起日之查詢訖日"取得MB_CLASS今天在報名起訖日的資料
    ''' </summary>
    ''' <param name="sMB_PLACE">課程地點</param>
    ''' <param name="D_S_SDATE">課程起日</param>
    ''' <param name="D_E_SDATE">課程起日之查詢訖日</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Function LoadAPLY_PLACE_SDATE_4(ByVal sMB_PLACE As String, ByVal D_S_SDATE As Date, ByVal D_E_SDATE As Date) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT *" & _
                     "  FROM MB_CLASS" & _
                     " WHERE (SYSDATE() BETWEEN MB_SDATE AND MB_EDATE) " & _
                     "   AND MB_PLACE = " & ProviderFactory.PositionPara & "MB_PLACE " & _
                     "   AND MB_SDATE BETWEEN " & ProviderFactory.PositionPara & "S_MB_SDATE " & _
                     "   AND " & ProviderFactory.PositionPara & "E_MB_SDATE "

            Dim paras(2) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("MB_PLACE", sMB_PLACE)
            paras(1) = ProviderFactory.CreateDataParameter("S_MB_SDATE", D_S_SDATE)
            paras(2) = ProviderFactory.CreateDataParameter("E_MB_SDATE", D_E_SDATE)

            Return Me.loadBySQLOnlyDs(sqlStr, paras)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function LoadAPLY_PLACE_SDATE_5(ByVal sMB_PLACE As String, ByVal D_S_SDATE As Date, ByVal D_E_SDATE As Date) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT *" & _
                     "  FROM MB_CLASS" & _
                     " WHERE (SYSDATE() < MB_SDATE AND SYSDATE() > MB_EAPLY) " & _
                     "   AND MB_PLACE = " & ProviderFactory.PositionPara & "MB_PLACE " & _
                     "   AND MB_SDATE BETWEEN " & ProviderFactory.PositionPara & "S_MB_SDATE " & _
                     "   AND " & ProviderFactory.PositionPara & "E_MB_SDATE "

            Dim paras(2) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("MB_PLACE", sMB_PLACE)
            paras(1) = ProviderFactory.CreateDataParameter("S_MB_SDATE", D_S_SDATE)
            paras(2) = ProviderFactory.CreateDataParameter("E_MB_SDATE", D_E_SDATE)

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
    ''' 根據傳入的課程起日取得MB_CLASS今天在報名起訖日的資料
    ''' </summary>
    ''' <param name="D_S_SDATE">課程起日</param>
    ''' <param name="D_E_SDATE">課程起日之查詢訖日</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Function LoadAPLY_SDATE_4(ByVal D_S_SDATE As Date, ByVal D_E_SDATE As Date) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT *" & _
                     "  FROM MB_CLASS" & _
                     " WHERE (SYSDATE() BETWEEN MB_SDATE AND MB_EDATE) " & _
                     "   AND MB_SDATE BETWEEN " & ProviderFactory.PositionPara & "S_MB_SDATE " & _
                     "   AND " & ProviderFactory.PositionPara & "E_MB_SDATE "

            Dim paras(1) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("S_MB_SDATE", D_S_SDATE)
            paras(1) = ProviderFactory.CreateDataParameter("E_MB_SDATE", D_E_SDATE)

            Return Me.loadBySQLOnlyDs(sqlStr, paras)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function LoadAPLY_SDATE_5(ByVal D_S_SDATE As Date, ByVal D_E_SDATE As Date) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT *" & _
                     "  FROM MB_CLASS" & _
                     " WHERE (SYSDATE() > MB_SDATE AND SYSDATE() < MB_EAPLY) " & _
                     "   AND MB_SDATE BETWEEN " & ProviderFactory.PositionPara & "S_MB_SDATE " & _
                     "   AND " & ProviderFactory.PositionPara & "E_MB_SDATE "

            Dim paras(1) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("S_MB_SDATE", D_S_SDATE)
            paras(1) = ProviderFactory.CreateDataParameter("E_MB_SDATE", D_E_SDATE)

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

#Region "關閉課程(已完成之課程)"
    ''' <summary>
    ''' 取得MB_CLASS今天在報名起訖日的資料
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Function LoadAPLY_3() As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT *" & _
                     "  FROM MB_CLASS" & _
                     " WHERE MB_EDATE <= SYSDATE() "

            Return Me.loadBySQLOnlyDs(sqlStr)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' 根據傳入的課程地點取得MB_CLASS今天在報名起訖日的資料
    ''' </summary>
    ''' <param name="sMB_PLACE">課程地點</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Function LoadAPLY_MB_PLACE_3(ByVal sMB_PLACE As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT *" & _
                     "  FROM MB_CLASS" & _
                     " WHERE MB_EDATE <= SYSDATE()" & _
                     "   AND MB_PLACE = " & ProviderFactory.PositionPara & "MB_PLACE "

            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_PLACE", sMB_PLACE)

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
    ''' 根據傳入的"課程地點","課程起日","課程起日之查詢訖日"取得MB_CLASS今天在報名起訖日的資料
    ''' </summary>
    ''' <param name="sMB_PLACE">課程地點</param>
    ''' <param name="D_S_SDATE">課程起日</param>
    ''' <param name="D_E_SDATE">課程起日之查詢訖日</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Function LoadAPLY_PLACE_SDATE_3(ByVal sMB_PLACE As String, ByVal D_S_SDATE As Date, ByVal D_E_SDATE As Date) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT *" & _
                     "  FROM MB_CLASS" & _
                     " WHERE MB_SDATE <= SYSDATE()" & _
                     "   AND MB_PLACE = " & ProviderFactory.PositionPara & "MB_PLACE " & _
                     "   AND MB_SDATE BETWEEN " & ProviderFactory.PositionPara & "S_MB_SDATE " & _
                     "   AND " & ProviderFactory.PositionPara & "E_MB_SDATE "

            Dim paras(2) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("MB_PLACE", sMB_PLACE)
            paras(1) = ProviderFactory.CreateDataParameter("S_MB_SDATE", D_S_SDATE)
            paras(2) = ProviderFactory.CreateDataParameter("E_MB_SDATE", D_E_SDATE)

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
    ''' 根據傳入的課程起日取得MB_CLASS今天在報名起訖日的資料
    ''' </summary>
    ''' <param name="D_S_SDATE">課程起日</param>
    ''' <param name="D_E_SDATE">課程起日之查詢訖日</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Function LoadAPLY_SDATE_3(ByVal D_S_SDATE As Date, ByVal D_E_SDATE As Date) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT *" & _
                     "  FROM MB_CLASS" & _
                     " WHERE MB_SDATE <= SYSDATE()" & _
                     "   AND MB_SDATE BETWEEN " & ProviderFactory.PositionPara & "S_MB_SDATE " & _
                     "   AND " & ProviderFactory.PositionPara & "E_MB_SDATE "

            Dim paras(1) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("S_MB_SDATE", D_S_SDATE)
            paras(1) = ProviderFactory.CreateDataParameter("E_MB_SDATE", D_E_SDATE)

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

#End Region

End Class
