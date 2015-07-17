Imports com.Azion.NET.VB

Public Class MB_NEWSList
    Inherits BosList

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.New("MB_NEWS", dbManager)
    End Sub

    Overrides Function newBos() As BosBase
        Dim bos As MB_NEWS = New MB_NEWS(MyBase.getDatabaseManager)
        Return bos
    End Function

#Region "Ted Function"
    ''' <summary>
    ''' 根據[CODEID]取得資料
    ''' </summary>
    ''' <param name="iCODEID">CODEID</param>
    ''' <returns>Integer 取得筆數</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Function LoadByCODEID_PREV(ByVal iCODEID As Integer, ByVal iSEQTIME As Decimal, ByVal iLIMIT As Integer) As Integer
        Try
            Dim sqlStr As String = String.Empty

            'CHGUID=TEMP為臨時連接MBSC Blog
            '將來需移除
            sqlStr = "SELECT * FROM MB_NEWS WHERE CODEID = " & ProviderFactory.PositionPara & "CODEID " & _
                     " AND SEQTIME <=" & ProviderFactory.PositionPara & "SEQTIME AND IFNULL(CHGUID,' ') <> 'TEMP' ORDER BY SEQTIME DESC LIMIT " & iLIMIT

            Dim paras(1) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("CODEID", iCODEID)
            paras(1) = ProviderFactory.CreateDataParameter("SEQTIME", iSEQTIME)

            Return Me.loadBySQLOnlyDs(sqlStr, paras)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function LoadByCODEID_PREV_FST(ByVal iCODEID As Integer, ByVal iSEQTIME As Decimal, ByVal iLIMIT As Integer) As Integer
        Try
            Dim sqlStr As String = String.Empty

            'CHGUID=TEMP為臨時連接MBSC Blog
            '將來需移除
            sqlStr = "SELECT * FROM MB_NEWS WHERE CODEID = " & ProviderFactory.PositionPara & "CODEID " & _
                     " AND SEQTIME <" & ProviderFactory.PositionPara & "SEQTIME AND IFNULL(CHGUID,' ') <> 'TEMP' ORDER BY SEQTIME DESC LIMIT " & iLIMIT

            Dim paras(1) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("CODEID", iCODEID)
            paras(1) = ProviderFactory.CreateDataParameter("SEQTIME", iSEQTIME)

            Return Me.loadBySQLOnlyDs(sqlStr, paras)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function LoadByCODEID_NEXT(ByVal iCODEID As Integer, ByVal iSEQTIME As Decimal, ByVal iLIMIT As Integer) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT * FROM MB_NEWS WHERE CODEID = " & ProviderFactory.PositionPara & "CODEID " & _
                     " AND SEQTIME >" & ProviderFactory.PositionPara & "SEQTIME AND IFNULL(CHGUID,' ') <> 'TEMP' ORDER BY SEQTIME DESC LIMIT " & iLIMIT

            Dim paras(1) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("CODEID", iCODEID)
            paras(1) = ProviderFactory.CreateDataParameter("SEQTIME", iSEQTIME)

            Return Me.loadBySQLOnlyDs(sqlStr, paras)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function LoadByCODEID_NEXT_FST(ByVal iCODEID As Integer, ByVal iSEQTIME As Decimal, ByVal iLIMIT As Integer) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT * FROM MB_NEWS WHERE CODEID = " & ProviderFactory.PositionPara & "CODEID " & _
                     " AND SEQTIME >" & ProviderFactory.PositionPara & "SEQTIME AND IFNULL(CHGUID,' ') <> 'TEMP' ORDER BY SEQTIME LIMIT " & iLIMIT

            Dim paras(1) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("CODEID", iCODEID)
            paras(1) = ProviderFactory.CreateDataParameter("SEQTIME", iSEQTIME)

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
    ''' 根據[CRETIME]取得資料
    ''' </summary>
    ''' <param name="iCRETIME">資料年月</param>
    ''' <param name="iLIMIT">筆數</param>
    ''' <returns>Integer 取得筆數</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Function LoadByTIME_PREV(ByVal iSEQTIME As Decimal, ByVal iLIMIT As Integer) As Integer
        Try
            Dim sqlStr As String = String.Empty

            'CHGUID=TEMP為臨時連接MBSC Blog
            '將來需移除
            sqlStr = "SELECT * " & _
                     "  FROM MB_NEWS " & _
                     " WHERE SEQTIME < " & ProviderFactory.PositionPara & "SEQTIME " & _
                     "      AND IFNULL(CHGUID,' ') <> 'TEMP' " & _
                     "ORDER BY SEQTIME DESC " & _
                     " LIMIT " & ProviderFactory.PositionPara & "LIMIT "

            Dim paras(1) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("SEQTIME", iSEQTIME)
            paras(1) = ProviderFactory.CreateDataParameter("LIMIT", iLIMIT)

            Return Me.loadBySQLOnlyDs(sqlStr, paras)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function LoadByTIME_PREV_FST(ByVal iSEQTIME As Decimal, ByVal iLIMIT As Integer) As Integer
        Try
            Dim sqlStr As String = String.Empty

            'CHGUID=TEMP為臨時連接MBSC Blog
            '將來需移除
            sqlStr = "SELECT * " & _
                     "  FROM MB_NEWS " & _
                     " WHERE SEQTIME < " & ProviderFactory.PositionPara & "SEQTIME " & _
                     "      AND IFNULL(CHGUID,' ') <> 'TEMP' " & _
                     "ORDER BY SEQTIME DESC " & _
                     " LIMIT " & ProviderFactory.PositionPara & "LIMIT "

            Dim paras(1) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("SEQTIME", iSEQTIME)
            paras(1) = ProviderFactory.CreateDataParameter("LIMIT", iLIMIT)

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
    ''' 根據[CRETIME]取得資料
    ''' </summary>
    ''' <param name="iCRETIME">資料年月</param>
    ''' <param name="iLIMIT">筆數</param>
    ''' <returns>Integer 取得筆數</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Function LoadByTIME_NEXT(ByVal iSEQTIME As Decimal, ByVal iLIMIT As Integer) As Integer
        Try
            Dim sqlStr As String = String.Empty

            'CHGUID=TEMP為臨時連接MBSC Blog
            '將來需移除
            sqlStr = "SELECT * " & _
                     "  FROM MB_NEWS " & _
                     " WHERE SEQTIME > " & ProviderFactory.PositionPara & "SEQTIME " & _
                     "      AND IFNULL(CHGUID,' ') <> 'TEMP' " & _
                     "ORDER BY SEQTIME DESC " & _
                     " LIMIT " & ProviderFactory.PositionPara & "LIMIT "

            Dim paras(1) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("SEQTIME", iSEQTIME)
            paras(1) = ProviderFactory.CreateDataParameter("LIMIT", iLIMIT)

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
    ''' 根據[CRETIME]取得資料
    ''' </summary>
    ''' <param name="iCRETIME">資料年月</param>
    ''' <param name="iLIMIT">筆數</param>
    ''' <returns>Integer 取得筆數</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Function LoadByTIME_NEXT_FST(ByVal iSEQTIME As Decimal, ByVal iLIMIT As Integer) As Integer
        Try
            Dim sqlStr As String = String.Empty

            'CHGUID=TEMP為臨時連接MBSC Blog
            '將來需移除
            sqlStr = "SELECT * " & _
                     "  FROM MB_NEWS " & _
                     " WHERE SEQTIME > " & ProviderFactory.PositionPara & "SEQTIME " & _
                     "      AND IFNULL(CHGUID,' ') <> 'TEMP' " & _
                     "ORDER BY SEQTIME " & _
                     " LIMIT " & ProviderFactory.PositionPara & "LIMIT "

            Dim paras(1) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("SEQTIME", iSEQTIME)
            paras(1) = ProviderFactory.CreateDataParameter("LIMIT", iLIMIT)

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
    ''' 根據"課程編號"取得資料
    ''' </summary>
    ''' <param name="iMB_SEQ">課程編號</param>
    ''' <returns>Integer 取得筆數</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Function LoadbyMB_SEQ(ByVal iMB_SEQ As Decimal) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT * FROM MB_NEWS " & _
                     "  WHERE MB_SEQ = " & ProviderFactory.PositionPara & "MB_SEQ "

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

    ''' <summary>
    ''' 取得最新20筆活動資料
    ''' </summary>
    ''' <returns>Integer 取得筆數</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Function LoadEvent() As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT SEQTIME,EVENT " & _
                     "  FROM MB_NEWS " & _
                     " WHERE EVENT IS NOT NULL " & _
                     "ORDER BY SEQTIME " & _
                     " LIMIT 20 "

            Return Me.loadBySQLOnlyDs(sqlStr)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function LoadByCRETIME(ByVal iCRETIME As Decimal) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT * FROM MB_NEWS WHERE CRETIME = " & ProviderFactory.PositionPara & "CRETIME "

            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("CRETIME", iCRETIME)

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


End Class
