Imports com.Azion.NET.VB

Public Class MB_NEWS
    Inherits BosBase

    Sub New()
        MyBase.New()
        Me.setPrimaryKeys()
    End Sub

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.new("MB_NEWS", dbManager)
    End Sub

#Region "Ted Function"
    ''' ----------------------------------------------------------------------------
    ''' <summary>
    ''' 根據[CRETIME]取得資料
    ''' </summary>
    ''' <param name="iCRETIME">CRETIME</param>
    ''' <returns>Boolean 是否有取得</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Function LoadByPK(ByVal iCRETIME As Decimal) As Boolean
        Try
            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("CRETIME", iCRETIME)

            Return Me.loadBySQL(para)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' 根據[CODEID]取得最大建立日期
    ''' </summary>
    ''' <param name="iCODEID">CODEID</param>
    ''' <returns>DECIMAL</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Function getMAX_SEQTIME(ByVal iCODEID As Decimal) As Decimal
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT MAX(SEQTIME) FROM MB_NEWS WHERE CODEID = " & ProviderFactory.PositionPara & "CODEID "

            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("CODEID", iCODEID)

            Dim objCRETIME As Object = Convert.DBNull

            If Me.getDatabaseManager.isTransaction Then
                objCRETIME = DBObject.ExecuteScalar(Me.getDatabaseManager.getTransaction, CommandType.Text, sqlStr, para)
            Else
                objCRETIME = DBObject.ExecuteScalar(Me.getDatabaseManager.getConnection, CommandType.Text, sqlStr, para)
            End If

            If IsNumeric(objCRETIME) Then
                Return CDec(objCRETIME)
            Else
                Dim D_NOW As Date = Now

                Dim iYYYYMMDD As Decimal = (D_NOW.Year * 10000) + (D_NOW.Month * 100) + D_NOW.Day

                Dim iHHMMSS As Decimal = (D_NOW.Hour * 10000) + (D_NOW.Minute * 100) + D_NOW.Second

                Return iYYYYMMDD * 1000000 + iHHMMSS
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
    ''' 取得最大建立日期
    ''' </summary>
    ''' <returns>DECIMAL</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Function getALL_MAX_SEQTIME() As Decimal
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT MAX(SEQTIME) FROM MB_NEWS"

            Dim objCRETIME As Object = Convert.DBNull

            If Me.getDatabaseManager.isTransaction Then
                objCRETIME = DBObject.ExecuteScalar(Me.getDatabaseManager.getTransaction, CommandType.Text, sqlStr)
            Else
                objCRETIME = DBObject.ExecuteScalar(Me.getDatabaseManager.getConnection, CommandType.Text, sqlStr)
            End If

            If IsNumeric(objCRETIME) Then
                Return CDec(objCRETIME)
            Else
                Dim D_NOW As Date = Now

                Dim iYYYYMMDD As Decimal = (D_NOW.Year * 10000) + (D_NOW.Month * 100) + D_NOW.Day

                Dim iHHMMSS As Decimal = (D_NOW.Hour * 10000) + (D_NOW.Minute * 100) + D_NOW.Second

                Return iYYYYMMDD * 1000000 + iHHMMSS
            End If
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function getMIN_SEQTIME() As Decimal
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT IFNULL(MIN(SEQTIME),0) FROM MB_NEWS "

            Dim iCNT As Object = Convert.DBNull

            If Me.getDatabaseManager.isTransaction Then
                iCNT = DBObject.ExecuteScalar(Me.getDatabaseManager.getTransaction.Connection, CommandType.Text, sqlStr)
            Else
                iCNT = DBObject.ExecuteScalar(Me.getDatabaseManager.getConnection, CommandType.Text, sqlStr)
            End If

            If IsNumeric(iCNT) Then
                Return iCNT
            Else
                Return 0
            End If
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function getMIN_SEQTIME_CODEID(ByVal iCODEID As Decimal) As Decimal
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT IFNULL(MIN(SEQTIME),0) FROM MB_NEWS WHERE CODEID = " & ProviderFactory.PositionPara & "CODEID "

            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("CODEID", iCODEID)

            Dim iCNT As Object = Convert.DBNull

            If Me.getDatabaseManager.isTransaction Then
                iCNT = DBObject.ExecuteScalar(Me.getDatabaseManager.getTransaction.Connection, CommandType.Text, sqlStr, para)
            Else
                iCNT = DBObject.ExecuteScalar(Me.getDatabaseManager.getConnection, CommandType.Text, sqlStr, para)
            End If

            If IsNumeric(iCNT) Then
                Return iCNT
            Else
                Return 0
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
