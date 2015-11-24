Imports com.Azion.NET.VB
Public Class AP_POSTLIST
    Inherits BosList
    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.New("AP_POST", dbManager)
    End Sub

    Overrides Function newBos() As BosBase
        Dim bos As AP_POST = New AP_POST(MyBase.getDatabaseManager)
        Return bos
    End Function
#Region "CHRIS"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 初始帶出縣市下拉
    ''' </summary>
    ''' <returns>Integer 資料筆數</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[chris]	2011/8/30	Created
    ''' </history>
    Function loadDataByCity() As Integer

        Dim sSQL As String

        sSQL = "select distinct CITY from AP_POST order by CITY "
        Try

            Return MyBase.loadBySQL(sSQL)

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
    ''' 依下拉縣市帶出鄉鎮市區下拉
    ''' </summary>
    ''' <param name="sCITY">縣市名稱</param>
    ''' <returns>Integer 資料筆數</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[chris]	2011/8/30	Created
    ''' </history>
    Function loadAREAByCITY(ByVal sCITY As String) As Integer
        Dim sSQL As String = String.Empty

        Try
            If Utility.isValidateData(sCITY) Then
                sSQL = "Select distinct AREA from AP_POST Where CITY =" & ProviderFactory.PositionPara & "CITY order by AREA"
                Dim paras(0) As System.Data.IDbDataParameter

                paras(0) = ProviderFactory.CreateDataParameter("CITY", sCITY.Trim)
                Return MyBase.loadBySQL(sSQL, paras)
            Else
                Return MyBase.loadBySQL(sSQL)
            End If

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
    ''' 依下拉縣市,鄉鎮市區帶出郵遞區號下拉
    ''' </summary>
    ''' <returns>Integer 資料筆數</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[chris]	2011/8/30	Created
    ''' </history>
    Function loadZIPByAREA(ByVal sAREA As String) As Integer
        Dim sSQL As String = String.Empty

        Try
            If Utility.isValidateData(sAREA) Then
                sSQL = "Select distinct ZIP from AP_POST where AREA =" & ProviderFactory.PositionPara & "AREA order by ZIP"
                Dim paras(0) As System.Data.IDbDataParameter

                paras(0) = ProviderFactory.CreateDataParameter("AREA", sAREA.Trim)
                Return MyBase.loadBySQL(sSQL, paras)
            Else
                Return MyBase.loadBySQL(sSQL)
            End If
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
    ''' 依下拉縣市,鄉鎮市區,郵遞區號撈出AP_POST資料
    ''' </summary>
    ''' <param name="sCITY">縣市名稱</param>
    ''' <param name="sAREA">鄉鎮市區名稱</param>
    ''' <returns>Integer 資料筆數</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[chris]	2011/8/30	Created
    ''' </history>
    Function loadDataByCITYAREA(ByVal sCITY As String, ByVal sAREA As String) As Integer
        Dim sSQL As String
        Try
            If Utility.isValidateData(sCITY) Then
                sSQL = "select * from AP_POST where CITY=" & ProviderFactory.PositionPara & "CITY  And AREA =" & ProviderFactory.PositionPara & "AREA order by CITY,AREA,ROAD,MEMO1"
                If Utility.isValidateData(sAREA) Then
                    Dim paras(1) As System.Data.IDbDataParameter
                    paras(0) = ProviderFactory.CreateDataParameter("CITY", sCITY.Trim)
                    paras(1) = ProviderFactory.CreateDataParameter("AREA", sAREA.Trim)
                    Return MyBase.loadBySQL(sSQL, paras)
                    'Return MyBase.ShowSelectSQL(sSQL, paras)
                End If
            End If
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function


#End Region


#Region "Nick Function"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 根據城市中文名,及鄉鎮中文名取得
    ''' 取得 AP_POST 的 record
    ''' </summary>
    ''' <returns>integer</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Nick]	2011/9/05	Created
    ''' </history>
    Public Function LoadByCityAndArea(ByVal sCITY As String, ByVal sAREA As String)
        Try
            Dim sqlStr As String = ""

            sqlStr = "select * from ap_post where trim(nvl(CITY,' '))=" & ProviderFactory.PositionPara & "CITY and " & _
                     " trim(nvl(AREA,' '))=" & ProviderFactory.PositionPara & "AREA "

            Dim paras(1) As IDbDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("CITY", sCITY)

            paras(1) = ProviderFactory.CreateDataParameter("AREA", sAREA)

            Return Me.loadBySQLOnlyDs(sqlStr, paras)
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 根據城市中文名取得
    ''' 取得 AP_POST 的 record
    ''' </summary>
    ''' <returns>integer</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Nick]	2011/9/05	Created
    ''' </history>
    Public Function LoadByCity(ByVal sCITY As String)
        Try
            Dim sqlStr As String = ""

            sqlStr = "select * from ap_post where trim(nvl(city,' '))=" & ProviderFactory.PositionPara & "city "

            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("city", sCITY)

            Return Me.loadBySQL(sqlStr, para)
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 依下拉縣市,鄉鎮市區,郵遞區號組成撈資料的SQL子句
    ''' </summary>
    ''' <param name="sCITY">縣市名稱</param>
    ''' <param name="sAREA">鄉鎮市區名稱</param>
    ''' <returns>String SQL子句</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Nick]	2012/1/13	Created
    ''' </history>
    Function GetSQLByCITYAREA(ByVal sCITY As String, ByVal sAREA As String) As String
        Dim sSQL As String
        Try
            If Utility.isValidateData(sCITY) Then
                sSQL = "select * from AP_POST where CITY=" & ProviderFactory.PositionPara & "CITY  And AREA =" & ProviderFactory.PositionPara & "AREA " 'order by CITY,AREA,ROAD,MEMO1"
                If Utility.isValidateData(sAREA) Then
                    Dim paras(1) As System.Data.IDbDataParameter
                    paras(0) = ProviderFactory.CreateDataParameter("CITY", sCITY.Trim)
                    paras(1) = ProviderFactory.CreateDataParameter("AREA", sAREA.Trim)
                    'Return MyBase.loadBySQL(sSQL, paras)
                    Return MyBase.ShowSelectSQL(sSQL, paras)
                End If
            End If
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

#Region "Jane Function"
    ''' <summary>
    ''' Load Data by Primary Key (CITY, AREA, ROAD)
    ''' </summary>
    ''' <param name="sCITY"></param>
    ''' <param name="sAREA"></param>
    ''' <param name="sROAD"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function LoadDataByPK(ByVal sCITY As String, ByVal sAREA As String, ByVal sROAD As String) As Integer
        Try
            If Utility.isValidateData(sCITY) AndAlso Utility.isValidateData(sAREA) AndAlso Utility.isValidateData(sROAD) Then
                Dim paras(2) As System.Data.IDbDataParameter
                paras(0) = ProviderFactory.CreateDataParameter("CITY", sCITY)
                paras(1) = ProviderFactory.CreateDataParameter("AREA", sAREA)
                paras(2) = ProviderFactory.CreateDataParameter("ROAD", sROAD)

                Return MyBase.loadBySQL(paras)
            End If
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 以Like方式查詢路名
    ''' </summary>
    ''' <param name="sCity"></param>
    ''' <param name="sArea"></param>
    ''' <param name="sRoad"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function QryRoadName(ByVal sCity As String, ByVal sArea As String, ByVal sRoad As String) As Integer
        Try
            Dim sqlStr As String = ""

            sqlStr = "select * from AP_POST " & _
                     " where CITY=" & ProviderFactory.PositionPara & "CITY" & _
                     " and AREA =" & ProviderFactory.PositionPara & "AREA" & _
                     " and ROAD like " & ProviderFactory.PositionPara & "ROAD"

            If Utility.isValidateData(sCity) AndAlso Utility.isValidateData(sArea) AndAlso Utility.isValidateData(sRoad) Then
                Dim paras(2) As System.Data.IDbDataParameter
                paras(0) = ProviderFactory.CreateDataParameter("CITY", sCity)
                paras(1) = ProviderFactory.CreateDataParameter("AREA", sArea)
                paras(2) = ProviderFactory.CreateDataParameter("ROAD", "%" & sRoad & "%")

                Return MyBase.loadBySQL(sqlStr, paras)
            End If
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

#Region "濟南昱勝添加"
#Region "Avril Function"
    ''' <summary>
    ''' 依縣市中文名及鄉鎮中文名查詢郵政區號
    ''' </summary>
    ''' <param name="sCITY">縣市中文名</param>
    ''' <param name="sAREA">鄉鎮中文名</param>
    ''' <returns>Boolean</returns>
    ''' <remarks>
    ''' [Avril] 2012/01/09	Created
    ''' </remarks>
    Function LoadDataByCityAndArea(ByVal sCITY As String, ByVal sAREA As String) As Boolean

        Try
            Dim sqlStr As String = ""

            sqlStr = "SELECT * FROM AP_POST WHERE CITY=" & ProviderFactory.PositionPara & "CITY AND " & _
                     " AREA =" & ProviderFactory.PositionPara & "AREA AND LEN(ZIP)=3"

            Dim paras(1) As IDbDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("CITY", sCITY)

            paras(1) = ProviderFactory.CreateDataParameter("AREA", sAREA)

            Return Me.loadBySQLOnlyDs(sqlStr, paras)
        Catch ex As Exception
            Throw ex
        End Try
    End Function

#End Region
#End Region
End Class
