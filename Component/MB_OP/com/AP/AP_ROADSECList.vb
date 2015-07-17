Imports com.Azion.NET.VB

Public Class AP_ROADSECList
    Inherits BosList

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.New("AP_ROADSEC", dbManager)
    End Sub

    Overrides Function newBos() As BosBase
        Dim bos As AP_ROADSEC = New AP_ROADSEC(MyBase.getDatabaseManager)
        'If (Me.getArrayPrimaryKeys.Count = 0) Then
        '    Me.setPrimaryKeys(bos.getArrayPrimaryKeys())
        'End If
        Return bos
    End Function

#Region "Thera Function"
    ''' <summary>
    ''' 取出縣市資料
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function loadAllCity() As Integer
        Try
            Dim sSQL As String = " select DISTINCT CITY,CITY_ID,CITY_CD   " & _
                                          "  from  AP_ROADSEC   " & _
                                          "  where OPTYPE <>'D'  " & _
                                          " order by CITY_CD  "


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
    ''' 傳入CITY_ID 資料, 取得縣市別資料
    ''' </summary>
    ''' <param name="sCityId">縣市代碼</param>
    ''' <returns>AREA_ID, AREA</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' </history>
    Function loadByCityId(ByVal sCityId As String) As Integer
        Try
            Dim sSQL As String = "select   DISTINCT AREA_ID, AREA   " & _
                                          "  from  AP_ROADSEC   " & _
                                          "  where CITY_ID=" & ProviderFactory.PositionPara & "CITY_ID" & _
                                          "   and   OPTYPE <>'D'  " & _
                                          " order by AREA_ID  "

            Dim paras(0) As System.Data.IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("CITY_ID", sCityId)
            Return MyBase.loadBySQL(sSQL, paras)

        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 傳入縣市/鄉鎮市區代碼，取得段小段資料
    ''' </summary>
    ''' <param name="sCityId">縣市代碼</param>
    ''' <param name="sAreaId">鄉鎮市區代碼</param>
    ''' <returns>MAINSEC, SEC_ID</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' </history>
    Function loadByAreaId(ByVal sCityId As String, ByVal sAreaId As String) As Integer
        Try
            Dim sSQL As String = "select   DISTINCT MAINSEC, SEC_ID , AREA_CENTER_ID   " & _
                                          "  from  AP_ROADSEC   " & _
                                          "  where CITY_ID=" & ProviderFactory.PositionPara & "CITY_ID" & _
                                          "   and   AREA_ID = " & ProviderFactory.PositionPara & "AREA_ID" & _
                                          "   and   OPTYPE <>'D'  " & _
                                          " order by SEC_ID  "

            Dim paras(1) As System.Data.IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("CITYID", sCityId)
            paras(1) = ProviderFactory.CreateDataParameter("AREAID", sAreaId)
            Return MyBase.loadBySQL(sSQL, paras)

        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 『沿用擔保品』
    ''' 傳入縣市代碼/鄉鎮市區代碼，查詢縣市代碼(地所)/鄉鎮市區代碼/中心代碼(地所)
    ''' </summary>
    ''' <param name="sCity"></param>
    ''' <param name="sArea"></param>
    ''' <returns></returns>
    ''' 20110817 thera add
    ''' <remarks></remarks>
    Public Function loadCityAreaByID(ByVal sCity As String, ByVal sArea As String) As Integer
        Try
            Dim sSQL As String = String.Empty

            sSQL = " select" & _
                   "  distinct CITY_ID, AREA_ID, AREA_CENTER_ID" & _
                   " from AP_ROADSEC" & _
                   " where CITY_CD=" & ProviderFactory.PositionPara & "CITY_CD" & _
                   " and  AREA_ID=" & ProviderFactory.PositionPara & "AREA_ID" & _
                   " and OPTYPE<>'D'"

            Dim paras(1) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("CITY_CD", sCity)
            paras(1) = ProviderFactory.CreateDataParameter("AREA_ID", sArea)

            Return MyBase.loadBySQL(sSQL, paras)

        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

#End Region

#Region "nick function"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 傳入CITY_CD 資料, 取得縣市別資料
    ''' </summary>
    ''' <param name="sCity">縣市</param>
    ''' <returns>AREA_ID, AREA</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' </history>
    Function loadByCity(ByVal sCity As String) As Integer
        Try
            Dim sSQL As String = "select   DISTINCT AREA_ID, AREA   " & _
                                          "  from  AP_ROADSEC   " & _
                                          "  where CITY=" & ProviderFactory.PositionPara & "CITY" & _
                                          " order by AREA_ID  "

            Dim paras(0) As System.Data.IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("CITY", sCity)
            Return MyBase.loadBySQL(sSQL, paras)

        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 取出縣市資料
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function loadCity() As Integer
        Try
            Dim sSQL As String = " select DISTINCT CITY,CITY_ID,CITY_CD  " & _
                                          "  from  AP_ROADSEC   " & _
                                          " order by CITY_CD  "


            Return MyBase.loadBySQL(sSQL)

        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

#End Region

#Region "Joe Function"
    Public Function getIdByName(ByVal sCity As String, ByVal sArea As String) As Integer
        Try
            Dim sSQL As String = String.Empty
            sSQL = " select" & _
                   "  distinct CITY_ID, AREA_ID, AREA_CENTER_ID" & _
                   " from AP_ROADSEC" & _
                   " where CITY=" & ProviderFactory.PositionPara & "CITY" & _
                   " and AREA=" & ProviderFactory.PositionPara & "AREA" & _
                   " and OPTYPE<>'D'"

            Dim paras(1) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("CITY", sCity)
            paras(1) = ProviderFactory.CreateDataParameter("AREA", sArea)

            Return MyBase.loadBySQL(sSQL, paras)
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

#Region "scott function"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 傳入CITY,AREA 資料, 取得資料
    ''' </summary>
    ''' <param name="sCity">縣市</param>
    ''' <param name="sAREA">鄉鎮市區</param>
    ''' <returns> </returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' </history>
    Function getdatabyareaid(ByVal sCity As String, ByVal sAREA As String) As Integer
        Try
            Dim sSQL As String = "select   DISTINCT AREA_ID, AREA ,AREA_CENTER ,AREA_CENTER_ID ,MAINSEC ,SEC_ID ,OPTYPE   " & _
                                          "  from  AP_ROADSEC   " & _
                                          "  where CITY=" & ProviderFactory.PositionPara & "CITY" & _
                                          " and  AREA=" & ProviderFactory.PositionPara & "AREA" & _
                                          " order by AREA_ID  "

            Dim paras(1) As System.Data.IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("CITY", sCity)
            paras(1) = ProviderFactory.CreateDataParameter("AREA", sAREA)
            Return MyBase.loadBySQL(sSQL, paras)

        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

#End Region

#Region "Ted Function"
    ''' <summary>
    ''' 取出縣市資料OPTYPE!=D
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function loadCityNOT_D(Optional ByVal bSORT As Boolean = True) As Integer
        Try
            Dim sSQL As String = " select DISTINCT CITY_ID, CITY  " & _
                                          "  from  AP_ROADSEC WHERE IFNULL(OPTYPE,' ')<>'D' "

            If bSORT Then
                sSQL &= " order by CITY_CD  "
            End If

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
    ''' 傳入CITY_CD 資料, 取得縣市別資料，OPTYPE!=D
    ''' </summary>
    ''' <param name="sCity">縣市</param>
    ''' <returns>AREA_ID, AREA</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' </history>
    Function loadByCityNOT_D(ByVal sCity As String) As Integer
        Try
            Dim sSQL As String = "select   DISTINCT AREA_ID, AREA   " & _
                                          "  from  AP_ROADSEC   " & _
                                          "  where CITY=" & ProviderFactory.PositionPara & "CITY AND IFNULL(OPTYPE,' ')<>'D' " & _
                                          " order by AREA_ID  "

            Dim paras(0) As System.Data.IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("CITY", sCity)
            Return MyBase.loadBySQL(sSQL, paras)

        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

End Class


