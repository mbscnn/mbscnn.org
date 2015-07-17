Imports com.Azion.NET.VB

Public Class AP_ROADSEC
    Inherits BosBase

    Sub New()
        MyBase.New()
    End Sub

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.new("AP_ROADSEC", dbManager)
    End Sub

#Region "Thera Function"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 根據傳入的縣市代碼,鄉鎮市區代碼,段小段代碼，取得AP_ROADSEC的資料
    ''' </summary>
    ''' <param name="sCityId">縣市代碼</param>
    ''' <param name="sAreaId">鄉鎮市區代碼</param>
    ''' <param name="sSectionId">段小段代碼</param>
    ''' <returns>Boolean 是否有取得</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    '''FROM    TEMP_ROADSEC
    ''' </history>
    Function loadByPK(ByVal sCityId As String, ByVal sAreaId As String, ByVal sSectionId As String) As Boolean
        Try
            Dim sSQL As String = "  SELECT     *  " & _
                                         "     FROM    AP_ROADSEC  " & _
                                         "  WHERE     CITY_ID =" & ProviderFactory.PositionPara & "CITY_ID  " & _
                                          "       AND    AREA_ID =" & ProviderFactory.PositionPara & "AREA_ID  " & _
                                          "       AND    SEC_ID =" & ProviderFactory.PositionPara & "SEC_ID  "

            Dim paras(2) As System.Data.IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("CITY_ID", sCityId)
            paras(1) = ProviderFactory.CreateDataParameter("AREA_ID", sAreaId)
            paras(2) = ProviderFactory.CreateDataParameter("SEC_ID", sSectionId)
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
    ''' 以sCityId與sAreaId查詢，取得AREANAME, MAINSEC名稱資料
    ''' </summary>
    ''' <param name="sCityId"></param>
    ''' <param name="sAreaId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function getCityIdAreaId(ByVal sCityId As String, ByVal sAreaId As String) As Boolean
        Try
            Dim sSQL As String = "  SELECT     CITY, AREA, MAINSEC  " & _
                                          "     FROM    AP_ROADSEC  " & _
                                          "  WHERE     CITY_ID =" & ProviderFactory.PositionPara & "CITY_ID  " & _
                                          "       AND    AREA_ID =" & ProviderFactory.PositionPara & "AREA_ID  "

            Dim paras(1) As System.Data.IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("CITY_ID", sCityId)
            paras(1) = ProviderFactory.CreateDataParameter("AREA_ID", sAreaId)

            Return MyBase.loadBySQL(sSQL, paras)
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
    ''' 『沿用擔保品』根據傳入的縣市/鄉鎮市區/段小段名稱，取得AP_ROADSEC.name的資料
    ''' </summary>
    ''' <param name="sCityName">縣市</param>
    ''' <param name="sAreaName">鄉鎮市區</param>
    ''' <returns>Boolean 是否有取得</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' </history>
    Function getCityAreaName(ByVal sCityName As String, ByVal sAreaName As String) As Boolean
        Try
            Dim sSQL As String = "  SELECT     CITY_ID, AREA_ID  " & _
                                          "     FROM    AP_ROADSEC  " & _
                                          "  WHERE     CITY =" & ProviderFactory.PositionPara & "CITY  " & _
                                          "       AND    AREA =" & ProviderFactory.PositionPara & "AREA  "

            Dim paras(1) As System.Data.IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("CITY", sCityName)
            paras(1) = ProviderFactory.CreateDataParameter("AREA", sAreaName)

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
    ''' 以sCityId、sAreaId、sSecName為條件，取得段小段ID
    ''' </summary>
    ''' <param name="sCityId"></param>
    ''' <param name="sAreaId"></param>
    ''' <param name="sSecId"></param>
    ''' <returns></returns>
    ''' 20110817 thea add
    ''' <remarks></remarks>
    Function loadCityidAreaidSecidBySEC_ID(ByVal sCityId As String, ByVal sAreaId As String, ByVal sSecId As String) As String
        Try
            Dim sSQL As String = "  SELECT     SEC_ID  " & _
                                          "     FROM    AP_ROADSEC  " & _
                                          "  WHERE     CITY_ID =" & ProviderFactory.PositionPara & "CITY_ID  " & _
                                          "       AND    AREA_ID =" & ProviderFactory.PositionPara & "AREA_ID  " & _
                                          "       AND    SEC_ID =" & ProviderFactory.PositionPara & "SEC_ID"

            Dim paras(2) As System.Data.IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("CITYID", sCityId)
            paras(1) = ProviderFactory.CreateDataParameter("AREAID", sAreaId)
            paras(2) = ProviderFactory.CreateDataParameter("SEC_ID", sSecId)

            MyBase.loadBySQL(sSQL, paras)

            Return Me.getString("SEC_ID")

        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Function getCityAreaSecName(ByVal sCityId As String, ByVal sAreaId As String, ByVal sSecName As String) As String
        Try
            Dim sSQL As String = "  SELECT     SEC_ID  " & _
                                          "     FROM    AP_ROADSEC  " & _
                                          "  WHERE     CITY_ID =" & ProviderFactory.PositionPara & "CITY_ID  " & _
                                          "       AND    AREA_ID =" & ProviderFactory.PositionPara & "AREA_ID  " & _
                                          "       AND    MAINSEC =" & ProviderFactory.PositionPara & "MAINSEC  "

            Dim paras(2) As System.Data.IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("CITYID", sCityId)
            paras(1) = ProviderFactory.CreateDataParameter("AREAID", sAreaId)
            paras(2) = ProviderFactory.CreateDataParameter("MAINSEC", sSecName)

            MyBase.loadBySQL(sSQL, paras)
            Return Me.getString("SEC_ID")
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


