Imports com.Azion.NET.VB
Public Class GA_CASEAPVList
    Inherits BosList

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.New("GA_CASEAPV", dbManager)
    End Sub

    Overrides Function newBos() As BosBase
        Return New GA_CASEAPV(MyBase.getDatabaseManager)
    End Function

#Region "Either Function"

    ''' <summary>
    ''' 以單位代碼取得案件核准資料
    ''' </summary>
    ''' <param name="sAPV_BRH_ID">單位代碼</param>
    ''' <returns>資料筆數</returns>
    ''' <remarks></remarks>
    ''' <history>
    ''' 	[Either]	2011/08/13	Created
    ''' </history>
    Function LoadByAPV_BRH_ID(ByVal sAPV_BRH_ID As String) As Integer
        Try
            Dim sSql As String
            Dim paras As Data.IDbDataParameter

            sSql = " SELECT A.*,B.BRID,B.BRCNAME,C.ID,C.LNAM,C.SNAM,D.STAFFID,D.USERNAME " & _
                   " FROM GA_CASEAPV A " & _
                   " LEFT JOIN BRANCH B ON A.APV_BRH_ID = B.BRID " & _
                   " LEFT JOIN AP_CUSTOMER C ON A.APV_APL_ID = C.ID " & _
                   " LEFT JOIN USERINFO D ON 'S'+RIGHT(REPLICATE('0', 6) + CAST([APV_RM_COD] as VARCHAR), 6) = D.STAFFID " & _
                   " WHERE A.APV_BRH_ID = " + ProviderFactory.PositionPara + "APV_BRH_ID "

            paras = ProviderFactory.CreateDataParameter("APV_BRH_ID", sAPV_BRH_ID)

            Return MyBase.loadBySQLOnlyDs(paras)
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 以授信戶統編取得案件核准資料
    ''' </summary>
    ''' <param name="sAPV_APL_ID">授信戶統編</param>
    ''' <returns>資料筆數</returns>
    ''' <remarks></remarks>
    ''' <history>
    ''' 	[Either]	2011/08/13	Created
    ''' </history>
    Function LoadByAPV_APL_ID(ByVal sAPV_APL_ID As String) As Integer
        Try
            Dim sSql As String
            Dim paras As Data.IDbDataParameter

            sSql = " SELECT A.*,B.BRID,B.BRCNAME,C.ID,C.LNAM,C.SNAM,D.STAFFID,D.USERNAME " & _
                   " FROM GA_CASEAPV A " & _
                   " LEFT JOIN BRANCH B ON A.APV_BRH_ID = B.BRID " & _
                   " LEFT JOIN AP_CUSTOMER C ON A.APV_APL_ID = C.ID " & _
                   " LEFT JOIN USERINFO D ON 'S'+RIGHT(REPLICATE('0', 6) + CAST([APV_RM_COD] as VARCHAR), 6) = D.STAFFID " & _
                   " WHERE A.APV_APL_ID = " + ProviderFactory.PositionPara + "APV_APL_ID "

            paras = ProviderFactory.CreateDataParameter("APV_APL_ID", sAPV_APL_ID)

            Return MyBase.loadBySQLOnlyDs(sSql, paras)
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 以案件編號取得案件核准資料
    ''' </summary>
    ''' <param name="sAPV_CAS_ID">案件編號</param>
    ''' <returns>資料筆數</returns>
    ''' <remarks></remarks>
    ''' <history>
    ''' 	[Either]	2011/08/13	Created
    ''' </history>
    Function LoadByAPV_CAS_ID(ByVal sAPV_CAS_ID As String) As Integer
        Try
            Dim sSql As String
            Dim paras As Data.IDbDataParameter

            sSql = " SELECT A.*,B.BRID,B.BRCNAME,C.ID,C.LNAM,C.SNAM,D.STAFFID,D.USERNAME " & _
                   " FROM GA_CASEAPV A " & _
                   " LEFT JOIN BRANCH B ON A.APV_BRH_ID = B.BRID " & _
                   " LEFT JOIN AP_CUSTOMER C ON A.APV_APL_ID = C.ID " & _
                   " LEFT JOIN USERINFO D ON 'S'+RIGHT(REPLICATE('0', 6) + CAST([APV_RM_COD] as VARCHAR), 6) = D.STAFFID " & _
                   " WHERE A.APV_CAS_ID = " + ProviderFactory.PositionPara + "APV_CAS_ID "

            paras = ProviderFactory.CreateDataParameter("APV_CAS_ID", sAPV_CAS_ID)

            Return MyBase.loadBySQLOnlyDs(sSql, paras)
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
    ''' 以單位代碼,事業群取得案件核准資料
    ''' </summary>
    ''' <param name="sAPV_BRH_ID">單位代碼</param>
    ''' <param name="sAPV_BIZ_TYP">事業群</param>
    ''' <returns>資料筆數</returns>
    ''' <remarks></remarks>
    ''' <history>
    ''' 	[Ted]	2011/08/13	Created
    ''' </history>
    Function LoadByAPV_BRH_IDandAPV_BIZ_TYP(ByVal sAPV_BRH_ID As String, ByVal sAPV_BIZ_TYP As String) As Integer
        Try
            Dim sSql As String

            sSql = " SELECT A.*,B.BRID,B.BRCNAME,C.ID,C.LNAM,C.SNAM,D.STAFFID,D.USERNAME " & _
                   " FROM GA_CASEAPV A " & _
                   " LEFT JOIN BRANCH B ON A.APV_BRH_ID = B.BRID " & _
                   " LEFT JOIN AP_CUSTOMER C ON A.APV_APL_ID = C.ID " & _
                   " LEFT JOIN USERINFO D ON 'S'+RIGHT(REPLICATE('0', 6) + CAST([APV_RM_COD] as VARCHAR), 6) = D.STAFFID " & _
                   " WHERE A.APV_BRH_ID = " + ProviderFactory.PositionPara + "APV_BRH_ID " & _
                   "       AND A.APV_BIZ_TYP = " & ProviderFactory.PositionPara & "APV_BIZ_TYP " & _
                   "       AND A.APV_MARK IS NULL "

            Dim paras(1) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("APV_BRH_ID", sAPV_BRH_ID)
            paras(1) = ProviderFactory.CreateDataParameter("APV_BIZ_TYP", sAPV_BIZ_TYP)

            Return MyBase.loadBySQLOnlyDs(sSql, paras)
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 以授信戶統編,事業群取得案件核准資料
    ''' </summary>
    ''' <param name="sAPV_APL_ID">授信戶統編</param>
    ''' <param name="sAPV_BIZ_TYP">事業群</param>
    ''' <returns>資料筆數</returns>
    ''' <remarks></remarks>
    ''' <history>
    ''' 	[Either]	2011/08/13	Created
    ''' </history>
    Function LoadByAPV_APL_IDandAPV_BIZ_TYP(ByVal sAPV_APL_ID As String, ByVal sAPV_BIZ_TYP As String) As Integer
        Try
            Dim sSql As String
            Dim paras(1) As Data.IDbDataParameter

            sSql = " SELECT A.*,B.BRID,B.BRCNAME,C.ID,C.LNAM,C.SNAM,D.STAFFID,D.USERNAME " & _
                   " FROM GA_CASEAPV A " & _
                   " LEFT JOIN BRANCH B ON A.APV_BRH_ID = B.BRID " & _
                   " LEFT JOIN AP_CUSTOMER C ON A.APV_APL_ID = C.ID " & _
                   " LEFT JOIN USERINFO D ON 'S'+RIGHT(REPLICATE('0', 6) + CAST([APV_RM_COD] as VARCHAR), 6) = D.STAFFID " & _
                   " WHERE A.APV_APL_ID = " + ProviderFactory.PositionPara + "APV_APL_ID " & _
                   "       AND APV_BIZ_TYP = " & ProviderFactory.PositionPara & "APV_BIZ_TYP " & _
                   "       AND A.APV_MARK IS NULL "

            paras(0) = ProviderFactory.CreateDataParameter("APV_APL_ID", sAPV_APL_ID)
            paras(1) = ProviderFactory.CreateDataParameter("APV_BIZ_TYP", sAPV_BIZ_TYP)

            Return MyBase.loadBySQLOnlyDs(sSql, paras)
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

    ''' <summary>
    ''' 以事業群,案件註記取得案件核准資料
    ''' </summary>
    ''' <param name="sAPV_BIZ_TYP">事業群</param>
    ''' <param name="sAPV_MARK">案件註記</param>
    ''' <returns>資料筆數</returns>
    ''' <remarks></remarks>
    ''' <history>
    ''' 	[Nick]	2011/09/14	Created
    ''' </history>
    Function LoadByBIZ_TYPandMARK(ByVal sAPV_BIZ_TYP As String, ByVal sAPV_MARK As String, Optional ByVal sAPV_APL_ID As String = "", Optional ByVal sAPV_BRID As String = "") As Integer
        Try
            Dim sSql As String = String.Empty
            Dim tmpPara(3) As IDbDataParameter '預定傳入參數
            Dim iCnt As Integer = -1
            Dim sUpCode As String = String.Empty
            If sAPV_BIZ_TYP = "1" Then
                sUpCode = "918"
            Else
                sUpCode = "924"
            End If

            sSql = " SELECT A.*,B.BRCNAME,C.LNAM,C.SNAM,D.USERNAME,E.TEXT " & _
                   " FROM GA_CASEAPV A " & _
                   " LEFT JOIN BRANCH B ON A.APV_BRH_ID = B.BRID " & _
                   " LEFT JOIN AP_CUSTOMER C ON A.APV_APL_ID = C.ID " & _
                   " LEFT JOIN USERINFO D ON 'S'+RIGHT(REPLICATE('0', 6) + CAST([APV_RM_COD] as VARCHAR), 6) = D.STAFFID " & _
                   " LEFT JOIN AP_CODE E ON A.APV_CHKLEV = E.NOTE and E.UPCODE = " & ProviderFactory.PositionPara & "UPCODE"

            iCnt += 1
            tmpPara(iCnt) = ProviderFactory.CreateDataParameter("UPCODE", sUpCode)

            If sAPV_BIZ_TYP <> Nothing Then
                sSql &= "  WHERE A.APV_BIZ_TYP = " & ProviderFactory.PositionPara & "APV_BIZ_TYP"
                iCnt += 1
                tmpPara(iCnt) = ProviderFactory.CreateDataParameter("APV_BIZ_TYP", sAPV_BIZ_TYP)
            End If


            If sAPV_APL_ID <> Nothing Then
                If sAPV_BIZ_TYP = Nothing Then
                    sSql &= " where  A.APV_APL_ID = " & ProviderFactory.PositionPara & "APV_APL_ID"
                Else
                    sSql &= IIf(iCnt > 0, " and ", " where ") & " A.APV_APL_ID = " & ProviderFactory.PositionPara & "APV_APL_ID"
                End If
                iCnt += 1
                tmpPara(iCnt) = ProviderFactory.CreateDataParameter("APV_APL_ID", sAPV_APL_ID)
            Else
                If sAPV_MARK <> Nothing Then
                    sSql &= IIf(iCnt > 0, " and ", " where ") & " A.APV_MARK = " & ProviderFactory.PositionPara & "APV_MARK"
                    iCnt += 1
                    tmpPara(iCnt) = ProviderFactory.CreateDataParameter("APV_MARK", sAPV_MARK)
                Else
                    sSql &= IIf(iCnt > 0, " and ", " where ") & "  A.APV_MARK is null "
                End If
            End If

            If sAPV_BRID <> "" Then
                sSql &= " AND A.APV_BRH_ID= " & ProviderFactory.PositionPara & "APV_BRH_ID"
                iCnt += 1
                tmpPara(iCnt) = ProviderFactory.CreateDataParameter("APV_BRH_ID", sAPV_BRID)
            End If


            Dim paras(iCnt) As System.Data.IDbDataParameter '正式要送出的參數
            For i As Integer = 0 To iCnt
                paras(i) = tmpPara(i)
            Next

            Return Me.loadBySQL(sSql, paras)

            'If sAPV_BIZ_TYP <> "" And sAPV_MARK <> "" Then
            '    Dim paras(2) As System.Data.IDbDataParameter
            '    paras(0) = ProviderFactory.CreateDataParameter("UPCODE", sUpCode)
            '    paras(1) = ProviderFactory.CreateDataParameter("APV_BIZ_TYP", sAPV_BIZ_TYP)
            '    paras(2) = ProviderFactory.CreateDataParameter("APV_MARK", sAPV_MARK)
            '    Return Me.loadBySQL(sSql, paras)
            'ElseIf sAPV_BIZ_TYP <> "" And sAPV_MARK = Nothing Then
            '    Dim paras(1) As System.Data.IDbDataParameter
            '    paras(0) = ProviderFactory.CreateDataParameter("UPCODE", sUpCode)
            '    paras(1) = ProviderFactory.CreateDataParameter("APV_BIZ_TYP", sAPV_BIZ_TYP)
            '    Return MyBase.loadBySQLOnlyDs(sSql, paras)
            'ElseIf sAPV_BIZ_TYP = Nothing And sAPV_MARK <> "" Then
            '    Dim paras(1) As System.Data.IDbDataParameter
            '    paras(0) = ProviderFactory.CreateDataParameter("UPCODE", sUpCode)
            '    paras(1) = ProviderFactory.CreateDataParameter("APV_MARK", sAPV_MARK)
            '    Return MyBase.loadBySQLOnlyDs(sSql, paras)
            'ElseIf sAPV_BIZ_TYP = Nothing And sAPV_MARK = Nothing Then
            '    Dim paras(0) As System.Data.IDbDataParameter
            '    paras(0) = ProviderFactory.CreateDataParameter("UPCODE", sUpCode)
            '    Return MyBase.loadBySQLOnlyDs(sSql, paras)
            'End If


        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

#End Region

    ''' <summary>
    ''' 以授信戶統編取得案件核准資料
    ''' </summary>
    ''' <param name="sAPV_APL_ID">授信戶統編</param>
    ''' <returns>資料筆數</returns>
    ''' <remarks></remarks>
    ''' <history>
    ''' 	[Steven]	2011/12/26	Created
    ''' </history>
    Function LoadByAPV_APL_IDandMARK(ByVal sAPV_APL_ID As String) As Integer
        Try
            Dim sSql As String
            Dim paras As Data.IDbDataParameter

            sSql = " SELECT * FROM GA_CASEAPV" & _
                   " WHERE APV_APL_ID = " + ProviderFactory.PositionPara + "APV_APL_ID " & _
                   " and APV_MARK='3' "

            paras = ProviderFactory.CreateDataParameter("APV_APL_ID", sAPV_APL_ID)

            Return MyBase.loadBySQLOnlyDs(sSql, paras)
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function
End Class
