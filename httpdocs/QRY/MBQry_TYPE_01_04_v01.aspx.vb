Imports MBSC.MB_OP
Imports com.Azion.NET.VB
Imports MBSC.UICtl
Public Class MBQry_TYPE_01_04_v01
    Inherits System.Web.UI.Page

    Dim m_sBY As String = String.Empty
    Dim m_sBM As String = String.Empty
    Dim m_sBD As String = String.Empty
    Dim m_sEY As String = String.Empty
    Dim m_sEM As String = String.Empty
    Dim m_sED As String = String.Empty
    Dim m_sMB_MEMTYP As String = String.Empty
    Dim m_sMB_FEETYPE As String = String.Empty
    Dim m_DT_UpCode7 As DataTable = Nothing
    Dim m_sUpCode7 As String = String.Empty
    Dim m_sUpCode28 As String = String.Empty
    Dim m_DT_UpCode28 As DataTable = Nothing
    Dim m_DT_UpCode23 As DataTable = Nothing
    Dim m_sUpcode23 As String = String.Empty

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.m_sBY = "" & Request.QueryString("BY")
            Me.m_sBM = "" & Request.QueryString("BM")
            Me.m_sBD = "" & Request.QueryString("BD")
            Me.m_sEY = "" & Request.QueryString("EY")
            Me.m_sEM = "" & Request.QueryString("EM")
            Me.m_sED = "" & Request.QueryString("ED")
            Me.m_sMB_MEMTYP = "" & Request.QueryString("MB_MEMTYP")
            Me.m_sMB_FEETYPE = "" & Request.QueryString("MB_FEETYPE")

            '所屬區代碼
            Me.m_sUpCode7 = com.Azion.EloanUtility.CodeList.getAppSettings("UpCode7")

            '會員類別
            Me.m_sUpCode28 = com.Azion.EloanUtility.CodeList.getAppSettings("UpCode28")

            '繳款方式
            m_sUpcode23 = com.Azion.EloanUtility.CodeList.getAppSettings("UpCode23")

            Me.Bind_TYPE_4()
        Catch ex As System.Threading.ThreadAbortException
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

    Sub Bind_TYPE_4()
        Dim DT_MB_MEMREV As DataTable = Nothing
        Dim DBManager As DatabaseManager = UIShareFun.getDataBaseManager
        Dim objStreamReader As System.IO.StreamReader = Nothing
        Try
            Dim D_BDay As Object = Convert.DBNull
            D_BDay = Me.ConvertWDate(Me.m_sBY, Me.m_sBM, Me.m_sBD)
            Dim D_EDay As Object = Convert.DBNull
            D_EDay = Me.ConvertWDate(Me.m_sEY, Me.m_sEM, Me.m_sED)

            Dim isPeriod As Boolean = False
            Dim iB_YYYYMMDD As Decimal = 0
            Dim iE_YYYYMMDD As Decimal = 0
            If IsDate(D_BDay) AndAlso IsDate(D_EDay) Then
                isPeriod = True

                iB_YYYYMMDD = CDate(D_BDay).Year * 10000 + CDate(D_BDay).Month * 100 + CDate(D_BDay).Day
                iE_YYYYMMDD = CDate(D_EDay).Year * 10000 + CDate(D_EDay).Month * 100 + CDate(D_EDay).Day
            End If

            Dim mbMEMREVList As New MB_MEMREVList(DBManager)
            If isPeriod AndAlso Not Utility.isValidateData(Me.m_sMB_MEMTYP) AndAlso Not Utility.isValidateData(Me.m_sMB_FEETYPE) Then
                mbMEMREVList.QRY_EXCEL_INCOME_1(iB_YYYYMMDD, iE_YYYYMMDD)
            ElseIf isPeriod AndAlso Utility.isValidateData(Me.m_sMB_MEMTYP) AndAlso Utility.isValidateData(Me.m_sMB_FEETYPE) Then
                mbMEMREVList.QRY_EXCEL_INCOME_2(iB_YYYYMMDD, iE_YYYYMMDD, Me.m_sMB_MEMTYP, Me.m_sMB_FEETYPE)
            ElseIf isPeriod AndAlso Not Utility.isValidateData(Me.m_sMB_MEMTYP) AndAlso Utility.isValidateData(Me.m_sMB_FEETYPE) Then
                mbMEMREVList.QRY_EXCEL_INCOME_3(iB_YYYYMMDD, iE_YYYYMMDD, Me.m_sMB_FEETYPE)
            ElseIf Not isPeriod AndAlso Utility.isValidateData(Me.m_sMB_MEMTYP) AndAlso Utility.isValidateData(Me.m_sMB_FEETYPE) Then
                mbMEMREVList.QRY_EXCEL_INCOME_4(Me.m_sMB_MEMTYP, Me.m_sMB_FEETYPE)
            ElseIf Not isPeriod AndAlso Utility.isValidateData(Me.m_sMB_MEMTYP) AndAlso Not Utility.isValidateData(Me.m_sMB_FEETYPE) Then
                mbMEMREVList.QRY_EXCEL_INCOME_5(Me.m_sMB_MEMTYP)
            ElseIf Not isPeriod AndAlso Not Utility.isValidateData(Me.m_sMB_MEMTYP) AndAlso Utility.isValidateData(Me.m_sMB_FEETYPE) Then
                mbMEMREVList.QRY_EXCEL_INCOME_6(Me.m_sMB_FEETYPE)
            End If

            DT_MB_MEMREV = mbMEMREVList.getCurrentDataSet.Tables(0)

            If System.IO.File.Exists(Server.MapPath(com.Azion.EloanUtility.UIUtility.getRootPath & "/QRY/MBQry_TYPE_01_04_v01.xml")) Then
                objStreamReader = New System.IO.StreamReader(Server.MapPath(com.Azion.EloanUtility.UIUtility.getRootPath & "/QRY/MBQry_TYPE_01_04_v01.xml"))
            Else
                Throw New Exception("EXCEL範本檔案(MBQry_TYPE_01_04_v01)不存在")
            End If

            Dim sbCtlContent As New System.Text.StringBuilder(objStreamReader.ReadToEnd())

            Dim objStringBuilder As New System.Text.StringBuilder

            For Each ROW As DataRow In DT_MB_MEMREV.Rows
                objStringBuilder.Append("<Row ss:AutoFitHeight=""1"" ss:Height=""50"" >")

                '區別
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & Me.getMB_AREA(DBManager, ROW("MB_AREA").ToString) & "</Data></Cell>")

                '委員
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & ROW("MB_LEADER").ToString & "</Data></Cell>")

                '日期
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & Me.get_DECIMAL_DATE(ROW("MB_TX_DATE")) & "</Data></Cell>")

                '會員編號
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & Me.getMB_MEMSEQ(ROW("MB_MEMSEQ"), ROW("MB_AREA")) & "</Data></Cell>")

                '捐款人
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & ROW("MB_NAME").ToString & "</Data></Cell>")

                '會費起訖日
                Dim iMB_MEMFEE_SY As Decimal = 0
                If IsNumeric(ROW("MB_MEMFEE_SY")) Then
                    iMB_MEMFEE_SY = ROW("MB_MEMFEE_SY")
                End If
                Dim iMB_MEMFEE_SM As Decimal = 0
                If IsNumeric(ROW("MB_MEMFEE_SM")) Then
                    iMB_MEMFEE_SM = ROW("MB_MEMFEE_SM")
                End If
                Dim sSYM As String = iMB_MEMFEE_SY.ToString & "/" & iMB_MEMFEE_SM.ToString & "/1"
                Dim sBDate As String = String.Empty
                If IsDate(sSYM) Then
                    sBDate = Utility.FillZero(iMB_MEMFEE_SY, 4) & Utility.FillZero(iMB_MEMFEE_SM, 2)
                End If

                Dim iMB_MEMFEE_EY As Decimal = 0
                If IsNumeric(ROW("MB_MEMFEE_EY")) Then
                    iMB_MEMFEE_EY = ROW("MB_MEMFEE_EY")
                End If
                Dim iMB_MEMFEE_EM As Decimal = 0
                If IsNumeric(ROW("MB_MEMFEE_EM")) Then
                    iMB_MEMFEE_EM = ROW("MB_MEMFEE_EM")
                End If
                Dim sEYM As String = String.Empty
                sEYM = iMB_MEMFEE_EY.ToString & "/" & iMB_MEMFEE_EM.ToString & "/1"
                Dim sEDate As String = String.Empty
                If IsDate(sEYM) Then
                    sEDate = Utility.FillZero(iMB_MEMFEE_EY, 4) & Utility.FillZero(iMB_MEMFEE_EM, 2)
                End If
                Dim sMB_MEMFEE_SYMEYM As String = String.Empty
                sMB_MEMFEE_SYMEYM = sBDate & "~" & sEDate
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & sMB_MEMFEE_SYMEYM & "</Data></Cell>")

                '會員種類
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & Me.getMB_FAMILY(DBManager, ROW("MB_MEMTYP")) & "</Data></Cell>")

                '繳款方式
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & Me.getMB_FEETYPE(DBManager, ROW("MB_FEETYPE")) & "</Data></Cell>")

                '每月金額
                If IsNumeric(ROW("MB_MONFEE")) Then
                    objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & Format(ROW("MB_MONFEE"), "#,##0") & "</Data></Cell>")
                Else
                    objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & String.Empty & "</Data></Cell>")
                End If

                '繳款金額
                If IsNumeric(ROW("MB_TOTFEE")) Then
                    objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & Format(ROW("MB_TOTFEE"), "#,##0") & "</Data></Cell>")
                Else
                    objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & String.Empty & "</Data></Cell>")
                End If

                '護僧基金
                If ROW("FUND1").ToString = "1" Then
                    objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & "V" & "</Data></Cell>")
                Else
                    objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & String.Empty & "</Data></Cell>")
                End If

                '建設基金
                If ROW("FUND2").ToString = "1" Then
                    objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & "V" & "</Data></Cell>")
                Else
                    objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & String.Empty & "</Data></Cell>")
                End If

                '弘法基金
                If ROW("FUND3").ToString = "1" Then
                    objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & "V" & "</Data></Cell>")
                Else
                    objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & String.Empty & "</Data></Cell>")
                End If

                '列印收據
                Select Case ROW("MB_PRINT").ToString
                    Case "1"
                        objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & "要列印" & "</Data></Cell>")
                    Case "0"
                        objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & "不列印" & "</Data></Cell>")
                    Case "Y"
                        objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & "已經列印" & "</Data></Cell>")
                    Case Else
                        objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & "不列印" & "</Data></Cell>")
                End Select

                '收據編號
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & ROW("MB_RECV_Y").ToString & ROW("MB_RECV_N").ToString & "</Data></Cell>")

                objStringBuilder.Append("</Row>")
            Next

            sbCtlContent.Replace("{{DATAROW}}", objStringBuilder.ToString)

            Dim iYYYYMMDD As Decimal = 0
            iYYYYMMDD = Now.Year * 10000 + Now.Month * 100 + Now.Day
            UIShareFun.downLoadFile(Me.Page, iYYYYMMDD.ToString & ".xls", sbCtlContent.ToString)
        Catch ex As System.Threading.ThreadAbortException
        Catch ex As Exception
            Throw
        Finally
            If Not IsNothing(DT_MB_MEMREV) Then
                DT_MB_MEMREV.Dispose()
            End If

            If Not IsNothing(objStreamReader) Then
                objStreamReader.Close()
            End If

            UIShareFun.releaseConnection(DBManager)
        End Try
    End Sub

    Function ConvertWDate(ByVal sYYY As String, ByVal sMM As String, ByVal sDD As String) As Object
        Try
            sYYY = Trim(sYYY)
            sMM = Trim(sMM)
            sDD = Trim(sDD)

            If IsNumeric(sYYY) AndAlso IsNumeric(sMM) AndAlso IsNumeric(sDD) Then
                Dim sYYYY As String = String.Empty
                sYYYY = CDec(sYYY).ToString

                Dim sDateStr As String = String.Empty
                sDateStr = sYYYY & "/" & sMM & "/" & sDD

                If IsDate(sDateStr) Then
                    Return CDate(sDateStr)
                Else
                    Return Convert.DBNull
                End If
            Else
                Return Convert.DBNull
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function

    Function getMB_AREA(ByVal DBManager As DatabaseManager, ByVal sMB_AREA As Object) As String
        Try
            If Utility.isValidateData(sMB_AREA) Then
                If IsNothing(Me.m_DT_UpCode7) Then
                    Dim apCODEList As New AP_CODEList(DBManager)
                    apCODEList.loadByUpCode(Me.m_sUpCode7)
                    Me.m_DT_UpCode7 = apCODEList.getCurrentDataSet.Tables(0)
                End If

                Dim ROW_SELECT() As DataRow = Nothing
                ROW_SELECT = Me.m_DT_UpCode7.Select("VALUE='" & sMB_AREA & "'")
                If Not IsNothing(ROW_SELECT) AndAlso ROW_SELECT.Length > 0 Then
                    Return ROW_SELECT(0)("TEXT").ToString
                End If
            End If

            Return String.Empty
        Catch ex As Exception
            Throw
        End Try
    End Function

    Function get_DECIMAL_DATE(ByVal MB_TX_DATE As Object) As String
        Try
            If Utility.isDecimalDate(MB_TX_DATE) Then
                Return Utility.DateTransfer(Left(MB_TX_DATE.ToString, 4) & "/" & MB_TX_DATE.ToString.Substring(4, 2) & "/" & Right(MB_TX_DATE.ToString, 2))
            End If

            Return String.Empty
        Catch ex As Exception
            Throw
        End Try
    End Function

    Function getMB_MEMSEQ(ByVal iMB_MEMSEQ As Object, ByVal sMB_AREA As Object) As String
        Try
            If IsNumeric(iMB_MEMSEQ) AndAlso Utility.isValidateData(sMB_AREA) Then
                Return sMB_AREA & "-" & com.Azion.EloanUtility.StrUtility.FillZero(CDec(iMB_MEMSEQ), 7)
            End If
            Return String.Empty
        Catch ex As Exception
            Throw
        End Try
    End Function

    Function getMB_FAMILY(ByVal DBManager As DatabaseManager, ByVal sMB_FAMILY As Object) As String
        Try
            If Utility.isValidateData(sMB_FAMILY) Then
                If IsNothing(Me.m_DT_UpCode28) Then
                    Dim apCODEList As New AP_CODEList(DBManager)
                    apCODEList.loadByUpCode(Me.m_sUpCode28)
                    Me.m_DT_UpCode28 = apCODEList.getCurrentDataSet.Tables(0)
                End If

                Dim ROW_SELECT() As DataRow = Nothing
                ROW_SELECT = Me.m_DT_UpCode28.Select("VALUE='" & sMB_FAMILY & "'")
                If Not IsNothing(ROW_SELECT) AndAlso ROW_SELECT.Length > 0 Then
                    Return ROW_SELECT(0)("TEXT").ToString
                End If
            End If

            Return String.Empty
        Catch ex As Exception
            Throw
        End Try
    End Function

    Function getMB_FEETYPE(ByVal DBManager As DatabaseManager, ByVal sMB_FEETYPE As Object) As String
        Try
            If Utility.isValidateData(sMB_FEETYPE) Then
                If IsNothing(Me.m_DT_UpCode23) Then
                    Dim apCODEList As New AP_CODEList(DBManager)
                    apCODEList.loadByUpCode(Me.m_sUpcode23)
                    Me.m_DT_UpCode23 = apCODEList.getCurrentDataSet.Tables(0)
                End If

                Dim ROW_SELECT() As DataRow = Me.m_DT_UpCode23.Select("VALUE='" & sMB_FEETYPE & "'")
                If Not IsNothing(ROW_SELECT) AndAlso ROW_SELECT.Length > 0 Then
                    Return ROW_SELECT(0)("TEXT")
                End If
            End If

            Return String.Empty
        Catch ex As Exception
            Throw
        End Try
    End Function
End Class