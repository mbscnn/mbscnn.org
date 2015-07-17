Imports MBSC.MB_OP
Imports MBSC.UICtl
Imports com.Azion.NET.VB

Public Class MBQry_TYPE_01_02_v01
    Inherits System.Web.UI.Page

    Dim m_sBY As String = String.Empty
    Dim m_sBM As String = String.Empty
    Dim m_sBD As String = String.Empty
    Dim m_sEY As String = String.Empty
    Dim m_sEM As String = String.Empty
    Dim m_sED As String = String.Empty
    Dim m_sMB_AREA As String = String.Empty
    Dim m_sMB_LEADER As String = String.Empty
    Dim m_DT_UpCode7 As DataTable = Nothing
    Dim m_sUpCode7 As String = String.Empty
    Dim DT_UPCODE15 As DataTable = Nothing
    Dim m_sUpcode15 As String = String.Empty


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.m_sBY = "" & Request.QueryString("BY")
            Me.m_sBM = "" & Request.QueryString("BM")
            Me.m_sBD = "" & Request.QueryString("BD")
            Me.m_sEY = "" & Request.QueryString("EY")
            Me.m_sEM = "" & Request.QueryString("EM")
            Me.m_sED = "" & Request.QueryString("ED")
            Me.m_sMB_AREA = "" & Request.QueryString("MB_AREA")
            Me.m_sMB_LEADER = "" & Request.QueryString("MB_LEADER")

            '所屬區代碼
            Me.m_sUpCode7 = com.Azion.EloanUtility.CodeList.getAppSettings("UpCode7")

            '功德項目
            Me.m_sUpcode15 = com.Azion.EloanUtility.CodeList.getAppSettings("UpCode15")

            Me.Bind_TYPE_2()
        Catch ex As System.Threading.ThreadAbortException
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

    Sub Bind_TYPE_2()
        Dim DBManager As DatabaseManager = UIShareFun.getDataBaseManager
        Dim objStreamReader As System.IO.StreamReader = Nothing
        Dim DT_MB_MEMREV As DataTable = Nothing
        Try
            If System.IO.File.Exists(Server.MapPath(com.Azion.EloanUtility.UIUtility.getRootPath & "/QRY/MBQry_TYPE_01_02_v01.xml")) Then
                objStreamReader = New System.IO.StreamReader(Server.MapPath(com.Azion.EloanUtility.UIUtility.getRootPath & "/QRY/MBQry_TYPE_01_02_v01.xml"))
            Else
                Throw New Exception("EXCEL範本檔案(MBQry_TYPE_01_02_v01)不存在")
            End If

            Dim D_BDay As Object = Convert.DBNull
            D_BDay = Me.ConvertWDate(Me.m_sBY, Me.m_sBM, Me.m_sBD)
            Dim D_EDay As Object = Convert.DBNull
            D_EDay = Me.ConvertWDate(Me.m_sEY, Me.m_sEM, Me.m_sED)

            Dim isPeriod As Boolean = False
            If IsDate(D_BDay) AndAlso IsDate(D_EDay) Then
                isPeriod = True
            End If

            Dim mbMEMREVList As New MB_MEMREVList(DBManager)
            If Not isPeriod AndAlso Not Utility.isValidateData(Me.m_sMB_AREA) AndAlso Not Utility.isValidateData(Me.m_sMB_LEADER) Then
                mbMEMREVList.QRY_EXCEL_1()
            ElseIf isPeriod AndAlso Not Utility.isValidateData(Me.m_sMB_AREA) AndAlso Not Utility.isValidateData(Me.m_sMB_LEADER) Then
                mbMEMREVList.QRY_EXCEL_2(D_BDay, D_EDay)
            ElseIf isPeriod AndAlso Utility.isValidateData(Me.m_sMB_AREA) AndAlso Not Utility.isValidateData(Me.m_sMB_LEADER) Then
                mbMEMREVList.QRY_EXCEL_3(D_BDay, D_EDay, Me.m_sMB_AREA)
            ElseIf isPeriod AndAlso Not Utility.isValidateData(Me.m_sMB_AREA) AndAlso Utility.isValidateData(Me.m_sMB_LEADER) Then
                mbMEMREVList.QRY_EXCEL_4(D_BDay, D_EDay, Me.m_sMB_LEADER)
            ElseIf isPeriod AndAlso Utility.isValidateData(Me.m_sMB_AREA) AndAlso Utility.isValidateData(Me.m_sMB_LEADER) Then
                mbMEMREVList.QRY_EXCEL_5(D_BDay, D_EDay, Me.m_sMB_AREA, Me.m_sMB_LEADER)
            ElseIf Not isPeriod AndAlso Utility.isValidateData(Me.m_sMB_AREA) AndAlso Not Utility.isValidateData(Me.m_sMB_LEADER) Then
                mbMEMREVList.QRY_EXCEL_6(Me.m_sMB_AREA)
            ElseIf Not isPeriod AndAlso Utility.isValidateData(Me.m_sMB_AREA) AndAlso Utility.isValidateData(Me.m_sMB_LEADER) Then
                mbMEMREVList.QRY_EXCEL_7(Me.m_sMB_AREA, Me.m_sMB_LEADER)
            End If
            DT_MB_MEMREV = mbMEMREVList.getCurrentDataSet.Tables(0)

            Dim sbCtlContent As New System.Text.StringBuilder(objStreamReader.ReadToEnd())

            Dim objStringBuilder As New System.Text.StringBuilder

            For Each ROW As DataRow In DT_MB_MEMREV.Rows
                If IsNumeric(ROW("MB_MEMSEQ")) Then
                    Me.Bind_TYPE_2_DTL(DBManager, ROW("MB_MEMSEQ"), objStringBuilder)
                End If
            Next

            sbCtlContent.Replace("{{DATAROW}}", objStringBuilder.ToString)

            Dim iYYYYMMDD As Decimal = 0
            iYYYYMMDD = Now.Year * 10000 + Now.Month * 100 + Now.Day
            UIShareFun.downLoadFile(Me.Page, iYYYYMMDD.ToString & ".xls", sbCtlContent.ToString)
        Catch ex As System.Threading.ThreadAbortException
        Catch ex As Exception
            Throw
        Finally
            UIShareFun.releaseConnection(DBManager)

            If Not IsNothing(objStreamReader) Then
                objStreamReader.Close()
            End If

            If Not IsNothing(DT_MB_MEMREV) Then
                DT_MB_MEMREV.Dispose()
            End If
        End Try
    End Sub

    Sub Bind_TYPE_2_DTL(ByVal DBManager As DatabaseManager, ByVal iMB_MEMSEQ As Decimal, ByVal objStringBuilder As System.Text.StringBuilder)
        Dim DT_DTL As DataTable = Nothing
        Try
            Dim D_BDay As Object = Convert.DBNull
            D_BDay = Me.ConvertWDate(Me.m_sBY, Me.m_sBM, Me.m_sBD)
            Dim D_EDay As Object = Convert.DBNull
            D_EDay = Me.ConvertWDate(Me.m_sEY, Me.m_sEM, Me.m_sED)

            Dim isPeriod As Boolean = False
            If IsDate(D_BDay) AndAlso IsDate(D_EDay) Then
                isPeriod = True
            End If

            Dim mbMEMREVList As New MB_MEMREVList(DBManager)
            If Not isPeriod AndAlso Not Utility.isValidateData(Me.m_sMB_AREA) AndAlso Not Utility.isValidateData(Me.m_sMB_LEADER) Then
                mbMEMREVList.DTL_1(iMB_MEMSEQ)
            ElseIf isPeriod AndAlso Not Utility.isValidateData(Me.m_sMB_AREA) AndAlso Not Utility.isValidateData(Me.m_sMB_LEADER) Then
                mbMEMREVList.DTL_2(iMB_MEMSEQ, D_BDay, D_EDay)
            ElseIf isPeriod AndAlso Utility.isValidateData(Me.m_sMB_AREA) AndAlso Not Utility.isValidateData(Me.m_sMB_LEADER) Then
                mbMEMREVList.DTL_2(iMB_MEMSEQ, D_BDay, D_EDay)
            ElseIf isPeriod AndAlso Not Utility.isValidateData(Me.m_sMB_AREA) AndAlso Utility.isValidateData(Me.m_sMB_LEADER) Then
                mbMEMREVList.DTL_2(iMB_MEMSEQ, D_BDay, D_EDay)
            ElseIf isPeriod AndAlso Utility.isValidateData(Me.m_sMB_AREA) AndAlso Utility.isValidateData(Me.m_sMB_LEADER) Then
                mbMEMREVList.DTL_2(iMB_MEMSEQ, D_BDay, D_EDay)
            ElseIf Not isPeriod AndAlso Utility.isValidateData(Me.m_sMB_AREA) AndAlso Not Utility.isValidateData(Me.m_sMB_LEADER) Then
                mbMEMREVList.DTL_1(iMB_MEMSEQ)
            ElseIf Not isPeriod AndAlso Utility.isValidateData(Me.m_sMB_AREA) AndAlso Utility.isValidateData(Me.m_sMB_LEADER) Then
                mbMEMREVList.DTL_1(iMB_MEMSEQ)
            End If

            DT_DTL = mbMEMREVList.getCurrentDataSet.Tables(0)
            DT_DTL.Columns.Add("CNT", Type.GetType("System.Decimal"))

            Dim sqlRow As DataRow = DT_DTL.NewRow
            sqlRow("MB_TX_DATE") = Convert.DBNull
            sqlRow("MB_TOTFEE") = DT_DTL.Compute("SUM(MB_TOTFEE)", String.Empty)
            sqlRow("CNT") = DT_DTL.Rows.Count
            DT_DTL.Rows.Add(sqlRow)

            For i As Integer = 0 To DT_DTL.Rows.Count - 2
                Dim ROW As DataRow = DT_DTL.Rows(i)

                objStringBuilder.Append("<Row ss:AutoFitHeight=""1"" ss:Height=""50"" >")

                '繳款日期
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & ROW("MB_TX_DATE").ToString & "</Data></Cell>")

                '所屬區
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & Me.getMB_AREA(DBManager, ROW("MB_AREA").ToString) & "</Data></Cell>")

                '委員
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & ROW("MB_LEADER").ToString & "</Data></Cell>")

                '會員姓名
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & ROW("MB_NAME").ToString & "</Data></Cell>")

                '會員編號
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & Me.getMB_MEMSEQ(ROW("MB_MEMSEQ").ToString, ROW("MB_AREA").ToString) & "</Data></Cell>")

                '功德項目
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & Me.getMB_ITEMID(DBManager, ROW("MB_ITEMID").ToString) & "</Data></Cell>")

                '繳款金額
                If IsNumeric(ROW("MB_TOTFEE")) Then
                    objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & Format(ROW("MB_TOTFEE"), "#,##0") & "</Data></Cell>")
                Else
                    objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & "" & "</Data></Cell>")
                End If

                '筆數
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & ROW("CNT").ToString & "</Data></Cell>")

                objStringBuilder.Append("</Row>")
            Next

            objStringBuilder.Append("<Row ss:AutoFitHeight=""1"" ss:Height=""50"" >")
            Dim ROWLST As DataRow = DT_DTL.Rows(DT_DTL.Rows.Count - 1)
            objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & "小計" & "</Data></Cell>")
            objStringBuilder.Append("<Cell ss:MergeAcross=""4""><Data ss:Type=""String"">" & "" & "</Data></Cell>")
            '繳款金額
            If IsNumeric(ROWLST("MB_TOTFEE")) Then
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & Format(ROWLST("MB_TOTFEE"), "#,##0") & "</Data></Cell>")
            Else
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & "" & "</Data></Cell>")
            End If
            '筆數
            objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & ROWLST("CNT").ToString & "</Data></Cell>")
            objStringBuilder.Append("</Row>")
        Catch ex As Exception
            Throw
        Finally
            If Not IsNothing(DT_DTL) Then
                DT_DTL.Dispose()
            End If
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

    Function getMB_ITEMID(ByVal DBManager As DatabaseManager, ByVal sMB_ITEMID As Object) As String
        Try
            If Utility.isValidateData(sMB_ITEMID) Then
                If IsNothing(Me.DT_UPCODE15) Then
                    Dim apCODEList As New AP_CODEList(DBManager)
                    apCODEList.loadByUpCode(Me.m_sUpcode15)
                    Me.DT_UPCODE15 = apCODEList.getCurrentDataSet.Tables(0)
                End If

                Dim ROW_Select() As DataRow = Me.DT_UPCODE15.Select("VALUE='" & sMB_ITEMID & "'")
                If Not IsNothing(ROW_Select) AndAlso ROW_Select.Length > 0 Then
                    Return ROW_Select(0)("TEXT").ToString()
                End If
            End If

            Return String.Empty
        Catch ex As Exception
            Throw
        End Try
    End Function
End Class