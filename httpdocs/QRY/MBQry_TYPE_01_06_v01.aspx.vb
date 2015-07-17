Imports com.Azion.NET.VB
Imports MBSC.MB_OP
Imports MBSC.UICtl

Public Class MBQry_TYPE_01_06_v01
    Inherits System.Web.UI.Page

    Dim m_sMB_AREA As String = String.Empty
    Dim m_sMB_LEADER As String = String.Empty
    Dim m_DT_UpCode7 As DataTable = Nothing
    Dim m_sUpCode7 As String = String.Empty

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.m_sMB_AREA = "" & Request.QueryString("MB_AREA")

            Me.m_sMB_LEADER = "" & Request.QueryString("MB_LEADER")

            '所屬區代碼
            Me.m_sUpCode7 = com.Azion.EloanUtility.CodeList.getAppSettings("UpCode7")

            Me.Bind_TYPE_6()
        Catch ex As System.Threading.ThreadAbortException
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

    Sub Bind_TYPE_6()
        Dim DBManager As DatabaseManager = UIShareFun.getDataBaseManager
        Dim objStreamReader As System.IO.StreamReader = Nothing
        Dim DT_MB_VIP As DataTable = Nothing
        Try
            If System.IO.File.Exists(Server.MapPath(com.Azion.EloanUtility.UIUtility.getRootPath & "/QRY/MBQry_TYPE_01_06_v01.xml")) Then
                objStreamReader = New System.IO.StreamReader(Server.MapPath(com.Azion.EloanUtility.UIUtility.getRootPath & "/QRY/MBQry_TYPE_01_06_v01.xml"))
            Else
                Throw New Exception("EXCEL範本檔案(MBQry_TYPE_01_06_v01)不存在")
            End If

            Dim mbVIPList As New MB_VIPList(DBManager)
            If Not Utility.isValidateData(Me.m_sMB_AREA) AndAlso Not Utility.isValidateData(Me.m_sMB_LEADER) Then
                mbVIPList.EXCEL_1()
            ElseIf Utility.isValidateData(Me.m_sMB_AREA) AndAlso Not Utility.isValidateData(Me.m_sMB_LEADER) Then
                mbVIPList.EXCEL_2(Me.m_sMB_AREA)
            ElseIf Utility.isValidateData(Me.m_sMB_AREA) AndAlso Utility.isValidateData(Me.m_sMB_LEADER) Then
                mbVIPList.EXCEL_3(Me.m_sMB_AREA, Me.m_sMB_LEADER)
            ElseIf Not Utility.isValidateData(Me.m_sMB_AREA) AndAlso Utility.isValidateData(Me.m_sMB_LEADER) Then
                mbVIPList.EXCEL_4(Me.m_sMB_LEADER)
            End If

            DT_MB_VIP = mbVIPList.getCurrentDataSet.Tables(0)

            Dim sbCtlContent As New System.Text.StringBuilder(objStreamReader.ReadToEnd())

            Dim objStringBuilder As New System.Text.StringBuilder

            For Each ROW As DataRow In DT_MB_VIP.Rows
                Dim DT_MEMREV As DataTable = Nothing
                Try
                    Dim mbMEMREVList As New MB_MEMREVList(DBManager)
                    mbMEMREVList.LoadBySEQ(ROW("MB_MEMSEQ"))
                    DT_MEMREV = mbMEMREVList.getCurrentDataSet.Tables(0)
                    If IsNothing(DT_MEMREV) OrElse DT_MEMREV.Rows.Count = 0 Then
                        Dim ROW_NEW As DataRow = DT_MEMREV.NewRow
                        ROW_NEW("MB_MEMSEQ") = ROW("MB_MEMSEQ")
                        ROW_NEW("MB_AREA") = ROW("MB_AREA")
                        ROW_NEW("MB_LEADER") = ROW("MB_LEADER")
                        ROW_NEW("MB_NAME") = ROW("MB_NAME")

                        DT_MEMREV.Rows.Add(ROW_NEW)
                    End If

                    Dim sqlRow As DataRow = DT_MEMREV.NewRow
                    sqlRow("MB_MEMSEQ") = "-1"
                    sqlRow("MB_TOTFEE") = Utility.CheckNumNull(DT_MEMREV.Compute("SUM(MB_TOTFEE)", String.Empty))
                    DT_MEMREV.Rows.Add(sqlRow)

                    For Each ROW_MEMREV As DataRow In DT_MEMREV.Rows
                        objStringBuilder.Append("<Row ss:AutoFitHeight=""1"" ss:Height=""25"" >")

                        '所屬區
                        objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & Me.getMB_AREA(DBManager, ROW_MEMREV("MB_AREA").ToString) & "</Data></Cell>")

                        '委員
                        objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & ROW_MEMREV("MB_LEADER").ToString & "</Data></Cell>")

                        '會員編號
                        If ROW_MEMREV("MB_MEMSEQ") = "-1" Then
                            objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & "小計" & "</Data></Cell>")
                        Else
                            objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & Me.getMB_MEMSEQ(ROW_MEMREV("MB_MEMSEQ").ToString, ROW_MEMREV("MB_AREA").ToString) & "</Data></Cell>")
                        End If

                        '捐款人
                        objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & ROW_MEMREV("MB_NAME").ToString & "</Data></Cell>")

                        '繳款日期
                        Dim sMB_TX_DATE As String = String.Empty
                        sMB_TX_DATE = ROW_MEMREV("MB_TX_DATE").ToString
                        If IsDate(sMB_TX_DATE) Then
                            sMB_TX_DATE = Utility.DateTransfer(CDate(sMB_TX_DATE))
                        Else
                            sMB_TX_DATE = String.Empty
                        End If
                        objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & sMB_TX_DATE & "</Data></Cell>")

                        '繳款金額
                        Dim sMB_TOTFEE As String = String.Empty
                        If IsNumeric(ROW_MEMREV("MB_TOTFEE")) Then
                            sMB_TOTFEE = Utility.FormatDec(ROW_MEMREV("MB_TOTFEE"), "#,##0")
                        End If
                        objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & sMB_TOTFEE & "</Data></Cell>")

                        objStringBuilder.Append("</Row>")
                    Next
                Catch ex As Exception
                    Throw
                Finally
                    If Not IsNothing(DT_MEMREV) Then
                        DT_MEMREV.Dispose()
                    End If
                End Try
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

            If Not IsNothing(DT_MB_VIP) Then
                DT_MB_VIP.Dispose()
            End If
        End Try
    End Sub

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
End Class