Imports MBSC.MB_OP
Imports MBSC.UICtl
Imports com.Azion.NET.VB

Public Class MBQry_TYPE_01_05_v01
    Inherits System.Web.UI.Page

    Dim m_sBY As String = String.Empty
    Dim m_sEY As String = String.Empty
    Dim m_sMB_AREA As String = String.Empty
    Dim m_sMB_LEADER As String = String.Empty
    Dim m_DT_UpCode7 As DataTable = Nothing
    Dim m_sUpCode7 As String = String.Empty

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.m_sBY = "" & Request.QueryString("BY")
            Me.m_sEY = "" & Request.QueryString("EY")
            Me.m_sMB_AREA = "" & Request.QueryString("MB_AREA")
            Me.m_sMB_LEADER = "" & Request.QueryString("MB_LEADER")

            '所屬區代碼
            Me.m_sUpCode7 = com.Azion.EloanUtility.CodeList.getAppSettings("UpCode7")
            Me.Bind_TYPE_5()
        Catch ex As System.Threading.ThreadAbortException
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

    Sub Bind_TYPE_5()
        Dim DBManager As DatabaseManager = UIShareFun.getDataBaseManager
        Dim objStreamReader As System.IO.StreamReader = Nothing
        Dim DT_MB_MEMFEE As DataTable = Nothing
        Try
            If System.IO.File.Exists(Server.MapPath(com.Azion.EloanUtility.UIUtility.getRootPath & "/QRY/MBQry_TYPE_01_05_v01.xml")) Then
                objStreamReader = New System.IO.StreamReader(Server.MapPath(com.Azion.EloanUtility.UIUtility.getRootPath & "/QRY/MBQry_TYPE_01_05_v01.xml"))
            Else
                Throw New Exception("EXCEL範本檔案(MBQry_TYPE_01_02_v01)不存在")
            End If

            Dim isPeriod As Boolean = False
            If IsNumeric(Me.m_sBY) AndAlso IsNumeric(Me.m_sEY) Then
                isPeriod = True
            End If

            Dim mbMEMFEEList As New MB_MEMFEEList(DBManager)
            If Not isPeriod AndAlso Not Utility.isValidateData(Me.m_sMB_AREA) AndAlso Not Utility.isValidateData(Me.m_sMB_LEADER) Then
                mbMEMFEEList.QRY_EXCEL_1()
            ElseIf isPeriod AndAlso Not Utility.isValidateData(Me.m_sMB_AREA) AndAlso Not Utility.isValidateData(Me.m_sMB_LEADER) Then
                mbMEMFEEList.QRY_EXCEL_2(CDec(Me.m_sBY), CDec(Me.m_sEY))
            ElseIf isPeriod AndAlso Utility.isValidateData(Me.m_sMB_AREA) AndAlso Not Utility.isValidateData(Me.m_sMB_LEADER) Then
                mbMEMFEEList.QRY_EXCEL_3(CDec(Me.m_sBY), CDec(Me.m_sEY), Me.m_sMB_AREA)
            ElseIf isPeriod AndAlso Utility.isValidateData(Me.m_sMB_AREA) AndAlso Utility.isValidateData(Me.m_sMB_LEADER) Then
                mbMEMFEEList.QRY_EXCEL_4(CDec(Me.m_sBY), CDec(Me.m_sEY), Me.m_sMB_AREA, Me.m_sMB_LEADER)
            ElseIf Not isPeriod AndAlso Utility.isValidateData(Me.m_sMB_AREA) AndAlso Utility.isValidateData(Me.m_sMB_LEADER) Then
                mbMEMFEEList.QRY_EXCEL_5(Me.m_sMB_AREA, Me.m_sMB_LEADER)
            ElseIf Not isPeriod AndAlso Utility.isValidateData(Me.m_sMB_AREA) AndAlso Not Utility.isValidateData(Me.m_sMB_LEADER) Then
                mbMEMFEEList.QRY_EXCEL_6(Me.m_sMB_AREA)
            ElseIf Not isPeriod AndAlso Not Utility.isValidateData(Me.m_sMB_AREA) AndAlso Utility.isValidateData(Me.m_sMB_LEADER) Then
                mbMEMFEEList.QRY_EXCEL_7(Me.m_sMB_LEADER)
            End If

            DT_MB_MEMFEE = mbMEMFEEList.getCurrentDataSet.Tables(0)

            Dim sbCtlContent As New System.Text.StringBuilder(objStreamReader.ReadToEnd())

            Dim objStringBuilder As New System.Text.StringBuilder

            For Each ROW As DataRow In DT_MB_MEMFEE.Rows
                objStringBuilder.Append("<Row ss:AutoFitHeight=""1"" ss:Height=""50"" >")

                '所屬區
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & Me.getMB_AREA(DBManager, ROW("MB_AREA").ToString) & "</Data></Cell>")

                '委員
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & ROW("MB_LEADER").ToString & "</Data></Cell>")

                '會員編號
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & Me.getMB_MEMSEQ(ROW("MB_MEMSEQ").ToString, ROW("MB_AREA").ToString) & "</Data></Cell>")

                '捐款人
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & ROW("MB_NAME").ToString & "</Data></Cell>")

                '繳款年度
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & ROW("MB_YYYY").ToString & "</Data></Cell>")

                '1月
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & Me.getFEEAmt(ROW("MB_M1")) & "</Data></Cell>")

                '2月
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & Me.getFEEAmt(ROW("MB_M2")) & "</Data></Cell>")

                '3月
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & Me.getFEEAmt(ROW("MB_M3")) & "</Data></Cell>")

                '4月
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & Me.getFEEAmt(ROW("MB_M4")) & "</Data></Cell>")

                '5月
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & Me.getFEEAmt(ROW("MB_M5")) & "</Data></Cell>")

                '6月
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & Me.getFEEAmt(ROW("MB_M6")) & "</Data></Cell>")

                '7月
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & Me.getFEEAmt(ROW("MB_M7")) & "</Data></Cell>")

                '8月
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & Me.getFEEAmt(ROW("MB_M8")) & "</Data></Cell>")

                '9月
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & Me.getFEEAmt(ROW("MB_M9")) & "</Data></Cell>")

                '10月
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & Me.getFEEAmt(ROW("MB_M10")) & "</Data></Cell>")

                '11月
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & Me.getFEEAmt(ROW("MB_M11")) & "</Data></Cell>")

                '12月
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & Me.getFEEAmt(ROW("MB_M12")) & "</Data></Cell>")

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
            UIShareFun.releaseConnection(DBManager)

            If Not IsNothing(objStreamReader) Then
                objStreamReader.Close()
            End If

            If Not IsNothing(DT_MB_MEMFEE) Then
                DT_MB_MEMFEE.Dispose()
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

    Protected Function getFEEAmt(ByVal iAMT As Object) As String
        Try
            If IsNumeric(iAMT) Then
                Return Format(CDec(iAMT), "#,##0")
            Else
                Return "0"
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function
End Class