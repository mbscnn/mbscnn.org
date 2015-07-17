Imports MBSC.MB_OP
Imports MBSC.UICtl
Imports com.Azion.NET.VB

Public Class MBQry_TYPE_01_01_v01
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
    Dim m_DT_UpCode28 As DataTable = Nothing
    Dim m_sUpCode28 As String = String.Empty

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

            '會員類別
            Me.m_sUpCode28 = com.Azion.EloanUtility.CodeList.getAppSettings("UpCode28")

            Me.Bind_TYPE_1()
        Catch ex As System.Threading.ThreadAbortException
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

    Sub Bind_TYPE_1()
        Dim DBManager As DatabaseManager = UIShareFun.getDataBaseManager
        Dim objStreamReader As System.IO.StreamReader = Nothing
        Dim DT_MEMBER As DataTable = Nothing
        Try
            Dim D_BDay As Object = Convert.DBNull
            D_BDay = Me.ConvertWDate(Me.m_sBY, Me.m_sBM, Me.m_sBD)
            Dim D_EDay As Object = Convert.DBNull
            D_EDay = Me.ConvertWDate(Me.m_sEY, Me.m_sEM, Me.m_sED)

            Dim isPeriod As Boolean = False
            If IsDate(D_BDay) AndAlso IsDate(D_EDay) Then
                isPeriod = True
            End If

            Dim mbMEMBERList As New MB_MEMBERList(DBManager)
            If Not isPeriod AndAlso Not Utility.isValidateData(Me.m_sMB_AREA) AndAlso Not Utility.isValidateData(Me.m_sMB_LEADER) Then
                mbMEMBERList.QRY_EXCEL_1()
            ElseIf isPeriod AndAlso Not Utility.isValidateData(Me.m_sMB_AREA) AndAlso Not Utility.isValidateData(Me.m_sMB_LEADER) Then
                mbMEMBERList.QRY_EXCEL_2(D_BDay, D_EDay)
            ElseIf isPeriod AndAlso Utility.isValidateData(Me.m_sMB_AREA) AndAlso Not Utility.isValidateData(Me.m_sMB_LEADER) Then
                mbMEMBERList.QRY_EXCEL_3(D_BDay, D_EDay, Me.m_sMB_AREA)
            ElseIf isPeriod AndAlso Not Utility.isValidateData(Me.m_sMB_AREA) AndAlso Utility.isValidateData(Me.m_sMB_LEADER) Then
                mbMEMBERList.QRY_EXCEL_4(D_BDay, D_EDay, Me.m_sMB_LEADER)
            ElseIf isPeriod AndAlso Utility.isValidateData(Me.m_sMB_AREA) AndAlso Utility.isValidateData(Me.m_sMB_LEADER) Then
                mbMEMBERList.QRY_EXCEL_5(D_BDay, D_EDay, Me.m_sMB_AREA, Me.m_sMB_LEADER)
            ElseIf Not isPeriod AndAlso Utility.isValidateData(Me.m_sMB_AREA) AndAlso Not Utility.isValidateData(Me.m_sMB_LEADER) Then
                mbMEMBERList.QRY_EXCEL_6(Me.m_sMB_AREA)
            ElseIf Not isPeriod AndAlso Utility.isValidateData(Me.m_sMB_AREA) AndAlso Utility.isValidateData(Me.m_sMB_LEADER) Then
                mbMEMBERList.QRY_EXCEL_7(Me.m_sMB_AREA, Me.m_sMB_LEADER)
            End If
            DT_MEMBER = mbMEMBERList.getCurrentDataSet.Tables(0)

            If System.IO.File.Exists(Server.MapPath(com.Azion.EloanUtility.UIUtility.getRootPath & "/QRY/MBQry_TYPE_01_01_v01.xml")) Then
                objStreamReader = New System.IO.StreamReader(Server.MapPath(com.Azion.EloanUtility.UIUtility.getRootPath & "/QRY/MBQry_TYPE_01_01_v01.xml"))
            Else
                Throw New Exception("EXCEL範本檔案(MBQry_TYPE_01_01_v01)不存在")
            End If

            Dim sbCtlContent As New System.Text.StringBuilder(objStreamReader.ReadToEnd())

            Dim objStringBuilder As New System.Text.StringBuilder

            For Each ROW As DataRow In DT_MEMBER.Rows
                objStringBuilder.Append("<Row ss:AutoFitHeight=""1"" ss:Height=""50"" >")

                '所屬區
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & Me.getMB_AREA(DBManager, ROW("MB_AREA").ToString) & "</Data></Cell>")

                '委員
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & ROW("MB_LEADER").ToString & "</Data></Cell>")

                '會員姓名
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & ROW("MB_NAME").ToString & "</Data></Cell>")

                '會員編號
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & Me.getMB_MEMSEQ(ROW("MB_MEMSEQ").ToString, ROW("MB_AREA").ToString) & "</Data></Cell>")

                '會員類別
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & Me.getMB_FAMILY(DBManager, ROW("MB_FAMILY").ToString) & "</Data></Cell>")

                '電話
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & ROW("MB_TEL").ToString & "</Data></Cell>")

                '手機
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & ROW("MB_MOBIL").ToString & "</Data></Cell>")

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
            If Not IsNothing(DT_MEMBER) Then
                DT_MEMBER.Dispose()
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

    Public Function getMB_AREA(ByVal DBManager As DatabaseManager, ByVal sMB_AREA As Object) As String
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

    Public Function getMB_FAMILY(ByVal DBManager As DatabaseManager, ByVal sMB_FAMILY As Object) As String
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

End Class