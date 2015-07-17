Imports iTextSharp
Imports System.IO
Imports MBSC.MB_OP
Imports MBSC.UICtl
Imports com.Azion.NET.VB.DatabaseManager

Public Class MBMnt_Print_01_01_v02
    Inherits System.Web.UI.Page

    Dim m_sMB_MEMSEQ As String = String.Empty
    Dim m_sMB_SEQNO As String = String.Empty
    Dim m_sUpcode15 As String = String.Empty
    Dim DT_UPCODE_15 As DataTable = Nothing

    Dim m_DBManager As com.Azion.NET.VB.DatabaseManager = Nothing

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.m_sMB_MEMSEQ = "" & Request.QueryString("MB_MEMSEQ")
            Me.m_sMB_SEQNO = "" & Request.QueryString("MB_SEQNO")

            Me.m_DBManager = UIShareFun.getDataBaseManager

            '功德項目
            Me.m_sUpcode15 = com.Azion.EloanUtility.CodeList.getAppSettings("UpCode15")

            Dim mbMEMBER As New MB_MEMBER(Me.m_DBManager)
            mbMEMBER.loadByPK(CDec(Me.m_sMB_MEMSEQ))

            Dim mbMEMREV As New MB_MEMREV(Me.m_DBManager)
            mbMEMREV.loadByPK(CDec(Me.m_sMB_MEMSEQ), CDec(Me.m_sMB_SEQNO))

            '讀取列印PDF範本檔
            Dim sPdfTemplate As String = String.Empty
            sPdfTemplate = com.Azion.EloanUtility.UIUtility.getRootPath & "/MNT/MBSCPRT.pdf"
            Dim Pdfreader As iTextSharp.text.pdf.PdfReader = New iTextSharp.text.pdf.PdfReader(Server.MapPath(sPdfTemplate))

            '將列印PDF範本檔讀入MemoryStream
            Dim ba As New MemoryStream
            Dim Pdfstamper As iTextSharp.text.pdf.PdfStamper = New iTextSharp.text.pdf.PdfStamper(Pdfreader, ba)

            '設定中文字型
            Dim pdfFormFields As iTextSharp.text.pdf.AcroFields = Pdfstamper.AcroFields
            Dim baseFT As iTextSharp.text.pdf.BaseFont = iTextSharp.text.pdf.BaseFont.CreateFont("C:\Windows\Fonts\MSJHBD.ttc,1", iTextSharp.text.pdf.BaseFont.IDENTITY_H, iTextSharp.text.pdf.BaseFont.EMBEDDED)
            Dim chtFont As iTextSharp.text.Font = New iTextSharp.text.Font(baseFT, 24)

            Me.FillPDFForm(pdfFormFields, baseFT, "_1", mbMEMBER, mbMEMREV)
            Me.FillPDFForm(pdfFormFields, baseFT, "_2", mbMEMBER, mbMEMREV)

            Pdfstamper.FormFlattening = True
            Pdfstamper.Close()

            Response.Clear()
            Response.Buffer = True

            Response.Clear()
            Response.ClearContent()
            Response.ClearHeaders()

            Response.AddHeader("cache-control", "private")
            Response.AddHeader("P3P", "CP='CAO PSA OUR'")
            Response.ExpiresAbsolute = Now.AddDays(-1)
            Response.Expires = 0

            Response.Charset = "utf-8"
            Response.Buffer = True

            Response.ContentType = "application/pdf"
            Dim sPDFName As String = String.Empty
            If Utility.isValidateData(mbMEMREV.getAttribute("MB_RECNAME")) Then
                sPDFName = mbMEMREV.getString("MB_RECNAME") & ".pdf"
            Else
                sPDFName = mbMEMBER.getString("MB_NAME") & ".pdf"
            End If
            Response.AddHeader("Content-Disposition", "attachment;FileName=""" + System.Web.HttpUtility.UrlEncode(sPDFName) + """")
            Response.BinaryWrite(ba.ToArray)
            Response.Flush()
            Response.End()
        Catch ex As System.Threading.ThreadAbortException
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

    Private Sub MBMnt_Print_01_01_v02_Unload(sender As Object, e As EventArgs) Handles Me.Unload
        UIShareFun.releaseConnection(Me.m_DBManager)
    End Sub

    Sub FillPDFForm(ByVal pdfFormFields As iTextSharp.text.pdf.AcroFields, ByVal baseFT As iTextSharp.text.pdf.BaseFont, ByVal sFILLSEQ As String, ByVal mbMEMBER As MB_MEMBER, ByVal mbMEMREV As MB_MEMREV)
        Try
            '捐款人名稱
            pdfFormFields.SetFieldProperty("APPNAME" & sFILLSEQ, "textfont", baseFT, Nothing)
            If Utility.isValidateData(mbMEMREV.getAttribute("MB_RECNAME")) Then
                pdfFormFields.SetField("APPNAME" & sFILLSEQ, mbMEMREV.getString("MB_RECNAME"))
            Else
                pdfFormFields.SetField("APPNAME" & sFILLSEQ, mbMEMBER.getString("MB_NAME"))
            End If

            '繳款金額
            If IsNumeric(mbMEMREV.getAttribute("MB_TOTFEE")) AndAlso mbMEMREV.getDecimal("MB_TOTFEE") > 0 Then
                Dim i10000 As Decimal = Fix(mbMEMREV.getDecimal("MB_TOTFEE") / 10000)
                If i10000 > 0 Then
                    '捐款金額萬
                    pdfFormFields.SetFieldProperty("D10000" & sFILLSEQ, "textfont", baseFT, Nothing)
                    pdfFormFields.SetField("D10000" & sFILLSEQ, Utility.cypherConverCNum(i10000))
                End If

                Dim s1000 As String = String.Empty
                s1000 = Left(Right(mbMEMREV.getString("MB_TOTFEE"), 4), 1)
                If IsNumeric(s1000) AndAlso CDec(s1000) > 0 Then
                    '捐款金額千
                    pdfFormFields.SetFieldProperty("D1000" & sFILLSEQ, "textfont", baseFT, Nothing)
                    pdfFormFields.SetField("D1000" & sFILLSEQ, Utility.cypherConverCNum(CInt(s1000)))
                End If

                Dim s100 As String = String.Empty
                s100 = Left(Right(mbMEMREV.getString("MB_TOTFEE"), 3), 1)
                If IsNumeric(s100) AndAlso CDec(s100) > 0 Then
                    '捐款金額佰
                    pdfFormFields.SetFieldProperty("D100" & sFILLSEQ, "textfont", baseFT, Nothing)
                    pdfFormFields.SetField("D100" & sFILLSEQ, Utility.cypherConverCNum(CInt(s100)))
                End If

                Dim s10 As String = String.Empty
                s10 = Left(Right(mbMEMREV.getString("MB_TOTFEE"), 2), 1)
                If IsNumeric(s10) AndAlso CDec(s10) > 0 Then
                    '捐款金額十
                    pdfFormFields.SetFieldProperty("D10" & sFILLSEQ, "textfont", baseFT, Nothing)
                    pdfFormFields.SetField("D10" & sFILLSEQ, Utility.cypherConverCNum(CInt(s10)))
                End If

                Dim s1 As String = String.Empty
                s1 = Right(mbMEMREV.getString("MB_TOTFEE"), 1)
                If IsNumeric(s1) AndAlso CDec(s1) > 0 Then
                    '捐款金額元
                    pdfFormFields.SetFieldProperty("D1" & sFILLSEQ, "textfont", baseFT, Nothing)
                    pdfFormFields.SetField("D1" & sFILLSEQ, Utility.cypherConverCNum(CInt(s1)))
                End If
            End If

            '用途
            '功德項目
            pdfFormFields.SetFieldProperty("USE" & sFILLSEQ, "textfont", baseFT, Nothing)
            pdfFormFields.SetField("USE" & sFILLSEQ, Me.get_MB_ITEMID_TEXT(mbMEMREV.getString("MB_ITEMID")))

            '繳款日期
            Dim D_MB_TX_DATE As Object = Convert.DBNull
            D_MB_TX_DATE = Me.getDECIMAL_DATE(mbMEMREV.getAttribute("MB_TX_DATE"))
            If IsDate(D_MB_TX_DATE) AndAlso CDate(D_MB_TX_DATE).Year > 1911 Then
                '西元年
                pdfFormFields.SetFieldProperty("Y100" & sFILLSEQ, "textfont", baseFT, Nothing)
                pdfFormFields.SetField("Y100" & sFILLSEQ, CDate(D_MB_TX_DATE).Year)

                '月
                pdfFormFields.SetFieldProperty("Y10" & sFILLSEQ, "textfont", baseFT, Nothing)
                pdfFormFields.SetField("Y10" & sFILLSEQ, CDate(D_MB_TX_DATE).Month)

                '日
                pdfFormFields.SetFieldProperty("Y1" & sFILLSEQ, "textfont", baseFT, Nothing)
                pdfFormFields.SetField("Y1" & sFILLSEQ, CDate(D_MB_TX_DATE).Day)
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Public Function get_MB_ITEMID_TEXT(ByVal sMB_ITEMID As Object) As String
        Try
            If Utility.isValidateData(sMB_ITEMID) Then
                If IsNothing(Me.DT_UPCODE_15) Then
                    Dim apCODEList As New AP_CODEList(Me.m_DBManager)
                    apCODEList.loadByUpCode(Me.m_sUpcode15)
                    Me.DT_UPCODE_15 = apCODEList.getCurrentDataSet.Tables(0)
                End If

                Dim ROW_SELECT() As DataRow = Me.DT_UPCODE_15.Select("VALUE='" & sMB_ITEMID & "'")
                If Not IsNothing(ROW_SELECT) AndAlso ROW_SELECT.Length > 0 Then
                    Return ROW_SELECT(0)("TEXT").ToString()
                End If
            End If

            Return String.Empty
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function getDECIMAL_DATE(ByVal MB_TX_DATE As Object) As Object
        Try
            If Utility.isDecimalDate(MB_TX_DATE) Then
                Return CDate(Left(MB_TX_DATE.ToString, 4) & "/" & MB_TX_DATE.ToString.Substring(4, 2) & "/" & Right(MB_TX_DATE.ToString, 2))
            End If

            Return Convert.DBNull
        Catch ex As Exception
            Throw
        End Try
    End Function

End Class