Imports MBSC.MB_OP
Imports com.Azion.NET.VB
Imports MBSC.UICtl

Public Class MBMnt_Print_01_01_v01_aspx
    Inherits System.Web.UI.Page

    Dim m_DBManager As com.Azion.NET.VB.DatabaseManager

    Dim m_sUpcode15 As String = String.Empty
    Dim DT_UPCODE_15 As DataTable

    Dim m_sUpCode28 As String = String.Empty
    Dim DT_UPCODE_28 As DataTable

    Dim m_sUpcode23 As String = String.Empty
    Dim DT_UPCODE_23 As DataTable

    Dim m_sUpCode7 As String = String.Empty
    Dim DT_UPCODE_7 As DataTable

    Dim m_sUpcode31 As String = String.Empty
    Dim DT_UPCODE_31 As DataTable

    Dim m_sMB_MEMSEQ As String = String.Empty

    Dim m_sMB_SEQNO As String = String.Empty

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        m_DBManager = UIShareFun.getDataBaseManager

        '功德項目
        Me.m_sUpcode15 = com.Azion.EloanUtility.CodeList.getAppSettings("UpCode15")

        '會員類別
        Me.m_sUpCode28 = com.Azion.EloanUtility.CodeList.getAppSettings("UpCode28")

        '繳款方式
        Me.m_sUpcode23 = com.Azion.EloanUtility.CodeList.getAppSettings("UpCode23")

        '所屬區代碼
        Me.m_sUpCode7 = com.Azion.EloanUtility.CodeList.getAppSettings("UpCode7")

        '付款方式
        Me.m_sUpcode31 = com.Azion.EloanUtility.CodeList.getAppSettings("UpCode31")

        Me.m_sMB_MEMSEQ = "" & Request.QueryString("MB_MEMSEQ")

        Me.m_sMB_SEQNO = "" & Request.QueryString("MB_SEQNO")

        If IsNumeric(Me.m_sMB_MEMSEQ) AndAlso IsNumeric(Me.m_sMB_SEQNO) Then
            Me.FILL_FORM(CDec(Me.m_sMB_MEMSEQ), CDec(Me.m_sMB_SEQNO))
        End If

        Me.OutToWord()
    End Sub

    Sub OutToWord()
        Try
            Response.Clear()
            Response.Buffer = True
            Response.Charset = "big5"

            Page.EnableViewState = False

            '下面這行很重要， attachment 參數表示作為附件下載，您可以改成 online在線打開 
            'filename=PrintFile.doc 指定輸出文件的名稱，注意其附檔名和指定文件類型相符，
            '可以為：.doc .xls .txt .htm
            'Response.AppendHeader("Content-Disposition", "attachment;filename=PrintFile.doc")
            Dim sValue As String = "inline;filename=" & Me.m_sMB_MEMSEQ & Me.m_sMB_SEQNO & ".doc"
            'Dim sValue As String = "attachment;filename=" & _PrintLogic.mENCASEID & ".doc"
            Response.AppendHeader("Content-Disposition", sValue)

            Response.ContentEncoding = System.Text.Encoding.GetEncoding("big5")
            Response.ContentType = "application/ms-word"

            Dim TW As New System.IO.StringWriter
            Dim HW As New HtmlTextWriter(TW)

            Me.Render(HW)
            Response.Write(com.Azion.NET.VB.Titan.Utility.transferUnicodeToNCR(TW.ToString))
            Response.End()
        Catch ex As System.Threading.ThreadAbortException

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub MBMnt_Print_01_01_v01_aspx_Unload(sender As Object, e As System.EventArgs) Handles Me.Unload
        UIShareFun.releaseConnection(m_DBManager)
    End Sub

    Sub FILL_FORM(ByVal iMB_MEMSEQ As Decimal, ByVal iMB_SEQNO As Decimal)
        Try
            Dim mbMEMBER As New MB_MEMBER(Me.m_DBManager)
            mbMEMBER.loadByPK(iMB_MEMSEQ)

            Dim mbMEMREV As New MB_MEMREV(Me.m_DBManager)
            mbMEMREV.loadByPK(iMB_MEMSEQ, iMB_SEQNO)

            '法名/姓名
            Me.MB_NAME.Text = Trim(mbMEMBER.getString("MB_NAME"))

            '會員編號
            Me.MB_MEMSEQ.Text = Me.getMB_MEMSEQ(iMB_MEMSEQ, mbMEMBER.getString("MB_AREA"))

            '通訊地址
            Me.MB_ADDR.Text = mbMEMBER.getString("MB_CITY") & mbMEMBER.getString("MB_VLG") & mbMEMBER.getString("MB_ADDR")

            '繳款日期
            If Utility.isDecimalDate(mbMEMREV.getAttribute("MB_TX_DATE")) Then
                Me.MB_TX_DATE.Text = Me.get_DECIMAL_DATE(mbMEMREV.getAttribute("MB_TX_DATE"))
            Else
                Me.MB_TX_DATE.Text = String.Empty
            End If

            '功德項目
            Me.MB_ITEMID.Text = Me.get_MB_ITEMID_TEXT(mbMEMREV.getString("MB_ITEMID"))
            If mbMEMREV.getString("MB_ITEMID") = "A" Then
                Me.PLH_MB_MEMTYP.Visible = True
                Me.MB_MEMTYP.Text = Me.get_MB_MEMTYP_TEXT(mbMEMREV.getString("MB_MEMTYP"))
            Else
                Me.PLH_MB_MEMTYP.Visible = False
                Me.MB_MEMTYP.Text = String.Empty
            End If

            '繳款方式
            Me.MB_FEETYPE.Text = Me.get_MB_FEETYPE_TEXT(mbMEMREV.getString("MB_FEETYPE"))

            '繳款金額
            If IsNumeric(mbMEMREV.getAttribute("MB_TOTFEE")) Then
                Me.MB_TOTFEE.Text = Utility.FormatDec(mbMEMREV.getAttribute("MB_TOTFEE"), "#,##0")
            Else
                Me.MB_TOTFEE.Text = String.Empty
            End If
            '每月金額
            If mbMEMREV.getString("MB_FEETYPE") <> "D" Then
                Me.PLH_MB_TOTFEE_MM.Visible = True
                Dim sNOTE As String = String.Empty
                sNOTE = Me.get_MB_FEETYPE_NOTE(mbMEMREV.getString("MB_FEETYPE"))
                If IsNumeric(sNOTE) Then
                    Me.MB_TOTFEE_MM.Text = Utility.FormatDec(sNOTE, "#,##0")
                Else
                    Me.MB_TOTFEE_MM.Text = String.Empty
                End If
            Else
                Me.PLH_MB_TOTFEE_MM.Visible = False
                Me.MB_TOTFEE_MM.Text = String.Empty
            End If

            '收據捐款名稱
            Me.MB_RECNAME.Text = mbMEMREV.getString("MB_RECNAME")

            '所屬區
            Me.MB_AREA.Text = Me.getMB_AREA(mbMEMBER.getString("MB_AREA"))

            '所屬區委員
            Me.MB_LEADER.Text = mbMEMBER.getString("MB_LEADER")

            '會費期間
            If mbMEMREV.getString("MB_ITEMID") = "A" Then
                Me.MB_DESC.Visible = False
                Me.MB_DESC.Text = String.Empty

                Me.MB_MEMFEE.Visible = True
                Me.MB_MEMFEE.Text = (mbMEMREV.getDecimal("MB_MEMFEE_SY")) & "年" & mbMEMREV.getString("MB_MEMFEE_SM") & "月至" & _
                                    (mbMEMREV.getDecimal("MB_MEMFEE_EY")) & "年" & mbMEMREV.getString("MB_MEMFEE_EM") & "月"
            Else
                Me.MB_MEMFEE.Visible = False
                Me.MB_MEMFEE.Text = String.Empty

                Me.MB_DESC.Visible = True
                Me.MB_DESC.Text = mbMEMREV.getString("MB_DESC")
            End If

            '付款方式
            Me.MB_PAY_TYPE.Text = Me.get_MB_PAY_TYPE_TEXT(mbMEMREV.getString("MB_PAY_TYPE"))

            '票據到期日
            If Utility.isDecimalDate(mbMEMREV.getAttribute("NOTE_DUE_DATE")) Then
                Me.NOTE_DUE_DATE.Text = Me.get_DECIMAL_DATE(mbMEMREV.getAttribute("NOTE_DUE_DATE"))
            Else
                Me.NOTE_DUE_DATE.Text = String.Empty
            End If

            '票據號碼
            Me.NOTE_NO.Text = mbMEMREV.getString("NOTE_NO")

            '發票行銀行
            Me.NOTE_BANK.Text = Trim(mbMEMREV.getString("NOTE_BANK"))

            '發票行分行
            Me.NOTE_BR.Text = Trim(mbMEMREV.getString("NOTE_BR"))

            '發票人
            Me.NOTE_HOLDER.Text = Trim(mbMEMREV.getString("NOTE_HOLDER"))

            '票據金額
            Me.NOTE_AMT.Text = Utility.FormatDec(mbMEMREV.getDecimal("NOTE_AMT"), "#,##0")

            '補印次數
            Me.MB_REISU.Text = mbMEMREV.getString("MB_REISU")
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

#Region "Utility"
    Public Function getMB_MEMSEQ(ByVal iMB_MEMSEQ As Object, ByVal sMB_AREA As Object) As String
        Try
            If IsNumeric(iMB_MEMSEQ) AndAlso Utility.isValidateData(sMB_AREA) Then
                Return sMB_AREA & "-" & com.Azion.EloanUtility.StrUtility.FillZero(CDec(iMB_MEMSEQ), 7)
            End If
            Return String.Empty
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function get_DECIMAL_DATE(ByVal MB_TX_DATE As Object) As String
        Try
            If Utility.isDecimalDate(MB_TX_DATE) Then
                Return Utility.DateTransfer(Left(MB_TX_DATE.ToString, 4) & "/" & MB_TX_DATE.ToString.Substring(4, 2) & "/" & Right(MB_TX_DATE.ToString, 2))
            End If

            Return String.Empty
        Catch ex As Exception
            Throw ex
        End Try
    End Function

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

    Public Function get_MB_MEMTYP_TEXT(ByVal sMB_MEMTYP As Object) As String
        Try
            If Utility.isValidateData(sMB_MEMTYP) Then
                If IsNothing(Me.DT_UPCODE_28) Then
                    Dim apCODEList As New AP_CODEList(Me.m_DBManager)
                    apCODEList.loadByUpCode(Me.m_sUpCode28)
                    Me.DT_UPCODE_28 = apCODEList.getCurrentDataSet.Tables(0)
                End If

                Dim ROW_SELECT() As DataRow = Me.DT_UPCODE_28.Select("VALUE='" & sMB_MEMTYP & "'")
                If Not IsNothing(ROW_SELECT) AndAlso ROW_SELECT.Length > 0 Then
                    Return ROW_SELECT(0)("TEXT").ToString()
                End If
            End If

            Return String.Empty
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function get_MB_FEETYPE_TEXT(ByVal sMB_FEETYPE As Object) As String
        Try
            If Utility.isValidateData(sMB_FEETYPE) Then
                If IsNothing(Me.DT_UPCODE_23) Then
                    Dim apCODEList As New AP_CODEList(Me.m_DBManager)
                    apCODEList.loadByUpCode(Me.m_sUpcode23)
                    Me.DT_UPCODE_23 = apCODEList.getCurrentDataSet.Tables(0)
                End If

                Dim ROW_SELECT() As DataRow = Me.DT_UPCODE_23.Select("VALUE='" & sMB_FEETYPE & "'")
                If Not IsNothing(ROW_SELECT) AndAlso ROW_SELECT.Length > 0 Then
                    Return ROW_SELECT(0)("TEXT").ToString()
                End If
            End If

            Return String.Empty
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function get_MB_FEETYPE_NOTE(ByVal sMB_FEETYPE As Object) As String
        Try
            If Utility.isValidateData(sMB_FEETYPE) Then
                If IsNothing(Me.DT_UPCODE_23) Then
                    Dim apCODEList As New AP_CODEList(Me.m_DBManager)
                    apCODEList.loadByUpCode(Me.m_sUpcode23)
                    Me.DT_UPCODE_23 = apCODEList.getCurrentDataSet.Tables(0)
                End If

                Dim ROW_SELECT() As DataRow = Me.DT_UPCODE_23.Select("VALUE='" & sMB_FEETYPE & "'")
                If Not IsNothing(ROW_SELECT) AndAlso ROW_SELECT.Length > 0 Then
                    Return ROW_SELECT(0)("NOTE").ToString()
                End If
            End If

            Return String.Empty
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function getMB_AREA(ByVal sMB_AREA As Object) As String
        Try
            If Utility.isValidateData(sMB_AREA) Then
                If IsNothing(Me.DT_UPCODE_7) Then
                    Dim apCODEList As New AP_CODEList(Me.m_DBManager)
                    apCODEList.loadByUpCode(Me.m_sUpCode7)
                    Me.DT_UPCODE_7 = apCODEList.getCurrentDataSet.Tables(0)
                End If

                Dim ROW_SELECT() As DataRow = Me.DT_UPCODE_7.Select("VALUE='" & sMB_AREA & "'")
                If Not IsNothing(ROW_SELECT) AndAlso ROW_SELECT.Length > 0 Then
                    Return ROW_SELECT(0)("TEXT").ToString()
                End If
            End If

            Return String.Empty
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function get_MB_PAY_TYPE_TEXT(ByVal sMB_PAY_TYPE As Object) As String
        Try
            If Utility.isValidateData(sMB_PAY_TYPE) Then
                If IsNothing(Me.DT_UPCODE_31) Then
                    Dim apCODEList As New AP_CODEList(Me.m_DBManager)
                    apCODEList.loadByUpCode(Me.m_sUpcode31)
                    Me.DT_UPCODE_31 = apCODEList.getCurrentDataSet.Tables(0)
                End If

                Dim ROW_SELECT() As DataRow = Me.DT_UPCODE_31.Select("VALUE='" & sMB_PAY_TYPE & "'")
                If Not IsNothing(ROW_SELECT) AndAlso ROW_SELECT.Length > 0 Then
                    Return ROW_SELECT(0)("TEXT").ToString()
                End If
            End If

            Return String.Empty
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

End Class