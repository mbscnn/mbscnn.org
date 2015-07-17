Imports com.Azion.NET.VB
Imports MBSC.MB_OP
Imports MBSC.UICtl

Public Class MBQry_Member_01_v01
    Inherits System.Web.UI.Page

    Dim m_DBManager As DatabaseManager

    Dim m_sEXCEL As String = String.Empty
    Dim m_sMB_SEQ As String = String.Empty
    Dim m_sMB_BATCH As String = String.Empty

#Region "Page Events"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.m_DBManager = UIShareFun.getDataBaseManager

            Me.m_sEXCEL = "" & Request.QueryString("EXCEL")
            Me.m_sMB_SEQ = "" & Request.QueryString("MB_SEQ")
            Me.m_sMB_BATCH = "" & Request.QueryString("MB_BATCH")

            If Me.m_sEXCEL = "1" Then
                Me.doEXCEL(Me.m_sMB_SEQ, Me.m_sMB_BATCH)
            End If
        Catch ex As System.Threading.ThreadAbortException
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    Private Sub MBMnt_Class_01_v01_Unload(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Unload
        UIShareFun.releaseConnection(Me.m_DBManager)
    End Sub
#End Region

#Region "DateBind"
    'Bind Data
    Private Function bindData(ByVal sSEQ As String, ByVal sMB_BATCH As String) As Boolean
        Dim MB_MEMCLASSList As New MB_MEMCLASSList(m_DBManager)
        Dim MB_CLASS As New MB_CLASS(m_DBManager)
        Try
            If Not IsNumeric(sMB_BATCH) Then
                sMB_BATCH = "0"
            End If
            If IsNumeric(sSEQ) AndAlso IsNumeric(sMB_BATCH) AndAlso MB_MEMCLASSList.loadByMB_SEQ(CDec(sSEQ), CDec(sMB_BATCH)) > 0 Then
                If MB_CLASS.loadByPK(CInt(sSEQ), CDec(sMB_BATCH)) Then
                    lbl_CLASSNAME.Text = MB_CLASS.getString("MB_CLASS_NAME")
                    lbl_BATCH.Text = MB_CLASS.getString("MB_BATCH")
                End If
                dg_Member.DataSource = MB_MEMCLASSList.getCurrentDataSet.Tables(0).Select("ISNULL(MB_FWMK,' ')<>'3'")
                dg_Member.DataBind()

                divResult.Style.Item("display") = ""
                Return True
            Else
                dg_Member.DataSource = Nothing
                dg_Member.DataBind()
                divResult.Style.Item("display") = "none"
                Return False
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function

    Private Sub dg_Member_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles dg_Member.ItemDataBound
        If e.Item.ItemType = ListItemType.Item Or e.Item.ItemType = ListItemType.AlternatingItem Then
            Dim dItem As Object = e.Item.DataItem
            Dim MB_MEMBER As New MB_MEMBER(m_DBManager)

            Try
                If MB_MEMBER.loadByPK(CInt(dItem("MB_MEMSEQ").ToString)) Then
                    '報名人員
                    e.Item.Cells(0).Text = MB_MEMBER.getString("MB_NAME")

                    '課程備註
                    e.Item.Cells(2).Text = dItem("MB_MEMO").ToString

                    '正取備取
                    Dim mbCLASS As New MB_CLASS(Me.m_DBManager)
                    If mbCLASS.loadByPK(dItem("MB_SEQ").ToString, dItem("MB_BATCH").ToString) Then
                        'MB_FULL	decimal(3,0)		NO		0		select,insert,update,references	額滿人數
                        Dim iMB_FULL As Decimal = 0
                        iMB_FULL = mbCLASS.getDecimal("MB_FULL")
                        If IsNumeric(dItem("MB_SORTNO")) Then
                            If dItem("MB_SORTNO") <= iMB_FULL Then
                                e.Item.Cells(3).Text = "正取"
                            Else
                                e.Item.Cells(3).Text = "備取"
                            End If
                        End If
                    End If
                End If
            Catch ex As Exception
                UIShareFun.showErrMsg(Me, ex)
            End Try
        End If
    End Sub
#End Region

#Region "按鈕事件"
    '匯出Excel
    Protected Sub btn_QSEQ_Click(sender As Object, e As EventArgs) Handles btn_QSEQ.Click
        If Not Utility.isValidateData(Me.MB_BATCH.SelectedValue) Then
            com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "請選擇梯次")
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請選擇梯次")
            Return
        End If
        If Not bindData(txt_MB_SEQ.Text, Me.MB_BATCH.SelectedValue) Then Exit Sub
        Try
            'Dim st As New System.IO.StringWriter
            'Dim htw As New UI.HtmlTextWriter(st)

            'bindData(txt_MB_SEQ.Text)
            ''匯出Excel
            'Response.Clear()
            'Response.ContentType = "application/ms-excel"
            'Response.ContentEncoding = System.Text.Encoding.Default
            'Response.AddHeader("content-disposition", "attachment; filename=file.xls")

            'tblTitle.RenderControl(htw) '寫入Title
            'dg_Member.BorderWidth = Unit.Pixel(1)
            'dg_Member.RenderControl(htw) '寫入DataGrid
            'Response.Write(st.ToString.Replace("<br>", ""))
            'Response.Flush()
            'Response.Close()

            Dim sURL As String = String.Empty

            sURL = com.Azion.EloanUtility.UIUtility.getRootPath & "/QRY/MBQry_Member_01_v01.aspx?EXCEL=1&MB_SEQ=" & Me.txt_MB_SEQ.Text & "&MB_BATCH=" & Me.MB_BATCH.SelectedValue

            com.Azion.EloanUtility.UIUtility.showOpen(sURL)
        Catch ex As Exception
        End Try
    End Sub

    Sub doEXCEL(ByVal sMB_SEQ As String, ByVal sMB_BATCH As String)
        Dim st As New System.IO.StringWriter
        Dim htw As New UI.HtmlTextWriter(st)

        bindData(sMB_SEQ, sMB_BATCH)
        '匯出Excel
        Response.Clear()
        Response.ContentType = "application/ms-excel"
        Response.ContentEncoding = System.Text.Encoding.Default
        Response.AddHeader("content-disposition", "attachment; filename=file.xls")

        tblTitle.RenderControl(htw) '寫入Title
        dg_Member.BorderWidth = Unit.Pixel(1)
        dg_Member.RenderControl(htw) '寫入DataGrid
        Response.Write(st.ToString.Replace("<br>", ""))
        Response.Flush()
        Response.Close()
    End Sub

    '查詢
    Protected Sub btn_Confirm_Click(sender As Object, e As EventArgs) Handles btn_Confirm.Click
        Try
            If Not IsNumeric(Me.MB_BATCH.SelectedValue) Then
                com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "請選擇梯次")
                com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請選擇梯次")
                Return
            End If

            If Not bindData(txt_MB_SEQ.Text, Me.MB_BATCH.SelectedValue) Then
                txt_MB_SEQ.Text = ""
                btn_QSEQ.Visible = False
                UIShareFun.showErrMsg(Me, "查無資料")
            Else
                btn_QSEQ.Visible = True
            End If
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub
#End Region

    Private Sub txt_MB_SEQ_TextChanged(sender As Object, e As EventArgs) Handles txt_MB_SEQ.TextChanged
        Dim DT_MB_BATCH As DataTable = Nothing
        Try
            Me.MB_BATCH.Items.Clear()
            If IsNumeric(Me.txt_MB_SEQ.Text) Then
                Dim mbCLASSList As New MB_CLASSList(Me.m_DBManager)
                mbCLASSList.LoadByMB_SEQ(Me.txt_MB_SEQ.Text)

                DT_MB_BATCH = New DataView(mbCLASSList.getCurrentDataSet.Tables(0), String.Empty, "MB_BATCH", DataViewRowState.CurrentRows).ToTable(True, "MB_BATCH")

                For Each ROW As DataRow In DT_MB_BATCH.Rows
                    Dim ListItem As New ListItem
                    If ROW("MB_BATCH") <= 0 Then
                        ListItem.Text = "無梯次"
                    Else
                        ListItem.Text = "第" & ROW("MB_BATCH") & "梯次"
                    End If
                    ListItem.Value = ROW("MB_BATCH")

                    Me.MB_BATCH.Items.Add(ListItem)
                Next

                If Me.MB_BATCH.Items.Count > 1 Then
                    Me.MB_BATCH.Items.Insert(0, New ListItem("請選擇", String.Empty))
                End If
            End If
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        Finally
            If Not IsNothing(DT_MB_BATCH) Then
                DT_MB_BATCH.Dispose()
            End If
        End Try
    End Sub
End Class