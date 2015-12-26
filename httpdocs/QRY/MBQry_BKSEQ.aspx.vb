Imports MBSC.MB_OP
Imports com.Azion.NET.VB
Imports MBSC.UICtl
Imports com.Azion.EloanUtility

Public Class MBQry_BKSEQ
    Inherits System.Web.UI.Page

    Dim m_DBManager As com.Azion.NET.VB.DatabaseManager

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.m_DBManager = UIShareFun.getDataBaseManager
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

    Private Sub MBQry_BKSEQ_Unload(sender As Object, e As EventArgs) Handles Me.Unload
        UIShareFun.releaseConnection(Me.m_DBManager)
    End Sub

    Private Sub btn_Qry_Click(sender As Object, e As EventArgs) Handles btn_Qry.Click
        Try
            If Me.rbt_Cel.Checked = False AndAlso Me.rbt_Tel.Checked = False AndAlso Me.rbt_Name.Checked = False Then
                com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "請選擇查詢方式『手機』或『電話』或『姓名』")
                com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請選擇查詢方式『手機』或『電話』或『姓名』")
                Return
            End If

            If Me.rbt_Cel.Checked Then
                If Not Utility.isValidateData(Me.TXT_MB_MOBIL.Text) Then
                    com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "請輸入手機號碼")
                    com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請輸入手機號碼")
                    Return
                End If

                If Not IsNumeric(TXT_MB_MOBIL.Text) Then
                    com.Azion.EloanUtility.UIUtility.alert("手機號碼應為10個數字【範例:0933123456】，請檢查")
                    com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "手機號碼應為10個數字【範例:0933123456】，請檢查")
                    Return
                End If
            End If

            If Me.rbt_Tel.Checked Then
                If Not Utility.isValidateData(Me.TXT_MB_TEL_ZIP.Text) Then
                    com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "請輸入區碼")
                    com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請輸入區碼")
                    Return
                End If

                If Not IsNumeric(Me.TXT_MB_TEL_ZIP.Text) Then
                    com.Azion.EloanUtility.UIUtility.alert("區碼應為數字【範例:02】，請檢查")
                    com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "區碼應為數字【範例:02】，請檢查")
                    Return
                End If

                If Not Utility.isValidateData(Me.TXT_MB_TEL.Text) Then
                    com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "請輸入電話號碼")
                    com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請輸入電話號碼")
                    Return
                End If

                If Not IsNumeric(Me.TXT_MB_TEL.Text) Then
                    com.Azion.EloanUtility.UIUtility.alert("電話號碼應為數字【範例:23627968】，請檢查")
                    com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "電話號碼應為數字【範例:23627968】，請檢查")
                    Return
                End If
            End If

            If Me.rbt_Name.Checked Then
                If Not Utility.isValidateData(Me.TXT_MB_NAME.Text) Then
                    com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "請輸入姓名")
                    com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請輸入姓名")
                    Return
                End If

                Dim pattern As String = "[\p{P}\p{S}-[._]]"
                Dim mx As MatchCollection = Regex.Matches(Me.TXT_MB_NAME.Text.Trim, pattern)
                If mx.Count > 0 Then
                    com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "姓名不可輸入標點符號")
                    com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "姓名不可輸入標點符號")
                    Return
                End If

                pattern = "[0-9]"
                Dim mxn As MatchCollection = Regex.Matches(Me.TXT_MB_NAME.Text.Trim, pattern)
                If mxn.Count > 0 Then
                    com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "姓名不可輸入數字")
                    com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "姓名不可輸入數字")
                    Return
                End If
            End If

            If Me.rbt_Cel.Checked Then
                '手機
                Dim MB_MEMBERList As New MB_MEMBERList(Me.m_DBManager)
                MB_MEMBERList.Load_BKSEQ_MB_MOBIL(Me.TXT_MB_MOBIL.Text)
                Me.Bind_MEMBER(MB_MEMBERList.getCurrentDataSet.Tables(0))
            ElseIf Me.rbt_Tel.Checked Then
                '電話
                Dim MB_MEMBERList As New MB_MEMBERList(Me.m_DBManager)
                MB_MEMBERList.Load_BKSEQ_MB_TEL(Me.TXT_MB_TEL_ZIP.Text & Me.TXT_MB_TEL.Text)
                Me.Bind_MEMBER(MB_MEMBERList.getCurrentDataSet.Tables(0))
            ElseIf Me.rbt_Name.Checked Then
                '姓名
                Dim MB_MEMBERList As New MB_MEMBERList(Me.m_DBManager)
                MB_MEMBERList.Load_BKSEQ_MB_NAME(Me.TXT_MB_NAME.Text)
                Me.Bind_MEMBER(MB_MEMBERList.getCurrentDataSet.Tables(0))
            End If
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

    Sub Bind_MEMBER(ByVal DT_MB_MEMBER As DataTable)
        If DT_MB_MEMBER.Rows.Count > 0 Then
            Me.RP_BKSEQ.DataSource = DT_MB_MEMBER
            Me.RP_BKSEQ.DataBind()
            Me.PLH_DATA.Visible = True
            Me.DIV_NODATA.Visible = False
        Else
            Me.RP_BKSEQ.DataSource = Nothing
            Me.RP_BKSEQ.DataBind()
            Me.PLH_DATA.Visible = False
            Me.DIV_NODATA.Visible = True
        End If
        Me.DIV_Paras.Visible = False
        Me.DIV_BACK.Visible = True
    End Sub

    Private Sub RP_BKSEQ_ItemDataBound(sender As Object, e As RepeaterItemEventArgs) Handles RP_BKSEQ.ItemDataBound
        If e.Item.ItemType = ListItemType.Item OrElse e.Item.ItemType = ListItemType.AlternatingItem Then
            Dim DRV As DataRowView = CType(e.Item.DataItem, DataRowView)
            '通訊地址
            Dim MB_ADDR As Literal = e.Item.FindControl("MB_ADDR")
            Dim sMB_ADDR As String = DRV("MB_CITY").ToString & DRV("MB_VLG").ToString & DRV("MB_ADDR").ToString
            If com.Azion.EloanUtility.StrUtility.getNVarcharLength(sMB_ADDR) > 14 Then
                sMB_ADDR = com.Azion.EloanUtility.StrUtility.TrimNVarcharToCHARLength(sMB_ADDR, 14) & "XXXX"
            End If
            MB_ADDR.Text = sMB_ADDR

            '手機/電話
            Dim LTL_MB_MOBIL As Literal = e.Item.FindControl("LTL_MB_MOBIL")
            If IsNumeric(DRV("MB_MOBIL")) Then
                LTL_MB_MOBIL.Text = "XXX" & Right(DRV("MB_MOBIL").ToString, 3)
            ElseIf IsNumeric(DRV("MB_TEL")) Then
                LTL_MB_MOBIL.Text = "XXX" & Right(DRV("MB_TEL").ToString, 3)
            End If

            '是否為會員
            Dim LTL_MB_ACCT As Literal = e.Item.FindControl("LTL_MB_ACCT")
            If Utility.isValidateData(DRV("MB_EMAIL")) Then
                Dim MB_ACCT As New MB_ACCT(Me.m_DBManager)
                If MB_ACCT.LoadByPK(DRV("MB_EMAIL").ToString) Then
                    LTL_MB_ACCT.Text = "是"
                Else
                    LTL_MB_ACCT.Text = "否"
                End If
            Else
                LTL_MB_ACCT.Text = "否"
            End If
        End If
    End Sub

    Private Sub btnBack_Click(sender As Object, e As EventArgs) Handles btnBack.Click
        Me.RP_BKSEQ.DataSource = Nothing
        Me.RP_BKSEQ.DataBind()
        Me.PLH_DATA.Visible = False
        Me.DIV_NODATA.Visible = False

        Me.DIV_Paras.Visible = True
        Me.DIV_BACK.Visible = False
    End Sub

    Private Sub btnSingin_Click(sender As Object, e As EventArgs) Handles btnSingin.Click
        Try
            Dim sURL As String = String.Empty

            sURL = com.Azion.EloanUtility.UIUtility.getRootPath & "/MNT/MBSignIn_01_v01.aspx"

            Response.Redirect(sURL)
        Catch ex As System.Threading.ThreadAbortException
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub
End Class