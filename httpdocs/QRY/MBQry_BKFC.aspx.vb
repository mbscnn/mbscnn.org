Imports MBSC.MB_OP
Imports com.Azion.NET.VB
Imports MBSC.UICtl
Imports com.Azion.EloanUtility

Public Class MBQry_BKFC
    Inherits System.Web.UI.Page

    Dim m_sUpcode76 As String = String.Empty

    Dim m_sUSERID As String = String.Empty

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            '維護人員
            Me.m_sUpcode76 = com.Azion.EloanUtility.CodeList.getAppSettings("UpCode76")

            Me.m_sUSERID = Session("USERID")

            If Not Page.IsPostBack Then
                If Utility.isValidateData(Me.m_sUSERID) Then
                    If Me.is76User(Me.m_sUSERID) Then
                        Me.PLH_MEMBER.Visible = False
                        Me.PLH_Paras.Visible = True
                    Else
                        Me.MB_BKSEQ.Attributes("Readonly")=true
                        Me.MB_BKSEQ.Style.Item("border") = "none"
                        Me.MB_BKSEQ.Style.Item("background") = "transparent"
                        Using dbManager As DatabaseManager = UIShareFun.getDataBaseManager
                            Dim MB_MEMBERList As New MB_MEMBERList(dbManager)
                            MB_MEMBERList.Load_MB_EMAIL_BKSEQ(Me.m_sUSERID)
                            If MB_MEMBERList.getCurrentDataSet.Tables(0).Rows.Count > 1 Then
                                Me.RP_MEMBER.DataSource = MB_MEMBERList.getCurrentDataSet.Tables(0)
                                Me.RP_MEMBER.DataBind()
                                Me.PLH_MEMBER.Visible = True
                                Me.PLH_Paras.Visible = False
                            ElseIf MB_MEMBERList.getCurrentDataSet.Tables(0).Rows.Count = 1 Then
                                Me.PLH_MEMBER.Visible = False
                                Me.PLH_Paras.Visible = True
                                Me.MB_BKSEQ.Text = MB_MEMBERList.getCurrentDataSet.Tables(0).Rows(0)("MB_BKSEQ").ToString
                            End If
                        End Using
                    End If
                End If
            End If
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

    Private Sub btn_Qry_Click(sender As Object, e As EventArgs) Handles btn_Qry.Click
        Try
            If Not Utility.isValidateData(Me.MB_BKSEQ.Text) Then
                com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "請輸入護法會員編號")
                com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請輸入護法會員編號")
                Return
            ElseIf Not IsNumeric(Me.MB_BKSEQ.text) Then
                com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "護法會員編號請輸入數字")
                com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "護法會員編號請輸入數字")
                Return
            End If

            If Not Utility.isValidateData(Me.RBL_QRYTYP.SelectedValue) Then
                com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "請選擇查詢類別")
                com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請選擇查詢類別")
                Return
            End If

            If Me.RBL_QRYTYP.SelectedIndex = 0 Then
                Using dbManager As DatabaseManager = UIShareFun.getDataBaseManager
                    Dim MB_MEMREVList As New MB_MEMREVList(dbManager)
                    MB_MEMREVList.setSQLCondition(" ORDER BY B.MB_TX_DATE ")
                    MB_MEMREVList.Load_MB_BKSEQ(Me.MB_BKSEQ.Text)
                    If MB_MEMREVList.getCurrentDataSet.Tables(0).Rows.Count > 0 Then
                        Me.RP_MB_MEMREV.DataSource = MB_MEMREVList.getCurrentDataSet.Tables(0)
                        Me.RP_MB_MEMREV.DataBind()
                        Me.PLH_MB_MEMREV.Visible = True
                        Me.PLH_Paras.Visible = False
                    Else
                        com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "查無捐贈資料")
                        com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "查無捐贈資料")
                    End If
                End Using
            ElseIf Me.RBL_QRYTYP.SelectedIndex = 1 Then
                Using dbManager As DatabaseManager = UIShareFun.getDataBaseManager
                    Dim MB_MEMCLASSList As New MB_MEMCLASSList(dbManager)
                    MB_MEMCLASSList.setSQLCondition(" ORDER BY B.MB_SEQ ")
                    MB_MEMCLASSList.Load_MB_BKSEQ(Me.MB_BKSEQ.Text)
                    If MB_MEMCLASSList.getCurrentDataSet.Tables(0).Rows.Count > 0 Then
                        Me.RP_MB_MEMCLASS.DataSource = MB_MEMCLASSList.getCurrentDataSet.Tables(0)
                        Me.RP_MB_MEMCLASS.DataBind()
                        Me.PLH_MB_MEMCLASS.Visible = True
                        Me.PLH_Paras.Visible = False
                    Else
                        com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "查無參加課程資料")
                        com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "查無參加課程資料")
                    End If
                End Using
            End If
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

    Private Sub btn_Sign_Qry_Click(sender As Object, e As EventArgs) Handles btn_Sign_Qry.Click
        Try
            Dim sMB_BKSEQ As String = String.Empty
            For Each ITEM As RepeaterItem In Me.RP_MEMBER.Items
                Dim RB_CHOOSE As RadioButton = ITEM.FindControl("RB_CHOOSE")
                If RB_CHOOSE.Checked Then
                    Dim LTL_MB_BKSEQ As Literal = ITEM.FindControl("LTL_MB_BKSEQ")
                    sMB_BKSEQ = LTL_MB_BKSEQ.Text
                End If
            Next

            If Utility.isValidateData(sMB_BKSEQ) Then
                Me.MB_BKSEQ.Text = sMB_BKSEQ
                Me.PLH_MEMBER.Visible = False
                Me.PLH_Paras.Visible = True
            Else
                com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "請選取會員")
                com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請選取會員")
                Return
            End If
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

    Private Sub btnReQry_1_Click(sender As Object, e As EventArgs) Handles btnReQry_1.Click, btnReQry_2.Click
        Me.PLH_MB_MEMREV.Visible = False
        Me.PLH_MB_MEMCLASS.Visible = False
        Me.RP_MB_MEMCLASS.DataSource = Nothing
        Me.RP_MB_MEMCLASS.DataBind()
        Me.RP_MB_MEMREV.DataSource = Nothing
        Me.RP_MB_MEMREV.DataBind()

        If Me.RP_MEMBER.Items.Count > 1 Then
            Me.PLH_MEMBER.Visible = True
            Me.PLH_Paras.Visible = False
        Else
            Me.PLH_Paras.Visible = True
            Me.PLH_MEMBER.Visible = False
        End If
    End Sub

    Private Sub RP_MB_MEMCLASS_ItemDataBound(sender As Object, e As RepeaterItemEventArgs) Handles RP_MB_MEMCLASS.ItemDataBound
        If e.Item.ItemType = ListItemType.Item OrElse e.Item.ItemType = ListItemType.AlternatingItem Then
            Dim DRV As DataRowView = CType(e.Item.DataItem, DataRowView)

            Dim LTL_MB_BATCH As Literal = e.Item.FindControl("LTL_MB_BATCH")
            Using dbManager As DatabaseManager = UIShareFun.getDataBaseManager
                Dim MB_MEMBATCHList As New MB_MEMBATCHList(dbManager)
                MB_MEMBATCHList.LoadBySEQ(DRV("MB_MEMSEQ"), DRV("MB_SEQ"))
                Dim ROW_MB_ELECT() As DataRow = MB_MEMBATCHList.getCurrentDataSet.Tables(0).Select("ISNULL(MB_CHKFLAG,' ')='1'", "MB_ELECT")
                If Not IsNothing(ROW_MB_ELECT) AndAlso ROW_MB_ELECT.Length > 0 Then
                    For Each ROW As DataRow In ROW_MB_ELECT
                        If ROW("MB_BATCH") = 0 Then
                            LTL_MB_BATCH.Text &= "本課程無梯次，"
                        Else
                            LTL_MB_BATCH.Text &= ROW("MB_BATCH") & "，"
                        End If
                    Next
                End If

                Dim MB_CLASSList As New MB_CLASSList(dbManager)
                MB_CLASSList.LoadByMB_SEQ(DRV("MB_SEQ"))
                If MB_CLASSList.getCurrentDataSet.Tables(0).Rows.Count > 0 Then
                    Dim ROW As DataRow = MB_CLASSList.getCurrentDataSet.Tables(0).Rows(0)
                    '地點
                    Dim MB_PLACE As Literal = e.Item.FindControl("MB_PLACE")
                    MB_PLACE.Text = ROW("MB_PLACE").ToString

                    '課程名稱
                    Dim MB_CLASS_NAME As Literal = e.Item.FindControl("MB_CLASS_NAME")
                    MB_CLASS_NAME.Text = ROW("MB_CLASS_NAME").ToString

                    '課程起訖日
                    Dim MB_SDATE As Literal = e.Item.FindControl("MB_SDATE")
                    If IsDate(ROW("MB_SDATE").ToString) Then
                        MB_SDATE.Text = CDate(ROW("MB_SDATE").ToString).Year & "/" & CDate(ROW("MB_SDATE").ToString).Month & "/" & CDate(ROW("MB_SDATE").ToString).Day
                    End If

                    If IsDate(ROW("MB_EDATE").ToString) Then
                        MB_SDATE.Text &= "~" & CDate(ROW("MB_EDATE").ToString).Year & "/" & CDate(ROW("MB_EDATE").ToString).Month & "/" & CDate(ROW("MB_EDATE").ToString).Day
                    End If

                    '指導老師
                    Dim MB_TEACHER As Literal = e.Item.FindControl("MB_TEACHER")
                    MB_TEACHER.Text = ROW("MB_TEACHER").ToString
                End If
            End Using
            If Utility.isValidateData(LTL_MB_BATCH.Text) Then
                LTL_MB_BATCH.Text = Left(LTL_MB_BATCH.Text, LTL_MB_BATCH.Text.Length - 1)
            End If
        End If
    End Sub

    Private Sub RP_MB_MEMREV_ItemDataBound(sender As Object, e As RepeaterItemEventArgs) Handles RP_MB_MEMREV.ItemDataBound
        If e.Item.ItemType = ListItemType.Item OrElse e.Item.ItemType = ListItemType.AlternatingItem Then
            Dim DRV As DataRowView = CType(e.Item.DataItem, DataRowView)

            '繳款日期
            If Utility.isValidateData(DRV("MB_TX_DATE").ToString) Then
                Dim MB_TX_DATE As Literal = e.Item.FindControl("MB_TX_DATE")
                MB_TX_DATE.Text = Left(DRV("MB_TX_DATE").ToString, 4) & "/" & Left(Right(DRV("MB_TX_DATE").ToString, 4), 2) & "/" & Right(DRV("MB_TX_DATE").ToString, 2)
            End If

            '繳款金額
            If IsNumeric(DRV("MB_TOTFEE")) Then
                Dim MB_TOTFEE As Literal = e.Item.FindControl("MB_TOTFEE")
                MB_TOTFEE.Text = Utility.FormatDec(DRV("MB_TOTFEE"), "#,##0")
            End If

            Using dbManager As DatabaseManager = UIShareFun.getDataBaseManager
                '功德項目
                'Dim AP_CODEList As New AP_CODEList(dbManager)
                'AP_CODEList.loadByUpCode(15)
                'Dim ROW_MB_ITEMID() As DataRow = Nothing
                'ROW_MB_ITEMID = AP_CODEList.getCurrentDataSet.Tables(0).Select("VALUE='" & DRV("MB_ITEMID").ToString & "'")
                'If Not IsNothing(ROW_MB_ITEMID) AndAlso ROW_MB_ITEMID.Length > 0 Then
                '    Dim MB_ITEMID As Literal = e.Item.FindControl("MB_ITEMID")
                '    MB_ITEMID.Text = ROW_MB_ITEMID(0)("TEXT").ToString
                'End If

                Dim AP_CODE As New AP_CODE(dbManager)
                If AP_CODE.loadByValue(15, DRV("MB_ITEMID").ToString) Then
                    Dim MB_ITEMID As Literal = e.Item.FindControl("MB_ITEMID")
                    MB_ITEMID.Text = AP_CODE.getString("TEXT")
                End If

                '會員類別
                If DRV("MB_ITEMID").ToString = "A" Then
                    'AP_CODEList.clear()
                    'AP_CODEList.loadByUpCode(28)
                    'Dim ROW_MB_MEMTYP() As DataRow = Nothing
                    'ROW_MB_MEMTYP = AP_CODEList.getCurrentDataSet.Tables(0).Select("VALUE='" & DRV("MB_MEMTYP").ToString & "'")
                    'If Not IsNothing(ROW_MB_MEMTYP) AndAlso ROW_MB_MEMTYP.Length > 0 Then
                    '    Dim MB_MEMTYP As Literal = e.Item.FindControl("MB_MEMTYP")
                    '    MB_MEMTYP.Text = ROW_MB_MEMTYP(0)("TEXT").ToString
                    'End If

                    AP_CODE.clear()
                    If AP_CODE.loadByValue(28, DRV("MB_MEMTYP").ToString) Then
                        Dim MB_MEMTYP As Literal = e.Item.FindControl("MB_MEMTYP")
                        MB_MEMTYP.Text = AP_CODE.getString("TEXT")
                    End If
                End If

                '捐贈管道
                '還未開欄位
            End Using
        End If
    End Sub

    Private Sub RP_MEMBER_ItemDataBound(sender As Object, e As RepeaterItemEventArgs) Handles RP_MEMBER.ItemDataBound
        If e.Item.ItemType = ListItemType.Item OrElse e.Item.ItemType = ListItemType.AlternatingItem Then
            Dim DRV As DataRowView = CType(e.Item.DataItem, DataRowView)
            '通訊地址
            Dim LTL_MB_ADDR As Literal = e.Item.FindControl("LTL_MB_ADDR")
            LTL_MB_ADDR.Text = DRV("MB_CITY").ToString & DRV("MB_VLG").ToString & DRV("MB_ADDR").ToString
            '手機/電話
            Dim LTL_MB_TEL As Literal = e.Item.FindControl("LTL_MB_TEL")
            If Utility.isValidateData(DRV("MB_MOBIL")) Then
                LTL_MB_TEL.Text = DRV("MB_MOBIL").ToString
            ElseIf Utility.isValidateData(DRV("MB_TEL")) Then
                LTL_MB_TEL.Text = DRV("MB_TEL").ToString
            End If
        End If
    End Sub

    Function is76User(ByVal sUserId As String) As Boolean
        Using dbManager As DatabaseManager = UIShareFun.getDataBaseManager
            Dim apCODEList As New AP_CODEList(dbManager)
            apCODEList.loadByUpCode(Me.m_sUpcode76)
            Dim ROW_USER() As DataRow = Nothing
            ROW_USER = apCODEList.getCurrentDataSet.Tables(0).Select("TEXT='" & sUserId & "'")
            If Not IsNothing(ROW_USER) AndAlso ROW_USER.Length > 0 Then
                Return True
            Else
                Return False
            End If
        End Using
    End Function

End Class