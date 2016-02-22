Imports MBSC.MB_OP
Imports com.Azion.NET.VB
Imports MBSC.UICtl

Public Class MBMnt_APV_01_v01
    Inherits System.Web.UI.Page

    Dim m_DBManager As com.Azion.NET.VB.DatabaseManager

    Dim m_sUpCode1 As String = String.Empty
    Dim m_sUpCode7 As String = String.Empty
    Dim m_sUpCode28 As String = String.Empty

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.m_DBManager = UIShareFun.getDataBaseManager

            '學歷
            Me.m_sUpCode1 = com.Azion.EloanUtility.CodeList.getAppSettings("UpCode1")

            '所屬區代碼
            Me.m_sUpCode7 = com.Azion.EloanUtility.CodeList.getAppSettings("UpCode7")

            '會員類別
            Me.m_sUpCode28 = com.Azion.EloanUtility.CodeList.getAppSettings("UpCode28")
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

    Sub RB_CHOOSE_OnCheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            For Each objItem As RepeaterItem In Me.RP_APV.Items
                Dim RB_CHOOSE As RadioButton = objItem.FindControl("RB_CHOOSE")
                RB_CHOOSE.Checked = False
            Next

            Me.CLEAR_MB_MEMBER_TEMP()
            Dim ITEM As RepeaterItem = sender.namingcontainer
            Dim RB_CHOOSE_CH As RadioButton = ITEM.FindControl("RB_CHOOSE")
            RB_CHOOSE_CH.Checked = True
            Dim MB_DATE As HtmlInputHidden = ITEM.FindControl("MB_DATE")
            Me.Bind_MB_MEMBER_TEMP(CType(MB_DATE.Value, Long))
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

    Sub CLEAR_MB_MEMBER_TEMP()
        Try
            Me.PLH_MB_MEMBER.Visible = False

            Me.MB_NAME.Text = String.Empty

            Me.MB_DATE.Value = String.Empty

            Me.MB_SEX.Text = String.Empty

            Me.MB_BIRTH.Text = String.Empty

            Me.MB_IDENTIFY.SelectedIndex = -1

            Me.MB_MOBIL.Text = String.Empty

            Me.MB_TEL_Pre.Text = String.Empty

            Me.MB_TEL.Text = String.Empty

            Me.MB_EMAIL.Text = String.Empty

            Me.MB_ID.Text = String.Empty

            Me.MB_EDU.Text = String.Empty

            Me.MB_AREA.Text = String.Empty

            Me.MB_LEADER.Text = String.Empty

            Me.MB_CITY.Text = String.Empty

            Me.MB_VLG.Text = String.Empty

            Me.MB_ADDR.Text = String.Empty

            Me.MB_ZIP.Text = String.Empty

            Me.MB_CITY1.Text = String.Empty

            Me.MB_VLG1.Text = String.Empty

            Me.MB_ADDR2.Text = String.Empty

            Me.MB_ZIP1.Text = String.Empty

            Me.MB_FAMILY.Text = String.Empty

            Me.VIP.SelectedIndex = -1
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Sub Bind_MB_MEMBER_TEMP(ByVal iMB_DATE As Long)
        Try
            Me.PLH_MB_MEMBER.Visible = True
            Me.MB_DATE.Value = iMB_DATE

            Dim MB_MEMBER_TEMP As New MB_MEMBER_TEMP(Me.m_DBManager)
            If MB_MEMBER_TEMP.LoadByPK(iMB_DATE) Then
                '法名/ 姓名
                Me.MB_NAME.Text = Trim(MB_MEMBER_TEMP.getString("MB_NAME"))

                '性別
                If MB_MEMBER_TEMP.getString("MB_SEX") = "1" Then
                    Me.MB_SEX.Text = "男"
                ElseIf MB_MEMBER_TEMP.getString("MB_SEX") = "2" Then
                    Me.MB_SEX.Text = "女"
                End If

                '出生年月日
                If Utility.isValidateData(MB_MEMBER_TEMP.getAttribute("MB_BIRTH")) AndAlso IsDate(MB_MEMBER_TEMP.getAttribute("MB_BIRTH").ToString()) AndAlso CDate(MB_MEMBER_TEMP.getAttribute("MB_BIRTH").ToString()).Year > 1911 Then
                    Me.MB_BIRTH.Text = CDate(MB_MEMBER_TEMP.getAttribute("MB_BIRTH").ToString()).Year & "年" & _
                                       CDate(MB_MEMBER_TEMP.getAttribute("MB_BIRTH").ToString()).Month & "月" & _
                                       CDate(MB_MEMBER_TEMP.getAttribute("MB_BIRTH").ToString()).Day & "日"
                Else
                    Me.MB_BIRTH.Text = String.Empty
                End If

                '身分別
                Me.MB_IDENTIFY.SelectedIndex = -1
                If Utility.isValidateData(MB_MEMBER_TEMP.getString("MB_IDENTIFY")) Then
                    Dim sMB_IDENTIFY() As String = Split(MB_MEMBER_TEMP.getString("MB_IDENTIFY"), ",")
                    If Not IsNothing(sMB_IDENTIFY) AndAlso sMB_IDENTIFY.Length > 0 Then
                        For i As Integer = 0 To UBound(sMB_IDENTIFY)
                            If Not IsNothing(Me.MB_IDENTIFY.Items.FindByValue(sMB_IDENTIFY(i))) Then
                                Me.MB_IDENTIFY.Items.FindByValue(sMB_IDENTIFY(i)).Selected = True
                            End If
                        Next
                    End If
                End If
                com.Azion.EloanUtility.UIUtility.setControlRead(Me.MB_IDENTIFY)

                '手機
                Me.MB_MOBIL.Text = Trim(MB_MEMBER_TEMP.getString("MB_MOBIL"))

                '電話
                If Utility.isValidateData(Trim(MB_MEMBER_TEMP.getString("MB_TEL"))) Then
                    If InStr(Trim(MB_MEMBER_TEMP.getString("MB_TEL")), "-") > 0 Then
                        Dim sMB_TEL() As String = Split(Trim(MB_MEMBER_TEMP.getString("MB_TEL")), "-")
                        If Not IsNothing(sMB_TEL) AndAlso sMB_TEL.Length >= 1 Then
                            Me.MB_TEL_Pre.Text = Trim(sMB_TEL(0))
                            Me.MB_TEL.Text = Trim(sMB_TEL(1))
                        End If
                    Else
                        Me.MB_TEL_Pre.Text = String.Empty
                        Me.MB_TEL.Text = Trim(MB_MEMBER_TEMP.getString("MB_TEL"))
                    End If
                Else
                    Me.MB_TEL_Pre.Text = String.Empty
                    Me.MB_TEL.Text = String.Empty
                End If

                'e-mail
                Me.MB_EMAIL.Text = Trim(MB_MEMBER_TEMP.getString("MB_EMAIL"))

                '身分證字號
                Me.MB_ID.Text = Trim(MB_MEMBER_TEMP.getString("MB_ID"))

                '學歷
                Dim apCODEList As New AP_CODEList(Me.m_DBManager)
                apCODEList.loadByUpCode(Me.m_sUpCode1)
                Dim ROW_MB_EDU() As DataRow = Nothing
                ROW_MB_EDU = apCODEList.getCurrentDataSet.Tables(0).Select("VALUE='" & MB_MEMBER_TEMP.getString("MB_EDU") & "'")
                If Not IsNothing(ROW_MB_EDU) AndAlso ROW_MB_EDU.Length > 0 Then
                    Me.MB_EDU.Text = ROW_MB_EDU(0)("TEXT").ToString
                Else
                    Me.MB_EDU.Text = String.Empty
                End If

                '所屬區
                apCODEList.clear()
                apCODEList.loadByUpCode(Me.m_sUpCode7)
                Dim ROW_MB_AREA() As DataRow = Nothing
                ROW_MB_AREA = apCODEList.getCurrentDataSet.Tables(0).Select("VALUE='" & MB_MEMBER_TEMP.getString("MB_AREA") & "'")
                If Not IsNothing(ROW_MB_AREA) AndAlso ROW_MB_AREA.Length > 0 Then
                    Me.MB_AREA.Text = ROW_MB_AREA(0)("TEXT").ToString
                Else
                    Me.MB_AREA.Text = String.Empty
                End If

                '委員
                Me.MB_LEADER.Text = MB_MEMBER_TEMP.getString("MB_LEADER")

                '通訊地址
                Me.MB_CITY.Text = MB_MEMBER_TEMP.getString("MB_CITY")
                Me.MB_VLG.Text = MB_MEMBER_TEMP.getString("MB_VLG")
                Me.MB_ADDR.Text = Trim(MB_MEMBER_TEMP.getString("MB_ADDR"))
                '通訊郵遞區號
                Me.MB_ZIP.Text = Trim(MB_MEMBER_TEMP.getString("MB_ZIP"))

                '戶籍地址
                Me.MB_CITY1.Text = MB_MEMBER_TEMP.getString("MB_CITY1")
                Me.MB_VLG1.Text = MB_MEMBER_TEMP.getString("MB_VLG1")
                Me.MB_ADDR2.Text = Trim(MB_MEMBER_TEMP.getString("MB_ADDR1"))
                '戶籍郵遞區號
                Me.MB_ZIP1.Text = Trim(MB_MEMBER_TEMP.getString("MB_ZIP1"))

                '會員類別
                apCODEList.clear()
                apCODEList.loadByUpCode(m_sUpCode28)
                Dim ROW_MB_FAMILY() As DataRow = Nothing
                ROW_MB_FAMILY = apCODEList.getCurrentDataSet.Tables(0).Select("VALUE='" & MB_MEMBER_TEMP.getString("MB_FAMILY") & "'")
                If Not IsNothing(ROW_MB_FAMILY) AndAlso ROW_MB_FAMILY.Length > 0 Then
                    Me.MB_FAMILY.Text = ROW_MB_FAMILY(0)("TEXT").ToString
                Else
                    Me.MB_FAMILY.Text = String.Empty
                End If

                '護持會員
                Dim MB_VIP_TEMPList As New MB_VIP_TEMPList(Me.m_DBManager)
                MB_VIP_TEMPList.LoadByMB_DATE(CType(MB_DATE.Value, Long))
                For i As Integer = 0 To Me.VIP.Items.Count - 1
                    Dim sVIP As String = String.Empty
                    sVIP = Me.VIP.Items(i).Value
                    Dim ROW_VIP() As DataRow = MB_VIP_TEMPList.getCurrentDataSet.Tables(0).Select("VIP='" & sVIP & "'")
                    If Not IsNothing(ROW_VIP) AndAlso ROW_VIP.Length > 0 Then
                        Me.VIP.Items(i).Selected = True
                    End If
                Next
                com.Azion.EloanUtility.UIUtility.setControlRead(Me.VIP)
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Private Sub MBMnt_APV_01_v01_Unload(sender As Object, e As EventArgs) Handles Me.Unload
        UIShareFun.releaseConnection(Me.m_DBManager)
    End Sub

    Private Sub RBL_OPTYPE_SelectedIndexChanged(sender As Object, e As EventArgs) Handles RBL_OPTYPE.SelectedIndexChanged
        Try
            'If Not Utility.isValidateData(Me.RBL_OPTYPE.SelectedValue) Then
            '    com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "請選擇作業模式")
            '    com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請選擇作業模式")
            '    Return
            'End If

            Me.PLH_OPTYPE_1.Visible = False
            Me.CLEAR_MB_MEMBER_TEMP()
            Me.PLH_OPTYPE_2.Visible = False

            If Me.RBL_OPTYPE.SelectedValue = "1" Then
                Dim MB_MEMBER_TEMP As New MB_MEMBER_TEMPList(Me.m_DBManager)
                MB_MEMBER_TEMP.LoadNULL(Session("AREA"))
                If MB_MEMBER_TEMP.getCurrentDataSet.Tables(0).Rows.Count = 0 Then
                    com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "無核准入會資料")
                    com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "無核准入會資料")
                    Return
                End If

                Me.RP_APV.DataSource = MB_MEMBER_TEMP.getCurrentDataSet.Tables(0)
                Me.RP_APV.DataBind()
                Me.PLH_OPTYPE_1.Visible = True
            Else
                Dim MB_MEMBER_TEMP As New MB_MEMBER_TEMPList(Me.m_DBManager)
                MB_MEMBER_TEMP.Load_D(Session("AREA"))
                If MB_MEMBER_TEMP.getCurrentDataSet.Tables(0).Rows.Count = 0 Then
                    com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "無不核准入會資料")
                    com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "無不核准入會資料")
                    Return
                End If

                Me.RP_APV_D.DataSource = MB_MEMBER_TEMP.getCurrentDataSet.Tables(0)
                Me.RP_APV_D.DataBind()
                Me.PLH_OPTYPE_2.Visible = True
            End If
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

    Private Sub btn_D_Click(sender As Object, e As EventArgs) Handles btn_D.Click
        Try
            Dim mbMEMBER_TEMP As New MB_MEMBER_TEMP(Me.m_DBManager)
            If mbMEMBER_TEMP.LoadByPK(CType(Me.MB_DATE.Value, Long)) Then
                mbMEMBER_TEMP.setAttribute("MB_DELFLG", "D")
                mbMEMBER_TEMP.save()
            End If

            Dim MB_MEMBER_TEMP As New MB_MEMBER_TEMPList(Me.m_DBManager)
            MB_MEMBER_TEMP.LoadNULL(Session("AREA"))
            If MB_MEMBER_TEMP.getCurrentDataSet.Tables(0).Rows.Count = 0 Then
                com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "無核准入會資料")
                com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "無核准入會資料")

                Me.PLH_OPTYPE_1.Visible = False
                Me.RP_APV.DataSource = Nothing
                Me.RP_APV.DataBind()
                Me.PLH_OPTYPE_1.Visible = False
                Me.CLEAR_MB_MEMBER_TEMP()
            Else
                Me.RP_APV.DataSource = MB_MEMBER_TEMP.getCurrentDataSet.Tables(0)
                Me.RP_APV.DataBind()
                Me.PLH_OPTYPE_1.Visible = True
                Me.CLEAR_MB_MEMBER_TEMP()
            End If
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

    Private Sub btn_Y_Click(sender As Object, e As EventArgs) Handles btn_Y.Click
        Try
            Dim mbMEMBER_TEMP As New MB_MEMBER_TEMP(Me.m_DBManager)
            If mbMEMBER_TEMP.LoadByPK(CType(Me.MB_DATE.Value, Long)) Then
                Dim MB_MEMBERList As New MB_MEMBERList(m_DBManager)
                If Utility.isValidateData(mbMEMBER_TEMP.getString("MB_NAME")) AndAlso Utility.isValidateData(mbMEMBER_TEMP.getString("MB_MOBIL")) Then
                    MB_MEMBERList.Load_MB_NAME_MB_MOBIL(mbMEMBER_TEMP.getString("MB_NAME"), mbMEMBER_TEMP.getString("MB_MOBIL"))
                    If MB_MEMBERList.getCurrentDataSet.Tables(0).Rows.Count > 0 Then
                        com.Azion.EloanUtility.UIUtility.showJSMsg("【" & mbMEMBER_TEMP.getString("MB_NAME") & "】【" & mbMEMBER_TEMP.getString("MB_MOBIL") & "】已經是會員了，請再確認看看")
                        com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "【" & mbMEMBER_TEMP.getString("MB_NAME") & "】【" & mbMEMBER_TEMP.getString("MB_MOBIL") & "】已經是會員了，請再確認看看")

                        Return
                    End If
                End If

                If Utility.isValidateData(mbMEMBER_TEMP.getString("MB_NAME")) AndAlso Utility.isValidateData(mbMEMBER_TEMP.getString("MB_TEL")) Then
                    MB_MEMBERList.clear()
                    MB_MEMBERList.Load_MB_NAME_MB_TEL(mbMEMBER_TEMP.getString("MB_NAME"), mbMEMBER_TEMP.getString("MB_TEL"))
                    If MB_MEMBERList.getCurrentDataSet.Tables(0).Rows.Count > 0 Then
                        com.Azion.EloanUtility.UIUtility.showJSMsg("【" & mbMEMBER_TEMP.getString("MB_NAME") & "】【" & mbMEMBER_TEMP.getString("MB_TEL") & "】已經是會員了，請再確認看看")
                        com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "【" & mbMEMBER_TEMP.getString("MB_NAME") & "】【" & mbMEMBER_TEMP.getString("MB_TEL") & "】已經是會員了，請再確認看看")

                        Return
                    End If
                End If

                If Utility.isValidateData(mbMEMBER_TEMP.getString("MB_NAME")) AndAlso Utility.isValidateData(mbMEMBER_TEMP.getString("MB_EMAIL")) Then
                    MB_MEMBERList.clear()
                    MB_MEMBERList.Load_MB_NAME_MB_EMAIL(mbMEMBER_TEMP.getString("MB_NAME"), mbMEMBER_TEMP.getString("MB_EMAIL"))
                    If MB_MEMBERList.getCurrentDataSet.Tables(0).Rows.Count > 0 Then
                        com.Azion.EloanUtility.UIUtility.showJSMsg("【" & mbMEMBER_TEMP.getString("MB_NAME") & "】【" & mbMEMBER_TEMP.getString("MB_EMAIL") & "】已經是會員了，請再確認看看")
                        com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "【" & mbMEMBER_TEMP.getString("MB_NAME") & "】【" & mbMEMBER_TEMP.getString("MB_EMAIL") & "】已經是會員了，請再確認看看")

                        Return
                    End If
                End If

                Dim sProcName As String = String.Empty
                sProcName = LCase(com.Azion.NET.VB.Properties.getSchemaName) & ".MB_GET_MAXID"
                Dim inParaAL As New ArrayList
                Dim outParaAL As New ArrayList
                inParaAL.Add("01")
                inParaAL.Add("1")

                outParaAL.Add(7)

                Dim HT_Return As Hashtable = com.Azion.NET.VB.DBObject.getProcedureData(Me.m_DBManager, sProcName, inParaAL, outParaAL)
                Dim iMAXID As Decimal = HT_Return.Item("@IMAXID")

                Try
                    Me.m_DBManager.beginTran()

                    Dim mbMEMBER As New MB_MEMBER(Me.m_DBManager)
                    Dim mbMEMBERList As New MB_MEMBERList(Me.m_DBManager)
                    mbMEMBERList.loadByMB_MEMSEQ(iMAXID)

                    Dim objTEMP As BosAttributeList = mbMEMBER_TEMP.getBosAttributes
                    For i As Integer = 0 To objTEMP.size - 1
                        Dim sCOLNAME As String = String.Empty
                        sCOLNAME = objTEMP.item(i).getColName

                        If mbMEMBERList.getCurrentDataSet.Tables(0).Columns.Contains(sCOLNAME) Then
                            If mbMEMBERList.getCurrentDataSet.Tables(0).Columns(sCOLNAME).DataType.FullName = "MySql.Data.Types.MySqlDateTime" Then
                                If Utility.isValidateData(mbMEMBER_TEMP.getAttribute(sCOLNAME)) AndAlso IsDate(mbMEMBER_TEMP.getAttribute(sCOLNAME).ToString) Then
                                    mbMEMBER.setAttribute(sCOLNAME, CDate(mbMEMBER_TEMP.getAttribute(sCOLNAME).ToString))
                                Else
                                    mbMEMBER.setAttribute(sCOLNAME, Nothing)
                                End If
                            Else
                                mbMEMBER.setAttribute(sCOLNAME, mbMEMBER_TEMP.getAttribute(sCOLNAME))
                            End If
                        End If
                    Next
                    mbMEMBER.setAttribute("MB_MEMSEQ", iMAXID)
                    mbMEMBER.save()

                    mbMEMBER_TEMP.remove()

                    Dim mbVIP_TEMPList As New MB_VIP_TEMPList(Me.m_DBManager)
                    mbVIP_TEMPList.LoadByMB_DATE(CType(Me.MB_DATE.Value, Long))
                    Dim MB_VIP As New MB_VIP(Me.m_DBManager)
                    For i As Integer = 0 To mbVIP_TEMPList.size - 1
                        Dim MB_VIP_TEMP As MB_VIP_TEMP = CType(mbVIP_TEMPList.item(i), MB_VIP_TEMP)
                        MB_VIP.clear()
                        MB_VIP.setAttribute("MB_MEMSEQ", iMAXID)
                        MB_VIP.setAttribute("VIP", MB_VIP_TEMP.getAttribute("VIP"))
                        MB_VIP.setAttribute("ACCAMT", MB_VIP_TEMP.getAttribute("ACCAMT"))
                        MB_VIP.setAttribute("CHGUID", MB_VIP_TEMP.getAttribute("CHGUID"))
                        MB_VIP.setAttribute("CHGDATE", MB_VIP_TEMP.getAttribute("CHGDATE"))
                        MB_VIP.save()
                        MB_VIP_TEMP.remove()
                    Next

                    Me.m_DBManager.commit()
                Catch ex As Exception
                    Me.m_DBManager.Rollback()
                    Throw
                End Try

                Dim MB_MEMBER_TEMP As New MB_MEMBER_TEMPList(Me.m_DBManager)
                MB_MEMBER_TEMP.LoadNULL(Session("AREA"))
                If MB_MEMBER_TEMP.getCurrentDataSet.Tables(0).Rows.Count = 0 Then
                    com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "無核准入會資料")
                    com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "無核准入會資料")

                    Me.PLH_OPTYPE_1.Visible = False
                    Me.RP_APV.DataSource = Nothing
                    Me.RP_APV.DataBind()
                    Me.PLH_OPTYPE_1.Visible = False
                    Me.CLEAR_MB_MEMBER_TEMP()
                Else
                    Me.RP_APV.DataSource = MB_MEMBER_TEMP.getCurrentDataSet.Tables(0)
                    Me.RP_APV.DataBind()
                    Me.PLH_OPTYPE_1.Visible = True
                    Me.CLEAR_MB_MEMBER_TEMP()
                End If
            End If
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

    Private Sub RP_APV_D_ItemCommand(source As Object, e As RepeaterCommandEventArgs) Handles RP_APV_D.ItemCommand
        Try
            If e.CommandName = "CANCEL" Then
                Dim mbMEMBER_TEMP As New MB_MEMBER_TEMP(Me.m_DBManager)
                If mbMEMBER_TEMP.LoadByPK(CType(e.CommandArgument, Long)) Then
                    mbMEMBER_TEMP.setAttribute("MB_DELFLG", Nothing)
                    mbMEMBER_TEMP.save()
                End If
            ElseIf e.CommandName = "DEL" Then
                Try
                    Me.m_DBManager.beginTran()
                    Dim mbMEMBER_TEMP As New MB_MEMBER_TEMP(Me.m_DBManager)
                    If mbMEMBER_TEMP.LoadByPK(CType(e.CommandArgument, Long)) Then
                        mbMEMBER_TEMP.remove()
                    End If

                    Dim mbVIP_TEMPList As New MB_VIP_TEMPList(Me.m_DBManager)
                    mbVIP_TEMPList.LoadByMB_DATE(CType(e.CommandArgument, Long))
                    For i As Integer = 0 To mbVIP_TEMPList.size - 1
                        mbVIP_TEMPList.item(i).remove()
                    Next
                    Me.m_DBManager.commit()
                Catch ex As Exception
                    Me.m_DBManager.Rollback()
                    Throw
                End Try
            End If

            Dim MB_MEMBER_TEMP As New MB_MEMBER_TEMPList(Me.m_DBManager)
            MB_MEMBER_TEMP.Load_D(Session("AREA"))
            If MB_MEMBER_TEMP.getCurrentDataSet.Tables(0).Rows.Count = 0 Then
                com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "無不核准入會資料")
                com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "無不核准入會資料")
                Me.PLH_OPTYPE_2.Visible = False
                Me.RP_APV_D.DataSource = Nothing
                Me.RP_APV_D.DataBind()
            Else
                Me.RP_APV_D.DataSource = MB_MEMBER_TEMP.getCurrentDataSet.Tables(0)
                Me.RP_APV_D.DataBind()
                Me.PLH_OPTYPE_2.Visible = True
            End If
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

    Private Sub RP_APV_D_ItemDataBound(sender As Object, e As RepeaterItemEventArgs) Handles RP_APV_D.ItemDataBound
        Try
            If e.Item.ItemType = ListItemType.Item OrElse e.Item.ItemType = ListItemType.AlternatingItem Then
                Dim IMG_DEL As ImageButton = e.Item.FindControl("IMG_DEL")
                IMG_DEL.ImageUrl = com.Azion.EloanUtility.UIUtility.getRootPath & "/IMG/trashcan.gif"
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub
End Class