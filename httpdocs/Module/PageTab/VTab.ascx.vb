Imports MBSC.MB_OP

Public Class VTab
    Inherits System.Web.UI.UserControl

    Dim m_DBManager As com.Azion.NET.VB.DatabaseManager

    Dim m_iUPCODE_LV As Integer = 37
    Dim m_iUPCODE_78 As Integer = 78

    Dim m_iLEVEL As Integer = 0

    'Dim m_sCODEID As String = String.Empty

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.m_DBManager = UICtl.UIShareFun.getDataBaseManager

        'Me.m_sCODEID = "" & Request.QueryString("CLASS")

        '會員系統
        Me.m_iUPCODE_78 = com.Azion.EloanUtility.CodeList.getAppSettings("UPCODE78")

        If Not Page.IsPostBack Then
            Me.Bind_RP_Tab()
        End If
    End Sub

    Private Sub Page_Unload(sender As Object, e As EventArgs) Handles Me.Unload
        UICtl.UIShareFun.releaseConnection(Me.m_DBManager)
    End Sub

    Public Sub Bind_RP_Tab()
        Try
            If IsNumeric(Session("admin")) Then
                Me.m_iLEVEL = CInt(Session("admin"))
            End If

            Dim apCODEList As New AP_CODEList(Me.m_DBManager)
            apCODEList.setSQLCondition(" ORDER BY SORTNO ")
            apCODEList.loadByUpCode(Me.m_iUPCODE_LV)

            Dim ROW_LV1() As DataRow = apCODEList.getCurrentDataSet.Tables(0).Select("DISABLED='1' AND LEVEL<=" & Me.m_iLEVEL)
            Me.RP_Tab.DataSource = ROW_LV1
            Me.RP_Tab.DataBind()
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Private Sub RP_Tab_ItemDataBound(sender As Object, e As RepeaterItemEventArgs) Handles RP_Tab.ItemDataBound
        Try
            If e.Item.ItemType = ListItemType.Item OrElse e.Item.ItemType = ListItemType.AlternatingItem Then
                Dim DRV As DataRow = CType(e.Item.DataItem, DataRow)

                Dim apCODEList As New AP_CODEList(Me.m_DBManager)
                apCODEList.setSQLCondition(" ORDER BY SORTNO ")
                Dim ROW_LV2() As DataRow = Nothing
                If DRV("CODEID") = Me.m_iUPCODE_78 AndAlso Utility.isValidateData(Session("UserId")) Then
                    apCODEList.Load_TAB_AUTH(DRV("CODEID"), Session("UserId"))
                    ROW_LV2 = apCODEList.getCurrentDataSet.Tables(0).Select("DISABLED='1'")
                Else
                    apCODEList.loadByUpCode(DRV("CODEID"))
                    ROW_LV2 = apCODEList.getCurrentDataSet.Tables(0).Select("DISABLED='1' AND LEVEL<=" & Me.m_iLEVEL)
                End If

                Dim RP_Tab_LV2 As Repeater = e.Item.FindControl("RP_Tab_LV2")
                If Not IsNothing(ROW_LV2) AndAlso ROW_LV2.Length > 0 Then
                    Dim PLH_LV2 As PlaceHolder = e.Item.FindControl("PLH_LV2")
                    PLH_LV2.Visible = True
                    RP_Tab_LV2.DataSource = ROW_LV2
                    RP_Tab_LV2.DataBind()
                End If

                Dim HPL_LV1 As HyperLink = e.Item.FindControl("HPL_LV1")
                If Utility.isValidateData(DRV("NOTE")) Then
                    If DRV("CODEID").ToString = "38" Then
                        HPL_LV1.NavigateUrl = com.Azion.EloanUtility.UIUtility.getRootPath & "/" & DRV("NOTE")
                    Else
                        If InStr(DRV("NOTE"), "?") > 0 Then
                            HPL_LV1.NavigateUrl = com.Azion.EloanUtility.UIUtility.getRootPath & "/" & DRV("NOTE") & "&CLASS=" & DRV("CODEID")
                        Else
                            HPL_LV1.NavigateUrl = com.Azion.EloanUtility.UIUtility.getRootPath & "/" & DRV("NOTE") & "?CLASS=" & DRV("CODEID")
                        End If
                    End If
                Else
                    HPL_LV1.Attributes("href") = "#"
                End If
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Public Sub RP_Tab_LV2_ItemDataBound(sender As Object, e As RepeaterItemEventArgs)
        Try
            If e.Item.ItemType = ListItemType.Item OrElse e.Item.ItemType = ListItemType.AlternatingItem Then
                Dim DRV As DataRow = CType(e.Item.DataItem, DataRow)

                Dim sCLASS As String = String.Empty
                sCLASS = "" & Request.QueryString("CLASS")
                Dim HID_CODEID As HtmlGenericControl = e.Item.FindControl("HID_CODEID")
                If sCLASS = HID_CODEID.InnerText Then
                    Dim objItem As RepeaterItem = sender.namingcontainer
                    Dim Tab_Li As HtmlGenericControl = objItem.FindControl("Tab_Li")
                    Tab_Li.Attributes("class") = "active"

                    Dim Tab_Li_C As HtmlGenericControl = e.Item.FindControl("Tab_Li")
                    Tab_Li_C.Attributes("style") = "background: #003545;border-left: 5px solid #970000;"
                Else
                    If Utility.isValidateData(sCLASS) Then
                        Dim ap_C_CODEList As New AP_CODEList(Me.m_DBManager)
                        ap_C_CODEList.loadByUpCode(HID_CODEID.InnerText)
                        Dim ROW_CHK() As DataRow = Nothing
                        ROW_CHK = ap_C_CODEList.getCurrentDataSet.Tables(0).Select("CODEID=" & sCLASS)
                        If Not IsNothing(ROW_CHK) AndAlso ROW_CHK.Length > 0 Then
                            Dim objItem As RepeaterItem = sender.namingcontainer
                            Dim Tab_Li As HtmlGenericControl = objItem.FindControl("Tab_Li")
                            Tab_Li.Attributes("class") = "active"

                            Dim Tab_Li_C As HtmlGenericControl = e.Item.FindControl("Tab_Li")
                            Tab_Li_C.Attributes("style") = "background: #003545;border-left: 5px solid #970000;"
                        End If
                    End If
                End If

                Dim apCODEList As New AP_CODEList(Me.m_DBManager)
                apCODEList.setSQLCondition(" ORDER BY SORTNO ")
                Dim ROW_LV3() As DataRow = Nothing
                If DRV("CODEID") = Me.m_iUPCODE_78 AndAlso Utility.isValidateData(Session("UserId")) Then
                    apCODEList.Load_TAB_AUTH(DRV("CODEID"), Session("UserId"))
                    ROW_LV3 = apCODEList.getCurrentDataSet.Tables(0).Select("DISABLED='1'")
                Else
                    apCODEList.loadByUpCode(DRV("CODEID"))
                    ROW_LV3 = apCODEList.getCurrentDataSet.Tables(0).Select("DISABLED='1' AND LEVEL<=" & Me.m_iLEVEL)
                End If

                Dim RP_Tab_LV3 As Repeater = e.Item.FindControl("RP_Tab_LV3")
                If Not IsNothing(ROW_LV3) AndAlso ROW_LV3.Length > 0 Then
                    Dim PLH_LV3 As PlaceHolder = e.Item.FindControl("PLH_LV3")
                    PLH_LV3.Visible = True
                    RP_Tab_LV3.DataSource = ROW_LV3
                    RP_Tab_LV3.DataBind()
                End If

                Dim HPL_LV2 As HyperLink = e.Item.FindControl("HPL_LV2")
                If Utility.isValidateData(DRV("NOTE")) Then
                    If DRV("CODEID").ToString = "38" Then
                        HPL_LV2.NavigateUrl = com.Azion.EloanUtility.UIUtility.getRootPath & "/" & DRV("NOTE")
                    Else
                        If InStr(DRV("NOTE"), "?") > 0 Then
                            HPL_LV2.NavigateUrl = com.Azion.EloanUtility.UIUtility.getRootPath & "/" & DRV("NOTE") & "&CLASS=" & DRV("CODEID")
                        Else
                            HPL_LV2.NavigateUrl = com.Azion.EloanUtility.UIUtility.getRootPath & "/" & DRV("NOTE") & "?CLASS=" & DRV("CODEID")
                        End If
                    End If
                Else
                    HPL_LV2.Attributes("href") = "#"
                End If
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Public Sub RP_Tab_LV3_ItemDataBound(sender As Object, e As RepeaterItemEventArgs)
        Try
            If e.Item.ItemType = ListItemType.Item OrElse e.Item.ItemType = ListItemType.AlternatingItem Then
                Dim DRV As DataRow = CType(e.Item.DataItem, DataRow)

                Dim sCLASS As String = String.Empty
                sCLASS = "" & Request.QueryString("CLASS")
                Dim HID_CODEID As HtmlGenericControl = e.Item.FindControl("HID_CODEID")
                If sCLASS = HID_CODEID.InnerText Then
                    Dim objItem As RepeaterItem = sender.namingcontainer
                    Dim Tab_Li As HtmlGenericControl = objItem.FindControl("Tab_Li")
                    Tab_Li.Attributes("class") = "active"

                    Dim Tab_Li_C As HtmlGenericControl = e.Item.FindControl("Tab_Li")
                    Tab_Li_C.Attributes("style") = "background: #003545;border-left: 5px solid #970000;"

                    'Dim apCODE As New AP_CODE(Me.m_DBManager)
                    'apCODE.loadByPK(sCLASS)
                    'Dim objPItem As RepeaterItem = sender.namingcontainer
                    'Dim HID_P_CODEID As HtmlGenericControl = objPItem.FindControl("HID_CODEID")
                    'If HID_P_CODEID.InnerText = apCODE.getString("UPCODE") Then
                    '    Dim Tab_P_Li As HtmlGenericControl = objPItem.FindControl("Tab_Li")
                    '    Tab_P_Li.Attributes("class") = "active"

                    '    Dim objPPItem As RepeaterItem = Tab_P_Li.NamingContainer
                    '    Dim Tab_P_P_Li As HtmlGenericControl = objPItem.FindControl("Tab_Li")
                    '    Tab_P_P_Li.Attributes("class") = "active"
                    'End If
                End If

                Dim apCODEList As New AP_CODEList(Me.m_DBManager)
                apCODEList.setSQLCondition(" ORDER BY SORTNO ")
                Dim ROW_LV4() As DataRow = Nothing
                If DRV("CODEID") = Me.m_iUPCODE_78 AndAlso Utility.isValidateData(Session("UserId")) Then
                    apCODEList.Load_TAB_AUTH(DRV("CODEID"), Session("UserId"))
                    ROW_LV4 = apCODEList.getCurrentDataSet.Tables(0).Select("DISABLED='1'")
                Else
                    apCODEList.loadByUpCode(DRV("CODEID"))
                    ROW_LV4 = apCODEList.getCurrentDataSet.Tables(0).Select("DISABLED='1' AND LEVEL<=" & Me.m_iLEVEL)
                End If

                Dim RP_Tab_LV4 As Repeater = e.Item.FindControl("RP_Tab_LV4")
                If Not IsNothing(ROW_LV4) AndAlso ROW_LV4.Length > 0 AndAlso Not IsNothing(RP_Tab_LV4) Then
                    Dim PLH_LV4 As PlaceHolder = e.Item.FindControl("PLH_LV4")
                    PLH_LV4.Visible = True
                    RP_Tab_LV4.DataSource = ROW_LV4
                    RP_Tab_LV4.DataBind()
                End If

                Dim HPL_LV3 As HyperLink = e.Item.FindControl("HPL_LV3")
                If Utility.isValidateData(DRV("NOTE")) Then
                    If DRV("CODEID").ToString = "38" Then
                        HPL_LV3.NavigateUrl = com.Azion.EloanUtility.UIUtility.getRootPath & "/" & DRV("NOTE")
                    Else
                        If InStr(DRV("NOTE"), "?") > 0 Then
                            HPL_LV3.NavigateUrl = com.Azion.EloanUtility.UIUtility.getRootPath & "/" & DRV("NOTE") & "&CLASS=" & DRV("CODEID")
                        Else
                            HPL_LV3.NavigateUrl = com.Azion.EloanUtility.UIUtility.getRootPath & "/" & DRV("NOTE") & "?CLASS=" & DRV("CODEID")
                        End If
                    End If
                Else
                    HPL_LV3.Attributes("href") = "#"
                End If
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub
End Class