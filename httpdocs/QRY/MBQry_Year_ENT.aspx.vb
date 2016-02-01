
Imports MBSC.MB_OP
Imports com.Azion.NET.VB
Imports MBSC.UICtl

Public Class MBQry_Year_ENT
    Inherits System.Web.UI.Page

    Dim m_sYear As String = String.Empty
    Dim m_DBManager As DatabaseManager

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.m_DBManager = UIShareFun.getDataBaseManager
            If Not Page.IsPostBack Then
                Me.m_sYear = "" & Request.QueryString("YEAR")

                Me.LTL_ENT_YEAR.Text = Me.m_sYear

                Dim D_SF_SDATE As Date = New Date(Me.m_sYear, 1, 1)
                Dim D_EF_SDATE As Date = New Date(Me.m_sYear, 6, 30)
                Dim MB_YENTList As New MB_YENTList(Me.m_DBManager)
                MB_YENTList.Load_SDATE(D_SF_SDATE, D_EF_SDATE)
                Me.RP_ENT_P6.DataSource = MB_YENTList.getCurrentDataSet.Tables(0)
                Me.RP_ENT_P6.DataBind()

                Dim D_SE_SDATE As Date = New Date(Me.m_sYear, 7, 1)
                Dim D_EE_SDATE As Date = New Date(Me.m_sYear, 12, 31)
                MB_YENTList.clear()
                MB_YENTList.Load_SDATE(D_SE_SDATE, D_EE_SDATE)
                Me.RP_ENT_E6.DataSource = MB_YENTList.getCurrentDataSet.Tables(0)
                Me.RP_ENT_E6.DataBind()
            End If
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

    Private Sub MBQry_Year_ENT_Unload(sender As Object, e As EventArgs) Handles Me.Unload
        UIShareFun.releaseConnection(Me.m_DBManager)
    End Sub

    Private Sub RP_ENT_P6_ItemDataBound(sender As Object, e As RepeaterItemEventArgs) Handles RP_ENT_P6.ItemDataBound, RP_ENT_E6.ItemDataBound
        If e.Item.ItemType = ListItemType.Item OrElse e.Item.ItemType = ListItemType.AlternatingItem Then
            Dim LTL_SEQ As Literal = e.Item.FindControl("LTL_SEQ")
            If sender.id = "RP_ENT_P6" Then
                LTL_SEQ.Text = e.Item.ItemIndex + 1
            Else
                LTL_SEQ.Text = e.Item.ItemIndex + 1 + Me.RP_ENT_P6.Items.Count
            End If

            Dim LTL_DATE As Literal = e.Item.FindControl("LTL_DATE")
            Dim DRV As DataRowView = CType(e.Item.DataItem, DataRowView)
            Dim D_SDATE As Date = CDate(DRV("SDATE").ToString)
            Dim D_EDATE As Date = CDate(DRV("EDATE").ToString)
            If D_SDATE = D_EDATE Then
                Dim twCalendar As New System.Globalization.TaiwanCalendar
                Dim iDay As Integer = twCalendar.GetDayOfWeek(D_SDATE)
                Dim sDay As String = String.Empty
                Select Case iDay
                    Case 1
                        sDay = "一"
                    Case 2
                        sDay = "二"
                    Case 3
                        sDay = "三"
                    Case 4
                        sDay = "四"
                    Case 5
                        sDay = "五"
                    Case 6
                        sDay = "六"
                    Case 7
                        sDay = "日"
                End Select
                LTL_DATE.Text = D_SDATE.Month & "/" & Utility.FillZero(D_SDATE.Day, 2) & "(" & sDay & ")"
            Else
                If D_SDATE.Month = D_EDATE.Month Then
                    LTL_DATE.Text = D_SDATE.Month & "/" & Utility.FillZero(D_SDATE.Day, 2) & "─" & Utility.FillZero(D_EDATE.Day, 2)
                Else
                    LTL_DATE.Text = D_SDATE.Month & "/" & Utility.FillZero(D_SDATE.Day, 2) & "─" & D_EDATE.Month & "/" & Utility.FillZero(D_EDATE.Day, 2)
                End If
            End If
        End If
    End Sub
End Class