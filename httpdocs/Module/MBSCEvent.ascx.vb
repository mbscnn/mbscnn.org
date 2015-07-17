Imports com.Azion.NET.VB
Imports MBSC.MB_OP

Public Class MBSCEvent
    Inherits System.Web.UI.UserControl

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Not Page.IsPostBack Then
                Dim dbManager As DatabaseManager = MBSC.UICtl.UIShareFun.getDataBaseManager()
                Try
                    Dim mbNEWSList As New MB_NEWSList(dbManager)
                    mbNEWSList.LoadEvent()
                    Me.RP_EVENT.DataSource = mbNEWSList.getCurrentDataSet.Tables(0)
                    Me.RP_EVENT.DataBind()

                    If mbNEWSList.getCurrentDataSet.Tables(0).Rows.Count = 0 Then
                        Dim TD_EVENT As HtmlTableCell = Me.Page.FindControl("TD_EVENT")
                        TD_EVENT.Visible = False
                    End If
                Catch ex As Exception
                    Throw
                Finally
                    MBSC.UICtl.UIShareFun.releaseConnection(dbManager)
                End Try
            End If
        Catch ex As Exception
            MBSC.UICtl.UIShareFun.showErrMsg(Me.Page, ex)
        End Try
    End Sub

    Private Sub RP_EVENT_ItemCommand(source As Object, e As RepeaterCommandEventArgs) Handles RP_EVENT.ItemCommand
        Try
            Dim sURL As String = String.Empty
            sURL = com.Azion.EloanUtility.UIUtility.getRootPath & "/NewsList.aspx?CRETIME=" & e.CommandArgument
            Response.Redirect(sURL)
        Catch ex As System.Threading.ThreadAbortException
        Catch ex As Exception
            MBSC.UICtl.UIShareFun.showErrMsg(Me.Page, ex)
        End Try
    End Sub
End Class