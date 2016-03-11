Public Class MBMnt_ATO_MEMBER
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Private Sub btnATO_Click(sender As Object, e As EventArgs) Handles btnATO.Click
        Try

        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me,ex)
        End Try
    End Sub
End Class