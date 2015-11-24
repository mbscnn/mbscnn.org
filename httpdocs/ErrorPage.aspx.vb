Public Class ErrorPage
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim ex As Exception = Server.GetLastError()
     
        Dim sErr As String = Request.QueryString("ERR")

        If Not IsNothing(sErr) AndAlso sErr.Length > 0 Then
            'labMessage.Text = sErr
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, sErr)
            Exit Sub
        End If

        com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
    End Sub

End Class