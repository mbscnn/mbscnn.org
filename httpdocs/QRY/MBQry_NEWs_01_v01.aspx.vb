Imports com.Azion.NET.VB
Imports MBSC.MB_OP
Imports MBSC.UICtl

Public Class MBQry_NEWs_01_v01
    Inherits System.Web.UI.Page

    Dim m_DBManager As DatabaseManager = Nothing
    Dim m_sCRETIME As String = String.Empty

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            'Clear Cache
            Response.Expires = 0
            Response.Cache.SetNoStore()
            Response.AppendHeader("Pragma", "no-cache")

            Me.m_DBManager = UIShareFun.getDataBaseManager
            Me.m_sCRETIME = Request.QueryString("CRETIME")
            If Not Page.IsPostBack Then
                If IsNumeric(Me.m_sCRETIME) Then
                    Dim mbNEWS As New MB_NEWS(Me.m_DBManager)
                    If mbNEWS.LoadByPK(CDec(Me.m_sCRETIME)) Then
                        Me.LTL_CNTHTML.Text = mbNEWS.getString("CNTHTML")
                    Else
                        com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "查無資料")
                    End If
                Else
                    com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "查無資料")
                End If
            End If
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

    Private Sub MBQry_NEWs_01_v01_Unload(sender As Object, e As EventArgs) Handles Me.Unload
        UIShareFun.releaseConnection(Me.m_DBManager)
    End Sub
End Class