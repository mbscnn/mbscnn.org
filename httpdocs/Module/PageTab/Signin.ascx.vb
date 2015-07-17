Imports MBSC.MB_OP

Public Class Signin
    Inherits System.Web.UI.UserControl

    Dim m_DBManager As com.Azion.NET.VB.DatabaseManager

    Dim m_iUPCODE_LV As Integer = 37

    Dim m_iUPCODE_78 As Integer = 78

    Dim m_iLEVEL As Integer = 0

    Dim m_sCODEID As String = String.Empty

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If IsNumeric(Session("admin")) OrElse com.Azion.EloanUtility.ValidateUtility.isValidateData(Session("USERID")) Then
                'Me.PLH_LOGIN.Visible = False
                Me.PLH_LOGOUT.Visible = True
            Else
                'Me.PLH_LOGIN.Visible = True
                Me.PLH_LOGOUT.Visible = False
            End If

            Me.m_DBManager = UIShareFun.getDataBaseManager

            If IsNumeric(Session("admin")) Then
                Me.m_iLEVEL = CInt(Session("admin"))
            End If

            Me.m_sCODEID = "" & Request.QueryString("CLASS")

            '會員系統
            Me.m_iUPCODE_78 = com.Azion.EloanUtility.CodeList.getAppSettings("UPCODE78")

            If Not Page.IsPostBack Then
                If com.Azion.EloanUtility.ValidateUtility.isValidateData(Session("USERID")) Then
                    Dim mbACCT As New MB_ACCT(Me.m_DBManager)
                    If mbACCT.LoadByPK(Session("USERID")) Then
                        Me.LTL_MB_NAME.Text = mbACCT.getString("MB_NAME")
                    End If
                End If

                'Me.Bind_RP_Tab()

                'Me.IMGTop.ImageUrl = com.Azion.EloanUtility.UIUtility.getRootPath & "/img/mbscbanner02.jpg"

                If Not com.Azion.EloanUtility.ValidateUtility.isValidateData(Me.m_sCODEID) Then
                    Dim apCODEList As New AP_CODEList(Me.m_DBManager)
                    apCODEList.setSQLCondition(" ORDER BY SORTNO ")
                    apCODEList.loadByUpCode(Me.m_iUPCODE_LV)
                    Dim ROW_LV1() As DataRow = apCODEList.getCurrentDataSet.Tables(0).Select("DISABLED='1'")
                    If Not IsNothing(ROW_LV1) AndAlso ROW_LV1.Length > 0 Then
                        Me.m_sCODEID = ROW_LV1(0)("CODEID")
                    End If
                End If
                Dim apCODE As New AP_CODE(Me.m_DBManager)
                If apCODE.loadByPK(Me.m_sCODEID) Then
                    Me.LTL_TAB_TITLE.Text = apCODE.getString("TEXT")
                    Me.LTL_TAB_TITLE_HID.Text = apCODE.getString("TEXT")

                    'If Utility.isValidateData(apCODE.getString("TMPURL")) Then
                    '    Dim sScript As String = String.Empty
                    '    sScript = "<script language='javascript' >" & vbCrLf
                    '    sScript &= "window.open('" & apCODE.getString("TMPURL") & "',null,'height=' + window.screen.height + ', width=' + window.screen.width + ', top=0, left=0, toolbar=yes, menubar=yes, scrollbars=no, resizable=no,location=n o, status=no');" & vbCrLf
                    '    sScript &= "</" & "script>"
                    '    Response.Write(sScript)
                    'End If
                End If
            End If
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me.Page, ex)
        End Try
    End Sub

    Private Sub Page_Unload(sender As Object, e As System.EventArgs) Handles Me.Unload
        UIShareFun.releaseConnection(Me.m_DBManager)
    End Sub

    Private Sub btLogout_Click(sender As Object, e As System.EventArgs) Handles btLogout.Click
        Session("admin") = Nothing

        Session("USERID") = Nothing

        Session("BRID") = Nothing

        Session("AREA") = Nothing

        'Response.Redirect(Request.AppRelativeCurrentExecutionFilePath & Request.Url.Query)
        Response.Redirect(com.Azion.EloanUtility.UIUtility.getRootPath & "/NewsList.aspx")
    End Sub
End Class