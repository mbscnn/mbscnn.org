Imports com.Azion.NET.VB
Imports MBSC.MB_OP
Imports MBSC.UICtl

Public Class MBMnt_RESP_01_v01
    Inherits System.Web.UI.Page

    Dim m_DBManager As DatabaseManager = Nothing

    Dim m_sMB_MEMSEQ As String = String.Empty
    Dim m_sMB_SEQ As String = String.Empty
    Dim m_sMB_BATCH As String = String.Empty
    Dim m_sOPTYPE As String = String.Empty

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.m_DBManager = UIShareFun.getDataBaseManager

            Me.m_sMB_MEMSEQ = "" & Request.QueryString("MB_MEMSEQ")

            Me.m_sMB_SEQ = "" & Request.QueryString("MB_SEQ")

            Me.m_sMB_BATCH = "" & Request.QueryString("MB_BATCH")

            Me.m_sOPTYPE = "" & Request.QueryString("OPTYPE")

            If IsNumeric(Me.m_sMB_MEMSEQ) AndAlso IsNumeric(Me.m_sMB_SEQ) AndAlso IsNumeric(Me.m_sMB_BATCH) AndAlso Utility.isValidateData(Me.m_sOPTYPE) Then
                Dim MB_MEMCLASS As New MB_MEMCLASS(Me.m_DBManager)
                If MB_MEMCLASS.LoadByPK(Me.m_sMB_MEMSEQ, Me.m_sMB_SEQ, Me.m_sMB_BATCH) Then
                    MB_MEMCLASS.setAttribute("MB_RESP", UCase(Me.m_sOPTYPE))
                    MB_MEMCLASS.save()

                    Dim MB_CLASS As New MB_CLASS(Me.m_DBManager)
                    MB_CLASS.loadByPK(Me.m_sMB_SEQ, Me.m_sMB_BATCH)
                    Dim sMsg As String = String.Empty
                    If Me.m_sOPTYPE = "Y" Then
                        sMsg = "要出席"
                    ElseIf Me.m_sOPTYPE = "N" Then
                        sMsg = "不出席"
                    End If
                    sMsg &= MB_CLASS.getString("MB_CLASS_NAME")

                    com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "已確認您" & sMsg)

                    Dim sURL As String = String.Empty
                    sURL = Request.ApplicationPath() & "/NewsList.aspx"
                    Server.Transfer(sURL)
                Else
                    com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "查無報名資料【" & Me.m_sMB_MEMSEQ & "】【" & Me.m_sMB_SEQ & "】【" & Me.m_sMB_BATCH & "】【" & Me.m_sOPTYPE & "】")
                End If
            Else
                com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "參數錯誤【" & Me.m_sMB_MEMSEQ & "】【" & Me.m_sMB_SEQ & "】【" & Me.m_sMB_BATCH & "】【" & Me.m_sOPTYPE & "】")
            End If
        Catch ex As System.Threading.ThreadAbortException
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

    Private Sub MBMnt_RESP_01_v01_Unload(sender As Object, e As EventArgs) Handles Me.Unload
        UIShareFun.releaseConnection(Me.m_DBManager)
    End Sub
End Class