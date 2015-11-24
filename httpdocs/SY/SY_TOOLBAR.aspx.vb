Imports com.Azion.EloanUtility

''' <summary>
''' 程式說明：
''' 建立者：Lake
''' 建立日期：2012-04-06
''' </summary>

Imports com.Azion.NET.VB
Imports AUTH_OP

Public Class SY_TOOLBAR
    Inherits SYUIBase

    Public m_SSList As New System.Text.StringBuilder

#Region " Page Load "

    ''' <summary>
    ''' 頁面加載
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            monitorTimeOut()
            Response.AddHeader("cache-control", "private")
            Response.AddHeader("P3P", "CP='CAO PSA OUR'")
            Response.ExpiresAbsolute = Now.AddDays(-1)
            Response.Expires = 0 'prevents caching at the proxy server 

            ' 初始化參數
            initParas()

        Catch ex As Exception
            handleException(ex)
        End Try
    End Sub
#End Region


    ''' <summary>
    '''  Session出現超時則需要重新登陸頁面
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub monitorTimeOut()
        Dim bTimeOut As Boolean = False

        ' 判斷Session是否為空
        If IsNothing(Session("StaffInfo")) Then
            bTimeOut = True
        End If

        ' 如果為空 則跳轉到登陸頁面重新登入
        If bTimeOut Then
            Dim sbTimeout As New System.Text.StringBuilder
            sbTimeout.Append("<SCRIPT LANGUAGE='JAVASCRIPT'>" & vbCrLf)
            sbTimeout.Append("alert('作業逾時,請重新登入!');" & vbCrLf)
            sbTimeout.Append("window.top.location = 'SY_DEFAULT.aspx'" & vbCrLf)
            sbTimeout.Append("</" & "SCRIPT>" & vbCrLf)
            Response.Write(HttpUtility.JavaScriptStringEncode(sbTimeout.ToString))
        End If
    End Sub

#Region "Function"

    ''' <summary>
    ''' 初始化參數
    ''' </summary>
    ''' <remarks></remarks>
    ''' <history>
    ''' [Lake] 2012/04/06 Created
    ''' </history>
    Sub initParas()
          
        Try
            Dim sbSysList As New StringBuilder()
            Dim sySysIdList As New SY_SYSIDList(GetDatabaseManager())

            If Not m_sWorkingRoleId Is Nothing Then

                ' 根據登錄人員角色取得對應的系統功能
                If sySysIdList.loadByWorkingRoleId(m_sWorkingRoleId) Then
                    Dim dtSysIdList As DataTable = sySysIdList.getCurrentDataSet().Tables(0)

                    For i As Integer = 0 To 7
                        If i <= dtSysIdList.Rows.Count - 1 Then
                            For Each row As DataRow In dtSysIdList.Rows
                                m_SSList.Append("<td class='smallbar' width='70' style='vertical-align: top'><a href='SY_FUNCTIONLIST.aspx?sysid=" & row("SYSID") & "' target='leftFrame'>" & row("SYSNAME") & "</a></td>" & vbCrLf)
                                m_SSList.Append("<td class='smallbar' width='18'><img src='" & com.Azion.EloanUtility.UIUtility.getImgPath() & "top_blueline.gif' width='18' height='17' align='baseline' hspace='0' vspace='0'></td>" & vbCrLf)
                                i = i + 1
                            Next
                        Else
                            m_SSList.Append("<td class='smallbar' width='70'></td>" & vbCrLf)
                            m_SSList.Append("<td class='smallbar' width='18'><img src='" & com.Azion.EloanUtility.UIUtility.getImgPath() & "top_blueline.gif' width='18' height='17' align='baseline' hspace='0' vspace='0'></td>" & vbCrLf)
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            SYUIBase.showErrMsg(Me, ex)
        End Try
    End Sub
#End Region

#Region " Event"

    ''' <summary>
    ''' 錯誤信息顯示
    ''' </summary>
    ''' <param name="_ex"></param>
    ''' <remarks></remarks>
    Private Sub handleException(ByVal _ex As Exception)
        Dim sMsg As String = "[訊息]: " & _ex.ToString & "<BR>[追蹤訊息]: " & _ex.StackTrace
        outError.InnerHtml = sMsg
    End Sub
#End Region
End Class