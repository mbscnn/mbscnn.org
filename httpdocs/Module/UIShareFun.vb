Imports com.Azion.EloanUtility
Imports com.Azion.NET.VB
Imports MBSC.MB_OP

Public Class UIShareFun

#Region "DataBaseManager"

#Region "Property"

    Private Shared _sConnectionString As String = getConnectionString()
    Shared ReadOnly Property ConnectionString As String
        Get
            Return _sConnectionString
        End Get
    End Property

#End Region

    ''' <summary>
    ''' 取得連接字串
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function getConnectionString() As String
        Return System.Web.HttpContext.Current.Application.Item("BotDSN")
        '"Data Source=" & com.Azion.EloanUtility.FileUtility.getAppSettings("DBSource") & ";database=" & com.Azion.EloanUtility.FileUtility.getAppSettings("DataBase") & ";User ID=" & com.Azion.EloanUtility.FileUtility.getAppSettings("DBUserID") & ";Persist Security Info=True;Pooling=true;Min Pool Size=" & com.Azion.EloanUtility.FileUtility.getAppSettings("MinPool") & ";Max Pool Size=" & com.Azion.EloanUtility.FileUtility.getAppSettings("MaxPool") & ";Connection Lifetime=" & com.Azion.EloanUtility.FileUtility.getAppSettings("ConnectionLifetime") & ";Connect Timeout=" & com.Azion.EloanUtility.FileUtility.getAppSettings("ConnectTimeout") & ";Application Name=" & com.Azion.EloanUtility.FileUtility.getAppSettings("AppName") & ";Password=" & com.Azion.EloanUtility.EncryptUtility.Decrypto(com.Azion.EloanUtility.FileUtility.getAppSettings("DBPassword"))
    End Function

    ''' <summary>
    ''' 取得DatabaseManager Object
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Shared Function getDataBaseManager() As com.Azion.NET.VB.DatabaseManager
        Return com.Azion.NET.VB.DatabaseManager.getInstance(ConnectionString)
    End Function

    ''' <summary>
    ''' release DatabaseManager
    ''' </summary>
    ''' <param name="dbManager"></param>
    ''' <remarks></remarks>
    Shared Sub releaseConnection(ByRef dbManager As com.Azion.NET.VB.DatabaseManager)
        Try
            If Not IsNothing(dbManager) Then
                dbManager.releaseConnection()
            End If
        Catch ex As Exception
            Throw New com.Azion.NET.VB.BosException(ex, CType(ex.TargetSite, System.Reflection.MethodInfo))
        End Try
    End Sub

#End Region

#Region "錯誤資訊的開關 "
    Public Shared Sub showErrMsg(ByRef page As System.Web.UI.Page, ByVal obj As Object)
        Try
            'If UIShareFun.isTabPostBack(page) Then
            '    Dim pageTab As PageTab = CType(page.FindControl("PageTab"), PageTab)
            '    If TypeOf obj Is Exception Then
            '        pageTab.setException(CType(obj, Exception))
            '    Else
            '        pageTab.setException(New Exception(CType(obj, String)))
            '    End If
            'End If

            Dim lblErrorMsg As Label = CType(page.FindControl("lblErrorMsg"), System.Web.UI.WebControls.Label)
            Dim palErrorMsg As Panel = CType(page.FindControl("ErrorMsg"), System.Web.UI.WebControls.Panel)

            Dim sErrorMsg As String = String.Empty

            If TypeOf obj Is Exception Then
                Dim exception As Exception = CType(obj, Exception)
                Dim sb As New System.Text.StringBuilder

                Do
                    If Utility.isValidateData(exception.Message) Then
                        sb.Append(exception.Message.ToString & "。<br>")
                    End If

                    If Utility.isValidateData(exception.StackTrace) Then
                        sb.Append(exception.StackTrace.ToString & "。<br>")
                    End If

                    exception = exception.InnerException
                Loop Until (exception Is Nothing)

                'ENLogger.log.Debug(obj)
                sErrorMsg = sb.ToString
            Else
                sErrorMsg = CType(obj, String)
            End If

            If Not IsNothing(palErrorMsg) Then
                palErrorMsg.Visible = True
                lblErrorMsg.Visible = True
                lblErrorMsg.Text = sErrorMsg
            Else
                sErrorMsg = "<span style='color:red'>" & sErrorMsg & "</span>"

                HttpContext.Current.Response.Write(sErrorMsg)
            End If
        Catch ex As Exception
            Dim s As String = CType(obj, Exception).Message.ToString & "<br>" & CType(obj, Exception).StackTrace
        End Try
    End Sub

    Public Shared Sub closeErrMsg()
        Dim page As Page = CType(HttpContext.Current.Handler, Page)
        Try
            Dim lblErrorMsg As Label = CType(page.FindControl("lblErrorMsg"), System.Web.UI.WebControls.Label)
            Dim ErrorMsg As Panel = CType(page.FindControl("ErrorMsg"), System.Web.UI.WebControls.Panel)

            lblErrorMsg.Text = ""
            ErrorMsg.Visible = False
        Catch ex As Exception

        End Try
    End Sub

    Public Shared Function getEXInfo(ByVal obj As Object) As String
        Dim sErrorMsg As String = String.Empty

        If TypeOf obj Is Exception Then
            Dim exception As Exception = CType(obj, Exception)
            Dim sb As New System.Text.StringBuilder

            Do
                If Utility.isValidateData(exception.Message) Then
                    sb.Append(exception.Message.ToString & "。<br>")
                End If

                If Utility.isValidateData(exception.StackTrace) Then
                    sb.Append(exception.StackTrace.ToString & "。<br>")
                End If

                exception = exception.InnerException
            Loop Until (exception Is Nothing)

            sErrorMsg = sb.ToString
        Else
            sErrorMsg = CType(obj, String)
        End If

        Return sErrorMsg
    End Function
#End Region

#Region "Web"
    Public Shared Function getLoginUserid() As String
        Dim httpContext As System.Web.HttpContext = System.Web.HttpContext.Current
        Dim sUserId As String = ""

        If IsNothing(httpContext) Then Return sUserId

        If Utility.isValidateData(httpContext.Current.Session("userid")) Then
            sUserId = httpContext.Current.Session("userid")
        ElseIf Utility.isValidateData(httpContext.Current.Request("userid")) Then
            sUserId = httpContext.Current.Request("userid")
            httpContext.Current.Session("userid") = sUserId
        End If

        Return sUserId
    End Function

    Public Shared Function getLoginBrid() As String
        Dim httpContext As System.Web.HttpContext = System.Web.HttpContext.Current
        Dim sBrid As String = ""

        If IsNothing(httpContext) Then Return sBrid

        If Utility.isValidateData(httpContext.Current.Session("brid")) Then
            sBrid = httpContext.Current.Session("brid")
        ElseIf Utility.isValidateData(httpContext.Current.Request("brid")) Then
            sBrid = httpContext.Current.Request("brid")
            httpContext.Current.Session("brid") = sBrid
        End If

        Return sBrid
    End Function

    Public Shared Function getMailBody(ByVal sUID As String, ByVal sMail As String, ByVal sAPPNAME As String) As String
        Try
            Dim sb As New StringBuilder
            sb.Append("<P>")
            sb.Append("親愛的" & sAPPNAME & "您好：<BR/>")
            sb.Append("這封信是由佛陀原始正法中心發送的。<BR/>")
            sb.Append("非常感謝您註冊成為我們的會員。<BR/>")
            sb.Append("您需點擊下面的連結啟用您的帳號：<BR/>")
            'sb.Append("<a href='http://www.mbscnn.org/mbsc/mnt/MBMnt_Reg_01_v01.aspx?MOD=APV&UID=" & sUID & "&MAIL=" & sMail & "'>http://www.mbscnn.org/mbsc/mnt/MBMnt_Reg_01_v01.aspx?MOD=APV&UID=" & sUID & "&MAIL=" & sMail & "</a>")
            'sb.Append("<a href='http://www.mbscnn.org/mbsc/mnt/MBMnt_Reg_01_v01.aspx?MOD=APV&UID=" & sUID & "'>http://www.mbscnn.org/mbsc/mnt/MBMnt_Reg_01_v01.aspx?MOD=APV&UID=" & sUID & "</a>")
            'sb.Append("http://www.mbscnn.org/mbsc/mnt/MBMnt_Reg_01_v01.aspx?MOD=APV&UID=" & sUID)
            sb.Append("<a href='http://www.mbscnn.org/mnt/MBMnt_Reg_01_v01.aspx?MOD=APV&UID=" & sUID & "'>http://www.mbscnn.org/mnt/MBMnt_Reg_01_v01.aspx?MOD=APV&UID=" & sUID & "</a>")
            sb.Append("<BR/>")
            sb.Append("<span style='color:red'>(如果上面不是連結形式，請將該位址貼到瀏覽器網址欄)</span>")
            sb.Append("<BR/>")
            sb.Append("帳號啟用後，請您記得至會員系統填寫入會申請單！")
            sb.Append("<BR/>")
            sb.Append("感謝您的訪問，祝順心愉快！")
            sb.Append("<BR/>")
            sb.Append("此致")
            sb.Append("</P>")

            Return sb.ToString
        Catch ex As Exception
            Throw
        End Try
    End Function

#End Region

#Region "會員取號"
    Public Shared Function getVadID(ByVal dbManager As com.Azion.NET.VB.DatabaseManager) As String
        Try
            Dim sProcName As String = String.Empty
            sProcName = LCase(com.Azion.NET.VB.Properties.getSchemaName) & ".MB_GET_MAXID"
            Dim inParaAL As New ArrayList
            Dim outParaAL As New ArrayList
            inParaAL.Add("01")
            inParaAL.Add("1")
            outParaAL.Add(7)
            Dim HT_Return As Hashtable = com.Azion.NET.VB.DBObject.getProcedureData(dbManager, sProcName, inParaAL, outParaAL)
            Dim iMAXID As Decimal = HT_Return.Item("@IMAXID")
            Return com.Azion.EloanUtility.StrUtility.FillZero(iMAXID, 7)
        Catch ex As Exception
            Throw
        Finally
        End Try
    End Function

#End Region

    Public Shared Sub downLoadFile(ByVal page As System.Web.UI.Page, ByVal sFileName As String, ByVal str As String)
        page.Response.ClearHeaders()
        page.Response.Clear()
        page.Response.Expires = 0
        page.Response.Buffer = True
        '原本有編碼的問題,出來的EXCEL內容會出現亂碼,是因為系統採"繁體中文BIG5"的關係,改成UTF-8就行了[原因不明,待查]
        page.Response.ContentEncoding = System.Text.Encoding.UTF8
        page.Response.AddHeader("Accept-Language", "zh-tw")
        page.Response.ContentType = "application/octet-stream; charset=iso-8859-1"  '"Application/octet-stream"
        'page.Response.ContentType = "Application/vnd.ms-excel" 'GAQryCase02的 System.Text.Encoding.UTF8.GetString(System.Text.Encoding.UTF8.GetBytes(sFileName))
        page.Response.AddHeader("content-disposition", "attachment; filename=" & Chr(34) & System.Web.HttpUtility.UrlEncode(sFileName, System.Text.Encoding.UTF8) & Chr(34))
        page.Response.Write(str)
        page.Response.End()
    End Sub


End Class
