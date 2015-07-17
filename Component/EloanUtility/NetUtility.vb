Option Explicit On
Option Strict On

Imports System.IO
Imports System.Net
Imports System.Net.Sockets
Imports System.Web
Imports System.Net.Mail

''' <summary>
''' 提供網路相關函式
''' Ftp、HTTP....
''' [Titan] 	2011/07/19	Created
''' </summary>
Public Class NetUtility

    ''' <summary>
    ''' fileName上傳的檔案
    ''' </summary>
    ''' <param name="sFileNmae"> c:\abc.xml</param>
    ''' <param name="sUploadURL">ftp://127.0.0.1</param>
    ''' <param name="sUserName">使用者FTP登入帳號</param>
    ''' <param name="sPassword">使用者登入密碼</param>
    ''' <remarks>
    ''' [Titan] 	2011/07/19	Created
    ''' </remarks>  
    ''' <exception cref="Sockets.SocketException"></exception>
    Public Shared Sub FtpUpload(ByVal sFileNmae As String, ByVal sUploadURL As String, ByVal sUserName As String, ByVal sPassword As String)
        Dim fi As FileInfo

        Dim uploadResponse As FtpWebResponse = Nothing

        Try
            If My.Computer.FileSystem.FileExists(sFileNmae) Then
                fi = New FileInfo(sFileNmae)
            Else
                Throw New FileNotFoundException("File Not Found", sFileNmae)
            End If

            If sUploadURL.Contains("ftp://") Then
                sUploadURL = sUploadURL.Trim
                sUploadURL = sUploadURL.Split(CChar(":"))(1).Replace("//", "") & "/"
                sUploadURL = sUploadURL.Replace("//", "/")
                sUploadURL = String.Format("ftp://{0}{1}", sUploadURL, fi.Name)
            Else
                sUploadURL = sUploadURL.Trim & "/"
                sUploadURL = sUploadURL.Replace("//", "/")
                sUploadURL = String.Format("ftp://{0}{1}", sUploadURL, fi.Name)
            End If

            Dim uploadRequest As FtpWebRequest = CType(WebRequest.Create(sUploadURL), FtpWebRequest)
            uploadRequest.Method = WebRequestMethods.Ftp.UploadFile '設定Method上傳檔案
            uploadRequest.Proxy = Nothing
            uploadRequest.KeepAlive = False
            uploadRequest.UseBinary = True
            uploadRequest.ContentLength = fi.Length

            If (sUserName.Length > 0) Then '如果需要帳號登入
                Dim nc As NetworkCredential = New NetworkCredential(sUserName, sPassword)
                uploadRequest.Credentials = nc
            End If

            Const bufferLength As Integer = 2048
            Dim buffer(bufferLength) As Byte
            Dim contentLen As Integer

            Using fileStream = File.Open(sFileNmae, FileMode.Open)
                Try
                    contentLen = fileStream.Read(buffer, 0, bufferLength)
                    Using requestStream = uploadRequest.GetRequestStream()
                        Try
                            While Not (contentLen = 0)
                                requestStream.Write(buffer, 0, contentLen)
                                contentLen = fileStream.Read(buffer, 0, bufferLength)
                            End While
                        Catch ex As Exception
                            Throw ex
                        Finally
                            If Not IsNothing(requestStream) Then requestStream.Close()
                        End Try
                    End Using
                Catch ex As Exception
                    Throw ex
                Finally
                    If Not IsNothing(fileStream) Then fileStream.Close()
                End Try
            End Using

        Catch ex As Sockets.SocketException
            Throw ex
        Catch ex As Exception
            Throw ex
        Finally
            If Not IsNothing(uploadResponse) Then uploadResponse.Close()
        End Try
    End Sub

    ''' <summary>
    ''' ping tool
    ''' </summary>
    ''' <param name="hostNameOrAddress">IP or HostName</param>
    ''' <param name="portnum">port number</param>
    ''' <returns>Boolean</returns>
    ''' <remarks>
    ''' [Titan] 	2011/07/19	Created
    ''' </remarks>  
    ''' <exception cref="Sockets.SocketException"></exception>
    Public Shared Function pingSite(ByVal hostNameOrAddress As String, ByVal portnum As Integer) As Boolean
        Try
            Dim sock As System.Net.Sockets.Socket = New System.Net.Sockets.Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
            'Console.WriteLine("Testing " & hostNameOrAddress & ":" & portnum & "...")

            sock.Connect(hostNameOrAddress, portnum)
            If (sock.Connected = True) Then
                sock.Close()
                sock = Nothing
                Return True
            End If
        Catch ex As SocketException
            Throw ex
        End Try

        Return False

    End Function

    ''' <summary>
    ''' socket client
    ''' </summary>
    ''' <param name="hostNameOrAddress">IP or HostName</param>
    ''' <param name="portnum">port number</param>
    ''' <param name="sMsg">Msg</param>
    ''' <param name="iTimeOut">Time Out</param>
    ''' <returns>Boolean</returns>
    ''' <remarks>
    ''' [Titan] 	2011/07/19	Created
    ''' </remarks>  
    ''' <exception cref="Sockets.SocketException"></exception>
    Public Shared Function socketSend(ByVal hostNameOrAddress As String, ByVal portnum As Integer, ByVal sMsg As String, Optional iTimeOut As Integer = 0) As String

        Using clientSocket As New System.Net.Sockets.TcpClient()

            Try
                clientSocket.Connect(hostNameOrAddress, portnum)
                clientSocket.SendTimeout = iTimeOut

                Using serverStream As NetworkStream = clientSocket.GetStream()

                    If serverStream.CanWrite Then
                        Dim outStream As Byte() = System.Text.Encoding.Default.GetBytes(sMsg)
                        serverStream.Write(outStream, 0, outStream.Length)
                        serverStream.Flush()
                    End If

                    If serverStream.CanRead Then
                        Dim inStream(3800) As Byte
                        Dim sServerMsg As Text.StringBuilder = New Text.StringBuilder()
                        Dim i As Integer = 0
                        ' Incoming message may be larger than the buffer size.
                        Do
                            i = serverStream.Read(inStream, 0, inStream.Length)
                            sServerMsg.AppendFormat("{0}", System.Text.Encoding.Default.GetString(inStream, 0, i))
                        Loop While serverStream.DataAvailable

                        Return sServerMsg.ToString()
                    Else
                        Console.WriteLine("Sorry.  You cannot read from this NetworkStream.")
                        Return String.Empty
                    End If
                End Using
            Catch ex As System.Net.Sockets.SocketException
                Throw ex
            Catch ex As Exception
                Throw ex
            End Try
        End Using
    End Function


    ''' <summary>
    ''' 檔案上傳，將硬碟檔案上傳至主機
    ''' </summary>
    ''' <param name="sCaseID">案件編號</param>
    ''' <param name="sCASEID_GA">擔保品案件編號</param>
    ''' <param name="sUserID">使用者帳號</param>
    ''' <param name="sBRID">使用者分行</param>
    ''' <param name="sSUBSYSID">子系統編號</param>
    ''' <param name="objPostedFile">上傳檔案</param>
    ''' <returns>String</returns>
    ''' <remarks>
    ''' [Dora] 	2011/08/05	Created
    ''' </remarks>  
    Public Shared Function UploadByIDMService(ByVal sCaseID As String, ByVal sCASEID_GA As String, ByVal sUserID As String, ByVal sBRID As String, ByVal sSUBSYSID As String, ByVal objPostedFile As HttpPostedFile) As String
        Try
            ImgUtility.checkFileSize(objPostedFile)
            Dim sURL As String = ""

            If UCase(Left(sUserID, 1)) = "S" Then
                sUserID = Right(sUserID, 6)
            End If

            Dim sPAGECOUNT As String = "0"
            Dim sBIZ_CODE As String = ""
            Dim sTASK_CODE As String = ""
            Dim sDOC_CODE As String = ""
            Dim sDOCNAME As String = ""
            If Not ValidateUtility.isValidateData(sCASEID_GA) Then
                sCASEID_GA = "NA"
            End If

            sURL = FileUtility.getAppSettings("WSURL_IDM_ELOAN")
            sURL &= "UploadFile.aspx?CASEID=" & sCaseID & _
                   "&FILENAME=" & System.IO.Path.GetFileName(objPostedFile.FileName) & _
                   "&BRID=" & sBRID & "&SUBSYSID=" & sSUBSYSID & "&PAGECOUNT=" & sPAGECOUNT & "&FILESRC=UP&FILEVALID=Y&BIZ_CODE=" & sBIZ_CODE & _
                   "&TASK_CODE=" & sTASK_CODE & "&DOC_CODE=" & sDOC_CODE & _
                   "&STAFFID=" & sUserID & "&DOCNAME=" & sDOCNAME & "&CASEID_GA=" & sCASEID_GA

            Dim objWebClient As New System.Net.WebClient()
            ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf RemoteCertificateValidationCallback)

            'objWebClient = New System.Net.WebClient()
            Dim objStream As System.IO.Stream = objPostedFile.InputStream
            Dim byteData(CInt(objStream.Length - 1)) As Byte
            objStream.Read(byteData, 0, CInt(objStream.Length))
            Dim byeWebClient() As Byte = objWebClient.UploadData(sURL, byteData)

            Dim sReturn As String = System.Text.Encoding.UTF8.GetString(byeWebClient)

            Return sReturn
        Catch ex As System.Threading.ThreadAbortException

        Catch ex As Exception
            Throw ex

        End Try

        Return String.Empty
    End Function



    ''' <summary>
    ''' 檔案刪除，根據傳入的案件編號、擔保品案件編號以及檔案名稱，將已上傳檔案進行刪除
    ''' </summary>
    ''' <param name="sCaseID">案件編號</param>
    ''' <param name="sCASEID_GA">擔保品案件編號</param>
    ''' <param name="sFILENAME">檔案名稱</param>
    ''' <returns>String</returns>
    ''' <remarks>
    ''' [Dora] 	2011/08/05	Created
    ''' </remarks>  
    Public Shared Function DeleteFile(ByVal sCaseID As String, ByVal sCASEID_GA As String, ByVal sFILENAME As String) As String
        Try
            If Not ValidateUtility.isValidateData(sCASEID_GA) Then
                sCASEID_GA = "NA"
            End If

            Dim sURL As String = FileUtility.getAppSettings("WSURL_IDM_ELOAN")
            sURL &= "DeleteFile.aspx?CASEID=" & sCaseID & _
                                 "&CASEID_GA=" & sCASEID_GA & "&FILENAME=" & sFILENAME & "&DEL_TYPE=DELETE"

            Dim objWebClient As New System.Net.WebClient
            ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf RemoteCertificateValidationCallback)

            Dim byeWebClient() As Byte = objWebClient.DownloadData(sURL)
            Return System.Text.Encoding.UTF8.GetString(byeWebClient)
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 檔案下載，開啟已上傳的檔案
    ''' </summary>
    ''' <param name="sCaseID">案件編號</param>
    ''' <param name="sCASEID_GA">擔保品案件編號</param>
    ''' <param name="sFILENAME">檔案名稱</param>
    ''' <remarks>
    ''' [Dora] 	2011/08/05	Created
    ''' </remarks>  
    Public Shared Sub DownloadFile(ByVal sCaseID As String, ByVal sCASEID_GA As String, ByVal sFILENAME As String)
        Try
            If Not ValidateUtility.isValidateData(sCASEID_GA) Then
                sCASEID_GA = "NA"
            End If
            Dim IdmUrl As String = FileUtility.getAppSettings("WSURL_IDM_ELOAN")
            Dim sURL As String = IdmUrl & "DownloadFile.aspx?CASEID=" & sCaseID & "&CASEID_GA=" & sCASEID_GA & "&FILENAME=" & sFILENAME

            Dim objWebClient As New System.Net.WebClient()
            ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf RemoteCertificateValidationCallback)

            Dim TifimgPath As String = FileUtility.getAppSettings("TifimgPath")
            Dim sOpenAddress As String = TifimgPath & sFILENAME
            objWebClient.DownloadFile(sURL, sOpenAddress)

            Dim ImageURL As String = FileUtility.getAppSettings("ImageURL")
            UIUtility.showOpen(ImageURL & sFILENAME)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Shared Function downloadIDMFile(ByVal sCaseID As String, ByVal sCASEID_GA As String, ByVal sFILENAME As String) As Byte()
        Try
            If Not ValidateUtility.isValidateData(sCASEID_GA) Then
                sCASEID_GA = "NA"
            End If
            Dim IdmUrl As String = FileUtility.getAppSettings("WSURL_IDM_ELOAN")
            Dim sURL As String = IdmUrl & "DownloadFile.aspx?CASEID=" & sCaseID & "&CASEID_GA=" & sCASEID_GA & "&FILENAME=" & sFILENAME

            Dim webClient As New System.Net.WebClient()
            ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf RemoteCertificateValidationCallback)
             
            Return webClient.DownloadData(sURL)

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Shared Function downloadeLandFile(ByVal sURL As String) As Byte()
        Try
            Dim webClient As New System.Net.WebClient()
            ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf RemoteCertificateValidationCallback)

            Return webClient.DownloadData(sURL)

        Catch ex As Exception
            Throw ex
        End Try
    End Function


    Private Shared Function RemoteCertificateValidationCallback(ByVal sender As Object, ByVal certificate As System.Security.Cryptography.X509Certificates.X509Certificate, ByVal chain As System.Security.Cryptography.X509Certificates.X509Chain, ByVal sslPolicyErrors As Net.Security.SslPolicyErrors) As Boolean
        Return True
    End Function

#Region "send mail"
    Public Shared Sub sendMail(ByVal sSubject As String, ByVal sBody As String, ByVal sToAddress As String, ByVal sFromAddress As String, ByVal sSMTP As String, ByVal sSMTPAcc As String, ByVal sSMTPPwd As String, ByVal sSMTPPort As String, Optional sFromName As String = "", Optional isSSL As Boolean = False)
        ' Mail Message Setting 
        Dim fromEmail As String = sFromAddress

        Dim from As New MailAddress(fromEmail, sFromName, System.Text.Encoding.UTF8)

        If com.Azion.EloanUtility.ValidateUtility.isValidateData(sToAddress) Then
            sToAddress = sToAddress.Trim
            'Dim s As String = String.Empty

            'For Each sToAdd As String In sToAddress.Split(CChar(","))
            '    sToAdd = sToAdd.Trim.Insert(0, "<")
            '    sToAdd = sToAdd.Trim.Insert(sToAdd.Length, ">")
            '    s = s & sToAdd & ","
            'Next
            'sToAddress = s.Remove(s.Length() - 1)
        Else
            Exit Sub
        End If

        'Dim MailAddress As MailAddress = New MailAddress("amyh@azion.com.tw", "janehsu@azion.com.tw", "michellehsieh36@yahoo.com.tw", "titanchen@azion.com.tw")

        'Dim mail As New MailMessage(from, New MailAddress(sToAddress, sSubject, System.Text.Encoding.Default))
        Dim mail As New MailMessage()
        'Dim subject As String = "Test Subject"
        mail.Subject = sSubject
        mail.SubjectEncoding = System.Text.Encoding.UTF8

        'Dim body As String = "Boteloan Mseeage"
        mail.Body = sBody
        mail.BodyEncoding = System.Text.Encoding.UTF8
        mail.IsBodyHtml = False
        mail.Priority = MailPriority.High
        'mail.Attachments.Add(New Mail.Attachment("c:\temp.jpg"))

        mail.IsBodyHtml = False '是否為HTML格式 
        mail.Priority = System.Net.Mail.MailPriority.Normal

        ' mail.To.Add(sToAddress)
        'SMTP Setting 
        Dim client As New SmtpClient()
        client.Host = sSMTP

        If com.Azion.EloanUtility.ValidateUtility.isValidateData(sSMTPPort) Then
            client.Port = CInt(sSMTPPort)
        End If

        If com.Azion.EloanUtility.ValidateUtility.isValidateData(sSMTPAcc) And com.Azion.EloanUtility.ValidateUtility.isValidateData(sSMTPPwd) Then
            client.Credentials = New NetworkCredential(sSMTPAcc, sSMTPPwd)
        End If

        client.EnableSsl = isSSL

        mail.From = from

        For Each sToAdd As String In sToAddress.Split(CChar(";"))
            If sToAdd.Length > 1 Then
                mail.To.Add(New MailAddress(sToAdd, sToAdd.Split(CChar("@"))(0), System.Text.Encoding.Default))
            End If
        Next
        ' Send Mail 
        Try
            client.Send(mail)
        Catch ex As Exception
            Throw ex
        End Try

        AddHandler client.SendCompleted, AddressOf client_SendCompleted

        ' Sent Compeleted Eevet 
    End Sub

    Private Shared Sub client_SendCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.AsyncCompletedEventArgs)
        If e.[Error] IsNot Nothing Then
            'MessageBox.Show(e.[Error].ToString())
            Throw New Exception(e.[Error].ToString())
        End If
    End Sub

    ''' <summary>
    ''' 完整的寄信功能
    ''' </summary>
    ''' <param name="sMailTos">收信人E-mail Address</param>
    ''' <param name="sCcs">副本E-mail Address</param>
    ''' <param name="sMailSub">主旨</param>
    ''' <param name="sMailBody">信件內容</param>
    ''' <param name="isBodyHtml">是否採用HTML格式</param>
    ''' <param name="sfilePaths">附檔在WebServer檔案總管路徑</param>
    ''' <param name="deleteFileAttachment">是否刪除在WebServer上的附件</param>
    ''' <returns>是否成功</returns>
    Public Shared Function GMail_Send(ByVal sMailTos As String(), ByVal sCcs As String(), ByVal sMailSub As String, ByVal sMailBody As String, ByVal isBodyHtml As Boolean, _
      ByVal sfilePaths As String(), ByVal deleteFileAttachment As Boolean) As Boolean
        Try
            '防呆
            'If String.IsNullOrEmpty(sMailFrom) Then
            '    '※有些公司的Mail Server會規定寄信人的Domain Name要是該Mail Server的Domain Name
            '    sMailFrom = com.Azion.EloanUtility.FileUtility.getAppSettings("MAILFROM")
            'End If

            '命名空間： System.Web.Mail已過時，http://msdn.microsoft.com/zh-tw/library/system.web.mail.mailmessage(v=vs.80).aspx
            '建立MailMessage物件
            Dim mms As New System.Net.Mail.MailMessage()

            '指定一位寄信人MailAddress
            mms.From = New MailAddress(com.Azion.EloanUtility.FileUtility.getAppSettings("MAILFROM"), com.Azion.EloanUtility.FileUtility.getAppSettings("MAILNAME"), System.Text.Encoding.UTF8)

            '信件主旨
            mms.Subject = sMailSub
            mms.SubjectEncoding = System.Text.Encoding.UTF8

            '信件內容
            mms.Body = sMailBody
            mms.BodyEncoding = System.Text.Encoding.UTF8

            '信件內容 是否採用Html格式
            mms.IsBodyHtml = isBodyHtml

            If sMailTos IsNot Nothing Then
                '防呆
                Dim i As Integer = 0
                While i < sMailTos.Length
                    '加入信件的收信人(們)address
                    If Not String.IsNullOrEmpty(sMailTos(i).Trim()) Then
                        mms.[To].Add(New MailAddress(sMailTos(i).Trim()))
                    End If
                    System.Math.Max(System.Threading.Interlocked.Increment(i), i - 1)
                End While
            End If

            'End if (MailTos !=null)//防呆
            If sCcs IsNot Nothing Then
                '防呆
                Dim i As Integer = 0
                While i < sCcs.Length
                    If Not String.IsNullOrEmpty(sCcs(i).Trim()) Then
                        '加入信件的副本(們)address
                        mms.CC.Add(New MailAddress(sCcs(i).Trim()))
                    End If
                    System.Math.Max(System.Threading.Interlocked.Increment(i), i - 1)
                End While
            End If
            'End if (Ccs!=null) //防呆

            Dim file__1 As Attachment = Nothing

            '宣告在這裡，待會要釋放物件用
            If sfilePaths IsNot Nothing Then
                '防呆
                '有夾帶檔案
                Dim i As Integer = 0
                While i < sfilePaths.Length
                    If Not String.IsNullOrEmpty(sfilePaths(i).Trim()) Then
                        file__1 = New Attachment(sfilePaths(i).Trim())
                        '加入信件的夾帶檔案

                        mms.Attachments.Add(file__1)
                    End If
                    System.Math.Max(System.Threading.Interlocked.Increment(i), i - 1)
                End While
            End If
            'End if (filePaths!=null)//防呆

            '或公司、客戶的smtp_server
            Dim client As New SmtpClient()
            client.Host = com.Azion.EloanUtility.FileUtility.getAppSettings("MAILSMTP")
            'Gmail smtp port
            client.Port = CInt(com.Azion.EloanUtility.FileUtility.getAppSettings("MAILPORT"))
            client.DeliveryMethod = SmtpDeliveryMethod.Network
            '開啟ssl模式
            'client.EnableSsl = True
            client.UseDefaultCredentials = False
            '以下可以省略，因為寄信不用帳密(除非客戶特別要求)
            client.Credentials = New System.Net.NetworkCredential(com.Azion.EloanUtility.FileUtility.getAppSettings("MAILFROM"), com.Azion.EloanUtility.FileUtility.getAppSettings("MAILPASSWORD"))

            '寄出一封信
            client.Send(mms)

            '#Region "釋放物件，避免程式lock住檔案"
            If file__1 IsNot Nothing Then
                file__1.Dispose()
                file__1 = Nothing
            End If
            client.Dispose()
            client = Nothing
            '#End Region

            '#Region "要刪除附檔"
            If deleteFileAttachment AndAlso sfilePaths IsNot Nothing AndAlso sfilePaths.Length > 0 Then
                For Each filePath As String In sfilePaths
                    File.Delete(filePath.Trim())
                Next
            End If
            '#End Region

            '成功
            Return True
        Catch ex As Exception
            Throw
        End Try
    End Function
#End Region

End Class
