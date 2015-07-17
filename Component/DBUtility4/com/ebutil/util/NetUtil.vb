Imports System.IO
Imports System.Net
Imports System.Net.Sockets
Imports System.Net.Mail
'Imports Big5EncoderFullDemo

Public Class NetUtil

    Shared Sub main()
        '<!-- 資料識別碼 資料型態為文字,組成方式為1-3系統別,4-11日期,12-19時間,20-24流水號 -->
        Console.WriteLine("MS1" & Date.Now.ToString(("yyyyMMddHHmmssff")))
        Console.WriteLine(Guid.NewGuid)

        'pingSite("10.253.6.170", 16887)
        Dim sInbox As String = "<?xml version=""1.0"" encoding=""BIG5""?>" & vbNewLine & _
"<!-- 時間校正 上行電文 -->" & vbNewLine & _
"<INPUT>" & vbNewLine & _
"  <AP_HEADER>" & vbNewLine & _
"    <ABEND></ABEND>" & vbNewLine & _
"    <FROM_HOST>BELOANE</FROM_HOST>" & vbNewLine & _
"    <FROM_SERVICE></FROM_SERVICE>" & vbNewLine & _
"    <FROM_DTAQ></FROM_DTAQ>" & vbNewLine & _
"    <TO_HOST>900</TO_HOST>" & vbNewLine & _
"    <TO_SERVICE></TO_SERVICE>" & vbNewLine & _
"    <TO_DTAQ>CGCTL</TO_DTAQ>" & vbNewLine & _
"    <MSG_ID>MI1200812311800000100005</MSG_ID>" & vbNewLine & _
"    <DATE>20081231</DATE>" & vbNewLine & _
"    <TIME>18000001</TIME>" & vbNewLine & _
"    <TXCD>MI01</TXCD>" & vbNewLine & _
"    <SUBCODE>00</SUBCODE>" & vbNewLine & _
"    <PRIORITY>6</PRIORITY>" & vbNewLine & _
"    <END_CODE>Y</END_CODE>" & vbNewLine & _
"  </AP_HEADER>" & vbNewLine & _
"  <AP_DATA>" & vbNewLine & _
"  </AP_DATA>" & vbNewLine & _
"</INPUT>"

        SocketSend1("192.168.8.108", 16887, sInbox)
        'SocketSend1("10.253.6.170", 16887, sInbox)

        Dim sInbox1 As String = "3,<?xml version=""1.0"" encoding=""BIG5""?>" & vbNewLine & _
"<!-- 時間校正 上行電文 -->" & vbNewLine & _
"<INPUT>" & vbNewLine & _
"  <AP_HEADER>" & vbNewLine & _
"    <ABEND></ABEND>" & vbNewLine & _
"    <FROM_HOST>BELOANE</FROM_HOST>" & vbNewLine & _
"    <FROM_SERVICE></FROM_SERVICE>" & vbNewLine & _
"    <FROM_DTAQ></FROM_DTAQ>" & vbNewLine & _
"    <TO_HOST>900</TO_HOST>" & vbNewLine & _
"    <TO_SERVICE></TO_SERVICE>" & vbNewLine & _
"    <TO_DTAQ>CGCTL</TO_DTAQ>" & vbNewLine & _
"    <MSG_ID>MI1200812311800000100005</MSG_ID>" & vbNewLine & _
"    <DATE>20081231</DATE>" & vbNewLine & _
"    <TIME>18000001</TIME>" & vbNewLine & _
"    <TXCD>MI01</TXCD>" & vbNewLine & _
"    <SUBCODE>00</SUBCODE>" & vbNewLine & _
"    <PRIORITY>6</PRIORITY>" & vbNewLine & _
"    <END_CODE>Y</END_CODE>" & vbNewLine & _
"  </AP_HEADER>" & vbNewLine & _
"  <AP_DATA>" & vbNewLine & _
"  </AP_DATA>" & vbNewLine & _
"</INPUT>"

        'SocketSend1("10.253.6.170", 16888, sInbox1)
        'SocketSend1("192.168.8.108", 8888, "我堃2")
        'SocketSend1("192.168.8.108", 8888, "bbbb")
        Console.ReadLine()
    End Sub

#Region "FTP"
    'fileName上傳的檔案ex : c:\abc.xml , uploadUrl上傳的FTP伺服器路徑ftp://127.0.0.1,UserName使用者FTP登入帳號 , Password使用者登入密碼
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
                sUploadURL = sUploadURL.Split(":")(1).Replace("//", "") & "/"
                sUploadURL = sUploadURL.Replace("//", "/")
                sUploadURL = "ftp://" & sUploadURL & fi.Name
            Else
                sUploadURL = sUploadURL.Trim & "/"
                sUploadURL = sUploadURL.Replace("//", "/")
                sUploadURL = "ftp://" & sUploadURL & fi.Name
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

            Dim bufferLength As Integer = 2048
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

        Catch ex As Exception
            Throw ex
        Finally
            If Not IsNothing(uploadResponse) Then uploadResponse.Close()
        End Try
    End Sub

#End Region

#Region "Tools"
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
#End Region

#Region "send Mail"

    Public Shared Sub sendMail(ByVal sSubject As String, ByVal sBody As String, ByVal sToAddress As String, ByVal sSMTP As String, ByVal sFromAddress As String)
        ' Mail Message Setting 
        Dim fromEmail As String = sFromAddress
        Dim fromName As String = "Boteloan Batch"
        Dim from As New MailAddress(fromEmail, fromName, System.Text.Encoding.UTF8)

        Dim toEmail As String = sToAddress.Split(",")(0)

        Dim mail As New MailMessage(from, New MailAddress(toEmail))

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

        mail.To.Add(sToAddress)
        'SMTP Setting 
        Dim client As New SmtpClient()
        client.Host = sSMTP
        'client.Port = 25
        'client.Credentials = New NetworkCredential("tita", "xxxxxx")
        'client.EnableSsl = True

        ' Send Mail 
        Try
            client.SendAsync(mail, sSubject)
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
#End Region

#Region "Socket"

    Public Shared Function socketSend(ByVal hostNameOrAddress As String, ByVal portnum As Integer, ByVal sMsg As String, Optional iTimeOut As Integer = 0) As Boolean

        Using clientSocket As New System.Net.Sockets.TcpClient()

            clientSocket.Connect(hostNameOrAddress, portnum)
            clientSocket.SendTimeout = iTimeOut

            Using serverStream As NetworkStream = clientSocket.GetStream()

                If serverStream.CanWrite Then
                    Dim outStream As Byte() = System.Text.Encoding.Default.GetBytes(sMsg)
                    serverStream.Write(outStream, 0, outStream.Length)
                    serverStream.Flush()
                End If

                If serverStream.CanRead Then
                    Dim inStream(1024) As Byte
                    Dim sServerMsg As Text.StringBuilder = New Text.StringBuilder()
                    Dim i As Integer = 0
                    ' Incoming message may be larger than the buffer size.
                    Do
                        i = serverStream.Read(inStream, 0, inStream.Length)
                        sServerMsg.AppendFormat("{0}", System.Text.Encoding.Default.GetString(inStream, 0, i))
                    Loop While serverStream.DataAvailable

                    ' Print out the received message to the console.
                    Console.WriteLine(("You received the following message : " + sServerMsg.ToString()))
                Else
                    Console.WriteLine("Sorry.  You cannot read from this NetworkStream.")
                End If
            End Using
        End Using
    End Function

    Public Shared Function SocketSendxxx(ByVal hostNameOrAddress As String, ByVal portnum As Integer, ByVal sMsg As String) As Boolean
        Dim clientSocket As New System.Net.Sockets.TcpClient()
        clientSocket.Connect(hostNameOrAddress, portnum)

        Dim serverStream As NetworkStream = clientSocket.GetStream()
        Dim outStream As Byte() = System.Text.Encoding.Default.GetBytes(sMsg)
        serverStream.Write(outStream, 0, outStream.Length)
        serverStream.Flush()

        Dim inStream(1024) As Byte
        Dim i As Integer = serverStream.Read(inStream, 0, CInt(clientSocket.ReceiveBufferSize))

        Dim returndata As String = System.Text.Encoding.Default.GetString(inStream)
        returndata = returndata.Substring(0, i - 1)
        Console.WriteLine("Data from Server : " + returndata)
    End Function

    Public Shared Function SocketSend1(ByVal hostNameOrAddress As String, ByVal portnum As Integer, ByVal sMsg As String) As Boolean

        Try
            Using sock As System.Net.Sockets.Socket = New System.Net.Sockets.Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
                'Console.WriteLine("Testing " & hostNameOrAddress & ":" & portnum & "...")
                sock.Connect(hostNameOrAddress, portnum)

                If (sock.Connected = True) Then
                    'Dim i As Integer = sock.Send(Text.Encoding.Default.GetBytes(sMsg)) ' 發送訊息
                    Dim i As Integer = sock.Send(Text.Encoding.Default.GetBytes(sMsg), sMsg.Length, SocketFlags.None) ' 發送訊息

                    Console.WriteLine("Sent {0} bytes.", i)
                    sock.Shutdown(SocketShutdown.Send) ' 傳送完後，關閉傳送端的socket，始接收端可正常接收資料

                    Dim bytes(10240) As Byte
                    ' Get reply from the server.
                    Dim byteCount As Integer = sock.Receive(bytes, sock.Available, SocketFlags.None) ' 接收回傳訊息
                    sock.Shutdown(SocketShutdown.Receive)
                    If byteCount > 0 Then
                        'Console.WriteLine(Text.Encoding.Default.GetString(bytes))
                        Dim returndata As String = System.Text.Encoding.Default.GetString(bytes)
                        returndata = returndata.Substring(0, byteCount)
                        Console.WriteLine("Data from Server : " + returndata)
                    End If
                End If

                'sock.Close()
            End Using
        Catch ex As SocketException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

#End Region

#Region "test"
    Private Declare Function DoFileDownload Lib "shdocvw.dll" (ByVal lpszFile As String) As Long

    Private Sub Command1_Click(ByVal sURL As String)
        Dim sFileUrl As String
        sFileUrl = StrConv("http://www.hosp.ncku.edu.tw/~cww/index.html", VbStrConv.None)
        DoFileDownload(sFileUrl)
    End Sub

    '第一個參數是僅當調用者是一個activex物件才使用,一般為null.
    '第二個參數就是要下載檔案的目標url,完整路徑. 
    '第三個是本地保存路徑,也是完整路徑
    '第四個是保留,必須為0
    '第五個是指向一個ibindstatuscallback介面的指標,這就類似一種回檔機制,你可以參考這些來活動當前下載進度,選擇是否繼續下載等等.
    Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
      (ByVal pCaller As Long, _
      ByVal szURL As String, _
      ByVal szFileName As String, _
      ByVal dwReserved As Long, _
      ByVal lpfnCB As Long) As Long
    Private Sub Command1_Click()
        Dim lReturn As Long

        lReturn = URLDownloadToFile(0, "http://www.hosp.ncku.edu.tw/~cww/index.htm", "C:\index.html", 0, 0)

        If lReturn = 0 Then
            MsgBox("Download Complete.", vbInformation + vbOKOnly)
        End If
    End Sub
#End Region

End Class

'-----------------------------------------------------------------------------------------------------------------





