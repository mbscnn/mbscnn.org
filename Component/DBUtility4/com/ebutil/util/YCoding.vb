Public Class YCoding
    Public Const GB2312 As Integer = 0
    Public Const GBK As Integer = 1
    Public Const HZ As Integer = 2
    Public Const BIG5 As Integer = 3
    Public Const CNS11643 As Integer = 4
    Public Const UTF8 As Integer = 5
    Public Const [UNICODE] As Integer = 6
    Public Const UNICODET As Integer = 7
    Public Const UNICODES As Integer = 8
    Public Const UTF32 As Integer = 9
    Public Const ISO2022CN As Integer = 10
    Public Const ISO2022CN_CNS As Integer = 11
    Public Const ISO2022CN_GB As Integer = 12
    Public Const ASCII As Integer = 13
    Public Const OTHER As Integer = 14


    Private Const TOTALTYPES As Integer = 15

    Public Shared MAXUNICODE As Integer = &HFFFF
    ' Names of the encodings as understood by Java
    Private Shared javaname(TOTALTYPES) As String
    ' Names of the encodings for human viewing
    Private Shared nicename(TOTALTYPES) As String
    ' Names of charsets as used in charset parameter of HTML Meta tag
    Private Shared htmlname(TOTALTYPES) As String

    ' Simplfied/Traditional character equivalence hashes
    Protected s2thash, t2shash As System.Collections.Hashtable

    ' lookup table for special characters
    Protected specialChars(128) As Boolean   ' = New Boolean(128)
    Protected e2i As System.Collections.Hashtable = New System.Collections.Hashtable
    Protected i2e As System.Collections.Hashtable = New System.Collections.Hashtable
    Private buffer(1000) As Byte '= New Byte(1000)

    Private m_sFilePath As String = ".\hcutf8.txt"

    Private Declare Function MultiByteToWideChar Lib "kernel32" ( _
    ByVal CodePage As Long, ByVal dwFlags As Long, _
    ByVal lpMultiByteStr As Byte(), ByVal cchMultiByte As Long, _
    ByVal lpWideCharStr As Byte(), ByVal cchWideChar As Long) As Long


    Private Declare Function WideCharToMultiByte Lib "kernel32" ( _
  ByVal CodePage As Long, ByVal dwFlags As Long, _
  ByVal lpWideCharStr As Byte(), ByVal cchWideChar As Long, _
  ByVal lpMultiByteStr As Byte(), ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long) As Long


    'static void Main(string[] args)
    '{
    'string str = "ÆèÆêÆìÆîÆð´ú¸Õ";

    'int ret;
    'byte[] buf = new byte[str.Length * 2];
    'ret = WideCharToMultiByte(950, 0, Encoding.Unicode.GetBytes(str), -1, buf, buf.Length, null, 0);

    'File.WriteAllBytes(@"E:\Test.txt", buf);

    'byte[] readbuf = File.ReadAllBytes(@"E:\Test.txt");
    'byte[] strbuf = new byte[readbuf.Length];
    'ret = MultiByteToWideChar(950, 0, readbuf, -1, strbuf, str.Length);

    'MessageBox.Show(Encoding.Unicode.GetString(strbuf));
    '}
    '} 

    Shared Sub main()



        'yCoding.transferHtmlToVB(ENDescribe.Describe)
        Dim coder As New YCoding
        Dim tc As String = "???"
        Dim s As String = "&#30908;&#999999;A&B;"
        Try
            Dim sc As String = coder.convertString(tc, YCoding.UNICODE, YCoding.HZ)
            Console.Write(sc)
            s = coder.transferNCRToUnicode(s)
            Console.Write(s)
        Catch ex As Exception
            Console.Write(ex)
        End Try

    End Sub

    Public Sub init()

        javaname(GB2312) = "GB2312"
        javaname(HZ) = "ASCII"
        javaname(GBK) = "GBK"
        javaname(ISO2022CN_GB) = "ISO2022CN_GB"
        javaname(BIG5) = "BIG5"
        javaname(CNS11643) = "EUC-TW"
        javaname(ISO2022CN_CNS) = "ISO2022CN_CNS"
        javaname(ISO2022CN) = "ISO2022CN"
        javaname(UTF8) = "UTF8"
        javaname([UNICODE]) = "Unicode"
        javaname(UNICODET) = "Unicode"
        javaname(UNICODES) = "Unicode"
        javaname(UTF32) = "UTF-32"
        javaname(ASCII) = "ASCII"
        javaname(OTHER) = "ISO8859_1"

        htmlname(GB2312) = "GB2312"
        htmlname(HZ) = "HZ-GB-2312"
        htmlname(GBK) = "GB2312"
        htmlname(ISO2022CN_GB) = "ISO-2022-CN-EXT"
        htmlname(BIG5) = "BIG5"
        htmlname(CNS11643) = "EUC-TW"
        htmlname(ISO2022CN_CNS) = "ISO-2022-CN-EXT"
        htmlname(ISO2022CN) = "ISO-2022-CN"
        htmlname(UTF8) = "UTF-8"
        htmlname([UNICODE]) = "UTF-16"
        htmlname(UNICODET) = "UTF-16"
        htmlname(UNICODES) = "UTF-16"
        htmlname(UTF32) = "UTF-32"
        htmlname(ASCII) = "ASCII"
        htmlname(OTHER) = "ISO8859-1"

        nicename(GB2312) = "GB-2312"
        nicename(HZ) = "HZ"
        nicename(GBK) = "GBK"
        nicename(ISO2022CN_GB) = "ISO2022CN-GB"
        nicename(BIG5) = "Big5"
        nicename(CNS11643) = "CNS11643"
        nicename(ISO2022CN_CNS) = "ISO2022CN-CNS"
        nicename(ISO2022CN) = "ISO2022 CN"
        nicename(UTF8) = "UTF-8"
        nicename([UNICODE]) = "Unicode"
        nicename(UNICODET) = "Unicode (Trad)"
        nicename(UNICODES) = "Unicode (Simp)"
        nicename(UTF32) = "UTF-32"
        nicename(ASCII) = "ASCII"
        nicename(OTHER) = "OTHER"
    End Sub


    Sub New()
        init()
        initialSpecialChar()
      
        s2thash = New System.Collections.Hashtable
        t2shash = New System.Collections.Hashtable

        Dim sr As System.IO.StreamReader = Nothing
        Dim str As String
        Dim sPath As String = ""

        Try
            'Dim br As New System.IO.BinaryReader(New System.IO.FileStream(m_sFilePath, System.IO.FileMode.Open, System.IO.FileAccess.Read))
            If Not System.Web.HttpContext.Current Is Nothing Then
                sPath = System.Web.HttpContext.Current.Request.MapPath((com.Azion.EloanUtility.UIUtility.getRootPath() & "/bin"))
            End If

            sr = System.IO.File.OpenText(sPath & m_sFilePath)

            Do
                str = sr.ReadLine()

                If Not IsNothing(str) AndAlso str.Trim.Length > 0 AndAlso Not str.Chars(0) = "#" Then
                    If Not s2thash.ContainsKey([String].Intern(str.Substring(0, 1))) Then
                        s2thash.Add([String].Intern(str.Substring(0, 1)), str.Substring(1, 1))
                    End If
                    For i As Integer = 1 To str.Length - 1
                        If Not t2shash.ContainsKey([String].Intern(str.Substring(i, 1))) Then
                            t2shash.Add([String].Intern(str.Substring(i, 1)), str.Substring(0, 1))
                        End If
                    Next
                End If
            Loop Until str Is Nothing

        Catch ex As Exception
            Throw ex
        Finally
            If Not IsNothing(sr) Then sr.Close()
        End Try
    End Sub


    Public Function convertString(ByVal dataline As String, ByVal source_encoding As Integer, ByVal target_encoding As Integer) As String

        Dim outline As System.Text.StringBuilder = New System.Text.StringBuilder
        Dim lineindex As Integer = 0

        If source_encoding = HZ Then
            dataline = hz2gb(dataline)
        End If

        While lineindex < dataline.Length()

            If ((source_encoding = GB2312 Or source_encoding = UNICODES Or source_encoding = ISO2022CN_GB Or source_encoding = GBK Or source_encoding = [UNICODE] Or source_encoding = HZ) And (target_encoding = BIG5 Or target_encoding = CNS11643 Or target_encoding = UNICODET Or target_encoding = ISO2022CN_CNS)) Then

                If (s2thash.ContainsKey(dataline.Substring(lineindex, 1)) = True) Then
                    Dim item As Object = s2thash.Item(dataline.Substring(lineindex, 1))
                    outline.Append(String.Intern(item))
                Else
                    outline.Append(dataline.Substring(lineindex, 1))
                End If
            ElseIf ((source_encoding = BIG5 Or source_encoding = CNS11643 Or source_encoding = UNICODET Or source_encoding = ISO2022CN_CNS Or source_encoding = GBK Or source_encoding = [UNICODE]) And (target_encoding = GB2312 Or target_encoding = UNICODES Or target_encoding = ISO2022CN_GB Or target_encoding = HZ)) Then

                If (t2shash.ContainsKey(dataline.Substring(lineindex, 1)) = True) Then
                    Dim item As Object = t2shash.Item(dataline.Substring(lineindex, 1))
                    outline.Append(String.Intern(item))
                Else
                    outline.Append(dataline.Substring(lineindex, 1))
                End If
            End If

            lineindex += 1
        End While

        If (target_encoding = HZ) Then
            ' Convert to look like HZ
            Return gb2hz(outline.ToString())
        End If

        Return outline.ToString()
    End Function
    'Numeric Character Reference(NCR)(html Unicode)  Âà  Unicode 
    Public Function transferNCRToUnicode(ByVal source As String) As String
        Dim resultBuffer As System.Text.StringBuilder = New System.Text.StringBuilder
        Dim entity As String = String.Empty
        Dim i As Integer = 0
        Dim semi As Integer = 0

[continue]: While (i < source.Length())
            Dim ch As Char = source.Chars(i)
            'UP LOGIC > After "&" must have "#" Then Run Condition
            If (ch = "&") AndAlso (i + 1) < source.Length() AndAlso source.Chars(i + 1) = "#" Then
                semi = source.IndexOf(";", i + 1)
                If (semi = -1) Then
                    resultBuffer.Append(ch)
                    i += 1
                    GoTo [continue]
                End If
                '&#30908;
                entity = source.Substring(i + 1, semi - (i + 1)) ' get &#nnnn;
                Dim iso As Integer = 0
                If (entity.Chars(0) = "#") Then
                    Try
                        iso = entity.Substring(1)

                    Catch ex As Exception

                    End Try
                Else
                    iso = e2i.Item(entity)
                End If

                If (iso = 0 Or iso.CompareTo(MAXUNICODE) > 0) Then
                    resultBuffer.Append(ch)
                    i += 1
                    'i = semi
                    GoTo [continue]
                Else
                    i = semi
                    Dim s As String = ChrW(entity.Substring(1)).ToString
                    'Dim a As String = Hex(entity.Substring(1))
                    'Dim abyte0() As Byte = digits2Bytes(a, 16)
                    'Dim s As String = byteToString(abyte0, javaname([UNICODE]))

                    'System.Text.Encoding.GetEncoding(javaname([UNICODE])).GetChars(System.Text.Encoding.GetEncoding(javaname([UNICODE])).GetBytes(s))
                    Try
                        resultBuffer.Append(System.Text.Encoding.GetEncoding(javaname([UNICODE])).GetChars(System.Text.Encoding.GetEncoding(javaname([UNICODE])).GetBytes(s)))
                    Catch ex As Exception
                        Throw ex
                    End Try
                End If
            Else
                resultBuffer.Append(ch)
            End If
            i += 1

        End While
        Return resultBuffer.ToString()
    End Function

    Public Function digits2Bytes(ByVal s As String, ByVal i As Integer) As Byte()

        If (i = 0) Then
            Dim abyte0(s.Length()) As Byte

            For k As Integer = 0 To k < s.Length()
                abyte0(k) = CType(Char.GetNumericValue(s.Chars(k)), Byte)
            Next

            Return abyte0
        End If

        Dim j As Integer = 0
        Dim l As Integer = digitsPerByte(i)
        Dim i1 As Integer = 0
        Dim j1 As Integer = 0

        For k1 As Integer = 0 To k1 < s.Length()
            Dim c As Char = s.Chars(k1)
            Dim l1 As Integer = Char.GetNumericValue(c)

            If (l1 < 0 Or l1 >= i) Then
                If Not (j1 = 0) Then
                    j += 1
                    buffer(j) = CType(i1, Byte)
                    i1 = 0
                    j1 = 0
                Else
                    i1 = i1 * i + l1

                    If (++j1 >= l) Then
                        j += 1
                        buffer(j) = CType(i1, Byte)
                        i1 = 0
                        j1 = 0
                    End If
                End If
            End If
        Next

        If (j1 > 0) Then
            j += 1
            buffer(j) = CType(i1, Byte)
        End If

        Dim abyte1(j) As Byte

        Array.Copy(buffer, 0, abyte1, 0, j)
        Return abyte1
    End Function

    Public Function digitsPerByte(ByVal i As Integer) As Integer
        If (i = 0) Then
            Return 1
        End If
        Dim j As Integer = 1

        Dim k As Integer = i
        While k < 256
            j += 1
            k *= i
        End While

        Return j
    End Function

    Public Function byteToString(ByVal abyte0() As Byte, ByVal coding As String) As String
        Try
            If (Not coding.Equals(javaname([UNICODE])) And Not coding.Equals(javaname(UTF32))) Then
                Return System.Text.Encoding.GetEncoding(coding).GetString(abyte0)
            End If
        Catch ex As Exception
            Return ""
        End Try

        Dim stringbuffer As System.Text.StringBuilder = New System.Text.StringBuilder
        If (coding.Equals(javaname(UTF32))) Then
            Dim i As Integer = 0

            While i < abyte0.Length

                Dim k As Integer = (&HFF And abyte0(i))
                Dim i1 As Integer = IIf(i + 1 < abyte0.Length(), (&HFF And abyte0(i + 1)), 0)
                Dim k1 As Integer = IIf(i + 2 < abyte0.Length(), (&HFF And abyte0(i + 2)), 0)
                Dim l1 As Integer = IIf(i + 3 < abyte0.Length(), (&HFF And abyte0(i + 3)), 0)

                Dim l2 As Long = k << 24 Or i1 << 16 Or k1 << 8 Or l1

                If (l2 <= 65535L) Then
                    stringbuffer.Append(ChrW(l2))
                Else
                    l2 -= &H10000L

                    stringbuffer.Append(ChrW((l2 >> 10 And 1023L) + 55296L))
                    stringbuffer.Append(ChrW((l2 And 1023L) + 56320L))
                End If
                i += 4
            End While
        Else
            Dim j As Integer = 0

            While j < abyte0.Length
                Dim l As Integer = &HFF And abyte0(j)

                Dim j1 As Integer = IIf(j + 1 < abyte0.Length(), (&HFF And abyte0(j + 1)), 0)
                stringbuffer.Append(ChrW(l << 8 Or j1))
                j += 2
            End While
        End If

        Return stringbuffer.ToString()
    End Function

    Public Function hz2gb(ByVal hzstring As String) As String
        Dim hzbytes(2) As Byte
        Dim gbchar(2) As Byte
        Dim gbstring As System.Text.StringBuilder = New System.Text.StringBuilder

        Try
            'hzbytes = System.Text.Encoding.GetEncoding("iso-8859-1").GetBytes(hzstring.ToCharArray)
            hzbytes = System.Text.Encoding.GetEncoding("iso-8859-1").GetBytes(hzstring)
        Catch ex As Exception
            'Throw ex
            Return hzstring
        End Try

        ' Convert to look like equivalent Unicode of GB
        Dim byteindex As Integer = 0
        While byteindex < hzbytes.Length()
            If (hzbytes(byteindex) = &H7E) Then
                If (hzbytes(byteindex + 1) = &H7B) Then
                    byteindex += 2
                    While (byteindex < hzbytes.Length)
                        If (hzbytes(byteindex) = &H7E And hzbytes(byteindex + 1) = &H7D) Then
                            byteindex += 1
                            Exit While
                        ElseIf (hzbytes(byteindex) = &HA Or hzbytes(byteindex) = &HD) Then
                            gbstring.Append(hzbytes(byteindex))
                            Exit While
                        End If
                        gbchar(0) = hzbytes(byteindex) + (&H80)
                        gbchar(1) = hzbytes(byteindex + 1) + (&H80)

                        'Dim sr As System.IO.BinaryReader
                        Try
                            Dim str As New String(gbchar(0) & gbchar(1))
                            'sr = New System.IO.BinaryReader(str.Chars, System.Text.Encoding.GetEncoding("gb2312"))
                            gbstring.Append(gbchar)
                        Catch ex As Exception
                            Throw ex
                        Finally
                            'sr.Close()
                        End Try
                        byteindex += 2
                    End While
                ElseIf (hzbytes(byteindex + 1) = &H7E) Then
                    ' ~~ becomes ~
                    gbstring.Append("~")
                Else 'false alarm
                    gbstring.Append(hzbytes(byteindex))
                End If
            Else
                gbstring.Append(hzbytes(byteindex))
            End If

            byteindex += 1
        End While


        Return gbstring.ToString()
    End Function

    Public Function gb2hz(ByVal gbstring As String) As String

        Dim hzbuffer As System.Text.StringBuilder = New System.Text.StringBuilder
        Dim gbbytes(2) As Byte
        Dim i As Integer = 0
        Dim terminated As Boolean = False

        Try
            gbbytes = System.Text.Encoding.GetEncoding("GB2312").GetBytes(gbstring)
        Catch ex As Exception
            Return gbstring
        End Try

        While (i < gbbytes.Length())

            If (gbbytes(i) < 0) Then
                hzbuffer.Append("~{")
                terminated = False

                While (i < gbbytes.Length)

                    If (gbbytes(i) = &HA Or gbbytes(i) = &HD) Then
                        hzbuffer.Append("~}" + gbbytes(i))
                        terminated = True
                        Exit While
                    ElseIf (gbbytes(i) >= 0) Then
                        hzbuffer.Append("~}" + gbbytes(i))
                        terminated = True
                        Exit While
                    End If

                    hzbuffer.Append((gbbytes(i) + 256 - &H80))
                    hzbuffer.Append((gbbytes(i + 1) + 256 - &H80))
                    i += 2
                End While

                If (terminated = False) Then
                    hzbuffer.Append("~}")
                End If
            Else
                If (gbbytes(i) = &H7E) Then
                    hzbuffer.Append("~~")
                Else
                    hzbuffer.Append(gbbytes(i))
                End If
            End If

            i += 1
        End While


        Return hzbuffer.ToString
    End Function

    '** ************** HTML unicode <--> JAVA unicode (Start) ****************** */
    ' transferJAVAtoHTML
    'protected boolean[] specialChars = new boolean[128]
    Protected Sub initialSpecialChar()
        For i As Integer = 0 To 127
            specialChars(i) = False
        Next

        Try

            Dim ch As Char = "<"

            specialChars(AscW("<")) = True
            specialChars(AscW(">")) = True
            specialChars(AscW("&")) = True
            specialChars(AscW("\")) = True
            specialChars(AscW("\""")) = True
        Catch ex As Exception
            Throw ex
        End Try


        Dim entities(,) As Object = {{"#39", 39}, {"quot", 34}, {"amp", 38}, {"lt", 60}, {"gt", 62}, {"nbsp", 160}, {"copy", 169}, {"reg", 174}, {"Agrave", 192}, {"Aacute", 193}, {"Acirc", 194}, {"Atilde", 195}, {"Auml", 196}, {"Aring", 197}, {"AElig", 198}, {"Ccedil", 199}, {"Egrave", 200}, {"Eacute", 201}, {"Ecirc", 202}, {"Euml", 203}, {"Igrave", 204}, {"Iacute", 205}, {"Icirc", 206}, {"Iuml", 207}, {"ETH", 208}, {"Ntilde", 209}, {"Ograve", 210}, {"Oacute", 211}, {"Ocirc", 212}, {"Otilde", 213}, {"Ouml", 214}, {"Oslash", 216}, {"Ugrave", 217}, {"Uacute", 218}, {"Ucirc", 219}, {"Uuml", 220}, {"Yacute", 221}, {"THORN", 222}, {"szlig", 223}, {"agrave", 224}, {"aacute", 225}, {"acirc", 226}, {"atilde", 227}, {"auml", 228}, {"aring", 229}, {"aelig", 230}, {"ccedil", 231}, {"egrave", 232}, {"eacute", 233}, {"ecirc", 234}, {"euml", 235}, {"igrave", 236}, {"iacute", 237}, {"icirc", 238}, {"iuml", 239}, {"eth", 240}, {"ntilde", 241}, {"ograve", 242}, {"oacute", 243}, {"ocirc", 244}, {"otilde", 245}, {"ouml", 246}, {"oslash", 248}, {"ugrave", 249}, {"uacute", 250}, {"ucirc", 251}, {"uuml", 252}, {"yacute", 253}, {"thorn", 254}, {"yuml", 255}, {"euro", 8364}}

        For i As Integer = 0 To entities.GetLength(0) - 1
            e2i.Add(entities(i, 0), entities(i, 1))
            i2e.Add(entities(i, 1), entities(i, 0))
        Next
    End Sub

    Public Shared Function ToHexString(ByVal bytes() As Byte) As String

        Dim hexStr As String = ""
        Dim i As Integer
        For i = 0 To bytes.Length - 1
            hexStr = hexStr + Hex(bytes(i))
        Next i
        Return hexStr

    End Function 'ToHexString


End Class
