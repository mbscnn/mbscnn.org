Option Explicit On
Option Strict On

Imports System.Text.RegularExpressions

''' <summary>
''' 提供資料驗證相關函式
''' [Titan] 	2011/07/19	Created
''' </summary>
Public NotInheritable Class ValidateUtility
    ''' <summary>
    ''' 驗證有無資料
    ''' </summary>
    ''' <param name="obj"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' [Titan] 	2011/07/19	Created
    ''' </remarks>  
    Public Shared Function isValidateData(ByVal obj As Object) As Boolean
        If IsDBNull(obj) Then
            Return False
        ElseIf IsNothing(obj) Then
            Return False
        ElseIf TypeOf obj Is String Then
            Dim str As String = CType(obj, String)
            If (str <> Nothing) Then
                Return True
            End If
            Return False
        ElseIf TypeOf obj Is Date Then
            Return IsDate(obj)
        ElseIf TypeOf obj Is Boolean Then
            Return CBool(obj)
        End If

        Return True
    End Function


    '更新Request或把String轉換成NameValueCollection
    Public Shared Function clearReqQry(ByVal _OldRqst As Object, Optional ByVal _NewRqst As String = "") As System.Collections.Specialized.NameValueCollection
        Dim NVC As New System.Collections.Specialized.NameValueCollection

        '讀取舊Request
        If TypeOf _OldRqst Is String AndAlso ValidateUtility.isValidateData(CType(_OldRqst, String)) Then
            Dim sOldRqst As String = CType(_OldRqst, String)
            '去掉www.xxx.com.tw/xxx.aspx?
            If sOldRqst.Split(CChar("?")).Length > 1 Then sOldRqst = sOldRqst.Split(CChar("?"))(1)
            For Each s As String In sOldRqst.Split(CChar("&"))
                If ValidateUtility.isValidateData(s) Then
                    Dim sKey As String = s.Split(CChar("="))(0)
                    Dim sValue As String = s.Split(CChar("="))(1)
                    If ValidateUtility.isValidateData(sValue) Then
                        Dim str() As String = sValue.Split(CType(",", Char))
                        sValue = str(str.Length - 1)
                        NVC.Item(sKey) = sValue '原本有什就留住，不過濾，取最後一個值
                    Else
                        NVC.Item(sKey) = "" '原本有什就留住，不過濾
                    End If
                End If
            Next
        ElseIf TypeOf _OldRqst Is System.Collections.Specialized.NameValueCollection Then
            NVC.Add(CType(_OldRqst, System.Collections.Specialized.NameValueCollection))
        End If

        If Not ValidateUtility.isValidateData(_NewRqst) Then Return NVC

        '更新 Request
        For Each s As String In _NewRqst.Split(CChar("&"))
            If ValidateUtility.isValidateData(s) Then
                Dim sKey As String = s.Split(CChar("="))(0)
                Dim sValue As String = s.Split(CChar("="))(1)
                If ValidateUtility.isValidateData(sValue) Then
                    Dim str() As String = sValue.Split(CType(",", Char))
                    sValue = str(str.Length - 1)
                    NVC.Item(sKey) = sValue '原本有什就留住，不過濾，取最後一個值
                Else
                    NVC.Item(sKey) = ""
                End If
            End If
        Next

        Return NVC
    End Function

    Public Shared Function NVCtoString(ByRef NVC As Collections.Specialized.NameValueCollection) As String
        Dim sb As New System.Text.StringBuilder
        Dim bAdded As Boolean = False '判斷是否已經加入第一個key
        Dim sDelRequestKey As String = "Del_Request"
        Dim sDelRequest As String

        'DEL_REQUEST:刪除Request某個項目
        'ex: 刪除Request("Display")和Request("MDisplayMode") ,Del_Request=FinalPage;MDisplayMode
        Try
            If ValidateUtility.isValidateData(NVC.Item(sDelRequestKey)) Then
                sDelRequest = NVC.Item(sDelRequestKey)
                For Each sKey As String In sDelRequest.Split(CChar(";"))
                    NVC.Remove(sKey)
                Next
                NVC.Remove(sDelRequestKey)
            End If
        Catch ex As System.NotSupportedException
            'NVC 如果是唯讀就離開
        End Try

        For i As Integer = 0 To NVC.Count - 1
            Dim sKey As String = NVC.GetKey(i)
            Dim sValue As String = NVC.Item(i)
            sValue = System.Web.HttpUtility.UrlEncode(sValue, System.Web.HttpContext.Current.Response.ContentEncoding)

            If ValidateUtility.isValidateData(sKey) Then  'Key存在才加入
                If bAdded Then sb.Append("&") '如果已經有key就用"&"分格
                If ValidateUtility.isValidateData(sValue) AndAlso sValue.IndexOf(",") > -1 Then
                    Dim aryValue() As String = sValue.Split(CChar(","))
                    sb.Append(sKey & "=" & aryValue(UBound(aryValue)))
                Else
                    sb.Append(sKey & "=" & sValue)
                End If
                bAdded = True
            End If
        Next

        Return sb.ToString
    End Function

#Region "証號辨識"
    Public Shared Function isCOMPUser(ByVal sCUSTOMERID As String) As Boolean
        Try
            If Not (UCase(com.Azion.EloanUtility.ValidateUtility.ChPersonalID1(sCUSTOMERID)) = "TRUE" OrElse _
                    UCase(com.Azion.EloanUtility.ValidateUtility.ChIDNTax(sCUSTOMERID)) = "TRUE" OrElse _
                    UCase(com.Azion.EloanUtility.ValidateUtility.ChresidenceID(sCUSTOMERID)) = "TRUE") Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    '[身份證字號驗證]
    Public Shared Function ChPersonalID1(ByVal sEId As String) As String
        Dim Letter As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        Dim MirrorNumber() = {10, 11, 12, 13, 14, 15, 16, 17, 34, 18, 19, 20, 21, 22, 35, 23, 24, 25, 26, 27, 28, 29, 32, 30, 31, 33}
        Dim Checker() = {1, 9, 8, 7, 6, 5, 4, 3, 2, 1, 1}
        Dim N1 As String = UCase(Left(sEId, 1))
        Dim I, N1Ten, N1Unit, Total As Integer
        Dim Pos As Byte = CByte(InStr(Letter, N1))
        If Pos <= 0 Then

            Return "身份證字號第 1 碼必須為英文字母"
            Exit Function
        End If
        If Len(sEId) < 10 Then

            Return "身份證字號共有 10 碼"
            Exit Function
        End If
        For I = 2 To 10
            If Not (Mid(sEId, I, 1) >= "0" And Mid(sEId, I, 1) <= "9") Then

                Return "身份證字號第 2 ~ 9 碼必須為數字"
                Exit Function
            End If
        Next
        If Not (Mid(sEId, 2, 1) = "1" Or Mid(sEId, 2, 1) = "2") Then

            Return "身份證字號第 2 碼必須為 1 或 2"
            Exit Function
        End If
        N1Ten = CInt(Mid(CStr(MirrorNumber(Pos - 1)), 1, 1))
        N1Unit = CInt(Mid(CStr(MirrorNumber(Pos - 1)), 2, 1))
        Total = N1Ten * Checker(0) + N1Unit * Checker(1)
        For I = 2 To 10
            Total = Total + CInt(Mid(sEId, I, 1)) * Checker(I)
        Next
        If Total Mod 10 = 0 Then
            Return "true"
        Else

            Return "身份證字號錯誤！"
            Exit Function
        End If
    End Function

    '[營利事業統編驗證]
    Public Shared Function ChCompanyID1(ByVal sEId As String) As String
        Dim Checker() = {1, 2, 1, 2, 1, 2, 4, 1}
        Dim I, N1, Total As Integer

        If Len(sEId) <> 8 Then

            Return "營利事業統編共有 8 碼"
        End If

        For I = 1 To 8
            If Not (Mid(sEId, I, 1) >= "0" And Mid(sEId, I, 1) <= "9") Then
                Return "營利事業統編必須為數字"
            End If
        Next

        Total = 0
        For I = 1 To 8                                      '正常情況
            N1 = CInt(Mid(sEId, I, 1)) * Checker(I - 1)
            Total += N1 Mod 10 + N1 \ 10
        Next

        If Total Mod 10 <> 0 Then
            '(joyce2005_4_26)---------------------------------------------------------------
            If Mid(sEId, 7, 1) = "7" Then                       '例外情況(第7位為7時)
                Dim sum_D7, a4, b4, sum2_D7, a5, b5 As Integer

                sum_D7 = CInt(CInt(Mid(sEId, 7, 1)) * 4)
                a4 = CInt(Mid(CStr(sum_D7), 1, 1))
                b4 = CInt(Mid(CStr(sum_D7), 2, 1))
                sum2_D7 = a4 + b4
                a5 = CInt(Mid(CStr(sum2_D7), 1, 1))
                b5 = CInt(Mid(CStr(sum2_D7), 2, 1))
                Total = Total - a4 - b4 + a5
            End If
            '/------------------------------------------------------------------------------
        End If

        If Total Mod 10 = 0 Then
            Return "true"
        Else
            Return "營利事業統編錯誤！"
        End If
    End Function

    Public Shared Function CHSpecialID(ByVal sIDBAN As String) As String
        Try
            sIDBAN = UCase(Trim(sIDBAN))
            If sIDBAN.Length >= 3 AndAlso Left(sIDBAN, 3) = "OBU" Then
                Return "true"
            ElseIf sIDBAN.Length >= 2 AndAlso _
                (sIDBAN.Substring(1, 1) = "A" OrElse sIDBAN.Substring(1, 1) = "B" OrElse sIDBAN.Substring(1, 1) = "C" OrElse sIDBAN.Substring(1, 1) = "D") Then
                '檢核居留證號碼
                Return ChresidenceID(sIDBAN)
            Else
                '外國人稅籍編號
                '外國人稅籍編號最後2碼應為英文字母
                '外國人稅籍編號前8碼應為數字
                If sIDBAN.Length = 10 Then
                    Dim sCAPITALS As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
                    If InStr(sCAPITALS, Right(sIDBAN, 1)) > 0 AndAlso InStr(sCAPITALS, sIDBAN.Substring(8, 1)) > 0 AndAlso IsNumeric(Left(sIDBAN, 8)) Then
                        Return ChIDNTax(sIDBAN)
                    End If
                End If
            End If

            Return "false"
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    '[外國人稅籍編號]
    Public Shared Function ChIDNTax(ByVal sIDBAN As String) As String
        Try
            sIDBAN = Trim(sIDBAN)

            If sIDBAN.Length <> 10 Then
                Return "外國人稅籍編號應為10碼, 但輸入值不為10碼!"
            End If

            Dim AR_IDBAN(9) As String
            For i As Integer = 0 To sIDBAN.Length - 1
                AR_IDBAN(i) = sIDBAN.Substring(i, 1)
            Next

            Dim sCAPITALS As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
            If InStr(sCAPITALS, AR_IDBAN(8)) = 0 OrElse InStr(sCAPITALS, AR_IDBAN(9)) = 0 Then
                Return "外國人稅籍編號最後2碼應為英文字母!"
            End If

            For i As Integer = 0 To 7
                If Not IsNumeric(AR_IDBAN(i)) Then
                    Return "外國人稅籍編號前8碼應為數字, 但輸入值的第" & i + 1 & "碼不為數字!"
                End If
            Next

            Dim sYear As String = sIDBAN.Substring(0, 4)
            Dim sMonth As String = sIDBAN.Substring(4, 2)
            Dim sDay As String = sIDBAN.Substring(6, 2)

            If Not IsDate(sYear & "/" & sMonth & "/" & sDay) Then
                Return "外國人稅籍編號中出生日不合理!"
            End If

            Return "true"
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    '[居留證號碼]
    Public Shared Function ChresidenceID(ByVal sIDBAN As String) As String
        Try
            sIDBAN = Trim(sIDBAN)

            If sIDBAN.Length <> 10 Then
                Return "居留證號碼錯誤!"
            End If

            If Not IsNumeric(Right(sIDBAN, 8)) Then
                Return "居留證號碼錯誤!"
            End If

            Dim sCAPITALS As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
            If InStr(sCAPITALS, Left(sIDBAN, 1)) = 0 Then
                Return "居留證號碼錯誤!"
            End If

            If InStr(sCAPITALS, sIDBAN.Substring(1, 1)) = 0 Then
                Return "居留證號碼錯誤!"
            End If

            Dim shead As String = "ABCDEFGHJKLMNPQRSTUVXYWZIO"

            Dim sFSTTemp As String = (InStr(shead, Left(sIDBAN, 1)) - 1 + 10).ToString
            Dim sSECTemp As String = ((InStr(shead, sIDBAN.Substring(1, 1)) - 1 + 10) Mod 10).ToString
            Dim sTHDTemp As String = Right(sIDBAN, 8)

            Dim sCHKIDBAN As String = sFSTTemp & sSECTemp & sTHDTemp
            If sCHKIDBAN.Length = 11 Then
                Dim iCHKNUM As Decimal = 0

                iCHKNUM = CInt(sCHKIDBAN.Substring(0, 1))
                iCHKNUM += (CInt(sCHKIDBAN.Substring(1, 1)) * 9)
                iCHKNUM += (CInt(sCHKIDBAN.Substring(2, 1)) * 8)
                iCHKNUM += (CInt(sCHKIDBAN.Substring(3, 1)) * 7)
                iCHKNUM += (CInt(sCHKIDBAN.Substring(4, 1)) * 6)
                iCHKNUM += (CInt(sCHKIDBAN.Substring(5, 1)) * 5)
                iCHKNUM += (CInt(sCHKIDBAN.Substring(6, 1)) * 4)
                iCHKNUM += (CInt(sCHKIDBAN.Substring(7, 1)) * 3)
                iCHKNUM += (CInt(sCHKIDBAN.Substring(8, 1)) * 2)
                iCHKNUM += CInt(sCHKIDBAN.Substring(9, 1))
                iCHKNUM += CInt(sCHKIDBAN.Substring(10, 1))

                '判斷是否可整除
                If (iCHKNUM Mod 10) = 0 Then
                    Return "true"
                Else
                    Return "居留證號碼錯誤!"
                End If
            Else
                Return "居留證號碼錯誤!"
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

#Region "e-Mail辨識"
    Public Shared Function isEmail(ByVal sInputEmail As String) As Boolean
        'sInputEmail = NulltoString(sInputEmail)
        Dim strRegex As String = "^([a-zA-Z0-9_\-\.]+)@((\[[0-9]{1,3}" + "\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([a-zA-Z0-9\-]+\" + ".)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$"
        Dim re As New Regex(strRegex)
        If re.IsMatch(sInputEmail) Then
            Return (True)
        Else
            Return (False)
        End If
    End Function
#End Region

End Class
