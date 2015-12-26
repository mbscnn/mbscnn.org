Option Explicit On
Option Strict On

Imports System.Text
''' <summary>
''' 提供字串相關函式
''' Format....
''' [Titan] 	2011/07/19	Created
''' </summary>
Public Class StrUtility
    ''' <summary>
    ''' 除了[a-zA-Z0-9_.@-]外，其餘轉為String.Empty
    ''' </summary>
    ''' <param name="str"></param>
    ''' <returns>String</returns>
    ''' <remarks>
    ''' [Titan] 	2011/07/19	Created
    ''' </remarks>  
    Public Shared Function stripCommas(ByVal str As String) As String
        Dim strReg As String = System.Text.RegularExpressions.Regex.Replace(str, "[^\w\.@-]", "")
        Return strReg
    End Function
    ''' <summary>
    ''' 從strSplit字串中先取得strStart位置，再從擷取後的位置找strEnd位置，再擷取完畢。
    ''' </summary>
    ''' <param name="strSplit">input String</param>
    ''' <param name="strStart">Start String</param>
    ''' <param name="strEnd">End String</param>
    ''' <param name="bolPreserveStartEnd">Boolean 是否保留原本搜尋的字串</param>
    ''' <param name="shtToLowOrUpp">Short 全部轉換大(1)小(0)寫</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' [Titan] 	2011/07/19	Created
    ''' </remarks>  
    Public Shared Function strSplitStartEnd(ByRef strSplit As String, ByRef strStart As String, ByRef strEnd As String, Optional ByVal bolPreserveStartEnd As Boolean = False, Optional ByVal shtToLowOrUpp As Short = 0) As String
        Dim str As New StringBuilder(strSplit)

        '// 全部轉換大小寫
        Select Case shtToLowOrUpp
            Case 0 '小寫
                str = New StringBuilder(strSplit.ToLower)
                strStart = strStart.ToLower
                strEnd = strEnd.ToLower
            Case 1 '大寫
                str = New StringBuilder(strSplit.ToUpper)
                strStart = strStart.ToUpper
                strEnd = strEnd.ToUpper
        End Select

        '// 尋找 strStart字串位置.
        Dim intStrPosition As Integer = str.ToString.IndexOf(strStart)
        '// 找不到字串時 會傳回-1 所以要避免發生錯誤的判斷.
        If (intStrPosition <> -1) Then

            '// 是否保留原本搜尋的字串
            If (bolPreserveStartEnd) = False Then

                '// Start字串長度也要刪除
                intStrPosition += strStart.Length
            End If

            '// 擷取Start字串完畢.
            str = New StringBuilder(str.ToString.Substring(intStrPosition))

        End If

        '// 尋找 strEnd結尾位置.
        intStrPosition = str.ToString.IndexOf(strEnd)

        '// 找不到字串時 會傳回-1 所以要避免發生錯誤的判斷.
        If (intStrPosition <> -1) Then

            '// 是否保留原本搜尋的字串.
            If (bolPreserveStartEnd) = True Then

                '// End字串長度也要刪除
                intStrPosition += strEnd.Length
            End If
            str.Remove(intStrPosition, str.Length - intStrPosition)
        End If

        Return (str.ToString)
    End Function

    ''' <summary>
    ''' 員編格式化: 傳回7碼員編, 也可傳回安泰的5碼員編(EnTie5Code=True)
    ''' </summary>
    ''' <param name="inUserID"></param>
    ''' <param name="EnTie5Code"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function FormatUserID(ByVal inUserID As String, Optional ByVal EnTie5Code As Boolean = False) As String
        Dim sRtn As String = String.Empty
        Dim strPrefix As String = "S"
        If Not EnTie5Code Then
            If Left(inUserID, 1) = strPrefix Then
                sRtn = inUserID
            Else
                sRtn = strPrefix & FillZero(inUserID, 6)
            End If
        Else
            If Len(inUserID) = 7 Then
                sRtn = Right(inUserID, 5)
            Else
                sRtn = inUserID
            End If
        End If
        Return sRtn
    End Function

    ''' <summary>
    ''' 補足0位函式
    ''' </summary>
    ''' <param name="sBeAdd">輸入字串</param>
    ''' <param name="iLength">補足長度(字串總長度)</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function FillZero(ByVal sBeAdd As String, ByVal iLength As Integer) As String
        If sBeAdd = "-1911.00.00" Then Return ""
        Dim RtnStr As String = ""
        Dim iloop As Integer
        If Len(sBeAdd) < iLength Then
            RtnStr = sBeAdd
            For iloop = 1 To iLength - Len(sBeAdd)
                RtnStr = "0" & RtnStr
            Next
        Else
            RtnStr = sBeAdd
        End If
        Return RtnStr
    End Function

    ''' <summary>
    ''' 左靠右補
    ''' </summary>
    ''' <param name="sINData">輸入字串</param>
    ''' <param name="sRFill">右補字串</param>
    ''' <param name="iLength">右補長度(字串總長度)</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function FillRightLength(ByVal sINData As String, ByVal sRFill As String, ByVal iLength As Integer) As String
        Try
            If Trim(sINData).Length < iLength Then
                Dim sReturn As String = String.Empty
                sReturn = Trim(sINData)

                For i As Integer = 1 To iLength - sReturn.Length
                    sReturn = sReturn & sRFill
                Next

                Return sReturn
            Else
                Return sINData
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 轉換日期函式-西曆轉中曆 YYYY/MM/DD->YYYMMDD
    ''' </summary>
    ''' <param name="D_Date">西元日期</param>
    ''' <returns>民國年月日</returns>
    ''' <remarks></remarks>
    Public Shared Function ToYYYMMDD(ByVal D_Date As Object) As String
        If Not IsDate(D_Date) Then
            Return String.Empty
        End If

        Dim InDate As Date = CType(D_Date, Date)
        Dim sRtnStr As String = String.Empty
        Dim sCyear As String = String.Empty

        sCyear = CStr(CInt(Year(InDate)) - 1911)

        sRtnStr = sCyear & FillZero(CStr(Month(InDate)), 2) & FillZero(CStr(Day(InDate)), 2)

        Return sRtnStr
    End Function

    ''' <summary>
    ''' 轉換NVarchar String to Char length
    ''' </summary>
    ''' <param name="sNVarchar">NVarchar String</param>
    ''' <param name="iCHARLength">DB CHAR長度</param>
    ''' <returns>string</returns>
    ''' <remarks></remarks>
    Public Shared Function TrimNVarcharToCHARLength(ByVal sNVarchar As Object, ByVal iCHARLength As Decimal) As String
        Try
            If Not IsDBNull(sNVarchar) AndAlso Not IsNothing(sNVarchar) AndAlso CStr(sNVarchar).Length > 0 Then
                Dim sSOURCE As String = CStr(sNVarchar)

                'Dim objByte() As Byte = System.Text.Encoding.GetEncoding("Utf-8").GetBytes(sSOURCE)

                Dim sReturn As String = String.Empty

                'ReDim objByte(iCHARLength)

                'Dim encoding_ASCII As New System.Text.ASCIIEncoding()
                'Dim encoding_UTF8 As New System.Text.UTF8Encoding

                'sReturn = encoding_ASCII.GetString(objByte)
                'sReturn = encoding_UTF8.GetString(objByte)

                Dim iCount As Decimal = 0
                For i As Integer = 0 To sSOURCE.Length - 1
                    Dim sTEMP As String = String.Empty
                    sTEMP = sSOURCE.Substring(i, 1)
                    iCount += Encoding.GetEncoding("Big5").GetByteCount(sTEMP)
                    If iCount <= iCHARLength Then
                        sReturn &= sTEMP
                    Else
                        Exit For
                    End If
                Next

                Return sReturn
            End If

            Return String.Empty
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' 轉換NVarchar String to Char length
    ''' </summary>
    ''' <param name="sNVarchar">NVarchar String</param>
    ''' <param name="iCHARLength">DB CHAR長度</param>
    ''' <returns>string</returns>
    ''' <remarks></remarks>
    Public Shared Function getNVarcharLength(ByVal sNVarchar As Object) As Decimal
        Try
            If Not IsDBNull(sNVarchar) AndAlso Not IsNothing(sNVarchar) AndAlso CStr(sNVarchar).Length > 0 Then
                Dim sSOURCE As String = CStr(sNVarchar)

                'Dim objByte() As Byte = System.Text.Encoding.GetEncoding("Utf-8").GetBytes(sSOURCE)

                Dim sReturn As String = String.Empty

                'ReDim objByte(iCHARLength)

                'Dim encoding_ASCII As New System.Text.ASCIIEncoding()
                'Dim encoding_UTF8 As New System.Text.UTF8Encoding

                'sReturn = encoding_ASCII.GetString(objByte)
                'sReturn = encoding_UTF8.GetString(objByte)

                Dim iCount As Decimal = 0
                For i As Integer = 0 To sSOURCE.Length - 1
                    Dim sTEMP As String = String.Empty
                    sTEMP = sSOURCE.Substring(i, 1)
                    iCount += Encoding.GetEncoding("Big5").GetByteCount(sTEMP)

                    'iCount += Encoding.ASCII.GetByteCount(sTEMP)
                Next

                Return iCount
            End If

            Return 0
        Catch ex As Exception
            Throw
        End Try
    End Function
End Class
