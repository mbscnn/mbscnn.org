Option Explicit On
Option Strict On
''' <summary>
''' 提供日期相關函式
''' Format...
''' [Titan] 	2011/07/19	Created
''' </summary>
Public Class DateUtility
    ''' <summary>
    ''' DateUtility.showTWYYMMDD(Date)-->100.07
    ''' </summary>
    ''' <param name="data">date</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' [Titan] 	2011/07/19	Created
    ''' </remarks>  
    Public Shared Function showTWYYMM(ByVal data As Object) As String
        Return showTWDate(data, "YYMM", ".")
    End Function
    ''' <summary>
    '''  DateUtility.showTWYYMMDD(Date)-->100.07.20
    ''' </summary>
    ''' <param name="data">date</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' [Titan] 	2011/07/19	Created
    ''' </remarks>  
    Public Shared Function showTWYYMMDD(ByVal data As Object) As String
        Return showTWDate(data, String.Empty, ".")
    End Function

    ''' <summary>
    ''' 格式化民國日期
    ''' DateUtility.showTWDate(Date,"YYMM","-")-->100-07
    ''' DateUtility.showTWDate(Date,"YYMM",".")-->100.07
    ''' DateUtility.showTWDate(Date,"","")-->10007
    ''' DateUtility.showTWDate(Date,"",".")-->100.07.20
    ''' DateUtility.showTWDate(Date,"","-")-->100-07-20
    ''' DateUtility.showTWDate(Date,"","")-->1000720
    ''' </summary>
    ''' <param name="data">date</param>
    ''' <param name="sFormat">YYMM or YYMMDD</param>
    ''' <param name="sSplit">分隔符號</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' [Titan] 	2011/07/19	Created
    ''' </remarks>  
    Public Shared Function showTWDate(ByVal data As Object, ByVal sFormat As String, ByVal sSplit As String) As String
        Try
            If Not IsDate(data) Then
                Return ""
            End If

            Dim twCalendar As New System.Globalization.TaiwanCalendar
            Dim dateTime As New DateTime(CDate(data).Year, CDate(data).Month, CDate(data).Day, New System.Globalization.GregorianCalendar)

            Dim sY As String = CStr(twCalendar.GetYear(dateTime)).ToString.PadLeft(3, CChar("0"))
            Dim sM As String = CStr(twCalendar.GetMonth(dateTime)).PadLeft(2, CChar("0"))
            Dim sD As String = CStr(twCalendar.GetDayOfMonth(dateTime)).PadLeft(2, CChar("0"))

            If (sFormat.ToUpper.Equals("YYMM")) Then
                Return sY & sSplit & sM
            Else
                Return sY & sSplit & sM & sSplit & sD
            End If
        Catch ex As Exception
            Return ""
        End Try


    End Function
    ''' <summary>
    ''' DateUtility.showADYYMMDD(Date)-->2011.07
    ''' </summary>
    ''' <param name="data">date</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' [Titan] 	2011/07/19	Created
    ''' </remarks>  
    Public Shared Function showADYYMM(ByVal data As Object) As String
        Return showADDate(data, "YYYYMM", ".")
    End Function
    ''' <summary>
    ''' DateUtility.showADYYMMDD(Date)-->2011.07.20
    ''' </summary>
    ''' <param name="data">date</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' [Titan] 	2011/07/19	Created
    ''' </remarks>  
    Public Shared Function showADYYMMDD(ByVal data As Object) As String
        Return showADDate(data, String.Empty, ".")
    End Function

    ''' <summary>
    ''' 格式化西元日期
    ''' DateUtility.showADDate(Date,"YYMM","-")-->2011-07
    ''' DateUtility.showADDate(Date,"YYMM",".")-->2011.07
    ''' DateUtility.showADDate(Date,"","")-->201107
    ''' DateUtility.showADDate(Date,"",".")-->2011.07.20
    ''' DateUtility.showADDate(Date,"","-")-->2011-07-20
    ''' DateUtility.showADDate(Date,"","")-->20110720
    ''' </summary>
    ''' <param name="data">date</param>
    ''' <param name="sFormat">YYMM or YYMMDD</param>
    ''' <param name="sSplit">分隔符號</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' [Titan] 	2011/07/19	Created
    ''' </remarks>  
    Public Shared Function showADDate(ByVal data As Object, ByVal sFormat As String, ByVal sSplit As String) As String
        Try
            If Not IsDate(data) Then
                Return ""
            End If

            Dim sY As String = CDate(data).Year.ToString
            Dim sM As String = CDate(data).Month.ToString.PadLeft(2, CChar("0"))
            Dim sD As String = CDate(data).Day.ToString.PadLeft(2, CChar("0"))

            If (sFormat.ToUpper.Equals("YYYYMM")) Then
                Return sY & sSplit & sM
            Else
                Return sY & sSplit & sM & sSplit & sD
            End If
        Catch ex As Exception
            Return ""
        End Try 
    End Function


#Region "1.3.10 DateTransfer() - 轉換日期函式-西曆轉中曆 YYYY/MM/DD->YYY.MM.DD"
    Public Shared Function dateTransfer(ByVal pDate As Object) As String
        If Not IsDate(pDate) Then
            Return ""
        End If
        Dim InDate As Date = CType(pDate, Date)
        Dim RtnStr As String = ""
        Dim Cyear As String = ""
        Cyear = CStr(CInt(Year(InDate)) - 1911).PadLeft(3, CChar("0"))
        RtnStr = Cyear & "." & CStr(Month(InDate)).PadLeft(2, CChar("0")) & "." & CStr(Day(InDate)).PadLeft(2, CChar("0"))
        Return RtnStr
    End Function
#End Region

End Class
