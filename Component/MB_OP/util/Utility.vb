Public Class Utility

#Region "Date"

    '取得民國年月日 return 民國95年04月20日
    Public Shared Function getROCYYMMDD() As String
        Dim time As DateTime = Now()
        Dim iY As Int16 = Int16.Parse(time.Year.ToString)
        Dim sYear As String = (iY - 1911).ToString '民國年
        Dim sMonth As String = time.Month.ToString
        Dim sDay As String = time.Day.ToString
        If (sMonth.Length = 1) Then
            sMonth = "0" & sMonth
        End If
        If (sDay.Length = 1) Then
            sDay = "0" & sDay
        End If
        Return "民國" & (sYear) & "年" & (sMonth) & "月" & (sDay) & "日"
    End Function

    '取得民國年月日 
    'return 95年04月20日
    Public Shared Function getROCYYMMDD1() As String
        Dim time As DateTime = Now()
        Dim iY As Int16 = Int16.Parse(time.Year.ToString)
        Dim sYear As String = (iY - 1911).ToString '民國年
        Dim sMonth As String = time.Month.ToString
        Dim sDay As String = time.Day.ToString
        If (sMonth.Length = 1) Then
            sMonth = "0" & sMonth
        End If
        If (sDay.Length = 1) Then
            sDay = "0" & sDay
        End If
        Return (sYear) & "年" & (sMonth) & "月" & (sDay) & "日"
    End Function

    '取得現在民國年月日 
    'return 95年04月20日,95年04月底
    Public Shared Function getROCYYMMDD2() As String
        Dim time As DateTime = Now()
        Dim iY As Int16 = Int16.Parse(time.Year.ToString)
        Dim sYear As String = (iY - 1911).ToString '民國年
        Dim sMonth As String = time.Month.ToString
        Dim sDay As String = time.Day.ToString
        If (sMonth.Length = 1) Then
            sMonth = "0" & sMonth
        End If
        If (sDay.Length = 1) Then
            sDay = "0" & sDay
        End If
        Return (sYear) & "年" & (sMonth) & "月" & (sDay) & "日," & (sYear) & "年" & (sMonth) & "月底"
    End Function

    '取得民國年月(這個月的最後一天)日 
    'input 200601 or 200602 or ....
    'return 95年01月31日 or 95年02月28日 or 95年03月31日 or ....
    Public Shared Function getROCYYMMDD3(ByVal str As String) As String
        str = str.Trim
        If (str.Length <> 6) Then
            Throw New Exception("日期格式不正確 [" & str & "]")
        End If
        str = str.Insert(4, "/") & "/01"
        Dim time As DateTime = Convert.ToDateTime(str)
        Dim iY As Int16 = Int16.Parse(time.Year.ToString)
        Dim sYear As String = (iY - 1911).ToString '民國年
        Dim sMonth As String = time.Month.ToString
        Dim sDay As String = Date.DaysInMonth(iY, time.Month)
        If (sMonth.Length = 1) Then
            sMonth = "0" & sMonth
        End If
        If (sDay.Length = 1) Then
            sDay = "0" & sDay
        End If
        Return (sYear) & "年" & (sMonth) & "月" & (sDay) & "日"
    End Function

    '取得民國年月(這個月的最後一天)日 return 95年01月31日,95年01月底
    Public Shared Function getROCYYMMDD4(ByVal str As String) As String
        str = str.Trim
        If (str.Length <> 6) Then
            Throw New Exception("日期格式不正確 [" & str & "]")
        End If
        str = str.Insert(4, "/") & "/01"
        Dim time As DateTime = Convert.ToDateTime(str)
        Dim iY As Int16 = Int16.Parse(time.Year.ToString)
        Dim sYear As String = (iY - 1911).ToString '民國年
        Dim sMonth As String = time.Month.ToString
        Dim sDay As String = Date.DaysInMonth(iY, time.Month)
        If (sMonth.Length = 1) Then
            sMonth = "0" & sMonth
        End If
        If (sDay.Length = 1) Then
            sDay = "0" & sDay
        End If
        Return (sYear) & "年" & (sMonth) & "月" & (sDay) & "日," & (sYear) & "年" & (sMonth) & "月底"
    End Function

    '取得民國年月日 
    'input str="2006/4/19"
    'return 95年04月19日
    Public Shared Function getROCYYMMDD5(ByVal str As String) As String
        Dim time As DateTime = Convert.ToDateTime(str)
        Dim iY As Int16 = Int16.Parse(time.Year.ToString)
        Dim sYear As String = (iY - 1911).ToString '民國年
        Dim sMonth As String = time.Month.ToString
        Dim sDay As String = time.Day
        If (sMonth.Length = 1) Then
            sMonth = "0" & sMonth
        End If
        If (sDay.Length = 1) Then
            sDay = "0" & sDay
        End If
        Return (sYear) & "年" & (sMonth) & "月" & (sDay) & "日"
    End Function

    '取得民國年月日 
    'input str="2006/4/19"
    'return 95.04.19
    Public Shared Function getROCYYMMDD6(ByVal str As String) As String

        If IsNothing(str) OrElse str.Length = 0 Then
            Return ""
        End If

        Dim time As DateTime = Convert.ToDateTime(str)
        Dim iY As Int16 = Int16.Parse(time.Year.ToString)
        Dim sYear As String = (iY - 1911).ToString '民國年
        Dim sMonth As String = time.Month.ToString
        Dim sDay As String = time.Day
        If (sMonth.Length = 1) Then
            sMonth = "0" & sMonth
        End If
        If (sDay.Length = 1) Then
            sDay = "0" & sDay
        End If
        Return (sYear) & "." & (sMonth) & "." & (sDay)
    End Function

    '取得民國年 return 民國95年
    Public Shared Function getROCYY() As String
        Dim time As DateTime = Now()
        Dim iY As Int16 = Int16.Parse(time.Year.ToString)
        Dim sYear As String = (iY - 1911).ToString '民國年
        Return "民國" & (sYear) & "年"
    End Function

    '取得民國年 
    'input 200602 
    'return 95
    Public Shared Function getROCYY(ByVal str As String) As String
        str = str.Trim
        If (str.Length <> 6) Then
            Throw New Exception("日期格式不正確 [" & str & "]")
        End If
        str = str.Insert(4, "/") & "/01"
        Dim time As DateTime = Convert.ToDateTime(str)
        Dim iY As Int16 = Int16.Parse(time.Year.ToString)
        Dim sYear As String = (iY - 1911).ToString '民國年
        Dim sMonth As String = time.Month.ToString
        If (sMonth.Length = 1) Then
            sMonth = "0" & sMonth
        End If
        Return (sYear)

    End Function


    '取得民國年月 
    'input 200602 
    'return 9502
    Public Shared Function getROCYYYMM(ByVal str As String) As String
        str = str.Trim
        If (str.Length <> 6) Then
            Throw New Exception("日期格式不正確 [" & str & "]")
        End If
        str = str.Insert(4, "/") & "/01"
        Dim time As DateTime = Convert.ToDateTime(str)
        Dim iY As Int16 = Int16.Parse(time.Year.ToString)
        Dim sYear As String = (iY - 1911).ToString '民國年
        Dim sMonth As String = time.Month.ToString
        If (sMonth.Length = 1) Then
            sMonth = "0" & sMonth
        End If
        Return sYear & sMonth

    End Function

    '取得民國年月 return 95年04月
    Public Shared Function getROCYYMM() As String
        Dim time As DateTime = Now()
        Dim iY As Int16 = Int16.Parse(time.Year.ToString)
        Dim sYear As String = (iY - 1911).ToString '民國年
        Dim sMonth As String = time.Month.ToString
        If (sMonth.Length = 1) Then
            sMonth = "0" & sMonth
        End If
        Return (sYear) & "年" & (sMonth) & "月"
    End Function

    '取得民國年月 return 95年04月
    Public Shared Function getROCYYMM(ByVal str As String) As String
        str = str.Trim
        If (str.Length <> 6) Then
            Throw New Exception("日期格式不正確 [" & str & "]")
        End If
        str = str.Insert(4, "/") & "/01"
        Dim time As DateTime = Convert.ToDateTime(str)
        Dim iY As Int16 = Int16.Parse(time.Year.ToString)
        Dim sYear As String = (iY - 1911).ToString '民國年
        Dim sMonth As String = time.Month.ToString
        If (sMonth.Length = 1) Then
            sMonth = "0" & sMonth
        End If
        Return (sYear) & "年" & (sMonth) & "月"
    End Function

    '取得   民國年月 本季最後一月
    'input  199901 or 1999902 or 1999903
    'return 民國88年03月
    Public Shared Function getROCQuarterYYMM(ByVal str As String) As String

        Dim strQua As String = getQuarterYYMM(str).Split(",")(0) '1999903
        Return getROCYYMM(strQua)
    End Function

    '西元加減年月日 return 1999/04/03
    Public Shared Function getYYMMDD(ByVal iYear As Int16, ByVal iMonth As Int16, ByVal iDay As Int16) As String
        Dim time As DateTime = Now().AddYears(iYear).AddMonths(iMonth).AddDays(iDay)
        Dim sCEYear As String = time.Year.ToString
        Dim sMonth As String = time.Month.ToString
        Dim sDay As String = time.Day.ToString
        If (sMonth.Length = 1) Then
            sMonth = "0" & sMonth
        End If
        If (sDay.Length = 1) Then
            sDay = "0" & sDay
        End If
        Return sCEYear & "/" & sMonth & "/" & sDay
    End Function

    '取得  西元年 return 1999/04/06
    Public Shared Function getYYMMDD() As String
        Dim time As DateTime = Now()
        Dim sMonth As String = time.Month.ToString
        Dim sDay As String = time.Day.ToString
        If (sMonth.Length = 1) Then
            sMonth = "0" & sMonth
        End If
        If (sDay.Length = 1) Then
            sDay = "0" & sDay
        End If
        Return (time.Year.ToString) & "/" & (sMonth) & "/" & (sDay)
    End Function

    '取得  西元 年月,年(-1)月 return 199904,199903 
    Public Shared Function getYYdecMM1() As String
        Dim time1 As DateTime = Now()
        Dim sYear1 As String = time1.Year.ToString
        Dim sMonth1 As String = time1.Month.ToString

        If (sMonth1.Length = 1) Then
            sMonth1 = "0" & sMonth1
        End If

        Dim time2 As DateTime = time1.AddMonths(-1)
        Dim sYear2 As String = time2.Year.ToString
        Dim sMonth2 As String = time2.Month.ToString

        If (sMonth2.Length = 1) Then
            sMonth2 = "0" & sMonth2
        End If

        Return sYear1 & sMonth1 & "," & sYear2 & sMonth2
    End Function

    '取得  西元 年月,年(-1)月 
    'input 199904
    'return 199904,199903
    Public Shared Function getYYdecMM1(ByVal str As String) As String
        str = str.Trim
        Dim str1 As String
        If (str.Length <> 6) Then
            Throw New Exception("日期格式不正確 [" & str & "]")
        End If
        str1 = str.Insert(4, "/") & "/01"
        Dim time As DateTime = Convert.ToDateTime(str1)

        Dim time2 As DateTime = time.AddMonths(-1)
        Dim sYear2 As String = time2.Year.ToString
        Dim sMonth2 As String = time2.Month.ToString

        If (sMonth2.Length = 1) Then
            sMonth2 = "0" & sMonth2
        End If
        Return str & "," & sYear2 & sMonth2
    End Function

    '取得  西元 年月,(-1)年月 
    'input 199904
    'return 199904,199804
    Public Shared Function getdecYYMM2() As String
        Dim time1 As DateTime = Now()
        Dim sYear1 As String = time1.Year.ToString
        Dim sMonth1 As String = time1.Month.ToString
        Dim time2 As DateTime = time1.AddYears(-1)
        Dim sYear2 As String = time2.Year.ToString
        Dim sMonth2 As String = time2.Month.ToString

        If (sMonth1.Length = 1) Then
            sMonth1 = "0" & sMonth1
        End If
        If (sMonth2.Length = 1) Then
            sMonth2 = "0" & sMonth2
        End If

        Return sYear1 & sMonth1 & "," & sYear2 & sMonth2
    End Function

    '取得  西元 年月,(-1)年月 
    'input 200601
    'return 200601,200501
    Public Shared Function getdecYYMM2(ByVal str As String) As String
        str = str.Trim
        If (str.Length <> 6) Then
            Throw New Exception("日期格式不正確 [" & str & "]")
        End If
        Dim str1 As String
        str1 = str.Insert(4, "/") & "/01" '200601 ==> 2006/01/01

        Dim time As DateTime = Convert.ToDateTime(str1)

        Dim time1 As DateTime = time.AddYears(-1)
        Dim sYear1 As String = time1.Year.ToString
        Dim sMonth1 As String = time1.Month.ToString

        If (sMonth1.Length = 1) Then
            sMonth1 = "0" & sMonth1
        End If

        Return str & "," & sYear1 & sMonth1
    End Function

    '取得  西元 年月,年(-1)月 ,(-1)年月 
    'input 200601
    'return 200601,200512,200501
    Public Shared Function getdecYYdecMM3(ByVal str As String) As String
        Dim sb As New System.Text.StringBuilder
        str = getYYdecMM1(str)
        sb.Append(str)
        str = getdecYYMM2(str.Split(",")(0))
        sb.Append(",").Append(str.Split(",")(1))
        Return sb.ToString
    End Function

    '取得  西元年月
    'input 9501 民國年月
    'return 200601
    Public Shared Function getROC2ADYYYYMM(ByVal str As String) As String
        str = str.Trim

        If str.Length = 5 Then
            str = getROC2ADYYYY(str.Substring(0, 3)) & str.Substring(3, 2)  'ROC 2 AD '10001 ==> 201101
        ElseIf str.Length = 4 Then
            str = getROC2ADYYYY(str.Substring(0, 2)) & str.Substring(2, 2)  'ROC 2 AD '9501 ==> 200601
        ElseIf str.Length = 3 Then
            str = getROC2ADYYYY(str.Substring(0, 1)) & str.Substring(1, 2) 'ROC 2 AD '501 ==> 191601
        Else
            Throw New Exception("(getROC2ADYYMM()):日期格式不正確 [" & str & "]")
        End If
        Return str
    End Function

    '取得  西元年月日
    'input 0950112
    'return 2006/01/01
    Public Shared Function getROC2ADYYMMDD(ByVal str As String) As String
        str = str.PadLeft(7, "0"c)
        If IsNothing(str) OrElse str.Length <> 7 Then
            Return str
        End If
        Dim sY As String = getROC2ADYYYY(str.Substring(0, 3))
        Dim sM As String = str.Substring(3, 2)
        Dim sD As String = str.Substring(5, 2)
        Return sY & "/" & sM & "/" & sD
    End Function

    '取得  民國年月
    'input 95.01.12
    'return 20060101
    Public Shared Function getROC2ADYMD(ByVal str As String, ByVal sSplit As String) As String

        sSplit = sSplit.Trim
        If IsNothing(str) OrElse str.Length = 0 Then
            Return ""
        Else
            str = str.Trim
            If (str.Split(sSplit).Length <> 3) Then
                Return str
            End If
        End If
        Dim sY As String = str.Split(sSplit)(0).Trim
        Dim sM As String = str.Split(sSplit)(1).Trim
        Dim sD As String = str.Split(sSplit)(2).Trim

        Try
            If (sM.Length = 1) Then
                sM = "0" & sM
            End If
            If (sD.Length = 1) Then
                sD = "0" & sD
            End If
            str = getROC2ADYYYY(sY) & sM & sD
        Catch ex As Exception
            Throw New Exception("(getROC2ADYMD()):日期格式不正確 [" & str & "]")
        End Try

        Return str
    End Function

    '取得  民國年月
    'input 95,01,12
    'return 20060101
    Public Shared Function getROC2ADYMD(ByVal str As String) As String

        If IsNothing(str) OrElse str.Length = 0 Then
            Return ""
        Else
            str = str.Trim
            If (str.Split(".").Length <> 3) Then
                Return str
            End If
        End If
        Dim sY As String = str.Split(".")(0).Trim
        Dim sM As String = str.Split(".")(1).Trim
        Dim sD As String = str.Split(".")(2).Trim

        Try
            If (sM.Length = 1) Then
                sM = "0" & sM
            End If
            If (sD.Length = 1) Then
                sD = "0" & sD
            End If
            str = getROC2ADYYYY(sY) & sM & sD
        Catch ex As Exception
            Throw New Exception("(getROC2ADYMD()):日期格式不正確 [" & str & "]")
        End Try

        Return str
    End Function

    '取得  西元年
    'input 95 民國年
    'return 2006
    Public Shared Function getROC2ADYYYY(ByVal str As String) As String
        str = str.Trim
        If str.Length = 3 Then
            str = Int16.Parse(str.Substring(0, 3)) + 1911  'ROC 2 AD '100 ==> 2011
        ElseIf str.Length = 2 Then
            str = Int16.Parse(str.Substring(0, 2)) + 1911  'ROC 2 AD '95 ==> 2006
        ElseIf str.Length = 1 Then
            str = Int16.Parse(str.Substring(0, 1)) + 1911  'ROC 2 AD '5 ==> 1916
        Else
            Throw New Exception("(getROC2ADYYYY()):日期格式不正確 [" & str & "]")
        End If
        Return str
    End Function

    '取得目前  西元 年月 return 199904
    Public Shared Function getYYMM() As String
        Dim time1 As DateTime = Now()
        Dim sYear1 As String = time1.Year.ToString
        Dim sMonth1 As String = time1.Month.ToString

        If (sMonth1.Length = 1) Then
            sMonth1 = "0" & sMonth1
        End If
        Return sYear1 & sMonth1
    End Function

    '以目前的時間 取得  西元年月 本季底月,前一季底月 
    'return 199903,1998012  
    Public Shared Function getQuarterYYMM() As String

        Dim time1 As DateTime = Now
        Dim sYear1 As String = time1.Year.ToString
        Dim sMonth1 As String = DatePart(DateInterval.Quarter, time1) * 3

        Dim time2 As DateTime = time1.AddMonths(-3)
        Dim sYear2 As String = time2.Year.ToString
        Dim sMonth2 As String = DatePart(DateInterval.Quarter, time2) * 3
        If (sMonth1.Length = 1) Then
            sMonth1 = "0" & sMonth1
        End If
        If (sMonth2.Length = 1) Then
            sMonth2 = "0" & sMonth2
        End If

        Return sYear1 & sMonth1 & "," & sYear2 & sMonth2
    End Function

    '取得   西元年月 本季,前一季 
    'input  199901 or 199902 or 199903
    'return         199903,1998012  
    Public Shared Function getQuarterYYMM(ByVal str As String) As String
        str = str.Trim
        If (str.Length <> 6) Then
            Throw New Exception("日期格式不正確 [" & str & "]")
        End If
        str = str.Insert(4, "/") & "/01"
        Dim time1 As DateTime = Convert.ToDateTime(str)
        Dim sYear1 As String = time1.Year.ToString
        Dim sMonth1 As String = DatePart(DateInterval.Quarter, time1) * 3

        Dim time2 As DateTime = time1.AddMonths(-3)
        Dim sYear2 As String = time2.Year.ToString
        Dim sMonth2 As String = DatePart(DateInterval.Quarter, time2) * 3
        If (sMonth1.Length = 1) Then
            sMonth1 = "0" & sMonth1
        End If
        If (sMonth2.Length = 1) Then
            sMonth2 = "0" & sMonth2
        End If

        Return sYear1 & sMonth1 & "," & sYear2 & sMonth2
    End Function

    '取得  西元年月 本季的三各月份 return 200603,200602,200601 or 200606,200605,200604(由大月-->小月)
    Public Shared Function getQuarterYYMM2() As String
        Dim sb As New System.Text.StringBuilder
        Dim time1 As DateTime = Now
        Dim sYear1 As String = time1.Year.ToString

        Dim sMonth1 As String
        sMonth1 = DatePart(DateInterval.Quarter, time1) * 3
        If (sMonth1.Length = 1) Then
            sMonth1 = "0" & sMonth1
        End If
        Dim iM As Int16 = Int16.Parse(sMonth1)
        Dim time2 As DateTime = time1.AddMonths(-3) '前一季
        Dim iM2 As Int16 = DatePart(DateInterval.Quarter, time2) * 3
        Dim sMonth2 As String
        If (iM < iM2) Then
            iM2 = 0
        End If
        For iM = iM To iM2 Step -1
            If iM = iM2 Or iM = 0 Then
                Exit For
            End If
            sMonth2 = iM
            If (sMonth2.Length = 1) Then
                sMonth2 = "0" & sMonth2
            End If
            sb.Append(sYear1).Append(sMonth2).Append(",")
        Next
        sb.Remove(sb.Length - 1, 1)
        Return sb.ToString
    End Function

    '取得  西元年月 本季的三各月份 return 200603,200602,200601 or 200606,200605,200604
    Public Shared Function getQuarterYYMM2(ByVal str As String) As String
        Dim sb As New System.Text.StringBuilder
        str = str.Trim
        If (str.Length <> 6) Then
            Throw New Exception("日期格式不正確 [" & str & "]")
        End If
        str = str.Insert(4, "/") & "/01"
        Dim time1 As DateTime = Convert.ToDateTime(str)
        Dim sYear1 As String = time1.Year.ToString

        Dim sMonth1 As String
        sMonth1 = DatePart(DateInterval.Quarter, time1) * 3
        If (sMonth1.Length = 1) Then
            sMonth1 = "0" & sMonth1
        End If
        Dim iM As Int16 = Int16.Parse(sMonth1)
        Dim time2 As DateTime = time1.AddMonths(-3) '前一季
        Dim iM2 As Int16 = DatePart(DateInterval.Quarter, time2) * 3
        Dim sMonth2 As String
        If (iM < iM2) Then
            iM2 = 0
        End If
        For iM = iM To iM2 Step -1
            If iM = iM2 Or iM = 0 Then
                Exit For
            End If
            sMonth2 = iM
            If (sMonth2.Length = 1) Then
                sMonth2 = "0" & sMonth2
            End If
            sb.Append(sYear1).Append(sMonth2).Append(",")
        Next

        sb.Remove(sb.Length - 1, 1)
        Return sb.ToString
    End Function

    'input 200607
    'return 200601,200602,200603,200604,200605,200606
    'input 200601
    'return 200501,200502,200503,200504,200505,200506,200507,200508,200509,200510,200511,200512
    Public Shared Function getAnnualYM(ByVal str As String) As String
        Dim sb As New System.Text.StringBuilder
        str = str.Trim
        If (str.Length <> 6) Then
            Throw New Exception("日期格式不正確 [" & str & "]")
        End If
        str = str.Insert(4, "/") & "/01" '組成日期格式2000/03/01
        Dim time As DateTime = Convert.ToDateTime(str)
        Dim iYear As Int16 = time.Year
        Dim iMonth As Int16 = time.Month
        If (iMonth <= 7) And (iMonth <> 1) Then
            For i As Int16 = 1 To 6
                sb.Append(iYear).Append(0).Append(i).Append(",")
            Next
        ElseIf (iMonth = 1) Then
            For i As Int16 = 1 To 12
                If i <= 9 Then
                    sb.Append(iYear - 1).Append(0).Append(i).Append(",")
                Else
                    sb.Append(iYear - 1).Append(i).Append(",")
                End If

            Next
        End If
        If (sb.Length > 1) Then
            sb.Remove(sb.Length - 1, 1)
        End If

        Return sb.ToString
    End Function

    'input 200607
    'return 200510,200511,200512,200601,200602,200603
    'input 200601
    'return 200510,200511,200512,200601,200602,200603,200604,200605,200606,200607,200608,200609
    Public Shared Function getAnnualYM1(ByVal str As String) As String
        Dim sb As New System.Text.StringBuilder
        str = str.Trim
        If (str.Length = 4) Then
            str = (Int16.Parse(str.Substring(0, 4)) + 1) & "01"
        End If
        Dim iMonth As Int16 = str.Substring(4.2)
        If (iMonth <= 7) And (iMonth <> 1) Then
            '取去年的最後一季,的三各月份
            Dim sDate As String = str
            sDate = Int16.Parse(sDate.Substring(0, 4)) - 1 & "12"
            Dim sQMonth() As String = getQuarterYYMM2(sDate).Split(",")
            sb.Append(sQMonth(2)).Append(",").Append(sQMonth(1)).Append(",").Append(sQMonth(0)).Append(",")
        ElseIf (iMonth = 1) Then
            Dim sDate As String = str
            sDate = Int16.Parse(sDate.Substring(0, 4)) - 2 & "12"
            Dim sQMonth() As String = getQuarterYYMM2(sDate).Split(",")
            sb.Append(sQMonth(2)).Append(",").Append(sQMonth(1)).Append(",").Append(sQMonth(0)).Append(",")
        End If


        str = str.Insert(4, "/") & "/01" '組成日期格式2000/03/01
        Dim time As DateTime = Convert.ToDateTime(str)
        Dim iYear As Int16 = time.Year
        iMonth = time.Month

        If (iMonth <= 7) And (iMonth <> 1) Then
            For i As Int16 = 1 To 3
                sb.Append(iYear).Append(0).Append(i).Append(",")
            Next
        ElseIf (iMonth = 1) Then
            For i As Int16 = 1 To 9
                If i <= 9 Then
                    sb.Append(iYear - 1).Append(0).Append(i).Append(",")
                Else
                    sb.Append(iYear - 1).Append(i).Append(",")
                End If

            Next
        End If
        If (sb.Length > 1) Then
            sb.Remove(sb.Length - 1, 1)
        End If

        Return sb.ToString
    End Function

    '取得加減年月日  西元年,民國年 return 1999/04/20,88/04/20
    Public Shared Function getYMD(ByVal iYear As Int16, ByVal iMonth As Int16, ByVal iDay As Int16) As String
        Dim time As DateTime = Now().AddYears(iYear).AddMonths(iMonth).AddDays(iDay)
        Dim sCEYear As String = time.Year.ToString '西元年
        Dim iY As Int16 = Int16.Parse(time.Year.ToString)
        Dim sYear As String = (iY - 1911).ToString '民國年
        Dim sMonth As String = time.Month.ToString
        Dim sDay As String = time.Day.ToString
        If (sMonth.Length = 1) Then
            sMonth = "0" & sMonth
        End If
        If (sDay.Length = 1) Then
            sDay = "0" & sDay
        End If
        Return sCEYear & "/" & sMonth & "/" & sDay & "," & sYear & "/" & sMonth & "/" & sDay
    End Function

    '取得加減年月日  西元年月,民國年月 return 1999/04,88/04
    Public Shared Function getYM(ByVal iYear As Int16, ByVal iMonth As Int16) As String
        Dim time As DateTime = Now().AddYears(iYear).AddMonths(iMonth)
        Dim sCEYear As String = time.Year.ToString '西元年
        Dim iY As Int16 = Int16.Parse(time.Year.ToString)
        Dim sYear As String = (iY - 1911).ToString '民國年
        Dim sMonth As String = time.Month.ToString
        If (sMonth.Length = 1) Then
            sMonth = "0" & sMonth
        End If
        Return sCEYear & "/" & sMonth & "," & sYear & "/" & sMonth
    End Function

    '取得加減年月日  西元年月,民國年月 return 199904,88/04
    Public Shared Function getYM1(ByVal iYear As Int16, ByVal iMonth As Int16) As String
        Dim time As DateTime = Now().AddYears(iYear).AddMonths(iMonth)
        Dim sCEYear As String = time.Year.ToString '西元年
        Dim iY As Int16 = Int16.Parse(time.Year.ToString)
        Dim sYear As String = (iY - 1911).ToString '民國年
        Dim sMonth As String = time.Month.ToString
        If (sMonth.Length = 1) Then
            sMonth = "0" & sMonth
        End If
        Return sCEYear & sMonth & "," & sYear & "/" & sMonth
    End Function

    '取得  西元年,民國年 return 1999/4/20,88/4/20
    Public Shared Function getYMD() As String
        Dim time As DateTime = Now()
        Dim iY As Int16 = Int16.Parse(time.Year.ToString)
        Dim sYear As String = (iY - 1911).ToString '民國年
        Return (time.Year.ToString) & "/" & (time.Month.ToString) & "/" & (time.Day.ToString) & "," & (sYear) & "/" & (time.Month.ToString) & "/" & (time.Day.ToString)
    End Function

    '取得  西元年 return 1999
    Public Shared Function getYY()
        Return (Now().Year.ToString)
    End Function

    Public Shared Function getMM()
        Return (Now().Month.ToString)
    End Function

    Public Shared Function getDD()
        Return (Now().Day.ToString)
    End Function


    Public Shared Function showTWDate(ByVal data As Date, ByVal sFormat As String, ByVal sSplit As String) As String

        If Utility.isValidateData(data) Then
            Dim twCalendar As New System.Globalization.TaiwanCalendar
            Dim dateTime As New dateTime(data.Year, data.Month, data.Day, New System.Globalization.GregorianCalendar)

            If dateTime.Year > 1911 Then
                Dim sY As String = twCalendar.GetYear(dateTime)

                Dim sM As String = twCalendar.GetMonth(dateTime)

                Dim sD As String = twCalendar.GetDayOfMonth(dateTime)

                If sM < 10 Then
                    sM = "0" + sM
                End If

                If sD < 10 Then
                    sD = "0" + sD
                End If

                If (sFormat.ToUpper.Equals("YYMM")) Then
                    Return sY & sSplit & sM
                ElseIf (sFormat.ToUpper.Equals("YYMMDD")) Then
                    Return sY & sSplit & sM & sSplit & sD
                ElseIf (sFormat.ToUpper.Equals("YYYMMDD")) Then
                    If sY < 100 Then
                        sY = "0" + sY
                    End If
                    Return sY & sSplit & sM & sSplit & sD
                ElseIf (sFormat.ToUpper.Equals("YYYMM")) Then
                    If sY < 100 Then
                        sY = "0" + sY
                    End If
                    Return sY & sSplit & sM
                Else
                    Return sY & sM & sD
                End If
            Else
                Return ""
            End If
        Else
            Return ""
        End If
    End Function

    Public Shared Function showGregorianDate(ByVal data As Object, ByVal sFormat As String, ByVal sSplit As String) As String

        If Not IsNothing(data) Or Not IsDBNull(data) Then
            Dim greCalendar As New System.Globalization.GregorianCalendar
            Dim dateTime As New dateTime(data.Year, data.Month, data.Day, New System.Globalization.GregorianCalendar)
            Dim sY As String = greCalendar.GetYear(dateTime)
            Dim sM As String = greCalendar.GetMonth(dateTime)
            Dim sD As String = greCalendar.GetDayOfMonth(dateTime)

            If sY < 100 Then
                sY = "0" + sY
            End If

            If sM < 10 Then
                sM = "0" + sM
            End If

            If sD < 10 Then
                sD = "0" + sD
            End If

            If (sFormat.ToUpper.Equals("YYYYMM")) Then
                Return sY & sSplit & sM
            ElseIf (sFormat.ToUpper.Equals("YYYYMMDD")) Then
                Return sY & sSplit & sM & sSplit & sD
            Else
                Return sY & sM & sD
            End If
        Else
            Return ""
        End If
    End Function

    Public Shared Function showTWYYMM(ByVal data As Object) As String
        Return showTWDate(data, "YYMM", ".")
    End Function

    Public Shared Function showTWYYMMDD(ByVal data As Object) As String
        If IsNothing(data) Or IsDBNull(data) Then
            Return ""
        End If
        Return showTWDate(data, "YYMMDD", ".")
    End Function

    Public Shared Function showTWYYYMMDD(ByVal data As Object) As String
        If IsNothing(data) Or IsDBNull(data) Then
            Return ""
        End If

        Return showTWDate(data, "YYYMMDD", ".")
    End Function

    '[轉換日期函式-西曆轉中曆 YYYY/MM/DD->YYYMMDD]
    Public Shared Function DateTransfer1(ByVal pDate As Object) As String
        If Not IsDate(pDate) Then
            Return ""
        End If
        Dim InDate As Date = CType(pDate, Date)
        Dim RtnStr As String = ""
        Dim Cyear As String = ""
        Cyear = CStr(CInt(Year(InDate)) - 1911)
        'modify by ted 20061128 0951128 --> 95.11.28
        'RtnStr = FillZero(Cyear, 3) & FillZero(CStr(Month(InDate)), 2) & FillZero(CStr(Day(InDate)), 2)
        RtnStr = FillZero(Cyear, 3) & FillZero(CStr(Month(InDate)), 2) & FillZero(CStr(Day(InDate)), 2)
        Return RtnStr
    End Function

    '[轉換日期函式-西曆轉中曆 YYYY/MM/DD HH/MM/SS->YYY.MM.DD HH:MM:SS]
    Public Shared Function DateTransfer2(ByVal pDate As Object) As String
        If Not IsDate(pDate) Then
            Return ""
        End If
        Dim InDate As Date = CType(pDate, Date)
        Dim RtnStr As String = ""
        Dim Cyear As String = ""
        Cyear = CStr(CInt(Year(InDate)) - 1911)
        'modify by ted 20061128 0951128 --> 95.11.28
        'RtnStr = FillZero(Cyear, 3) & FillZero(CStr(Month(InDate)), 2) & FillZero(CStr(Day(InDate)), 2)
        RtnStr = Cyear & "." & FillZero(CStr(Month(InDate)), 2) & "." & FillZero(CStr(Day(InDate)), 2) & " " & FillZero(CStr(Hour(pDate)), 2) & ":" & FillZero(CStr(Minute(pDate)), 2) & ":" & FillZero(CStr(Second(pDate)), 2)
        Return RtnStr
    End Function

    '[轉換日期函式-西曆轉中曆 YYYY/MM/DD HH/MM->YYY.MM.DD HH:MM]
    Public Shared Function DateTransfer3(ByVal pDate As Object) As String
        If Not IsDate(pDate) Then
            Return ""
        End If
        Dim InDate As Date = CType(pDate, Date)
        Dim RtnStr As String = ""
        Dim Cyear As String = ""
        Cyear = CStr(CInt(Year(InDate)) - 1911)
        'modify by ted 20061128 0951128 --> 95.11.28
        'RtnStr = FillZero(Cyear, 3) & FillZero(CStr(Month(InDate)), 2) & FillZero(CStr(Day(InDate)), 2)
        RtnStr = Cyear & "." & FillZero(CStr(Month(InDate)), 2) & "." & FillZero(CStr(Day(InDate)), 2) & " " & FillZero(CStr(Hour(pDate)), 2) & ":" & FillZero(CStr(Minute(pDate)), 2)
        Return RtnStr
    End Function


    '[取得民國年]
    Public Shared Function DateTransferYYYY(ByVal pDate As Object) As String
        If IsDate(pDate) Then
            Dim InDate As Date = CType(pDate, Date)
            DateTransferYYYY = CStr(CInt(Year(InDate) - 1911))
        Else
            DateTransferYYYY = ""
        End If
    End Function

    '轉換民國年為西元年
    Public Shared Function BackToDATE(ByVal InDate As String) As Object
        Dim RtnDate As Object = Convert.DBNull
        Dim AryDate As Array
        Dim wkStr As String
        Dim RtnStr As String = ""
        If InStr(1, InDate, ".") <> 0 Then
            AryDate = Split(InDate, ".")
            If AryDate.Length >= 3 Then
                wkStr = CType(CType(AryDate(0), Int16) + 1911, String) & "/" & AryDate(1) & "/" & AryDate(2)
                If IsDate(wkStr) Then
                    RtnDate = CDate(wkStr)
                End If
            End If
        Else
            RtnStr = InDate
            If RtnStr.Length > 4 Then
                wkStr = CType(CType(RtnStr.Substring(0, RtnStr.Length - 4), Int16) + 1911, String) & "/" & RtnStr.Substring(RtnStr.Length - 4, 2) & "/" & RtnStr.Substring(RtnStr.Length - 2, 2)
                If IsDate(wkStr) Then
                    RtnDate = CDate(wkStr)
                End If
            End If
        End If
        Return RtnDate
    End Function
#End Region

#Region "Format"

    Public Shared Function stripCommas(ByVal str As String) As String
        Dim strReg As String = System.Text.RegularExpressions.Regex.Replace(str, "[^\w\.@-]", "")
        Return strReg
    End Function

    Public Shared Function formatCoin(ByVal obj As Object) As String
        If IsNothing(obj) Then
            'Return "-"
            Return " "
        Else
            Dim str As String = CType(obj, String)
            If (str.Trim.Length > 0) AndAlso Not str.Equals("-") Then
                Try
                    str = stripCommas(str)
                    Return Format(Decimal.Parse(str), "#,##0")
                Catch ex As Exception
                    Return str
                End Try
            Else
                'Return "-"
                Return str
            End If
        End If
    End Function

    Public Shared Function formatCoin1(ByVal obj As Object) As String
        If IsNothing(obj) Then
            'Return "-"
            Return " "
        Else
            Dim str As String = CType(obj, String)
            If (str.Trim.Length > 0) AndAlso Not str.Equals("-") Then
                Try
                    str = stripCommas(str)
                    Return Format(Decimal.Parse(str), "#,##0.00")
                Catch ex As Exception
                    Return str
                End Try
            Else
                'Return "-"
                Return str
            End If
        End If
    End Function

    Public Shared Function formatCoin2(ByVal num As Object, Optional ByVal n As Integer = 2) As String
        Return Format(System.Math.Round(num, n), "#,##0.00")
    End Function

    Public Shared Function formatPercen(ByVal obj As Object) As String
        If IsNothing(obj) Then
            'Return "-%"
            Return " "
        Else
            Dim str As String = CType(obj, String)
            If (str.Trim.Length > 0) AndAlso Not str.Equals("-%") Then
                Try
                    str = stripCommas(str)
                    Return Format(Decimal.Parse(str), "#,##0.00") & "%"
                Catch ex As Exception
                    Return str
                End Try
            Else
                'Return "-"
                Return str
            End If
        End If
    End Function

    Public Shared Function converEmptyStr2Nothing(ByVal str As String) As String

        If Not IsNothing(str) AndAlso (str.Length = 0) Then
            Return Nothing
        Else
            Return str
        End If
    End Function

#End Region

#Region "Web Console Message"

    '在 Console顯示Exception
    Public Shared Sub Debug(ByVal ex As Exception)
        System.Diagnostics.Debug.WriteLine(ex)
    End Sub

    '在 Console顯示訊息
    Public Shared Sub Debug(ByVal message As String)
        System.Diagnostics.Debug.WriteLine(message)
    End Sub

    '在 Console顯示訊息
    Public Shared Sub Debug(ByVal obj As Object)
        System.Diagnostics.Debug.WriteLine(obj)
    End Sub

#End Region

#Region "String"

    Public Shared Function getSubString(ByVal sSourceString As String, ByVal sStartString As String, ByVal sEndString As String) As String
        Dim sn, en, nLen As Integer
        nLen = Len(sStartString)
        sn = InStr(sSourceString, sStartString)
        en = InStr(sn + nLen, sSourceString, sEndString)

        If sn = 0 Or en = 0 Then
            Return vbNullString
        Else
            Return Mid(sSourceString, sn + nLen, en - (sn + nLen))
        End If
    End Function
    '左補
    Public Shared Function padl(ByVal sSource As String, ByVal iLen As Integer, ByVal sPadding As String)
        Dim sData As New System.Text.StringBuilder(sSource)
        While (sData.Length() < iLen)
            sData.Insert(0, sPadding)
        End While
        Return sData.ToString()
    End Function

    '右補
    Public Shared Function padr(ByVal sSource As String, ByVal iLen As Integer, ByVal sPadding As String)
        Dim sData As New System.Text.StringBuilder(sSource)
        While (sData.Length() < iLen)
            sData.Append(sPadding)
        End While
        Return sData.ToString()
    End Function

    'ASCII轉Unicode
    Public Shared Function native2ascii(ByVal str As String) As String
        Dim code As Integer
        Dim sb As New System.Text.StringBuilder(255)
        Dim ch As CharEnumerator = str.GetEnumerator
        While ch.MoveNext
            Dim c As Char = ch.Current
            Dim iAsc As Integer = AscW(c) 'Unicode 字碼
            If iAsc > 255 Then
                sb.Append("\\u")
                code = (iAsc >> 8)
                Dim tmp As String = code.ToString("X")
                If tmp.Length = 1 Then
                    sb.Append("0")
                End If
                sb.Append(tmp)
                code = (iAsc And &HFF)
                tmp = code.ToString("X")
                If tmp.Length = 1 Then
                    sb.Append("0")
                End If
                sb.Append(tmp)
            Else
                sb.Append(c)
            End If
        End While
        'Dim code As Integer
        'Dim chars As Char() = str.ToCharArray
        'Dim sb As New System.Text.StringBuilder(255)
        'For i As Integer = 0 To chars.Length - 1
        '    Dim c As Char = chars(i)
        '    Dim iAsc As Integer = AscW(c) 'Unicode 字碼
        '    If iAsc > 255 Then
        '        sb.Append("\\u")
        '        code = (iAsc >> 8)
        '        Dim tmp As String = code.ToString("X")
        '        If tmp.Length = 1 Then
        '            sb.Append("0")
        '        End If
        '        sb.Append(tmp)
        '        code = (iAsc And &HFF)
        '        tmp = code.ToString("X")
        '        If tmp.Length = 1 Then
        '            sb.Append("0")
        '        End If
        '        sb.Append(tmp)
        '    Else
        '        sb.Append(c)
        '    End If
        'Next
        Return (sb.ToString())
    End Function

    'ASCII轉html Unicode
    Public Shared Function ascii2Unicode(ByVal str As String) As String
        'Dim code As Integer
        Dim chars As Char() = str.ToCharArray
        Dim sb As New System.Text.StringBuilder(255)
        For i As Integer = 0 To chars.Length - 1
            Dim c As Char = chars(i)
            Dim iAsc As Integer = AscW(c) 'Unicode 字碼             
            sb.Append("&#").Append(iAsc).Append(";")
        Next
        Return sb.ToString()
    End Function

    Public Shared Function string2ByteLength(ByVal str As String) As Integer
        Dim iSum As Integer = 0
        ' Console.WriteLine(System.Globalization.CultureInfo.CurrentCulture.TextInfo.ANSICodePage())
        Dim ch As CharEnumerator = str.GetEnumerator
        While ch.MoveNext
            Dim c As Char = ch.Current
            Dim iAsc As Integer = AscW(c) 'Unicode 字碼
            If iAsc >= 0 AndAlso iAsc <= 255 Then
                iSum += 1
            Else
                iSum += 2
            End If

            'Dim i As Integer = Asc(c)
            'If i <= 255 AndAlso i >= 0 Then  '值從 0 到 255 的單一位元組字元集
            '    iSum += 1
            '    'Console.WriteLine(c)
            'ElseIf i <= 32767 AndAlso i >= -32768 Then '值從 -32768 到 32767 的雙位元組字元集
            '    iSum += 2
            '    'Console.WriteLine(c)
            'End If
        End While
        Return iSum
    End Function
#End Region

#Region "ValidateData"

    Public Shared Function isValidateData(ByVal sInfo As String) As Boolean
        If (Not IsNothing(sInfo) AndAlso sInfo.Trim.Length > 0) Then
            Return True
        End If
        Return False
    End Function

    Public Shared Function isValidateData(ByVal obj As Object) As Boolean
        If IsDBNull(obj) Then
            Return False
        ElseIf IsNothing(obj) Then
            Return False
        ElseIf TypeOf obj Is String Then
            Dim s As String = CType(obj, String)
            Return isValidateData(s)
        ElseIf TypeOf obj Is Date Then
            Dim oDate As Date = CType(obj, Date)
            If oDate.Year < 1911 AndAlso oDate.Month < 1 AndAlso oDate.Month > 12 Then
                Return False
            End If
        ElseIf TypeOf obj Is Integer Then
            Return True
        ElseIf TypeOf obj Is Boolean Then
            Return obj
        Else
            Dim sObj As String = String.Empty
            sObj = obj.ToString
            If IsDate(sObj) Then
                Dim oDate As Date = CType(sObj, Date)
                If oDate.Year < 1911 AndAlso oDate.Month < 1 AndAlso oDate.Month > 12 Then
                    Return False
                Else
                    Return True
                End If
            End If
        End If
        Return True
    End Function


    Public Shared Function checkDBNull(ByVal dbObject As Object) As Object
        If IsDBNull(dbObject) Then
            Return Nothing
        Else
            Return dbObject
        End If
    End Function

    Public Shared Function DbObject2Str(ByVal dbObject As Object) As String
        dbObject = checkDBNull(dbObject)
        If isValidateData(dbObject) Then
            Dim str As String = CType(dbObject, String)
            str = StrConv(str.Trim, VbStrConv.Narrow)
            Return str
        End If
        Return ""
    End Function

    Public Shared Function isDecimalDate(ByVal MB_TX_DATE As Object) As Boolean
        Try
            If Utility.isValidateData(MB_TX_DATE) AndAlso MB_TX_DATE.ToString.Length = 8 Then
                Dim sMB_TX_DATE As String = MB_TX_DATE.ToString

                Dim sCHKDate As String = String.Empty

                sCHKDate = CDate(Left(sMB_TX_DATE, 4) & "/" & sMB_TX_DATE.Substring(4, 2) & "/" & Right(sMB_TX_DATE, 2))

                Return IsDate(sCHKDate)
            End If

            Return False
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

#Region "阿拉伯數字轉為大寫國字"

    Private Shared ReadOnly ar() As String = {"零", "壹", "貳", "參", "肆", "伍", "陸", "柒", "捌", "玖"}
    Private Shared ReadOnly cName() As String = {"", "", "拾", "佰", "仟", "萬", "拾", "佰", "仟", "億", "拾", "佰", "仟"}

    Private Shared ReadOnly cENName() As String = {"", "", "拾", "佰", "仟", "萬", "拾", "佰", "仟", "億", "拾", "佰", "仟", "兆", "拾", "佰", "仟"}

    Public Shared Function GetChineseNumber(ByVal number As Integer) As String
        Dim chineseNumber As String() = {"零", "一", "二", "三", "四", "五", "六", "七", "八", "九"}
        Dim unit As String() = {"", "十", "百", "千", "萬", "十萬", "百萬", "千萬", "億", "十億", "百億", "千億", "兆", "十兆", "百兆", "千兆"}
        Dim ret As New System.Text.StringBuilder()
        Dim inputNumber As String = number.ToString()
        Dim idx As Integer = inputNumber.Length
        Dim needAppendZero As Boolean = False
        For Each c As Char In inputNumber
            idx -= 1
            If Asc(c) > Asc("0") Then
                If needAppendZero Then
                    ret.Append(chineseNumber(0))
                    needAppendZero = False
                End If
                ret.Append(chineseNumber(Asc(c) - Asc("0")) & unit(idx))
            Else
                needAppendZero = True
            End If
        Next

        If ret.Length = 0 Then
            Return chineseNumber(0)
        Else
            Return ret.ToString()
        End If
    End Function

    Public Shared Function cypherConverCNum(ByVal sNum As String) As String
        Dim cNum As Integer
        Dim cunit As String
        Dim i As Integer
        Dim cZero As Integer
        Dim conver As String
        Dim cLast As String
        conver = ""
        cLast = ""
        cZero = 0
        i = 1
        For j As Integer = Len(sNum) To 1 Step -1
            cNum = Val(Mid(sNum, i, 1))
            cunit = cName(j) '取出位數 
            If cNum = 0 Then '判斷取出的數字是否為0,如果是0,則記錄共有幾0 
                cZero = cZero + 1
                If (InStr("萬億", cunit) > 0 And (cLast = "")) Then '如果取出的是萬,億,則位數以萬億來補 
                    cLast = cunit
                End If
            Else
                If cZero > 0 Then '如果取出的數字0有n個,則以零代替所有的0 
                    If InStr("萬億", Right(conver, 1)) = 0 Then
                        conver = conver + cLast '如果最後一位不是億,萬,則最後一位補上"億萬" 
                    End If
                    conver = conver + "零"
                    cZero = 0
                    cLast = ""
                End If
                conver = conver + ar(cNum) + cunit '如果取出的數字沒有0,則是中文數字+單位 
            End If
            i = i + 1
        Next
        '判斷數字的最後一位是否為0,如果最後一位為0,則把萬億補上 
        If InStr("萬億", Right(conver, 1)) = 0 Then
            conver = conver + cLast '如果最後一位不是億,萬,則最後一位補上"億萬" 
        End If

        'Return "新台幣" & conver & "元整"
        Return conver
    End Function

    'Add By Ted 20080701
    '阿拉伯數字轉國字大寫
    '只能到仟兆 超過仟兆 回傳1,9001,123,456,789,000
    Public Shared Function cypherConverCNumEN(ByVal sNum As String, Optional ByVal isCUT1000 As Boolean = True) As String
        Try
            If InStr(sNum, ".") > 0 Then
                sNum = sNum.Split(".")(0)
            End If

            If sNum.Length <= 16 Then
                Dim cNum As Integer
                Dim cunit As String
                Dim i As Integer
                Dim cZero As Integer
                Dim conver As String
                Dim cLast As String
                conver = ""
                cLast = ""
                cZero = 0
                i = 1

                '判斷是否取到元 否則取到千元
                Dim iMinCVT As Integer = 3
                If isCUT1000 = False Then
                    iMinCVT = 1
                End If

                For j As Integer = Len(sNum) To 3 Step -1
                    cNum = Val(Mid(sNum, i, 1))
                    cunit = cENName(j) '取出位數 
                    If cNum = 0 Then '判斷取出的數字是否為0,如果是0,則記錄共有幾0 
                        cZero = cZero + 1
                        If (InStr("萬億", cunit) > 0 And (cLast = "")) Then '如果取出的是萬,億,則位數以萬億來補 
                            cLast = cunit
                        End If
                    Else
                        If cZero > 0 Then '如果取出的數字0有n個,則以零代替所有的0 
                            If InStr("萬億", Right(conver, 1)) = 0 Then
                                conver = conver + cLast '如果最後一位不是億,萬,則最後一位補上"億萬" 
                            End If

                            'conver = conver + "零"

                            cZero = 0
                            cLast = ""
                        End If
                        conver = conver + ar(cNum) + cunit '如果取出的數字沒有0,則是中文數字+單位 
                    End If
                    i = i + 1
                Next
                '判斷數字的最後一位是否為0,如果最後一位為0,則把萬億補上 
                If InStr("萬億", Right(conver, 1)) = 0 Then
                    conver = conver + cLast '如果最後一位不是億,萬,則最後一位補上"億萬" 
                End If

                'Return "新台幣" & conver & "元整"
                Return conver
            Else
                Return FormatDec(sNum, "#,##0")
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

#Region "web"

    '<td width="33%" class="tactive" nowrap="nowrap">授信戶：07663290&nbsp;&nbsp;力麗企業股份有限公司</td>
    'return 授信戶：07663290力麗企業股份有限公司

    '<tr class="tactive">
    '<td width="33%" class="tactive" nowrap="nowrap">授信戶：07663290&nbsp;&nbsp;力麗企業股份有限公司</td>
    '<td width="34%" align="center" class="tactive" nowrap="nowrap">步驟代碼: 1400</td>
    '<td width="33%" align="right" class="tactive" nowrap="nowrap">核定階層： 董事會或常董會</td>
    '</tr>
    'return 授信戶：07663290力麗企業股份有限公司步驟代碼: 1400核定階層： 董事會或常董會
    Public Shared Function gFiltHtmlInString(ByVal rstrIn As String, Optional ByVal isSplit As Boolean = False) As String
        Try
            Dim sb As New System.Text.StringBuilder
            Dim sSplit As String = "^"
            Dim str As String
            Dim iStart As Integer = 0

            Do While iStart < rstrIn.Length
                str = rstrIn.Substring(iStart, 1)
                If str.Equals("<") Then
                    Dim i As Integer = rstrIn.Substring(iStart).IndexOf(">") + iStart
                    If (i > 0) Then
                        If isSplit AndAlso sb.Length > 0 And rstrIn.Substring(iStart, 2).Equals("</") Then
                            sb.Append(sSplit)
                        End If
                        iStart = i + 1
                    Else
                        sb.Append(str)
                        iStart += 1
                    End If
                Else
                    sb.Append(str)
                    iStart += 1
                End If
            Loop
            '歸位/換行字元 
            sb.Replace(Microsoft.VisualBasic.vbCrLf, "")
            '歸位字元。
            sb.Replace(Microsoft.VisualBasic.vbCr, "")
            '換行字元
            sb.Replace(Microsoft.VisualBasic.vbLf, "")
            '新行字元
            sb.Replace(Microsoft.VisualBasic.vbNewLine, "")
            '具有 0 值的字元。
            sb.Replace(Microsoft.VisualBasic.vbNullChar, "")
            '定位字元
            sb.Replace(Microsoft.VisualBasic.vbTab, "")
            '退格鍵 (Backspace)
            sb.Replace(Microsoft.VisualBasic.vbBack, "")
            sb.Replace("&nbsp;", " ")
            If isSplit Then
                Return sb.Remove(sb.ToString.Trim().Length() - 1, 1).ToString
            Else
                Return sb.ToString.Trim()
            End If

        Catch ex As Exception
            Throw ex
        End Try

    End Function
    '[取得所有URL參數 回傳URL字串]
    Public Shared Function getURLParas(ByVal request As System.Web.HttpRequest) As String
        Dim sbURL As New System.Text.StringBuilder
        Dim sCol As System.Collections.Specialized.NameValueCollection

        sCol = request.QueryString
        Try
            For Each sReqKey As String In sCol.AllKeys
                If Not IsNothing(sReqKey) Then
                    'Dim sReqValue As String = sCol.Item(sReqKey)
                    'sbURL.Append(sReqKey & "=" & sReqValue & "&")
                    Dim s As String = request.QueryString(sReqKey)
                    If Array.IndexOf(sCol.AllKeys, sReqKey) <> Array.LastIndexOf(sCol.AllKeys, sReqKey) Then
                        Dim sReqValue As String() = request.QueryString(sReqKey).Split(",")
                        s = sReqValue(sReqValue.Length - 1)
                    End If

                    If Utility.isValidateData(s) Then
                        sbURL.Append(sReqKey & "=" & s & "&")
                    End If
                End If
            Next
            If sbURL.Length > 0 Then
                Return sbURL.Remove(sbURL.ToString().Trim().Length() - 1, 1).ToString
            Else
                Return ""
            End If
        Catch ex As Exception
            Throw ex
        End Try
        Return ""
    End Function

    Public Shared Function getHtml(ByVal page As System.Web.UI.Page) As String
        Dim sw As System.IO.StringWriter = Nothing
        Dim tw As System.Web.UI.HtmlTextWriter = Nothing

        Try
            sw = New System.IO.StringWriter
            tw = New System.Web.UI.HtmlTextWriter(sw)
            page.RenderControl(tw)
            Return sw.ToString()
        Finally
            sw.Close()
            tw.Close()
        End Try
    End Function

    Public Shared Function getHtml(ByVal page As System.Web.UI.Page, ByVal sw As System.IO.StringWriter) As String
        Dim tw As System.Web.UI.HtmlTextWriter = Nothing
        Try

            tw = New System.Web.UI.HtmlTextWriter(sw)
            page.RenderControl(tw)
            Return sw.ToString()
        Finally
            sw.Close()
            tw.Close()
        End Try
    End Function

    Public Shared Sub downLoadFile(ByVal page As System.Web.UI.Page, ByVal sFileName As String, ByVal str As String)
        'Dim page As New System.Web.UI.Page

        page.Response.ClearHeaders()
        page.Response.Clear()
        page.Response.Expires = 0
        page.Response.Buffer = True

        page.Response.AddHeader("Accept-Language", "zh-tw")
        page.Response.ContentType = "application/octet-stream; charset=iso-8859-1"  '"Application/octet-stream"
        page.Response.AddHeader("content-disposition", "attachment; filename=" & Chr(34) & System.Web.HttpUtility.UrlEncode(sFileName, System.Text.Encoding.UTF8) & Chr(34))

        page.Response.Write(str)
        page.Response.End()
    End Sub

    Public Shared Function getWebPage(ByVal sUrl As String) As String
        Dim req As System.Net.HttpWebRequest
        Dim resp As System.Net.HttpWebResponse
        Dim stream As System.IO.Stream = Nothing
        Dim streamRead As System.IO.StreamReader = Nothing
        Try
            req = System.Net.WebRequest.Create(sUrl)
            resp = req.GetResponse()
            stream = resp.GetResponseStream()
            streamRead = New System.IO.StreamReader(stream, System.Text.Encoding.Default)
            Return streamRead.ReadToEnd
        Catch ex As Exception
            Throw ex
        Finally
            stream.Close()
            streamRead.Close()
        End Try
    End Function

    'Public Shared Function getWebPageX(ByVal sURL As String) As String
    '    Try
    '        '此範例需安裝  MSXML 4.0 Service Pack 2 (Microsoft XML Core Services) (msxmlcht.msi)
    '        Dim xmlhttp As New MSXML2.ServerXMLHTTP60
    '        xmlhttp.setTimeouts(30000, 30000, 30000, 30000)
    '        xmlhttp.open("POST", sURL)
    '        'add by ted 20070206
    '        '忽略所有憑證檢查
    '        xmlhttp.setOption(MSXML2.SERVERXMLHTTP_OPTION.SXH_OPTION_IGNORE_SERVER_SSL_CERT_ERROR_FLAGS, 13056)
    '        'add by ted End here
    '        xmlhttp.send()

    '        If xmlhttp.statusText = "OK" Then
    '            'add by ted 20070305
    '            '移除歸位/換行字元
    '            Dim sResult As String = ""
    '            sResult = xmlhttp.responseText
    '            Return sResult
    '        Else
    '            Return xmlhttp.statusText
    '        End If
    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Function
#End Region

#Region "字串組合"

    '[組成地址字串]
    'Public Shared Function getAddressComb(ByVal pType As String, ByVal pCity As String, _
    '         ByVal pVlg As String, ByVal pUcy_Sec As String, ByVal pNhr_SSub As String, _
    '         ByVal pRod_Lno As String, ByVal pTnl_LSub As String, _
    '         Optional ByVal pAly As String = "", Optional ByVal pNO As String = "", _
    '         Optional ByVal pFlr As String = "", Optional ByVal pFSub As String = "", _
    '         Optional ByVal pRom As String = "") As String
    '    Dim sStr As String
    '    Dim rtnValue As String

    '    If pCity <> "" AndAlso Trim(pCity) <> "" AndAlso Trim(pCity) <> "0" Then
    '        sStr &= pCity
    '    End If
    '    If pVlg <> "" AndAlso Trim(pVlg) <> "" AndAlso Trim(pVlg) <> "0" Then
    '        sStr &= pVlg
    '    End If

    '    Select Case pType
    '        Case "A"
    '            If pUcy_Sec <> "" AndAlso Trim(pUcy_Sec) <> "" AndAlso Trim(pUcy_Sec) <> "0" Then
    '                sStr &= pUcy_Sec & "段"
    '            End If
    '            If pNhr_SSub <> "" AndAlso Trim(pNhr_SSub) <> "" AndAlso Trim(pNhr_SSub) <> "0" Then
    '                sStr &= pNhr_SSub & "小段"
    '            End If
    '            If pRod_Lno <> "" AndAlso Trim(pRod_Lno) <> "" AndAlso Trim(pRod_Lno) <> "0" Then
    '                sStr &= "地號 " & pRod_Lno
    '            End If
    '            If pTnl_LSub <> "" AndAlso Trim(pTnl_LSub) <> "" AndAlso Trim(pTnl_LSub) <> "0" Then
    '                sStr &= " ─ " & pTnl_LSub
    '            End If
    '        Case "B"
    '            If pUcy_Sec <> "" AndAlso Trim(pUcy_Sec) <> "" AndAlso Trim(pUcy_Sec) <> "0" Then
    '                sStr &= pUcy_Sec & "里"
    '            End If
    '            If pNhr_SSub <> "" AndAlso Trim(pNhr_SSub) <> "" AndAlso Trim(pNhr_SSub) <> "0" Then
    '                sStr &= pNhr_SSub & "鄰"
    '            End If
    '            If pRod_Lno <> "" AndAlso Trim(pRod_Lno) <> "" AndAlso Trim(pRod_Lno) <> "0" Then
    '                sStr &= pRod_Lno & "路"
    '            End If
    '            If pTnl_LSub <> "" AndAlso Trim(pTnl_LSub) <> "" AndAlso Trim(pTnl_LSub) <> "0" Then
    '                sStr &= pTnl_LSub & "巷"
    '            End If
    '            If pAly <> "" AndAlso Trim(pAly) <> "" AndAlso Trim(pAly) <> "0" Then
    '                sStr &= pAly & "弄"
    '            End If
    '            If pNO <> "" AndAlso Trim(pNO) <> "" AndAlso Trim(pNO) <> "0" Then
    '                sStr &= pNO & "號"
    '            End If
    '            If pFlr <> "" AndAlso Trim(pFlr) <> "" AndAlso Trim(pFlr) <> "0" Then
    '                sStr &= pFlr & "樓"
    '            End If
    '            If pFSub <> "" AndAlso Trim(pFSub) <> "" AndAlso Trim(pFSub) <> "0" Then
    '                sStr &= "之" & pFSub
    '            End If
    '            If pRom <> "" AndAlso Trim(pRom) <> "" AndAlso Trim(pRom) <> "0" Then
    '                sStr &= pRom & "室"
    '            End If
    '    End Select

    '    rtnValue = sStr

    '    Return rtnValue

    'End Function

    ''[設定分子分母顯示格式]
    'Public Shared Function getDividFormat(ByVal pNMR As String, ByVal pDMR As String) As String
    '    Dim rtnValue As String

    '    rtnValue = pNMR & "/" & pDMR

    '    Return rtnValue

    'End Function

    '[補足0位函式 參數(變數,長度)]
    Public Shared Function FillZero(ByVal sBeAdd As String, ByVal iLength As Integer) As String
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

    '[去除字串中/、.、-或自訂去除的符號，pValuep--輸入字串，pClsChar--欲替換的字元不輸入依預設的替換符號，pChgChar--替換後的字元不輸入預設為""]
    Public Shared Function chgString(ByVal pValue As Object, Optional ByVal pClsChar As String = "", Optional ByVal pChgChar As String = "") As String
        Dim sValue As String
        Dim rtnValue As String

        sValue = CStr(pValue)

        If pClsChar <> "" Then
            If pChgChar <> "" Then
                rtnValue = Replace(sValue, pClsChar, pChgChar)
            Else
                rtnValue = Replace(sValue, pClsChar, "")
            End If
        Else
            sValue = Replace(sValue, "/", "")
            sValue = Replace(sValue, ".", "")
            sValue = Replace(sValue, "-", "")
            rtnValue = sValue
        End If

        Return rtnValue
    End Function

    ''[處理TextBox單引雙引相關資料]
    'Public Shared Function Replace_HTML(ByVal rString)
    '    rString = rString.Trim().replace("'", "''")
    '    Replace_HTML = rString
    'End Function

    '由統編第2碼判斷先生或女士 1=先生 2=女士
    Public Shared Function getSEXByIDBAN(ByVal sIDBAN As String) As String
        Try
            If Not IsNothing(sIDBAN) AndAlso sIDBAN.Trim.Length <> 8 AndAlso _
                InStr(UCase(sIDBAN), "OBU") = 0 AndAlso sIDBAN.Trim.Length >= 2 Then

                Dim sTemp As String = ""
                sTemp = sIDBAN.Trim.Substring(1, 1)

                Select Case sTemp
                    Case "1"
                        Return "先生"
                    Case "2"
                        Return "女士"
                End Select
            End If

            Return ""
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Shared Function MapArabicToTW(ByVal iIndex As Integer) As String
        Dim sTWNUMBER As String = "零一二三四五六七八九"
        Try
            If iIndex <= 9 Then
                Return sTWNUMBER.Substring(iIndex, 1)
            End If

            Return ""
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    '取得打勾符號
    Public Shared Function getMarkAssign() As String
        Try
            Dim sRTN As String = ""
            sRTN = "<span lang=EN-US style='font-family:""Wingdings 2"";mso-ascii-font-family:""Times New Roman"";mso-hansi-font-family:""Times New Roman"";mso-char-type:symbol;mso-symbol-font-family:""Wingdings 2""' > " & _
                   "<span style='mso-char-type:symbol;mso-symbol-font-family:""Wingdings 2""' >R</span>" & _
                   "</span>"
            Return sRTN
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    '取得空格符號
    Public Shared Function getBlankAssign() As String
        Try
            Dim sRTN As String = ""
            sRTN = "<span lang=EN-US style='font-family:""Wingdings 2"";mso-ascii-font-family:""Times New Roman"";mso-hansi-font-family:""Times New Roman"";mso-char-type:symbol;mso-symbol-font-family:""Wingdings 2""' > " & _
                   "<span style='mso-char-type:symbol;mso-symbol-font-family:""Wingdings 2""' >&pound;</span>" & _
                   "</span>"
            Return sRTN
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    '取代&符號供XML Tag使用
    Public Shared Function replaceAndForXML(ByVal sXML As String) As String
        Try
            If Utility.isValidateData(sXML) Then
                Dim sRTN As String = ""
                sRTN = Replace(sXML, "&", "&#38;")

                Return sRTN
            End If

            Return ""
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

#Region "証號辨識"

    '[身份證字號驗證]
    Public Shared Function ChPersonalID1(ByVal sEId As String) As String
        Dim Letter As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        Dim MirrorNumber() = {10, 11, 12, 13, 14, 15, 16, 17, 34, 18, 19, 20, 21, 22, 35, 23, 24, 25, 26, 27, 28, 29, 32, 30, 31, 33}
        Dim Checker() = {1, 9, 8, 7, 6, 5, 4, 3, 2, 1, 1}
        Dim N1 As String = UCase(Left(sEId, 1))
        Dim I, N1Ten, N1Unit, Total As Integer
        Dim Pos As Byte = InStr(Letter, N1)
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
        N1Ten = Mid(MirrorNumber(Pos - 1), 1, 1)
        N1Unit = Mid(MirrorNumber(Pos - 1), 2, 1)
        Total = N1Ten * Checker(0) + N1Unit * Checker(1)
        For I = 2 To 10
            Total = Total + Mid(sEId, I, 1) * Checker(I)
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
            N1 = Mid(sEId, I, 1) * Checker(I - 1)
            Total += N1 Mod 10 + N1 \ 10
        Next

        If Total Mod 10 <> 0 Then
            '(joyce2005_4_26)---------------------------------------------------------------
            If Mid(sEId, 7, 1) = "7" Then                       '例外情況(第7位為7時)
                Dim sum_D7, a4, b4, sum2_D7, a5, b5 As Integer

                sum_D7 = CInt(Mid(sEId, 7, 1) * 4)
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

    '[帳號檢核]

    '帳號檢核規則
    '帳號:  9  8  3  0  0  1  0  6  0  3  3  4
    '＊  6  5  4  3  2   7  6  5  4  3  2
    '9*6+8*5+3*4+0*3+0*3+1*7+0*6+6*5+0*4+3*3+3*2 = 158

    '158/9 →商數= 17       餘數=5     檢查碼=9-5→ 4
    Public Shared Function isValidAccount(ByVal sAccount As String) As Boolean
        isValidAccount = False
        Dim myAL As New ArrayList
        Try
            If Not IsNothing(sAccount) AndAlso sAccount.Length = 12 Then
                Dim iFormula() As Integer = {6, 5, 4, 3, 2, 7, 6, 5, 4, 3, 2}
                Dim iSum As Integer = 0
                For i As Integer = 0 To sAccount.Length - 2
                    Dim sTarget As String = ""
                    sTarget = sAccount.Substring(i, 1)
                    If IsNumeric(sTarget) Then
                        iSum += CInt(sTarget) * iFormula(i)
                    End If
                Next

                Dim iMod As Integer = 0
                iMod = iSum Mod 9

                Dim sResult As String = ""
                sResult = sAccount.Substring(11, 1)
                If IsNumeric(sResult) AndAlso (9 - iMod) = CInt(sResult) Then
                    Return True
                End If
            End If
        Catch ex As Exception
            Throw ex
        Finally
            If Not myAL Is Nothing Then myAL = Nothing
        End Try
    End Function

#End Region

#Region "數學運算"

    '[取得輸入值的總和]
    Public Shared Function calTotal(ByVal pS1 As String, ByVal pS2 As String, _
                    Optional ByVal pS3 As String = "", Optional ByVal pS4 As String = "", _
                    Optional ByVal pS5 As String = "") As String
        Dim iTotal As Decimal
        Dim rtnValue As String

        iTotal = 0

        If pS1 <> "" AndAlso Trim(pS1) <> "" Then
            iTotal += CDec(pS1)
        End If
        If pS2 <> "" AndAlso Trim(pS2) <> "" Then
            iTotal += CDec(pS2)
        End If
        If pS3 <> "" AndAlso Trim(pS3) <> "" Then
            iTotal += CDec(pS3)
        End If
        If pS4 <> "" AndAlso Trim(pS4) <> "" Then
            iTotal += CDec(pS4)
        End If
        If pS5 <> "" AndAlso Trim(pS5) <> "" Then
            iTotal += CDec(pS5)
        End If

        rtnValue = FormatDec(CStr(iTotal), "###,##0.00")

        Return rtnValue
    End Function

    '[取得輸入值的乘積]
    Public Shared Function getMultiValue(ByVal pS1 As String, ByVal pS2 As String, _
                    Optional ByVal pS3 As String = "", Optional ByVal pS4 As String = "") As String
        Dim iSummary As Decimal
        Dim rtnValue As String

        If pS1 <> "" AndAlso Trim(pS1) <> "" AndAlso pS2 <> "" AndAlso Trim(pS2) <> "" Then
            iSummary = CDec(pS1) * CDec(pS2)
        Else
            iSummary = 0
        End If
        If pS3 <> "" AndAlso Trim(pS3) <> "" Then
            iSummary = iSummary * CDec(pS3)
        End If
        If pS4 <> "" AndAlso Trim(pS4) <> "" Then
            iSummary = iSummary * CDec(pS4)
        End If

        rtnValue = FormatDec(CStr(iSummary), "###,##0.00")

        Return rtnValue
    End Function

    '[取得輸入值的百分比]
    Public Shared Function getDicountPercent(ByVal pS1 As String, ByVal pS2 As String) As String
        Dim iDivis As Decimal
        Dim rtnValue As String

        If pS1 <> "" AndAlso Trim(pS1) <> "" AndAlso pS2 <> "" AndAlso Trim(pS2) <> "" Then
            iDivis = CDec(pS1) / CDec(pS2)
        Else
            iDivis = 0
        End If
        rtnValue = FormatDec(CStr(iDivis * 100), "##0.00") & "%"

        Return rtnValue

    End Function

    '[取得輸入值的成數]
    Public Shared Function getDivisTen(ByVal pS1 As String, ByVal pS2 As String) As String
        Dim iDivis As Decimal
        Dim rtnValue As String

        If pS1 <> "" AndAlso Trim(pS1) <> "" AndAlso pS2 <> "" AndAlso Trim(pS2) <> "" Then
            iDivis = CDec(pS1) / CDec(pS2)
        Else
            iDivis = 0
        End If
        rtnValue = FormatDec(CStr(iDivis * 100), "##0.0")

        Return rtnValue
    End Function

    '[取得輸入值的格式]
    Public Shared Function FormatDec(ByVal InDecimal As Object, ByVal strFormat As String) As String
        Dim RtnStr As String = ""
        If TypeName(InDecimal) <> "Nothing" AndAlso IsNumeric(InDecimal) Then
            RtnStr = Format(CDec(InDecimal), strFormat)
        End If
        Return RtnStr
    End Function

    '[小數點兩位格式]
    Public Shared Function getFloat2Format(ByVal pDollar As String) As String
        Dim dValue As Decimal
        Dim rtnvalue As String

        If Not IsNothing(pDollar) AndAlso Trim(pDollar) <> "" Then
            dValue = CDec(pDollar)
            rtnvalue = Format(dValue, "#,##0.00")
        Else
            rtnvalue = " "
        End If

        Return rtnvalue

    End Function

#End Region

#Region "ENBASE FUNCTION"
    
    Public Shared Function CheckNull(ByVal CheckValue) As String
        If IsDBNull(CheckValue) Then
            Return ""
        Else
            Return CheckValue
        End If
    End Function
     

    Public Shared Function CheckNumNull(ByVal CheckValue) As Object
        If IsNumeric(CheckValue) Then
            Return CDec(CheckValue)
        Else
            Return 0
        End If
    End Function
      
    '[轉換日期函式-西曆轉中曆 YYYY/MM/DD->YYY.MM.DD]
    Public Shared Function DateTransfer(ByVal pDate As Object) As String
        'MYSQL的Date型態和.NET不相容，所以先轉String
        Dim sDate As String = String.Empty
        If Utility.isValidateData(pDate) Then
            sDate = pDate.ToString
        End If
        If Not IsDate(sDate) Then
            Return ""
        End If
        Dim InDate As Date = CType(sDate, Date)
        Dim RtnStr As String = ""
        Dim Cyear As String = ""
        Cyear = CStr(CInt(Year(InDate)) - 1911)
        'modify by ted 20061128 0951128 --> 95.11.28
        'RtnStr = FillZero(Cyear, 3) & FillZero(CStr(Month(InDate)), 2) & FillZero(CStr(Day(InDate)), 2)
        If Cyear > 0 Then RtnStr = Cyear & "." & FillZero(CStr(Month(InDate)), 2) & "." & FillZero(CStr(Day(InDate)), 2)
        Return RtnStr
    End Function

    '[轉換日期函式-西曆轉中曆 YYYY/MM/DD->YYYMMDD]
    Public Shared Function ToYYYMMDD(ByVal pDate As Object) As String
        'MYSQL的Date型態和.NET不相容，所以先轉String
        Dim sDate As String = String.Empty
        If Utility.isValidateData(pDate) Then
            sDate = pDate.ToString
        End If
        If Not IsDate(sDate) Then
            Return ""
        End If
        Dim InDate As Date = CType(sDate, Date)
        Dim RtnStr As String = ""
        Dim Cyear As String = ""
        Cyear = CStr(CInt(Year(InDate)) - 1911)
        ' RtnStr = FillZero(Cyear,3) & FillZero(Cstr(Month(InDate)),2) & FillZero(Cstr(Day(InDate)),2)
        RtnStr = FillZero(Cyear, 3) & FillZero(CStr(Month(InDate)), 2) & FillZero(CStr(Day(InDate)), 2)
        Return RtnStr
    End Function

#End Region

#Region "ENCODING FUNCTIONs "

#Region " transferNCRToUnicode [參數]:[source][來源字串] "
    'Numeric Character Reference(NCR)(html Unicode)  轉  Unicode 
    Public Shared Function transferNCRToUnicode(ByVal source As String) As String
        Dim resultBuffer As System.Text.StringBuilder = New System.Text.StringBuilder
        Dim entity As String = String.Empty
        Dim i As Integer = 0
        Dim semi As Integer = 0
        Dim e2i As System.Collections.Hashtable = New System.Collections.Hashtable
        Dim MAXUNICODE As Integer = &HFFFF

bcontinue: While (i < source.Length())
            Dim ch As Char = source.Chars(i)
            'UP LOGIC > After "&" must have "#" Then Run Condition
            If (ch = "&") AndAlso (i + 1) < source.Length() AndAlso source.Chars(i + 1) = "#" Then
                semi = source.IndexOf(";", i + 1)
                If (semi = -1) Then
                    resultBuffer.Append(ch)
                    i += 1
                    GoTo bcontinue
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
                    GoTo bcontinue
                Else
                    i = semi
                    Dim s As String = ChrW(entity.Substring(1)).ToString
                    'Dim a As String = Hex(entity.Substring(1))
                    'Dim abyte0() As Byte = digits2Bytes(a, 16)
                    'Dim s As String = byteToString(abyte0, javaname([UNICODE]))

                    'System.Text.Encoding.GetEncoding(javaname([UNICODE])).GetChars(System.Text.Encoding.GetEncoding(javaname([UNICODE])).GetBytes(s))
                    Try
                        resultBuffer.Append(System.Text.Encoding.GetEncoding("Unicode").GetChars(System.Text.Encoding.GetEncoding("Unicode").GetBytes(s)))
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
#End Region 'RETURN [STRING]

#Region " getUrlEncode "
    Public Shared Function getUrlEncode(ByVal str As String) As String
        Dim sRtn As String = ""
        Try
            If Utility.isValidateData(str) Then
                Dim i1 As Integer = str.Split("&").Length - 1
                Dim j1 As Integer = 0
                For Each s1 As String In str.Split("&")
                    Dim i2 As Integer = s1.Split("=").Length - 1
                    Dim j2 As Integer = 0
                    For Each s2 As String In s1.Split("=")
                        sRtn &= System.Web.HttpUtility.UrlEncode(s2)
                        If j2 < i2 Then sRtn &= "="
                        j2 = +1
                    Next
                    If j1 < i1 Then sRtn &= "&"
                    j1 = +1
                Next
            End If
            Return sRtn
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

#End Region

#Region "執行外部程式(sqlldr)"
    Public Shared Function Shell(ByVal sExeFile As String, ByVal sArgument As String) As String

        Dim pShell As System.Diagnostics.Process

        pShell = New System.Diagnostics.Process

        '設定執行檔及參數
        pShell.StartInfo.FileName = sExeFile
        pShell.StartInfo.Arguments = sArgument

        '必須要設定以下兩個屬性才可將輸出結果導向
        pShell.StartInfo.ErrorDialog = True
        pShell.StartInfo.UseShellExecute = False
        pShell.StartInfo.RedirectStandardOutput = True

        '是否要在新視窗中啟動處理序。不顯示任何視窗
        pShell.StartInfo.CreateNoWindow = True

        '開始執行
        pShell.Start()

        '將StdOUT的結果轉為字串, 其中StandardOutput屬性類別為StreamReader
        Shell = pShell.StandardOutput.ReadToEnd()

        pShell.WaitForExit()
    End Function
#End Region

    '[頁面停駐特定點使用] add by JoviKuo
    Public Shared Sub setObjFocus(ByVal sClientid As String, ByVal page As System.Web.UI.Page)
        Dim script As String = ""

        script = "<script language='javascript'>"
        script &= "if (document.all('" & sClientid & "')!=undefined){"
        script &= "document.all('" & sClientid & "').scrollIntoView();"
        script &= "}"
        script &= "</script>"

        page.RegisterStartupScript("setFocus", script)
    End Sub


End Class
