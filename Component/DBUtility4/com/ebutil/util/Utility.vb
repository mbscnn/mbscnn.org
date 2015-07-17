Imports ICSharpCode.SharpZipLib.Zip.Compression
Imports System.Reflection
Namespace Titan

    Public Class Utility

#Region "Date"

        '���o����~��� return ����95�~04��20��
        Public Shared Function getROCYYMMDD() As String
            Dim time As DateTime = Now()
            Dim iY As Int16 = Int16.Parse(time.Year.ToString)
            Dim sYear As String = (iY - 1911).ToString '����~
            Dim sMonth As String = time.Month.ToString
            Dim sDay As String = time.Day.ToString
            If (sMonth.Length = 1) Then
                sMonth = "0" & sMonth
            End If
            If (sDay.Length = 1) Then
                sDay = "0" & sDay
            End If
            Return "����" & (sYear) & "�~" & (sMonth) & "��" & (sDay) & "��"
        End Function

        '���o����~��� 
        'return 95�~04��20��
        Public Shared Function getROCYYMMDD1() As String
            Dim time As DateTime = Now()
            Dim iY As Int16 = Int16.Parse(time.Year.ToString)
            Dim sYear As String = (iY - 1911).ToString '����~
            Dim sMonth As String = time.Month.ToString
            Dim sDay As String = time.Day.ToString
            If (sMonth.Length = 1) Then
                sMonth = "0" & sMonth
            End If
            If (sDay.Length = 1) Then
                sDay = "0" & sDay
            End If
            Return (sYear) & "�~" & (sMonth) & "��" & (sDay) & "��"
        End Function

        '���o�{�b����~��� 
        'return 95�~04��20��,95�~04�멳
        Public Shared Function getROCYYMMDD2() As String
            Dim time As DateTime = Now()
            Dim iY As Int16 = Int16.Parse(time.Year.ToString)
            Dim sYear As String = (iY - 1911).ToString '����~
            Dim sMonth As String = time.Month.ToString
            Dim sDay As String = time.Day.ToString
            If (sMonth.Length = 1) Then
                sMonth = "0" & sMonth
            End If
            If (sDay.Length = 1) Then
                sDay = "0" & sDay
            End If
            Return (sYear) & "�~" & (sMonth) & "��" & (sDay) & "��," & (sYear) & "�~" & (sMonth) & "�멳"
        End Function

        '���o����~��(�o�Ӥ몺�̫�@��)�� 
        'input 200601 or 200602 or ....
        'return 95�~01��31�� or 95�~02��28�� or 95�~03��31�� or ....
        Public Shared Function getROCYYMMDD3(ByVal str As String) As String
            str = str.Trim
            If (str.Length <> 6) Then
                Throw New Exception("����榡�����T [" & str & "]")
            End If
            str = str.Insert(4, "/") & "/01"
            Dim time As DateTime = Convert.ToDateTime(str)
            Dim iY As Int16 = Int16.Parse(time.Year.ToString)
            Dim sYear As String = (iY - 1911).ToString '����~
            Dim sMonth As String = time.Month.ToString
            Dim sDay As String = Date.DaysInMonth(iY, time.Month)
            If (sMonth.Length = 1) Then
                sMonth = "0" & sMonth
            End If
            If (sDay.Length = 1) Then
                sDay = "0" & sDay
            End If
            Return (sYear) & "�~" & (sMonth) & "��" & (sDay) & "��"
        End Function

        '���o����~��(�o�Ӥ몺�̫�@��)�� return 95�~01��31��,95�~01�멳
        Public Shared Function getROCYYMMDD4(ByVal str As String) As String
            str = str.Trim
            If (str.Length <> 6) Then
                Throw New Exception("����榡�����T [" & str & "]")
            End If
            str = str.Insert(4, "/") & "/01"
            Dim time As DateTime = Convert.ToDateTime(str)
            Dim iY As Int16 = Int16.Parse(time.Year.ToString)
            Dim sYear As String = (iY - 1911).ToString '����~
            Dim sMonth As String = time.Month.ToString
            Dim sDay As String = Date.DaysInMonth(iY, time.Month)
            If (sMonth.Length = 1) Then
                sMonth = "0" & sMonth
            End If
            If (sDay.Length = 1) Then
                sDay = "0" & sDay
            End If
            Return (sYear) & "�~" & (sMonth) & "��" & (sDay) & "��," & (sYear) & "�~" & (sMonth) & "�멳"
        End Function

        '���o����~��� 
        'input str="2006/4/19"
        'return 95�~04��19��
        Public Shared Function getROCYYMMDD5(ByVal str As String) As String
            Dim time As DateTime = Convert.ToDateTime(str)
            Dim iY As Int16 = Int16.Parse(time.Year.ToString)
            Dim sYear As String = (iY - 1911).ToString '����~
            Dim sMonth As String = time.Month.ToString
            Dim sDay As String = time.Day
            If (sMonth.Length = 1) Then
                sMonth = "0" & sMonth
            End If
            If (sDay.Length = 1) Then
                sDay = "0" & sDay
            End If
            Return (sYear) & "�~" & (sMonth) & "��" & (sDay) & "��"
        End Function

        '���o����~��� 
        'input str="2006/4/19"
        'return 95.04.19
        Public Shared Function getROCYYMMDD6(ByVal str As String) As String

            If IsNothing(str) OrElse str.Length = 0 Then
                Return ""
            End If

            Dim time As DateTime = Convert.ToDateTime(str)
            Dim iY As Int16 = Int16.Parse(time.Year.ToString)
            Dim sYear As String = (iY - 1911).ToString '����~
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

        '���o����~ return ����95�~
        Public Shared Function getROCYY() As String
            Dim time As DateTime = Now()
            Dim iY As Int16 = Int16.Parse(time.Year.ToString)
            Dim sYear As String = (iY - 1911).ToString '����~
            Return "����" & (sYear) & "�~"
        End Function

        '���o����~ 
        'input 200602 
        'return 95
        Public Shared Function getROCYY(ByVal str As String) As String
            str = str.Trim
            If (str.Length <> 6) Then
                Throw New Exception("����榡�����T [" & str & "]")
            End If
            str = str.Insert(4, "/") & "/01"
            Dim time As DateTime = Convert.ToDateTime(str)
            Dim iY As Int16 = Int16.Parse(time.Year.ToString)
            Dim sYear As String = (iY - 1911).ToString '����~
            Dim sMonth As String = time.Month.ToString
            If (sMonth.Length = 1) Then
                sMonth = "0" & sMonth
            End If
            Return (sYear)

        End Function

        '���o����~�� return 95�~04��
        Public Shared Function getROCYYMM() As String
            Dim time As DateTime = Now()
            Dim iY As Int16 = Int16.Parse(time.Year.ToString)
            Dim sYear As String = (iY - 1911).ToString '����~
            Dim sMonth As String = time.Month.ToString
            If (sMonth.Length = 1) Then
                sMonth = "0" & sMonth
            End If
            Return (sYear) & "�~" & (sMonth) & "��"
        End Function

        '���o����~�� return 95�~04��
        Public Shared Function getROCYYMM(ByVal str As String) As String
            str = str.Trim
            If (str.Length <> 6) Then
                Throw New Exception("����榡�����T [" & str & "]")
            End If
            str = str.Insert(4, "/") & "/01"
            Dim time As DateTime = Convert.ToDateTime(str)
            Dim iY As Int16 = Int16.Parse(time.Year.ToString)
            Dim sYear As String = (iY - 1911).ToString '����~
            Dim sMonth As String = time.Month.ToString
            If (sMonth.Length = 1) Then
                sMonth = "0" & sMonth
            End If
            Return (sYear) & "�~" & (sMonth) & "��"
        End Function

        '���o   ����~�� ���u�̫�@��
        'input  199901 or 1999902 or 1999903
        'return ����88�~03��
        Public Shared Function getROCQuarterYYMM(ByVal str As String) As String

            Dim strQua As String = getQuarterYYMM(str).Split(",")(0) '1999903
            Return getROCYYMM(strQua)
        End Function

        '�褸�[��~��� return 1999/04/03
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

        '���o  �褸�~ return 1999/04/06
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

        '���o  �褸 �~��,�~(-1)�� return 199904,199903 
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

        '���o  �褸 �~��,�~(-1)�� 
        'input 199904
        'return 199904,199903
        Public Shared Function getYYdecMM1(ByVal str As String) As String
            str = str.Trim
            Dim str1 As String
            If (str.Length <> 6) Then
                Throw New Exception("����榡�����T [" & str & "]")
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

        '���o  �褸 �~��,(-1)�~�� 
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

        '���o  �褸 �~��,(-1)�~�� 
        'input 200601
        'return 200601,200501
        Public Shared Function getdecYYMM2(ByVal str As String) As String
            str = str.Trim
            If (str.Length <> 6) Then
                Throw New Exception("����榡�����T [" & str & "]")
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

        '���o  �褸 �~��,�~(-1)�� ,(-1)�~�� 
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

        '���o  �褸�~��
        'input 9501 ����~��
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
                Throw New Exception("(getROC2ADYYMM()):����榡�����T [" & str & "]")
            End If
            Return str
        End Function

        '���o  �褸�~���
        'input 0950112
        'return 2006/01/01
        Public Shared Function getROC2ADYYMMDD(ByVal str As String) As String

            If IsNothing(str) OrElse str.Length <> 7 Then
                Return str
            End If
            Dim sY As String = getROC2ADYYYY(str.Substring(0, 3))
            Dim sM As String = str.Substring(3, 2)
            Dim sD As String = str.Substring(5, 2)
            Return sY & "/" & sM & "/" & sD
        End Function

        '���o  ����~��
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
                Throw New Exception("(getROC2ADYMD()):����榡�����T [" & str & "]")
            End Try

            Return str
        End Function

        '���o  ����~��
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
                Throw New Exception("(getROC2ADYMD()):����榡�����T [" & str & "]")
            End Try

            Return str
        End Function

        '���o  �褸�~
        'input 95 ����~
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
                Throw New Exception("(getROC2ADYYYY()):����榡�����T [" & str & "]")
            End If
            Return str
        End Function

        '���o�ثe  �褸 �~�� return 199904
        Public Shared Function getYYMM() As String
            Dim time1 As DateTime = Now()
            Dim sYear1 As String = time1.Year.ToString
            Dim sMonth1 As String = time1.Month.ToString

            If (sMonth1.Length = 1) Then
                sMonth1 = "0" & sMonth1
            End If
            Return sYear1 & sMonth1
        End Function

        '�H�ثe���ɶ� ���o  �褸�~�� ���u����,�e�@�u���� 
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

        '���o   �褸�~�� ���u,�e�@�u 
        'input  199901 or 199902 or 199903
        'return         199903,1998012  
        Public Shared Function getQuarterYYMM(ByVal str As String) As String
            str = str.Trim
            If (str.Length <> 6) Then
                Throw New Exception("����榡�����T [" & str & "]")
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

        '���o  �褸�~�� ���u���T�U��� return 200603,200602,200601 or 200606,200605,200604(�Ѥj��-->�p��)
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
            Dim time2 As DateTime = time1.AddMonths(-3) '�e�@�u
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

        '���o  �褸�~�� ���u���T�U��� return 200603,200602,200601 or 200606,200605,200604
        Public Shared Function getQuarterYYMM2(ByVal str As String) As String
            Dim sb As New System.Text.StringBuilder
            str = str.Trim
            If (str.Length <> 6) Then
                Throw New Exception("����榡�����T [" & str & "]")
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
            Dim time2 As DateTime = time1.AddMonths(-3) '�e�@�u
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
                Throw New Exception("����榡�����T [" & str & "]")
            End If
            str = str.Insert(4, "/") & "/01" '�զ�����榡2000/03/01
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
                '���h�~���̫�@�u,���T�U���
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


            str = str.Insert(4, "/") & "/01" '�զ�����榡2000/03/01
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

        '���o�[��~���  �褸�~,����~ return 1999/04/20,88/04/20
        Public Shared Function getYMD(ByVal iYear As Int16, ByVal iMonth As Int16, ByVal iDay As Int16) As String
            Dim time As DateTime = Now().AddYears(iYear).AddMonths(iMonth).AddDays(iDay)
            Dim sCEYear As String = time.Year.ToString '�褸�~
            Dim iY As Int16 = Int16.Parse(time.Year.ToString)
            Dim sYear As String = (iY - 1911).ToString '����~
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

        '���o�[��~���  �褸�~��,����~�� return 1999/04,88/04
        Public Shared Function getYM(ByVal iYear As Int16, ByVal iMonth As Int16) As String
            Dim time As DateTime = Now().AddYears(iYear).AddMonths(iMonth)
            Dim sCEYear As String = time.Year.ToString '�褸�~
            Dim iY As Int16 = Int16.Parse(time.Year.ToString)
            Dim sYear As String = (iY - 1911).ToString '����~
            Dim sMonth As String = time.Month.ToString
            If (sMonth.Length = 1) Then
                sMonth = "0" & sMonth
            End If
            Return sCEYear & "/" & sMonth & "," & sYear & "/" & sMonth
        End Function

        '���o�[��~���  �褸�~��,����~�� return 199904,88/04
        Public Shared Function getYM1(ByVal iYear As Int16, ByVal iMonth As Int16) As String
            Dim time As DateTime = Now().AddYears(iYear).AddMonths(iMonth)
            Dim sCEYear As String = time.Year.ToString '�褸�~
            Dim iY As Int16 = Int16.Parse(time.Year.ToString)
            Dim sYear As String = (iY - 1911).ToString '����~
            Dim sMonth As String = time.Month.ToString
            If (sMonth.Length = 1) Then
                sMonth = "0" & sMonth
            End If
            Return sCEYear & sMonth & "," & sYear & "/" & sMonth
        End Function

        '���o  �褸�~,����~ return 1999/4/20,88/4/20
        Public Shared Function getYMD() As String
            Dim time As DateTime = Now()
            Dim iY As Int16 = Int16.Parse(time.Year.ToString)
            Dim sYear As String = (iY - 1911).ToString '����~
            Return (time.Year.ToString) & "/" & (time.Month.ToString) & "/" & (time.Day.ToString) & "," & (sYear) & "/" & (time.Month.ToString) & "/" & (time.Day.ToString)
        End Function

        '���o  �褸�~ return 1999
        Public Shared Function getYY()
            Return (Now().Year.ToString)
        End Function

        Public Shared Function getMM()
            Return (Now().Month.ToString)
        End Function

        Public Shared Function getDD()
            Return (Now().Day.ToString)
        End Function


        Public Shared Function showTWDate(ByVal data As Object, ByVal sFormat As String, ByVal sSplit As String) As String

            If Not IsNothing(data) Then
                Dim twCalendar As New System.Globalization.TaiwanCalendar
                Dim dateTime As New dateTime(data.Year, data.Month, data.Day, New System.Globalization.GregorianCalendar)

                Dim sY As String = twCalendar.GetYear(dateTime)
                If sY < 100 Then
                    sY = "0" + sY
                End If
                Dim sM As String = twCalendar.GetMonth(dateTime)
                If sM < 10 Then
                    sM = "0" + sM
                End If

                Dim sD As String = twCalendar.GetDayOfMonth(dateTime)
                If sD < 10 Then
                    sD = "0" + sD
                End If

                If (sFormat.ToUpper.Equals("YYMM")) Then
                    Return sY & sSplit & sM
                ElseIf (sFormat.ToUpper.Equals("YYMMDD")) Then
                    Return sY & sSplit & sM & sSplit & sD
                Else
                    Return sY & sM & sD
                End If
            Else
                Return ""
            End If
        End Function

        Public Shared Function showGregorianDate(ByVal data As Object, ByVal sFormat As String, ByVal sSplit As String) As String

            If Not IsNothing(data) Then
                Dim greCalendar As New System.Globalization.GregorianCalendar
                Dim dateTime As New dateTime(data.Year, data.Month, data.Day, New System.Globalization.GregorianCalendar)
                Dim sY As String = greCalendar.GetYear(dateTime)
                Dim sM As String = greCalendar.GetMonth(dateTime)
                Dim sD As String = greCalendar.GetDayOfMonth(dateTime)

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
            If IsNothing(data) Then
                Return ""
            End If
            Return showTWDate(data, "YYMMDD", ".")
        End Function

        '[���o����~]
        Public Shared Function DateTransferYYYY(ByVal pDate As Object) As String
            If IsDate(pDate) Then
                Dim InDate As Date = CType(pDate, Date)
                DateTransferYYYY = CStr(CInt(Year(InDate) - 1911))
            Else
                DateTransferYYYY = ""
            End If
        End Function

        '�ഫ����~���褸�~
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

        '�b Console���Exception
        Public Shared Sub Debug(ByVal ex As Exception)
            System.Diagnostics.Debug.WriteLine(ex)
        End Sub

        '�b Console��ܰT��
        Public Shared Sub Debug(ByVal message As String)
            System.Diagnostics.Debug.WriteLine(message)
        End Sub

        '�b Console��ܰT��
        Public Shared Sub Debug(ByVal obj As Object)
            System.Diagnostics.Debug.WriteLine(obj)
        End Sub

#End Region

#Region "String"

        Public Shared Function string2ByteLength(ByVal str As String) As Integer
            Dim iSum As Integer = 0
            ' Console.WriteLine(System.Globalization.CultureInfo.CurrentCulture.TextInfo.ANSICodePage())
            Dim ch As CharEnumerator = str.GetEnumerator
            While ch.MoveNext
                Dim c As Char = ch.Current
                Dim i As Integer = Asc(c)
                If i <= 255 AndAlso i >= 0 Then  '�ȱq 0 �� 255 ����@�줸�զr����
                    iSum += 1
                    'Console.WriteLine(c)
                ElseIf i <= 32768 AndAlso i >= -32768 Then '�ȱq -32768 �� 32767 �����줸�զr����
                    iSum += 2
                    'Console.WriteLine(c)
                End If
            End While
            Return iSum
        End Function

        'Public Shared Function converStr2StrArray(ByVal str As String, ByVal sp1 As String, ByVal sp2 As String) As String(,)

        '    Dim arryStr(str.Split(sp1).Length - 1, str.Split(sp1)(0).Split(sp2).Length - 1) As String

        '    For i As Int16 = 0 To str.Split(sp1).Length - 1
        '        Dim strRow As String = str.Split(sp1)(i)
        '        For j As Int16 = 0 To strRow.Split(sp2).Length - 1
        '            Dim strCol As String = strRow.Split(",")(j)
        '            arryStr(i, j) = strCol
        '            Console.WriteLine("(" & i & "," & j & ")=" & strCol)
        '        Next j
        '    Next i
        '    Return arryStr
        'End Function

        Public Shared Function getString(ByVal sSource As String, ByVal sStartToken As String, ByVal sEndToken As String) As String
            Try

                'sSource = "Provider=OraOleDB.Oracle;max pool size=350;Password=KKKK;User ID=BBBB;Data Source=AAAAA"
                'sStartToken = "Password="
                'sEndToken = ";"
                Dim str As String
                Dim iPosition1 As Integer = sSource.ToLower.IndexOf(sStartToken.ToLower)
                Dim iStrLen As String = sSource.Length - iPosition1
                sSource = sSource.Substring(iPosition1, iStrLen)
                Dim iPosition2 As Integer = sSource.ToLower.IndexOf(sEndToken.ToLower.Trim)
                If iPosition2 = 0 Then
                    sSource = sSource.Substring(1, sSource.Length - 1)
                    iPosition2 = sSource.IndexOf(sEndToken.Trim)
                    str = sSource.Substring(sStartToken.Length - 1, iPosition2 - sStartToken.Length)
                ElseIf iPosition2 <> -1 Then
                    str = sSource.Substring(sStartToken.Length, iPosition2 - sStartToken.Length)
                Else
                    str = sSource.Substring(sStartToken.Length)
                End If

                Return str

            Catch ex As Exception
                Throw
            End Try
        End Function

        'Unicode �� Numeric Character Reference(NCR)(html Unicode)
        Public Shared Function transferUnicodeToNCR(ByVal str As String) As String
            Dim sb As New System.Text.StringBuilder
            Dim big5 As System.Text.Encoding = System.Text.Encoding.GetEncoding("big5")
            For Each c As Char In str
                '�P�_��X��Big5�A�ݷ|���|�ܦ��ݸ�
                Dim cInBig5 As String = big5.GetString(big5.GetBytes(New Char() {c}))
                '��Ӥ��O�ݸ��A��X���ܰݸ��A�P�w�����r
                If (c <> "?" And cInBig5 = "?") Then
                    'Dim iAsc As Integer = AscW(c) 'Unicode �r�X             
                    'sb.Append("&#").Append(iAsc).Append(";")
                    sb.AppendFormat("&#{0};", Convert.ToInt32(c))
                Else
                    sb.Append(c)
                End If
            Next
            Return sb.ToString()
        End Function

        'Numeric Character Reference(NCR)(html Unicode) �� Unicode 
        Public Shared Function transferNCRToUnicode(ByVal str As String) As String
            For Each m As System.Text.RegularExpressions.Match In System.Text.RegularExpressions.Regex.Matches(str, "&#(?<ncr>\d+?);")
                str = str.Replace(m.Value, Convert.ToChar(Integer.Parse(m.Groups("ncr").Value)).ToString())
            Next
            Return str
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
            ElseIf TypeOf obj Is Integer Then
                Return True
            ElseIf TypeOf obj Is Boolean Then
                Return obj
            ElseIf TypeOf obj Is System.Data.DataSet Then
                Return IIf(IsNothing(obj), False, isValidateData(CType(obj, System.Data.DataSet).Tables(0)))
            ElseIf TypeOf obj Is System.Data.DataTable Then
                Return IIf(IsNothing(obj), False, True)
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

#End Region

#Region "���ԧB�Ʀr�ର�j�g��r"

        Private Shared ReadOnly ar() As String = {"�s", "��", "�L", "��", "�v", "��", "��", "�m", "��", "�h"}
        Private Shared ReadOnly cName() As String = {"", "", "�B", "��", "�a", "�U", "�B", "��", "�a", "��", "�B", "��", "�a"}

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
                cunit = cName(j) '���X��� 
                If cNum = 0 Then '�P�_���X���Ʀr�O�_��0,�p�G�O0,�h�O���@���X0 
                    cZero = cZero + 1
                    If (InStr("�U��", cunit) > 0 And (cLast = "")) Then '�p�G���X���O�U,��,�h��ƥH�U���Ӹ� 
                        cLast = cunit
                    End If
                Else
                    If cZero > 0 Then '�p�G���X���Ʀr0��n��,�h�H�s�N���Ҧ���0 
                        If InStr("�U��", Right(conver, 1)) = 0 Then
                            conver = conver + cLast '�p�G�̫�@�줣�O��,�U,�h�̫�@��ɤW"���U" 
                        End If
                        conver = conver + "�s"
                        cZero = 0
                        cLast = ""
                    End If
                    conver = conver + ar(cNum) + cunit '�p�G���X���Ʀr�S��0,�h�O����Ʀr+��� 
                End If
                i = i + 1
            Next
            '�P�_�Ʀr���̫�@��O�_��0,�p�G�̫�@�쬰0,�h��U���ɤW 
            If InStr("�U��", Right(conver, 1)) = 0 Then
                conver = conver + cLast '�p�G�̫�@�줣�O��,�U,�h�̫�@��ɤW"���U" 
            End If

            'Return "�s�x��" & conver & "����"
            Return conver
        End Function
#End Region

#Region "web"

        '[���o�Ҧ�URL�Ѽ� �^��URL�r��]
        Public Shared Function getURLParas(ByVal request As System.Web.HttpRequest) As String
            Dim sbURL As New System.Text.StringBuilder
            Dim sCol As System.Collections.Specialized.NameValueCollection

            sCol = request.QueryString
            Try
                For Each sReqKey As String In sCol.AllKeys
                    If Not IsNothing(sReqKey) Then
                        Dim sReqValue As String = sCol.Item(sReqKey)
                        sbURL.Append(sReqKey & "=" & sReqValue & "&")
                    End If
                Next
                If sbURL.Length > 0 Then
                    Return sbURL.Remove(sbURL.ToString().Trim().Length() - 1, 1).ToString
                Else
                    Return ""
                End If
            Catch ex As Exception
                Throw
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
                If Not IsNothing(sw) Then sw.Close()
                If Not IsNothing(tw) Then tw.Close()
            End Try
        End Function


        Public Shared Function getHtml(ByVal page As System.Web.UI.Page, ByVal sw As System.IO.StringWriter) As String
            Dim tw As System.Web.UI.HtmlTextWriter = Nothing
            Try
                sw = New System.IO.StringWriter
                tw = New System.Web.UI.HtmlTextWriter(sw)
                page.RenderControl(tw)
                Return sw.ToString()
            Finally
                If Not IsNothing(sw) Then sw.Close()
                If Not IsNothing(tw) Then tw.Close()
            End Try
        End Function
#End Region

#Region "Compress"

        Public Shared Function DeCompress(ByVal bytes() As Byte) As Byte()
            Dim inflaterInputStream As ICSharpCode.SharpZipLib.Zip.Compression.Streams.InflaterInputStream = New ICSharpCode.SharpZipLib.Zip.Compression.Streams.InflaterInputStream(New System.IO.MemoryStream(bytes))
            Dim memoryStream As System.IO.MemoryStream = New System.IO.MemoryStream
            Dim iSize As Int32
            Dim byteWriteData(4096) As Byte
            While (True)
                iSize = inflaterInputStream.Read(byteWriteData, 0, byteWriteData.Length)
                If (iSize > 0) Then
                    memoryStream.Write(byteWriteData, 0, iSize)
                Else
                    Exit While
                End If
            End While

            inflaterInputStream.Close()
            Return memoryStream.ToArray()
        End Function

        '<summary>                                                                                                                                                                                 
        '�����YViewState�r��                                                                                                                                                                       
        '</summary>                                                                                                                                                                                
        '<param name="pViewState">ViewState�r��</param>                                                                                                                                            
        '<returns>��^�y���줸�հ}�C</returns>  
        Public Shared Function DeCompress(ByVal sViewState As String) As Byte()
            '�NBase64�r���ഫ���줸�հ}�C  
            Dim bytes As Byte() = System.Convert.FromBase64String(sViewState)
            Dim inflaterInputStream As ICSharpCode.SharpZipLib.Zip.Compression.Streams.InflaterInputStream = New ICSharpCode.SharpZipLib.Zip.Compression.Streams.InflaterInputStream(New System.IO.MemoryStream(bytes))
            '�Ыؤ䴩�O����s�x���y 
            Dim memoryStream As System.IO.MemoryStream = New System.IO.MemoryStream
            Dim iSize As Int32
            Dim byteWriteData(4096) As Byte

            While (True)
                iSize = inflaterInputStream.Read(byteWriteData, 0, byteWriteData.Length)
                If (iSize > 0) Then
                    memoryStream.Write(byteWriteData, 0, iSize)
                Else
                    Exit While
                End If
            End While

            inflaterInputStream.Close()
            Return memoryStream.ToArray()
        End Function

        Public Shared Function Compress(ByVal bytes() As Byte) As Byte()
            Dim memoryStream As System.IO.MemoryStream = New System.IO.MemoryStream
            Dim deflater As deflater = New deflater(ICSharpCode.SharpZipLib.Zip.Compression.Deflater.BEST_COMPRESSION)
            Dim deflaterOutputStream As ICSharpCode.SharpZipLib.Zip.Compression.Streams.DeflaterOutputStream = New ICSharpCode.SharpZipLib.Zip.Compression.Streams.DeflaterOutputStream(memoryStream, deflater, 131072)
            deflaterOutputStream.Write(bytes, 0, bytes.Length)
            deflaterOutputStream.Close()
            Return memoryStream.ToArray()
        End Function

        ' <summary>   
        ' ��r��i�����Y   
        ' </summary>   
        '<param name="pViewState">ViewState�r��</param>   
        '<returns>��^�y���줸�հ}�C</returns>  
        Public Shared Function Compress(ByVal sViewStateStr As String, ByVal iZipLevel As Int32) As Byte()
            '�N�r���ഫ���줸�հ}�C
            Dim bytes() As Byte = System.Convert.FromBase64String(sViewStateStr)
            '�إ߰O���骺Stream 
            Dim memoryStream As System.IO.MemoryStream = New System.IO.MemoryStream
            Dim deflater As deflater = New deflater(iZipLevel)
            Dim deflaterOutputStream As ICSharpCode.SharpZipLib.Zip.Compression.Streams.DeflaterOutputStream = New ICSharpCode.SharpZipLib.Zip.Compression.Streams.DeflaterOutputStream(memoryStream, deflater, 131072)
            deflaterOutputStream.Write(bytes, 0, bytes.Length)
            deflaterOutputStream.Close()
            Return memoryStream.ToArray()
        End Function




#End Region

#Region "difference datatable"
        Public Shared Function difference(ByVal First As DataTable, ByVal Second As DataTable) As DataTable
            'Create Empty Table
            Dim table As DataTable = New DataTable("Difference")
            'Must use a Dataset to make use of a DataRelation object

            Dim ds As DataSet = New DataSet()
            Using (ds)
                'Add tables
                ds.Tables.AddRange(New DataTable() {First.Copy(), Second.Copy()})
                'Get Columns for DataRelation
                Dim firstcolumns(ds.Tables(0).Columns.Count - 1) As DataColumn

                For i As Integer = 0 To firstcolumns.Length - 1
                    firstcolumns(i) = ds.Tables(0).Columns(i)
                Next

                Dim secondcolumns(ds.Tables(1).Columns.Count - 1) As DataColumn
                For i As Integer = 0 To secondcolumns.Length - 1
                    secondcolumns(i) = ds.Tables(1).Columns(i)
                Next

                'Create DataRelation
                Dim r As DataRelation = New DataRelation(String.Empty, firstcolumns, secondcolumns, False)
                ds.Relations.Add(r)

                'Create columns for return table
                For i As Integer = 0 To First.Columns.Count - 1
                    table.Columns.Add(First.Columns(i).ColumnName, First.Columns(i).DataType)
                Next

                'If First Row not in Second, Add to return table.
                table.BeginLoadData()

                For Each parentrow As DataRow In ds.Tables(0).Rows
                    Dim childrows() As DataRow = parentrow.GetChildRows(r)

                    If (IsNothing(childrows) Or childrows.Length = 0) Then
                        table.LoadDataRow(parentrow.ItemArray, True)
                    End If

                Next
                table.EndLoadData()

            End Using

            Return table
        End Function
#End Region

#Region "���o�һݲե�"

        ' <summary>
        ' ���o�һݲե�
        ' </summary>
        ' <param name="containGAC">�O�_�n�]�tGAC�ե�</param>
        ' <param name="finder">�L�o�禡</param>
        ' <returns></returns>

        'public static IEnumerable<Assembly> GetAssemblies(bool containGAC = false, Func<Assembly, bool> finder = null)

        Public Shared Function GetAssemblies(Optional containGAC As Boolean = False, Optional finder As Action(Of Assembly, Boolean) = Nothing) As IEnumerable(Of Assembly)
            Dim app As AppDomain = AppDomain.CurrentDomain

            Dim dynamicDirectory As String = IIf(app.SetupInformation.CachePath = "", String.Empty, app.SetupInformation.CachePath)
            Dim list As New List(Of Assembly)
            For Each assembly As Assembly In app.GetAssemblies()
                '�]�����i��|ShadowCopy
                'If (assembly.IsDynamic OrElse Not assembly.Location.Contains(dynamicDirectory)) Then
                If (Not assembly.Location.Contains(dynamicDirectory)) Then
                    Continue For
                End If
                If (Not containGAC AndAlso assembly.GlobalAssemblyCache) Then
                    Continue For
                End If
                list.Add(assembly)
            Next
            Return list
        End Function

        ' <summary>
        '�M��Ҧ��l��Type
        ' </summary>
        ' <param name="baseType">��¦Type</param>
        ' <param name="containGAC">�O�_�n�]�tGAC�ե�</param>
        ' <param name="findNonPublic">�O�_�n�]�t�p��Type</param>
        ' <returns></returns>
        'public static IEnumerable<Type> FindDerivedTypes(Type baseType, bool containGAC = false, bool findNonPublic = false
        Public Shared Function findDerivedTypes(ByVal baseType As Type, Optional containGAC As Boolean = False, Optional findNonPublic As Boolean = False) As IEnumerable(Of Type)
            'Dim derivedTypes As List(Of IEnumerable) = Nothing
            Dim derivedTypes As New Dictionary(Of Type, IEnumerable)
            SyncLock (derivedTypes)

                Dim derived As IEnumerable(Of Type) = Nothing
                '�]���ۤv�g���ե󳣤��|��bGAC���A�ҥH�w�]�|�L�o��GAC���ե�A�[�ַj�M�t��
                If Not derivedTypes.TryGetValue(baseType, derived) Then
                    Dim list As List(Of Type) = New List(Of Type)
                    'oreach (var assembly in GetAssemblies(containGAC))
                    For Each assembly In GetAssemblies(False)

                        For Each type As Type In assembly.GetTypes

                            If type.Equals(baseType) Then
                                Continue For
                            End If

                            If (type.IsNotPublic OrElse type.IsNestedPrivate) Then
                                If Not findNonPublic Then
                                    Continue For
                                End If
                            End If

                            If (type.IsAssignableFrom(baseType)) Then
                                list.Add(type)
                            End If
                        Next
                    Next

                    derived = list
                    derivedTypes.Add(baseType, list)
                End If
                Return derived
            End SyncLock
        End Function
#End Region

#Region "For SQL Server"
        Public Shared Function getSQLDIF(ByVal sSQLWORD As String) As String
            Try
                Select Case UCase(sSQLWORD)
                    Case "MINUS"
                        If com.Azion.NET.VB.Properties.getProvider = ProviderType.SqlClient Then
                            Return " EXCEPT "
                        Else
                            Return " " & sSQLWORD & " "
                        End If
                    Case "NVL"
                        If com.Azion.NET.VB.Properties.getProvider = ProviderType.SqlClient Then
                            Return " ISNULL "
                        Else
                            Return " " & sSQLWORD & " "
                        End If
                    Case "SYSDATE"
                        If com.Azion.NET.VB.Properties.getProvider = ProviderType.SqlClient Then
                            Return " GETDATE() "
                        Else
                            Return " " & sSQLWORD & " "
                        End If

                    Case "||"
                        If com.Azion.NET.VB.Properties.getProvider = ProviderType.SqlClient Then
                            Return " + "
                        Else
                            Return " " & sSQLWORD & " "
                        End If

                    Case "SUBSTR"
                        If com.Azion.NET.VB.Properties.getProvider = ProviderType.SqlClient Then
                            Return " SUBSTRING "
                        Else
                            Return " " & sSQLWORD & " "
                        End If

                    Case "SUBSTRB"
                        If com.Azion.NET.VB.Properties.getProvider = ProviderType.SqlClient Then
                            Return " SUBSTRING "
                        Else
                            Return " " & sSQLWORD & " "
                        End If

                    Case "TO_CHAR"
                        If com.Azion.NET.VB.Properties.getProvider = ProviderType.SqlClient Then
                            Return " STR "
                        Else
                            Return " " & sSQLWORD & " "
                        End If

                    Case "TO_DATE"
                        If com.Azion.NET.VB.Properties.getProvider = ProviderType.SqlClient Then
                            Return " CAST "
                        Else
                            Return " " & sSQLWORD & " "
                        End If
                    Case Else
                        Return String.Empty
                End Select
            Catch ex As Exception
                Throw ex
            End Try
        End Function

        Public Shared Function getSQLDIF2(ByVal sSQLWORD As String) As String
            Try
                If com.Azion.NET.VB.Properties.getProvider = ProviderType.SqlClient Then
                    sSQLWORD = Microsoft.VisualBasic.Strings.Replace(sSQLWORD, "{MINUS}", " EXCEPT ", CompareMethod.Text)
                    sSQLWORD = Microsoft.VisualBasic.Strings.Replace(sSQLWORD, "{NVL}", " ISNULL ", CompareMethod.Text)
                    sSQLWORD = Microsoft.VisualBasic.Strings.Replace(sSQLWORD, "{SYSDATE}", " GETDATE() ", CompareMethod.Text)
                    sSQLWORD = Microsoft.VisualBasic.Strings.Replace(sSQLWORD, "{||}", " + ", CompareMethod.Text)
                    sSQLWORD = Microsoft.VisualBasic.Strings.Replace(sSQLWORD, "{SUBSTR}", " SUBSTRING ", CompareMethod.Text)
                    sSQLWORD = Microsoft.VisualBasic.Strings.Replace(sSQLWORD, "{SUBSTRB}", " SUBSTRING ", CompareMethod.Text)
                    sSQLWORD = Microsoft.VisualBasic.Strings.Replace(sSQLWORD, "{TO_CHAR}", " STR ", CompareMethod.Text)
                    sSQLWORD = Microsoft.VisualBasic.Strings.Replace(sSQLWORD, "{TO_DATE}", " CAST ", CompareMethod.Text)
                Else
                    sSQLWORD = Microsoft.VisualBasic.Strings.Replace(sSQLWORD, "{MINUS}", " MINUS ", CompareMethod.Text)
                    sSQLWORD = Microsoft.VisualBasic.Strings.Replace(sSQLWORD, "{NVL}", " NVL ", CompareMethod.Text)
                    sSQLWORD = Microsoft.VisualBasic.Strings.Replace(sSQLWORD, "{SYSDATE}", " SYSDATE ", CompareMethod.Text)
                    sSQLWORD = Microsoft.VisualBasic.Strings.Replace(sSQLWORD, "{||}", " || ", CompareMethod.Text)
                    sSQLWORD = Microsoft.VisualBasic.Strings.Replace(sSQLWORD, "{SUBSTR}", " SUBSTR ", CompareMethod.Text)
                    sSQLWORD = Microsoft.VisualBasic.Strings.Replace(sSQLWORD, "{SUBSTRB}", " SUBSTRB ", CompareMethod.Text)
                    sSQLWORD = Microsoft.VisualBasic.Strings.Replace(sSQLWORD, "{TO_CHAR}", " TO_CHAR ", CompareMethod.Text)
                    sSQLWORD = Microsoft.VisualBasic.Strings.Replace(sSQLWORD, "{TO_DATE}", " TO_DATE ", CompareMethod.Text)
                End If

                Return sSQLWORD

            Catch ex As Exception
                Throw
            End Try
        End Function

#End Region

        Public Shared Function getMySqlDateTime(ByVal D_NET As Object) As Object
            Try
                If IsDate(D_NET) AndAlso CDate(D_NET).Year > 1911 Then
                    Dim D_RTN As MySql.Data.Types.MySqlDateTime = New MySql.Data.Types.MySqlDateTime(CDate(D_NET))
                    Return D_RTN
                Else
                    Return Convert.DBNull
                End If
            Catch ex As Exception
                Throw ex
            End Try
        End Function

    End Class

       
End Namespace