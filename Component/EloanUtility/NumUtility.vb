Option Explicit On
Option Strict On

''' <summary>
''' 提供數字相關函式
''' Format....
''' [Titan] 	2011/07/19	Created
''' </summary>
Public Class NumUtility

#Region "Format"

    ''' <summary>
    ''' 金額的表示, 每3位數(千)加逗號，到小數點第二位，超過會自動四捨五入
    ''' if isZero2Dash = True then  0 --> 顯示 -(Dash) else 0 -->顯示 0
    ''' IntegerUtility.str2Amount(0,True)--> 這個會顯示 -
    ''' IntegerUtility.str2Amount(0,False)--> 這個會顯示 0.00
    ''' </summary>
    ''' <param name="obj">Object</param>
    ''' <param name="isZero2Dash">Boolean</param>
    ''' <returns>String</returns>
    ''' <remarks>
    ''' [Titan] 	2011/07/19	Created
    ''' </remarks>  
    Public Shared Function str2Amount(ByVal obj As Object, Optional isZero2Dash As Boolean = False) As String
        If IsNumeric(obj) Then
            If isZero2Dash Then
                Return String.Format("{0:#,##0.00;(#,##0.00);-}", CDec(obj))
            End If
            Return String.Format("{0:#,##0.00;(#,##0.00);0.00}", CDec(obj))
        End If

        If isZero2Dash Then
            Return "-"
        End If

        Return "0"
    End Function

    ''' <summary>
    ''' 百分比
    ''' if isZero2Dash = True then  0 --> 顯示 -(Dash) else 0 -->顯示 0
    ''' IntegerUtility.str2Amount(0,True)--> 這個會顯示 -
    ''' IntegerUtility.str2Amount(0,False)--> 這個會顯示 0.00%
    ''' </summary>
    ''' <param name="obj">Object</param>
    ''' <param name="isZero2Dash">Boolean</param>
    ''' <returns>String</returns>
    ''' <remarks>
    ''' [Titan] 	2011/07/19	Created
    ''' </remarks>  
    Public Shared Function str2Percen(ByVal obj As Object, Optional isZero2Dash As Boolean = False) As String
        If IsNumeric(obj) Then
            If isZero2Dash Then
                Return String.Format("{0:0.00;(0.00);-}", CDbl(obj)) & "%"
            End If
            Return String.Format("{0:0.00;(0.00);0.00}", CDbl(obj)) & "%"
        End If

        If isZero2Dash Then
            Return "-"
        End If

        Return "0.00%"
    End Function

#End Region

#Region "阿拉伯數字轉為大寫國字"

    Private Shared ReadOnly ar() As String = {"零", "壹", "貳", "參", "肆", "伍", "陸", "柒", "捌", "玖"}
    Private Shared ReadOnly cName() As String = {"", "", "拾", "佰", "仟", "萬", "拾", "佰", "仟", "億", "拾", "佰", "仟"}
    ''' <summary>
    ''' 阿拉伯數字轉為大寫國字
    ''' </summary>
    ''' <param name="sNum"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' [Titan] 	2011/07/19	Created
    ''' </remarks>  
    Public Shared Function cypherConverCNum(ByVal sNum As String) As String
        Dim cNum As Integer
        Dim cUnit As String
        Dim i As Integer = 1
        Dim cZero As Integer = 0
        Dim sConver As String = String.Empty
        Dim cLast As String = String.Empty


        For j As Integer = Len(sNum) To 1 Step -1
            cNum = CInt(Val(Mid(sNum, i, 1)))
            cUnit = cName(j) '取出位數 
            If cNum = 0 Then '判斷取出的數字是否為0,如果是0,則記錄共有幾0 
                cZero = cZero + 1
                If (InStr("萬億", cUnit) > 0 And (cLast = "")) Then '如果取出的是萬,億,則位數以萬億來補 
                    cLast = cUnit
                End If
            Else
                If cZero > 0 Then '如果取出的數字0有n個,則以零代替所有的0 
                    If InStr("萬億", Right(sConver, 1)) = 0 Then
                        sConver = sConver + cLast '如果最後一位不是億,萬,則最後一位補上"億萬" 
                    End If
                    sConver = sConver + "零"
                    cZero = 0
                    cLast = ""
                End If
                sConver = sConver + ar(cNum) + cUnit '如果取出的數字沒有0,則是中文數字+單位 
            End If
            i = i + 1
        Next
        '判斷數字的最後一位是否為0,如果最後一位為0,則把萬億補上 
        If InStr("萬億", Right(sConver, 1)) = 0 Then
            sConver = sConver + cLast '如果最後一位不是億,萬,則最後一位補上"億萬" 
        End If

        'Return "新台幣" & conver & "元整"
        Return sConver
    End Function
#End Region
End Class
