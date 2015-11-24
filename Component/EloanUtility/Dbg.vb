Option Explicit On
Option Strict On

Imports System.Diagnostics


''' <summary>
''' 提供一組幫助您偵錯程式碼的方法和屬性
''' </summary>
''' <remarks>
''' 需要開啟 專案屬性->進階編譯選項->定義DEBUG常數 
''' </remarks>
Public Class Dbg

    'Public Shared enableBreak As Boolean = True

    ''' <summary>
    ''' 檢查條件；如果條件為 false，則會顯示列出呼叫堆疊的訊息方塊。
    ''' </summary>
    ''' <param name="condition"></param>
    ''' <param name="bBreak"></param>
    ''' <remarks></remarks>
    Public Shared Sub Assert(ByVal condition As Boolean, Optional ByVal bBreak As Boolean = True)
#If DEBUG Then
        Debug.Assert(condition)

#If Not _SKIP_BREAK Then
        If condition OrElse System.Diagnostics.Debugger.IsAttached = False OrElse bBreak = False Then
            Exit Sub
        End If

        System.Diagnostics.Debugger.Break()
#End If
#End If
    End Sub


    ''' <summary>
    ''' 檢查條件；如果條件為 false，則會輸出指定的訊息並且顯示列出呼叫堆疊的訊息方塊。
    ''' </summary>
    ''' <param name="condition"></param>
    ''' <param name="message"></param>
    ''' <param name="bBreak"></param>
    ''' <remarks></remarks>
    Public Shared Sub Assert(ByVal condition As Boolean, ByVal message As String, Optional ByVal bBreak As Boolean = True)
#If DEBUG Then
        Debug.Assert(condition, message)

#If Not _SKIP_BREAK Then
        If condition OrElse System.Diagnostics.Debugger.IsAttached = False OrElse bBreak = False Then
            Exit Sub
        End If

        System.Diagnostics.Debugger.Break()
#End If
#End If
    End Sub


    ''' <summary>
    ''' 檢查條件；如果條件為 false，則會輸出兩個指定的訊息並且顯示列出呼叫堆疊的訊息方塊。
    ''' </summary>
    ''' <param name="condition"></param>
    ''' <param name="message"></param>
    ''' <param name="detailMessage"></param>
    ''' <param name="bBreak"></param>
    ''' <remarks></remarks>
    Public Shared Sub Assert(ByVal condition As Boolean, ByVal message As String, ByVal detailMessage As String, Optional ByVal bBreak As Boolean = True)
#If DEBUG Then
        Debug.Assert(condition, message, detailMessage)

#If Not _SKIP_BREAK Then
        If condition OrElse System.Diagnostics.Debugger.IsAttached = False OrElse bBreak = False Then
            Exit Sub
        End If

        System.Diagnostics.Debugger.Break()
#End If
#End If
    End Sub


    ''' <summary>
    ''' 檢查條件；如果條件為 false，則會輸出兩個訊息 (簡單和格式化) 並且顯示列出呼叫堆疊的訊息方塊。
    ''' </summary>
    ''' <param name="condition"></param>
    ''' <param name="message"></param>
    ''' <param name="detailMessageFormat"></param>
    ''' <param name="bBreak"></param>
    ''' <param name="args"></param>
    ''' <remarks></remarks>
    Public Shared Sub Assert(ByVal condition As Boolean, ByVal message As String, ByVal detailMessageFormat As String, ByVal bBreak As Boolean, ByVal ParamArray args() As Object)
#If DEBUG Then
        Debug.Assert(condition, message, detailMessageFormat, args)

#If Not _SKIP_BREAK Then
        If condition OrElse System.Diagnostics.Debugger.IsAttached = False OrElse bBreak = False Then
            Exit Sub
        End If

        System.Diagnostics.Debugger.Break()
#End If
#End If
    End Sub

    ''' <summary>
    ''' 表示已附加偵錯工具的中斷點。
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared Sub Break()
#If DEBUG Then

#If Not _SKIP_BREAK Then
        If System.Diagnostics.Debugger.IsAttached Then
            System.Diagnostics.Debugger.Break()
        End If
#End If

#End If
    End Sub



    Public Shared Sub LogWrite(ByVal message As String,
                               Optional ByVal type As EventLogEntryType = EventLogEntryType.Error)
        Try
            Diagnostics.EventLog.WriteEntry("Eloan", message, type)
        Catch e As Exception
        End Try
    End Sub


End Class
