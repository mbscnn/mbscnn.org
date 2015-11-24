Option Explicit On
Option Strict On

Public Class Rules

    Private _Rule As ValidateType
    Private _Message As String
    Private _CompareTo As Object

    ''' <summary>
    ''' 驗證規則
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Rule() As ValidateType
        Get
            Return _Rule
        End Get
        Set(ByVal Value As ValidateType)
            _Rule = Value
        End Set
    End Property

    ''' <summary>
    ''' 提示信息
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Message() As String
        Get
            Return _Message
        End Get
        Set(ByVal Value As String)
            _Message = Value
        End Set
    End Property

    ''' <summary>
    ''' 被比較的欄位ID或字符或數字
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property CompareTo() As Object
        Get
            Return _CompareTo
        End Get
        Set(ByVal Value As Object)
            _CompareTo = Value
        End Set
    End Property

    ''' <summary>
    ''' 構造函數
    ''' </summary>
    ''' <param name="vType">驗證類型</param>
    ''' <param name="sMessage">提示信息</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal vType As ValidateType, ByVal sMessage As String)
        Me.Rule = vType
        Me.Message = sMessage
    End Sub

    ''' <summary>
    ''' 構造函數
    ''' </summary>
    ''' <param name="vType">驗證類型</param>
    ''' <param name="sMessage">提示信息</param>
    ''' <param name="oCompareTo">被比較的欄位ID或字符或數字</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal vType As ValidateType, ByVal sMessage As String, ByVal oCompareTo As Object)
        Me.Rule = vType
        Me.Message = sMessage
        Me.CompareTo = oCompareTo
    End Sub
End Class
