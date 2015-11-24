Option Explicit On
Option Strict On


Imports System.IO
Imports System.Data

''' <summary>
''' 登入者物件
''' </summary>
''' <remarks></remarks>
''' <history>
''' [Lake] 2012/04/06 Created
''' </history>
<SerializableAttribute()> _
Public Class StaffInfo

    Private _SessionID As String

    Private _LoginUserId As String
    Private _LoginUserName As String
    Private _LoginTopDepNo As String
    Private _LoginBrid As String
    Private _LoginBrCname As String
    Private _LoginRoleId As String

    'Private _BridList As DataTable
    'Private _Bra_DepNoList As String
     
    Private _WorkingStaffid As String
    Private _WorkingName As String
    Private _WorkingTopDepNo As String
    Private _WorkingBrid As String
    Private _WorkingBrCname As String
    Private _WorkingRoleId As String

    Private _CaseId As String
    Private _FuncList As New Hashtable

    Private _SysId As String

    ''' <summary>
    ''' 系統編號
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property SysId As String
        Get
            Return _SysId
        End Get
        Set(ByVal value As String)
            _SysId = value
        End Set
    End Property

    ''' <summary>
    ''' 登入者ID
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property LoginUserId As String
        Get
            Return _LoginUserId
        End Get
        Set(ByVal value As String)
            _LoginUserId = value
        End Set
    End Property

    ''' <summary>
    ''' 登入者名稱
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property LoginUserName As String
        Get
            Return _LoginUserName
        End Get
        Set(ByVal value As String)
            _LoginUserName = value
        End Set
    End Property

    ''' <summary>
    ''' 登入者默認單位
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property LoginBrid As String
        Get
            Return _LoginBrid
        End Get
        Set(ByVal value As String)
            _LoginBrid = value
        End Set
    End Property

    ''' <summary>
    ''' 登入者最上層部門
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property LoginTopDepNo As String
        Get
            Return _LoginTopDepNo
        End Get
        Set(ByVal value As String)
            _LoginTopDepNo = value
        End Set
    End Property

    ''' <summary>
    ''' 登入者默認角色
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property LoginRoleId As String
        Get
            Return _LoginRoleId
        End Get
        Set(ByVal value As String)
            _LoginRoleId = value
        End Set
    End Property

    ''' <summary>
    ''' 登入者id or所代理id
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property WorkingStaffid As String
        Get
            Return _WorkingStaffid
        End Get
        Set(ByVal value As String)
            _WorkingStaffid = value
        End Set
    End Property

    ''' <summary>
    ''' 登入者name or 所代理 name
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property WorkingName As String
        Get
            Return _WorkingName
        End Get
        Set(ByVal value As String)
            _WorkingName = value
        End Set
    End Property

    ''' <summary>
    ''' 登入者單位 or 所代理單位
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property WorkingBrid As String
        Get
            Return _WorkingBrid
        End Get
        Set(ByVal value As String)
            _WorkingBrid = value
        End Set
    End Property

    ''' <summary>
    ''' 部門名稱
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property WorkingBrCname As String
        Get
            Return _WorkingBrCname
        End Get
        Set(ByVal value As String)
            _WorkingBrCname = value
        End Set
    End Property

    ''' <summary>
    ''' 登入者部門最上層部門代號 or 所代理部門最上層部門代號
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property WorkingTopDepNo As String
        Get
            Return _WorkingTopDepNo
        End Get
        Set(ByVal value As String)
            _WorkingTopDepNo = value
        End Set
    End Property

    ''' <summary>
    ''' 登入者默認角色 or 所代理角色
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property WorkingRoleId As String
        Get
            Return _WorkingRoleId
        End Get
        Set(ByVal value As String)
            _WorkingRoleId = value
        End Set
    End Property

    ''' <summary>
    ''' 記錄案件編號，點選案件后賦值
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property CaseId As String
        Get
            Return _CaseId
        End Get
        Set(ByVal value As String)
            _CaseId = value
        End Set
    End Property

    ''' <summary>
    ''' 儲存交易資料 通過系統編號/詳細資料的組合組合在一起
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property FuncList As Hashtable
        Get
            Return _FuncList
        End Get
        Set(ByVal value As Hashtable)
            _FuncList = value
        End Set
    End Property

    ''' <summary>
    ''' 部門名稱
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property LoginBrCname As String
        Get
            Return _LoginBrCname
        End Get
        Set(ByVal value As String)
            _LoginBrCname = value
        End Set
    End Property


    Public Property SessionID As String
        Get
            Return _SessionID
        End Get
        Set(ByVal value As String)
            _SessionID = value
        End Set
    End Property
End Class
