Imports System
Imports System.Data
Imports System.Configuration
Imports System.Collections
Imports System.Web
Imports System.Web.Security
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.UI.WebControls.WebParts
Imports System.Web.UI.HtmlControls

''' <summary>
''' 委派宣告
''' </summary>
Public Delegate Sub BindDataDelegate()

''' <summary>
''' 分頁自定義控件屬性
''' </summary>
Partial Public Class PagerMenu
    Inherits System.Web.UI.UserControl

    ''' <summary>
    ''' 委託Grid綁定數據的方法
    ''' </summary>
    Public BindData As BindDataDelegate

    ''' <summary>
    ''' 分頁筆數,預設為20
    ''' </summary>
    Private _pageSize As Integer = 0

    ''' <summary>
    ''' grid當前顯示頁面的索引
    ''' </summary>
    Private _pageIndex As Integer = 0

    ''' <summary>
    ''' Grid顯示資料總筆數
    ''' </summary>
    Private _itemCount As Integer = 0

    ''' <summary>
    ''' 是否自動隱藏: true=是/false=否
    ''' </summary>
    Private _autoHidden As Boolean = False

    ''' <summary>
    ''' 目標Grid
    ''' </summary>
    Private _gvTarget As New DataGrid


#Region "屬性"

    ''' <summary>
    ''' 取得或設定Grid每頁顯示的資料筆數
    ''' </summary>
    Public Property PageSize() As Integer
        Get
            Return Me._pageSize
        End Get
        Set(value As Integer)
            Me._pageSize = value
        End Set
    End Property

    ''' <summary>
    ''' 取得或設定Grid當前顯示頁面的索引
    ''' </summary>
    Public Property CurrentPageIndex() As Integer
        Get
            Return Me._pageIndex
        End Get
        Set(value As Integer)
            Me._pageIndex = value
        End Set
    End Property

    ''' <summary>
    ''' 取得或設定Grid顯示資料總筆數
    ''' </summary>
    Public Property ItemCount() As Integer
        Get
            Return Me._itemCount
        End Get
        Set(value As Integer)
            Me._itemCount = value
        End Set
    End Property

    ''' <summary>
    ''' 設定或取得導航控件是否自動隱藏
    ''' </summary>
    Public Property AutoHidden() As Boolean
        Get
            Return _autoHidden
        End Get
        Set(value As Boolean)
            _autoHidden = value
        End Set
    End Property

    ''' <summary>
    ''' 設定或取得導航控件對應的目標Grid
    ''' </summary>
    Public Property Target() As DataGrid
        Get
            Return _gvTarget
        End Get
        Set(value As DataGrid)
            _gvTarget = value
        End Set
    End Property


#Region "取得或設定第一頁按鈕的文字"

    ''' <summary>
    ''' 取得或設定第一頁按鈕的文字
    ''' </summary>
    Public Property FirstPageText() As String
        Get
            ' 取得第一頁按鈕的Text值
            Return Me.btnNavFirst.Text
        End Get
        Set(value As String)

            ' 取得第一頁按鈕的Text值
            Me.btnNavFirst.Text = value
        End Set
    End Property

#End Region

#Region "取得或設定上一頁按鈕的文字"

    ''' <summary>
    ''' 取得或設定上一頁按鈕的文字
    ''' </summary>
    Public Property PreviousPageText() As String
        Get
            ' 取得上一頁按鈕的Text值
            Return Me.btnNavPrevious.Text
        End Get
        Set(value As String)

            ' 取得上一頁按鈕的Text值
            Me.btnNavPrevious.Text = value
        End Set
    End Property
#End Region

#Region "取得或設定下一頁按鈕的文字"

    ''' <summary>
    ''' 取得或設定下一頁按鈕的文字
    ''' </summary>
    Public Property NextPageText() As String
        Get
            ' 取得下一頁按鈕的Text值
            Return Me.btnNavNext.Text
        End Get
        Set(value As String)

            ' 取得下一頁按鈕的Text值
            Me.btnNavNext.Text = value
        End Set
    End Property
#End Region

#Region "取得或設定最後一頁按鈕的文字"

    ''' <summary>
    ''' 取得或設定最後一頁按鈕的文字
    ''' </summary>
    Public Property LastPageText() As String
        Get
            ' 取得最後一頁按鈕的Text值
            Return btnNavLast.Text
        End Get
        Set(value As String)

            ' 取得最後一頁按鈕的Text值
            Me.btnNavLast.Text = value
        End Set
    End Property
#End Region

#End Region

    ''' <summary>
    ''' 頁面加載
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub Page1_Load(sender As Object, e As EventArgs)
        If Not Page.IsPostBack Then
            If Request("prePageIndex") <> Nothing AndAlso Request("prePageIndex") <> "" AndAlso Request("page") = Me.Page.[GetType]().Name Then
                Dim index As Integer = Integer.Parse(Request("prePageIndex"))

                _gvTarget.CurrentPageIndex = index
            Else
                Me._gvTarget.CurrentPageIndex = 0
                Me.SetButtonState(0)
            End If
        End If
    End Sub

    ''' <summary>
    ''' 設定導航對象對應的目標Grid以及針對Grid數據綁定的事件
    ''' </summary>
    ''' <param name="gvTarget">目標DataGrid對象</param>
    ''' <param name="delegateBindData">數據綁定事件</param>
    Public Sub SetTarget(gvTarget As DataGrid, delegateBindData As BindDataDelegate)
        Me._gvTarget = gvTarget
        Me.BindData = delegateBindData

        AddHandler _gvTarget.DataBinding, AddressOf GridDataBinding

        Me.SetStyle()

        Page1_Load(gvTarget, New EventArgs())
    End Sub

    ''' <summary>
    ''' 設定導航對象對應的目標Grid以及針對Grid數據綁定的事件
    ''' </summary>
    ''' <param name="gvTarget">目標DataGrid對象</param>
    ''' <param name="delegateBindData">數據綁定事件</param>
    ''' <param name="PageSize">Grid顯示資料筆數</param>
    Public Sub SetTarget(gvTarget As DataGrid, delegateBindData As BindDataDelegate, PageSize As Integer)
        Me._gvTarget = gvTarget
        Me.BindData = delegateBindData

        AddHandler _gvTarget.DataBinding, AddressOf GridDataBinding
        Me.SetStyle(PageSize)

        Page1_Load(gvTarget, New EventArgs())
    End Sub

    ''' <summary>
    ''' 前置參數賦值
    ''' </summary>
    Public Sub SetStyle()

        ' 設定分頁筆數
        Me._pageSize = Convert.ToInt32(com.Azion.EloanUtility.CodeList.getAppSettings("PageSize"))

        ' 啟用Grid分頁功能
        Me._gvTarget.AllowPaging = True

        ' 設定Grid每頁顯示筆數
        Me._gvTarget.PageSize = _pageSize

        ' 設定Grid本身的導航按鈕隱藏
        Me._gvTarget.PagerStyle.Visible = False
    End Sub


    Public Sub SetStyle(PageSize As Integer)
        ' 設定分頁筆數
        Me._pageSize = PageSize

        ' 啟用Grid分頁功能
        Me._gvTarget.AllowPaging = True

        ' 設定Grid每頁顯示筆數
        Me._gvTarget.PageSize = _pageSize

        ' 設定Grid本身的導航按鈕隱藏
        Me._gvTarget.PagerStyle.Visible = False
    End Sub

    ''' <summary>
    ''' 初始化
    ''' </summary>
    Public Sub Clear()
        Me._gvTarget.CurrentPageIndex = 0
    End Sub

    ''' <summary>		
    ''' 導航按鈕點擊事件
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub NavigationButtonClick(sender As Object, e As System.EventArgs)
        Try

            Dim direction As String = (DirectCast(sender, LinkButton)).CommandName

            Select Case direction.ToUpper()
                Case "FIRST"
                    Me._gvTarget.CurrentPageIndex = 0
                    Exit Select
                Case "PREVIOUS"
                    Me._gvTarget.CurrentPageIndex = Math.Max(_gvTarget.CurrentPageIndex - 1, 0)
                    Exit Select
                Case "NEXT"
                    Me._gvTarget.CurrentPageIndex = Math.Min(_gvTarget.CurrentPageIndex + 1, _gvTarget.PageCount - 1)
                    Exit Select
                Case "LAST"
                    Me._gvTarget.CurrentPageIndex = Math.Max(_gvTarget.PageCount - 1, 0)
                    Exit Select
                Case Else
                    Exit Select
            End Select

            ' 調用綁定Grid數據的的方法
            Me.BindData()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ''' <summary>
    ''' DataGrid綁定數據前,設定導航條
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Overridable Sub GridDataBinding(sender As Object, e As System.EventArgs)
        Try
            Dim newCount As Integer = 0
            Dim PageCount As Integer = 0

            If _gvTarget.DataSource Is Nothing Then
                SetButtonState(0)
                Return
            End If

            If _gvTarget.DataSource.[GetType]().ToString().ToLower() = "system.data.datatable" Then
                newCount = (DirectCast(_gvTarget.DataSource, DataTable)).Rows.Count
            ElseIf _gvTarget.DataSource.[GetType]().ToString().ToLower() = "system.data.dataview" Then
                newCount = (DirectCast(_gvTarget.DataSource, DataView)).Count
            ElseIf _gvTarget.DataSource.[GetType]().ToString().ToLower() = "system.data.dataset" Then
                newCount = (DirectCast(_gvTarget.DataSource, DataSet)).Tables(0).Rows.Count
            End If

            If newCount > 0 Then
                If newCount Mod _pageSize = 0 Then
                    PageCount = newCount / _pageSize
                Else
                    PageCount = Math.Floor((newCount / _pageSize)) + 1
                End If

                If _gvTarget.CurrentPageIndex > PageCount - 1 Then
                    _gvTarget.CurrentPageIndex = PageCount - 1
                End If
            Else
                PageCount = 0
                _gvTarget.CurrentPageIndex = 0
            End If

            Me._itemCount = newCount
            Me._pageIndex = Me._gvTarget.CurrentPageIndex

            Dim index As Integer = If((PageCount = 0), 0, Me._pageIndex + 1)


            ' 判斷導航控件是否顯示
            If Me._autoHidden Then
                Me.tbPagingNavigation.Visible = ((newCount - 1) / Me._pageSize >= 0)
            Else
                Me.tbPagingNavigation.Visible = True
            End If

            lblPageIndex.Text = index
            lblTotalPage.Text = PageCount
            lblTotalRecord.Text = PageSize

            ddlJumpPage.Items.Clear()

            If lblTotalPage.Text <> "" Then
                For a = 1 To Convert.ToInt32(lblTotalPage.Text)
                    ddlJumpPage.Items.Insert(a - 1, New ListItem(a, a))
                Next
            End If

            ddlJumpPage.SelectedValue = index

            ' 設定導航按鈕的狀態
            Me.SetButtonState(PageCount)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ''' <summary>
    ''' 設定導航按鈕的狀態
    ''' </summary>
    ''' <param name="pageCount"></param>
    Public Sub SetButtonState(pageCount As Integer)
        Me.btnNavFirst.Enabled = (_gvTarget.CurrentPageIndex > 0)
        Me.imgBtnFirst.Enabled = (_gvTarget.CurrentPageIndex > 0)

        Me.btnNavPrevious.Enabled = (_gvTarget.CurrentPageIndex > 0)
        Me.imgBtnPrevious.Enabled = (_gvTarget.CurrentPageIndex > 0)

        Me.btnNavNext.Enabled = (_gvTarget.CurrentPageIndex < pageCount - 1)
        Me.imgBtnNext.Enabled = (_gvTarget.CurrentPageIndex < pageCount - 1)

        Me.btnNavLast.Enabled = (_gvTarget.CurrentPageIndex < pageCount - 1)
        Me.imgBtnLast.Enabled = (_gvTarget.CurrentPageIndex < pageCount - 1)
    End Sub

    ''' <summary>
    ''' 跳頁
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub ddlJumpPage_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlJumpPage.SelectedIndexChanged
        Try
            Me._gvTarget.CurrentPageIndex = Convert.ToInt32(ddlJumpPage.SelectedValue) - 1

            ' 調用綁定DataGrid數據的方法
            Me.BindData()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
End Class