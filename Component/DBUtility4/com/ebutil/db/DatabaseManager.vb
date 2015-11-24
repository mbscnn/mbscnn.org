Option Explicit On
Option Strict On

Imports System.Data
Imports System.Collections.Concurrent
Imports com.Azion.EloanUtility

Public Class DatabaseManager
    'Inherits HttpHandler
    Implements System.IDisposable

    Private m_connection As IDbConnection
    Private m_transaction As IDbTransaction

    Dim m_bTransaction As Boolean
    Dim m_bCloseTransaction As Boolean

    Private m_sDataSource As String

    Private m_lLastUsed As Long
    Private m_sPath As String = String.Empty

    Private m_bUse As Boolean = True
    Protected m_sConstructorCallStack As String = String.Empty

    Public m_nWriteCount As Integer = 0
    Public TRANSACTION_WARNING As Boolean = True     '寫入第二次時，會檢查是否使用Transaction。若設為FALSE，可停用Transaction檢查

    Public Shared CREATEDBYUI_WARNING As Boolean = True     '檢查是否由UIBASE所建立的連線。若設為FALSE，可停用UIBASE所建立的連線檢查
    Public Shared AVOID_TABLE_LOCKING As Boolean = False


    Protected Shared m_ConnectionQueue As New ConcurrentQueue(Of IDbConnection)         '未關閉的Connection
    Protected Shared m_ConnectionExceptionQueue As New ConcurrentQueue(Of Exception)    '檢查未關閉Connection時產生的Exception

#Region "Property"

    ReadOnly Property ConnectionString() As String
        Get
            Return m_connection.ConnectionString()
        End Get
    End Property

    ReadOnly Property ConnectionTimeout() As Integer
        Get
            Return m_connection.ConnectionTimeout()
        End Get
    End Property

    ReadOnly Property Database() As String
        Get
            Return m_connection.Database()
        End Get
    End Property

    ReadOnly Property State() As Integer
        Get
            Return m_connection.State()
        End Get
    End Property

    Property LoginTime() As Long
        Get
            Return m_lLastUsed
        End Get
        Set(ByVal Value As Long)
            m_lLastUsed = Value
        End Set
    End Property

#End Region

#Region "get DatabaseManager"

    Public Shared Function getInstance(ByVal sDatabaseName As String, ByVal sUserName As String, ByVal sPassword As String, Optional ByVal iMaxConn As Integer = 0, Optional ByVal bCreatedByUIBase As Boolean = False) As DatabaseManager
        Dim dbManager As DatabaseManager = Nothing

        Try
            If IsNothing(dbManager) Then
                dbManager = newDatabase(sDatabaseName, sUserName, sPassword)
            End If

#If DEBUG Then
            'If bCreatedByUIBase = False AndAlso CREATEDBYUI_WARNING = True Then
            '    Dim l_CurrentStack As New System.Diagnostics.StackTrace(True)
            '    Trace.WriteLine(vbCrLf & "自行建立的Connection無法自動被釋放，且容易造成所更新的TABLE被鎖住：")
            '    Trace.WriteLine(l_CurrentStack.ToString() & vbCrLf)

            '    'Dbg.Assert(Process.GetCurrentProcess.ProcessName.ToUpper.CompareTo("W3WP") <> 0, "建立Connection要改成使用UIBASE建立Connection")
            'End If
#End If

            Return dbManager

        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function getInstance(ByVal sConnectionString As String, Optional ByVal iMaxConn As Integer = 0, Optional ByVal bCreatedByUIBase As Boolean = False) As DatabaseManager
        Dim dbManager As DatabaseManager = Nothing

        Try
            If IsNothing(dbManager) Then
                dbManager = newDatabase(sConnectionString)

                If DatabaseManager.AVOID_TABLE_LOCKING = True Then
                    Try
                        TinyDb.DbTable.getNewBosBase(dbManager).ExecuteNonQuery("SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED --MSG: 避免LOCKING，但會造成Dirty Read，開發時使用", Nothing)
                    Catch ex As Exception
                    End Try
                End If

            End If

#If DEBUG Then
            'If bCreatedByUIBase = False AndAlso CREATEDBYUI_WARNING = True Then
            '    Dim l_CurrentStack As New System.Diagnostics.StackTrace(True)
            '    Trace.WriteLine(vbCrLf & "自行建立的Connection無法自動被釋放，且容易造成所更新的TABLE被鎖住：")
            '    Trace.WriteLine(l_CurrentStack.ToString() & vbCrLf)

            '    'Dbg.Assert(Process.GetCurrentProcess.ProcessName.ToUpper.CompareTo("W3WP") <> 0, "建立Connection要改成使用UIBASE建立Connection")
            'End If
#End If

            Return dbManager

        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function getInstance(Optional ByVal bCreatedByUIBase As Boolean = False) As DatabaseManager

        Dim dbManager As DatabaseManager = Nothing
        Try
            If IsNothing(dbManager) Then
                dbManager = newDatabase()

                If DatabaseManager.AVOID_TABLE_LOCKING = True Then
                    Try
                        TinyDb.DbTable.getNewBosBase(dbManager).ExecuteNonQuery("SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED --MSG: 避免LOCKING，但會造成Dirty Read，開發時使用", Nothing)
                    Catch ex As Exception
                    End Try
                End If
            End If

#If DEBUG Then
            'If bCreatedByUIBase = False AndAlso CREATEDBYUI_WARNING = True Then
            '    Dim l_CurrentStack As New System.Diagnostics.StackTrace(True)
            '    Trace.WriteLine(vbCrLf & "自行建立的Connection無法自動被釋放，且容易造成所更新的TABLE被鎖住：")
            '    Trace.WriteLine(l_CurrentStack.ToString() & vbCrLf)

            '    'Dbg.Assert(Process.GetCurrentProcess.ProcessName.ToUpper.CompareTo("W3WP") <> 0, "建立Connection要改成使用UIBASE建立Connection")
            'End If
#End If

            Return dbManager
        Catch ex As System.Threading.ThreadAbortException
            Console.Write(ex)
        Catch ex As Exception
            Throw
        End Try
        Return Nothing
    End Function

    Friend Shared Function newDatabase(ByVal sDNS As String, ByVal sUserName As String, ByVal sPassword As String, Optional ByVal iMaxConn As Integer = 0) As DatabaseManager
        Dim dbManager As DatabaseManager = Nothing
        Dim sConnectionString As String = com.Azion.NET.VB.Properties.getDSN()
        Try
            If IsNothing(dbManager) Then
                dbManager = New DatabaseManager
                dbManager.open(sConnectionString)
                dbManager.setUse(True)
                dbManager.LoginTime = Now.Ticks
            End If

            Return dbManager
        Catch ex As Exception
            Throw
        End Try
    End Function

    Protected Friend Shared Function newDatabase(ByVal sConnectionString As String) As DatabaseManager
        Dim dbManager As DatabaseManager = Nothing

        Try
            If IsNothing(dbManager) Then
                dbManager = New DatabaseManager
                dbManager.open(sConnectionString)
                dbManager.setUse(True)
                dbManager.LoginTime = Now.Ticks
            End If

            Return dbManager
        Catch ex As Exception
            Throw
        End Try
    End Function

    Friend Shared Function newDatabase() As DatabaseManager
        Dim dbManager As DatabaseManager = Nothing

        Try
            If IsNothing(dbManager) Then
                dbManager = New DatabaseManager
                dbManager.open()
                dbManager.setUse(True)
                dbManager.LoginTime = Now.Ticks
            End If

            Return dbManager
        Catch ex As Exception
            Throw
        End Try
    End Function
#End Region

    Sub dispose() Implements System.IDisposable.Dispose
        Me.releaseConnection()
        GC.SuppressFinalize(Me)
    End Sub

    Sub releaseConnection()
        If Me.m_bTransaction And Not Me.m_bCloseTransaction Then
            Me.Rollback()
            Throw New Exception("未關閉 Transaction ! 交易已經 Rollback !")
        End If

        closeConnection()
    End Sub

    Private Sub New()
#If DEBUG Then
        Dim l_CurrentStack As New System.Diagnostics.StackTrace(True)
        m_sConstructorCallStack = l_CurrentStack.ToString()


        'If m_sConstructorCallStack.Contains("於 ASP.ga_ge3_ge3200_aspx.ProcessRequest(HttpContext context)") AndAlso
        '    m_sConstructorCallStack.Contains("於 com.Azion.UITools.UIBase.Page_PreInit(Object sender, EventArgs e) 於 D:\EnTie\Ent\PageTab\UIBase\UIBase.vb: 行 60") Then
        '    System.Diagnostics.Debugger.Break()
        'End If

#End If

        '關閉上次未關閉的Connection
        Dim conn As IDbConnection = Nothing

        Try
            While m_ConnectionQueue.TryDequeue(conn)
                If Not conn Is Nothing Then conn.Close()
            End While
        Catch ex As Exception
        End Try


#If DEBUG Then
        '上一次Connection未關閉時產生的Exception
        Dim except As Exception = Nothing
        If m_ConnectionExceptionQueue.TryDequeue(except) Then
            Throw except
        End If
#Else
        If Environment.MachineName = "WIN-QRLAKF0WDIF" orelse Environment.MachineName = "TSRVNT170" Then
            Dim except2 As Exception = Nothing
            If m_ConnectionExceptionQueue.TryDequeue(except2) Then
                Throw except2
            End If
        End If
#End If

    End Sub

    Private Sub New(ByVal sDatabaseName As String, ByVal sUserName As String, ByVal sPassword As String, ByVal iMaxConn As Integer)
        init(sDatabaseName, sUserName, sPassword, iMaxConn)

#If DEBUG Then
        Dim l_CurrentStack As New System.Diagnostics.StackTrace(True)
        m_sConstructorCallStack = l_CurrentStack.ToString()
#End If

        '關閉上次未關閉的Connection
        Dim conn As IDbConnection = Nothing

        Try
            While m_ConnectionQueue.TryDequeue(conn)
                If Not conn Is Nothing Then conn.Close()
            End While
        Catch ex As Exception
        End Try


#If DEBUG Then
        '上一次Connection未關閉時產生的Exception
        Dim except As Exception = Nothing
        If m_ConnectionExceptionQueue.TryDequeue(except) Then
            Throw except
        End If
#Else
        If Environment.MachineName = "WIN-QRLAKF0WDIF" orelse Environment.MachineName = "TSRVNT170" Then
            Dim except2 As Exception = Nothing
            If m_ConnectionExceptionQueue.TryDequeue(except2) Then
                Throw except2
            End If
        End If
#End If

    End Sub


    Protected Overrides Sub Finalize()

        '在Finalizer時將未關閉的Connection放入Queue，在非Finalizer時關閉Connection
#If DEBUG Then
        If Not IsNothing(m_connection) Then

            Dim sDebug As String
            sDebug = "Connection 沒有關閉，建立連線時的CALL STACK內容：" & vbCrLf & m_sConstructorCallStack

            If System.Diagnostics.Debugger.IsAttached = True Then
                System.Diagnostics.Trace.WriteLine(sDebug)
                System.Diagnostics.Debugger.Break()         'Connection 沒有關閉
                'Throw New Exception(sDebug)
            End If

            'm_ConnectionQueue.Enqueue(m_connection)     '新增至Queue，等到非Finalizer時釋放

            '新增至Queue，等到非Finalizer時Throw Exception
            'm_ConnectionExceptionQueue.Enqueue(New Exception(sDebug))
        End If
#End If
        Try
            m_ConnectionQueue.Enqueue(m_connection)     '新增至Queue，等到非Finalizer時釋放
        Catch ex As Exception
        End Try
        'Me.releaseConnection()
    End Sub


    Sub init(ByVal sConnectionString As String, Optional ByVal iMaxConn As Integer = 1, Optional ByVal iMinConn As Integer = 0)
        m_sDataSource = sConnectionString
        setDataSource(m_sDataSource)
    End Sub

    Private Sub init(ByVal sDatabaseName As String, ByVal sUserName As String, ByVal sPassword As String, Optional ByVal iMaxConn As Integer = 1, Optional ByVal iMinConn As Integer = 0)
        m_sDataSource = "Provider=OraOLEDB.Oracle.1;Password=" & sPassword & ";User ID=" & sUserName & ";Data Source=" & sDatabaseName & ";MAX POOL SIZE=" & iMaxConn & ";MIN POOL SIZE=" & iMinConn & "; Pooling=True; Enlist=True"
        setDataSource(m_sDataSource)
    End Sub

    Private Sub setDataSource(ByVal sDataSource As String)
        m_sDataSource = sDataSource
    End Sub

    Public Function getDataSource() As String
        Return m_sDataSource
    End Function


    Private Function open(ByVal sConnectionString As String) As IDbConnection
        Dim strConn As String = sConnectionString
        setDataSource(strConn)
        Try
            If IsNothing(Me.m_connection) Then
                m_connection = ProviderFactory.CreateConnection
                m_connection.ConnectionString = strConn
            End If

            If (m_connection.State = ConnectionState.Closed) Then
                m_connection.Open()
            End If

            Return m_connection
        Catch oleEx As System.Runtime.InteropServices.ExternalException
            If Not IsNothing(m_connection) Then
                Me.releaseConnection()
            End If
            'Throw New Exception("ConnectionString=[" & strConn & "] <br>" & oleEx.Message)
            Throw
        Catch ex As Exception
            If Not IsNothing(m_connection) Then
                Me.releaseConnection()
            End If
            Throw
        Finally
            'do nothing
        End Try
    End Function

    Private Function open() As IDbConnection
        Dim strConn As String = Properties.getDataSource()
        setDataSource(strConn)
        Try
            If IsNothing(Me.m_connection) Then
                m_connection = ProviderFactory.CreateConnection
                m_connection.ConnectionString = strConn
            End If

            If (m_connection.State = ConnectionState.Closed) Then
                m_connection.Open()
            End If

            Return m_connection
        Catch oleEx As System.Runtime.InteropServices.ExternalException
            If Not IsNothing(m_connection) Then
                Me.releaseConnection()
            End If
            'Throw New Exception("ConnectionString=[" & strConn & "] <br>" & oleEx.Message)
            Throw
        Catch ex As Exception
            If Not IsNothing(m_connection) Then
                Me.releaseConnection()
            End If
            Throw
        Finally
            'do nothing
        End Try
    End Function

    Public Function getConnection() As IDbConnection
        Return Me.m_connection
    End Function

    Public Function getTransaction() As IDbTransaction
        Return Me.m_transaction
    End Function

    Function isTransaction() As Boolean
        Return m_bTransaction
    End Function

    Function isCloseTransaction() As Boolean
        Return m_bCloseTransaction
    End Function

    Function beginTran() As Boolean
        If IsNothing(m_connection) Then
            m_bTransaction = False
            Return Nothing
        End If
        Try
            If IsNothing(m_transaction) Then

                If (AVOID_TABLE_LOCKING) Then
                    m_transaction = m_connection.BeginTransaction(System.Data.IsolationLevel.ReadUncommitted)
                Else
                    m_transaction = m_connection.BeginTransaction()
                End If

                m_bCloseTransaction = False
            End If
            m_bTransaction = True

        Catch ex As Exception
            m_bTransaction = False

            Throw
        End Try
        Return m_bTransaction
    End Function

    Sub commit()
        If IsNothing(m_connection) Then
            Return
        End If
        Try
            If (m_bTransaction) Then
                m_bTransaction = False
                m_bCloseTransaction = True
                m_transaction.Commit()
                m_transaction = Nothing
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Sub Rollback()
        If IsNothing(m_connection) Then
            Return
        End If
        Try
            If (m_bTransaction) AndAlso Not (m_bCloseTransaction) Then
                m_bTransaction = False
                m_bCloseTransaction = True
                m_transaction.Rollback()
                m_transaction = Nothing
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Function getUseObject() As String
        Return m_sPath
    End Function

    Sub setUseObject(ByVal str As String)
        m_sPath = str
    End Sub

    Friend Sub setUse(ByVal isUse As Boolean)
        m_bUse = isUse

        If isUse Then
            If Not System.Web.HttpContext.Current Is Nothing Then
                m_sPath = System.Web.HttpContext.Current.Request.CurrentExecutionFilePath.ToString
            End If
        End If
    End Sub

    Function getUse() As Boolean
        Return m_bUse
    End Function

    Sub IncreaseWriteCount()

#If DEBUG Then
        If m_bTransaction = True OrElse TRANSACTION_WARNING = False Then
            Return
        End If

        m_nWriteCount = m_nWriteCount + 1

        If m_nWriteCount >= 2 Then
            Dim sDebug As String
            sDebug = "執行多個寫入時需使用TRANSACTION，或使用DatabaseManager.TRANSACTION_WARNING = False關閉檢查"

            System.Diagnostics.Debug.Assert(False, sDebug)

            If System.Diagnostics.Debugger.IsAttached = True Then
                System.Diagnostics.Debugger.Break()
            End If

            'Throw New Exception(sDebug)
        End If
#End If

    End Sub


    Sub closeConnection()
        If Not IsNothing(Me.m_connection) Then
            Me.m_connection.Close()
            'Me.Finalize()
            'Me.m_connection.Dispose() 
            'System.Data.IDbConnection.ReleaseObjectPool()        
            Me.m_connection = Nothing
        End If
    End Sub

    Public Shared Sub closePool()
        Try
            Oracle.DataAccess.Client.OracleConnection.ClearAllPools()
        Catch ex As Exception
        End Try
    End Sub

End Class
