Option Explicit On
Option Strict On

Imports System.Data
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

    Public Shared Function getInstance(ByVal sDatabaseName As String, ByVal sUserName As String, ByVal sPassword As String, Optional ByVal iMaxConn As Integer = 0) As DatabaseManager
        Dim dbManager As DatabaseManager = Nothing

        Try
            If IsNothing(dbManager) Then
                dbManager = newDatabase(sDatabaseName, sUserName, sPassword)
            End If

            Return dbManager

        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function getInstance(ByVal sConnectionString As String, Optional ByVal iMaxConn As Integer = 0) As DatabaseManager
        Dim dbManager As DatabaseManager = Nothing

        Try
           If IsNothing(dbManager) Then
                dbManager = newDatabase(sConnectionString)
            End If

            Return dbManager

        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function getInstance() As DatabaseManager

        Dim dbManager As DatabaseManager = Nothing
        Try
            If IsNothing(dbManager) Then
                dbManager = newDatabase()
            End If

            Return dbManager
        Catch ex As System.Threading.ThreadAbortException
            Console.Write(ex)
        Catch ex As Exception
            Throw New DataBaseManagerException(ex, CType(ex.TargetSite, Reflection.MethodInfo))
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
            Throw New DataBaseManagerException(ex, CType(ex.TargetSite, Reflection.MethodInfo))
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
            Throw New DataBaseManagerException(ex, CType(ex.TargetSite, Reflection.MethodInfo))
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
            Throw New DataBaseManagerException(ex, CType(ex.TargetSite, Reflection.MethodInfo))
        End Try
    End Function
#End Region

    Sub dispose() Implements System.IDisposable.Dispose
        Me.releaseConnection()
    End Sub

    Sub releaseConnection()
        If Me.m_bTransaction And Not Me.m_bCloseTransaction Then
            Me.Rollback()
            Throw New DataBaseManagerException("未關閉 Transaction ! 交易已經 Rollback !")
        End If

        closeConnection()
    End Sub

    Private Sub New()
    End Sub

    Private Sub New(ByVal sDatabaseName As String, ByVal sUserName As String, ByVal sPassword As String, ByVal iMaxConn As Integer)
        init(sDatabaseName, sUserName, sPassword, iMaxConn)
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
            Throw New Exception(oleEx.Message)
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
            Throw New Exception(oleEx.Message)
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
                m_transaction = m_connection.BeginTransaction()
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

    Sub closeConnection()
        If Not IsNothing(Me.m_connection) Then
            Me.m_connection.Close()
            Me.Finalize()
            'Me.m_connection.Dispose() 
            'System.Data.IDbConnection.ReleaseObjectPool()        
        End If
    End Sub

    Public Shared Sub closePool()
        Try
            If Properties.getProvider = ProviderType.SqlClient Then

            ElseIf Properties.getProvider = ProviderType.ODP Then
                Oracle.DataAccess.Client.OracleConnection.ClearAllPools()
            End If

        Catch ex As Exception
        End Try
    End Sub

#Region "預計廢除的Function"
    <ObsoleteAttribute("This function will be removed from future Versions.20110715")> _
    Private Function getSchemaName() As String
        'Return Titan.Utility.getString(m_sDataSource, "User ID=", ";")
        If com.Azion.NET.VB.Properties.getProvider = ProviderType.SqlClient Then
            Return FileUtil.getAppSettings("", "")
        ElseIf com.Azion.NET.VB.Properties.getProvider = ProviderType.ODP Then
            Return Titan.Utility.getString(m_sDataSource, "User ID=", ";")
        End If
        Return ""
    End Function
#End Region
End Class
