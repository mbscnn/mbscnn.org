Imports System.Collections

Imports System.Web
Imports System.Web.SessionState
Public MustInherit Class Properties
    'Inherits HttpHandler
    'Web 一定要設為 False
    'Public Shared bConsoleMode As Boolean = False
    'Private Shared m_sConfiguration As String = String.Empty  '= "C:\DBProfile.txt"

    Private Shared m_bConsoleMode As Boolean = False
    Public Shared Property bConsoleMode As Boolean
        Get
            Return m_bConsoleMode
        End Get
        Set(ByVal value As Boolean)
            m_bConsoleMode = value
        End Set
    End Property

    Private Shared m_sConfiguration As String = String.Empty  '= "C:\DBProfile.txt"
    Public Shared ReadOnly Property sConfiguration As String
        Get
            Return m_sConfiguration
        End Get
    End Property

#Region "inner function"

    <System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.Synchronized)> _
    Public Shared Sub setConfiguration(Optional ByVal sConfiguration As String = "C:\MSBCConf\MSBC.config")
        'SyncLock m_sConfiguration
        m_sConfiguration = sConfiguration
        'If IsNothing(m_props) Then
        '    m_props = Hashtable.Synchronized(New Hashtable)
        '    getProperties(getDBProfilePath())
        'End If
        'End SyncLock
    End Sub

    Protected Friend Shared Function getString(ByVal sKey As String) As String
        Try
            Dim sValue As String = com.Azion.EloanUtility.FileUtility.getAppSettings(sKey, getDBProfilePath())
            If com.Azion.EloanUtility.ValidateUtility.isValidateData(sValue) Then
                Return sValue
            Else
                Throw New Exception
            End If
        Catch ex As Exception
            Throw New ProviderException("無法取得參數[" & sKey & "]的值")
        End Try
    End Function

    Private Shared Function getDBProfilePath() As String
        Try
            If m_sConfiguration = Nothing Then
                'setConfiguration()
                Throw New ProviderException("無法讀取設定檔，請使用Properties.setConfiguration(""路徑"")設定?:[ex:C:\MSBCConf\MSBC.config]")
            End If

            If Not System.IO.File.Exists(m_sConfiguration) Then
                Throw New ProviderException("config is not found:" & m_sConfiguration)
            End If
            Return m_sConfiguration
        Catch ex As Exception

            Throw New ProviderException(ex, Reflection.MethodBase.GetCurrentMethod)
        End Try
    End Function

#End Region

#Region "取得開啟資料庫的字串"

    Public Shared Function getDSN() As String

        ' If Properties.bConsoleMode OrElse System.Web.HttpContext.Current Is Nothing Then
        If System.Web.HttpContext.Current Is Nothing Then
            Dim oProviderType As ProviderType = Properties.getProvider()
            If oProviderType = ProviderType.OleDb Then
                Return "Provider=OraOleDB.Oracle;Pooling=false;Connection Lifetime=90;Password=" & getString("DBPassword") & ";User ID=" & getString("DBUserID").ToUpper & ";Data Source=" & getString("DBSource").ToUpper
            ElseIf oProviderType = ProviderType.SqlClient Then
                'Return "Data Source=" & getString("DBSource") & ";database=" & getString("DataBase") & ";User ID=" & getString("DBUserID") & ";Password=" & getString("DBPassword") & ";Persist Security Info=True;Pooling=true;Min Pool Size=" & getString("MinPool") & ";Max Pool Size=" & getString("MaxPool") & ";Connection Lifetime=" & getString("ConnectionLifetime") & ";Connect Timeout=" & getString("ConnectTimeout") & ";Application Name=" & getString("AppName")
                Return "Data Source=" & getString("DBSource") & ";database=" & getString("DataBase") & ";User ID=" & getString("DBUserID") & ";Persist Security Info=True;Pooling=true;Min Pool Size=" & getString("MinPool") & ";Max Pool Size=" & getString("MaxPool") & ";Connection Lifetime=" & getString("ConnectionLifetime") & ";Connect Timeout=" & getString("ConnectTimeout") & ";Application Name=" & getString("AppName") & ";Password=" & com.Azion.EloanUtility.EncryptUtility.Decrypto(getString("DBPassword"))
            ElseIf oProviderType = ProviderType.OracleClient Then
                Return "Max Pool Size=" & getMaxSize() & ";Min Pool Size=" & getMinSize() & ";Pooling=true;Connection Lifetime=120;Connection Timeout=60;Password=" & getString("DBPassword") & ";User ID=" & getString("DBUserID").ToUpper & ";Data Source=" & getString("DBSource").ToUpper
            ElseIf oProviderType = ProviderType.ODP Then
                Return "Max Pool Size=" & getMaxSize() & ";Min Pool Size=" & getMinSize() & ";Pooling=true;Connection Lifetime=120;Connection Timeout=60;Enlist=false;User ID=" & getString("DBUserID").ToUpper & ";Password=" & getString("DBPassword") & ";Data Source=" & getString("DBSource").ToUpper '& ";Statement Cache Size=20"   'in ODP.NET version 10.1.0.3. 
            ElseIf oProviderType = ProviderType.IBMDADB2 Then
                Return "Provider=IBMDADB2;Database=SFA;Hostname=10.8.210.40;Protocol=TCPIP; Port=60001;Uid=sfausr;Pwd=sfausr;"
            ElseIf oProviderType = ProviderType.MySqlClient Then
                Return "server=" & getString("DBSource") & ";uid=" & getString("DBUserID") & ";pwd=" & getString("DBPassword") & ";database=" & getString("SchemaName") & ";CharSet=utf8;"
            End If
        End If
        Throw New ProviderException("connection string is null")
    End Function

    Public Shared Function getDataSource() As String

        'If Properties.bConsoleMode OrElse System.Web.HttpContext.Current Is Nothing Then
        If System.Web.HttpContext.Current Is Nothing Then
            Return getDSN()
        Else
            Return System.Web.HttpContext.Current.Application.Item("BotDSN")
        End If
    End Function

    Shared Function getSchemaName() As String 
        If Properties.getProvider() = ProviderType.SqlClient OrElse Properties.getProvider() = ProviderType.MySqlClient Then
            Return getString("SchemaName")
        End If
        Return getString("DBUserID")
    End Function

    Private Shared Function getDBPassword() As String
        Return com.Azion.EloanUtility.EncryptUtility.Decrypto(getString("DBPassword"))
    End Function

    Shared Function getDBSource() As String
        Return getString("DBSource")
    End Function

    Shared Function getWebAppSettings(ByVal sKey As String) As String
        Dim app As Object = System.Web.Configuration.WebConfigurationManager.GetWebApplicationSection("appSettings")
        If Not app Is Nothing Then
            Return app.Item(sKey)
        End If

        Return "" 
    End Function
#End Region

#Region "pooling" 
    Shared Function getProvider() As ProviderType
         
        Dim sProvider As String = getString("Provider") 'CType(m_props.Item("Provider"), String)
        If sProvider.ToUpper.Equals("OLEDB") Then
            Return ProviderType.OleDb
        ElseIf sProvider.ToUpper.Equals("ODP") Then
            Return ProviderType.ODP
        ElseIf sProvider.ToUpper.Equals("ORACLECLIENT") Then
            Return ProviderType.OracleClient
        ElseIf sProvider.ToUpper.Equals("SQLCLIENT") Then
            Return ProviderType.SqlClient
        ElseIf sProvider.ToUpper.Equals("IBMDADB2") Then
            Return ProviderType.IBMDADB2
        ElseIf sProvider.ToUpper.Equals("MYSQLCLIENT") Then
            Return ProviderType.MySqlClient
        End If
    End Function

    Shared Function getMinSize() As Integer

        Dim sMinSize As String = getString("MinSize")

        If Titan.Utility.isValidateData(sMinSize) Then
            Return CType(sMinSize, Integer)
        End If
        Return 5
    End Function

    Shared Function getMaxSize() As Integer
 
        Dim sMaxSize As String = getString("MaxSize")

        If Titan.Utility.isValidateData(sMaxSize) Then
            Return CType(sMaxSize, Integer)
        End If

        Return 10
    End Function

    Shared Function getMaxUseIdelTime() As Double
         
        Dim sMAX_USE_IDLE_TIME As String = getString("MAX_USE_IDLE_TIME")
        If Titan.Utility.isValidateData(sMAX_USE_IDLE_TIME) Then
            Return CType(sMAX_USE_IDLE_TIME, Double)
        End If
        Return 3
    End Function

    Shared Function getMonitorUseIdelTime() As Double
         
        Dim sMONITOR_USE_IDLE_TIME As String = getString("MONITOR_USE_IDLE_TIME")
        If Titan.Utility.isValidateData(sMONITOR_USE_IDLE_TIME) Then
            Return CType(sMONITOR_USE_IDLE_TIME, Double)
        End If
        Return 1
    End Function

    Shared Function getMaxPoolIdelTime() As Double
         
        Dim sMAX_POOL_IDLE_TIME As String = getString("MAX_POOL_IDLE_TIME")
        If Titan.Utility.isValidateData(sMAX_POOL_IDLE_TIME) Then
            Return CType(sMAX_POOL_IDLE_TIME, Double)
        End If
        Return 2
    End Function
     
    Shared Function isMonitorPool() As Boolean
        'Dim bMonitor As Boolean = False
         
        'Dim sMONITORPOOL As String = getString("MONITORPOOL")
        'If Titan.Utility.isValidateData(sMONITORPOOL) Then
        '    Return CType(sMONITORPOOL, Boolean)
        'End If
        'Return bMonitor

        Return False
    End Function
#End Region

End Class
