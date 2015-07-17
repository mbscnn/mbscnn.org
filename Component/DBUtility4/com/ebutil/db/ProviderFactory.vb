Option Explicit On
Option Strict On

' <summary>
' The collection of ADO.NET data providers that are supported by <see cref="ProviderFactory"/>.
' </summary>
Public Enum ProviderType

    ' <summary>
    ' The OLE DB (<see cref="System.Data.OleDb"/>) .NET data provider.
    ' </summary>
    OleDb = 0
    ' <summary>
    ' The SQL Server (<see cref="System.Data.SqlClient"/>) .NET data provider.
    ' </summary>
    SqlClient = 1
    ' <summary>
    ' The Oracle Server (<see cref="System.Data.OracleClient"/>) .NET data provider.
    ' </summary>
    OracleClient = 2
    ' <summary>
    ' The  Oracle Server (<see cref="Oracle.DataAccess.Client"/>) .ODP.NET provider. 
    ' </summary>
    ODP = 3
    ' <summary>
    ' The  IB2 DB2  .NET data provider.. 
    ' </summary>
    IBMDADB2 = 4

    ' <summary>
    ' The  MySql Server (<see cref="MySql.Data"/>) .MySql.NET provider. 
    ' </summary>
    MySqlClient = 5
End Enum 'ProviderType


' <summary>
' The <b>ProviderFactory</b> class abstracts ADO.NET relational data providers through creator methods which return
' the underlying <see cref="System.Data"/> interface.
' </summary>
' <remarks>
' </remarks>
Public NotInheritable Class ProviderFactory

    Friend Shared m_ProviderType As ProviderType = Properties.getProvider()

    Friend Shared CONNECTION_TYPES As Type = ProviderFactory.initConnectionClassType
    'Friend Shared COMMAND_TYPES As Type = ProviderFactory.initCommandClassType
    'Friend Shared DATA_ADAPTER_TYPES As Type = ProviderFactory.initDataAdapterClassType
    Friend Shared DATA_PARAMETER_TYPES As Type = ProviderFactory.initDataParameterClassType

    Friend Shared m_bDebug As Boolean = False


    Private Sub New()
    End Sub 'New
     
#Region "Property"



    Shared ReadOnly Property PositionPara() As String
        Get
            Select Case m_ProviderType
                Case ProviderType.ODP
                    Return ":" '«Ý¬d
                Case ProviderType.OleDb
                    Return ":"
                Case ProviderType.OracleClient
                    Return ":"
                Case ProviderType.SqlClient
                    Return "@"
                Case ProviderType.IBMDADB2
                    Return "?"
                Case ProviderType.MySqlClient
                    Return "?"
            End Select
            Return ""
        End Get
    End Property

#End Region

#Region "InitClassType"
    Private Shared Function initConnectionClassType() As Type
        Try
            Select Case m_ProviderType
                Case ProviderType.OleDb
                    Return GetType(System.Data.OleDb.OleDbConnection)
                Case ProviderType.SqlClient
                    Return GetType(System.Data.SqlClient.SqlConnection)
                Case ProviderType.OracleClient
                    Return GetType(System.Data.OracleClient.OracleConnection)
                Case ProviderType.ODP
                    Return GetType(Oracle.DataAccess.Client.OracleConnection)
                Case ProviderType.MySqlClient
                    Return GetType(MySql.Data.MySqlClient.MySqlConnection)
                Case Else
                    Throw New ProviderException("ProviderFactory.initConnectionClassType ¿ù»~!")
            End Select
        Catch ex As Exception

            Throw New ProviderException(ex, CType(ex.TargetSite, Reflection.MethodInfo))
        End Try
    End Function

    Private Shared Function initCommandClassType() As Type
        Try
            Select Case m_ProviderType
                Case ProviderType.OleDb
                    Return GetType(System.Data.OleDb.OleDbCommand)
                Case ProviderType.SqlClient
                    Return GetType(System.Data.SqlClient.SqlCommand)
                Case ProviderType.OracleClient
                    Return GetType(System.Data.OracleClient.OracleCommand)
                Case ProviderType.ODP
                    Return GetType(Oracle.DataAccess.Client.OracleCommand)
                Case ProviderType.MySqlClient
                    Return GetType(MySql.Data.MySqlClient.MySqlCommand)
                Case Else
                    Throw New Exception("ProviderFactory.initCommandClassType ¿ù»~!")
            End Select
        Catch ex As Exception

            Throw New ProviderException(ex, CType(ex.TargetSite, Reflection.MethodInfo))
        End Try
    End Function

    Private Shared Function initDataAdapterClassType() As Type
        Try
            Select Case m_ProviderType
                Case ProviderType.OleDb
                    Return GetType(System.Data.OleDb.OleDbDataAdapter)
                Case ProviderType.SqlClient
                    Return GetType(System.Data.SqlClient.SqlDataAdapter)
                Case ProviderType.OracleClient
                    Return GetType(System.Data.OracleClient.OracleDataAdapter)
                Case ProviderType.ODP
                    Return GetType(Oracle.DataAccess.Client.OracleDataAdapter)
                Case ProviderType.MySqlClient
                    Return GetType(MySql.Data.MySqlClient.MySqlDataAdapter)
                Case Else
                    Throw New Exception("ProviderFactory.initDataAdapterClassType ¿ù»~!")
            End Select
        Catch ex As Exception

            Throw New ProviderException(ex, CType(ex.TargetSite, Reflection.MethodInfo))
        End Try
    End Function

    Private Shared Function initDataParameterClassType() As Type
        Try
            Select Case m_ProviderType
                Case ProviderType.OleDb
                    Return GetType(System.Data.OleDb.OleDbParameter)
                Case ProviderType.SqlClient
                    Return GetType(System.Data.SqlClient.SqlParameter)
                Case ProviderType.OracleClient
                    Return GetType(System.Data.OracleClient.OracleParameter)
                Case ProviderType.ODP
                    Return GetType(Oracle.DataAccess.Client.OracleParameter)
                Case ProviderType.MySqlClient
                    Return GetType(MySql.Data.MySqlClient.MySqlParameter)
                Case Else
                    Throw New Exception("ProviderFactory.initDataParameterClassType ¿ù»~!")
            End Select
        Catch ex As Exception

            Throw New ProviderException(ex, CType(ex.TargetSite, Reflection.MethodInfo))
        End Try
    End Function

#End Region

#Region "CreateConnection"

    Public Overloads Shared Function CreateConnection() As System.Data.IDbConnection
        Dim conn As System.Data.IDbConnection = Nothing

        Try
            conn = System.Data.Common.DbProviderFactories.GetFactory(CONNECTION_TYPES.Namespace).CreateConnection 'CType(Activator.CreateInstance(CONNECTION_TYPES), System.Data.IDbConnection)
        Catch ex As Exception

            Throw New ProviderException(ex, CType(ex.TargetSite, Reflection.MethodInfo))
        End Try

        Return conn
    End Function 'CreateConnection

    Public Overloads Shared Function CreateConnection(ByVal connectionString As String) As System.Data.IDbConnection
        Dim conn As System.Data.IDbConnection = Nothing
        Dim args As Object() = {connectionString}

        Try
            conn = System.Data.Common.DbProviderFactories.GetFactory(CONNECTION_TYPES.Namespace).CreateConnection ' CType(Activator.CreateInstance(CONNECTION_TYPES, args), System.Data.IDbConnection)
            conn.ConnectionString = connectionString
        Catch ex As Exception

            Throw New ProviderException(ex, CType(ex.TargetSite, Reflection.MethodInfo))
        End Try

        Return conn
    End Function 'CreateConnection
#End Region

#Region "CreateCommand"

    Public Overloads Shared Function CreateCommand() As System.Data.Common.DbCommand 'System.Data.IDbCommand
        Dim cmd As System.Data.Common.DbCommand 'System.Data.IDbCommand = Nothing

        Try
            cmd = System.Data.Common.DbProviderFactories.GetFactory(CONNECTION_TYPES.Namespace).CreateCommand 'CType(Activator.CreateInstance(COMMAND_TYPES), System.Data.IDbCommand)

            If TypeOf (cmd) Is Oracle.DataAccess.Client.OracleCommand Then
                CType(cmd, Oracle.DataAccess.Client.OracleCommand).InitialLONGFetchSize = -1
                CType(cmd, Oracle.DataAccess.Client.OracleCommand).InitialLOBFetchSize = -1
            End If
        Catch ex As Exception

            Throw New ProviderException(ex, CType(ex.TargetSite, Reflection.MethodInfo))
        End Try

        Return cmd
    End Function 'CreateCommand

    Public Overloads Shared Function CreateCommand(ByVal cmdText As String) As System.Data.Common.DbCommand 'System.Data.IDbCommand
        Dim cmd As System.Data.Common.DbCommand 'System.Data.IDbCommand = Nothing
        Dim args As Object() = {cmdText}

        Try
            cmd = System.Data.Common.DbProviderFactories.GetFactory(CONNECTION_TYPES.Namespace).CreateCommand 'CType(Activator.CreateInstance(COMMAND_TYPES, args), System.Data.IDbCommand)
            cmd.CommandText = cmdText
            If TypeOf (cmd) Is Oracle.DataAccess.Client.OracleCommand Then
                CType(cmd, Oracle.DataAccess.Client.OracleCommand).InitialLONGFetchSize = -1
                CType(cmd, Oracle.DataAccess.Client.OracleCommand).InitialLOBFetchSize = -1
            End If
        Catch ex As Exception

            Throw New ProviderException(ex, CType(ex.TargetSite, Reflection.MethodInfo))
        End Try
        Return cmd
    End Function 'CreateCommand

    Public Overloads Shared Function CreateCommand(ByVal cmdText As String, ByVal connection As System.Data.IDbConnection) As System.Data.Common.DbCommand 'System.Data.IDbCommand
        Dim cmd As System.Data.Common.DbCommand 'System.Data.IDbCommand = Nothing
        Dim args As Object() = {cmdText, connection}

        Try
            cmd = System.Data.Common.DbProviderFactories.GetFactory(CONNECTION_TYPES.Namespace).CreateCommand ' CType(Activator.CreateInstance(COMMAND_TYPES, args), System.Data.IDbCommand)
            cmd.CommandText = cmdText
            cmd.Connection = CType(connection, System.Data.Common.DbConnection)
            If TypeOf (cmd) Is Oracle.DataAccess.Client.OracleCommand Then
                CType(cmd, Oracle.DataAccess.Client.OracleCommand).InitialLONGFetchSize = -1
                CType(cmd, Oracle.DataAccess.Client.OracleCommand).InitialLOBFetchSize = -1
            End If
        Catch ex As Exception

            Throw New ProviderException(ex, CType(ex.TargetSite, Reflection.MethodInfo))
        End Try

        Return cmd
    End Function 'CreateCommand

    Public Overloads Shared Function CreateCommand(ByVal cmdText As String, ByVal connection As System.Data.IDbConnection, ByVal transaction As System.Data.IDbTransaction) As System.Data.Common.DbCommand 'System.Data.IDbCommand
        Dim cmd As System.Data.Common.DbCommand 'System.Data.IDbCommand = Nothing
        Dim args As Object() = {cmdText, connection, transaction}

        Try
            cmd = System.Data.Common.DbProviderFactories.GetFactory(CONNECTION_TYPES.Namespace).CreateCommand ' CType(Activator.CreateInstance(COMMAND_TYPES, args), System.Data.IDbCommand)
            cmd.CommandText = cmdText
            cmd.Connection = CType(connection, System.Data.Common.DbConnection)
            cmd.Transaction = CType(transaction, System.Data.Common.DbTransaction)
            If TypeOf (cmd) Is Oracle.DataAccess.Client.OracleCommand Then
                CType(cmd, Oracle.DataAccess.Client.OracleCommand).InitialLONGFetchSize = -1
                CType(cmd, Oracle.DataAccess.Client.OracleCommand).InitialLOBFetchSize = -1
            End If
        Catch ex As Exception

            Throw New ProviderException(ex, CType(ex.TargetSite, Reflection.MethodInfo))
        End Try
        Return cmd
    End Function 'CreateCommand
#End Region

#Region "CreateDataAdapter"

    Public Overloads Shared Function CreateDataAdapter() As System.Data.Common.DbDataAdapter
        Dim da As System.Data.Common.DbDataAdapter = Nothing

        Try
            'da = CType(Activator.CreateInstance(DATA_ADAPTER_TYPES), System.Data.Common.DbDataAdapter)
            da = System.Data.Common.DbProviderFactories.GetFactory(CONNECTION_TYPES.Namespace).CreateDataAdapter
        Catch ex As Exception

            Throw New ProviderException(ex, CType(ex.TargetSite, Reflection.MethodInfo))
        End Try
        Return da
    End Function 'CreateDataAdapter

    Public Overloads Shared Function CreateDataAdapter(ByVal selectCommand As IDbCommand) As System.Data.Common.DbDataAdapter
        Dim da As System.Data.Common.DbDataAdapter = Nothing
        'Dim args As Object() = {selectCommand}

        Try
            'da = CType(Activator.CreateInstance(DATA_ADAPTER_TYPES, args), System.Data.Common.DbDataAdapter)
            da = ProviderFactory.CreateDataAdapter()
            da.SelectCommand = CType(selectCommand, Common.DbCommand)
        Catch ex As Exception

            Throw New ProviderException(ex, CType(ex.TargetSite, Reflection.MethodInfo))
        End Try

        Return da
    End Function 'CreateDataAdapter

    Public Overloads Shared Function CreateDataAdapter(ByVal selectCommandText As String, ByVal connection As System.Data.IDbConnection) As System.Data.Common.DbDataAdapter
        Dim da As System.Data.Common.DbDataAdapter = Nothing
        'Dim args As Object() = {selectCommandText, selectConnection}

        Try
            'da = CType(Activator.CreateInstance(DATA_ADAPTER_TYPES, args), System.Data.Common.DbDataAdapter)
            da = ProviderFactory.CreateDataAdapter()
            da.SelectCommand = CType(connection.CreateCommand, Common.DbCommand)
            da.SelectCommand.CommandText = selectCommandText
        Catch ex As Exception

            Throw New ProviderException(ex, CType(ex.TargetSite, Reflection.MethodInfo))
        End Try

        Return da
    End Function 'CreateDataAdapter

    Public Overloads Shared Function CreateDataAdapter(ByVal selectCommandText As String, ByVal connectionString As String) As System.Data.Common.DbDataAdapter
        Dim da As System.Data.Common.DbDataAdapter = Nothing
        'Dim args As Object() = {selectCommandText, selectConnectionString}
        Try
            Dim connection As System.Data.IDbConnection = ProviderFactory.CreateConnection(connectionString)
            'da = CType(Activator.CreateInstance(DATA_ADAPTER_TYPES, args), System.Data.Common.DbDataAdapter)
            da = ProviderFactory.CreateDataAdapter(selectCommandText, connection)
        Catch ex As Exception

            Throw New ProviderException(ex, CType(ex.TargetSite, Reflection.MethodInfo))
        End Try

        Return da
    End Function 'CreateDataAdapter
#End Region

#Region "CreateCommandBuilder"

    Public Overloads Shared Sub CommandBuilderDeriveParameters(ByVal iDbCommand As IDbCommand)

        If m_ProviderType = ProviderType.OleDb Then
            System.Data.OleDb.OleDbCommandBuilder.DeriveParameters(CType(iDbCommand, OleDb.OleDbCommand))
        ElseIf m_ProviderType = ProviderType.SqlClient Then
            System.Data.SqlClient.SqlCommandBuilder.DeriveParameters(CType(iDbCommand, SqlClient.SqlCommand))
        ElseIf m_ProviderType = ProviderType.ODP Then
            Oracle.DataAccess.Client.OracleCommandBuilder.DeriveParameters(CType(iDbCommand, Oracle.DataAccess.Client.OracleCommand))
        ElseIf m_ProviderType = ProviderType.OracleClient Then
            System.Data.OracleClient.OracleCommandBuilder.DeriveParameters(CType(iDbCommand, OracleClient.OracleCommand))
        ElseIf m_ProviderType = ProviderType.MySqlClient Then
            MySql.Data.MySqlClient.MySqlCommandBuilder.DeriveParameters(CType(iDbCommand, MySql.Data.MySqlClient.MySqlCommand))
        End If
    End Sub

    Private Overloads Shared Function CreateCommandBuilder() As System.Data.Common.DbCommandBuilder

        Try
            Return System.Data.Common.DbProviderFactories.GetFactory(CONNECTION_TYPES.Namespace).CreateCommandBuilder()
        Catch ex As Exception

            Throw New ProviderException(ex, CType(ex.TargetSite, Reflection.MethodInfo))
        End Try

        Return Nothing
    End Function

    Private Overloads Shared Function CreateCommandBuilder(ByVal da As System.Data.Common.DbDataAdapter) As System.Data.Common.DbCommandBuilder
        Dim cb As System.Data.Common.DbCommandBuilder = Nothing
        Try

            cb = ProviderFactory.CreateCommandBuilder()
            cb.DataAdapter = da

        Catch ex As Exception

            Throw New ProviderException(ex, CType(ex.TargetSite, Reflection.MethodInfo))
        End Try
        Return cb
    End Function
#End Region

#Region "CreateDataParameter"


    Public Overloads Shared Function CreateDataParameter() As System.Data.IDbDataParameter
        Dim param As System.Data.IDbDataParameter = Nothing

        Try
            'param = CType(Activator.CreateInstance(DATA_PARAMETER_TYPES), System.Data.IDbDataParameter)
            param = System.Data.Common.DbProviderFactories.GetFactory(CONNECTION_TYPES.Namespace).CreateParameter
        Catch ex As Exception

            Throw New ProviderException(ex, CType(ex.TargetSite, Reflection.MethodInfo))
        End Try

        Return param
    End Function 'CreateDataParameter

    Public Overloads Shared Function CreateDataParameter(ByVal parameterName As String, ByVal value As Object) As System.Data.IDbDataParameter
        Dim param As System.Data.IDbDataParameter = Nothing
        Dim args As Object() = {parameterName, value}

        Try
            If args(1) Is Nothing Then
                args(1) = DBNull.Value
            End If
            param = CType(Activator.CreateInstance(DATA_PARAMETER_TYPES, args), System.Data.IDbDataParameter)

        Catch ex As Exception

            Throw New ProviderException(ex, CType(ex.TargetSite, Reflection.MethodInfo))
        End Try

        Return param
    End Function 'CreateDataParameter

    Public Overloads Shared Function CreateDataParameter(ByVal parameterName As String, ByVal dbType As System.Data.DbType, ByVal value As Object) As System.Data.IDbDataParameter
        Dim param As System.Data.IDbDataParameter = CreateDataParameter(parameterName, dbType)

        Try
            If Not (param Is Nothing) Then
                param.Value = value
            End If
        Catch ex As Exception

            Throw New ProviderException(ex, CType(ex.TargetSite, Reflection.MethodInfo))
        End Try

        Return param
    End Function 'CreateDataParameter

    Public Overloads Shared Function CreateDataParameter(ByVal parameterName As String, ByVal dbType As System.Data.DbType) As System.Data.IDbDataParameter
        Dim param As System.Data.IDbDataParameter = CreateDataParameter()

        If Not (param Is Nothing) Then
            param.ParameterName = parameterName
            param.DbType = dbType
        End If

        Return param
    End Function 'CreateDataParameter

    'Public Overloads Shared Function CreateDataParameter(ByVal parameterName As String, ByVal dbType As System.Data.DbType, ByVal size As Integer) As System.Data.IDbDataParameter
    '    Dim param As System.Data.IDbDataParameter = CreateDataParameter()

    '    If Not (param Is Nothing) Then
    '        param.ParameterName = parameterName
    '        param.DbType = dbType
    '        param.Size = size
    '    End If

    '    Return param
    'End Function 'CreateDataParameter

    Public Overloads Shared Function CreateDataParameter(ByVal parameterName As String, ByVal dbType As System.Data.DbType, ByVal size As Integer, ByVal sourceColumn As String) As System.Data.IDbDataParameter
        Dim param As System.Data.IDbDataParameter = CreateDataParameter()

        If Not (param Is Nothing) Then
            param.ParameterName = parameterName
            param.DbType = dbType
            param.Size = size
            param.SourceColumn = sourceColumn
        End If

        Return param
    End Function 'CreateDataParameter

#End Region

    Public Shared Function getFieldType(ByVal iType As Integer) As String
        If m_ProviderType = ProviderType.OleDb Then
            Return DbUtility.getOLETypeName(iType).ToString
        ElseIf m_ProviderType = ProviderType.ODP Then
            Return "Oracle.DataAccess.Client.OracleDbType." & DbUtility.getODPTypeName(iType)
        ElseIf m_ProviderType = ProviderType.OracleClient Then
            Return DbUtility.getOracleTypeName(iType)
        ElseIf m_ProviderType = ProviderType.SqlClient Then
            Return "System.Data.SqlDbType." & DbUtility.getSqlTypeName(iType)
        ElseIf m_ProviderType = ProviderType.MySqlClient Then
            Return DbUtility.getMySqlTypeName(iType)
        End If
        Return Nothing
    End Function

End Class 'ProviderFactory '

