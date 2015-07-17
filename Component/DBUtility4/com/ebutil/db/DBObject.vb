Option Explicit On
Option Strict On

Imports System.Data
Imports System.Collections

Public Class DBObject
    'Inherits properties

#Region "ExecuteNonQuery"

    Protected Overloads Shared Function ExecuteNonQuery(ByVal connectionString As String, _
                                                   ByVal commandType As System.Data.CommandType, _
                                                   ByVal commandText As String) As Integer
        ' Pass through the call providing null for the set of OleDbParameters
        Return ExecuteNonQuery(connectionString, commandType, commandText, CType(Nothing, IDataParameter()))
    End Function ' ExecuteNonQuery

    Protected Overloads Shared Function ExecuteNonQuery(ByVal connectionString As String, _
                                                     ByVal commandType As System.Data.CommandType, _
                                                     ByVal commandText As String, _
                                                     ByVal ParamArray commandParameters() As IDataParameter) As Integer
        If (connectionString Is Nothing OrElse connectionString.Length = 0) Then Throw New ArgumentNullException("connectionString")
        ' Create & open a OleDbConnection, and dispose of it after we are done
        Dim connection As IDbConnection = Nothing
        Try
            connection = ProviderFactory.CreateConnection(connectionString)
            connection.Open()

            ' Call the overload that takes a connection in place of the connection string
            Return ExecuteNonQuery(connection, commandType, commandText, commandParameters)
        Finally
            If Not connection Is Nothing Then connection.Close()
        End Try
    End Function ' ExecuteNonQuery

    Public Overloads Shared Function ExecuteNonQuery(ByVal dbManager As DatabaseManager, _
                                              ByVal commandType As System.Data.CommandType, _
                                              ByVal commandText As String, _
    ByVal ParamArray commandParameters() As IDataParameter) As Integer
        If (dbManager.isTransaction) Then
            Return ExecuteNonQuery(dbManager.getTransaction, commandType, commandText, commandParameters)
        Else
            Return ExecuteNonQuery(dbManager.getConnection, commandType, commandText, commandParameters)
        End If
    End Function ' ExecuteNonQuery

    Public Overloads Shared Function ExecuteNonQuery(ByVal connection As IDbConnection, _
                                                 ByVal commandType As CommandType, _
                                                 ByVal commandText As String, _
                                                 ByVal ParamArray commandParameters() As IDataParameter) As Integer

        If (connection Is Nothing) Then Throw New ArgumentNullException("connection")

        ' Create a command and prepare it for execution
        Dim cmd As IDbCommand = ProviderFactory.CreateCommand
        Dim retval As Integer
        Dim mustCloseConnection As Boolean = False

        If TypeOf (cmd) Is Oracle.DataAccess.Client.OracleCommand Then
            'CType(cmd, Oracle.DataAccess.Client.OracleCommand).AddRowid = True
            CType(cmd, Oracle.DataAccess.Client.OracleCommand).InitialLONGFetchSize = -1
            CType(cmd, Oracle.DataAccess.Client.OracleCommand).InitialLOBFetchSize = -1
        End If

        PrepareCommand(cmd, connection, CType(Nothing, IDbTransaction), commandType, commandText, commandParameters, mustCloseConnection)

        ' Finally, execute the command
        retval = cmd.ExecuteNonQuery()

        ' Detach the OleDbParameters from the command object, so they can be used again
        cmd.Parameters.Clear()

        If (mustCloseConnection) Then connection.Close()

        Return retval
    End Function ' ExecuteNonQuery

    ' Execute a OleDbCommand (that returns no resultset) against the specified OleDbTransaction
    ' using the provided parameters.
    ' e.g.:  
    ' Dim result As Integer = ExecuteNonQuery(trans, CommandType.StoredProcedure, "GetOrders", new OleDbParameter("@prodid", 24))
    ' Parameters:
    ' -transaction - a valid OleDbTransaction 
    ' -commandType - the CommandType (stored procedure, text, etc.) 
    ' -commandText - the stored procedure name or T-OleDb command 
    ' -commandParameters - an array of OleDbParamters used to execute the command 
    ' Returns: An int representing the number of rows affected by the command 
    Public Overloads Shared Function ExecuteNonQuery(ByVal transaction As IDbTransaction, _
                                                     ByVal commandType As CommandType, _
                                                     ByVal commandText As String, _
                                                     ByVal ParamArray commandParameters() As IDataParameter) As Integer

        If (transaction Is Nothing) Then Throw New ArgumentNullException("transaction")
        If Not (transaction Is Nothing) AndAlso (transaction.Connection Is Nothing) Then Throw New ArgumentException("The transaction was rollbacked or commited, please provide an open transaction.", "transaction")

        ' Create a command and prepare it for execution
        Dim cmd As IDbCommand = ProviderFactory.CreateCommand
        Dim retval As Integer
        Dim mustCloseConnection As Boolean = False

        If TypeOf (cmd) Is Oracle.DataAccess.Client.OracleCommand Then
            'CType(cmd, Oracle.DataAccess.Client.OracleCommand).AddRowid = True
            CType(cmd, Oracle.DataAccess.Client.OracleCommand).InitialLONGFetchSize = -1
            CType(cmd, Oracle.DataAccess.Client.OracleCommand).InitialLOBFetchSize = -1
        End If

        PrepareCommand(cmd, transaction.Connection, transaction, commandType, commandText, commandParameters, mustCloseConnection)

        ' Finally, execute the command
        retval = cmd.ExecuteNonQuery()

        ' Detach the OleDbParameters from the command object, so they can be used again
        cmd.Parameters.Clear()

        Return retval
    End Function ' ExecuteNonQuery



#End Region

#Region "ExecuteDataset"


    Public Overloads Shared Function ExecuteDataset(ByVal dbManager As DatabaseManager, _
                                             ByVal commandType As System.Data.CommandType, _
                                             ByVal commandText As String, _
   ByVal ParamArray commandParameters() As IDataParameter) As DataSet
        If (dbManager.isTransaction) Then
            Return ExecuteDataset(dbManager.getTransaction, commandType, commandText, commandParameters)
        Else
            Return ExecuteDataset(dbManager.getConnection, commandType, commandText, commandParameters)
        End If
    End Function ' ExecuteNonQuery


    ' Execute a OleDbCommand (that returns a resultset) against the specified OleDbTransaction
    ' using the provided parameters.
    ' e.g.:  
    ' Dim ds As Dataset = ExecuteDataset(trans, CommandType.StoredProcedure, "GetOrders", new OleDbParameter("@prodid", 24))
    ' Parameters
    ' -transaction - a valid OleDbTransaction 
    ' -commandType - the CommandType (stored procedure, text, etc.)
    ' -commandText - the stored procedure name or T-OleDb command
    ' -commandParameters - an array of OleDbParamters used to execute the command
    ' Returns: A dataset containing the resultset generated by the command
    Public Overloads Shared Function ExecuteDataset(ByVal transaction As IDbTransaction, _
                                                    ByVal commandType As CommandType, _
                                                    ByVal commandText As String, _
                                                    ByVal ParamArray commandParameters() As IDataParameter) As DataSet
        If (transaction Is Nothing) Then Throw New ArgumentNullException("transaction")
        If Not (transaction Is Nothing) AndAlso (transaction.Connection Is Nothing) Then Throw New ArgumentException("The transaction was rollbacked or commited, please provide an open transaction.", "transaction")

        ' Create a command and prepare it for execution
        Dim cmd As IDbCommand = ProviderFactory.CreateCommand
        Dim ds As New DataSet
        Dim dataAdatpter As System.Data.Common.DataAdapter = Nothing
        Dim mustCloseConnection As Boolean = False

        If TypeOf (cmd) Is Oracle.DataAccess.Client.OracleCommand Then
            'CType(cmd, Oracle.DataAccess.Client.OracleCommand).AddRowid = True
            CType(cmd, Oracle.DataAccess.Client.OracleCommand).InitialLONGFetchSize = -1
            CType(cmd, Oracle.DataAccess.Client.OracleCommand).InitialLOBFetchSize = -1
        End If

        PrepareCommand(cmd, transaction.Connection, transaction, commandType, commandText, commandParameters, mustCloseConnection)

        Try
            ' Create the DataAdapter & DataSet
            dataAdatpter = ProviderFactory.CreateDataAdapter(cmd)
            'dataAdatpter.MissingMappingAction = MissingMappingAction.Passthrough
            'dataAdatpter.MissingSchemaAction = MissingSchemaAction.AddWithKey
            'ds.EnforceConstraints = False

            ' Fill the DataSet using default values for DataTable names, etc
            dataAdatpter.Fill(ds)

            ' Detach the OleDbParameters from the command object, so they can be used again
            cmd.Parameters.Clear()
        Finally
            If (Not dataAdatpter Is Nothing) Then dataAdatpter.Dispose()
        End Try

        ' Return the dataset
        Return ds

    End Function ' ExecuteDataset

    Public Overloads Shared Function ExecuteDataset(ByVal connection As IDbConnection, _
                                                    ByVal commandType As CommandType, _
                                                    ByVal commandText As String, _
                                                    ByVal ParamArray commandParameters() As IDataParameter) As DataSet
        If (connection Is Nothing) Then Throw New ArgumentNullException("connection")
        ' Create a command and prepare it for execution
        Dim cmd As IDbCommand = ProviderFactory.CreateCommand
        Dim ds As New DataSet
        Dim dataAdatpter As System.Data.Common.DataAdapter = Nothing
        Dim mustCloseConnection As Boolean = False

        If TypeOf (cmd) Is Oracle.DataAccess.Client.OracleCommand Then
            'CType(cmd, Oracle.DataAccess.Client.OracleCommand).AddRowid = True
            CType(cmd, Oracle.DataAccess.Client.OracleCommand).InitialLONGFetchSize = -1
            CType(cmd, Oracle.DataAccess.Client.OracleCommand).InitialLOBFetchSize = -1
        End If

        PrepareCommand(cmd, connection, CType(Nothing, IDbTransaction), commandType, commandText, commandParameters, mustCloseConnection)

        Try
            ' Create the DataAdapter & DataSet

            dataAdatpter = ProviderFactory.CreateDataAdapter(cmd)
            'dataAdatpter.MissingMappingAction = MissingMappingAction.Passthrough
            'dataAdatpter.MissingSchemaAction = MissingSchemaAction.AddWithKey
            'ds.EnforceConstraints = False
            'Utility.Debug("xxx=" & cmd.CommandText.ToString)
            ' Fill the DataSet using default values for DataTable names, etc
            dataAdatpter.Fill(ds)

            ' Detach the OleDbParameters from the command object, so they can be used again
            cmd.Parameters.Clear()
        Finally
            If (Not dataAdatpter Is Nothing) Then dataAdatpter.Dispose()
        End Try
        If (mustCloseConnection) Then connection.Close()

        ' Return the dataset
        Return ds
    End Function ' ExecuteDataset

    Public Overloads Shared Function ExecuteDataset(ByVal connectionString As String, _
                                                ByVal commandType As CommandType, _
                                                ByVal commandText As String, _
                                                ByVal ParamArray commandParameters() As IDataParameter) As DataSet

        If (connectionString Is Nothing OrElse connectionString.Length = 0) Then Throw New ArgumentNullException("connectionString")

        ' Create & open a OleDbConnection, and dispose of it after we are done
        Dim connection As IDbConnection = Nothing
        Try
            connection = ProviderFactory.CreateConnection(connectionString)
            connection.Open()

            ' Call the overload that takes a connection in place of the connection string
            Return ExecuteDataset(connection, commandType, commandText, commandParameters)
        Finally
            If Not connection Is Nothing Then connection.Close()
        End Try
    End Function ' ExecuteDataset
#End Region

#Region "ExecuteScalar"

    ' Execute a OleDbCommand (that returns a 1x1 resultset and takes no parameters) against the database specified in 
    ' the connection string. 
    ' e.g.:  
    ' Dim orderCount As Integer = CInt(ExecuteScalar(connString, CommandType.StoredProcedure, "GetOrderCount"))
    ' Parameters:
    ' -connectionString - a valid connection string for a OleDbConnection 
    ' -commandType - the CommandType (stored procedure, text, etc.) 
    ' -commandText - the stored procedure name or T-OleDb command 
    ' Returns: An object containing the value in the 1x1 resultset generated by the command
    Protected Overloads Shared Function ExecuteScalar(ByVal connectionString As String, _
                                                   ByVal commandType As CommandType, _
                                                   ByVal commandText As String) As Object
        ' Pass through the call providing null for the set of OleDbParameters
        Return ExecuteScalar(connectionString, commandType, commandText, CType(Nothing, IDataParameter()))
    End Function ' ExecuteScalar

    ' Execute a OleDbCommand (that returns a 1x1 resultset) against the database specified in the connection string 
    ' using the provided parameters.
    ' e.g.:  
    ' Dim orderCount As Integer = Cint(ExecuteScalar(connString, CommandType.StoredProcedure, "GetOrderCount", new OleDbParameter("@prodid", 24)))
    ' Parameters:
    ' -connectionString - a valid connection string for a OleDbConnection 
    ' -commandType - the CommandType (stored procedure, text, etc.) 
    ' -commandText - the stored procedure name or T-OleDb command 
    ' -commandParameters - an array of OleDbParamters used to execute the command 
    ' Returns: An object containing the value in the 1x1 resultset generated by the command 
    Public Overloads Shared Function ExecuteScalar(ByVal connectionString As String, _
                                                   ByVal commandType As CommandType, _
                                                   ByVal commandText As String, _
                                                   ByVal ParamArray commandParameters() As IDataParameter) As Object
        If (connectionString Is Nothing OrElse connectionString.Length = 0) Then Throw New ArgumentNullException("connectionString")
        ' Create & open a OleDbConnection, and dispose of it after we are done.
        Dim connection As IDbConnection = Nothing
        Try
            connection = ProviderFactory.CreateConnection(connectionString)
            connection.Open()

            ' Call the overload that takes a connection in place of the connection string
            Return ExecuteScalar(connection, commandType, commandText, commandParameters)
        Finally
            If Not connection Is Nothing Then connection.Close()
        End Try
    End Function ' ExecuteScalar

    ' Execute a OleDbCommand (that returns a 1x1 resultset and takes no parameters) against the provided OleDbConnection. 
    ' e.g.:  
    ' Dim orderCount As Integer = CInt(ExecuteScalar(conn, CommandType.StoredProcedure, "GetOrderCount"))
    ' Parameters:
    ' -connection - a valid OleDbConnection 
    ' -commandType - the CommandType (stored procedure, text, etc.) 
    ' -commandText - the stored procedure name or T-OleDb command 
    ' Returns: An object containing the value in the 1x1 resultset generated by the command 
    Public Overloads Shared Function ExecuteScalar(ByVal connection As IDbConnection, _
                                                   ByVal commandType As CommandType, _
                                                   ByVal commandText As String) As Object
        ' Pass through the call providing null for the set of OleDbParameters
        Return ExecuteScalar(connection, commandType, commandText, CType(Nothing, IDataParameter()))
    End Function ' ExecuteScalar

    ' Execute a OleDbCommand (that returns a 1x1 resultset) against the specified OleDbConnection 
    ' using the provided parameters.
    ' e.g.:  
    ' Dim orderCount As Integer = CInt(ExecuteScalar(conn, CommandType.StoredProcedure, "GetOrderCount", new OleDbParameter("@prodid", 24)))
    ' Parameters:
    ' -connection - a valid OleDbConnection 
    ' -commandType - the CommandType (stored procedure, text, etc.) 
    ' -commandText - the stored procedure name or T-OleDb command 
    ' -commandParameters - an array of OleDbParamters used to execute the command 
    ' Returns: An object containing the value in the 1x1 resultset generated by the command 
    Public Overloads Shared Function ExecuteScalar(ByVal connection As IDbConnection, _
                                                   ByVal commandType As CommandType, _
                                                   ByVal commandText As String, _
                                                   ByVal ParamArray commandParameters() As IDataParameter) As Object

        If (connection Is Nothing) Then Throw New ArgumentNullException("connection")

        ' Create a command and prepare it for execution
        Dim cmd As IDbCommand = ProviderFactory.CreateCommand
        Dim retval As Object
        Dim mustCloseConnection As Boolean = False

        If TypeOf (cmd) Is Oracle.DataAccess.Client.OracleCommand Then
            'CType(cmd, Oracle.DataAccess.Client.OracleCommand).AddRowid = True
            CType(cmd, Oracle.DataAccess.Client.OracleCommand).InitialLONGFetchSize = -1
            CType(cmd, Oracle.DataAccess.Client.OracleCommand).InitialLOBFetchSize = -1
        End If

        PrepareCommand(cmd, connection, CType(Nothing, IDbTransaction), commandType, commandText, commandParameters, mustCloseConnection)

        ' Execute the command & return the results
        retval = cmd.ExecuteScalar()

        ' Detach the OleDbParameters from the command object, so they can be used again
        cmd.Parameters.Clear()

        If (mustCloseConnection) Then connection.Close()

        Return retval

    End Function ' ExecuteScalar

    Public Overloads Shared Function ExecuteScalar(ByVal transaction As IDbTransaction, _
                                                   ByVal commandType As CommandType, _
                                                   ByVal commandText As String) As Object
        ' Pass through the call providing null for the set of OleDbParameters
        Return ExecuteScalar(transaction, commandType, commandText, CType(Nothing, IDataParameter()))
    End Function ' ExecuteScalar

    ' Execute a OleDbCommand (that returns a 1x1 resultset) against the specified OleDbTransaction
    ' using the provided parameters.
    ' e.g.:  
    ' Dim orderCount As Integer = CInt(ExecuteScalar(trans, CommandType.StoredProcedure, "GetOrderCount", new OleDbParameter("@prodid", 24)))
    ' Parameters:
    ' -transaction - a valid OleDbTransaction  
    ' -commandType - the CommandType (stored procedure, text, etc.) 
    ' -commandText - the stored procedure name or T-OleDb command 
    ' -commandParameters - an array of OleDbParamters used to execute the command 
    ' Returns: An object containing the value in the 1x1 resultset generated by the command 
    Public Overloads Shared Function ExecuteScalar(ByVal transaction As IDbTransaction, _
                                                   ByVal commandType As CommandType, _
                                                   ByVal commandText As String, _
                                                   ByVal ParamArray commandParameters() As IDataParameter) As Object
        If (transaction Is Nothing) Then Throw New ArgumentNullException("transaction")
        If Not (transaction Is Nothing) AndAlso (transaction.Connection Is Nothing) Then Throw New ArgumentException("The transaction was rollbacked or commited, please provide an open transaction.", "transaction")

        ' Create a command and prepare it for execution
        Dim cmd As IDbCommand = ProviderFactory.CreateCommand
        Dim retval As Object
        Dim mustCloseConnection As Boolean = False

        If TypeOf (cmd) Is Oracle.DataAccess.Client.OracleCommand Then
            'CType(cmd, Oracle.DataAccess.Client.OracleCommand).AddRowid = True
            CType(cmd, Oracle.DataAccess.Client.OracleCommand).InitialLONGFetchSize = -1
            CType(cmd, Oracle.DataAccess.Client.OracleCommand).InitialLOBFetchSize = -1
        End If

        PrepareCommand(cmd, transaction.Connection, transaction, commandType, commandText, commandParameters, mustCloseConnection)

        ' Execute the command & return the results
        retval = cmd.ExecuteScalar()

        ' Detach the OleDbParameters from the command object, so they can be used again
        cmd.Parameters.Clear()

        Return retval
    End Function ' ExecuteScalar




#End Region

#Region "ExecuteReader"

    Public Overloads Shared Function ExecuteReader(ByVal dbManager As DatabaseManager, _
                                                   ByVal commandType As System.Data.CommandType, _
                                                   ByVal commandText As String, _
                                                   ByVal ParamArray commandParameters() As IDataParameter) As System.Data.IDataReader
        If (dbManager.isTransaction) Then
            Return ExecuteReader(dbManager.getTransaction, commandType, commandText, commandParameters)
        Else
            Return ExecuteReader(dbManager.getConnection, commandType, commandText, commandParameters)
        End If
    End Function ' ExecuteNonQuery

    ' this enum is used to indicate whether the connection was provided by the caller, or created by SqlHelper, so that
    ' we can set the appropriate CommandBehavior when calling ExecuteReader()
    Private Enum SqlConnectionOwnership
        'Connection is owned and managed by SqlHelper
        Internal
        'Connection is owned and managed by the caller
        [External]
    End Enum 'SqlConnectionOwnership

    ' Create and prepare a OleDbCommand, and call ExecuteReader with the appropriate CommandBehavior.
    ' If we created and opened the connection, we want the connection to be closed when the DataReader is closed.
    ' If the caller provided the connection, we want to leave it to them to manage.
    ' Parameters:
    ' -connection - a valid OleDbConnection, on which to execute this command 
    ' -transaction - a valid OleDbTransaction, or 'null' 
    ' -commandType - the CommandType (stored procedure, text, etc.) 
    ' -commandText - the stored procedure name or T-SQL command 
    ' -commandParameters - an array of SqlParameters to be associated with the command or 'null' if no parameters are required 
    ' -connectionOwnership - indicates whether the connection parameter was provided by the caller, or created by SqlHelper 
    ' Returns: System.Data.Common.DbDataRecord containing the results of the command 
    Private Overloads Shared Function ExecuteReader(ByVal connection As IDbConnection, _
                                                    ByVal transaction As IDbTransaction, _
                                                    ByVal commandType As CommandType, _
                                                    ByVal commandText As String, _
                                                    ByVal commandParameters() As IDataParameter, _
                                                     ByVal connectionOwnership As SqlConnectionOwnership) As System.Data.IDataReader
        'create a command and prepare it for execution
        'Dim cmd As IDbCommand = connection.CreateCommand 'ProviderFactory.CreateCommand
        Dim cmd As IDbCommand = ProviderFactory.CreateCommand
        cmd.Connection = connection

        'create a reader
        Dim dr As System.Data.IDataReader
        Dim mustCloseConnection As Boolean = False
        PrepareCommand(cmd, connection, transaction, commandType, commandText, commandParameters, mustCloseConnection)

        ' call ExecuteReader with the appropriate CommandBehavior
        If connectionOwnership = SqlConnectionOwnership.External Then
            dr = cmd.ExecuteReader '(CommandBehavior.SequentialAccess)
        Else
            dr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
        End If

        'detach the SqlParameters from the command object, so they can be used again
        'cmd.Parameters.Clear()
        If (mustCloseConnection) Then connection.Close()
        Return dr
    End Function 'ExecuteReader

    ' Execute a OleDbCommand (that returns a resultset and takes no parameters) against the database specified in 
    ' the connection string. 
    ' e.g.:  
    ' Dim dr As System.Data.Common.DbDataRecord = ExecuteReader(connString, CommandType.StoredProcedure, "GetOrders")
    ' Parameters:
    ' -connectionString - a valid connection string for a OleDbConnection 
    ' -commandType - the CommandType (stored procedure, text, etc.) 
    ' -commandText - the stored procedure name or T-SQL command 
    ' Returns: a System.Data.Common.DbDataRecord containing the resultset generated by the command 
    Private Overloads Shared Function ExecuteReader(ByVal connectionString As String, _
                                                   ByVal commandType As CommandType, _
                                                   ByVal commandText As String) As System.Data.IDataReader
        'pass through the call providing null for the set of SqlParameters
        Return ExecuteReader(connectionString, commandType, commandText, CType(Nothing, IDataParameter()))
    End Function 'ExecuteReader

    ' Execute a OleDbCommand (that returns a resultset) against the database specified in the connection string 
    ' using the provided parameters.
    ' e.g.:  
    ' Dim dr As System.Data.Common.DbDataRecord = ExecuteReader(connString, CommandType.StoredProcedure, "GetOrders", new OleDbParameter("@prodid", 24))
    ' Parameters:
    ' -connectionString - a valid connection string for a OleDbConnection 
    ' -commandType - the CommandType (stored procedure, text, etc.) 
    ' -commandText - the stored procedure name or T-SQL command 
    ' -commandParameters - an array of SqlParamters used to execute the command 
    ' Returns: a System.Data.Common.DbDataRecord containing the resultset generated by the command 
    Private Overloads Shared Function ExecuteReader(ByVal connectionString As String, _
                                                   ByVal commandType As CommandType, _
                                                   ByVal commandText As String, _
                                                   ByVal ParamArray commandParameters() As IDataParameter) As System.Data.IDataReader
        'create & open a OleDbConnection
        Dim cn As IDbConnection = ProviderFactory.CreateConnection(connectionString)
        cn.Open()

        Try
            'call the private overload that takes an internally owned connection in place of the connection string
            Return ExecuteReader(cn, CType(Nothing, IDbTransaction), commandType, commandText, commandParameters, SqlConnectionOwnership.Internal)
        Catch
            'if we fail to return the SqlDatReader, we need to close the connection ourselves
            cn.Close()
        End Try
        Return Nothing
    End Function 'ExecuteReader

    ' Execute a stored procedure via a OleDbCommand (that returns a resultset) against the database specified in 
    ' the connection string using the provided parameter values.  This method will discover the parameters for the 
    ' stored procedure, and assign the values based on parameter order.
    ' This method provides no access to output parameters or the stored procedure's return value parameter.
    ' e.g.:  
    ' Dim dr As System.Data.Common.DbDataRecord = ExecuteReader(connString, "GetOrders", 24, 36)
    ' Parameters:
    ' -connectionString - a valid connection string for a OleDbConnection 
    ' -spName - the name of the stored procedure 
    ' -parameterValues - an array of objects to be assigned as the input values of the stored procedure 
    ' Returns: a System.Data.Common.DbDataRecord containing the resultset generated by the command 
    Private Overloads Shared Function ExecuteReader(ByVal connectionString As String, _
                                                   ByVal spName As String, _
                                                   ByVal ParamArray parameterValues() As Object) As System.Data.IDataReader
        Dim commandParameters As IDataParameter()

        'if we receive parameter values, we need to figure out where they go
        If Not (parameterValues Is Nothing) And parameterValues.Length > 0 Then
            'pull the parameters for this stored procedure from the parameter cache (or discover them & populate the cache)
            commandParameters = SqlHelperParameterCache.GetSpParameterSet(connectionString, spName)

            'assign the provided values to these parameters based on parameter order
            AssignParameterValues(commandParameters, parameterValues)

            'call the overload that takes an array of SqlParameters
            Return ExecuteReader(connectionString, CommandType.StoredProcedure, spName, commandParameters)
            'otherwise we can just call the SP without params
        Else
            Return ExecuteReader(connectionString, CommandType.StoredProcedure, spName)
        End If
    End Function 'ExecuteReader

    ' Execute a OleDbCommand (that returns a resultset and takes no parameters) against the provided OleDbConnection. 
    ' e.g.:  
    ' Dim dr As System.Data.Common.DbDataRecord = ExecuteReader(conn, CommandType.StoredProcedure, "GetOrders")
    ' Parameters:
    ' -connection - a valid OleDbConnection 
    ' -commandType - the CommandType (stored procedure, text, etc.) 
    ' -commandText - the stored procedure name or T-SQL command 
    ' Returns: a System.Data.Common.DbDataRecord containing the resultset generated by the command 
    Private Overloads Shared Function ExecuteReader(ByVal connection As IDbConnection, _
                                                   ByVal commandType As CommandType, _
                                                   ByVal commandText As String) As System.Data.IDataReader

        Return ExecuteReader(connection, commandType, commandText, CType(Nothing, IDataParameter()))

    End Function 'ExecuteReader

    ' Execute a OleDbCommand (that returns a resultset) against the specified OleDbConnection 
    ' using the provided parameters.
    ' e.g.:  
    ' Dim dr As System.Data.Common.DbDataRecord = ExecuteReader(conn, CommandType.StoredProcedure, "GetOrders", new OleDbParameter("@prodid", 24))
    ' Parameters:
    ' -connection - a valid OleDbConnection 
    ' -commandType - the CommandType (stored procedure, text, etc.) 
    ' -commandText - the stored procedure name or T-SQL command 
    ' -commandParameters - an array of SqlParamters used to execute the command 
    ' Returns: a System.Data.Common.DbDataRecord containing the resultset generated by the command 
    Public Overloads Shared Function ExecuteReader(ByVal connection As IDbConnection, _
                                                   ByVal commandType As CommandType, _
                                                   ByVal commandText As String, _
                                                   ByVal ParamArray commandParameters() As IDataParameter) As System.Data.IDataReader
        'pass through the call to private overload using a null transaction value
        Return ExecuteReader(connection, CType(Nothing, IDbTransaction), commandType, commandText, commandParameters, SqlConnectionOwnership.External)

    End Function 'ExecuteReader

    '' Execute a stored procedure via a OleDbCommand (that returns a resultset) against the specified OleDbConnection 
    '' using the provided parameter values.  This method will discover the parameters for the 
    '' stored procedure, and assign the values based on parameter order.
    '' This method provides no access to output parameters or the stored procedure's return value parameter.
    '' e.g.:  
    '' Dim dr As System.Data.Common.DbDataRecord = ExecuteReader(conn, "GetOrders", 24, 36)
    '' Parameters:
    '' -connection - a valid OleDbConnection 
    '' -spName - the name of the stored procedure 
    '' -parameterValues - an array of objects to be assigned as the input values of the stored procedure 
    '' Returns: a System.Data.Common.DbDataRecord containing the resultset generated by the command 
    'Public Overloads Shared Function ExecuteReader(ByVal connection As IDbConnection, _
    '                                               ByVal spName As String, _
    '                                               ByVal ParamArray parameterValues() As Object) As System.Data.Common.DbDataRecord
    '    'pass through the call using a null transaction value
    '    'Return ExecuteReader(connection, CType(Nothing, IDbTransaction), spName, parameterValues)

    '    Dim commandParameters AS IDataParameter()

    '    'if we receive parameter values, we need to figure out where they go
    '    If Not (parameterValues Is Nothing) And parameterValues.Length > 0 Then
    '        commandParameters = SqlHelperParameterCache.GetSpParameterSet(connection.ConnectionString, spName)

    '        AssignParameterValues(commandParameters, parameterValues)

    '        Return ExecuteReader(CommandType.StoredProcedure, spName, commandParameters)
    '        'otherwise we can just call the SP without params
    '    Else
    '        Return ExecuteReader(connection, CommandType.StoredProcedure, spName)
    '    End If

    'End Function 'ExecuteReader

    ' Execute a OleDbCommand (that returns a resultset and takes no parameters) against the provided OleDbTransaction.
    ' e.g.:  
    ' Dim dr As System.Data.Common.DbDataRecord = ExecuteReader(trans, CommandType.StoredProcedure, "GetOrders")
    ' Parameters:
    ' -transaction - a valid OleDbTransaction  
    ' -commandType - the CommandType (stored procedure, text, etc.) 
    ' -commandText - the stored procedure name or T-SQL command 
    ' Returns: a System.Data.Common.DbDataRecord containing the resultset generated by the command 
    Private Overloads Shared Function ExecuteReader(ByVal transaction As IDbTransaction, _
                                                   ByVal commandType As CommandType, _
                                                   ByVal commandText As String) As System.Data.IDataReader
        'pass through the call providing null for the set of SqlParameters
        Return ExecuteReader(transaction, commandType, commandText, CType(Nothing, IDataParameter()))
    End Function 'ExecuteReader

    ' Execute a OleDbCommand (that returns a resultset) against the specified OleDbTransaction
    ' using the provided parameters.
    ' e.g.:  
    ' Dim dr As System.Data.Common.DbDataRecord = ExecuteReader(trans, CommandType.StoredProcedure, "GetOrders", new OleDbParameter("@prodid", 24))
    ' Parameters:
    ' -transaction - a valid OleDbTransaction 
    ' -commandType - the CommandType (stored procedure, text, etc.)
    ' -commandText - the stored procedure name or T-SQL command 
    ' -commandParameters - an array of SqlParamters used to execute the command 
    ' Returns: a System.Data.Common.DbDataRecord containing the resultset generated by the command 
    Public Overloads Shared Function ExecuteReader(ByVal transaction As IDbTransaction, _
                                                   ByVal commandType As CommandType, _
                                                   ByVal commandText As String, _
                                                   ByVal ParamArray commandParameters() As IDataParameter) As System.Data.IDataReader
        'pass through to private overload, indicating that the connection is owned by the caller
        Return ExecuteReader(transaction.Connection, transaction, commandType, commandText, commandParameters, SqlConnectionOwnership.External)
    End Function 'ExecuteReader

    '' Execute a stored procedure via a OleDbCommand (that returns a resultset) against the specified OleDbTransaction 
    '' using the provided parameter values.  This method will discover the parameters for the 
    '' stored procedure, and assign the values based on parameter order.
    '' This method provides no access to output parameters or the stored procedure's return value parameter.
    '' e.g.:  
    '' Dim dr As System.Data.Common.DbDataRecord = ExecuteReader(trans, "GetOrders", 24, 36)
    '' Parameters:
    '' -transaction - a valid OleDbTransaction 
    '' -spName - the name of the stored procedure 
    '' -parameterValues - an array of objects to be assigned as the input values of the stored procedure
    '' Returns: a System.Data.Common.DbDataRecord containing the resultset generated by the command
    'Public Overloads Shared Function ExecuteReader(ByVal transaction As IDbTransaction, _
    '                                               ByVal spName As String, _
    '                                               ByVal ParamArray parameterValues() As Object) As System.Data.Common.DbDataRecord
    '    Dim commandParameters AS IDataParameter()

    '    'if we receive parameter values, we need to figure out where they go
    '    If Not (parameterValues Is Nothing) And parameterValues.Length > 0 Then
    '        commandParameters = SqlHelperParameterCache.GetSpParameterSet(transaction.Connection.ConnectionString, spName)

    '        AssignParameterValues(commandParameters, parameterValues)

    '        Return ExecuteReader(transaction, CommandType.StoredProcedure, spName, commandParameters)
    '        'otherwise we can just call the SP without params
    '    Else
    '        Return ExecuteReader(transaction, CommandType.StoredProcedure, spName)
    '    End If
    'End Function 'ExecuteReader

#End Region

#Region "private utility methods"

    ' This method opens (if necessary) and assigns a connection, transaction, command type and parameters 
    ' to the provided command.
    ' Parameters:
    ' -command - the OleDbCommand to be prepared
    ' -connection - a valid OleDbConnection, on which to execute this command
    ' -transaction - a valid OleDbTransaction, or ' null' 
    ' -commandType - the CommandType (stored procedure, text, etc.)
    ' -commandText - the stored procedure name or T-OleDb command
    ' -commandParameters - an array of OleDbParameters to be associated with the command or ' null' if no parameters are required
    Private Shared Sub PrepareCommand(ByVal command As IDbCommand, _
                                      ByVal connection As IDbConnection, _
                                      ByVal transaction As IDbTransaction, _
                                      ByVal commandType As CommandType, _
                                      ByVal commandText As String, _
                                      ByVal commandParameters() As IDataParameter, ByRef mustCloseConnection As Boolean)

        If (command Is Nothing) Then Throw New ArgumentNullException("command")
        If (commandText Is Nothing OrElse commandText.Length = 0) Then Throw New ArgumentNullException("commandText")

        ' If the provided connection is not open, we will open it
        If connection.State <> ConnectionState.Open Then
            connection.Open()
            mustCloseConnection = True
        Else
            mustCloseConnection = False
        End If

        ' Associate the connection with the command
        command.Connection = connection

        ' Set the command text (stored procedure name or OleDb statement)
        command.CommandText = commandText

        ' If we were provided a transaction, assign it.
        If Not (transaction Is Nothing) Then
            If transaction.Connection Is Nothing Then Throw New ArgumentException("The transaction was rollbacked or commited, please provide an open transaction.", "transaction")
            command.Transaction = transaction
        End If

        ' Set the command type
        command.CommandType = commandType

        ' Attach the command parameters if they are provided
        If Not (commandParameters Is Nothing) Then
            AttachParameters(command, commandParameters)
        End If
        Return
    End Sub ' PrepareCommand

    ' This method is used to attach array of OleDbParameters to a OleDbCommand.
    ' This method will assign a value of DbNull to any parameter with a direction of
    ' InputOutput and a value of null.  
    ' This behavior will prevent default values from being used, but
    ' this will be the less common case than an intended pure output parameter (derived as InputOutput)
    ' where the user provided no input value.
    ' Parameters:
    ' -command - The command to which the parameters will be added
    ' -commandParameters - an array of OleDbParameters to be added to command
    Private Shared Sub AttachParameters(ByVal command As IDbCommand, ByVal commandParameters() As IDataParameter)
        If (command Is Nothing) Then Throw New ArgumentNullException("command")
        If (Not commandParameters Is Nothing) Then
            Dim p As IDataParameter
            For Each p In commandParameters

                If (Not p Is Nothing) Then
                    ' Check for derived output value with no value assigned
                    If (p.Direction = ParameterDirection.InputOutput OrElse p.Direction = ParameterDirection.Input) AndAlso p.Value Is Nothing Then
                        p.Value = DBNull.Value
                    End If
                    command.Parameters.Add(p)
                End If

            Next p
        End If
    End Sub ' AttachParameters

    ' This method assigns an array of values to an array of SqlParameters.
    ' Parameters:
    ' -commandParameters - array of SqlParameters to be assigned values
    ' -array of objects holding the values to be assigned
    Private Shared Sub AssignParameterValues(ByVal commandParameters() As IDataParameter, ByVal parameterValues() As Object)

        Dim i As Short
        Dim j As Short

        If (commandParameters Is Nothing) And (parameterValues Is Nothing) Then
            'do nothing if we get no data
            Return
        End If

        ' we must have the same number of values as we pave parameters to put them in
        If commandParameters.Length <> parameterValues.Length Then
            Throw New ArgumentException("Parameter count does not match Parameter Value count.")
        End If

        'value array
        j = CShort(commandParameters.Length - 1)
        For i = 0 To j
            commandParameters(i).Value = parameterValues(i)
        Next

    End Sub 'AssignParameterValues

    Private Shared Sub createParameters(ByVal command As IDbCommand, ByVal commandParameters() As IDataParameter, ByVal values As ArrayList, ByVal sizes As ArrayList)

        'Dim iParaCount As Int16 = commandParameters.Length

        'If (values.Count + sizes.Count <> iParaCount) Then
        '    Throw New Exception("Parameters is not correct" & values.Count + sizes.Count & "<>" & iParaCount)
        'End If

        Dim p As IDataParameter
        Dim index As Int16 = 0
        Dim inCount As Int16 = 0
        Dim outCount As Int16 = 0

        For Each p In command.Parameters
            'Console.WriteLine("ParameterName=" & p.ParameterName)

            If p.Direction = ParameterDirection.Input Then
                p.Value = values.Item(inCount)
                inCount = CShort(inCount + 1)
            ElseIf p.Direction = ParameterDirection.Output Or p.Direction = ParameterDirection.InputOutput Then
                'If p.OleDbType = OleDbType.VarChar Then
                If p.DbType = System.Data.DbType.String OrElse p.DbType = DbType.AnsiString Then
                    CType(p, IDbDataParameter).Size = CInt(sizes.Item(outCount))
                    outCount = CShort(outCount + 1)
                End If
            End If

            commandParameters(index) = p
            index = CShort(index + 1)
        Next p

        command.Parameters.Clear()
    End Sub
#End Region

    ' SqlHelperParameterCache provides functions to leverage a static cache of procedure parameters, and the
    ' ability to discover parameters for stored procedures at run-time.
    Public NotInheritable Class SqlHelperParameterCache

#Region "private methods, variables, and constructors"


        'Since this class provides only static methods, make the default constructor private to prevent 
        'instances from being created with "new SqlHelperParameterCache()".
        Private Sub New()
        End Sub 'New 

        Private Shared paramCache As Hashtable = Hashtable.Synchronized(New Hashtable)

        ' resolve at run time the appropriate set of SqlParameters for a stored procedure
        ' Parameters:
        ' - connectionString - a valid connection string for a OleDbConnection
        ' - spName - the name of the stored procedure
        ' - includeReturnValueParameter - whether or not to include their return value parameter>
        ' Returns: OleDbParameter()
        Private Shared Function DiscoverSpParameterSet(ByVal connectionString As String, _
                                                       ByVal spName As String, _
                                                       ByVal includeReturnValueParameter As Boolean, _
                                                       ByVal ParamArray parameterValues() As Object) As IDataParameter()

            Dim cn As IDbConnection = ProviderFactory.CreateConnection(connectionString)
            Dim cmd As IDbCommand = ProviderFactory.CreateCommand(spName, cn)
            cmd.CommandTimeout = cn.ConnectionTimeout

            Dim discoveredParameters() As IDataParameter

            Try
                cn.Open()
                cmd.CommandType = CommandType.StoredProcedure
                ProviderFactory.CommandBuilderDeriveParameters(cmd)
                If Not includeReturnValueParameter Then
                    cmd.Parameters.RemoveAt(0)
                End If

                discoveredParameters = New IDataParameter(cmd.Parameters.Count - 1) {}
                cmd.Parameters.CopyTo(discoveredParameters, 0)
            Finally
                cmd.Dispose()
                cn.Close()

            End Try

            Return discoveredParameters

        End Function 'DiscoverSpParameterSet

        'deep copy of cached OleDbParameter array
        Private Shared Function CloneParameters(ByVal originalParameters() As IDataParameter) As IDataParameter()

            Dim i As Short
            Dim j As Short = CShort(originalParameters.Length - 1)
            Dim clonedParameters(j) As IDataParameter

            For i = 0 To j
                clonedParameters(i) = CType(CType(originalParameters(i), ICloneable).Clone, IDataParameter)
            Next

            Return clonedParameters
        End Function 'CloneParameters



#End Region

#Region "caching functions"

        ' add parameter array to the cache
        ' Parameters
        ' -connectionString - a valid connection string for a OleDbConnection 
        ' -commandText - the stored procedure name or T-SQL command 
        ' -commandParameters - an array of SqlParamters to be cached 
        Private Shared Sub CacheParameterSet(ByVal connectionString As String, _
                                            ByVal commandText As String, _
                                            ByVal ParamArray commandParameters() As IDataParameter)
            Dim hashKey As String = connectionString + ":" + commandText

            paramCache(hashKey) = commandParameters
        End Sub 'CacheParameterSet

        ' retrieve a parameter array from the cache
        ' Parameters:
        ' -connectionString - a valid connection string for a OleDbConnection 
        ' -commandText - the stored procedure name or T-SQL command 
        ' Returns: an array of SqlParamters 
        Private Shared Function GetCachedParameterSet(ByVal connectionString As String, ByVal commandText As String) As IDataParameter()
            Dim hashKey As String = connectionString + ":" + commandText
            Dim cachedParameters As IDataParameter() = CType(paramCache(hashKey), IDataParameter())

            If cachedParameters Is Nothing Then
                Return Nothing
            Else
                Return CloneParameters(cachedParameters)
            End If
        End Function 'GetCachedParameterSet

#End Region

#Region "Parameter Discovery Functions"
        ' Retrieves the set of SqlParameters appropriate for the stored procedure
        ' 
        ' This method will query the database for this information, and then store it in a cache for future requests.
        ' Parameters:
        ' -connectionString - a valid connection string for a OleDbConnection 
        ' -spName - the name of the stored procedure 
        ' Returns: an array of SqlParameters
        Protected Friend Overloads Shared Function GetSpParameterSet(ByVal connectionString As String, ByVal spName As String) As IDataParameter()
            Return GetSpParameterSet(connectionString, spName, False)
        End Function 'GetSpParameterSet 

        ' Retrieves the set of SqlParameters appropriate for the stored procedure
        ' 
        ' This method will query the database for this information, and then store it in a cache for future requests.
        ' Parameters:
        ' -connectionString - a valid connection string for a OleDbConnection
        ' -spName - the name of the stored procedure 
        ' -includeReturnValueParameter - a bool value indicating whether the return value parameter should be included in the results 
        ' Returns: an array of SqlParameters 
        Private Overloads Shared Function GetSpParameterSet(ByVal connectionString As String, _
                                                           ByVal spName As String, _
                                                           ByVal includeReturnValueParameter As Boolean) As IDataParameter()

            Dim cachedParameters() As IDataParameter
            Dim hashKey As String

            hashKey = String.Format("{0}:{1}{2}", connectionString, spName, IIf(includeReturnValueParameter = True, ":include ReturnValue Parameter", ""))

            cachedParameters = CType(paramCache(hashKey), IDataParameter())

            If (cachedParameters Is Nothing) Then
                paramCache(hashKey) = DiscoverSpParameterSet(connectionString, spName, includeReturnValueParameter)
                cachedParameters = CType(paramCache(hashKey), IDataParameter())

            End If

            Return CloneParameters(cachedParameters)

        End Function 'GetSpParameterSet
#End Region

    End Class 'SqlHelperParameterCache 

#Region "Protected Function"
    'add by ted 20070116
    Public Shared Function getProcedureData(ByVal dbManager As DatabaseManager, _
                                               ByVal SPName As String, ByVal values As ArrayList, ByVal sizes As ArrayList) As Hashtable
        Try
            If Not IsNothing(dbManager) Then
                If IsNothing(dbManager.getTransaction) Then
                    Return getProcedureData(dbManager.getConnection, SPName, values, sizes)
                Else
                    Return getProcedureData(dbManager.getTransaction, SPName, values, sizes)
                End If
            End If
            Return Nothing
        Catch ex As Exception
            Throw New Exception(ex.Message & ex.StackTrace)
        End Try
    End Function

    Protected Shared Function getProcedureData(ByVal transaction As IDbTransaction, _
                                               ByVal SPName As String, ByVal values As ArrayList, ByVal sizes As ArrayList) As Hashtable
        Dim cmd As IDbCommand = ProviderFactory.CreateCommand
        Try
            cmd.Connection = transaction.Connection
            cmd.CommandType = CommandType.StoredProcedure
            'modify by ted 20070117
            cmd.CommandText = SPName 'combineSName(SPName) 'schemaName.SPName
            '耝Stored Procedure把计戈癟
            ProviderFactory.CommandBuilderDeriveParameters(cmd)

            Dim iParaCount As Integer = cmd.Parameters.Count

            Dim paras(iParaCount - 1) As IDataParameter
            cmd.Parameters.CopyTo(paras, 0)
            createParameters(cmd, paras, values, sizes)

            ExecuteDataset(transaction, cmd.CommandType, cmd.CommandText, paras)

            Dim htbOutValue As New Hashtable
            For i As Integer = values.Count To iParaCount - 1
                htbOutValue.Add(paras(i).ParameterName.ToUpper, IIf(IsDBNull(paras(i).Value), Nothing, paras(i).Value))
            Next

            Return htbOutValue
        Catch ex As Exception
            Throw New Exception(ex.Message & ex.StackTrace)
        End Try
    End Function

    Protected Shared Function getProcedureData(ByVal connection As IDbConnection, _
                                               ByVal SPName As String, ByVal values As ArrayList, ByVal sizes As ArrayList) As Hashtable
        Dim cmd As IDbCommand = ProviderFactory.CreateCommand
        Try
            cmd.Connection = connection
            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandText = SPName
            'cmd.CommandText = combineSName(SPName)
            '耝Stored Procedure把计戈癟
            ProviderFactory.CommandBuilderDeriveParameters(cmd)

            Dim iParaCount As Integer = cmd.Parameters.Count

            Dim paras(iParaCount - 1) As IDataParameter
            cmd.Parameters.CopyTo(paras, 0)
            createParameters(cmd, paras, values, sizes)

            ExecuteDataset(connection, cmd.CommandType, cmd.CommandText, paras)

            Dim htbOutValue As New Hashtable
            For i As Integer = values.Count To iParaCount - 1
                htbOutValue.Add(paras(i).ParameterName.ToUpper, IIf(IsDBNull(paras(i).Value), Nothing, paras(i).Value))
            Next

            Return htbOutValue
        Catch ex As Exception
            Throw New Exception(ex.Message & ex.StackTrace)
        End Try
    End Function

#End Region

    'Public Shared Function combineSName(ByVal sObjName As String) As String
    '    Return (" " & SCHEMANAME & "." & sObjName & " ").ToUpper
    'End Function

#Region "UpdateDataset"

    'Public Overloads Shared Function UpdateDataset(ByVal ds As DataSet, _
    '                                                ByVal updateCmd As OleDbCommand, _
    '                                                ByVal transaction As IDbTransaction, _
    '                                                ByVal commandType As CommandType, _
    '                                                ByVal commandText As String, _
    '                                                ByVal ParamArray commandParameters() AS IDataParameter) As DataSet
    '    If (transaction Is Nothing) Then Throw New ArgumentNullException("transaction")
    '    If Not (transaction Is Nothing) AndAlso (transaction.Connection Is Nothing) Then Throw New ArgumentException("The transaction was rollbacked or commited, please provide an open transaction.", "transaction")

    '    ' Create a command and prepare it for execution
    '    Dim cmd As IDbCommand = ProviderFactory.CreateCommand

    '    Dim dataAdatpter As System.Data.Common.DataAdapter
    '    Dim mustCloseConnection As Boolean = False

    '    'Provider=MSDAORA.1;max pool size=350;Password=pruser;User ID=pruser;Data Source=test;
    '    PrepareCommand(cmd, transaction.Connection, transaction, commandType, commandText, commandParameters, mustCloseConnection)

    '    Try
    '        dataAdatpter = ProviderFactory.CreateDataAdapter(cmd)

    '        Dim oleCmdBuilder As OleDbCommandBuilder = New OleDbCommandBuilder(dataAdatpter)

    '        'Utility.Debug("  oleCmdBuilder.SelectCommand=" & oleCmdBuilder.DataAdapter.SelectCommand.CommandText.ToString())

    '        dataAdatpter.UpdateCommand = updateCmd
    '        'Utility.Debug("  oleCmdBuilder.UpdateCommand=" & oleCmdBuilder.DataAdapter.UpdateCommand.CommandText.ToString())

    '        'Utility.Debug("  oleCmdBuilder.InsertCommand=" & oleCmdBuilder.DataAdapter.InsertCommand.CommandText.ToString())

    '        'Utility.Debug("========" & dataAdatpter.Update(ds) & "========")

    '        ' Detach the OleDbParameters from the command object, so they can be used again
    '        cmd.Parameters.Clear()
    '    Finally
    '        If (Not dataAdatpter Is Nothing) Then dataAdatpter.Dispose()
    '    End Try

    '    ' Return the dataset
    '    Return ds

    'End Function ' UpdateDataset


    'Public Overloads Shared Function UpdateDataset(ByVal updateDs As DataSet, _
    '                                                  ByVal updateCmd As OleDbCommand, _
    '                                                  ByVal connection As IDbConnection, _
    '                                                  ByVal commandType As CommandType, _
    '                                                  ByVal commandText As String, _
    '                                                  ByVal ParamArray commandParameters() AS IDataParameter) As DataSet
    '    If (connection Is Nothing) Then Throw New ArgumentNullException("connection")
    '    ' Create a command and prepare it for execution
    '    Dim cmd As IDbCommand = ProviderFactory.CreateCommand

    '    Dim dataAdatpter As System.Data.Common.DataAdapter
    '    Dim mustCloseConnection As Boolean = False

    '    PrepareCommand(cmd, connection, CType(Nothing, IDbTransaction), commandType, commandText, commandParameters, mustCloseConnection)

    '    Try
    '        ' Create the DataAdapter & DataSet
    '        dataAdatpter = New OleDbDataAdapter

    '        Dim oleCmdBuilder As OleDbCommandBuilder = New OleDbCommandBuilder(dataAdatpter)

    '        ' Utility.Debug("  oleCmdBuilder.SelectCommand=" & oleCmdBuilder.DataAdapter.SelectCommand.CommandText.ToString())

    '        dataAdatpter.UpdateCommand = updateCmd
    '        'Utility.Debug("  oleCmdBuilder.UpdateCommand=" & oleCmdBuilder.DataAdapter.UpdateCommand.CommandText.ToString())

    '        'For Each para AS IDataParameter In updateCmd.Parameters
    '        '    Utility.Debug(para.ParameterName & "=" & para.Value)
    '        'Next




    '        'Utility.Debug("========" & dataAdatpter.Update(updateDs) & "========")

    '        ' Detach the OleDbParameters from the command object, so they can be used again
    '        cmd.Parameters.Clear()
    '    Finally
    '        If (Not dataAdatpter Is Nothing) Then dataAdatpter.Dispose()
    '    End Try
    '    If (mustCloseConnection) Then connection.Close()

    '    ' Return the dataset
    '    Return updateDs

    'End Function ' UpdateDataset
#End Region

    Private Shared Function Logger() As Object
        Throw New NotImplementedException
    End Function

End Class