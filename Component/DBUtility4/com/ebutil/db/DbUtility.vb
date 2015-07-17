Option Explicit On
Option Strict On

Imports System.Collections
Imports System.Data.OleDb
Imports System.Data

Public Class DbUtility

#Region "getField"

    Shared Function getField(ByVal iValue As Long, ByVal defaultValue As Long) As Long

        If (Nothing = iValue) Then
            Return defaultValue
        Else
            Return iValue
        End If
    End Function

    Shared Function getField(ByVal iValue As Long) As Long
        Return getField(iValue, 0)
    End Function

    Shared Function getField(ByVal iValue As String, ByVal defaultValue As String) As String

        If (Nothing = iValue) Then
            Return defaultValue
        Else
            Return iValue.Trim()
        End If
    End Function

    Shared Function getField(ByVal iValue As String) As String
        Return getField(iValue, "")
    End Function

    Shared Function getField(ByVal iValue As Integer, ByVal defaultValue As Integer) As Integer

        If (Nothing = iValue) Then
            Return defaultValue
        Else
            Return iValue
        End If
    End Function

    Shared Function getField(ByVal iValue As Integer) As Integer
        Return getField(iValue, 0)
    End Function
    '
    Shared Function getField(ByVal iValue As Decimal, ByVal defaultValue As Decimal) As Decimal

        If (Nothing = iValue) Then
            Return defaultValue
        Else
            Return iValue
        End If
    End Function

    Shared Function getField(ByVal iValue As Decimal) As Decimal
        Return getField(iValue, 0)
    End Function

    Shared Function getField(ByVal dValue As DateTime, ByVal defaultValue As DateTime) As DateTime

        If (Nothing = dValue) Then
            Return defaultValue
        Else
            Return dValue
        End If
    End Function

    Shared Function getField(ByVal iValue As DateTime) As DateTime
        Return getField(iValue, Nothing)
    End Function

    Shared Function getField(ByVal iValue As Int16, ByVal defaultValue As Int16) As Int16

        If (Nothing = iValue) Then
            Return defaultValue
        Else
            Return iValue
        End If
    End Function

    Shared Function getField(ByVal iValue As Int16) As Int16
        Return getField(iValue, CShort(0))
    End Function

    '將DBNULL設為Nothing
    Public Shared Function checkDBNull(ByVal ds As DataSet, ByVal columnName As String, ByVal index As Int16) As Object
        Dim obj As Object = ds.Tables(0).DefaultView.Item(index).Item(columnName)
        Return IIf(IsDBNull(obj), Nothing, obj)
    End Function
#End Region

#Region "loadSchemaInfo"

    Public Shared Function loadSchemaInfo(ByVal databaseManager As DatabaseManager, ByVal sTableName As String) As DataTable
        Dim schemaTable As DataTable = Nothing
        Dim ds As New DataSet(sTableName)
        Dim sSQL As String = "SELECT * FROM " & sTableName '& " where rownum=1"
        Dim cmd As IDbCommand = ProviderFactory.CreateCommand
        Dim dataAdatpter As System.Data.Common.DataAdapter = ProviderFactory.CreateDataAdapter(cmd)
        Dim transaction As IDbTransaction = databaseManager.getTransaction
        Dim connection As IDbConnection = databaseManager.getConnection

        Try
            cmd.Connection = connection
            If Not (transaction Is Nothing) Then
                If transaction.Connection Is Nothing Then Throw New ArgumentException("The transaction was rollbacked or commited, please provide an open transaction.", "transaction")
                cmd.Transaction = transaction
            End If
            cmd.CommandText = sSQL
            'dataAdatpter.MissingMappingAction = MissingMappingAction.Passthrough
            'dataAdatpter.MissingSchemaAction = MissingSchemaAction.AddWithKey
            'ds.EnforceConstraints = False
            schemaTable = dataAdatpter.FillSchema(ds, SchemaType.Mapped)(0)
            'Dim s As String = ds.GetXmlSchema
        Catch ex As Exception
            Throw
        Finally
            If (Not dataAdatpter Is Nothing) Then dataAdatpter.Dispose()
            If (Not ds Is Nothing) Then ds.Dispose()
        End Try

        If Properties.getProvider() = ProviderType.OleDb Then
            cmd.Parameters.Clear()
            dataAdatpter = ProviderFactory.CreateDataAdapter(cmd)
            Dim dtPK As DataTable
            Try
                ds = New DataSet(sTableName)
                sSQL = "select distinct  t.column_name " & _
                        "  from (select *" & _
                        "          from user_constraints" & _
                        "         where owner = '" & Properties.getSchemaName().ToUpper & "'" & _
                        "           and constraint_type = 'P') m," & _
                        "       (select *" & _
                        "          from user_cons_columns" & _
                        "         where owner = '" & Properties.getSchemaName().ToUpper & "'" & _
                        "         order by table_name, column_name, position) t" & _
                        " where m.table_name = t.table_name" & _
                        "   and t.constraint_name = m.constraint_name" & _
                        "   and t.table_name =  " & ProviderFactory.PositionPara & "TABLENAME"

                cmd.CommandText = sSQL
                cmd.Parameters.Add(com.Azion.NET.VB.ProviderFactory.CreateDataParameter("TABLENAME", sTableName))

                'dataAdatpter.MissingMappingAction = MissingMappingAction.Passthrough
                'dataAdatpter.MissingSchemaAction = MissingSchemaAction.AddWithKey
                'ds.EnforceConstraints = False
                dataAdatpter.Fill(ds)
                dtPK = ds.Tables(0)

                Dim i As Integer = 0
                Dim dataColumns(dtPK.Rows.Count - 1) As DataColumn

                For Each row As DataRow In dtPK.Rows
                    For Each dataColumn As DataColumn In schemaTable.Columns
                        If dataColumn.ColumnName.ToString = CStr(row(0)) Then
                            dataColumns(i) = dataColumn
                            i += 1
                            Exit For
                        End If
                    Next
                Next

                schemaTable.PrimaryKey = dataColumns

            Catch ex As Exception
                Throw
            Finally
                If (Not cmd Is Nothing) Then cmd.Dispose()
                If (Not ds Is Nothing) Then ds.Dispose()
                If (Not dataAdatpter Is Nothing) Then dataAdatpter.Dispose()
            End Try
        End If
        Return schemaTable
    End Function
#End Region

#Region "getPrimaryKey"

#Region "DatabaseManager"
    Public Shared Function getPrimaryKey(ByVal sTableName As String, ByVal databaseManager As DatabaseManager) As ArrayList

        If Properties.getProvider() = ProviderType.OleDb Then
            Return getOLEPrimaryKey(sTableName, CType(databaseManager.getConnection, OleDb.OleDbConnection))
        ElseIf Properties.getProvider() = ProviderType.SqlClient Then
            Return getSQLPrimaryKey(sTableName, databaseManager)
        ElseIf Properties.getProvider() = ProviderType.ODP Then
            Return getODPPrimaryKey(sTableName, databaseManager)
        End If

        Return getPrimaryKey(sTableName, databaseManager.getConnection)

    End Function
#End Region

#Region "IDbConnection"
    Public Shared Function getPrimaryKey(ByVal sTableName As String, ByVal connection As System.Data.IDbConnection) As ArrayList

        Dim aryColumnName As New ArrayList
        Dim schemaTable As DataTable = Nothing
        Dim dataReader As System.Data.IDataReader = Nothing
        Dim command As System.Data.IDbCommand = Nothing
        'Dim oracleParameter As Oracle.DataAccess.Client.OracleParameter
        Dim sSQL As String
        Try

            sSQL = "SELECT * FROM " & sTableName
            '取得欄位的資訊
            command = connection.CreateCommand

            command.CommandType = CommandType.Text
            command.CommandText = sSQL

            dataReader = command.ExecuteReader(CommandBehavior.KeyInfo) 'Or CommandBehavior.CloseConnection

            'get the schema table
            schemaTable = dataReader.GetSchemaTable()

            For i As Integer = 0 To schemaTable.Rows.Count - 1
                If CBool(schemaTable.Rows(i).Item("IsKey")) Then
                    'Console.WriteLine("PrimaryKeys=" & (schemaTable.Rows(i)!ColumnName.ToString)) 'COLUMN_NAME
                    aryColumnName.Add((schemaTable.Rows(i)!ColumnName.ToString))
                End If
            Next i


            Return aryColumnName
        Catch ex As Exception

            Throw
        Finally
            If Not IsNothing(dataReader) Then dataReader.Close()
            If Not IsNothing(command) Then command.Dispose()
        End Try
    End Function
#End Region

#Region "ODP"
    Public Shared Function getODPPrimaryKey(ByVal sTableName As String, ByVal databaseManager As DatabaseManager) As ArrayList

        Dim aryColumnName As New ArrayList
        Dim schemaTable As DataTable

        Dim sSQL As String
        Try

            sSQL = "SELECT * FROM " & Properties.getSchemaName().ToUpper & "." & sTableName
            '取得欄位的資訊
            Using command As System.Data.IDbCommand = CType(databaseManager.getConnection, Oracle.DataAccess.Client.OracleConnection).CreateCommand

                command.Transaction = CType(databaseManager.getTransaction, Oracle.DataAccess.Client.OracleTransaction)
                command.CommandType = CommandType.Text
                command.CommandText = sSQL

                Dim para As Oracle.DataAccess.Client.OracleParameter = CType(command.CreateParameter(), Oracle.DataAccess.Client.OracleParameter)
                para.ParameterName = "TABLE_NAME"
                para.Value = sTableName

                command.Parameters.Add(para)

                Using dataReader As System.Data.IDataReader = command.ExecuteReader(CommandBehavior.KeyInfo)
                    schemaTable = dataReader.GetSchemaTable()
                    For i As Integer = 0 To schemaTable.Rows.Count - 1
                        If CBool(schemaTable.Rows(i).Item("IsKey")) Then
                            'Titan.Utility.Debug("PrimaryKeys=" & (schemaTable.Rows(i)!ColumnName.ToString)) 'COLUMN_NAME
                            aryColumnName.Add((schemaTable.Rows(i)!ColumnName.ToString))
                        End If
                    Next i
                End Using
            End Using

            Return aryColumnName
        Catch ex As Exception
            Throw
        End Try
    End Function
#End Region

#Region "SQLClient"
    Public Shared Function getSQLPrimaryKey(ByVal sTableName As String, ByVal databaseManager As DatabaseManager) As ArrayList
        Dim aryColumnName As New ArrayList
        Dim sSQL As String
        Try

            sSQL = "  SELECT i1.TABLE_NAME, i2.COLUMN_NAME" & _
                    "          FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS i1" & _
                    "         INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE i2" & _
                    "            ON i1.CONSTRAINT_NAME = i2.CONSTRAINT_NAME" & _
                    "         WHERE i1.CONSTRAINT_TYPE = 'PRIMARY KEY'" & _
                    "           and i1.TABLE_NAME =" & ProviderFactory.PositionPara & "TABLE_NAME"

            Using command As System.Data.IDbCommand = CType(databaseManager.getConnection, SqlClient.SqlConnection).CreateCommand

                command.Transaction = CType(databaseManager.getTransaction, SqlClient.SqlTransaction)
                command.CommandType = CommandType.Text
                command.CommandText = sSQL

                Dim para As SqlClient.SqlParameter = CType(command.CreateParameter(), SqlClient.SqlParameter)
                para.ParameterName = "TABLE_NAME"
                para.Value = sTableName

                command.Parameters.Add(para)

                Using dataReader As System.Data.IDataReader = command.ExecuteReader()
                    While dataReader.Read
                        aryColumnName.Add(dataReader("COLUMN_NAME"))
                    End While
                End Using
            End Using

            Return aryColumnName
        Catch ex As Exception

            Throw
        End Try
    End Function
#End Region

#Region "OleDbConnection"

    Public Shared Function getOLEPrimaryKey(ByVal sTableName As String, ByVal connection As System.Data.OleDb.OleDbConnection) As ArrayList
        Dim aryColumnName As New ArrayList
        Dim schemaTable As DataTable
        Try

            '取得欄位的資訊
            schemaTable = connection.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Primary_Keys, _
                     New Object() {Nothing, Properties.getSchemaName().ToUpper, sTableName})
            For i As Integer = 0 To schemaTable.Rows.Count - 1
                'Console.WriteLine("PrimaryKeys=" & (schemaTable.Rows(i)!COLUMN_NAME.ToString)) 'COLUMN_NAME
                aryColumnName.Add((schemaTable.Rows(i)!COLUMN_NAME.ToString))
            Next i

            Return aryColumnName
        Catch ex As Exception

            Throw
        End Try
    End Function

#End Region

#End Region

#Region "getTable"

#Region "DatabaseManager"

    Public Overloads Shared Function getArrayTable(ByVal sTableName As String, ByVal databaseManager As DatabaseManager) As ArrayList
        Dim aryColumnName As New ArrayList
        Dim schemaTable As DataTable

        Try

            schemaTable = getTable(sTableName, databaseManager)

            For i As Integer = 0 To schemaTable.Rows.Count - 1
                'Console.WriteLine("ColumnName=" & (schemaTable.Rows(i)!ColumnName.ToString)) 'COLUMN_NAME
                aryColumnName.Add((schemaTable.Rows(i)!ColumnName.ToString))
            Next i

            Return aryColumnName
        Catch ex As Exception

            Throw
        End Try
    End Function

    Public Overloads Shared Function getTable(ByVal sTableName As String, ByVal databaseManager As DatabaseManager) As DataTable

        Dim schemaTable As DataTable = Nothing
        Dim dataReader As System.Data.IDataReader = Nothing
        Dim command As System.Data.IDbCommand = Nothing
        'Dim oracleParameter As Oracle.DataAccess.Client.OracleParameter
        Dim sSQL As String
        Try

            sSQL = "SELECT * FROM " & sTableName
            '取得欄位的資訊
            command = ProviderFactory.CreateCommand()
            command.Connection = databaseManager.getConnection

            If databaseManager.isTransaction Then
                command.Transaction = databaseManager.getTransaction
            End If

            command.CommandType = CommandType.Text
            command.CommandText = sSQL

            dataReader = command.ExecuteReader(CommandBehavior.SchemaOnly)

            'get the schema table
            schemaTable = dataReader.GetSchemaTable()

            Return schemaTable
        Catch ex As Exception

            Throw
        Finally
            If Not IsNothing(dataReader) Then dataReader.Close()
            If Not IsNothing(command) Then command.Dispose()
        End Try
    End Function

#End Region

#Region "IDbConnection"

    Public Overloads Shared Function getTable(ByVal sTableName As String, ByVal connection As System.Data.IDbConnection) As DataTable

        Dim schemaTable As DataTable = Nothing
        Dim dataReader As System.Data.IDataReader = Nothing
        Dim command As System.Data.IDbCommand = Nothing
        'Dim oracleParameter As Oracle.DataAccess.Client.OracleParameter
        Dim sSQL As String
        Try

            sSQL = "SELECT * FROM " & sTableName
            '取得欄位的資訊
            command = connection.CreateCommand

            command.CommandType = CommandType.Text
            command.CommandText = sSQL

            dataReader = command.ExecuteReader(CommandBehavior.SchemaOnly)

            'get the schema table
            schemaTable = dataReader.GetSchemaTable()

            Return schemaTable
        Catch ex As Exception

            Throw
        Finally
            If Not IsNothing(dataReader) Then dataReader.Close()
            If Not IsNothing(command) Then command.Dispose()
        End Try
    End Function

#End Region

#End Region

#Region "DB data type"

    'bool  
    'byte 
    'byte[]  
    'char      
    'DateTime  
    'Decimal  
    'double  
    'float 
    'Guid  
    'Int16  
    'Int32  
    'Int64  
    'object  
    'string  
    'TimeSpan  
    'UInt16  
    'UInt32  
    'UInt64 
    'ms-help://MS.MSDNQTR.2003FEB.1028/cpref/html/frlrfSystemDataDbTypeClassTopic.htm
    Public Shared Function getFieldType(ByVal sTtype As String) As System.Data.DbType
        Select Case sTtype
            Case "System.Boolean"
                Return DbUtility.getDBType(3)
            Case "System.Byte"
                Return DbUtility.getDBType(2)
            Case "System.Char"
                Return DbUtility.getDBType(16)
            Case "System.String"
                Return DbUtility.getDBType(16)
            Case "System.Decimal"
                Return DbUtility.getDBType(7)
            Case "System.DateTime"
                Return DbUtility.getDBType(6)
            Case "System.Double"
                Return DbUtility.getDBType(8)
            Case "System.Int32"
                Return DbUtility.getDBType(11)
            Case "System.Int16"
                Return DbUtility.getDBType(10)
            Case "System.Int64"
                Return DbUtility.getDBType(12)
            Case "System.Single"
                Return DbUtility.getDBType(15)
            Case "System.Guid"
                Return DbUtility.getDBType(9)
            Case "System.DateTime2"
                Return DbUtility.getDBType(26)
            Case "System.DateTimeOffset"
                Return DbUtility.getDBType(27)
            Case Else
                Throw New Exception("未知型別 sTtype=[" & sTtype & "]")
        End Select

    End Function

    Public Shared Function getDBType(ByVal iType As Integer) As System.data.DbType

        Select Case iType
            Case 0
                Return System.Data.DbType.AnsiString
            Case 1
                Return System.Data.DbType.Binary
            Case 2
                Return System.Data.DbType.Byte
            Case 3
                Return System.Data.DbType.Boolean
            Case 4
                Return System.Data.DbType.Currency
            Case 5
                Return System.Data.DbType.Date
            Case 6
                Return System.Data.DbType.DateTime
            Case 7
                Return System.Data.DbType.Decimal
            Case 8
                Return System.Data.DbType.Double
            Case 9
                Return System.Data.DbType.Guid
            Case 10
                Return System.Data.DbType.Int16
            Case 11
                Return System.Data.DbType.Int32
            Case 12
                Return System.Data.DbType.Int64
            Case 13
                Return System.Data.DbType.Object
            Case 14
                Return System.Data.DbType.[SByte]
            Case 15
                Return System.Data.DbType.Single
            Case 16
                Return System.Data.DbType.String
            Case 17
                Return System.Data.DbType.Time
            Case 18
                Return System.Data.DbType.UInt16
            Case 19
                Return System.Data.DbType.UInt32
            Case 20
                Return System.Data.DbType.UInt64
            Case 21
                Return System.Data.DbType.VarNumeric
            Case 22
                Return System.Data.DbType.AnsiStringFixedLength
            Case 23
                Return System.Data.DbType.StringFixedLength
            Case 26
                Return System.Data.DbType.DateTime2
            Case 27
                Return System.Data.DbType.DateTimeOffset
        End Select
    End Function

    Public Shared Function getOLETypeName(ByVal iType As Integer) As String

        Select Case iType
            Case 0
                Return "System.Data.OleDb.OleDbType.Empty"
            Case 2
                Return "System.Data.OleDb.OleDbType.SmallInt"
            Case 3
                Return "System.Data.OleDb.OleDbType.Integer"
            Case 4
                Return "System.Data.OleDb.OleDbType.Single"
            Case 5
                Return "System.Data.OleDb.OleDbType.Double"
            Case 6
                Return "System.Data.OleDb.OleDbType.Currency"
            Case 7
                Return "System.Data.OleDb.OleDbType.Date"
            Case 8
                Return "System.Data.OleDb.OleDbType.BSTR"
            Case 9
                Return "System.Data.OleDb.OleDbType.IDispatch"
            Case 10
                Return "System.Data.OleDb.OleDbType.Error"
            Case 11
                Return "System.Data.OleDb.OleDbType.Boolean"
            Case 12
                Return "System.Data.OleDb.OleDbType.Variant"
            Case 13
                Return "System.Data.OleDb.OleDbType.IUnknown"
            Case 14
                Return "System.Data.OleDb.OleDbType.Decimal"
            Case 16
                Return "System.Data.OleDb.OleDbType.TinyInt"
            Case 17
                Return "System.Data.OleDb.OleDbType.UnsignedTinyInt"
            Case 18
                Return "System.Data.OleDb.OleDbType.UnsignedSmallInt"
            Case 19
                Return "System.Data.OleDb.OleDbType.UnsignedInt"
            Case 21
                Return "System.Data.OleDb.OleDbType.UnsignedBigInt"
            Case 20
                Return "System.Data.OleDb.OleDbType.BigInt"
            Case 64
                Return "System.Data.OleDb.OleDbType.Filetime"
            Case 72
                Return "System.Data.OleDb.OleDbType.Guid"
            Case 128
                Return "System.Data.OleDb.OleDbType.Binary"
            Case 129
                Return "System.Data.OleDb.OleDbType.Char"
            Case 130
                Return "System.Data.OleDb.OleDbType.WChar"
            Case 131
                Return "System.Data.OleDb.OleDbType.Numeric"
            Case 133
                Return "System.Data.OleDb.OleDbType.DBDate"
            Case 134
                Return "System.Data.OleDb.OleDbType.DBTime"
            Case 135
                Return "System.Data.OleDb.OleDbType.DBTimeStamp"
            Case 138
                Return "System.Data.OleDb.OleDbType.PropVariant"
            Case 139
                Return "System.Data.OleDb.OleDbType.VarNumeric"
            Case 200
                Return "System.Data.OleDb.OleDbType.VarChar"
            Case 201
                Return "System.Data.OleDb.OleDbType.LongVarChar"
            Case 202
                Return "System.Data.OleDb.OleDbType.VarWChar"
            Case 203
                Return "System.Data.OleDb.OleDbType.LongVarWChar"
            Case 204
                Return "System.Data.OleDb.OleDbType.VarBinary"
            Case 205
                Return "System.Data.OleDb.OleDbType.LongVarBinary"
        End Select

        Return ""
    End Function

    Public Shared Function getOracleTypeName(ByVal iType As Integer) As String
        Select Case iType
            Case 1
                Return "System.Data.OracleClient.OracleType.BFile"
            Case 2
                Return "System.Data.OracleClient.OracleType.Blob"
            Case 23
                Return "System.Data.OracleClient.OracleType.Byte"
            Case 3
                Return "System.Data.OracleClient.OracleType.Char"
            Case 4
                Return "System.Data.OracleClient.OracleType.Clob"
            Case 5
                Return "System.Data.OracleClient.OracleType.Cursor"
            Case 6
                Return "System.Data.OracleClient.OracleType.DateTime"
            Case 30
                Return "System.Data.OracleClient.OracleType.Double"
            Case 29
                Return "System.Data.OracleClient.OracleType.Float"
            Case 27
                Return "System.Data.OracleClient.OracleType.Int16"
            Case 28
                Return "System.Data.OracleClient.OracleType.Int32"
            Case 7
                Return "System.Data.OracleClient.OracleType.IntervalDayToSecond"
            Case 8
                Return "System.Data.OracleClient.OracleType.IntervalYearToMonth"
            Case 9
                Return "System.Data.OracleClient.OracleType.LongRaw"
            Case 10
                Return "System.Data.OracleClient.OracleType.LongVarChar"
            Case 11
                Return "System.Data.OracleClient.OracleType.NChar"
            Case 12
                Return "System.Data.OracleClient.OracleType.NClob"
            Case 13
                Return "System.Data.OracleClient.OracleType.Number"
            Case 14
                Return "System.Data.OracleClient.OracleType.NVarChar"
            Case 15
                Return "System.Data.OracleClient.OracleType.Raw"
            Case 16
                Return "System.Data.OracleClient.OracleType.RowId"
            Case 26
                Return "System.Data.OracleClient.OracleType.SByte"
            Case 18
                Return "System.Data.OracleClient.OracleType.Timestamp"
            Case 19
                Return "System.Data.OracleClient.OracleType.TimestampLocal"
            Case 20
                Return "System.Data.OracleClient.OracleType.TimestampWithTZ"
            Case 24
                Return "System.Data.OracleClient.OracleType.UInt16"
            Case 25
                Return "System.Data.OracleClient.OracleType.UInt32"
            Case 22
                Return "System.Data.OracleClient.OracleType.VarChar"
        End Select
        Return ""
    End Function

    Public Shared Function getSqlTypeName(ByVal iType As Integer) As String
        Return getSqlDBType(iType).ToString
    End Function
    'System.Data.SqlDbType.BigInt()
    Public Shared Function getSqlDBType(ByVal iType As Integer) As System.Data.SqlDbType

        Select Case iType
            Case 0
                Return System.Data.SqlDbType.BigInt
            Case 1
                Return System.Data.SqlDbType.Binary
            Case 2
                Return System.Data.SqlDbType.Bit
            Case 3
                Return System.Data.SqlDbType.Char
            Case 31
                Return System.Data.SqlDbType.Date
            Case 4
                Return System.Data.SqlDbType.DateTime
            Case 33
                Return System.Data.SqlDbType.DateTime2
            Case 34
                Return System.Data.SqlDbType.DateTimeOffset
            Case 5
                Return System.Data.SqlDbType.Decimal
            Case 6
                Return System.Data.SqlDbType.Float
            Case 7
                Return System.Data.SqlDbType.Image
            Case 8
                Return System.Data.SqlDbType.Int
            Case 9
                Return System.Data.SqlDbType.Money
            Case 10
                Return System.Data.SqlDbType.NChar
            Case 11
                Return System.Data.SqlDbType.NText
            Case 12
                Return System.Data.SqlDbType.NVarChar
            Case 13
                Return System.Data.SqlDbType.Real
            Case 15
                Return System.Data.SqlDbType.SmallDateTime
            Case 16
                Return System.Data.SqlDbType.SmallInt
            Case 17
                Return System.Data.SqlDbType.SmallMoney
            Case 30
                Return System.Data.SqlDbType.Structured
            Case 18
                Return System.Data.SqlDbType.Text
            Case 32
                Return System.Data.SqlDbType.Time
            Case 19
                Return System.Data.SqlDbType.Timestamp
            Case 20
                Return System.Data.SqlDbType.TinyInt
            Case 29
                Return System.Data.SqlDbType.Udt
            Case 14
                Return System.Data.SqlDbType.UniqueIdentifier
            Case 21
                Return System.Data.SqlDbType.VarBinary
            Case 22
                Return System.Data.SqlDbType.VarChar
            Case 23
                Return System.Data.SqlDbType.Variant
            Case 25
                Return System.Data.SqlDbType.Xml
        End Select

        Return Nothing
    End Function

    Public Shared Function getODPTypeName(ByVal iType As Integer) As String
        Return getODPDBType(iType).ToString
    End Function

    Public Shared Function getODPDBType(ByVal iType As Integer) As Oracle.DataAccess.Client.OracleDbType
        Select Case iType
            Case 101
                Return Oracle.DataAccess.Client.OracleDbType.BFile
            Case 102
                Return Oracle.DataAccess.Client.OracleDbType.Blob
            Case 103
                Return Oracle.DataAccess.Client.OracleDbType.Byte
            Case 104
                Return Oracle.DataAccess.Client.OracleDbType.Char
            Case 105
                Return Oracle.DataAccess.Client.OracleDbType.Clob
            Case 106
                Return Oracle.DataAccess.Client.OracleDbType.Date
            Case 107
                Return Oracle.DataAccess.Client.OracleDbType.Decimal
            Case 108
                Return Oracle.DataAccess.Client.OracleDbType.Double
            Case 109
                Return Oracle.DataAccess.Client.OracleDbType.Long
            Case 110
                Return Oracle.DataAccess.Client.OracleDbType.LongRaw
            Case 111
                Return Oracle.DataAccess.Client.OracleDbType.Int16
            Case 112
                Return Oracle.DataAccess.Client.OracleDbType.Int32
            Case 113
                Return Oracle.DataAccess.Client.OracleDbType.Int64
            Case 114
                Return Oracle.DataAccess.Client.OracleDbType.IntervalDS
            Case 115
                Return Oracle.DataAccess.Client.OracleDbType.IntervalYM
            Case 116
                Return Oracle.DataAccess.Client.OracleDbType.NClob
            Case 117
                Return Oracle.DataAccess.Client.OracleDbType.NChar
            Case 119
                Return Oracle.DataAccess.Client.OracleDbType.NVarchar2
            Case 120
                Return Oracle.DataAccess.Client.OracleDbType.Raw
            Case 121
                Return Oracle.DataAccess.Client.OracleDbType.RefCursor
            Case 122
                Return Oracle.DataAccess.Client.OracleDbType.Single
            Case 123
                Return Oracle.DataAccess.Client.OracleDbType.TimeStamp
            Case 124
                Return Oracle.DataAccess.Client.OracleDbType.TimeStampLTZ
            Case 125
                Return Oracle.DataAccess.Client.OracleDbType.TimeStampTZ
            Case 126
                Return Oracle.DataAccess.Client.OracleDbType.Varchar2
            Case 127
                Return Oracle.DataAccess.Client.OracleDbType.XmlType
            Case 128
                Return Oracle.DataAccess.Client.OracleDbType.Array
            Case 129
                Return Oracle.DataAccess.Client.OracleDbType.Object
            Case 130
                Return Oracle.DataAccess.Client.OracleDbType.Ref
            Case 132
                Return Oracle.DataAccess.Client.OracleDbType.BinaryDouble
            Case 133
                Return Oracle.DataAccess.Client.OracleDbType.BinaryFloat
        End Select
        Return Nothing
    End Function

    Public Shared Function getMySqlTypeName(ByVal iType As Integer) As String

        Select Case iType
            Case 600
                Return "MySql.Data.MySqlClient.MySqlDbType.Binary"
            Case 16
                Return "MySql.Data.MySqlClient.MySqlDbType.Bit"
            Case 252
                Return "MySql.Data.MySqlClient.MySqlDbType.Blob"
            Case 1
                Return "MySql.Data.MySqlClient.MySqlDbType.Byte"
            Case 10
                Return "MySql.Data.MySqlClient.MySqlDbType.Date"
            Case 12
                Return "MySql.Data.MySqlClient.MySqlDbType.DateTime"
            Case 0
                Return "MySql.Data.MySqlClient.MySqlDbType.Decimal"
            Case 5
                Return "MySql.Data.MySqlClient.MySqlDbType.Double"
            Case 247
                Return "MySql.Data.MySqlClient.MySqlDbType.Enum"
            Case 4
                Return "MySql.Data.MySqlClient.MySqlDbType.Float"
            Case 255
                Return "MySql.Data.MySqlClient.MySqlDbType.Geometry"
            Case 800
                Return "MySql.Data.MySqlClient.MySqlDbType.Guid"
            Case 2
                Return "MySql.Data.MySqlClient.MySqlDbType.Int16"
            Case 9
                Return "MySql.Data.MySqlClient.MySqlDbType.Int24"
            Case 3
                Return "MySql.Data.MySqlClient.MySqlDbType.Int32"
            Case 8
                Return "MySql.Data.MySqlClient.MySqlDbType.Int64"
            Case 251
                Return "MySql.Data.MySqlClient.MySqlDbType.LongBlob"
            Case 751
                Return "MySql.Data.MySqlClient.MySqlDbType.LongText"
            Case 250
                Return "MySql.Data.MySqlClient.MySqlDbType.MediumBlob"
            Case 750
                Return "MySql.Data.MySqlClient.MySqlDbType.MediumText"
            Case 14
                Return "MySql.Data.MySqlClient.MySqlDbType.Newdate"
            Case 246
                Return "MySql.Data.MySqlClient.MySqlDbType.NewDecimal"
            Case 248
                Return "MySql.Data.MySqlClient.MySqlDbType.Set"
            Case 254
                Return "MySql.Data.MySqlClient.MySqlDbType.String"
            Case 752
                Return "MySql.Data.MySqlClient.MySqlDbType.Text"
            Case 11
                Return "MySql.Data.MySqlClient.MySqlDbType.Time"
            Case 7
                Return "MySql.Data.MySqlClient.MySqlDbType.Timestamp"
            Case 249
                Return "MySql.Data.MySqlClient.MySqlDbType.TinyBlob"
            Case 749
                Return "MySql.Data.MySqlClient.MySqlDbType.TinyText"
            Case 501
                Return "MySql.Data.MySqlClient.MySqlDbType.UByte"
            Case 502
                Return "MySql.Data.MySqlClient.MySqlDbType.UInt16"
            Case 509
                Return "MySql.Data.MySqlClient.MySqlDbType.UInt24"
            Case 503
                Return "MySql.Data.MySqlClient.MySqlDbType.UInt32"
            Case 508
                Return "MySql.Data.MySqlClient.MySqlDbType.UInt64"
            Case 601
                Return "MySql.Data.MySqlClient.MySqlDbType.VarBinary"
            Case 253
                Return "MySql.Data.MySqlClient.MySqlDbType.VarChar"
            Case 15
                Return "MySql.Data.MySqlClient.MySqlDbType.VarString"
            Case 12
                Return "MySql.Data.MySqlClient.MySqlDbType.Year"
        End Select
        Return ""
    End Function

#End Region
End Class
