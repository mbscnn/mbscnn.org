Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Data
Imports System.Diagnostics

'#If __MSSQL Then
Imports System.Data.SqlClient
'#End If


Namespace TinyDb

    '#If MSSQL Then

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Public Structure CmdParameter
        Public parameterName As String
        Public parameterValue As Object
        Public parameterIsPK As Boolean

        ''' <summary>
        ''' 使用指定的parameterName和parameterValue，初始化 CmdParameter 類別的新執行個體。
        ''' </summary>
        ''' <param name="name"></param>
        ''' <param name="value"></param>
        ''' <param name="isPK">
        ''' 可不填,isPK是InsertUpdate()及Update()時使用在判斷哪些筆記錄需更新。
        ''' 若設為TRUE，InsertUpdate()及Update()內組SQL語法時會將isPK=TRUE的參數加入至 [WHERE] 條件後。
        ''' 例如，將STAFFID="S123456"記錄的USERNAME設為"S234567"
        ''' dbTable.Update("USERINFO", New CmdParameter(){ New CmdParameter("STAFFID","S123456", True),
        '''                                                New CmdParameter("USERNAME","S234567")   });
        ''' Update()會重組SQL語法為
        ''' Update USERINFO SET USERNAME="S234567" Where STAFFID="S123456"; 
        ''' 
        ''' STAFFID是Where後的條件之一，因此要將isPK設為1
        ''' </param>
        ''' <remarks></remarks>
        Public Sub New(ByVal name As String, ByVal value As Object, Optional ByVal isPK As Boolean = False)
            parameterName = name

            If TypeOf value Is String Then
                If value = String.Empty Then
                    value = System.DBNull.Value
                End If
            ElseIf TypeOf value Is Integer Then
                If Not IsNumeric(value) Then
                    value = System.DBNull.Value
                End If
            ElseIf TypeOf value Is Decimal Then
                If Not IsNumeric(value) Then
                    value = System.DBNull.Value
                End If
            ElseIf TypeOf value Is Long Then
                If Not IsNumeric(value) Then
                    value = System.DBNull.Value
                End If
            ElseIf TypeOf value Is Double Then
                If Not IsNumeric(value) Then
                    value = System.DBNull.Value
                End If
            ElseIf TypeOf value Is Date Then
                If Not IsDate(value) Then
                    value = System.DBNull.Value
                End If
            ElseIf TypeOf value Is DateTime Then
                If Not IsDate(value) Then
                    value = System.DBNull.Value
                End If
            ElseIf TypeOf value Is DBNull Then
                value = System.DBNull.Value
            End If
            parameterValue = value
            parameterIsPK = isPK
        End Sub

    End Structure


    Public Structure SqlCommandQueueItem
        Public strSQL As String
        Public cmdParameter As CmdParameter()
    End Structure


    ''' <summary>
    ''' 簡易的資料庫存取 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    Public Class DbTable
        Protected m_dbTransaction As IDbTransaction
        Protected m_dbConnection As IDbConnection
        Protected m_queueDeferredSQL As New ArrayList

#If _FULL_TINYDB Then
        Public Sub New(dbConnection As IDbConnection)
            m_dbConnection = dbConnection
        End Sub

        Public Sub New(dbTransaction As IDbTransaction)
            m_dbTransaction = dbTransaction
            m_dbConnection = dbTransaction.Connection
        End Sub
#End If


        Public Sub New(dbManager As com.Azion.NET.VB.DatabaseManager)
            m_dbConnection = dbManager.getConnection()
            If dbManager.isTransaction Then
                m_dbTransaction = dbManager.getTransaction()
            End If
        End Sub

        Public Property DBTransaction() As IDbTransaction
            Get
                Return m_dbTransaction
            End Get

            Set(value As IDbTransaction)
                m_dbTransaction = value
            End Set
        End Property

        Public Property DbConnection() As IDbConnection
            Get
                Return m_dbConnection
            End Get

            Set(value As IDbConnection)
                m_dbConnection = value
            End Set
        End Property

        Public ReadOnly Property DeferredSQLCount() As Integer
            Get
                Return m_queueDeferredSQL.Count
            End Get
        End Property
        Protected Function GetParamter(ByVal parameterName As String, _
                                       ByVal parameterValue As Object) As CmdParameter
            Return New CmdParameter(parameterName, parameterValue)
        End Function

        Protected Function GetSqlCommandQueue(ByVal strSQL As String, _
                                         ByVal parameters As CmdParameter()) As SqlCommandQueueItem


            Dim sqlCommandQueue As New SqlCommandQueueItem
            Dim cmdParameterArrayList As New ArrayList


#If DEBUG Then
            Dim strDbg As String = strSQL
#End If

            If parameters IsNot Nothing Then
                For Each param In parameters
                    cmdParameterArrayList.Add(param)

                    strSQL = Microsoft.VisualBasic.Strings.Replace(strSQL, _
                        "@" & param.parameterName & "@", ProviderFactory.PositionPara & param.parameterName, _
                        CompareMethod.Text)

                    '組合顯示在追踪視窗上的SQL
#If DEBUG Then
                    If TypeOf param.parameterValue Is String Then
                        strDbg = Microsoft.VisualBasic.Strings.Replace(strDbg, _
                                    "@" & param.parameterName & "@", "'" & param.parameterValue.ToString() & "'", _
                                    CompareMethod.Text)
                    Else
                        strDbg = Microsoft.VisualBasic.Strings.Replace(strDbg, _
                                    "@" & param.parameterName & "@", param.parameterValue.ToString(), _
                                    CompareMethod.Text)
                    End If
#End If
                Next
            End If

#If DEBUG Then
            Trace.WriteLine(strDbg) '顯示SQL在追踪視窗上
#End If

            sqlCommandQueue.strSQL = strSQL
            sqlCommandQueue.cmdParameter = cmdParameterArrayList.ToArray(GetType(CmdParameter))
            Return sqlCommandQueue

        End Function


        Protected Function GetSqlCommand(ByVal strSQL As String, _
                                         ByVal parameters As CmdParameter(), _
                                         Optional ByVal commandType As CommandType = Nothing) As IDbCommand

            Dim sqlCommandQueue As SqlCommandQueueItem
            sqlCommandQueue = GetSqlCommandQueue(strSQL, parameters)

            Dim sqlCmd As SqlCommand

            sqlCmd = m_dbConnection.CreateCommand()
            sqlCmd.Parameters.Clear()

            If m_dbTransaction IsNot Nothing Then
                sqlCmd.Transaction = m_dbTransaction
            End If

            If Not IsNothing(commandType) Then
                sqlCmd.CommandType = commandType
            End If

            sqlCmd.CommandText = sqlCommandQueue.strSQL

            For Each param In sqlCommandQueue.cmdParameter
                sqlCmd.Parameters.AddWithValue(param.parameterName, param.parameterValue)
            Next

            Return sqlCmd
        End Function

        ''' <summary>
        ''' 執行查詢，並傳回DataReader
        ''' </summary>
        ''' <param name="strSQL"></param>
        ''' <param name="parameters"></param>
        ''' <param name="commandType"></param>
        ''' <returns>傳回IDataReader，需自行釋放傳回的IDataReader物件 (IDataReader.Dispose())</returns>
        ''' <remarks></remarks>
        Public Function ExecuteReader(ByVal strSQL As String, _
                                      ByVal parameters As CmdParameter(), _
                                      Optional ByVal commandType As CommandType = Nothing) As IDataReader
            Return GetSqlCommand(strSQL, parameters, commandType).ExecuteReader()
        End Function


        ''' <summary>
        ''' 將查詢的資訊寫入至DataTable內，由DataReader轉入至DataTable
        ''' </summary>
        ''' <param name="strSQL"></param>
        ''' <param name="parameters"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetDataTable(ByVal strSQL As String, _
                                     ByVal parameters As CmdParameter()) As DataTable
            Using dr As IDataReader = ExecuteReader(strSQL, parameters)
                Dim dt As New DataTable()
                dt.Load(dr)
                Return dt
            End Using
        End Function

        ''' <summary>
        ''' 執行查詢，並傳回查詢所傳回的結果集第一個資料列的第一個資料行。會忽略其他的資料行或資料列。
        ''' </summary>
        ''' <param name="strSQL"></param>
        ''' <param name="parameters"></param>
        ''' <param name="commandType"></param>
        ''' <returns>結果集中第一個資料列的第一個資料行；如果結果集是空的，則為 null 參考 (在 Visual Basic 中為 Nothing)。最多傳回 2033 個字元。</returns>
        ''' <remarks></remarks>
        ''' <sample>
        ''' '是否是管理單位
        ''' '如果是管理單位，則可以設定所有人的密碼
        ''' Dim dbTable As New DbTable(dbManager.getConnection)
        ''' Dim objInt As Integer
        ''' objInt = dbTable.ExecuteScalar( _
        '''      "SELECT Count(*) from ROLEGROUP_GRANT_BRID where BRID = @BRID@", _
        '''      New CmdParameter() {New CmdParameter("BRID", Session("BRID"))})
        ''' 
        ''' If (objInt = 0) Then
        '''      '不是管理單位
        ''' Else
        '''      '是管理單位
        ''' End If        ''' 
        ''' </sample>
        Public Function ExecuteScalar(ByVal strSQL As String, _
                                      ByVal parameters As CmdParameter(), _
                                      Optional ByVal commandType As CommandType = Nothing) As Object
            Return GetSqlCommand(strSQL, parameters, commandType).ExecuteScalar()
        End Function

        ''' <summary>
        ''' 針對連接執行陳述式，並傳回受影響的資料列數目。
        ''' </summary>
        ''' <param name="strSQL"></param>
        ''' <param name="parameters"></param>
        ''' <param name="commandType"></param>
        ''' <param name="deferredExecute"></param>
        ''' <returns>受影響的資料列數目。</returns>
        ''' <remarks></remarks>
        ''' <SAMPLE>
        ''' Dim table as new DbTable(dbManager.getConnection)
        ''' 
        ''' table.ExecuteNonQuert ("Delete from user where USERID = @USERID@ AND USERNAME = @USERNAME@", 
        '''                         New CmdParameter(){ New CmdParameter("USERID","S123456"),
        '''                                             New CmdParameter("USERNAME","S234567")   })
        ''' </SAMPLE>
        Public Function ExecuteNonQuery(ByVal strSQL As String, _
                                        ByVal parameters As CmdParameter(), _
                                        Optional ByVal commandType As CommandType = Nothing, _
                                        Optional ByVal deferredExecute As Boolean = False) As Integer


            'rename parameter for multiple sql command

            If deferredExecute = True OrElse m_queueDeferredSQL.Count > 0 Then

                Dim paramAppend As String = "_R" & m_queueDeferredSQL.Count().ToString()
                Dim sOldName, sNewName As String

                For nIndex As Integer = 0 To parameters.Count - 1
                    sOldName = parameters(nIndex).parameterName
                    sNewName = parameters(nIndex).parameterName & paramAppend

                    parameters(nIndex).parameterName = sNewName
                    strSQL = strSQL.Replace("@" & sOldName & "@", "@" & sNewName & "@")
                Next
            End If

            ' Insert command to Queue
            m_queueDeferredSQL.Add(GetSqlCommandQueue(strSQL, parameters))

            ' If it is deferred execution, return 0 
            If deferredExecute = True OrElse m_queueDeferredSQL.Count = 0 Then
                Return 0
            End If

            ' 組全部的SQL及parameters
            Dim strAllSql As New StringBuilder
            Dim laParameters As New ArrayList

            For Each deferredSQL As SqlCommandQueueItem In m_queueDeferredSQL
                If strAllSql.Length > 0 Then
                    strAllSql.Append(" ; ")
                End If

                strAllSql.Append(deferredSQL.strSQL)

                For Each param In deferredSQL.cmdParameter
                    laParameters.Add(param)
                Next
            Next

            Dim sqlCmd As SqlCommand

            sqlCmd = m_dbConnection.CreateCommand()
            sqlCmd.Parameters.Clear()

            If m_dbTransaction IsNot Nothing Then
                sqlCmd.Transaction = m_dbTransaction
            End If

            If Not IsNothing(commandType) Then
                sqlCmd.CommandType = commandType
            End If

            sqlCmd.CommandText = strAllSql.ToString()
            For Each param As CmdParameter In laParameters
                sqlCmd.Parameters.AddWithValue(param.parameterName, param.parameterValue)
            Next

            '清空DeferredSQL
            m_queueDeferredSQL.Clear()

            Return sqlCmd.ExecuteNonQuery()
        End Function


    End Class
End Namespace
