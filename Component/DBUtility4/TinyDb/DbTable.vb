Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Data
Imports System.Diagnostics
Imports com.Azion.EloanUtility

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

#If Not _TINYDB Then
        Protected m_dbManager As com.Azion.NET.VB.DatabaseManager
#End If

        Protected m_dbTransaction As IDbTransaction
        Protected m_dbConnection As IDbConnection
        Protected m_queueDeferredSQL As New ArrayList

        ''' <summary>
        ''' TABLE的名稱
        ''' </summary>
        ''' <remarks></remarks>
        Public m_strTableName As String


        ''' <summary>
        ''' 當搜尋到指字串時，中斷執行 
        ''' </summary>
        ''' strDebugBreak(0) = "Update FlowStep S003835"
        ''' strDebugBreak(1) = "Insert FlowStep S003835"
        ''' 當更新或新增至FLOWSTEP內的記錄有S003835的員編就中斷
        ''' <remarks></remarks>
        Public Shared strDebugBreak(10) As String

#If _TINYDB Then
        Public Sub New(ByVal strTableName As String, ByVal dbConnection As IDbConnection)
            m_strTableName = strTableName
            m_dbConnection = dbConnection
        End Sub

        Public Sub New(ByVal strTableName As String, ByVal dbTransaction As IDbTransaction)
            m_strTableName = strTableName
            m_dbTransaction = dbTransaction
            m_dbConnection = dbTransaction.Connection
        End Sub


        Public Sub New(ByVal dbConnection As IDbConnection)
            m_dbConnection = dbConnection
        End Sub

        Public Sub New(ByVal dbTransaction As IDbTransaction)
            m_dbTransaction = dbTransaction
            m_dbConnection = dbTransaction.Connection
        End Sub
#Else
        Public Sub New(ByVal strTableName As String, dbManager As com.Azion.NET.VB.DatabaseManager)
            m_strTableName = strTableName
            m_dbManager = dbManager

            If Not IsNothing(dbManager) Then
                m_dbConnection = dbManager.getConnection()
                If dbManager.isTransaction Then
                    m_dbTransaction = dbManager.getTransaction()
                End If
            End If
        End Sub

        Public Sub New(dbManager As com.Azion.NET.VB.DatabaseManager)
            m_dbManager = dbManager

            If Not IsNothing(dbManager) Then
                m_dbConnection = dbManager.getConnection()
                If dbManager.isTransaction Then
                    m_dbTransaction = dbManager.getTransaction()
                End If
            End If
        End Sub
#End If

        Public Shared Function getNewBosBase(ByVal databaseManager As DatabaseManager) As DbTable
            Return New DbTable(databaseManager)
        End Function


        Public Property DBTransaction() As IDbTransaction
            Get
                Return m_dbTransaction
            End Get

            Set(value As IDbTransaction)
                Dbg.Assert(IsNothing(m_dbTransaction) OrElse IsNothing(m_dbTransaction.Connection))

                If Not IsNothing(m_dbTransaction) Then
                    If Not IsNothing(m_dbTransaction.Connection) Then
                        m_dbTransaction.Connection.Dispose()
                    End If
                End If

                m_dbConnection = value.Connection
                m_dbTransaction = value
                m_dbManager = Nothing
            End Set
        End Property

        Public Property DbConnection() As IDbConnection
            Get
                Return m_dbConnection
            End Get

            Set(value As IDbConnection)

                Dbg.Assert(IsNothing(m_dbConnection))

                If Not IsNothing(m_dbConnection) Then
                    m_dbConnection.Dispose()
                End If

                m_dbConnection = value
                m_dbTransaction = Nothing
                m_dbManager = Nothing
            End Set
        End Property

        Public Property DatabaseManager() As com.Azion.NET.VB.DatabaseManager
            Get
                Return m_dbManager
            End Get

            Set(value As com.Azion.NET.VB.DatabaseManager)

                Dbg.Assert(IsNothing(m_dbManager))

                If Not IsNothing(m_dbManager) Then
                    m_dbManager.dispose()
                End If

                m_dbConnection = Nothing
                m_dbTransaction = Nothing
                m_dbManager = value

                If Not IsNothing(m_dbManager) Then
                    m_dbConnection = m_dbManager.getConnection()
                    If m_dbManager.isTransaction Then
                        m_dbTransaction = m_dbManager.getTransaction()
                    End If
                End If

            End Set
        End Property

        Public ReadOnly Property DeferredSQLCount() As Integer
            Get
                Return m_queueDeferredSQL.Count
            End Get
        End Property

        Protected Function PARAMETER(ByVal parameterName As String, _
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
                        If Not IsDBNull(param.parameterValue) AndAlso Not IsNothing(param.parameterValue) Then
                            strDbg = Microsoft.VisualBasic.Strings.Replace(strDbg, _
                                        "@" & param.parameterName & "@", param.parameterValue.ToString(), _
                                        CompareMethod.Text)
                        Else
                            strDbg = Microsoft.VisualBasic.Strings.Replace(strDbg, _
                                        "@" & param.parameterName & "@", "NULL", _
                                        CompareMethod.Text)
                        End If
                    End If
#End If
                Next
            End If

#If DEBUG Then
            Trace.WriteLine(strDbg) '顯示SQL在追踪視窗上

            '當特定SQL語法時中斷執行
            Dim bBreak As Boolean = False

            For Each sBreak In strDebugBreak

                If IsNothing(sBreak) OrElse sBreak.Length = 0 Then
                    Continue For
                End If

                Dim sSplit() As String = Split(sBreak)
                Dim bFound As Boolean = True

                For Each ssSplit In sSplit
                    If strDbg.IndexOf(ssSplit) = -1 Then
                        bFound = False
                        Exit For
                    End If
                Next

                If bFound = True Then
                    bBreak = True
                    Exit For
                End If
            Next

            If bBreak Then
                Dbg.Break()
            End If
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
        Public Overridable Function GetDataTable(ByVal strSQL As String, _
                                     ByVal parameters As CmdParameter(), _
                                     Optional ByVal commandType As CommandType = Nothing) As DataTable
            Using dr As IDataReader = ExecuteReader(strSQL, parameters, commandType)
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
        Public Overridable Function ExecuteScalar(ByVal strSQL As String, _
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
        Public Overridable Function ExecuteNonQuery(ByVal strSQL As String, _
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
                If Not IsDBNull(param.parameterValue) AndAlso Not IsNothing(param.parameterValue) Then
                    sqlCmd.Parameters.AddWithValue(param.parameterName, param.parameterValue)
                Else
                    sqlCmd.Parameters.AddWithValue(param.parameterName, Convert.DBNull)
                End If
            Next

            '清空DeferredSQL
            m_queueDeferredSQL.Clear()

            Return sqlCmd.ExecuteNonQuery()
        End Function


    End Class
End Namespace
