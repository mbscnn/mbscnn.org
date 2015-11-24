Option Explicit On
Option Strict On
Imports System.Runtime.CompilerServices
Imports System.Runtime.Serialization.Formatters.Binary
Imports System.Data
Imports System.Collections
Imports System.IO
Imports System.Xml.Serialization
Imports System.Text
Imports com.Azion.EloanUtility
Imports System.Text.RegularExpressions

''' <summary>
''' 
''' </summary>
''' <remarks></remarks>
Public Class BosParameter
    Public parameterName As String
    Public parameterValue As Object

    ''' <summary> 
    ''' 使用指定的parameterName和parameterValue，初始化 CmdParameter 類別的新執行個體。
    ''' </summary>
    ''' <param name="name"></param>
    ''' <param name="value"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal name As String, ByVal value As Object)
        parameterName = name

        If TypeOf value Is String Then
            If value Is String.Empty Then
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
    End Sub

End Class


Public Class StructDbMetaCacheItem
    Public m_arrayPrimaryKeys As HashSet(Of String)
    Public m_attributes As Hashtable
    Public m_bosAttributeList As BosAttributeList
End Class


Public Class BosParamsList
    Inherits List(Of BosParameter)

    Public Overloads Function Add(ByVal parameterName As String, ByVal parameterValue As Object) As String
        Add(New BosParameter(parameterName, parameterValue))
        Return parameterName
    End Function

    Public Overloads Function Add(ByVal parameterValue As Object) As String
        Dim paramName As String
        paramName = "DYMPARAM_" & Count.ToString

        Add(New BosParameter(paramName, parameterValue))
        Return ProviderFactory.PositionPara & paramName
    End Function

    Public Function ToIParameterArray() As IDataParameter()
        Dim iDP As New List(Of IDataParameter)

        For Each bosParam As BosParameter In Me
            iDP.Add(ProviderFactory.CreateDataParameter(bosParam.parameterName, bosParam.parameterValue))
        Next

        Return iDP.ToArray

    End Function

End Class


Public Class BosTable
    Dim m_Tables As Dictionary(Of String, BosBase)
    Dim m_DatabaseManager As DatabaseManager

    Sub New(ByVal databaseManager As DatabaseManager)
        m_DatabaseManager = databaseManager
    End Sub

    Default ReadOnly Property myProperty(ByVal tableName As String) As BosBase
        Get
            If IsNothing(m_Tables) Then
                m_Tables = New Dictionary(Of String, BosBase)
            End If

            If m_Tables.ContainsKey(tableName) = False Then
                m_Tables(tableName) = New BosBase(tableName, m_DatabaseManager)
            End If

            Return m_Tables(tableName)
        End Get
    End Property
End Class


Public Class SQLTrace
    Public SQL As String
    Public Parameter As BosParameter()

    Public Shared Function Create(ByVal sql As String, ByVal params As BosParameter()) As SQLTrace
        Dim sqlTrace As New SQLTrace
        sqlTrace.SQL = sql
        sqlTrace.Parameter = params
        Return sqlTrace
    End Function
End Class


Partial Public Class BosBase
    Implements IEnumerable, IEnumerator ', System.ICloneable

    Protected Shared m_metaCache As Hashtable
    Public Shared SQL_DUMP As Boolean = True

    '一個SQL指令最多只能更新1000筆資料
    Public m_nNoQueryThreshold As Integer = 10000

    '單一物件內最後執行的SQL
    Protected m_TraceLastSQL As SQLTrace

    '所有物件最後執行的SQL，Group By ThreadId
    Protected Shared m_GlobalTraceLastSQL As New Dictionary(Of Integer, SQLTrace)
    Public Shared Property GlobalTraceLastSQL() As SQLTrace
        Get
            Try
                Return m_GlobalTraceLastSQL(System.Threading.Thread.CurrentThread.ManagedThreadId)
            Catch ex As Exception
                Return Nothing
            End Try
        End Get

        Set(ByVal value As SQLTrace)
            m_GlobalTraceLastSQL(System.Threading.Thread.CurrentThread.ManagedThreadId) = value
        End Set
    End Property


    Sub New(ByVal sTableName As String)
        setTableName(sTableName)
    End Sub


    ''' <summary>
    ''' 初始化 BosBase 類別的新執行個體 (Instance)。
    ''' </summary>
    ''' <param name="sTableName">傳入Table Name(大寫)</param>
    ''' <param name="databaseManager">傳入 DatabaseManager</param>
    ''' <remarks></remarks>
    '    Sub New(ByVal sTableName As String, ByVal databaseManager As DatabaseManager)

    '        setDatabaseManager(databaseManager)
    '        setTableName(sTableName)

    '        '如果沒有，建立新HASH來保留TABLE的資訊
    '        If IsNothing(m_metaCache) Then
    '            m_metaCache = New Hashtable(2000)
    '        End If

    '        Dim dbMetaCacheItem As StructDbMetaCacheItem


    '        If m_metaCache.ContainsKey(sTableName) Then
    '            ' 從CAHCE中取回資料，如果資料已在快取內
    '            dbMetaCacheItem = CType(m_metaCache(sTableName), StructDbMetaCacheItem)

    '            m_arrayPrimaryKeys = (CType(dbMetaCacheItem.m_arrayPrimaryKeys, HashSet(Of String)))
    '            m_attributes = dbMetaCacheItem.m_attributes

    '            Dim attributes As New Hashtable(m_attributes.Count)
    '            Dim bosAttributeList As New BosAttributeList
    '            Dim bosAttribute As BosAttribute

    '            For Each attribute_key As String In m_attributes.Keys
    '                bosAttribute = CType(m_attributes(attribute_key), BosAttribute).newInstance()
    '                attributes.Add(attribute_key.Clone(), bosAttribute)
    '                bosAttributeList.add(bosAttribute)
    '            Next

    '            m_attributes = attributes
    '            m_bosAttributeList = bosAttributeList

    '        Else
    '            ' 第一次從DB取回資料，並存在快取內
    '            Me.initMetaData(sTableName)

    '            dbMetaCacheItem = New StructDbMetaCacheItem
    '            dbMetaCacheItem.m_arrayPrimaryKeys = m_arrayPrimaryKeys
    '            dbMetaCacheItem.m_attributes = m_attributes
    '            dbMetaCacheItem.m_bosAttributeList = m_bosAttributeList

    '            m_metaCache(sTableName) = dbMetaCacheItem
    '        End If

    '        '避免快取內的資料被修改
    '        m_attributes = CType(m_attributes.Clone, Hashtable)

    '#If DEBUG Then
    '        'Dbg.Assert(m_arrayPrimaryKeys.Count > 0)
    '        Dbg.Assert(m_attributes.Count > 0 AndAlso m_bosAttributeList.size > 0, "Table的欄位需大於0")
    '#End If
    '    End Sub


    Public Shared Function getNewBosBase(ByVal sTableName As String, ByVal databaseManager As DatabaseManager) As BosBase
        Return New BosBase(sTableName, databaseManager)
    End Function


    Public Sub setDbManager(ByVal databaseManager As DatabaseManager)
        Me.m_databaseManager = databaseManager
    End Sub

    Public Shared ReadOnly Property PARAMETER(ByVal parameterName As String, _
                       ByVal parameterValue As Object) As BosParameter
        Get
            Return New BosParameter(parameterName, parameterValue)
        End Get
    End Property



    Public Shared ReadOnly Property PARAM_ARRAY(ByVal ParamArray objs() As Object) As BosParameter()
        Get
            Dim arrayBosParameters As New BosParamsList

            Dbg.Assert((objs.Count Mod 2) = 0, "參數數目不正確")

            For i As Integer = 0 To UBound(objs, 1) Step 2
                arrayBosParameters.Add(CStr(objs(i)), objs(i + 1))
            Next i

            Return arrayBosParameters.ToArray()
        End Get
    End Property



    Private Function ExecuteReader(ByVal strSQL As String, _
                                     ByVal parameters() As BosParameter, _
                                     Optional ByVal commandType As CommandType = CommandType.Text) As IDataReader

        Dim nParametersCount As Integer = -1

        If IsNothing(parameters) = False Then
            nParametersCount = parameters.Count - 1
        End If

        Dim para(nParametersCount) As System.Data.IDbDataParameter
        Dim sbSql As New StringBuilder(strSQL)

        If IsNothing(parameters) = False Then
            For n As Integer = 0 To parameters.Count - 1
                sbSql.Replace("@" & parameters(n).parameterName & "@", ProviderFactory.PositionPara & parameters(n).parameterName)
                para(n) = ProviderFactory.CreateDataParameter(parameters(n).parameterName, parameters(n).parameterValue)
            Next
        End If

        TRACE(strSQL, parameters)

        Return CType(ExecuteReader(commandType, sbSql.ToString(), para), IDataReader)
    End Function

    Public Overridable Function ExecuteNonQuery(ByVal strSQL As String, _
                                                ByVal ParamArray objs() As Object) As Integer

        Return ExecuteNonQuery(strSQL, PARAM_ARRAY(objs), CommandType.Text)
    End Function


    Public Overridable Function ExecuteNonQuery(ByVal strSQL As String, _
                                ByVal parameters() As BosParameter, _
                                Optional ByVal commandType As CommandType = CommandType.Text) As Integer

        Dim nParametersCount As Integer = -1
        Dim nResult As Integer

        If IsNothing(parameters) = False Then
            nParametersCount = parameters.Count - 1
        End If

        Dim para(nParametersCount) As System.Data.IDbDataParameter
        Dim sbSql As New StringBuilder(strSQL)

        'If IsNothing(parameters) = False Then
        For n As Integer = 0 To nParametersCount
            sbSql.Replace("@" & parameters(n).parameterName & "@", ProviderFactory.PositionPara & parameters(n).parameterName)
            para(n) = ProviderFactory.CreateDataParameter(parameters(n).parameterName, _
                                                          IIf(IsNothing(parameters(n).parameterValue), Convert.DBNull, parameters(n).parameterValue))
        Next
        'End If

        TRACE(strSQL, parameters)

        nResult = ExecuteNonQuery(commandType, sbSql.ToString(), para)

        If nResult > m_nNoQueryThreshold Then
            Throw New Exception("更新超過" & m_nNoQueryThreshold.ToString & "筆資料")
        End If

        Return nResult
    End Function

    Public Overridable Function ExecuteScalar(ByVal strSQL As String, _
                                            ByVal ParamArray objs() As Object) As Object

        Return ExecuteScalar(strSQL, PARAM_ARRAY(objs), CommandType.Text)
    End Function

    Public Overridable Function ExecuteScalar(ByVal strSQL As String, _
                              ByVal parameters() As BosParameter, _
                              Optional ByVal commandType As CommandType = CommandType.Text) As Object

        Dim nParametersCount As Integer = -1

        strSQL = strSQL.Trim

        If String.IsNullOrEmpty(strSQL) = False AndAlso Left(strSQL, 1) = "[" AndAlso Right(strSQL, 1) = "]" Then

            Dim sFieldName As String = Mid(strSQL, 2, strSQL.Length - 2)
            Dim sTableName As String = getTableName()

            Dim rgx As New Regex("\[(?<TABLE>\w*)\]\.\[(?<FIELD>\w*)\]", RegexOptions.IgnoreCase)
            Dim matches As MatchCollection = rgx.Matches(strSQL)

            For Each match As Match In matches
                Dim groups As GroupCollection = match.Groups
                '取得TABLE及FILED
                sTableName = groups.Item("TABLE").Value
                sFieldName = groups.Item("FIELD").Value
            Next

            Dim sb As New StringBuilder("select top 1 " & sFieldName & " from " & sTableName & vbCrLf)

            Dim bAnd As Boolean = False

            If (IsNothing(parameters) = False AndAlso parameters.Length > 0) Then
                sb.Append(" where ")

                For Each param As BosParameter In parameters
                    If bAnd Then
                        sb.Append("   and")
                    End If
                    'sb.Append(" " & param.parameterName & " = @" & param.parameterName & "@" & vbCrLf)
                    If IsNothing(param.parameterValue) = True OrElse IsDBNull(param.parameterValue) = True Then
                        sb.Append(" " & param.parameterName & " is null" & vbCrLf)
                    Else
                        sb.Append(" " & param.parameterName & " = @" & param.parameterName & "@" & vbCrLf)
                    End If
                    bAnd = True
                Next
            End If

            strSQL = sb.ToString
        End If

        If IsNothing(parameters) = False Then
            nParametersCount = parameters.Count - 1
        End If

        Dim para(nParametersCount) As System.Data.IDbDataParameter
        Dim sbSql As New StringBuilder(strSQL)

        'If IsNothing(parameters) = False Then
        For n As Integer = 0 To nParametersCount
            sbSql.Replace("@" & parameters(n).parameterName & "@", ProviderFactory.PositionPara & parameters(n).parameterName)
            para(n) = ProviderFactory.CreateDataParameter(parameters(n).parameterName, parameters(n).parameterValue)
        Next
        'End If

        TRACE(strSQL, parameters)

        Return ExecuteScalar(commandType, sbSql.ToString, para)
    End Function


    ''' <summary>
    ''' 將查詢的資訊寫入至DataTable內，由DataReader轉入至DataTable
    ''' </summary>
    ''' <param name="strSQL">若為NOTHING，程式會自行組SQL字串</param>
    ''' <param name="objs"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetDataTable(ByVal strSQL As String, _
                                         ByVal ParamArray objs() As Object) As DataTable
        Return GetDataTable(strSQL, PARAM_ARRAY(objs))
    End Function


    ''' <summary>
    ''' 將查詢的資訊寫入至DataTable內，由DataReader轉入至DataTable
    ''' </summary>
    ''' <param name="strSQL">若為NOTHING，程式會自行組SQL字串</param>
    ''' <param name="parameters"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overridable Function GetDataTable(ByVal strSQL As String, _
                                 ByVal parameters As BosParameter(), _
                                 Optional ByVal commandType As CommandType = CommandType.Text) As DataTable
        Try
            If String.IsNullOrEmpty(strSQL) Then

                Dim sb As New StringBuilder("select * from " & getTableName() & vbCrLf)
                Dim bAnd As Boolean = False

                If (IsNothing(parameters) = False AndAlso parameters.Length > 0) Then
                    sb.Append(" where ")

                    For Each param As BosParameter In parameters
                        If bAnd Then
                            sb.Append("   and")
                        End If

                        If Left(param.parameterName, 1) = "!" Then
                            param.parameterName = Mid(param.parameterName, 2)

                            If IsNothing(param.parameterValue) = True OrElse IsDBNull(param.parameterValue) = True Then
                                sb.Append(" " & param.parameterName & " is not null" & vbCrLf)
                            Else
                                sb.Append(" " & param.parameterName & " <> @" & param.parameterName & "@" & vbCrLf)
                            End If
                        Else
                            If IsNothing(param.parameterValue) = True OrElse IsDBNull(param.parameterValue) = True Then
                                sb.Append(" " & param.parameterName & " is null" & vbCrLf)
                            Else
                                sb.Append(" " & param.parameterName & " = @" & param.parameterName & "@" & vbCrLf)
                            End If
                        End If

                        bAnd = True
                    Next
                End If

                strSQL = sb.ToString()
            Else
                strSQL = strSQL.Trim

                If String.IsNullOrEmpty(strSQL) = False AndAlso Left(strSQL, 1) = "[" AndAlso Right(strSQL, 1) = "]" Then
                    Dim sFieldName As String = Mid(strSQL, 2, strSQL.Length - 2)
                    Dim sTableName As String = getTableName()

                    Dim rgx As New Regex("\[(?<TABLE>\w*)\]\.\[(?<FIELD>\w*)\]", RegexOptions.IgnoreCase)
                    Dim matches As MatchCollection = rgx.Matches(strSQL)

                    For Each match As Match In matches
                        Dim groups As GroupCollection = match.Groups
                        '取得TABLE及FILED
                        sTableName = groups.Item("TABLE").Value
                        sFieldName = groups.Item("FIELD").Value
                    Next

                    Dim sb As New StringBuilder("select " & sFieldName & " from " & sTableName & vbCrLf)
                    Dim bAnd As Boolean = False

                    If (IsNothing(parameters) = False AndAlso parameters.Length > 0) Then
                        sb.Append(" where ")

                        For Each param As BosParameter In parameters
                            If bAnd Then
                                sb.Append("   and")
                            End If

                            If Left(param.parameterName, 1) = "!" Then
                                param.parameterName = Mid(param.parameterName, 2)

                                If IsNothing(param.parameterValue) = True OrElse IsDBNull(param.parameterValue) = True Then
                                    sb.Append(" " & param.parameterName & " is not null" & vbCrLf)
                                Else
                                    sb.Append(" " & param.parameterName & " <> @" & param.parameterName & "@" & vbCrLf)
                                End If
                            Else
                                If IsNothing(param.parameterValue) = True OrElse IsDBNull(param.parameterValue) = True Then
                                    sb.Append(" " & param.parameterName & " is null" & vbCrLf)
                                Else
                                    sb.Append(" " & param.parameterName & " = @" & param.parameterName & "@" & vbCrLf)
                                End If
                            End If

                            bAnd = True
                        Next
                    End If

                    strSQL = sb.ToString
                End If
            End If

            Using dr As IDataReader = ExecuteReader(strSQL, parameters, commandType)
                Dim dt As New DataTable()
                dt.Load(dr)
                Return dt
            End Using
        Catch ex As Exception
            Throw
        End Try

    End Function



    Public Function InsertUpdate(ByVal ParamArray objs() As Object) As Integer
        Return InsertUpdate(PARAM_ARRAY(objs))
    End Function


    ''' <summary>
    ''' 若已有記錄，則更新記錄。否則新增記錄。
    ''' </summary>
    ''' <param name="parameters"></param>
    ''' <returns>受影響的資料列數目。</returns>
    ''' <remarks>會以PK來判斷是否是新記錄還是舊記錄</remarks>
    Public Function InsertUpdate(ByVal parameters As BosParameter()) As Integer
#If DEBUG Then
        Dbg.Assert(IsNothing(getTableName()) = False AndAlso getTableName().Length > 0, "Table不存在")
#End If
        Return InsertUpdate(getTableName(), parameters)
    End Function


    Public Function InsertUpdate(ByVal datarow As DataRow) As Integer

        Dim arrayBosParameters As New BosParamsList

        For Each column As DataColumn In datarow.Table.Columns
            If Not IsNothing(datarow(column.ColumnName)) Then
                arrayBosParameters.Add(column.ColumnName, datarow(column.ColumnName))
            End If
        Next

        Return InsertUpdate(arrayBosParameters.ToArray())

    End Function


    ''' <summary>
    ''' 若已有記錄，則更新記錄。否則新增記錄。
    ''' </summary>
    ''' <param name="tableName"></param>
    ''' <param name="parameters"></param>
    ''' <returns>受影響的資料列數目。</returns>
    ''' <remarks>會以PK來判斷是否是新記錄還是舊記錄</remarks>
    ''' <SAMPLE>
    ''' </SAMPLE>
    Public Function InsertUpdate(ByVal tableName As String, ByVal parameters As BosParameter()) As Integer
        Dim nValue As Integer = 0

        nValue = Update(tableName, parameters)

        If nValue = 0 Then
            m_databaseManager.m_nWriteCount = Math.Max(0, m_databaseManager.m_nWriteCount - 1)
            nValue = Insert(tableName, parameters)
        End If

#If DEBUG Then
        Dbg.Assert(nValue > 0, "至少要有一行被更新或新增")
        '至少要有一行被更新或新增
#End If
        Return nValue

    End Function

    ''' <summary>
    ''' 新資資料
    ''' </summary>
    ''' <param name="objs"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Insert(ByVal ParamArray objs() As Object) As Integer
        Return Insert(PARAM_ARRAY(objs))
    End Function


    Public Function Insert(ByVal parameters As BosParameter()) As Integer

#If DEBUG Then
        Dbg.Assert(IsNothing(getTableName()) = False AndAlso getTableName().Length > 0, "Table不存在")
#End If
        Return Insert(getTableName(), parameters)
    End Function

    Public Function Insert(ByVal datarow As DataRow) As Integer

        Dim arrayBosParameters As New BosParamsList

        For Each column As DataColumn In datarow.Table.Columns
            If Not IsNothing(datarow(column.ColumnName)) Then
                arrayBosParameters.Add(column.ColumnName, datarow(column.ColumnName))
            End If
        Next

        Return Insert(arrayBosParameters.ToArray())

    End Function


    ''' <summary>
    ''' 新增記錄
    ''' </summary>
    ''' <param name="tableName"></param>
    ''' <param name="parameters"></param>
    ''' <returns>受影響的資料列數目</returns>
    ''' <remarks></remarks>
    Public Function Insert(ByVal tableName As String, _
                                 ByVal parameters As BosParameter()) As Integer

        Dim strInsertUpdate1, strInsertUpdate2 As New StringBuilder
        Dim parametersInsertUpdate As New BosParamsList
        Dim strInsertUpdate As String
        Dim nValue As Integer

        ' 新增資料, 組SQL
        'strInsertUpdate1.Clear()
        'strInsertUpdate2.Clear()
        'parametersInsertUpdate.Clear()
        'strInsertUpdate = ""

        For Each param In parameters

            If Left(param.parameterName, 1) = "!" Then
                param.parameterName = Mid(param.parameterName, 2)
            End If

            '檢查欄位是否正確
            If String.Compare(m_sTableName, tableName, True) = 0 Then
#If DEBUG Then
                Dbg.Assert(m_attributes.ContainsKey(param.parameterName) = True, param.parameterName & "不是欄位")
#End If
                If m_attributes.ContainsKey(param.parameterName) = False Then
                    Throw New Exception(param.parameterName & "不是欄位")
                End If
            End If
            If m_arrayPrimaryKeys.Contains(param.parameterName) AndAlso IsNothing(param.parameterValue) Then
#If DEBUG Then
                Dbg.Assert(IsNothing(param.parameterValue) = False, "PK欄位的值不可為NULL")
#End If
                Throw New Exception("PK欄位的值不可為NULL: [" & param.parameterName & "]")
            End If

            If strInsertUpdate1.Length <> 0 Then
                strInsertUpdate1.Append(" , ")
                strInsertUpdate2.Append(" , ")
            End If

            strInsertUpdate1.Append(" " & param.parameterName & " ")
            strInsertUpdate2.Append(" @" & param.parameterName & "@ ")
            parametersInsertUpdate.Add(param)
        Next

#If DEBUG Then
        Dbg.Assert(strInsertUpdate1.Length > 0 OrElse strInsertUpdate2.Length > 0, "Insert語法錯誤")
#End If

        strInsertUpdate = _
            "insert into " & tableName & " ( " & strInsertUpdate1.ToString() & _
            " ) values ( " & strInsertUpdate2.ToString() & " )"

        nValue = ExecuteNonQuery(strInsertUpdate, parametersInsertUpdate.ToArray())

        'Dbg.Assert(nValue > 0)    '應該至少會有一行被更新或新增
        Return nValue

    End Function

    Public Function Update(ByVal ParamArray objs() As Object) As Integer
        Return Update(PARAM_ARRAY(objs))
    End Function

    ''' <summary>
    ''' 更新記錄。
    ''' </summary>
    ''' <param name="parameters"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Update(ByVal parameters As BosParameter()) As Integer
#If DEBUG Then
        Dbg.Assert(IsNothing(getTableName()) = False AndAlso getTableName().Length > 0, "Table不存在")
#End If
        Return Update(getTableName(), parameters)
    End Function


    Public Function Update(ByVal datarow As DataRow) As Integer

        Dim arrayBosParameters As New BosParamsList

        For Each column As DataColumn In datarow.Table.Columns
            If Not IsNothing(datarow(column.ColumnName)) Then
                arrayBosParameters.Add(column.ColumnName, datarow(column.ColumnName))
            End If
        Next

        Return Update(arrayBosParameters.ToArray())

    End Function


    ''' <summary>
    ''' 更新記錄。
    ''' </summary>
    ''' <param name="tableName"></param>
    ''' <param name="parameters"></param>
    ''' <returns>受影響的資料列數目。</returns>
    ''' <remarks></remarks>
    ''' <SAMPLE>
    ''' </SAMPLE>
    Public Function Update(ByVal tableName As String, _
                                 ByVal parameters As BosParameter()) As Integer
        Dim listPk, listNotPk As New BosParamsList
        ''Dim strInsert, strUpdate As String

        '區分PK及非PK
        For Each param As BosParameter In parameters

            Dim sFieldName As String = param.parameterName
            If Left(sFieldName, 1) = "!" Then
                sFieldName = Mid(sFieldName, 2)
            End If

#If DEBUG Then
            '檢查欄位是否正確
            If String.Compare(m_sTableName, tableName, True) = 0 Then

                Dbg.Assert(m_attributes.ContainsKey(sFieldName) = True, sFieldName & "不是欄位")

                If m_attributes.ContainsKey(sFieldName) = False Then
                    Throw New Exception(sFieldName & "不是欄位")
                End If
            End If
#End If

            If m_arrayPrimaryKeys.Contains(sFieldName) Then
                listPk.Add(param)
            Else
                listNotPk.Add(param)
            End If
        Next

        '#If DEBUG Then
        '        If String.Compare(m_sTableName, tableName, True) = 0 Then
        '            '檢查是否有PK，或PK數目是否正確
        '            'Dbg.Assert(listPk.Count = m_arrayPrimaryKeys.Count AndAlso listPk.Count > 0, "PK數目不正確")
        '        End If
        '#End If

        If listPk.Count = 0 Then
            Throw New BosException("PK的數量不能為0")
        End If


        Dim strInsertUpdate1, strInsertUpdate2 As New StringBuilder
        Dim parametersInsertUpdate As New BosParamsList
        Dim strInsertUpdate As String
        Dim nValue As Integer

        ' 更新資料, 組SQL
        For Each param In listNotPk
            If strInsertUpdate1.Length <> 0 Then
                strInsertUpdate1.Append(" , ")
            End If
            strInsertUpdate1.Append(" " & param.parameterName & " = @" & param.parameterName & "@ ")
            parametersInsertUpdate.Add(param)
        Next

        For Each param In listPk
#If DEBUG Then
            Dbg.Assert(IsNothing(param.parameterValue) = False, "PK欄位的值不可為NULL")
#End If
            If IsNothing(param.parameterValue) Then
                Throw New Exception("PK欄位的值不可為NULL: [" & param.parameterName & "]")
            End If

            If strInsertUpdate2.Length > 0 Then
                strInsertUpdate2.Append(" and ")
            End If
            'strInsertUpdate2.Append(" " & param.parameterName & " = @" & param.parameterName & "@ ")

            If Left(param.parameterName, 1) = "!" Then
                param.parameterName = Mid(param.parameterName, 2)

                If IsNothing(param.parameterValue) = True OrElse IsDBNull(param.parameterValue) = True Then
                    strInsertUpdate2.Append(" " & param.parameterName & " is not null" & vbCrLf)
                Else
                    strInsertUpdate2.Append(" " & param.parameterName & " <> @" & param.parameterName & "@" & vbCrLf)
                End If
            Else
                If IsNothing(param.parameterValue) = True OrElse IsDBNull(param.parameterValue) = True Then
                    strInsertUpdate2.Append(" " & param.parameterName & " is null" & vbCrLf)
                Else
                    strInsertUpdate2.Append(" " & param.parameterName & " = @" & param.parameterName & "@" & vbCrLf)
                End If
            End If


            parametersInsertUpdate.Add(param)
        Next

#If DEBUG Then
        Dbg.Assert(strInsertUpdate1.Length > 0 OrElse strInsertUpdate2.Length > 0, "Update 的語法不正確")
#End If
        nValue = 0

        If strInsertUpdate1.Length > 0 AndAlso strInsertUpdate2.Length > 0 Then
            strInsertUpdate = _
                "update " & tableName & " set " & strInsertUpdate1.ToString() & _
                " where " & strInsertUpdate2.ToString()

            nValue = ExecuteNonQuery(strInsertUpdate, parametersInsertUpdate.ToArray())
        End If

        Return nValue
    End Function


    ''' <summary>
    ''' 取得記錄筆數。
    ''' </summary>
    ''' <param name="parameters"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>

    Public Function GetCount(ByVal parameters As BosParameter()) As Integer

#If DEBUG Then
        Dbg.Assert(IsNothing(getTableName()) = False AndAlso getTableName().Length > 0, "Table不存在")
#End If

        Return GetCount(getTableName(), parameters)
    End Function

    <Obsolete("This method is deprecated, use GetCount(PARAM_ARRAY(param1, param2, ...) instead.")> _
    Public Function GetCount(ByVal ParamArray objs() As Object) As Integer
        Return GetCount(getTableName(), PARAM_ARRAY(objs))
    End Function


    Public Function Delete(ByVal ParamArray objs() As Object) As Integer
        Return Delete(PARAM_ARRAY(objs))
    End Function


    ''' <summary>
    ''' 刪除記錄
    ''' </summary>
    ''' <param name="parameters"></param>
    ''' <returns>受影響的資料列數目</returns>
    ''' <remarks></remarks>
    Public Function Delete(ByVal parameters As BosParameter()) As Integer
#If DEBUG Then
        Dbg.Assert(IsNothing(getTableName()) = False AndAlso getTableName().Length > 0, "Table不存在")
#End If
        Return Delete(getTableName(), parameters)
    End Function



    ''' <summary>
    ''' 刪除記錄
    ''' </summary>
    ''' <param name="tableName"></param>
    ''' <param name="parameters"></param>
    ''' <returns>受影響的資料列數目</returns>
    ''' <remarks></remarks>
    Public Function Delete(ByVal tableName As String, _
                                 ByVal parameters As BosParameter()) As Integer

        Dim sbDelete As New StringBuilder
        Dim parametersDelete As New BosParamsList
        Dim strDelete As String
        Dim nValue As Integer

        ' 新增資料, 組SQL

        For Each param In parameters
            'If Not IsNothing(param.parameterValue) Then
            If sbDelete.Length <> 0 Then
                sbDelete.Append(" and ")
            End If

#If DEBUG Then
            If String.Compare(m_sTableName, tableName, True) = 0 Then
                '檢查是否應為PK
                'Dbg.Assert(m_arrayPrimaryKeys.Contains(param.parameterName.Trim.ToUpper), param.parameterName.Trim & "應為PK", False)

                Dim sFieldName As String = param.parameterName
                If Left(sFieldName, 1) = "!" Then
                    sFieldName = Mid(sFieldName, 2)
                End If

                '檢查欄位是否正確
                Dbg.Assert(m_attributes.ContainsKey(sFieldName) = True, sFieldName & "不是欄位")
            End If
#End If

            'sbDelete.Append(" " & param.parameterName & " = @" & param.parameterName & "@ ")

            If Left(param.parameterName, 1) = "!" Then
                param.parameterName = Mid(param.parameterName, 2)

                If IsNothing(param.parameterValue) = True OrElse IsDBNull(param.parameterValue) = True Then
                    sbDelete.Append(" " & param.parameterName & " is not null" & vbCrLf)
                Else
                    sbDelete.Append(" " & param.parameterName & " <> @" & param.parameterName & "@" & vbCrLf)
                End If
            Else
                If IsNothing(param.parameterValue) = True OrElse IsDBNull(param.parameterValue) = True Then
                    sbDelete.Append(" " & param.parameterName & " is null" & vbCrLf)
                Else
                    sbDelete.Append(" " & param.parameterName & " = @" & param.parameterName & "@" & vbCrLf)
                End If
            End If


            parametersDelete.Add(param)
            'End If
        Next

#If DEBUG Then
        Dbg.Assert(sbDelete.Length > 0, "Delete語法不正確")
#End If
        If sbDelete.Length = 0 Then
            Throw New Exception("刪除必須要有條件")
        End If

        strDelete = _
            "delete from " & tableName & " where " & sbDelete.ToString()

        nValue = ExecuteNonQuery(strDelete, parametersDelete.ToArray())

        Return nValue

    End Function


    ''' <summary>
    ''' 取得記錄筆數。
    ''' </summary>
    ''' <param name="tableName"></param>
    ''' <param name="parameters"></param>
    ''' <returns>資料列數目。</returns>
    ''' <remarks></remarks>
    ''' <SAMPLE>
    ''' </SAMPLE>
    Public Function GetCount(ByVal tableName As String, _
                                 ByVal parameters As BosParameter()) As Integer
        Dim listFilter As New BosParamsList
        ''Dim strInsert, strUpdate As String

        For Each param In parameters
            listFilter.Add(param)
        Next

        Dim sbFilter As New StringBuilder
        Dim parametersFilter As New BosParamsList
        Dim nValue As Integer

        ' 更新資料, 組SQL

        sbFilter.Append("select  count(*) from " & tableName)

        If listFilter.Count > 0 Then
            'sbFilter.Append(" WHERE ")

            Dim sbFilterWhere As New StringBuilder

            For Each param In listFilter
                'If Not IsNothing(param.parameterValue) Then
                If sbFilterWhere.Length <> 0 Then
                    sbFilterWhere.Append(" and ")
                End If

#If DEBUG Then
                '檢查欄位是否正確

                If String.Compare(m_sTableName, tableName, True) = 0 Then
                    Dim sFieldName As String = param.parameterName
                    If Left(sFieldName, 1) = "!" Then
                        sFieldName = Mid(sFieldName, 2)
                    End If

                    Dbg.Assert(m_attributes.ContainsKey(sFieldName) = True, sFieldName & "不是欄位")
                End If
#End If

                If Left(param.parameterName, 1) = "!" Then
                    param.parameterName = Mid(param.parameterName, 2)

                    If IsNothing(param.parameterValue) = True OrElse IsDBNull(param.parameterValue) = True Then
                        sbFilterWhere.Append(" " & param.parameterName & " is not null" & vbCrLf)
                    Else
                        sbFilterWhere.Append(" " & param.parameterName & " <> @" & param.parameterName & "@" & vbCrLf)
                    End If
                Else
                    If IsNothing(param.parameterValue) = True OrElse IsDBNull(param.parameterValue) = True Then
                        sbFilterWhere.Append(" " & param.parameterName & " is null" & vbCrLf)
                    Else
                        sbFilterWhere.Append(" " & param.parameterName & " = @" & param.parameterName & "@" & vbCrLf)
                    End If
                End If

                parametersFilter.Add(param)
            Next

            If (sbFilterWhere.Length > 0) Then
                sbFilter.Append(" where ")
                sbFilter.Append(sbFilterWhere)
            End If
        End If

        nValue = CInt(ExecuteScalar(sbFilter.ToString(), parametersFilter.ToArray()))
        Return nValue
    End Function


    ''' <summary>
    ''' 取得DataRowCollection。若不為NOTHING，則一定至少有一筆資料
    ''' </summary>
    ''' <param name="strSQL">若為NOTHING，程式會自行組SQL字串</param>
    ''' <param name="objs"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Obsolete("This method is deprecated, use GetDataRowCollection2 instead.")> _
    Public Function GetDataRowCollection(ByVal strSQL As String, _
                                         ByVal ParamArray objs() As Object) As DataRowCollection
        Return GetDataRowCollection(strSQL, PARAM_ARRAY(objs))
    End Function

    ''' <summary>
    ''' 取得DataRowCollection。若不為NOTHING，則一定至少有一筆資料
    ''' </summary>
    ''' <param name="strSQL">若為NOTHING，程式會自行組SQL字串</param>
    ''' <param name="parameters"></param>
    ''' <param name="commandType"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Obsolete("This method is deprecated, use GetDataRowCollection2 instead.")> _
    Public Function GetDataRowCollection(ByVal strSQL As String, _
                                 ByVal parameters As BosParameter(), _
                                 Optional ByVal commandType As CommandType = CommandType.Text) As DataRowCollection

        Dim dt As DataTable
        dt = GetDataTable(strSQL, parameters, commandType)

        Dbg.Assert(Not IsNothing(dt), "查詢沒有資料")
        If IsNothing(dt) Then
            Return Nothing
        End If

        If IsNothing(dt.Rows) Then
            Return Nothing
        End If

        If dt.Rows.Count = 0 Then
            Return Nothing
        End If

        Return dt.Rows

    End Function


    ''' <summary>
    ''' 取得DataRowCollection。
    ''' </summary>
    ''' <param name="strSQL">若為NOTHING，程式會自行組SQL字串</param>
    ''' <param name="objs"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetDataRowCollection2(ByVal strSQL As String, _
                                         ByVal ParamArray objs() As Object) As DataRowCollection
        Return GetDataRowCollection2(strSQL, PARAM_ARRAY(objs))
    End Function

    ''' <summary>
    ''' 取得DataRowCollection。
    ''' </summary>
    ''' <param name="strSQL">若為NOTHING，程式會自行組SQL字串</param>
    ''' <param name="parameters"></param>
    ''' <param name="commandType"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetDataRowCollection2(ByVal strSQL As String, _
                                 ByVal parameters As BosParameter(), _
                                 Optional ByVal commandType As CommandType = CommandType.Text) As DataRowCollection

        Dim dt As DataTable
        dt = GetDataTable(strSQL, parameters, commandType)

        Dbg.Assert(Not IsNothing(dt), "查詢沒有資料")
        If IsNothing(dt) Then
            Return Nothing
        End If

        Return dt.Rows

    End Function


    ''' <summary>
    ''' 取得DataRow。若不為NOTHING，則一定至少有一筆資料
    ''' </summary>
    ''' <param name="strSQL">若為NOTHING，程式會自行組SQL字串</param>
    ''' <param name="objs"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetDataRow(ByVal strSQL As String, _
                               ByVal ParamArray objs() As Object) As DataRow
        Return GetDataRow(strSQL, PARAM_ARRAY(objs))
    End Function


    ''' <summary>
    ''' 取得DataRow。若不為NOTHING，則一定只有一筆資料
    ''' </summary>
    ''' <param name="strSQL">若為NOTHING，程式會自行組SQL字串</param>
    ''' <param name="parameters"></param>
    ''' <param name="commandType"></param>
    ''' <returns>若不為NOTHING，則傳回一筆資料</returns>
    ''' <remarks></remarks>
    Public Function GetDataRow(ByVal strSQL As String, _
                                 ByVal parameters As BosParameter(), _
                                 Optional ByVal commandType As CommandType = CommandType.Text) As DataRow

        Dim drc As DataRowCollection
        drc = GetDataRowCollection2(strSQL, parameters, commandType)

        'Dbg.Assert(Not IsNothing(drc), "查詢沒有資料")
        If drc.Count = 0 Then
            Return Nothing
        End If

#If DEBUG Then
        '若有多筆資料則警告，(可能QUERY語法不正確導致多筆資料)
        Debug.Assert(drc.Count = 1, "資料筆數應等於1")
#End If
        Return (drc(0))

    End Function



    Public Shared lastID As Integer = 0
    Protected Sub TRACE(ByVal strSQL As String, ByVal parameters As BosParameter())
        Try

            m_TraceLastSQL = SQLTrace.Create(strSQL, parameters)
            GlobalTraceLastSQL = m_TraceLastSQL

#If DEBUG Then
            If SQL_DUMP AndAlso System.Diagnostics.Debugger.IsAttached = True Then
                lastID = lastID + 1

                Dim strTrace As String
                strTrace = GetSQLStatement(strSQL, parameters)

                System.Diagnostics.Trace.WriteLine("--SQL " & lastID.ToString() & ":" & vbCrLf & strTrace & ";" & vbCrLf)
            End If
#End If

            If Environment.MachineName.ToUpper = "WIN-QRLAKF0WDIF" Then
                lastID = lastID + 1

                Dim strTrace2 As String
                strTrace2 = GetSQLStatement(strSQL, parameters)

                Dim tempPath As String = System.IO.Path.GetTempPath() & "\SQL.LOG"

                WriteEventLog("--SQL " & lastID.ToString() & ":" & vbCrLf & strTrace2 & ";" & vbCrLf,
                              tempPath)
            End If

        Catch ex As Exception

        End Try

    End Sub

    Public Shared Sub WriteEventLog(ByVal sEvent As String, ByVal sFile As String)
        Dim fs As FileStream = Nothing
        Dim sw As StreamWriter = Nothing
        Dim sr As StreamReader = Nothing
        Dim listString As New ArrayList
        Dim fiFile As FileInfo

        Try
            fiFile = New FileInfo(sFile)

            If fiFile.Length > 10000000 Then
                sr = New StreamReader(sFile, System.Text.Encoding.UTF8)

                Do While sr.Peek() >= 0
                    listString.Add(sr.ReadLine())
                Loop
                sr.Close()
                File.Delete(sFile)
                fs = New FileStream(sFile, FileMode.Create)
                sw = New StreamWriter(fs, System.Text.Encoding.UTF8)

                If listString.Count > 0 Then
                    For n As Integer = listString.Count >> 1 To listString.Count - 1
                        sw.WriteLine(listString(n))
                    Next
                End If
            End If
        Catch ex As Exception
        End Try


        Try
            If IsNothing(fs) Then
                If File.Exists(sFile) Then
                    fs = New FileStream(sFile, FileMode.Append)
                Else
                    fs = New FileStream(sFile, FileMode.Create)
                End If
            End If

            If IsNothing(sw) Then
                sw = New StreamWriter(fs, System.Text.Encoding.UTF8)
            End If

            sw.WriteLine(sEvent)
        Catch ex As Exception

        Finally
            Try
                sr.Close()
            Catch ex As Exception
            End Try

            Try
                sw.Close()
            Catch ex As Exception
            End Try

            Try
                fs.Close()
            Catch ex As Exception
            End Try
        End Try
    End Sub

    ''' <summary>
    ''' 取得SQL語法
    ''' </summary>
    ''' <param name="strSQL"></param>
    ''' <param name="parameters"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetSQLStatement(ByVal strSQL As String, ByVal parameters As BosParameter()) As String
        Dim sbTrace As New StringBuilder(strSQL)

        If IsNothing(parameters) = False Then
            For n As Integer = 0 To parameters.Count - 1
                If TypeOf parameters(n).parameterValue Is String Then
                    sbTrace.Replace("@" & parameters(n).parameterName & "@", "'" & parameters(n).parameterValue.ToString & "'")
                    Continue For
                End If

                If IsNothing(parameters(n).parameterValue) Then
                    sbTrace.Replace("@" & parameters(n).parameterName & "@", "NULL")
                    Continue For
                End If

                If TypeOf parameters(n).parameterValue Is DateTime Then
                    sbTrace.Replace("@" & parameters(n).parameterName & "@", "'" & _
                                    CDate(parameters(n).parameterValue).ToString("yyyy-MM-dd hh:mm:ss.fff") & "'")
                    Continue For
                End If

                sbTrace.Replace("@" & parameters(n).parameterName & "@", parameters(n).parameterValue.ToString)
            Next
        End If

        Return sbTrace.ToString
    End Function

    ''' <summary>
    ''' 取得最後執行的SQL語法
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property GetLastSQL() As String
        Get
            Return GetSQLStatement(m_TraceLastSQL.SQL, m_TraceLastSQL.Parameter)
        End Get
    End Property

    Public Shared ReadOnly Property GetGlobalLastSQL() As String
        Get
            Return GetSQLStatement(GlobalTraceLastSQL.SQL, GlobalTraceLastSQL.Parameter)
        End Get
    End Property


    ''' <summary>
    ''' 將從DB讀出來的資料轉成指定的資料類別，如果是空的，則轉成NOTHING
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="Expression"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overloads Shared Function CDbType(Of T)(ByVal Expression As Object) As T

        If IsNothing(Expression) OrElse IsDBNull(Expression) Then
            Return Nothing
        End If

        Return CType(Expression, T)
    End Function

    ''' <summary>
    ''' 將從DB讀出來的資料轉成指定的資料類別，如果是空的，則轉成預設的值(tDefault)
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="Expression"></param>
    ''' <param name="tDefault"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overloads Shared Function CDbType(Of T)(ByVal Expression As Object, ByVal tDefault As T) As T

        If IsNothing(Expression) OrElse IsDBNull(Expression) Then
            Return tDefault
        End If

        Return CType(Expression, T)
    End Function



    ''' <summary>
    ''' 從DB讀出資料並轉成字串，如果是DBNULL，則轉成NOTHING
    ''' </summary>
    ''' <param name="strSQL_Field"></param>
    ''' <param name="objs"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function QueryString(ByVal strSQL_Field As String, _
                                ByVal ParamArray objs() As Object) As String
        Return QueryValue(Of String)(strSQL_Field, Nothing, objs)
    End Function


    ''' <summary>
    ''' 從DB讀出一筆資料並轉成指定的資料類別，如果是空的，則轉成預設的值(DBNULLValue)
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="strSQL_Field"></param>
    ''' <param name="DBNULLValue"></param>
    ''' <param name="objs"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function QueryValue(Of T)(ByVal strSQL_Field As String, _
                                     ByVal DBNULLValue As T, _
                                     ByVal ParamArray objs() As Object) As T

        Return CDbType(Of T)(ExecuteScalar(strSQL_Field, objs), DBNULLValue)
    End Function

    ''' <summary>
    ''' 從DB讀出多筆資料並轉成指定的資料類別，如果是空的，則轉成預設的值(DBNULLValue)
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="strSQL_Field"></param>
    ''' <param name="DBNULLValue"></param>
    ''' <param name="objs"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function QueryList(Of T)(ByVal strSQL_Field As String, _
                                          ByVal DBNULLValue As Object, _
                                          ByVal ParamArray objs() As Object) As List(Of T)
        Dim drc As DataRowCollection
        Dim objList As New List(Of T)

        drc = GetDataRowCollection2(strSQL_Field, objs)

        For Each dr As DataRow In drc
            objList.Add(CDbType(Of T)(dr(0), CType(DBNULLValue, T)))
        Next

        Return objList
    End Function


    ''' <summary>
    ''' 從DB讀出多筆資料並轉成指定的資料類別，如果是空的，則轉成預設的值(DBNULLValue)
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="strSQL_Field"></param>
    ''' <param name="DBNULLValue"></param>
    ''' <param name="objs"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function QueryArray(Of T)(ByVal strSQL_Field As String, _
                                          ByVal DBNULLValue As Object, _
                                          ByVal ParamArray objs() As Object) As T()

        Return QueryList(Of T)(strSQL_Field, DBNULLValue, objs).ToArray

    End Function

End Class
