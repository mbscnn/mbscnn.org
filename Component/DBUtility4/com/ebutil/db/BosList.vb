Option Explicit On
Option Strict On
Imports System.Runtime.Serialization
Imports System.Security.Permissions
Imports System.ComponentModel
Imports System.Data
Imports System.Collections

''' <summary>
'''  The List of the Business objects (BOS)。
''' </summary>
''' <remarks>
''' BosList 實作IEnumerator、IEnumerable 介面
''' 所有實作多筆 Table Object物件須繼承BosList。
''' 
''' </remarks>
''' <example> This sample shows how to inherits the BosList class based on EN_CASEMAIN which has the primary key
''' <code>
''' '請注意命名方法 
''' '----------------------------------------------
''' '|Table Name  | BosBase     |  BosList         |
''' '|------------| ------------|  ----------------|
''' '|EN_CASEMAIN | EN_CASEMAIN |  EN_CASEMAINList |
''' '----------------------------------------------
''' Public Class EN_CASEMAINList Inherits BosList
'''     '建構子
'''     Sub New(ByVal dbManager As DatabaseManager)
'''          MyBase.New("EN_CASEMAIN", dbManager)
'''     End Sub
''' 
'''     'must Overrides Function newBos
'''     Overrides Function newBos() As BosBase
'''          Return New EN_CASEMAIN(MyBase.getDatabaseManager)
'''     End Function
'''     
'''     '案例一:不寫SQL語法，直接給參數
'''     '內部運作語法:[select * from en_casemain where GOBRID=sBrId]
'''     Function loadByBrIdEx1(ByVal sBrId As String) As Integer
'''          Try
'''               If (Not IsNothing(sBrId) AndAlso sBrId.Length > 0) Then
'''                    Dim paras(0) As System.Data.IDbDataParameter
'''                    paras(0) = ProviderFactory.CreateDataParameter("GOBRID", sBrId)
'''                    Return MyBase.loadBySQL(paras)
'''               End If
'''          Catch ex As ProviderException
'''               Throw ex
'''          Catch ex As BosException
'''               Throw ex
'''          Catch ex As Exception
'''               Throw ex
'''          End Try
''' 
'''          Return 0
'''     End Function
''' 
'''     '案例二:寫SQL語法，並給參數
'''     '內部運作語法:[programmer 給的SQL = select * from en_casemain where GOBRID=sBrId]
'''     Function loadByBrIdEx2(ByVal sBrId As String) As Integer
'''          Try
'''               If (Not IsNothing(sBrId) AndAlso sBrId.Length > 0) Then
'''                    Dim sSQL As String="select * from en_casemain where GOBRID" + ProviderFactory.PositionPara + "sBrId"
'''                    Dim paras(0) As System.Data.IDbDataParameter
'''                    paras(0) = ProviderFactory.CreateDataParameter("GOBRID", sBrId)
'''                    Return MyBase.loadBySQL(sSQL, paras)
'''               End If
'''          Catch ex As ProviderException
'''               Throw ex
'''          Catch ex As BosException
'''               Throw ex
'''          Catch ex As Exception
'''               Throw ex
'''          End Try
''' 
'''          Return 0
'''     End Function
''' 
'''    '案例三:寫Join SQL語法，並給參數
'''    '內部運作語法:[programmer 給的SQL = select  a.caseid,a.brid,b.caseitemid caseitemid,b.loanitem 授信項目名稱,b.datadate 資料日期 from En_Casemain a ,En_Casedetailcond  b where a.caseid=b.caseid and b.caseid=sCaseID]
'''    '如果是多各不同Table Join在一起時，
'''    '在此物件(EN_CASEMAINList)須手動增加欄位屬性，才可取得不同Table(En_Casedetailcond)的欄位(caseitemid、loanitem、datadate)
'''     Function loadCaseItem(ByVal sCaseId As String) As Integer
'''          Try
'''               If (Not IsNothing(sCaseId) AndAlso sCaseId.Length > 0) Then
'''                    Dim sqlStr As String = ""
'''                    sqlStr = "select  a.caseid,a.brid,b.caseitemid caseitemid,b.loanitem 授信項目名稱,b.datadate 資料日期 from En_Casemain a ,En_Casedetailcond  b where a.caseid=b.caseid " + _
'''                    " and CASEID=" + ProviderFactory.PositionPara + "CASEID "
''' 
'''                    '下面這三個欄位屬於En_Casedetailcond的，並非屬於En_CaseMain的欄位
'''                    '因此須將En_Casedetailcond的三個欄位手動加入EN_CASEMAINList的物件，如此此能取得這三各欄位的值
'''                    Me.addAttribute("caseitemid", System.Data.DbType.Decimal)'如果資料型別是Number-->System.Data.DbType.Decimal
'''                    Me.addAttribute("授信項目名稱", System.Data.DbType.String)'如果資料型別是String-->System.Data.DbType.String
'''                    Me.addAttribute("資料日期", System.Data.DbType.Date)'如果資料型別是Date-->System.Data.DbType.Date
''' 
'''                    Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("CASEID", sCaseId)
''' 
'''                    Return Me.loadBySQL(sqlStr, para)
'''               End If
'''          Catch ex As ProviderException
'''               Throw ex
'''          Catch ex As BosException
'''               Throw ex
'''          Catch ex As Exception
'''               Throw ex
'''          End Try
''' 
'''          Return 0
'''     End Function
''' End Class
''' </code>
''' </example>
''' <history>
''' 	[Titan]	2009/01/20	modify
''' </history>
Public MustInherit Class BosList ''<Serializable()>
    'Inherits SerializableObject
    Implements System.Collections.IEnumerable, System.Collections.IEnumerator ', System.ICloneable ', ISerializable, IDeserializationCallback
    'ENOP 組件名稱
    '    Private m_sDLLName As String '= Properties.m_sDLLName
    'ENOP 命名空間
    'Private m_sNameSpase As String  '= Properties.m_sNameSpase

    ''' <summary>
    ''' 連線物件
    ''' </summary>
    ''' <remarks></remarks>
    Private m_databaseManager As DatabaseManager

    ''' <summary>
    ''' Table Name
    ''' </summary>
    ''' <remarks></remarks>
    Private m_sTableName As String

    ''' <summary>
    ''' 存放BosBase物件陣列
    ''' </summary>
    ''' <remarks></remarks>
    Private m_objects As New ArrayList
    ''' <summary>
    ''' 存放BosAttrMeta欄位物件陣列
    ''' </summary>
    ''' <remarks></remarks>
    Private m_BosAttrMetas As New ArrayList

    ''' <summary>
    ''' 依需求調整Condition
    ''' </summary>
    ''' <remarks></remarks>
    Private m_sSQLCondition As String

    ''' <summary>
    ''' 目前記錄的指標
    ''' </summary>
    ''' <remarks></remarks>
    Private m_currentIndex As Integer = -1

    ''' <summary>
    ''' Current DataSet
    ''' </summary>
    ''' <remarks></remarks>
    Private m_dsCurrent As DataSet

    ''' <summary>
    ''' SQL script
    ''' </summary>
    ''' <remarks></remarks>
    Private m_sSQL As String

    Public sErrorMsg As String = ""

    'Public m_keepDs As DataSet = Nothing

#Region "建構子"

    ''' <summary>
    ''' 初始化 BosList 類別的新執行個體 (Instance)。
    ''' </summary>
    ''' <param name="sTableName">傳入Table Name(大寫)</param>
    ''' <param name="databaseManager">傳入 DatabaseManager</param>
    ''' <remarks>建構子
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/01/20	modify
    ''' </history>
    Sub New(ByVal sTableName As String, ByVal databaseManager As DatabaseManager)
        'ENOP 命名空間
        'm_sNameSpase = Me.GetType.Namespace
        ''ENOP 組件名稱
        'm_sDLLName = Me.GetType.Module.Name.ToUpper.Replace(".DLL", "")
        setDatabaseManager(databaseManager)
        setTableName(sTableName)
    End Sub

    'Private Sub New(ByVal sTableName As String, ByVal databaseManager As DatabaseManager, ByVal ds As DataSet)
    '    Me.m_databaseManager = databaseManager
    '    setConnection(databaseManager.getConnection)
    '    If (databaseManager.isTransaction) Then
    '        setTransaction(databaseManager.getTransaction())
    '    End If
    '    setTableName(sTableName)
    'End Sub

    'Private Sub New(ByVal sTableName As String, ByVal connection As OleDbConnection)
    '    setConnection(connection)
    '    setTableName(sTableName)
    'End Sub

    'Private Sub New(ByVal sTableName As String, ByVal transaction As OleDbTransaction)
    '    setTransaction(transaction)
    '    setConnection(transaction.Connection)
    '    setTableName(sTableName)
    'End Sub
#End Region

#Region "load SQL"

    ''' <summary>
    ''' 取得相關Table的所有記錄。
    ''' </summary>
    ''' <returns>Integer 傳回記錄的筆數。</returns>
    ''' <remarks>
    ''' select * from [your TABLE_NAME]
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/01/20	modify
    ''' </history>
    Public Function loadAllData(Optional ByVal isDataSet As Boolean = False) As Integer
        Dim sSQL As String = " SELECT * FROM " + Me.getSchemaTableName
        If isDataSet Then
            Return loadBySQLOnlyDs(sSQL)
        End If
        Return loadBySQL(sSQL)
    End Function

#Region "loadBySQL"
    ''' <summary>
    ''' 使用DataAdatpter.Fill(DataTable)方法，取的相關Table的記錄。
    ''' </summary>
    ''' <param name="sSQL">傳入SQL語法。</param>
    ''' <returns>Integer 傳回記錄的筆數。</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/01/20	modify
    ''' </history>
    Protected Function loadBySQL(ByVal sSQL As String) As Integer
        Return loadBySQL(sSQL, Nothing)
    End Function

    ''' <summary>
    ''' 使用DataAdatpter.Fill(DataTable)方法，取的相關Table的記錄。
    ''' </summary>
    ''' <param name="commandParameters">傳入IDataParameter陣列。</param>
    ''' <returns>Integer 傳回記錄的筆數。</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/01/20	modify
    ''' </history>
    Protected Function loadBySQL(ByVal ParamArray commandParameters() As IDataParameter) As Integer
        Dim sbSQL As New System.Text.StringBuilder

        sbSQL.Append(" SELECT * FROM " + Me.getSchemaTableName())
        If (commandParameters.GetLength(0) > 0) Then
            sbSQL.Append(" WHERE ")
        End If
        For Each para As IDataParameter In commandParameters
            'for poc
            If Me.getDatabaseManager.getDataSource.IndexOf("IBM") = -1 Then
                sbSQL.Append(para.ParameterName & "=" & ProviderFactory.PositionPara & para.ParameterName & " AND ")
            ElseIf Me.getDatabaseManager.getDataSource.IndexOf("IBM") <> -1 Then
                sbSQL.Append(para.ParameterName & "=? AND ")
            Else
                sbSQL.Append(para.ParameterName & "=" & ProviderFactory.PositionPara & para.ParameterName & " AND ")
            End If
        Next
        sbSQL.Remove(sbSQL.Length - 4, 4)

        Return loadBySQL(sbSQL.ToString, commandParameters)
    End Function

    ''' <summary>
    ''' 使用DataAdatpter.Fill(DataTable)方法，取的相關Table的記錄。
    ''' </summary>
    ''' <param name="sSQL">傳入SQL語法。</param>
    ''' <param name="commandParameters">傳入IDataParameter陣列。</param>
    ''' <returns>Integer 傳回記錄的筆數</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/01/20	modify
    ''' </history>
    Protected Function loadBySQL(ByVal sSQL As String, ByVal ParamArray commandParameters() As IDataParameter) As Integer
        Try
            If Not IsNothing(commandParameters) Then
                Dim pars(commandParameters.Length - 1) As IDataParameter
                commandParameters.CopyTo(pars, 0)
            End If

            m_sSQL = sSQL
            m_sSQL &= " " & getSQLCondition()
            If ProviderFactory.m_bDebug Then
                Titan.Utility.Debug(ShowSelectSQL(m_sSQL, commandParameters))
            End If

#If DEBUG Then
            Trace.WriteLine(ShowSelectSQL(m_sSQL, commandParameters))
#End If


            Dim ds As DataSet = Me.ExecuteDataset(CommandType.Text, m_sSQL, commandParameters)
            Me.setSQLCondition("")

            If Not IsNothing(ds) AndAlso ds.Tables.Count > 0 Then
                constructData(ds.Tables(0))
                Return size()
            End If
            Return 0
        Catch ex As Exception
            Throw New BosException(Me.getSchemaTableName & ":(Exception)", ex, CType(ex.TargetSite, Reflection.MethodInfo))
        End Try
    End Function
#End Region

#Region "loadBySQLOnlyDs"
    ''' <summary>
    ''' 使用DataAdatpter.Fill(DataTable)方法，取的相關Table的記錄。
    ''' </summary>
    ''' <param name="sSQL">傳入SQL語法。</param>
    ''' <returns>Integer 傳回記錄的筆數。</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/01/20	modify
    ''' </history>
    Protected Function loadBySQLOnlyDs(ByVal sSQL As String) As Integer
        Return loadBySQLOnlyDs(sSQL, Nothing)
    End Function

    ''' <summary>
    ''' 使用DataAdatpter.Fill(DataTable)方法，取的相關Table的記錄。
    ''' </summary>
    ''' <param name="commandParameters">傳入IDataParameter陣列。</param>
    ''' <returns>Integer 傳回記錄的筆數。</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/01/20	modify
    ''' </history>
    Protected Function loadBySQLOnlyDs(ByVal ParamArray commandParameters() As IDataParameter) As Integer
        Dim sbSQL As New System.Text.StringBuilder

        sbSQL.Append(" SELECT * FROM " + Me.getSchemaTableName())
        If (commandParameters.GetLength(0) > 0) Then
            sbSQL.Append(" WHERE ")
        End If
        For Each para As IDataParameter In commandParameters
            'for poc
            If Me.getDatabaseManager.getDataSource.IndexOf("IBM") = -1 Then
                sbSQL.Append(para.ParameterName & "=" & ProviderFactory.PositionPara & para.ParameterName & " AND ")
            ElseIf Me.getDatabaseManager.getDataSource.IndexOf("IBM") <> -1 Then
                sbSQL.Append(para.ParameterName & "=? AND ")
            Else
                sbSQL.Append(para.ParameterName & "=" & ProviderFactory.PositionPara & para.ParameterName & " AND ")
            End If
        Next
        sbSQL.Remove(sbSQL.Length - 4, 4)

        Return loadBySQLOnlyDs(sbSQL.ToString, commandParameters)
    End Function

    ''' <summary>
    ''' 使用DataAdatpter.Fill(DataTable)方法，取的相關Table的記錄。
    ''' </summary>
    ''' <param name="sSQL">傳入SQL語法。</param>
    ''' <param name="commandParameters">傳入IDataParameter陣列。</param>
    ''' <returns>Integer 傳回記錄的筆數</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/01/20	modify
    ''' </history>
    Protected Function loadBySQLOnlyDs(ByVal sSQL As String, ByVal ParamArray commandParameters() As IDataParameter) As Integer
        Try
            If Not IsNothing(commandParameters) Then
                Dim pars(commandParameters.Length - 1) As IDataParameter
                commandParameters.CopyTo(pars, 0)
            End If

            m_sSQL = sSQL
            m_sSQL &= " " & getSQLCondition()
            If ProviderFactory.m_bDebug Then
                Titan.Utility.Debug(ShowSelectSQL(m_sSQL, commandParameters))
            End If

            Dim ds As DataSet = Me.ExecuteDataset(CommandType.Text, m_sSQL, commandParameters)
            Me.setSQLCondition("")

            If Not IsNothing(ds) Then
                If Not IsNothing(ds.Tables(0)) Then
                    If Not IsNothing(ds.Tables(0).Rows) Then
                        Return ds.Tables(0).Rows.Count
                    End If
                End If
            End If
            Return 0
        Catch ex As Exception
            Throw New BosException(Me.getSchemaTableName & ":(Exception)", ex, CType(ex.TargetSite, Reflection.MethodInfo))
        End Try
    End Function

#End Region

#Region "loadReaderBySQL"

    Protected Function loadReaderBySQL(ByVal sSQL As String, ByVal ParamArray commandParameters() As IDataParameter) As Integer
        Dim IReader As IDataReader = Nothing
        Try
            If Not IsNothing(commandParameters) Then
                Dim pars(commandParameters.Length - 1) As IDataParameter
                commandParameters.CopyTo(pars, 0)
            End If

            m_sSQL = sSQL
            m_sSQL &= " " & getSQLCondition()
            If ProviderFactory.m_bDebug Then
                Titan.Utility.Debug(ShowSelectSQL(m_sSQL, commandParameters))
            End If

            Using databaseManager As DatabaseManager = databaseManager.getInstance

                IReader = Me.ExecuteReader(databaseManager, CommandType.Text, m_sSQL, commandParameters)
                Me.setSQLCondition("")

                If Not IsNothing(IReader) Then
                    constructIRDData(IReader)
                    Return size()
                End If
            End Using
           
            Return 0
        Catch ex As Exception
            Throw New BosException(Me.getSchemaTableName & ":(Exception)", ex, CType(ex.TargetSite, Reflection.MethodInfo))
        Finally
            If Not IsNothing(IReader) Then IReader.Close()
        End Try
    End Function
#End Region

#Region "loadReaderBySQL"

    Protected Function loadBosBySQL(ByVal sSQL As String, ByVal ParamArray commandParameters() As IDataParameter) As Integer
        Dim IReader As IDataReader = Nothing
        Try
            If Not IsNothing(commandParameters) Then
                Dim pars(commandParameters.Length - 1) As IDataParameter
                commandParameters.CopyTo(pars, 0)
            End If

            m_sSQL = sSQL
            m_sSQL &= " " & getSQLCondition()
            If ProviderFactory.m_bDebug Then
                Titan.Utility.Debug(ShowSelectSQL(m_sSQL, commandParameters))
            End If

            Using databaseManager As DatabaseManager = databaseManager.getInstance

                IReader = Me.ExecuteReader(databaseManager, CommandType.Text, m_sSQL, commandParameters)
                Me.setSQLCondition("")

                If Not IsNothing(IReader) Then
                    constructBosData(IReader)
                    Return size()
                End If
            End Using

            Return 0
        Catch ex As Exception
            Throw New BosException(Me.getSchemaTableName & ":(Exception)", ex, CType(ex.TargetSite, Reflection.MethodInfo))
        Finally
            If Not IsNothing(IReader) Then IReader.Close()
        End Try
    End Function

    ''' <summary>
    ''' 將多筆物件回充填至單筆Bos物件(初始化單筆物件)。不產生DataTable
    ''' </summary>
    ''' <param name="IReader">傳入IDataReader。</param>
    ''' <remarks>
    ''' 大型物件，走constructIRDData
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/01/20	modify
    ''' </history>
    Protected Sub constructBosData(ByVal IReader As System.Data.IDataReader)
        clear()
        Try
            '大型物件，走reader 不回傳DataTable
            'Dim dt As DataTable = createDataTable(IReader)
            Dim dataTable As New DataTable(Me.getTableName())
            dataTable.Columns.AddRange(CreateDataColumns(IReader))
            'dataTable.BeginLoadData()

            While IReader.Read
                Dim bos As BosBase = newBos()

                Me.addFields(bos)
                bos.constructIRDData(IReader, DataTable.Columns)

                add(bos)

                Dim values(IReader.FieldCount - 1) As Object
                IReader.GetValues(values)
                ' dataTable.Rows.Add(values)
            End While

            'dataTable.EndLoadData()

            Me.setCurrentDataSet(New DataSet)
            'Me.getCurrentDataSet.Tables.Add(dataTable)

        Catch ex As Exception
            Throw New BosException(Me.getSchemaTableName & ":(Exception)", ex, CType(ex.TargetSite, Reflection.MethodInfo))
        End Try
    End Sub
#End Region

    ''' <summary>
    ''' 偵錯時，Show Select SQL。
    ''' </summary>
    ''' <param name="sSQL">傳入SQL語法。</param>
    ''' <param name="commandParameters">傳入IDataParameter陣列。</param>
    ''' <returns>String</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/01/20	modify
    ''' </history>
    Function ShowSelectSQL(ByVal sSQL As String, ByVal ParamArray commandParameters() As IDataParameter) As String

        Try
            If Not IsNothing(commandParameters) Then
                Dim sbSQL As New System.Text.StringBuilder(sSQL)
                For Each para As IDataParameter In commandParameters
                    If Not Titan.Utility.isValidateData(para.Value) Then
                        sbSQL = sbSQL.Replace(ProviderFactory.PositionPara & para.ParameterName, "null")
                    Else
                        If para.DbType = DbType.String Then
                            sbSQL = sbSQL.Replace(ProviderFactory.PositionPara & para.ParameterName, String.Format("'{0}'", para.Value))
                            sbSQL = sbSQL.Replace(ProviderFactory.PositionPara & " " & para.ParameterName, String.Format("'{0}'", para.Value))
                            sbSQL = sbSQL.Replace(" " & ProviderFactory.PositionPara & para.ParameterName, String.Format("'{0}'", para.Value))
                            sbSQL = sbSQL.Replace(" " & ProviderFactory.PositionPara & " " & para.ParameterName, String.Format("'{0}'", para.Value))
                        Else
                            sbSQL = sbSQL.Replace(String.Format("{0}{1}", ProviderFactory.PositionPara, para.ParameterName), CStr(para.Value))
                            sbSQL = sbSQL.Replace(ProviderFactory.PositionPara & " " & para.ParameterName, CStr(para.Value))
                            sbSQL = sbSQL.Replace(" " & ProviderFactory.PositionPara & para.ParameterName, CStr(para.Value))
                            sbSQL = sbSQL.Replace(" " & ProviderFactory.PositionPara & " " & para.ParameterName, CStr(para.Value))
                        End If
                    End If
                Next
                Return sbSQL.ToString
            End If
        Catch ex As Exception
        End Try
        Return sSQL
    End Function

#End Region

#Region "MustOverride"
    ''' <summary>
    ''' 每個繼承BosList的物件，Must Override this Function
    ''' 將多筆物件回充填至單筆Bos物件(初始化單筆物件)
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    MustOverride Function newBos() As BosBase
    ' MustOverride Function newBos(ByVal dataBaseManager As DatabaseManager) As BosBase
#End Region

#Region "getter/setter"

#Region "Table Name"
    ''' <summary>
    ''' 設定Table Name。
    ''' </summary>
    ''' <param name="sTableName">傳入Table Name。</param> 
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/01/20	modify
    ''' </history>
    Private Sub setTableName(ByVal sTableName As String)
        m_sTableName = sTableName
    End Sub

    ''' <summary>
    ''' 取得Table Name。
    ''' </summary>
    ''' <returns>String 傳回Table Name</returns>
    ''' <remarks>
    ''' EN_CASEMAIN
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/01/20	modify
    ''' </history>
    Public Function getTableName() As String
        Return Me.m_sTableName.ToUpper
    End Function

    ''' <summary>
    ''' 取得Schema Table Name。
    ''' </summary>
    ''' <returns>String 傳回Schema Table Name</returns>
    ''' <remarks>
    ''' BOTELOAN.EN_CASEMAIN
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/01/20	modify
    ''' </history>
    Public Function getSchemaTableName() As String
        'If com.Azion.NET.VB.Properties.getProvider = ProviderType.SqlClient Then
        '    Return m_sTableName
        'Else
        '    'Return getSchemaName().ToUpper & "." & m_sTableName
        '    Return Properties.getString("DBUserID").ToUpper & "." & m_sTableName
        'End If
        Return Properties.getSchemaName().ToUpper & "." & m_sTableName
    End Function

#End Region

#Region "Current DataSet"
    ''' <summary>
    ''' 設定Curr DataTable's Name。
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/01/20	modify
    ''' </history>
    Private Sub setDataTableName()
        If (Me.getCurrentDataSet.Tables.Count > 0) Then
            Me.getCurrentDataSet.Tables(0).TableName = Me.getTableName
        End If
    End Sub
    ''' <summary>
    ''' 設定Current DataSet。
    ''' </summary>
    ''' <param name="dataSet">傳入DataSet。</param>
    ''' <remarks>
    ''' 將物件由memory重新載入
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/01/20	modify
    ''' </history>
    Protected Sub setCurrentDataSet(ByVal dataSet As DataSet)
        Me.m_dsCurrent = dataSet
    End Sub

    ''' <summary>
    ''' 取得Current DataSet。
    ''' </summary>
    ''' <returns>DataSet 傳回物件 Current DataSet</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/01/20	modify
    ''' </history>
    Public Function getCurrentDataSet() As DataSet
        Return m_dsCurrent
    End Function
#End Region

#Region "DatabaseManager"
    ''' <summary>
    ''' 設定DatabaseManager Object。
    ''' </summary>
    ''' <param name="databaseManager">傳入DatabaseManager。</param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/01/20	modify
    ''' </history>
    Private Sub setDatabaseManager(ByVal databaseManager As DatabaseManager)
        Me.m_databaseManager = databaseManager
    End Sub
    ''' <summary>
    ''' 取得DatabaseManager。Object
    ''' </summary>
    ''' <returns>DatabaseManager 傳回DatabaseManager</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/01/20	modify
    ''' </history>
    Protected Function getDatabaseManager() As DatabaseManager
        Return Me.m_databaseManager
    End Function

#End Region

#Region "SQL Condition"
    ''' <summary>
    ''' 設定SQL Condition。
    ''' </summary>
    ''' <param name="sSQLCondition">傳入SQL Condition。</param>
    ''' <remarks>此物件可重複使用
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/01/20	modify
    ''' </history>
    Public Sub setSQLCondition(ByVal sSQLCondition As String)
        m_sSQLCondition = sSQLCondition
    End Sub

    ''' <summary>
    ''' 取得SQL Condition。
    ''' </summary>
    ''' <returns>String 傳回SQL Condition</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/01/20	modify
    ''' </history>
    Private Function getSQLCondition() As String
        Return m_sSQLCondition
    End Function
#End Region

#End Region

#Region "實作IEnumerator、IEnumerable 介面"

    ''' <summary>
    ''' 傳回可透過集合來重複的列舉值。
    ''' </summary>
    ''' <remarks>實作IEnumerator、IEnumerable 介面。
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/01/20	modify
    ''' </history>
    Function GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return CType(Me, IEnumerator)
    End Function

    ''' <summary>
    ''' 設定列舉值至它的初始位置，這是在集合中第一個元素之前。
    ''' </summary>
    ''' <remarks>實作IEnumerator、IEnumerable 介面。
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/01/20	modify
    ''' </history>

    Sub Reset() Implements IEnumerator.Reset
        m_currentIndex = -1
    End Sub

    ''' <summary>
    ''' 將列舉值推前至下一個集合的元素。
    ''' </summary>
    ''' <remarks>實作IEnumerator、IEnumerable 介面。
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/01/20	modify
    ''' </history>
    Function MoveNext() As Boolean Implements IEnumerator.MoveNext
        If m_currentIndex < Me.size - 1 Then
            m_currentIndex = m_currentIndex + 1
            Return True
        Else
            Return False
        End If
    End Function

    ''' <summary>
    ''' 取得集合中目前的項目。
    ''' </summary>
    ''' <remarks>實作IEnumerator、IEnumerable 介面。
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/01/20	modify
    ''' </history>
    ReadOnly Property Current() As Object Implements IEnumerator.Current
        Get
            Return Me.m_objects(m_currentIndex)
        End Get
    End Property
#End Region

#Region "實作ICloneable介面"
    'Function Clone() As Object Implements System.ICloneable.Clone

    '    '首先我們建立指定類型的一個實例 
    '    Dim newObject As BosList = Activator.CreateInstance(Me.GetType())
    '    '我們取得新的類型實例的欄位陣列。
    '    Dim fields() As System.Reflection.FieldInfo = newObject.GetType().GetFields()
    '    Dim i As Integer = 0

    '    For Each fi As System.Reflection.FieldInfo In fields
    '        '我們判斷欄位是否支援ICloneable介面。
    '        Dim ICloneType As System.Type = fi.FieldType.GetInterface("ICloneable", True)
    '        If Not IsNothing(ICloneType) Then
    '            '取得物件的Icloneable介面。
    '            Dim IClone As ICloneable = CType(fi.GetValue(Me), ICloneable)
    '            '使用克隆方法給欄位設定新值。
    '            fields(i).SetValue(newObject, IClone.Clone())
    '        Else
    '            ' 如果該欄位部支援Icloneable介面，直接設置即可。
    '            fields(i).SetValue(newObject, fi.GetValue(Me))
    '        End If

    '        '檢查該物件是否支援IEnumerable介面，如果支援， 
    '        '枚舉其所有項並檢查他們是否支援IList 或 IDictionary 介面。            
    '        Dim IEnumerableType As System.Collections.IEnumerable = fi.FieldType.GetInterface("IEnumerable", True)
    '        If Not IsNothing(IEnumerableType) Then
    '            ' 取得該欄位的IEnumerable介面
    '            Dim IEnum As System.Collections.IEnumerable = CType(fi.GetValue(Me), System.Collections.IEnumerable)
    '            '這個版本支援IList 或 IDictionary 介面來?代集合。
    '            Dim IListType As System.Type = fields(i).FieldType.GetInterface("IList", True)
    '            Dim IDicType As System.Type = fields(i).FieldType.GetInterface("IDictionary", True)

    '            Dim j As Integer = 0
    '            If Not IsNothing(IListType) Then
    '                '取得IList介面。
    '                Dim list As IList = CType(fields(i).GetValue(newObject), System.Collections.IList)
    '                For Each obj As Object In IEnum
    '                    '查看當前項是否支援支援ICloneable 介面。
    '                    ICloneType = obj.GetType().GetInterface("ICloneable", True)
    '                    If Not IsNothing(ICloneType) Then
    '                        '如果支援ICloneable 介面，
    '                        '我們用它?設置列表中的物件的克隆
    '                        Dim cloned As ICloneable = CType(obj, ICloneable)
    '                        list(j) = cloned.Clone()
    '                    End If
    '                    '注意：如果列表中的項不支援ICloneable介面，那?
    '                    '在克隆列表的項將與原列表對應項相同
    '                    '（只要該類型是引用類型）
    '                    j += 1
    '                Next
    '            ElseIf Not IsNothing(IDicType) Then
    '                '取得IDictionary 介面
    '                Dim dic As IDictionary = CType(fields(i).GetValue(newObject), IDictionary)
    '                j = 0
    '                For Each de As DictionaryEntry In IEnum
    '                    '查看當前項是否支援支援ICloneable 介面。
    '                    ICloneType = de.Value.GetType().GetInterface("ICloneable", True)
    '                    If Not IsNothing(ICloneType) Then
    '                        Dim cloned As ICloneable = CType(de.Value, ICloneable)
    '                        dic(de.Key) = cloned.Clone()
    '                    End If
    '                    j += 1
    '                Next
    '            End If
    '        End If
    '        i += 1
    '    Next
    '    Return newObject
    'End Function
#End Region

#Region "inner function"

#Region " for ExecuteReader function"
    ''' <summary>
    ''' 將多筆物件回充填至單筆Bos物件(初始化單筆物件)。
    ''' </summary>
    ''' <param name="IReader">傳入IDataReader。</param>
    ''' <remarks>
    ''' 大型物件，走constructIRDData
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/01/20	modify
    ''' </history>
    Protected Sub constructIRDData(ByVal IReader As System.Data.IDataReader)
        clear()
        Try
            '大型物件，走reader 不回傳DataTable
            'Dim dt As DataTable = createDataTable(IReader)
            Dim dataTable As New DataTable(Me.getTableName())
            dataTable.Columns.AddRange(CreateDataColumns(IReader))
            dataTable.BeginLoadData()

            While IReader.Read
                Dim bos As BosBase = newBos()

                Me.addFields(bos)
                bos.constructIRDData(IReader, dataTable.Columns)

                add(bos)

                Dim values(IReader.FieldCount - 1) As Object
                IReader.GetValues(values)
                dataTable.Rows.Add(values)
            End While

            dataTable.EndLoadData()

            Me.setCurrentDataSet(New DataSet)
            Me.getCurrentDataSet.Tables.Add(dataTable)

        Catch ex As Exception 
            Throw New BosException(Me.getSchemaTableName & ":(Exception)", ex, CType(ex.TargetSite, Reflection.MethodInfo))
        End Try
    End Sub

    ''' <summary>
    ''' 產生DataTable的欄位
    ''' </summary>
    ''' <param name="IReader">傳入IDataReader。</param>
    ''' <returns>DataColumn()</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/01/20	modify
    ''' </history>
    Private Function CreateDataColumns(ByVal IReader As System.Data.IDataReader) As DataColumn()
        Dim cols(IReader.FieldCount - 1) As DataColumn

        For i As Integer = 0 To cols.Length - 1
            cols(i) = New DataColumn(IReader.GetName(i), IReader.GetFieldType(i))
        Next

        Return cols
    End Function

#End Region

#Region " for ExecuteDataSet function"

    ''' <summary>
    ''' 將多筆物件回充填至單筆Bos物件(初始化單筆物件)。
    ''' </summary>
    ''' <param name="dataTable">傳入DataTable。</param>
    ''' <remarks>
    ''' 可以被Overridable
    ''' you can see EN_04_RPT_ATTACH04_ITEMList class。
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/01/20	modify
    ''' </history>
    Protected Overridable Sub constructData(ByVal dataTable As System.Data.DataTable)
        clear()
        Try

            For Each row As DataRow In dataTable.Rows
                Dim bos As BosBase = newBos()
                'synchronized
                'SyncLock bos
                'If Not Me.getTableName.ToUpper.Equals("EN_LINK_TAB") Then
                '    Me.newBosBase(bos)
                'End If
                Me.addFields(bos)
                bos.constructData(row)
                add(bos)
                'End SyncLock
            Next
        Catch ex As Exception
            Throw New BosException(Me.getSchemaTableName & ":(Exception)", ex, CType(ex.TargetSite, Reflection.MethodInfo))
        End Try
    End Sub
#End Region

#Region "Execute"

    ''' <summary>
    ''' ExecuteReader
    ''' </summary>
    ''' <param name="commandType">傳入CommandType</param>
    ''' <param name="commandText">傳入commandText</param>
    ''' <param name="commandParameters">傳入IDataParameter()</param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/01/20	modify
    ''' </history>
    Private Function ExecuteReader(ByVal dataBaseManager As DatabaseManager, ByVal commandType As CommandType, _
                                ByVal commandText As String, _
                                ByVal ParamArray commandParameters() As IDataParameter) As IDataReader
        Try
            Dim IReader As IDataReader
            If Me.m_databaseManager.isTransaction AndAlso Not Me.m_databaseManager.isCloseTransaction Then
                dataBaseManager.beginTran()
                IReader = DBObject.ExecuteReader(dataBaseManager, commandType, commandText, commandParameters)
                'setDataTableName()
                Return IReader
            ElseIf Not IsNothing(Me.getDatabaseManager.getConnection()) Then
                IReader = DBObject.ExecuteReader(dataBaseManager, commandType, commandText, commandParameters)
                'setDataTableName()
                Return IReader
            Else
                Return Nothing
            End If
        Catch ex As Exception
            Throw New BosException(Me.getSchemaTableName & ":(Exception)", ex, CType(ex.TargetSite, Reflection.MethodInfo))
        End Try
    End Function

    ''' <summary>
    ''' ExecuteDataset
    ''' </summary>
    ''' <param name="commandType">傳入CommandType</param>
    ''' <param name="commandText">傳入commandText</param>
    ''' <param name="commandParameters">傳入IDataParameter()</param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/01/20	modify
    ''' </history>
    Private Function ExecuteDataset(ByVal commandType As CommandType, _
                                ByVal commandText As String, _
                                ByVal ParamArray commandParameters() As IDataParameter) As DataSet
        Try

            If Me.m_databaseManager.isTransaction AndAlso Not Me.m_databaseManager.isCloseTransaction Then
                Me.setCurrentDataSet(DBObject.ExecuteDataset(Me.getDatabaseManager.getTransaction, commandType, commandText, commandParameters))
                setDataTableName()
                Return Me.getCurrentDataSet()
            ElseIf Not IsNothing(Me.getDatabaseManager.getConnection()) Then
                Me.setCurrentDataSet(DBObject.ExecuteDataset(Me.getDatabaseManager.getConnection, commandType, commandText, commandParameters))
                setDataTableName()
                Return Me.getCurrentDataSet()
            Else
                Return Nothing
            End If
        Catch ex As Exception
            Throw New BosException(Me.getSchemaTableName & ":(Exception)", ex, CType(ex.TargetSite, Reflection.MethodInfo))
        End Try
    End Function
#End Region
    ''' <summary>
    ''' 加載Bos Object 的欄位名稱、值及資料型別等屬性值
    ''' </summary>
    ''' <param name="bos">傳入BosBase</param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/01/20	modify
    ''' </history>
    Private Overloads Function addFields(ByVal bos As BosBase) As BosBase
        For Each attrMeta As BosAttrMeta In m_BosAttrMetas
            Dim attr As New BosAttribute
            attr.setAttrMeta(attrMeta)
            bos.addAttribute(attr)
        Next
        Return bos
    End Function

    ''' <summary>
    ''' 提供非此 BosList 物件之加載欄位屬性值，如案例三
    ''' </summary>
    ''' <param name="attr">傳入BosAttrMeta</param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/01/20	modify
    ''' </history>
    Protected Sub addAttribute(ByVal attr As BosAttrMeta)
        m_BosAttrMetas.Add(attr)
    End Sub

    ''' <summary>
    ''' 提供非此 BosList 物件之加載欄位屬性值，如案例三
    ''' </summary>
    ''' <param name="sColName">傳入Column Name</param>
    ''' <param name="DBType">傳入System.Data.DbType</param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/01/20	modify
    ''' </history>
    Protected Sub addAttribute(ByVal sColName As String, ByVal DBType As System.Data.DbType)
        Dim attrMeta As New BosAttrMeta
        attrMeta.setColName(sColName.ToUpper)
        attrMeta.setDataType(DBType)
        m_BosAttrMetas.Add(attrMeta)
    End Sub

    ''' <summary>
    ''' 加載BosList object
    ''' </summary>
    ''' <param name="bos">傳入BosBase</param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/01/20	modify
    ''' </history>
    Public Sub add(ByVal bos As BosBase)
        SyncLock m_objects.SyncRoot
            m_objects.Add(bos)
        End SyncLock
    End Sub

    ''' <summary>
    ''' 移除單筆BosList object
    ''' </summary>
    ''' <param name="bos">傳入BosBase</param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/01/20	modify
    ''' </history>
    Public Sub remove(ByVal bos As BosBase)
        m_objects.Remove(bos)
    End Sub

    ''' <summary>
    ''' 清除所有BosList object
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/01/20	modify
    ''' </history>
    Public Sub clear()
        Me.setSQLCondition("")
        m_objects.Clear()
        Me.Reset()
    End Sub

    ''' <summary>
    ''' 取得BosList object
    ''' </summary>
    ''' <param name="index">傳入Integer</param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/01/20	modify
    ''' </history>
    Public Function item(ByVal index As Integer) As BosBase
        Dim bos As BosBase = Nothing
        Try
            bos = CType(m_objects.Item(index), BosBase)
        Catch ex As Exception
            Throw New BosException(Me.getSchemaTableName & ":(Exception)", ex, CType(ex.TargetSite, Reflection.MethodInfo))
        End Try
        Return bos
    End Function

    ''' <summary>
    ''' 取得BosList object size
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/01/20	modify
    ''' </history>
    Public Function size() As Integer
        Return m_objects.Count
    End Function
#End Region

#Region "預計實作的Function"
#Region "Serializable"

    'Public Sub New(ByVal info As SerializationInfo, ByVal context As StreamingContext)

    '    MyBase.New(info, context)
    'End Sub

    ''A method called when serializing a Singleton.
    '<SecurityPermissionAttribute(SecurityAction.LinkDemand, _
    'Flags:=SecurityPermissionFlag.SerializationFormatter)> _
    '             Public Overridable Sub GetObjectData(ByVal info As SerializationInfo, _
    '   ByVal context As StreamingContext) _
    '   Implements ISerializable.GetObjectData

    '    info.AddValue("m_SezMe", Me)
    '    info.AddValue("m_objects", m_objects)
    '    info.AddValue("m_attributes", m_attributes)
    '    info.AddValue("m_transaction", m_transaction)
    '    info.AddValue("m_connection", m_connection)
    '    info.AddValue("m_databaseManager", m_databaseManager)
    '    info.AddValue("m_sTableName", m_sTableName)
    '    info.AddValue("m_arrayPrimaryKeys", m_arrayPrimaryKeys)
    '    info.AddValue("m_sOrderBySQL", m_sOrderBySQL)

    '    info.AddValue("m_currentIndex", m_currentIndex)
    '    info.AddValue("m_dsCurrent", m_dsCurrent)

    '    Utility.Debug("=========info============" & info.FullTypeName)
    '    Utility.Debug("=========info.MemberCount()============" & info.MemberCount())
    'End Sub

    'Private Sub OnDeserialization(ByVal sender As Object) _
    '   Implements IDeserializationCallback.OnDeserialization
    '    ' After being deserialized
    'End Sub

    'Sub New()
    'End Sub
#End Region
#End Region

#Region "預計廢除的Function"

#Region "update for datatable"
    ''' <summary>
    ''' 取得Schema Name(User ID)。
    ''' </summary>
    ''' <returns>String 傳回Schema Table Name</returns>
    ''' <remarks>
    ''' BOTELOAN
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/01/20	modify
    ''' </history>
    <ObsoleteAttribute("This function will be removed from future Versions.20110715")> _
    Private Function getSchemaName() As String
        Return "" 'Me.m_databaseManager.getSchemaName()
    End Function

    ''' <summary>
    ''' This function will be removed from future Versions.Use another function 'loadReaderBySQL'
    ''' </summary>
    ''' <param name="sSQL">傳入SQL語法。</param>
    ''' <param name="commandParameters">傳入IDataParameter陣列。</param>
    ''' <returns>Integer 傳回記錄的筆數</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/01/20	modify
    ''' </history>
    ''' 
    <ObsoleteAttribute("This function will be removed from future Versions.Use another function 'loadReaderBySQL'")> _
    Private Function loadReadeBySQL(ByVal sSQL As String, ByVal ParamArray commandParameters() As IDataParameter) As Integer
        loadReaderBySQL(sSQL, commandParameters)
    End Function
#End Region
#End Region

End Class
