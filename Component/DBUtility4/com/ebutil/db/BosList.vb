Option Explicit On
Option Strict On
Imports System.Runtime.Serialization
Imports System.Security.Permissions
Imports System.ComponentModel
Imports System.Data
Imports System.Collections

''' <summary>
'''  The List of the Business objects (BOS)�C
''' </summary>
''' <remarks>
''' BosList ��@IEnumerator�BIEnumerable ����
''' �Ҧ���@�h�� Table Object�����~��BosList�C
''' 
''' </remarks>
''' <example> This sample shows how to inherits the BosList class based on EN_CASEMAIN which has the primary key
''' <code>
''' '�Ъ`�N�R�W��k 
''' '----------------------------------------------
''' '|Table Name  | BosBase     |  BosList         |
''' '|------------| ------------|  ----------------|
''' '|EN_CASEMAIN | EN_CASEMAIN |  EN_CASEMAINList |
''' '----------------------------------------------
''' Public Class EN_CASEMAINList Inherits BosList
'''     '�غc�l
'''     Sub New(ByVal dbManager As DatabaseManager)
'''          MyBase.New("EN_CASEMAIN", dbManager)
'''     End Sub
''' 
'''     'must Overrides Function newBos
'''     Overrides Function newBos() As BosBase
'''          Return New EN_CASEMAIN(MyBase.getDatabaseManager)
'''     End Function
'''     
'''     '�רҤ@:���gSQL�y�k�A�������Ѽ�
'''     '�����B�@�y�k:[select * from en_casemain where GOBRID=sBrId]
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
'''     '�רҤG:�gSQL�y�k�A�õ��Ѽ�
'''     '�����B�@�y�k:[programmer ����SQL = select * from en_casemain where GOBRID=sBrId]
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
'''    '�רҤT:�gJoin SQL�y�k�A�õ��Ѽ�
'''    '�����B�@�y�k:[programmer ����SQL = select  a.caseid,a.brid,b.caseitemid caseitemid,b.loanitem �«H���ئW��,b.datadate ��Ƥ�� from En_Casemain a ,En_Casedetailcond  b where a.caseid=b.caseid and b.caseid=sCaseID]
'''    '�p�G�O�h�U���PTable Join�b�@�_�ɡA
'''    '�b������(EN_CASEMAINList)����ʼW�[����ݩʡA�~�i���o���PTable(En_Casedetailcond)�����(caseitemid�Bloanitem�Bdatadate)
'''     Function loadCaseItem(ByVal sCaseId As String) As Integer
'''          Try
'''               If (Not IsNothing(sCaseId) AndAlso sCaseId.Length > 0) Then
'''                    Dim sqlStr As String = ""
'''                    sqlStr = "select  a.caseid,a.brid,b.caseitemid caseitemid,b.loanitem �«H���ئW��,b.datadate ��Ƥ�� from En_Casemain a ,En_Casedetailcond  b where a.caseid=b.caseid " + _
'''                    " and CASEID=" + ProviderFactory.PositionPara + "CASEID "
''' 
'''                    '�U���o�T������ݩ�En_Casedetailcond���A�ëD�ݩ�En_CaseMain�����
'''                    '�]�����NEn_Casedetailcond���T������ʥ[�JEN_CASEMAINList������A�p��������o�o�T�U��쪺��
'''                    Me.addAttribute("caseitemid", System.Data.DbType.Decimal)'�p�G��ƫ��O�ONumber-->System.Data.DbType.Decimal
'''                    Me.addAttribute("�«H���ئW��", System.Data.DbType.String)'�p�G��ƫ��O�OString-->System.Data.DbType.String
'''                    Me.addAttribute("��Ƥ��", System.Data.DbType.Date)'�p�G��ƫ��O�ODate-->System.Data.DbType.Date
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
    'ENOP �ե�W��
    '    Private m_sDLLName As String '= Properties.m_sDLLName
    'ENOP �R�W�Ŷ�
    'Private m_sNameSpase As String  '= Properties.m_sNameSpase

    ''' <summary>
    ''' �s�u����
    ''' </summary>
    ''' <remarks></remarks>
    Private m_databaseManager As DatabaseManager

    ''' <summary>
    ''' Table Name
    ''' </summary>
    ''' <remarks></remarks>
    Private m_sTableName As String

    ''' <summary>
    ''' �s��BosBase����}�C
    ''' </summary>
    ''' <remarks></remarks>
    Private m_objects As New ArrayList
    ''' <summary>
    ''' �s��BosAttrMeta��쪫��}�C
    ''' </summary>
    ''' <remarks></remarks>
    Private m_BosAttrMetas As New ArrayList

    ''' <summary>
    ''' �̻ݨD�վ�Condition
    ''' </summary>
    ''' <remarks></remarks>
    Private m_sSQLCondition As String

    ''' <summary>
    ''' �ثe�O��������
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

#Region "�غc�l"

    ''' <summary>
    ''' ��l�� BosList ���O���s������� (Instance)�C
    ''' </summary>
    ''' <param name="sTableName">�ǤJTable Name(�j�g)</param>
    ''' <param name="databaseManager">�ǤJ DatabaseManager</param>
    ''' <remarks>�غc�l
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/01/20	modify
    ''' </history>
    Sub New(ByVal sTableName As String, ByVal databaseManager As DatabaseManager)
        'ENOP �R�W�Ŷ�
        'm_sNameSpase = Me.GetType.Namespace
        ''ENOP �ե�W��
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
    ''' ���o����Table���Ҧ��O���C
    ''' </summary>
    ''' <returns>Integer �Ǧ^�O�������ơC</returns>
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
    ''' �ϥ�DataAdatpter.Fill(DataTable)��k�A��������Table���O���C
    ''' </summary>
    ''' <param name="sSQL">�ǤJSQL�y�k�C</param>
    ''' <returns>Integer �Ǧ^�O�������ơC</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/01/20	modify
    ''' </history>
    Protected Function loadBySQL(ByVal sSQL As String) As Integer
        Return loadBySQL(sSQL, Nothing)
    End Function

    ''' <summary>
    ''' �ϥ�DataAdatpter.Fill(DataTable)��k�A��������Table���O���C
    ''' </summary>
    ''' <param name="commandParameters">�ǤJIDataParameter�}�C�C</param>
    ''' <returns>Integer �Ǧ^�O�������ơC</returns>
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
    ''' �ϥ�DataAdatpter.Fill(DataTable)��k�A��������Table���O���C
    ''' </summary>
    ''' <param name="sSQL">�ǤJSQL�y�k�C</param>
    ''' <param name="commandParameters">�ǤJIDataParameter�}�C�C</param>
    ''' <returns>Integer �Ǧ^�O��������</returns>
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
    ''' �ϥ�DataAdatpter.Fill(DataTable)��k�A��������Table���O���C
    ''' </summary>
    ''' <param name="sSQL">�ǤJSQL�y�k�C</param>
    ''' <returns>Integer �Ǧ^�O�������ơC</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/01/20	modify
    ''' </history>
    Protected Function loadBySQLOnlyDs(ByVal sSQL As String) As Integer
        Return loadBySQLOnlyDs(sSQL, Nothing)
    End Function

    ''' <summary>
    ''' �ϥ�DataAdatpter.Fill(DataTable)��k�A��������Table���O���C
    ''' </summary>
    ''' <param name="commandParameters">�ǤJIDataParameter�}�C�C</param>
    ''' <returns>Integer �Ǧ^�O�������ơC</returns>
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
    ''' �ϥ�DataAdatpter.Fill(DataTable)��k�A��������Table���O���C
    ''' </summary>
    ''' <param name="sSQL">�ǤJSQL�y�k�C</param>
    ''' <param name="commandParameters">�ǤJIDataParameter�}�C�C</param>
    ''' <returns>Integer �Ǧ^�O��������</returns>
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
    ''' �N�h������^�R��ܳ浧Bos����(��l�Ƴ浧����)�C������DataTable
    ''' </summary>
    ''' <param name="IReader">�ǤJIDataReader�C</param>
    ''' <remarks>
    ''' �j������A��constructIRDData
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/01/20	modify
    ''' </history>
    Protected Sub constructBosData(ByVal IReader As System.Data.IDataReader)
        clear()
        Try
            '�j������A��reader ���^��DataTable
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
    ''' �����ɡAShow Select SQL�C
    ''' </summary>
    ''' <param name="sSQL">�ǤJSQL�y�k�C</param>
    ''' <param name="commandParameters">�ǤJIDataParameter�}�C�C</param>
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
    ''' �C���~��BosList������AMust Override this Function
    ''' �N�h������^�R��ܳ浧Bos����(��l�Ƴ浧����)
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    MustOverride Function newBos() As BosBase
    ' MustOverride Function newBos(ByVal dataBaseManager As DatabaseManager) As BosBase
#End Region

#Region "getter/setter"

#Region "Table Name"
    ''' <summary>
    ''' �]�wTable Name�C
    ''' </summary>
    ''' <param name="sTableName">�ǤJTable Name�C</param> 
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/01/20	modify
    ''' </history>
    Private Sub setTableName(ByVal sTableName As String)
        m_sTableName = sTableName
    End Sub

    ''' <summary>
    ''' ���oTable Name�C
    ''' </summary>
    ''' <returns>String �Ǧ^Table Name</returns>
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
    ''' ���oSchema Table Name�C
    ''' </summary>
    ''' <returns>String �Ǧ^Schema Table Name</returns>
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
    ''' �]�wCurr DataTable's Name�C
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
    ''' �]�wCurrent DataSet�C
    ''' </summary>
    ''' <param name="dataSet">�ǤJDataSet�C</param>
    ''' <remarks>
    ''' �N�����memory���s���J
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/01/20	modify
    ''' </history>
    Protected Sub setCurrentDataSet(ByVal dataSet As DataSet)
        Me.m_dsCurrent = dataSet
    End Sub

    ''' <summary>
    ''' ���oCurrent DataSet�C
    ''' </summary>
    ''' <returns>DataSet �Ǧ^���� Current DataSet</returns>
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
    ''' �]�wDatabaseManager Object�C
    ''' </summary>
    ''' <param name="databaseManager">�ǤJDatabaseManager�C</param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/01/20	modify
    ''' </history>
    Private Sub setDatabaseManager(ByVal databaseManager As DatabaseManager)
        Me.m_databaseManager = databaseManager
    End Sub
    ''' <summary>
    ''' ���oDatabaseManager�CObject
    ''' </summary>
    ''' <returns>DatabaseManager �Ǧ^DatabaseManager</returns>
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
    ''' �]�wSQL Condition�C
    ''' </summary>
    ''' <param name="sSQLCondition">�ǤJSQL Condition�C</param>
    ''' <remarks>������i���ƨϥ�
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/01/20	modify
    ''' </history>
    Public Sub setSQLCondition(ByVal sSQLCondition As String)
        m_sSQLCondition = sSQLCondition
    End Sub

    ''' <summary>
    ''' ���oSQL Condition�C
    ''' </summary>
    ''' <returns>String �Ǧ^SQL Condition</returns>
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

#Region "��@IEnumerator�BIEnumerable ����"

    ''' <summary>
    ''' �Ǧ^�i�z�L���X�ӭ��ƪ��C�|�ȡC
    ''' </summary>
    ''' <remarks>��@IEnumerator�BIEnumerable �����C
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/01/20	modify
    ''' </history>
    Function GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return CType(Me, IEnumerator)
    End Function

    ''' <summary>
    ''' �]�w�C�|�Ȧܥ�����l��m�A�o�O�b���X���Ĥ@�Ӥ������e�C
    ''' </summary>
    ''' <remarks>��@IEnumerator�BIEnumerable �����C
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/01/20	modify
    ''' </history>

    Sub Reset() Implements IEnumerator.Reset
        m_currentIndex = -1
    End Sub

    ''' <summary>
    ''' �N�C�|�ȱ��e�ܤU�@�Ӷ��X�������C
    ''' </summary>
    ''' <remarks>��@IEnumerator�BIEnumerable �����C
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
    ''' ���o���X���ثe�����ءC
    ''' </summary>
    ''' <remarks>��@IEnumerator�BIEnumerable �����C
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

#Region "��@ICloneable����"
    'Function Clone() As Object Implements System.ICloneable.Clone

    '    '�����ڭ̫إ߫��w�������@�ӹ�� 
    '    Dim newObject As BosList = Activator.CreateInstance(Me.GetType())
    '    '�ڭ̨��o�s��������Ҫ����}�C�C
    '    Dim fields() As System.Reflection.FieldInfo = newObject.GetType().GetFields()
    '    Dim i As Integer = 0

    '    For Each fi As System.Reflection.FieldInfo In fields
    '        '�ڭ̧P�_���O�_�䴩ICloneable�����C
    '        Dim ICloneType As System.Type = fi.FieldType.GetInterface("ICloneable", True)
    '        If Not IsNothing(ICloneType) Then
    '            '���o����Icloneable�����C
    '            Dim IClone As ICloneable = CType(fi.GetValue(Me), ICloneable)
    '            '�ϥΧJ����k�����]�w�s�ȡC
    '            fields(i).SetValue(newObject, IClone.Clone())
    '        Else
    '            ' �p�G����쳡�䴩Icloneable�����A�����]�m�Y�i�C
    '            fields(i).SetValue(newObject, fi.GetValue(Me))
    '        End If

    '        '�ˬd�Ӫ���O�_�䴩IEnumerable�����A�p�G�䴩�A 
    '        '�T�|��Ҧ������ˬd�L�̬O�_�䴩IList �� IDictionary �����C            
    '        Dim IEnumerableType As System.Collections.IEnumerable = fi.FieldType.GetInterface("IEnumerable", True)
    '        If Not IsNothing(IEnumerableType) Then
    '            ' ���o����쪺IEnumerable����
    '            Dim IEnum As System.Collections.IEnumerable = CType(fi.GetValue(Me), System.Collections.IEnumerable)
    '            '�o�Ӫ����䴩IList �� IDictionary ������?�N���X�C
    '            Dim IListType As System.Type = fields(i).FieldType.GetInterface("IList", True)
    '            Dim IDicType As System.Type = fields(i).FieldType.GetInterface("IDictionary", True)

    '            Dim j As Integer = 0
    '            If Not IsNothing(IListType) Then
    '                '���oIList�����C
    '                Dim list As IList = CType(fields(i).GetValue(newObject), System.Collections.IList)
    '                For Each obj As Object In IEnum
    '                    '�d�ݷ�e���O�_�䴩�䴩ICloneable �����C
    '                    ICloneType = obj.GetType().GetInterface("ICloneable", True)
    '                    If Not IsNothing(ICloneType) Then
    '                        '�p�G�䴩ICloneable �����A
    '                        '�ڭ̥Υ�?�]�m�C�������󪺧J��
    '                        Dim cloned As ICloneable = CType(obj, ICloneable)
    '                        list(j) = cloned.Clone()
    '                    End If
    '                    '�`�N�G�p�G�C���������䴩ICloneable�����A��?
    '                    '�b�J���C�����N�P��C��������ۦP
    '                    '�]�u�n�������O�ޥ������^
    '                    j += 1
    '                Next
    '            ElseIf Not IsNothing(IDicType) Then
    '                '���oIDictionary ����
    '                Dim dic As IDictionary = CType(fields(i).GetValue(newObject), IDictionary)
    '                j = 0
    '                For Each de As DictionaryEntry In IEnum
    '                    '�d�ݷ�e���O�_�䴩�䴩ICloneable �����C
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
    ''' �N�h������^�R��ܳ浧Bos����(��l�Ƴ浧����)�C
    ''' </summary>
    ''' <param name="IReader">�ǤJIDataReader�C</param>
    ''' <remarks>
    ''' �j������A��constructIRDData
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/01/20	modify
    ''' </history>
    Protected Sub constructIRDData(ByVal IReader As System.Data.IDataReader)
        clear()
        Try
            '�j������A��reader ���^��DataTable
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
    ''' ����DataTable�����
    ''' </summary>
    ''' <param name="IReader">�ǤJIDataReader�C</param>
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
    ''' �N�h������^�R��ܳ浧Bos����(��l�Ƴ浧����)�C
    ''' </summary>
    ''' <param name="dataTable">�ǤJDataTable�C</param>
    ''' <remarks>
    ''' �i�H�QOverridable
    ''' you can see EN_04_RPT_ATTACH04_ITEMList class�C
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
    ''' <param name="commandType">�ǤJCommandType</param>
    ''' <param name="commandText">�ǤJcommandText</param>
    ''' <param name="commandParameters">�ǤJIDataParameter()</param>
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
    ''' <param name="commandType">�ǤJCommandType</param>
    ''' <param name="commandText">�ǤJcommandText</param>
    ''' <param name="commandParameters">�ǤJIDataParameter()</param>
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
    ''' �[��Bos Object �����W�١B�Ȥθ�ƫ��O���ݩʭ�
    ''' </summary>
    ''' <param name="bos">�ǤJBosBase</param>
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
    ''' ���ѫD�� BosList ���󤧥[������ݩʭȡA�p�רҤT
    ''' </summary>
    ''' <param name="attr">�ǤJBosAttrMeta</param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/01/20	modify
    ''' </history>
    Protected Sub addAttribute(ByVal attr As BosAttrMeta)
        m_BosAttrMetas.Add(attr)
    End Sub

    ''' <summary>
    ''' ���ѫD�� BosList ���󤧥[������ݩʭȡA�p�רҤT
    ''' </summary>
    ''' <param name="sColName">�ǤJColumn Name</param>
    ''' <param name="DBType">�ǤJSystem.Data.DbType</param>
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
    ''' �[��BosList object
    ''' </summary>
    ''' <param name="bos">�ǤJBosBase</param>
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
    ''' �����浧BosList object
    ''' </summary>
    ''' <param name="bos">�ǤJBosBase</param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/01/20	modify
    ''' </history>
    Public Sub remove(ByVal bos As BosBase)
        m_objects.Remove(bos)
    End Sub

    ''' <summary>
    ''' �M���Ҧ�BosList object
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
    ''' ���oBosList object
    ''' </summary>
    ''' <param name="index">�ǤJInteger</param>
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
    ''' ���oBosList object size
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

#Region "�w�p��@��Function"
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

#Region "�w�p�o����Function"

#Region "update for datatable"
    ''' <summary>
    ''' ���oSchema Name(User ID)�C
    ''' </summary>
    ''' <returns>String �Ǧ^Schema Table Name</returns>
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
    ''' <param name="sSQL">�ǤJSQL�y�k�C</param>
    ''' <param name="commandParameters">�ǤJIDataParameter�}�C�C</param>
    ''' <returns>Integer �Ǧ^�O��������</returns>
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
