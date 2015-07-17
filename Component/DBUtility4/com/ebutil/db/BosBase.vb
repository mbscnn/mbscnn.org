Option Explicit On
Option Strict On
Imports System.Runtime.CompilerServices
Imports System.Runtime.Serialization.Formatters.Binary
Imports System.Data
Imports System.Collections
Imports System.IO
Imports System.Xml.Serialization


''' <summary>
'''  The Business objects (BOS)�C
''' </summary>
''' <remarks>
''' Bosbase ��@IEnumerator�BIEnumerable ����
''' �Ҧ���@�浧 Table Object�����~��Bosbase�C
''' 
''' </remarks>
''' <example> This sample shows how to inherits the Bosbase class based on EN_CASEMAIN which has the primary key
''' <code>
''' '�Ъ`�N�R�W��k 
''' '----------------------------------------------
''' '|Table Name  | BosBase     |  BosList         |
''' '|------------| ------------|  ----------------|
''' '|EN_CASEMAIN | EN_CASEMAIN |  EN_CASEMAINList |
''' '----------------------------------------------
''' Public Class EN_CASEMAIN Inherits Bosbase
'''     '�غc�l
'''     Sub New()
'''          MyBase.New()
'''          Me.setPrimaryKeys()
'''     End Sub
'''     Sub New(ByVal dbManager As DatabaseManager)
'''         MyBase.new("EN_CASEMAIN", dbManager)
'''     End Sub
''' 
'''     'must Overrides Function newBos
'''     Overrides Function newBos() As BosBase
'''          Return New EN_CASEMAIN(MyBase.getDatabaseManager)
'''     End Function
'''     
'''     '�רҤ@:���gSQL�y�k�A�������Ѽ�
'''     '�����B�@�y�k:[select * from en_casemain where CaseID=sCaseID]
'''     Function loadByPK(ByVal sCaseID As String) As Integer
'''          Try
'''               If (Not IsNothing(sCaseID) AndAlso sCaseID.Length > 0) Then
'''                    Dim paras(0) As System.Data.IDbDataParameter
'''                    paras(0) = ProviderFactory.CreateDataParameter("CaseID", sCaseID)
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
'''     '�����B�@�y�k:[programmer ����SQL = select * from en_casemain where CaseID=sCaseID]
'''     Function loadByPK(ByVal sCaseID As String) As Integer
'''          Try
'''               If (Not IsNothing(sCaseID) AndAlso sCaseID.Length > 0) Then
'''                    Dim sSQL As String="select * from en_casemain where CaseID" + ProviderFactory.PositionPara + "CaseID"
'''                    Dim paras(0) As System.Data.IDbDataParameter
'''                    paras(0) = ProviderFactory.CreateDataParameter("CaseID", sCaseID)
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
'''    '�����B�@�y�k:[programmer ����SQL = select  a.caseid,a.brid,b.caseitemid caseitemid,b.loanitem �«H���ئW��,b.datadate ��Ƥ�� from En_Casemain a ,En_Casedetailcond  b where a.caseid=b.caseid and b.caseid=sCaseID and b.caseitemid=1]
'''    '�p�G�O�h�U���PTable Join�b�@�_�ɡA
'''    '�b������(EN_CASEMAINList)����ʼW�[����ݩʡA�~�i���o���PTable(En_Casedetailcond)�����(caseitemid�Bloanitem�Bdatadate)
'''     Function loadCaseItem(ByVal sCaseId As String) As Integer
'''          Try
'''               If (Not IsNothing(sCaseId) AndAlso sCaseId.Length > 0) Then
'''                    Dim sqlStr As String = ""
'''                    sqlStr = "select  a.caseid,a.brid,b.caseitemid caseitemid,b.loanitem �«H���ئW��,b.datadate ��Ƥ�� from En_Casemain a ,En_Casedetailcond  b where a.caseid=b.caseid " + _
'''                    " and CASEID=" + ProviderFactory.PositionPara + "CASEID and b.caseitemid=1 "
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
''' 	[Titan]	2010/05/20	modify
''' </history>

Public Class BosBase
    Implements IEnumerable, IEnumerator ', System.ICloneable

    'ENOP �ե�W��
    'Private m_sDLLName As String 
    'ENOP �R�W�Ŷ�
    'Private m_sNameSpase As String e

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
    ''' �s�� Column Name �P BosAttribute Object
    ''' </summary>
    ''' <remarks></remarks>
    Private m_attributes As New Hashtable

    ''' <summary>
    ''' �s�� BosAttribute Object
    ''' </summary>
    ''' <remarks></remarks>
    Private m_bosAttributeList As New BosAttributeList

    ''' <summary>
    ''' ����O�_���J
    ''' </summary>
    ''' <remarks></remarks>
    Private m_bIsLoaded As Boolean

    ''' <summary>
    ''' �s�� PrimaryKey's Column Name �P value
    ''' </summary>
    ''' <remarks></remarks>
    Private m_hashPrimaryKeys As New Hashtable

    ''' <summary>
    ''' �s�� PrimaryKey's Column Name
    ''' </summary>
    ''' <remarks></remarks>
    Protected Friend m_arrayPrimaryKeys As New ArrayList

    ''' <summary>
    ''' Current DataSet
    ''' </summary>
    ''' <remarks></remarks>
    Private m_dataSet As DataSet

    ''' <summary>
    ''' �ثe�O��������
    ''' </summary>
    ''' <remarks></remarks>
    Private m_currentIndex As Integer = -1

    ''' <summary>
    ''' �O��PK�O�_���J
    ''' </summary>
    ''' <remarks></remarks>
    Private m_bloadPk As Boolean

    Public sErrorMsg As String = ""

    ' Public MustOverride Function loadByPK() As Boolean

#Region "�غc�l"

    ''' <summary>
    ''' ��l�� BosBase ���O���s������� (Instance)�C
    ''' </summary>
    ''' <remarks>�غc�l
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/01/20	modify
    ''' </history>
    Sub New()
    End Sub

    ''' <summary>
    ''' ��l�� BosBase ���O���s������� (Instance)�C
    ''' </summary>
    ''' <param name="sTableName">�ǤJTable Name(�j�g)</param>
    ''' <param name="databaseManager">�ǤJ DatabaseManager</param>
    ''' <remarks>�غc�l
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/01/20	modify
    ''' </history>
    Sub New(ByVal sTableName As String, ByVal databaseManager As DatabaseManager)
        Me.setDatabaseManager(databaseManager)
        'If (databaseManager.isTransaction) Then
        '    setTransaction(databaseManager.getTransaction)
        'End If                
        'Me.setOperateOP()
        'ENOP �R�W�Ŷ�
        'm_sNameSpase = Me.GetType.Namespace
        ''ENOP �ե�W��
        'm_sDLLName = Me.GetType.Module.Name.ToUpper.Replace(".DLL", "")
        Me.initMetaData(sTableName)
        'Me.setPrimaryKeys()
    End Sub

    'Private Sub setOperateOP()
    '    Me.m_sDLLName = Me.GetType.Module.ScopeName.Split(".")(0)
    '    Me.m_sNameSpase = Me.GetType.Namespace
    'End Sub
#End Region

#Region "load SQL"

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
    Function loadBySQL(ByVal sSQL As String) As Boolean
        Dim bResult As Boolean = loadBySQL(sSQL, Nothing)
        Return bResult
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
    Protected Function loadData(ByVal ParamArray commandParameters() As IDataParameter) As Boolean
        Return loadBySQL(commandParameters)
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
    Function loadBySQL(ByVal ParamArray commandParameters() As IDataParameter) As Boolean

        Dim sSQL As String = " SELECT * FROM " + Me.getSchemaTableName()
        If (commandParameters.GetLength(0) > 0) Then
            sSQL &= " WHERE "
        End If
        For Each para As IDataParameter In commandParameters
            'for poc

            If Me.getDatabaseManager.getDataSource.IndexOf("IBM") = -1 Then
                sSQL &= para.ParameterName & "=" & ProviderFactory.PositionPara & "" & para.ParameterName & " AND "
            ElseIf Me.getDatabaseManager.getDataSource.IndexOf("IBM") <> -1 Then
                sSQL &= para.ParameterName & "=? AND "
            Else
                sSQL &= para.ParameterName & "=" & ProviderFactory.PositionPara & "" & para.ParameterName & " AND "
            End If

        Next
        sSQL = sSQL.Trim()
        sSQL = sSQL.Substring(0, sSQL.Length() - 4)
        Return loadBySQL(sSQL, commandParameters)
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
    Function loadBySQL(ByVal sSQL As String, ByVal ParamArray commandParameters() As IDataParameter) As Boolean
        Dim bResult As Boolean = False
        setIsLoaded(bResult)

        Try
            Dim ds As DataSet
            ds = Me.ExecuteDataset(CommandType.Text, sSQL, commandParameters)
            If Not IsNothing(ds) Then
                For Each row As DataRow In ds.Tables(0).Rows
                    bResult = constructData(row)
                Next
                setIsLoaded(bResult)
                Return bResult
            End If
            Return bResult
        Catch ex As Exception
            Throw New BosException(Me.getSchemaTableName & ":(Exception)", ex, CType(ex.TargetSite, System.Reflection.MethodInfo))
        End Try
    End Function

    Protected Function loadBySQLScalar(ByVal ParamArray commandParameters() As IDataParameter) As Object

        Dim sSQL As String = " SELECT * FROM " + Me.getSchemaTableName()
        If (commandParameters.GetLength(0) > 0) Then
            sSQL &= " WHERE "
        End If
        For Each para As IDataParameter In commandParameters
            'for poc

            If Me.getDatabaseManager.getDataSource.IndexOf("IBM") = -1 Then
                sSQL &= para.ParameterName & "=" & ProviderFactory.PositionPara & "" & para.ParameterName & " AND "
            ElseIf Me.getDatabaseManager.getDataSource.IndexOf("IBM") <> -1 Then
                sSQL &= para.ParameterName & "=? AND "
            Else
                sSQL &= para.ParameterName & "=" & ProviderFactory.PositionPara & "" & para.ParameterName & " AND "
            End If

        Next
        sSQL = sSQL.Trim()
        sSQL = sSQL.Substring(0, sSQL.Length() - 4)
        Return loadBySQLScalar(sSQL, commandParameters)
    End Function

    Protected Function loadBySQLScalar(ByVal sSQL As String) As Object
        Dim obj As Object = loadBySQLScalar(sSQL, Nothing)
        Return obj
    End Function

    Protected Function loadBySQLScalar(ByVal sSQL As String, ByVal ParamArray commandParameters() As IDataParameter) As Object
        Try
            Dim obj As Object
            obj = Me.ExecuteScalar(CommandType.Text, sSQL, commandParameters)

            Return obj
        Catch ex As Exception
            Throw New BosException(Me.getSchemaTableName & ":(Exception)", ex, CType(ex.TargetSite, System.Reflection.MethodInfo))
        End Try
    End Function

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
    Friend Sub setTableName(ByVal sTableName As String)
        Me.m_sTableName = sTableName
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
        Return Me.m_sTableName
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
    Private Function getSchemaTableName() As String
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
    ''' ���oCurrent DataSet�C
    ''' </summary>
    ''' <returns>DataSet �Ǧ^���� Current DataSet</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/01/20	modify
    ''' </history>
    Public Function getDataSet() As DataSet
        Return m_dataSet
    End Function

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
        m_dataSet = dataSet
    End Sub

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
    Friend Sub setDatabaseManager(ByVal databaseManager As DatabaseManager)
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
    Function getDatabaseManager() As DatabaseManager
        Return Me.m_databaseManager
    End Function
#End Region

#Region "PrimaryKey"
    ''' <summary>
    ''' �]�w PrimaryKeys
    ''' </summary>
    ''' <remarks> 
    ''' �p�GTable�S���]�w���A�@�w�ݭnOverridable�o��Function�A
    ''' �_�h������L�@��
    ''' </remarks>
    Public Overridable Sub setPrimaryKeys()
    End Sub

    ''' <summary>
    ''' get PrimaryKeys
    ''' </summary>
    ''' <returns>ArrayList �^��PrimaryKey's Column Name </returns>
    ''' <remarks></remarks>
    Public Function getArrayPrimaryKeys() As ArrayList
        Return Me.m_arrayPrimaryKeys
    End Function

    ''' <summary>
    ''' get PrimaryKeys
    ''' </summary>
    ''' <returns>Hashtable �^��PrimaryKey's key and Value </returns>
    ''' <remarks></remarks>
    Public Function getHashPrimaryKeys() As Hashtable
        Return Me.m_hashPrimaryKeys
    End Function
    ''' <summary>
    ''' set Hash PrimaryKeys
    ''' </summary>
    ''' <param name="hashtable">�ǤJPrimaryKey's key and Value (Hashtable)</param>
    ''' <remarks></remarks>
    Protected Sub setHashPrimaryKeys(ByVal hashtable As Hashtable)
        Me.m_hashPrimaryKeys = hashtable
    End Sub
#End Region

#Region "IsLoaded"
    ''' <summary>
    ''' �]�w����O�_���J
    ''' </summary>
    ''' <param name="bIsLoaded"></param>
    ''' <remarks></remarks>
    Public Sub setIsLoaded(ByVal bIsLoaded As Boolean)
        m_bIsLoaded = bIsLoaded
    End Sub
    ''' <summary>
    ''' ���o����O�_���J
    ''' </summary>
    ''' <returns>Boolean True:���J�AFalse:�L���J</returns>
    ''' <remarks></remarks>
    Public Function isLoaded() As Boolean
        Return m_bIsLoaded
    End Function
#End Region

#Region "get/set Value"

#Region "Attribute"

    ''' <summary>
    ''' set value of Column
    ''' </summary>
    ''' <param name="sAttrName">String �ǤJColumn Name</param>
    ''' <param name="value">Object �ǤJColumn's value</param>
    ''' <remarks></remarks>
    Sub setAttribute(ByVal sAttrName As String, ByVal value As Object)
        sAttrName = sAttrName.Trim
        Dim attr As BosAttribute = CType(m_attributes.Item(sAttrName), BosAttribute)

        If Not IsNothing(attr) Then
            'If Not IsNothing(value) AndAlso Not IsDBNull(value) Then
            '    If attr.getDataType = System.Data.DbType.String Then
            '        value = Titan.Utility.transferUnicodeToNCR(value.ToString)
            '    End If
            'End If
            attr.setValue(value, True)
        End If
    End Sub

    ''' <summary>
    ''' get value of Column
    ''' </summary>
    ''' <param name="sAttrName">String �ǤJColumn Name</param>
    ''' <returns>Object Column's Value</returns>
    ''' <remarks>�^�Ǹ�Ʈw���O</remarks>
    Function getAttribute(ByVal sAttrName As String) As Object
        Dim attr As BosAttribute = CType(m_attributes.Item(sAttrName), BosAttribute)

        If Not IsNothing(attr) Then
            Dim value As Object = attr.getValue()
            If Not IsNothing(value) AndAlso Not IsDBNull(value) Then
                If attr.getDataType = System.Data.DbType.String Then
                    Return Titan.Utility.transferNCRToUnicode(value.ToString) 'coding.transferNCRToUnicode(value)
                End If
            End If
            Return value
        Else
            Return Nothing
        End If
    End Function

#End Region

#Region "Long"
    ''' <summary>
    ''' set Long value of Column
    ''' </summary>
    ''' <param name="sAttrName">String �ǤJColumn Name</param>
    ''' <param name="iAttrValue">Long �ǤJColumn's value</param>
    ''' <remarks>
    ''' OleDb.OleDbType.BigInt 
    ''' </remarks>
    Sub setLong(ByVal sAttrName As String, ByVal iAttrValue As Long)
        setAttribute(sAttrName, CType(iAttrValue, Long))
    End Sub

    ''' <summary>
    ''' get Long value of Column
    ''' </summary>
    ''' <param name="sAttrName">String �ǤJColumn Name</param>
    ''' <returns>Long Column's Value</returns>
    ''' <remarks>
    ''' OleDb.OleDbType.BigInt
    ''' </remarks>
    Function getLong(ByVal sAttrName As String) As Long
        Return DbUtility.getField(CType(getAttribute(sAttrName), Long))
    End Function
#End Region

#Region "String"

    ''' <summary>
    ''' set String value of Column
    ''' </summary>
    ''' <param name="sAttrName">String �ǤJColumn Name</param>
    ''' <param name="sAttrValue">String �ǤJColumn's value</param>
    ''' <remarks>
    ''' OleDb.OleDbType.BSTR
    ''' OleDb.OleDbType.Char
    ''' OleDb.OleDbType.LongVarChar
    ''' OleDb.OleDbType.LongVarWChar
    ''' OleDb.OleDbType.VarChar
    ''' OleDb.OleDbType.VarWChar
    ''' OleDb.OleDbType.WChar
    ''' </remarks>
    Sub setString(ByVal sAttrName As String, ByVal sAttrValue As String)
        If Titan.Utility.isValidateData(sAttrValue) Then
            sAttrValue = sAttrValue.Trim
        End If
        setAttribute(sAttrName, CType(sAttrValue, String))
    End Sub
    ''' <summary>
    ''' set String value of Column
    ''' </summary>
    ''' <param name="sAttrName">String �ǤJColumn Name</param>
    ''' <returns>String Column's Value</returns>
    ''' <remarks>
    ''' OleDb.OleDbType.BSTR
    ''' OleDb.OleDbType.Char
    ''' OleDb.OleDbType.LongVarChar
    ''' OleDb.OleDbType.LongVarWChar
    ''' OleDb.OleDbType.VarChar
    ''' OleDb.OleDbType.VarWChar
    ''' OleDb.OleDbType.WChar
    ''' </remarks>
    Function getString(ByVal sAttrName As String) As String

        Return DbUtility.getField(CType(getAttribute(sAttrName), String)).Trim
    End Function

#End Region

#Region "Int"
    ''' <summary>
    ''' set Integer value of Column
    ''' </summary>
    ''' <param name="sAttrName">String �ǤJColumn Name</param>
    ''' <param name="iAttrValue">Integer �ǤJColumn's value</param>
    ''' <remarks>
    ''' OleDb.OleDbType.Integer  
    ''' </remarks>
    Sub setInt(ByVal sAttrName As String, ByVal iAttrValue As Integer)
        setAttribute(sAttrName, CType(iAttrValue, Integer))
    End Sub

    ''' <summary>
    ''' get Integer value of Column
    ''' </summary>
    ''' <param name="sAttrName">String �ǤJColumn Name</param>
    ''' <returns>Integer Column's Value</returns>
    ''' <remarks>
    ''' OleDb.OleDbType.Integer  
    ''' </remarks>
    Function getInt(ByVal sAttrName As String) As Integer
        Dim str As String = getString(sAttrName)
        If str.Equals("") Then
            str = "0"
        End If
        Return DbUtility.getField(CType(str, Integer))
    End Function
#End Region

#Region "Decimal"

    ''' <summary>
    ''' set Decimal value of Column
    ''' </summary>
    ''' <param name="sAttrName">String �ǤJColumn Name</param>
    ''' <param name="decAttrValue">Decimal �ǤJColumn's value</param>
    ''' <remarks>
    ''' OleDb.OleDbType.Numeric 
    ''' OleDb.OleDbType.Decimal
    ''' </remarks>
    Sub setDecimal(ByVal sAttrName As String, ByVal decAttrValue As Decimal)
        setAttribute(sAttrName, CType(decAttrValue, Decimal))
    End Sub
    ''' <summary>
    ''' get Decimal value of Column
    ''' </summary>
    ''' <param name="sAttrName">String �ǤJColumn Name</param>
    ''' <returns>Decimal Column's Value</returns>
    ''' <remarks>
    ''' OleDb.OleDbType.Numeric 
    ''' OleDb.OleDbType.Decimal
    ''' </remarks>
    Function getDecimal(ByVal sAttrName As String) As Decimal
        Return DbUtility.getField(CType(getAttribute(sAttrName), Decimal))
    End Function
#End Region

#Region "Date"

    ''' <summary>
    ''' set Date value of Column
    ''' </summary>
    ''' <param name="sAttrName">String �ǤJColumn Name</param>
    ''' <param name="dAttrValue">DateTime �ǤJColumn's value</param>
    ''' <remarks>
    ''' OleDb.OleDbType.Date
    ''' OleDb.OleDbType.DBDate
    ''' OleDb.OleDbType.DBTimeStamp
    ''' </remarks>
    Sub setDateTime(ByVal sAttrName As String, ByVal dAttrValue As DateTime)
        setAttribute(sAttrName, CType(dAttrValue, DateTime))
    End Sub
    ''' <summary>
    ''' get Date value of Column
    ''' </summary>
    ''' <param name="sAttrName">String �ǤJColumn Name</param>
    ''' <returns>Date Column's Value</returns>
    ''' <remarks>
    ''' OleDb.OleDbType.Date
    ''' OleDb.OleDbType.DBDate
    ''' OleDb.OleDbType.DBTimeStamp
    ''' </remarks>
    Function getDateTime(ByVal sAttrName As String) As DateTime
        If Not IsNothing(getAttribute(sAttrName)) Then
            Return DbUtility.getField(CType(getAttribute(sAttrName), DateTime))
        End If
        Return Nothing
    End Function
#End Region

#Region "SmallInt"

    ''' <summary>
    ''' set Int16 value of Column
    ''' </summary>
    ''' <param name="sAttrName">String �ǤJColumn Name</param>
    ''' <param name="iAttrValue">Int16 �ǤJColumn's value</param>
    ''' <remarks>
    ''' OleDb.OleDbType.SmallInt
    ''' </remarks>
    Sub setSmallInt(ByVal sAttrName As String, ByVal iAttrValue As Int16)
        setAttribute(sAttrName, CType(iAttrValue, Int16))
    End Sub

    ''' <summary>
    ''' get Int16 value of Column
    ''' </summary>
    ''' <param name="sAttrName">String �ǤJColumn Name</param>
    ''' <returns>Int16 Column's Value</returns>
    ''' <remarks>
    ''' OleDb.OleDbType.SmallInt
    ''' </remarks>
    Function getSmallInt(ByVal sAttrName As String) As Int16
        Return DbUtility.getField(CType(getAttribute(sAttrName), Int16))
    End Function
#End Region

#End Region

#Region "BosAttribut"

    Public Function getBosAttributes() As BosAttributeList
        Return m_bosAttributeList
    End Function
#End Region

#Region "���ѫD�� BosBase ���󤧥[������ݩʭ�"

    ''' <summary>
    ''' ���ѫD�� BosBase ���󤧥[������ݩʭ�
    ''' </summary>
    ''' <param name="sColName">�ǤJColumn Name</param>
    ''' <param name="DBType">�ǤJSystem.Data.DbType</param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/01/20	modify
    ''' </history>
    Public Function setAttMeta(ByVal sColName As String, ByVal DBType As Data.DbType) As BosAttribute
        Dim attribute As New BosAttribute
        Dim attrMeta As New BosAttrMeta
        attrMeta.setColName(sColName)
        attrMeta.setDataType(DBType)
        attribute.setAttrMeta(attrMeta)
        Return attribute
    End Function

    ''' <summary>
    ''' ���ѫD�� BosBase���󤧥[������ݩʭ�
    ''' </summary>
    ''' <param name="attr">�ǤJBosAttribute</param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/01/20	modify
    ''' </history>
    Protected Friend Sub addAttribute(ByVal attr As BosAttribute)
        Dim sColName As String = attr.getColName()
        If Not m_attributes.ContainsKey(sColName.ToUpper) Then
            m_attributes.Add(sColName.ToUpper, attr)
            m_bosAttributeList.add(attr)
        Else
            m_attributes.Item(sColName.ToUpper) = attr
            m_bosAttributeList.add(attr)
        End If
    End Sub

    ''' <summary>
    ''' ���ѫD�� BosBase ���󤧥[������ݩʭ�
    ''' </summary>
    ''' <param name="sColName">�ǤJColumn Name</param>
    ''' <param name="DBType">�ǤJSystem.Data.DbType</param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/01/20	modify
    ''' </history> 
    Public Sub addAttribute(ByVal sColName As String, ByVal DBType As System.Data.DbType)
        Me.addAttribute(setAttMeta(sColName, DBType))
    End Sub
#End Region

#End Region

#Region "inner function"

#Region "initMetaData"
    ''' <summary>
    ''' ��l��MetaData
    ''' </summary>
    ''' <param name="sTableName">�ǤJ Table Name</param>
    ''' <returns>Boolean MetaData�O�_���J</returns>
    ''' <remarks>
    ''' �p�G���w������DBMeta.dll�A�h�qDBMeta.dll���J
    ''' �p�G�S���h�����qDataBase���o
    ''' </remarks>
    Private Function initMetaData(ByVal sTableName As String) As Boolean
        Dim bResult As Boolean = False
        Dim objClass As Object

        setTableName(sTableName)
        Try
            objClass = ClassLoad.ClassforName("DBMeta", "Com.Azion.VB.DBMeta", getTableName() & "_dbMeta")
            'load �۹����� MetaData
            bResult = loadAttrMeta(CType(objClass, DBMetaData))
            'bResult = loadAttrMeta(sTableName)
        Catch ex As Exception
            'Logger.log.Error(ex)
            'load �۹����� MetaData
            bResult = loadAttrMeta(sTableName)
            'Throw
        End Try
        Return bResult
    End Function

    ''' <summary>
    ''' ��l��MetaData
    ''' </summary>
    ''' <param name="dBMetaData">�ǤJ DBMetaData</param>
    ''' <returns>Boolean MetaData�O�_���J</returns>
    ''' <remarks>
    ''' ��DBMeta.dll���JMeta Data
    ''' </remarks>
    Private Function loadAttrMeta(ByVal dBMetaData As DBMetaData) As Boolean
        Dim bResult As Boolean = False
        Try
            dBMetaData.init()
            If Not m_bloadPk Then
                dBMetaData.setPrimaryKeys()
                If (dBMetaData.getPrimaryKeys.Count > 0) Then
                    m_arrayPrimaryKeys = dBMetaData.getPrimaryKeys
                    m_bloadPk = True
                Else
                    Me.setPrimaryKeys()
                End If
            End If

            For Each hData As Hashtable In dBMetaData.getMetaArray
                ' attribute meta data (column name, type name)
                SyncLock hData.SyncRoot
                    Dim attrMeta As New BosAttrMeta
                    attrMeta.constructAttrMeta(hData)

                    Dim attr As New BosAttribute
                    attr.setAttrMeta(attrMeta)
                    'add attribute to bos
                    addAttribute(attr)
                End SyncLock
            Next
            bResult = True
        Catch ex As Exception
            Throw New BosException(Me.getSchemaTableName & ":(Exception)", ex, CType(ex.TargetSite, System.Reflection.MethodInfo))
        End Try
        Return bResult
    End Function

    ''' <summary>
    ''' ��l��MetaData
    ''' </summary>
    ''' <param name="sTableName">�ǤJ Table Name</param>
    ''' <returns>Boolean MetaData�O�_���J</returns>
    ''' <remarks>
    ''' ������DataBase���o
    ''' </remarks>
    Private Function loadAttrMeta(ByVal sTableName As String) As Boolean
        Dim schemaTable As DataTable = Nothing
        Dim bResult As Boolean = False

        Try
            'schemaTable = DbUtility.loadSchemaInfo(Me.m_databaseManager, sTableName)
            schemaTable = DbUtility.getTable(sTableName, Me.m_databaseManager)
            Dim colkeys As ArrayList = DbUtility.getPrimaryKey(sTableName, m_databaseManager)

            m_arrayPrimaryKeys.AddRange(colkeys)

            If (m_arrayPrimaryKeys.Count = 0) Then
                Me.setPrimaryKeys()
            End If


            For i As Integer = 0 To schemaTable.Rows.Count - 1
                Dim attrMeta As New BosAttrMeta
                attrMeta.constructAttrMeta(schemaTable.Rows(i))

                Dim attr As New BosAttribute
                attr.setAttrMeta(attrMeta)
                'add attribute to bos
                addAttribute(attr)
            Next i

            ''���o��쪺��T
            'For Each column As DataColumn In schemaTable.Columns
            '    Dim attrMeta As New BosAttrMeta
            '    attrMeta.constructAttrMeta(column)

            '    Dim attr As New BosAttribute
            '    attr.setAttrMeta(attrMeta)
            '    'add attribute to bos
            '    addAttribute(attr)
            'Next


            bResult = True
        Catch ex As Exception
            Throw New BosException(Me.getSchemaTableName & ":(Exception)", ex, CType(ex.TargetSite, System.Reflection.MethodInfo))
        Finally
            If (Not schemaTable Is Nothing) Then schemaTable.Dispose()
        End Try
        Return bResult
    End Function
#End Region

#Region "Execute"

#Region "ExecuteDataset"
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
                m_dataSet = DBObject.ExecuteDataset(Me.getDatabaseManager.getTransaction, commandType, commandText, commandParameters)
                Return m_dataSet
            ElseIf Not IsNothing(Me.getDatabaseManager.getConnection) Then
                m_dataSet = DBObject.ExecuteDataset(Me.getDatabaseManager.getConnection, commandType, commandText, commandParameters)
                Return m_dataSet
            Else
                Return Nothing
            End If
        Catch ex As Exception
            Throw New BosException(Me.getSchemaTableName & ":(Exception)", ex, CType(ex.TargetSite, System.Reflection.MethodInfo))
        End Try
    End Function
#End Region

#Region "ExecuteNonQuery"
    ''' <summary>
    ''' ExecuteNonQuery
    ''' </summary>
    ''' <param name="commandType">�ǤJCommandType</param>
    ''' <param name="commandText">�ǤJcommandText</param>
    ''' <param name="commandParameters">�ǤJIDataParameter()</param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/01/20	modify
    ''' </history>
    Private Function ExecuteNonQuery(ByVal commandType As CommandType, _
                                ByVal commandText As String, _
                                ByVal ParamArray commandParameters() As IDataParameter) As Integer
        Try
            If Me.m_databaseManager.isTransaction AndAlso Not Me.m_databaseManager.isCloseTransaction Then
                Return DBObject.ExecuteNonQuery(Me.getDatabaseManager.getTransaction, commandType, commandText, commandParameters)
            Else
                Return DBObject.ExecuteNonQuery(Me.getDatabaseManager.getConnection(), commandType, commandText, commandParameters)
            End If
        Catch ex As Exception
            Throw New BosException(Me.getSchemaTableName & ":(Exception)", ex, CType(ex.TargetSite, System.Reflection.MethodInfo))
        End Try
    End Function
#End Region

#Region "ExecuteScalar"
    ''' <summary>
    ''' ExecuteScalar
    ''' </summary>
    ''' <param name="commandType">�ǤJCommandType</param>
    ''' <param name="commandText">�ǤJcommandText</param>
    ''' <param name="commandParameters">�ǤJIDataParameter()</param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/01/20	modify
    ''' </history>
    Private Function ExecuteScalar(ByVal commandType As CommandType, _
                                      ByVal commandText As String, _
                                      ByVal ParamArray commandParameters() As IDataParameter) As Object
        Try
            If Me.m_databaseManager.isTransaction AndAlso Not Me.m_databaseManager.isCloseTransaction Then
                Return DBObject.ExecuteScalar(Me.getDatabaseManager.getTransaction, commandType, commandText, commandParameters)
            Else
                Return DBObject.ExecuteScalar(Me.getDatabaseManager.getConnection(), commandType, commandText, commandParameters)
                Return Nothing
            End If
        Catch ex As Exception
            Throw New BosException(Me.getSchemaTableName & ":(Exception)", ex, CType(ex.TargetSite, System.Reflection.MethodInfo))
        End Try
    End Function
#End Region

#Region "ExecuteReader"
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
    Private Function ExecuteReader(ByVal commandType As CommandType, _
                                    ByVal commandText As String, _
                                    ByVal ParamArray commandParameters() As IDataParameter) As Object
        Try
            If Me.m_databaseManager.isTransaction AndAlso Not Me.m_databaseManager.isCloseTransaction Then
                Return DBObject.ExecuteReader(Me.getDatabaseManager.getTransaction(), commandType, commandText, commandParameters)
            Else
                Return DBObject.ExecuteReader(Me.getDatabaseManager.getConnection(), commandType, commandText, commandParameters)
            End If
        Catch ex As Exception
            Throw New BosException(Me.getSchemaTableName & ":(Exception)", ex, CType(ex.TargetSite, System.Reflection.MethodInfo))
        End Try
    End Function
#End Region

#End Region

#Region "toString"

    Private Function toStringPK() As String
        Try
            Dim sb As New System.Text.StringBuilder("<PrimaryKeys>")

            If Not IsNothing(m_arrayPrimaryKeys) Then
                For Each sKey As String In Me.m_hashPrimaryKeys.Keys
                    sb.Append("<PrimaryKey>" & sKey & "</PrimaryKey>")
                    sb.Append("<Value>" & Me.m_hashPrimaryKeys.Item(sKey).ToString & "</Value>")
                Next
            End If
            sb.Append("</PrimaryKeys>")
            Return sb.ToString()
        Catch ex As Exception
        End Try
        Return ""
    End Function

    Public Overrides Function toString() As String
        Dim sTag As String = Me.getSchemaTableName()
        Dim sb As New System.Text.StringBuilder
        sb.Append("<" + sTag + ">")
        Me.toStringPK()
        For i As Integer = 0 To m_bosAttributeList.size - 1
            Dim bosAttribute As BosAttribute = m_bosAttributeList.item(i)
            sb.Append(bosAttribute.toString())
        Next

        sb.Append("</" + sTag + ">")

        Return sb.ToString()
    End Function
#End Region

#Region "�M������"

    Sub clear()

        Dim iEnum As IEnumerator = m_attributes.Values.GetEnumerator

        While iEnum.MoveNext
            Dim attr As BosAttribute = CType(iEnum.Current, BosAttribute)
            attr.clearValue()
        End While

        setIsLoaded(False)
    End Sub
#End Region
   
#End Region

#Region "construct"

    Protected Friend Function constructIRDData(ByVal IReader As System.Data.IDataReader, ByVal objDataColumnCollection As System.Data.DataColumnCollection) As Boolean
        Dim bResult As Boolean = False
        setIsLoaded(bResult)

        Try

            For i As Integer = 0 To m_bosAttributeList.size - 1
                Dim bosAttribute As bosAttribute = m_bosAttributeList.item(i)
                Dim value As Object = Nothing
                Dim sColumnName As String = bosAttribute.getColName

                If objDataColumnCollection.Contains(sColumnName) Then
                    Try
                        value = IReader(sColumnName)
                    Catch ex As Exception
                        Throw New BosException(Me.getSchemaTableName & ":(Exception)", ex, CType(ex.TargetSite, System.Reflection.MethodInfo))
                    End Try

                    CType(m_attributes.Item(sColumnName), bosAttribute).setValue(value)

                    If Not IsNothing(m_arrayPrimaryKeys) Then
                        If m_arrayPrimaryKeys.Contains(sColumnName) Then
                            If m_hashPrimaryKeys.ContainsKey(sColumnName) Then
                                m_hashPrimaryKeys.Item(sColumnName) = value
                            Else
                                m_hashPrimaryKeys.Add(sColumnName, value)
                            End If
                        End If
                    End If
                End If
            Next
            bResult = True
            setIsLoaded(bResult)
            Return bResult
        Catch ex As Exception
            Throw New BosException(Me.getSchemaTableName & ":(Exception)", ex, CType(ex.TargetSite, System.Reflection.MethodInfo))
        End Try
    End Function

    Public Function constructData(ByVal objDataRow As DataRow) As Boolean
        Dim bResult As Boolean = False
        setIsLoaded(bResult)

        Try
             
            For Each column As DataColumn In objDataRow.Table.Columns
                Dim sColName As String = column.ColumnName.ToUpper
                Dim value As Object = objDataRow.Item(sColName)
                If m_attributes.ContainsKey(sColName) Then
                    CType(m_attributes.Item(sColName), BosAttribute).setValue(value)

                    If Not IsNothing(m_arrayPrimaryKeys) Then
                        If m_arrayPrimaryKeys.Contains(sColName) Then
                            If m_hashPrimaryKeys.ContainsKey(sColName) Then
                                m_hashPrimaryKeys.Item(sColName) = value
                            Else
                                m_hashPrimaryKeys.Add(sColName, value)
                            End If
                        End If
                    End If
                End If
            Next

            'For i As Integer = 0 To m_bosAttributeList.size - 1

            '    Dim bosAttribute As BosAttribute = m_bosAttributeList.item(i)

            '    Dim value As Object = Nothing
            '    Dim sColName As String = bosAttribute.getColName()

            '    If objDataRow.Table.Columns.Contains(sColName) Then

            '        value = objDataRow.Item(sColName)
            '        CType(m_attributes.Item(sColName), BosAttribute).setValue(value)

            '        If Not IsNothing(m_arrayPrimaryKeys) Then
            '            If m_arrayPrimaryKeys.Contains(sColName) Then
            '                If m_hashPrimaryKeys.ContainsKey(sColName) Then
            '                    m_hashPrimaryKeys.Item(sColName) = value
            '                Else
            '                    m_hashPrimaryKeys.Add(bosAttribute.getColName, value)
            '                End If
            '            End If

            '        End If
            '    End If
            'Next
            bResult = True
            setIsLoaded(bResult)
            Return bResult
        Catch ex As Exception
            Throw New BosException(Me.getSchemaTableName & ":(Exception)", ex, CType(ex.TargetSite, System.Reflection.MethodInfo))
        End Try
    End Function

#End Region

#Region "generate SQL"

    Function save() As Integer
        If (m_arrayPrimaryKeys.Count = 0) Then
            Throw New BosException(Me.getSchemaTableName & ": �г]�w PrimaryKey")
        End If
        If (m_attributes.Keys.Count() <= 0) Then
            Return 0
        End If

        Dim iEffect As Integer = 0
        Try

            ' �S���]�wtransaction �A�htransaction�Ѧ��\��]�w
            'If Not (Me.m_databaseManager.isTransaction) Then
            '    '��DB�ۤv��
            '    'Me.getDatabaseManager.beginTran()
            '    'Me.setTransaction(Me.m_databaseManager.getTransaction)          
            'End If

            If (isLoaded()) Then
                iEffect = updateSQL()
            Else
                iEffect = insertSQL()
            End If

            'If Not (Me.m_databaseManager.isTransaction) Then
            '    ' Me.m_databaseManager.commit()
            'End If
        Catch ex As Exception
            'If Not (Me.m_databaseManager.isTransaction) Then
            '    ' Me.m_databaseManager.Rollback()
            'End If
            Throw New BosException(Me.getSchemaTableName & ":(Exception)", ex, CType(ex.TargetSite, System.Reflection.MethodInfo))
        End Try
        Return iEffect
    End Function

    Private Function insertSQL() As Integer

        Dim sbValueSQL As New System.Text.StringBuilder
        Dim sbParam As New System.Text.StringBuilder

        Dim valueParameters(m_attributes.Count - 1) As IDataParameter
        Dim iParaCount As Integer = 0

        Try
            For i As Integer = 0 To m_bosAttributeList.size - 1
                Dim bosAttribute As BosAttribute = m_bosAttributeList.item(i)
                Dim oDbType As System.Data.DbType = bosAttribute.getDataType
                Dim value As Object = bosAttribute.getValue()

                If bosAttribute.isSeted Then
                    Dim sColName As String = bosAttribute.getColName()
                    'Dim para As IDataParameter = ProviderFactory.CreateDataParameter(sColName, IIf(IsNothing(bosAttribute.getValue()), Convert.DBNull, bosAttribute.getValue()))
                    Dim para As IDataParameter = ProviderFactory.CreateDataParameter()

                    If ProviderFactory.m_ProviderType = ProviderType.ODP Then
                        CType(para, Oracle.DataAccess.Client.OracleParameter).OracleDbType = DbUtility.getODPDBType(bosAttribute.getProviderType())
                    End If

                    valueParameters(iParaCount) = para
                    sbValueSQL.Append(sColName & " , ")

                    If Me.getDatabaseManager.getDataSource.IndexOf("IBM") = -1 Then
                        sbParam.Append(ProviderFactory.PositionPara).Append(sColName).Append("  , ")
                    ElseIf Me.getDatabaseManager.getDataSource.IndexOf("IBM") <> -1 Then
                        sbParam.Append("?  , ")
                    Else
                        sbParam.Append(ProviderFactory.PositionPara).Append(sColName).Append("  , ")
                    End If

                    If (oDbType = DbType.Decimal OrElse oDbType = DbType.Double OrElse oDbType = DbType.Int16 OrElse oDbType = DbType.Int32 OrElse oDbType = DbType.Int32 OrElse oDbType = DbType.Int64 OrElse oDbType = DbType.Single) AndAlso Not Microsoft.VisualBasic.IsNumeric(value) Then
                        value = Nothing
                    ElseIf (oDbType = DbType.DateTime OrElse oDbType = DbType.Date) AndAlso Not Microsoft.VisualBasic.IsDate(value) Then
                        value = Nothing
                    End If

                    para.ParameterName = sColName
                    para.Value = IIf(IsNothing(value), Convert.DBNull, value)

                    iParaCount += 1 
                End If
            Next

            sbValueSQL = sbValueSQL.Remove(sbValueSQL.Length() - 2, 2)
            sbParam = sbParam.Remove(sbParam.Length() - 2, 2)

            Dim sbSQL As New System.Text.StringBuilder
            sbSQL.Append(" INSERT INTO ").Append(Me.getSchemaTableName()).Append(" ( ").Append(sbValueSQL.ToString).Append(" ) ")
            sbSQL.Append(" VALUES ( ").Append(sbParam.ToString).Append(" ) ")

            'Dim insertParas(iParaCount) As IDataParameter
            'valueParameters.CopyTo(insertParas, )

            If ProviderFactory.m_bDebug Then
                Titan.Utility.Debug(showInsertSQL())
            End If

            Return Me.ExecuteNonQuery(CommandType.Text, sbSQL.ToString, valueParameters)

        Catch ex As Exception
            Dim sSQL As String = showInsertSQL()
            Titan.Utility.Debug(sSQL)
            Throw New BosException(Me.getSchemaTableName & ":(Exception)", ex, CType(ex.TargetSite, System.Reflection.MethodInfo))
        End Try
    End Function
    ''' <summary>
    ''' Show Insert SQL�C
    ''' </summary> 
    ''' <returns>String</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/02/25	modify
    '''</history>
    Function showInsertSQL() As String
        Dim sbValueSQL As New System.Text.StringBuilder
        Dim sbShowParam As New System.Text.StringBuilder
        Dim sbShowSQL As New System.Text.StringBuilder
        For i As Integer = 0 To m_bosAttributeList.size - 1
            Dim bosAttribute As BosAttribute = m_bosAttributeList.item(i)
            Dim value As Object = bosAttribute.getValue()
            Dim sColName As String = bosAttribute.getColName()
            sbValueSQL.Append(sColName & " , ")

            If IsNothing(value) Then
                sbShowParam.Append(bosAttribute.getValue()).Append(" NULL , ")
            ElseIf bosAttribute.getDataType = DbType.String OrElse bosAttribute.getDataType = DbType.AnsiString Then
                sbShowParam.Append("'").Append(bosAttribute.getValue()).Append("'  , ")
            Else
                sbShowParam.Append(bosAttribute.getValue()).Append("  , ")
            End If
        Next

        sbValueSQL = sbValueSQL.Remove(sbValueSQL.Length() - 2, 2)
        sbShowParam = sbShowParam.Remove(sbShowParam.Length() - 2, 2)

        sbShowSQL.Append(" INSERT INTO ").Append(Me.getSchemaTableName()).Append(" ( ").Append(sbValueSQL.ToString).Append(" ) ")
        sbShowSQL.Append(" VALUES ( ").Append(sbShowParam.ToString).Append(" ) ")

        Return sbShowSQL.ToString
    End Function

    Private Function updateSQL() As Integer
        'PK���Ѽ�
        Dim sbPkeySQL As New System.Text.StringBuilder
        Dim pKeyParameters(m_arrayPrimaryKeys.Count - 1) As IDataParameter
        Dim iKeyCount As Integer = 0
        '��L�DPK���Ѽ�
        Dim sbValueSQL As New System.Text.StringBuilder
        Dim valueParaCount As Integer = m_attributes.Count
        Dim valueParameters(valueParaCount + m_arrayPrimaryKeys.Count - 1) As IDataParameter '+ pKeyParaCount
        Dim iValueCount As Integer = 0

        Try

            For i As Integer = 0 To m_bosAttributeList.size - 1
                Dim bosAttribute As BosAttribute = m_bosAttributeList.item(i)
                Dim sColName As String = bosAttribute.getColName()
                Dim value As Object = bosAttribute.getValue()
                Dim oDbType As System.Data.DbType = bosAttribute.getDataType

              
                If m_arrayPrimaryKeys.Contains(sColName) Then
                    If Me.getDatabaseManager.getDataSource.IndexOf("IBM") = -1 Then
                        sbPkeySQL.Append(sColName).Append("=").Append(ProviderFactory.PositionPara).Append("key" & iKeyCount).Append(" AND ")
                    ElseIf Me.getDatabaseManager.getDataSource.IndexOf("IBM") <> -1 Then
                        sbPkeySQL.Append(sColName).Append("=? AND ")
                    Else
                        sbPkeySQL.Append(sColName).Append("=").Append(ProviderFactory.PositionPara).Append("key" & iKeyCount).Append(" AND ")
                    End If
                    '(sColName, IIf(IsNothing(bosAttribute.getValue()), Convert.DBNull, bosAttribute.getValue()))
                    Dim para As IDataParameter = ProviderFactory.CreateDataParameter()

                    If ProviderFactory.m_ProviderType = ProviderType.ODP Then
                        CType(para, Oracle.DataAccess.Client.OracleParameter).OracleDbType = DbUtility.getODPDBType(bosAttribute.getProviderType())
                    End If

                    para.ParameterName = "key" & iKeyCount
                    para.Value = Me.m_hashPrimaryKeys.Item(sColName)
                    'Dim sParaColName As String = ProviderFactory.PositionPara & sColName
                    pKeyParameters(iKeyCount) = para 'ProviderFactory.CreateDataParameter(sColName, oDbType, Me.m_hashPrimaryKeys.Item(sColName))
                    iKeyCount += 1
                End If


                If bosAttribute.isSeted Then
                    If Me.getDatabaseManager.getDataSource.IndexOf("IBM") = -1 Then
                        sbValueSQL.Append(sColName).Append("=").Append(ProviderFactory.PositionPara).Append("value" & iValueCount).Append(" , ")
                    ElseIf Me.getDatabaseManager.getDataSource.IndexOf("IBM") <> -1 Then
                        sbValueSQL.Append(sColName).Append("=? , ")
                    Else
                        sbValueSQL.Append(sColName).Append("=").Append(ProviderFactory.PositionPara).Append("value" & iValueCount).Append(" , ")
                    End If
                    'ProviderFactory.PositionPara &
                    If (oDbType = DbType.Decimal OrElse oDbType = DbType.Double OrElse oDbType = DbType.Int16 OrElse oDbType = DbType.Int32 OrElse oDbType = DbType.Int32 OrElse oDbType = DbType.Int64 OrElse oDbType = DbType.Single) AndAlso Not Microsoft.VisualBasic.IsNumeric(value) Then
                        value = Nothing
                    ElseIf (oDbType = DbType.DateTime OrElse oDbType = DbType.Date) AndAlso Not Microsoft.VisualBasic.IsDate(value) Then
                        value = Nothing
                    End If

                    Dim para As IDataParameter = ProviderFactory.CreateDataParameter()

                    If ProviderFactory.m_ProviderType = ProviderType.ODP Then
                        CType(para, Oracle.DataAccess.Client.OracleParameter).OracleDbType = DbUtility.getODPDBType(bosAttribute.getProviderType())
                    End If

                    para.ParameterName = "value" & iValueCount
                    para.Value = IIf(IsNothing(value), Convert.DBNull, value)

                    valueParameters(iValueCount) = para ' ProviderFactory.CreateDataParameter(sColName, oDbType, IIf(IsNothing(value), Convert.DBNull, value))
                    iValueCount += 1
                End If
            Next

            If sbValueSQL.Length() > 0 Then
                sbValueSQL = sbValueSQL.Remove(sbValueSQL.Length() - 2, 2)
                sbPkeySQL = sbPkeySQL.Remove(sbPkeySQL.Length() - 4, 4)

                pKeyParameters.CopyTo(valueParameters, valueParameters.GetLength(0) - m_arrayPrimaryKeys.Count)

                Dim sbSQL As New System.Text.StringBuilder
                sbSQL.Append(" UPDATE ").Append(Me.getSchemaTableName()).Append(" Set ").Append(sbValueSQL.ToString)
                sbSQL.Append(" WHERE ").Append(sbPkeySQL.ToString)

                If ProviderFactory.m_bDebug Then
                    Titan.Utility.Debug(showUpdateSQL())
                End If

                Return Me.ExecuteNonQuery(CommandType.Text, sbSQL.ToString, valueParameters)
            End If
        Catch ex As Exception
            Dim sSQL As String = showUpdateSQL()
            Titan.Utility.Debug(sSQL)
            Throw New BosException(Me.getSchemaTableName & ":(Exception)", ex, CType(ex.TargetSite, System.Reflection.MethodInfo))
        End Try

        Return 0
    End Function
    ''' <summary>
    ''' Show Update SQL�C
    ''' </summary> 
    ''' <returns>String</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Titan]	2009/02/25	modify
    '''</history>
    Function showUpdateSQL() As String
        Dim sbShowPkeySQL As New System.Text.StringBuilder
        Dim sbShowValueSQL As New System.Text.StringBuilder
        Dim sbShowSQL As New System.Text.StringBuilder

        For i As Integer = 0 To m_bosAttributeList.size - 1
            Dim bosAttribute As BosAttribute = m_bosAttributeList.item(i)
            Dim value As Object = bosAttribute.getValue()
            Dim sColName As String = bosAttribute.getColName()
            Dim oDbType As System.Data.DbType = bosAttribute.getDataType
            If (oDbType = DbType.Decimal OrElse oDbType = DbType.Double OrElse oDbType = DbType.Int16 OrElse oDbType = DbType.Int32 OrElse oDbType = DbType.Int32 OrElse oDbType = DbType.Int64 OrElse oDbType = DbType.Single) AndAlso Not Microsoft.VisualBasic.IsNumeric(value) Then
                value = Nothing
            ElseIf (oDbType = DbType.DateTime OrElse oDbType = DbType.Date) AndAlso Not Microsoft.VisualBasic.IsDate(value) Then
                value = Nothing
            End If
            'update where ���󪺭ȶ��q�쥻select where���Ȩ��X(m_hashPrimaryKeys)
            If m_arrayPrimaryKeys.Contains(sColName) Then
                If bosAttribute.getDataType = DbType.String OrElse bosAttribute.getDataType = DbType.AnsiString Then
                    sbShowPkeySQL.Append(sColName).Append(" = '").Append(Me.m_hashPrimaryKeys.Item(sColName)).Append("' AND ")
                Else
                    sbShowPkeySQL.Append(sColName).Append(" = ").Append(Me.m_hashPrimaryKeys.Item(sColName)).Append(" AND ")
                End If
            End If

            If bosAttribute.isSeted Then
                If IsNothing(value) Then
                    sbShowValueSQL.Append(sColName).Append(" = NULL ").Append(" , ")
                ElseIf bosAttribute.getDataType = DbType.String OrElse bosAttribute.getDataType = DbType.AnsiString Then
                    sbShowValueSQL.Append(sColName).Append(" = '").Append(value).Append("' , ")
                Else
                    sbShowValueSQL.Append(sColName).Append(" = ").Append(value).Append(" , ")
                End If
            End If
        Next
        sbShowValueSQL = sbShowValueSQL.Remove(sbShowValueSQL.Length() - 2, 2)
        sbShowPkeySQL = sbShowPkeySQL.Remove(sbShowPkeySQL.Length() - 4, 4)

        sbShowSQL.Append(" UPDATE ").Append(Me.getSchemaTableName()).Append(" Set ").Append(sbShowValueSQL.ToString)
        sbShowSQL.Append(" WHERE ").Append(sbShowPkeySQL.ToString)

        Return sbShowSQL.ToString
    End Function

    Public Function remove() As Boolean
        If (m_arrayPrimaryKeys.Count = 0) Then
            Throw New BosException(Me.getSchemaTableName() & ": �г]�w PrimaryKey")
        End If
        Dim bResult As Boolean = False
        '�S�����J�����@�R�����\
        If (Not isLoaded()) Then
            Return True
        End If

        Dim sb As New System.Text.StringBuilder
        Dim i As Integer = 0
        Dim iPkeyCount As Integer = m_arrayPrimaryKeys.Count
        Dim commandParameters(iPkeyCount - 1) As IDataParameter
        Try
            For Each strField As String In m_attributes.Keys
                Dim strColName As String = strField
                Dim attr As BosAttribute = CType(m_attributes.Item(strColName), BosAttribute)

                For Each pKey As String In m_arrayPrimaryKeys
                    If (pKey.ToUpper.Equals(strColName)) Then
                        Dim sColName As String
                        If Me.getDatabaseManager.getDataSource.IndexOf("IBM") = -1 Then
                            sb.Append(attr.getColName() & "= " & ProviderFactory.PositionPara & attr.getColName() & " AND ")
                            sColName = ProviderFactory.PositionPara & attr.getColName()
                        ElseIf Me.getDatabaseManager.getDataSource.IndexOf("IBM") <> -1 Then
                            sb.Append(attr.getColName() & "=? AND ")
                            sColName = attr.getColName()
                        Else
                            sb.Append(attr.getColName() & "= " & ProviderFactory.PositionPara & attr.getColName() & " AND ")
                            sColName = ProviderFactory.PositionPara & attr.getColName()
                        End If
                        commandParameters(i) = ProviderFactory.CreateDataParameter(sColName, attr.getValue())
                        iPkeyCount -= 1
                        i += 1
                        Exit For
                    End If
                Next
                If (iPkeyCount = 0) Then
                    Exit For
                End If
            Next
            sb = sb.Remove(sb.ToString().Trim().Length() - 4, 4)

            Dim sbSQL As New System.Text.StringBuilder(" DELETE FROM ")

            sbSQL.Append(getSchemaTableName())
            sbSQL.Append(" WHERE ")
            sbSQL.Append(sb)

            If (Me.ExecuteNonQuery(CommandType.Text, sbSQL.ToString, commandParameters) > 0) Then
                bResult = True
                Me.setIsLoaded(False)
                Me.clear()
                Return bResult
            End If
            Return bResult
        Catch ex As Exception
            Throw New BosException(Me.getSchemaTableName & ":(Exception)", ex, CType(ex.TargetSite, System.Reflection.MethodInfo))
        End Try
    End Function
#End Region

#Region "��@IEnumerator�BIEnumerable ����"
    Function GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return CType(Me, IEnumerator)
    End Function

    Sub Reset() Implements IEnumerator.Reset
        m_currentIndex = -1
    End Sub

    Function MoveNext() As Boolean Implements IEnumerator.MoveNext
        If m_currentIndex < Me.m_attributes.Count - 1 Then
            m_currentIndex = m_currentIndex + 1
            Return True
        Else
            Return False
        End If
    End Function

    ReadOnly Property Current() As Object Implements IEnumerator.Current
        Get
            Return Me.m_attributes(m_currentIndex)
        End Get
    End Property

    Private Function Clone() As Object 'Implements System.ICloneable.Clone

        '�����ڭ̫إ߫��w�������@�ӹ�� 
        Dim newObject As BosBase = CType(Activator.CreateInstance(Me.GetType()), BosBase)
        '�ڭ̨��o�s��������Ҫ����}�C�C
        Dim fields() As System.Reflection.FieldInfo = newObject.GetType().GetFields()
        newObject.GetType().GetMembers()
        Dim i As Integer = 0

        For Each fi As System.Reflection.FieldInfo In fields
            '�ڭ̧P�_���O�_�䴩ICloneable�����C
            Dim ICloneType As System.Type = fi.FieldType.GetInterface("ICloneable", True)
            If Not IsNothing(ICloneType) Then
                '���o����Icloneable�����C
                Dim IClone As ICloneable = CType(fi.GetValue(Me), ICloneable)
                '�ϥΧJ����k�����]�w�s�ȡC
                fields(i).SetValue(newObject, IClone.Clone())
            Else
                ' �p�G����쳡�䴩Icloneable�����A�����]�m�Y�i�C
                fields(i).SetValue(newObject, fi.GetValue(Me))
            End If

            '�ˬd�Ӫ���O�_�䴩IEnumerable�����A�p�G�䴩�A 
            '�T�|��Ҧ������ˬd�L�̬O�_�䴩IList �� IDictionary �����C            
            Dim IEnumerableType As System.Type = fi.FieldType.GetInterface("IEnumerable", True)
            If Not IsNothing(IEnumerableType) Then
                ' ���o����쪺IEnumerable����
                Dim IEnum As System.Collections.IEnumerable = CType(fi.GetValue(Me), System.Collections.IEnumerable)
                '�o�Ӫ����䴩IList �� IDictionary ������?�N���X�C
                Dim IListType As System.Type = fields(i).FieldType.GetInterface("IList", True)
                Dim IDicType As System.Type = fields(i).FieldType.GetInterface("IDictionary", True)

                Dim j As Integer = 0
                If Not IsNothing(IListType) Then
                    '���oIList�����C
                    Dim list As IList = CType(fields(i).GetValue(newObject), System.Collections.IList)
                    For Each obj As Object In IEnum
                        '�d�ݷ�e���O�_�䴩�䴩ICloneable �����C
                        ICloneType = obj.GetType().GetInterface("ICloneable", True)
                        If Not IsNothing(ICloneType) Then
                            '�p�G�䴩ICloneable �����A
                            '�ڭ̥Υ�?�]�m�C�������󪺧J��
                            Dim cloned As ICloneable = CType(obj, ICloneable)
                            list(j) = cloned.Clone()
                        End If
                        '�`�N�G�p�G�C���������䴩ICloneable�����A��?
                        '�b�J���C�����N�P��C��������ۦP
                        '�]�u�n�������O�ޥ������^
                        j += 1
                    Next
                ElseIf Not IsNothing(IDicType) Then
                    '���oIDictionary ����
                    Dim dic As IDictionary = CType(fields(i).GetValue(newObject), IDictionary)
                    j = 0
                    For Each de As DictionaryEntry In IEnum
                        '�d�ݷ�e���O�_�䴩�䴩ICloneable �����C
                        ICloneType = de.Value.GetType().GetInterface("ICloneable", True)
                        If Not IsNothing(ICloneType) Then
                            Dim cloned As ICloneable = CType(de.Value, ICloneable)
                            dic(de.Key) = cloned.Clone()
                        End If
                        j += 1
                    Next
                End If
            End If
            i += 1
        Next
        Return newObject
    End Function
#End Region

#Region "�w�p�o����Function"
    ' ''' <summary>
    ' ''' ���oSchema Name(User ID)�C
    ' ''' </summary>
    ' ''' <returns>String �Ǧ^Schema Table Name</returns>
    ' ''' <remarks>
    ' ''' BOTELOAN
    ' ''' </remarks>
    ' ''' <history>
    ' ''' 	[Titan]	2009/01/20	modify
    ' ''' </history>
    <ObsoleteAttribute("This function will be removed from future Versions.20110715")> _
    Private Function getDBSchema() As String
        Return "" 'Me.getSchemaName()
    End Function

    ' ''' <summary>
    ' ''' ���oSchema Name(User ID)�C
    ' ''' </summary>
    ' ''' <returns>String �Ǧ^Schema Table Name</returns>
    ' ''' <remarks>
    ' ''' BOTELOAN
    ' ''' </remarks>
    ' ''' <history>
    ' ''' 	[Titan]	2009/01/20	modify
    ' ''' </history>
    <ObsoleteAttribute("This function will be removed from future Versions.20110715")> _
    Public Function getSchemaName() As String
        Return "" 'Me.m_databaseManager.getSchemaName()
    End Function
#End Region
End Class
