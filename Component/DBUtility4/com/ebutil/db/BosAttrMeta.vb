Option Explicit On
Option Strict On

Imports System.Data
Imports System.Collections
Public Class BosAttrMeta

    Dim m_sColName As String 'column name
    Dim m_iProviderType As Integer  'Provider type 
    Dim m_sProviderTypeName As String 'data type name
    'Dim m_iMaxLength As Integer

    Dim m_iDBType As System.data.DbType
    Sub New()
    End Sub

    Sub New(ByVal sColName As String, ByVal iDbType As System.data.DbType)
        setColName(sColName)
        setDataType(iDbType)
    End Sub

    Sub constructAttrMeta(ByVal column As DataColumn)
        setColName(column.ColumnName.ToString) 'COLUMN_NAME            
        setDataType(DbUtility.getFieldType(column.DataType.ToString)) 'DB_TYPE
       
    End Sub

    Sub constructAttrMeta(ByVal objRows As DataRow)
        Try
            setColName(objRows!ColumnName.ToString) 'COLUMN_NAME            
            setDataType(DbUtility.getFieldType(objRows!DataType.ToString)) 'DB_TYPE
            setProviderType(CType(objRows!ProviderType, Integer))  'ProviderType
            setProviderTypeName(ProviderFactory.getFieldType(Me.getProviderType()).ToString) 'ProviderTypeName

        Catch ex As Exception
            Throw New Exception
        End Try
    End Sub

    Sub constructAttrMeta(ByVal recHash As Hashtable)
        setColName(CType(recHash.Item("COLUMN_NAME"), String)) 'COLUMN_NAME
        setDataType(CType(recHash.Item("DB_TYPE"), DbType))
        setProviderType(CType(recHash.Item("PROVIDER_TYPE"), Integer))
        setProviderTypeName(ProviderFactory.getFieldType(CType(recHash.Item("PROVIDER_TYPE"), Integer)))
    End Sub

    Sub setColName(ByVal sColName As String)
        Me.m_sColName = sColName.ToUpper
    End Sub

    Function getColName() As String
        Return Me.m_sColName
    End Function

    Sub setProviderTypeName(ByVal sProviderTypeName As String)
        Me.m_sProviderTypeName = sProviderTypeName
    End Sub

    Function getProviderTypeName() As String
        Return Me.m_sProviderTypeName
    End Function

    Sub setProviderType(ByVal iProviderType As Integer)
        Me.m_iProviderType = iProviderType
    End Sub

    Function getProviderType() As Integer
        Return Me.m_iProviderType
    End Function

    Sub setDataType(ByVal iDBType As System.data.DbType)
        Me.m_iDBType = iDBType
    End Sub

    Function getDataType() As System.data.DbType
        Return Me.m_iDBType
    End Function
End Class

