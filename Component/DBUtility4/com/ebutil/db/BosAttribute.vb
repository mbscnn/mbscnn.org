Option Explicit On
Option Strict On

Public Class BosAttribute

    Dim m_attrMeta As BosAttrMeta
    Dim m_value As Object
    Dim m_isSeted As Boolean
    Sub New()
    End Sub

    Sub New(ByVal attrMeta As BosAttrMeta)
        setAttrMeta(attrMeta)
    End Sub

    Sub clearValue()
        m_value = Nothing
    End Sub

    Function getAttrMeta() As BosAttrMeta
        Return m_attrMeta
    End Function

    Sub setAttrMeta(ByVal attrMeta As BosAttrMeta)
        m_attrMeta = attrMeta
    End Sub

    Function getValue() As Object
        If IsDBNull(m_value) Then
            m_value = Nothing
        End If
        Return m_value
    End Function

    Sub setValue(ByVal value As Object, Optional ByVal isSeted As Boolean = False)
        If IsDBNull(value) Then
            value = Nothing
        End If
        m_value = value
        m_isSeted = isSeted
    End Sub

    Function isSeted() As Boolean
        Return m_isSeted
    End Function

    Function getColName() As String
        Return m_attrMeta.getColName()
    End Function

    Function getDataType() As System.data.DbType
        Return m_attrMeta.getDataType()
    End Function

    Function getProviderType() As Integer
        Return m_attrMeta.getProviderType()
    End Function

    'Public Overrides Function toString() As String
    '    Dim sValue As String = "" & CType(IIf(IsNothing(Me.getValue()), "", Me.getValue()), String)
    '    Dim sb As New System.Text.StringBuilder("<Attr>")
    '    sb.Append("<AttrID>" & Me.getColName() & "</AttrID>")
    '    sb.Append("<AttrValue>" & sValue.Trim() & "</AttrValue>")
    '    sb.Append("</Attr>")
    '    Return sb.ToString()
    'End Function
End Class
