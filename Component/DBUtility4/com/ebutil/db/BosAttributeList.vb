Option Explicit On
Option Strict On

Public Class BosAttributeList

    Dim m_attributes As New System.Collections.ArrayList

    Sub New()
    End Sub

    Sub add(ByVal attr As BosAttribute)
        m_attributes.Add(attr)
    End Sub

    Function add(ByVal index As Integer) As BosAttribute
        Return CType(m_attributes.Item(index), BosAttribute)
    End Function

    Sub clear()
        m_attributes.Clear()
    End Sub

    Function size() As Integer
        Return m_attributes.Count
    End Function

    Function item(ByVal i As Integer) As BosAttribute
        Return CType(m_attributes.Item(i), BosAttribute)
    End Function

End Class
