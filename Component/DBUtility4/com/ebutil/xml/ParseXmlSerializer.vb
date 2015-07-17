Imports System
Imports System.IO
Imports System.Xml
Imports System.Xml.Serialization

Public Class ParseXmlSerializer
    Private m_xmlSerializer As XmlSerializer
    Private type As Type
    Private m_sClassName As String = ""
    Private m_sDLLName As String = "EN_OP"
    Private m_sNameSpase As String = "Eloan.EN_OP"
    Private objClass As Object

    Public Sub New(ByVal type As type)
        Me.type = type
        Me.m_xmlSerializer = New XmlSerializer(Me.type)
    End Sub

    Public Sub New(ByVal sClassName As String)
        m_sClassName = sClassName
        objClass = ClassLoad.ClassforName(m_sDLLName, m_sNameSpase, m_sClassName)
        Me.type = objClass.GetType
        Me.m_xmlSerializer = New XmlSerializer(Me.type)
    End Sub

    Public Sub New(ByVal sDLLName As String, ByVal sNameSpase As String, ByVal sClassName As String)
        m_sDLLName = sDLLName
        m_sNameSpase = sNameSpase
        m_sClassName = sClassName
        objClass = ClassLoad.ClassforName(m_sDLLName, m_sNameSpase, m_sClassName)
        Me.type = objClass.GetType
        Me.m_xmlSerializer = New XmlSerializer(Me.type)
    End Sub

    Public Function Deserialize(ByVal xml As String) As Object
        Dim reader As TextReader = New StringReader(xml)
        Return Deserialize(reader)
    End Function
    Public Function Deserialize(ByVal doc As XmlDocument) As Object
        Dim reader = New StringReader(doc.OuterXml)
        Return Deserialize(reader)
    End Function
    Public Function Deserialize(ByVal reader As TextReader) As Object
        Dim o As Object = Me.m_xmlSerializer.Deserialize(reader)
        reader.Close()
        Return o
    End Function
    Public Function Serialize(ByVal obj As Object) As XmlDocument
        Dim xml As String = StringSerialize(obj)
        Dim doc As XmlDocument = New XmlDocument
        doc.PreserveWhitespace = True
        doc.LoadXml(xml)
        doc = Clean(doc)
        Return doc
    End Function
    Private Function StringSerialize(ByVal obj As Object) As String
        Dim w As TextWriter = WriterSerialize(obj)
        Dim xml As String = w.ToString()
        w.Close()
        Return xml.Trim()
    End Function
    Private Function WriterSerialize(ByVal obj As Object) As TextWriter
        Dim w As TextWriter = New StringWriter

        Me.m_xmlSerializer = New XmlSerializer(Me.type)
        Me.m_xmlSerializer.Serialize(w, obj)
        w.Flush()
        Return w
    End Function
    Private Function Clean(ByVal doc As XmlDocument) As XmlDocument
        doc.RemoveChild(doc.FirstChild)
        Dim first As XmlNode = doc.FirstChild
        For Each n As XmlNode In doc.ChildNodes
            If (n.NodeType = XmlNodeType.Element) Then
                first = n
                Exit For
            End If
        Next
        If Not first.Attributes Is Nothing Then
            Dim a As XmlAttribute = Nothing
            a = first.Attributes("xmlns:xsd")
            If Not a Is Nothing Then first.Attributes.Remove(a)
            a = first.Attributes("xmlns:xsi")
            If Not a Is Nothing Then first.Attributes.Remove(a)
        End If
        Return doc
    End Function

    Public Shared Function ReadXML(ByVal sXML As String, ByVal type As type) As Object
        Dim serializer As ParseXmlSerializer = New ParseXmlSerializer(type)
        Try
            Return serializer.Deserialize(sXML)
        Catch ex As Exception
            Throw
        End Try
        Return New Object
    End Function

    Public Shared Function ReadXML(ByVal sXML As String, ByVal sDLLName As String, ByVal sNameSpase As String, ByVal sClassName As String) As Object
        Dim serializer As ParseXmlSerializer = New ParseXmlSerializer(sDLLName, sNameSpase, sClassName)
        Try
            Return serializer.Deserialize(sXML)
        Catch ex As Exception
            Throw
        End Try
        Return New Object
    End Function

    Public Shared Function ReadFile(ByVal file As String, ByVal type As Type) As Object
        Dim serializer As ParseXmlSerializer = New ParseXmlSerializer(type)
        Try
            Dim xml As String = String.Empty
            Dim reader As StreamReader = New StreamReader(file, System.Text.Encoding.Default)
            xml = reader.ReadToEnd()
            reader.Close()
            reader = Nothing
            Return serializer.Deserialize(xml)
        Catch ex As Exception
            Throw
        End Try
        Return New Object
    End Function

    Public Shared Function ReadFile(ByVal file As String, ByVal sDLLName As String, ByVal sNameSpase As String, ByVal sClassName As String) As Object
        Dim serializer As ParseXmlSerializer = New ParseXmlSerializer(sDLLName, sNameSpase, sClassName)
        Try
            Dim xml As String = String.Empty
            Dim reader As StreamReader = New StreamReader(file, System.Text.Encoding.Default)
            xml = reader.ReadToEnd()
            reader.Close()
            reader = Nothing
            Return serializer.Deserialize(xml)
        Catch ex As Exception
            Throw
        End Try
        Return New Object
    End Function

    Public Shared Function ReadFile(ByVal file As String, ByVal sClassName As String) As Object
        Dim serializer As ParseXmlSerializer = New ParseXmlSerializer(sClassName)
        Try
            Dim xml As String = String.Empty
            Dim reader As StreamReader = New StreamReader(file, System.Text.Encoding.Default)
            xml = reader.ReadToEnd()
            reader.Close()
            reader = Nothing
            Return serializer.Deserialize(xml)
        Catch ex As Exception
            Throw
        End Try
        Return New Object
    End Function

    Public Shared Function WriteFile(ByVal file As String, ByVal obj As Object, ByVal sClassName As String) As Boolean
        Dim ok As Boolean = False
        Dim serializer As ParseXmlSerializer = New ParseXmlSerializer(sClassName)
        Try
            Dim xml As String = serializer.Serialize(obj).OuterXml
            Dim writer As StreamWriter = New StreamWriter(file, False)
            'writer.Write("<?xml version=""1.0"" encoding=""Big5""?>")
            'writer.Write("<?xml version=""1.0"" encoding=""UTF-8""?>")
            writer.Write(xml)
            writer.Flush()
            writer.Close()
            writer = Nothing
            ok = True
        Catch ex As Exception
            Throw
        End Try
        Return ok
    End Function

    Public Shared Function genXmlFile(ByVal obj As Object, ByVal sClassName As String) As String
        Dim parseXmlSerializer As ParseXmlSerializer = New ParseXmlSerializer(sClassName)
        Try
            Dim xml As String = parseXmlSerializer.Serialize(obj).OuterXml
            Return xml
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function genXmlFile(ByVal obj As Object, ByVal sDLLName As String, ByVal sNameSpase As String, ByVal sClassName As String) As Object
        Dim parseXmlSerializer As ParseXmlSerializer = New ParseXmlSerializer(sDLLName, sNameSpase, sClassName)
        Try
            Dim xml As String = parseXmlSerializer.Serialize(obj).OuterXml
            Return xml
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function genXmlFile(ByVal obj As Object, ByVal type As Type) As String
        Dim parseXmlSerializer As ParseXmlSerializer = New ParseXmlSerializer(type)
        Try
            Dim xml As String = parseXmlSerializer.Serialize(obj).OuterXml
            Return xml
        Catch ex As Exception
            Throw
        End Try
    End Function

End Class
