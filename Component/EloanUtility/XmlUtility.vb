Option Explicit On
Option Strict On

Public Class XmlUtility

    ''' <summary>
    ''' convert datatable to xml
    ''' </summary>
    ''' <param name="datatable"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function convertDataTable2XML(ByVal datatable As Data.DataTable) As String

        Dim sb As New Text.StringBuilder
        Using sw As New IO.StringWriter(sb)
            datatable.WriteXml(sw, Data.XmlWriteMode.IgnoreSchema, False)
        End Using

        Return sb.ToString
    End Function

    ''' <summary>
    ''' convert xml to datatable
    ''' </summary>
    ''' <param name="sXml"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function convertXML2DataTable(ByVal sXml As String) As Data.DataTable
        Dim ds As New Data.DataSet
        Dim dt As Data.DataTable = Nothing
        If sXml <> Nothing Then
            Dim xmlDocument As New Xml.XmlDocument
            xmlDocument.LoadXml(sXml)
            ds.ReadXml(New Xml.XmlNodeReader(xmlDocument))
            If ds.Tables.Count > 0 Then
                dt = ds.Tables(0)
                Return dt
            End If
        End If
        Return dt
    End Function
End Class
