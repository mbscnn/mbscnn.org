
<Serializable()> Public Class BaseException
    Inherits Exception

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub New(ByVal strMsg As String)
        MyBase.New(strMsg)
    End Sub

    Public Sub New(ByVal ex As Exception)
        MyBase.New(ex.Message, ex.InnerException)
    End Sub

    Public Sub New(ByVal ex As Exception, ByVal methodInfo As System.Reflection.MethodInfo)
        MyBase.New("Method¡G¡m" & methodPara(ex, methodInfo) & "¡n <br>" & vbCrLf & "Error Message¡G¡m" & getMsg(ex) & "¡n¡C")
    End Sub

    Public Sub New(ByVal sMessage As String, ByVal ex As Exception, ByVal methodInfo As System.Reflection.MethodInfo)
        MyBase.New("¿ù»~°T®§:" & sMessage & "¡C<br>" & vbCrLf & "Method¡G¡m" & methodPara(ex, methodInfo) & "¡n¡C<br>" & vbCrLf & "Error Message¡G¡m" & getMsg(ex) & "¡n¡C")
    End Sub

    'Public Sub New(ByVal sUrl As String, ByVal strMessage As String, ByVal Response As System.Web.HttpResponse)
    '    Response.Redirect(sUrl & "?sMsg=" & strMessage)
    'End Sub

    Public Sub New(ByVal sUrl As String, ByVal sCaseId As String, ByVal strMessage As String, ByVal Response As System.Web.HttpResponse)
        Response.Redirect(sUrl & "?sMsg=" & strMessage & "&CaseId=" & sCaseId)
    End Sub

    Public Sub New(ByVal ex As Exception, ByVal sUrl As String, ByVal Response As System.Web.HttpResponse)
        Dim strMessage As String = ex.Message
        Response.Redirect(sUrl & "?sMsg=" & strMessage)
    End Sub

    Private Shared Function getMsg(ByVal ex As Exception) As String
        Dim sb As New System.Text.StringBuilder
        Do
            If Not IsNothing(ex.Message) Then
                sb.Append(ex.Message.ToString)
            End If

            If Not IsNothing(ex.StackTrace) Then
                If Titan.Utility.isValidateData(sb.ToString) Then
                    sb.Append("<br>")
                End If

                sb.Append(ex.StackTrace.ToString & "¡C<br>")
            End If

            ex = ex.InnerException
        Loop Until (ex Is Nothing)

        Return sb.ToString
    End Function

    Public Shared Function methodPara(ByVal ex As Exception, ByVal methodInfo As System.Reflection.MethodBase) As String
        Dim sb As New System.Text.StringBuilder

        sb.Append(ex.TargetSite.DeclaringType.FullName & "." & methodInfo.Name & "(")
        For Each para As System.Reflection.ParameterInfo In methodInfo.GetParameters
            sb.Append(para.ParameterType).Append(" ").Append(para.Name).Append(",")
        Next
        If sb.Chars(sb.Length - 1).ToString.Equals(",") Then
            sb.Remove(sb.Length - 1, 1)
        End If
        sb.Append(")")
        Return sb.ToString
    End Function

End Class
