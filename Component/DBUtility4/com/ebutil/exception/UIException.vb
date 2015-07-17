<Serializable()> Public Class UIException
    Inherits BaseException

    Public Sub New(ByVal strMsg As String)
        MyBase.New(strMsg)
    End Sub

    Public Sub New(ByVal ex As Exception, ByVal methodInfo As System.Reflection.MethodInfo)
        MyBase.New(ex, methodInfo)
    End Sub

    Public Sub New(ByVal sMessage As String, ByVal ex As Exception, ByVal methodInfo As System.Reflection.MethodInfo)
        MyBase.New(sMessage, ex, methodInfo)
    End Sub
End Class
