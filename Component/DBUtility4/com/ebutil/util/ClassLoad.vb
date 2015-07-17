Public Class ClassLoad

    Public Shared Function ClassforName(ByVal strDLLName As String, ByVal strNameSpase As String, ByVal strClassName As String, ByVal ParamArray args As Object()) As Object
        '檢核所需的Object
        Dim assem As System.reflection.Assembly
        Dim objType As Type
        Dim obj As Object

        Try
            assem = System.Reflection.Assembly.Load(strDLLName)

            objType = assem.GetType(strNameSpase & "." & strClassName) '"BotEloan.PR.PR_MAIN"

            obj = System.Activator.CreateInstance(objType, args)

        Catch argumentNullException As argumentNullException 'type 為 Null 參照 (即 Visual Basic 中的 Nothing)。 
            ' Console.WriteLine(argumentNullException)
            Throw New Exception("找不到 CLASS")

        Catch targetInvocationException As System.Reflection.TargetInvocationException '被呼叫的建構函式擲回例外狀況。
            Throw New Exception("被呼叫的建構函式擲回例外狀況")

        Catch methodAccessException As MethodAccessException ' 呼叫端沒有使用權限來呼叫這個建構函式。 
            Throw New Exception("呼叫端沒有使用權限來呼叫這個建構函式")
        Catch missingMethodException As MissingMethodException ' 找不到相符的公用建構函式。 
            Throw New Exception("找不到相符的公用建構函式")

        Catch memberAccessException As memberAccessException ' 無法建立抽象類別 (Abstract Class) 的執行個體，或這個成員是以晚期繫結機制叫用的。 
            Throw New Exception("無法建立抽象類別 (Abstract Class) 的執行個體，或這個成員是以晚期繫結機制叫用的")

      
        Catch methodAccessException As TypeLoadException ' type 不是有效的型別。 
            Throw New Exception("找不到 CLASS")

        Catch ex As Exception
            assem = Nothing
            obj = Nothing
            Throw
        End Try
        Return obj
    End Function

    Public Shared Function ClassforName(ByVal strDLLName As String, ByVal strNameSpase As String, ByVal strClassName As String) As Object
        '檢核所需的Object
        Dim assem As System.Reflection.Assembly
        Dim objType As Type
        Dim obj As Object

        Try
            assem = System.Reflection.Assembly.Load(strDLLName)

            objType = assem.GetType(strNameSpase & "." & strClassName) '"BotEloan.PR.PR_MAIN"

            obj = System.Activator.CreateInstance(objType)
        Catch argumentNullException As ArgumentNullException 'type 為 Null 參照 (即 Visual Basic 中的 Nothing)。 
            'Console.WriteLine(argumentNullException)
            Throw New Exception("找不到 CLASS")

        Catch targetInvocationException As System.Reflection.TargetInvocationException '被呼叫的建構函式擲回例外狀況。
            Throw New Exception("被呼叫的建構函式擲回例外狀況")

        Catch methodAccessException As MethodAccessException ' 呼叫端沒有使用權限來呼叫這個建構函式。 
            Throw New Exception("呼叫端沒有使用權限來呼叫這個建構函式")

        Catch missingMethodException As MissingMethodException ' 找不到相符的公用建構函式。 
            Throw New Exception("找不到相符的公用建構函式")

        Catch memberAccessException As MemberAccessException ' 無法建立抽象類別 (Abstract Class) 的執行個體，或這個成員是以晚期繫結機制叫用的。 
            Throw New Exception("無法建立抽象類別 (Abstract Class) 的執行個體，或這個成員是以晚期繫結機制叫用的")

        Catch methodAccessException As TypeLoadException ' type 不是有效的型別。 
            Throw New Exception("找不到 CLASS")

        Catch ex As Exception
            assem = Nothing
            obj = Nothing
            Throw
        End Try
        Return obj
    End Function
End Class
