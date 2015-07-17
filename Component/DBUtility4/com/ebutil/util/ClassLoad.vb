Public Class ClassLoad

    Public Shared Function ClassforName(ByVal strDLLName As String, ByVal strNameSpase As String, ByVal strClassName As String, ByVal ParamArray args As Object()) As Object
        '�ˮ֩һݪ�Object
        Dim assem As System.reflection.Assembly
        Dim objType As Type
        Dim obj As Object

        Try
            assem = System.Reflection.Assembly.Load(strDLLName)

            objType = assem.GetType(strNameSpase & "." & strClassName) '"BotEloan.PR.PR_MAIN"

            obj = System.Activator.CreateInstance(objType, args)

        Catch argumentNullException As argumentNullException 'type �� Null �ѷ� (�Y Visual Basic ���� Nothing)�C 
            ' Console.WriteLine(argumentNullException)
            Throw New Exception("�䤣�� CLASS")

        Catch targetInvocationException As System.Reflection.TargetInvocationException '�Q�I�s���غc�禡�Y�^�ҥ~���p�C
            Throw New Exception("�Q�I�s���غc�禡�Y�^�ҥ~���p")

        Catch methodAccessException As MethodAccessException ' �I�s�ݨS���ϥ��v���өI�s�o�ӫغc�禡�C 
            Throw New Exception("�I�s�ݨS���ϥ��v���өI�s�o�ӫغc�禡")
        Catch missingMethodException As MissingMethodException ' �䤣��۲Ū����Ϋغc�禡�C 
            Throw New Exception("�䤣��۲Ū����Ϋغc�禡")

        Catch memberAccessException As memberAccessException ' �L�k�إߩ�H���O (Abstract Class) ���������A�γo�Ӧ����O�H�ߴ�ô������s�Ϊ��C 
            Throw New Exception("�L�k�إߩ�H���O (Abstract Class) ���������A�γo�Ӧ����O�H�ߴ�ô������s�Ϊ�")

      
        Catch methodAccessException As TypeLoadException ' type ���O���Ī����O�C 
            Throw New Exception("�䤣�� CLASS")

        Catch ex As Exception
            assem = Nothing
            obj = Nothing
            Throw
        End Try
        Return obj
    End Function

    Public Shared Function ClassforName(ByVal strDLLName As String, ByVal strNameSpase As String, ByVal strClassName As String) As Object
        '�ˮ֩һݪ�Object
        Dim assem As System.Reflection.Assembly
        Dim objType As Type
        Dim obj As Object

        Try
            assem = System.Reflection.Assembly.Load(strDLLName)

            objType = assem.GetType(strNameSpase & "." & strClassName) '"BotEloan.PR.PR_MAIN"

            obj = System.Activator.CreateInstance(objType)
        Catch argumentNullException As ArgumentNullException 'type �� Null �ѷ� (�Y Visual Basic ���� Nothing)�C 
            'Console.WriteLine(argumentNullException)
            Throw New Exception("�䤣�� CLASS")

        Catch targetInvocationException As System.Reflection.TargetInvocationException '�Q�I�s���غc�禡�Y�^�ҥ~���p�C
            Throw New Exception("�Q�I�s���غc�禡�Y�^�ҥ~���p")

        Catch methodAccessException As MethodAccessException ' �I�s�ݨS���ϥ��v���өI�s�o�ӫغc�禡�C 
            Throw New Exception("�I�s�ݨS���ϥ��v���өI�s�o�ӫغc�禡")

        Catch missingMethodException As MissingMethodException ' �䤣��۲Ū����Ϋغc�禡�C 
            Throw New Exception("�䤣��۲Ū����Ϋغc�禡")

        Catch memberAccessException As MemberAccessException ' �L�k�إߩ�H���O (Abstract Class) ���������A�γo�Ӧ����O�H�ߴ�ô������s�Ϊ��C 
            Throw New Exception("�L�k�إߩ�H���O (Abstract Class) ���������A�γo�Ӧ����O�H�ߴ�ô������s�Ϊ�")

        Catch methodAccessException As TypeLoadException ' type ���O���Ī����O�C 
            Throw New Exception("�䤣�� CLASS")

        Catch ex As Exception
            assem = Nothing
            obj = Nothing
            Throw
        End Try
        Return obj
    End Function
End Class
