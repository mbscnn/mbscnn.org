Imports com.Azion.NET.VB

Public Class MB_CLASS_MList
    Inherits BosList

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.New("MB_CLASS_M", dbManager)
    End Sub

    Overrides Function newBos() As BosBase
        Dim bos As MB_CLASS_M = New MB_CLASS_M(MyBase.getDatabaseManager)
        Return bos
    End Function
End Class
