Imports com.Azion.NET.VB

Public Class MB_MBIMPList
    Inherits BosList

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.New("MB_MBIMP", dbManager)
    End Sub

    Overrides Function newBos() As BosBase
        Dim bos As MB_MBIMP = New MB_MBIMP(MyBase.getDatabaseManager)
        Return bos
    End Function
End Class
