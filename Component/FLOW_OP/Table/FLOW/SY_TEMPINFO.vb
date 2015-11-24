Option Explicit On
Option Strict On

Imports com.Azion.NET.VB
Imports FLOW_OP.TABLE

Namespace TABLE

    Public Class SY_TEMPINFO
        Inherits SY_TABLEBASE

        Public Sub New(ByVal dbManager As com.Azion.NET.VB.DatabaseManager)
            MyBase.new("SY_TEMPINFO", dbManager)
        End Sub



    End Class

End Namespace