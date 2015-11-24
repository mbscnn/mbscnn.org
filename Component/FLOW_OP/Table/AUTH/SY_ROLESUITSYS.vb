Option Explicit On
Option Strict On

Imports com.Azion.NET.VB

Namespace TABLE

    Public Class SY_ROLESUITSYS
        Inherits BosBase


        ''' <summary>
        ''' Constructor
        ''' </summary>
        ''' <param name="dbManager">DatabaseManager</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal dbManager As com.Azion.NET.VB.DatabaseManager)
            MyBase.new("SY_ROLESUITSYS", dbManager)
        End Sub


    End Class

End Namespace
