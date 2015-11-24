Option Explicit On
Option Strict On

Imports com.Azion.NET.VB

Namespace TABLE

    Public Class SY_STEP_NO
        Inherits SY_TABLEBASE

        Public Const SystemStep_Start As String = "START"
        Public Const SystemStep_End As String = "END"
        Public Const SystemStep_Cancel As String = "CANCEL"
        Public Const SystemStep_Splitter As String = "SPLITTER"
        Public Const SystemStep_Joiner As String = "JOINER"
        Public Const SystemStep_Next As String = "NEXT"

        ''' <summary>
        ''' Constructor
        ''' </summary>
        ''' <param name="dbManager">DatabaseManager</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal dbManager As com.Azion.NET.VB.DatabaseManager)
            MyBase.new("SY_STEP_NO", dbManager)
        End Sub

    End Class

End Namespace
