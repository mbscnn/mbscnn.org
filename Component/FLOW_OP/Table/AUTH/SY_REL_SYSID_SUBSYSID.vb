﻿Option Explicit On
Option Strict On

Imports com.Azion.NET.VB

Namespace TABLE

    Public Structure _SYSID_SUBSYSID
        Public sSysid As String
        Public sSubsysid As String
    End Structure


    Public Class SY_REL_SYSID_SUBSYSID
        Inherits BosBase

        ''' <summary>
        ''' Constructor
        ''' </summary>
        ''' <param name="dbManager">DatabaseManager</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal dbManager As com.Azion.NET.VB.DatabaseManager)
            MyBase.new("SY_REL_SYSID_SUBSYSID", dbManager)
        End Sub


    End Class

End Namespace
