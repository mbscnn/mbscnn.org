Option Explicit On
Option Strict On

Imports com.Azion.NET.VB
Imports System.Text
Imports System.Text.RegularExpressions

Namespace TABLE

    Public Class SY_CONFIGURATION
        Inherits SY_TABLEBASE


        ''' <summary>
        ''' Constructor
        ''' </summary>
        ''' <param name="dbManager">DatabaseManager</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal dbManager As com.Azion.NET.VB.DatabaseManager)
            MyBase.new("SY_CONFIGURATION", dbManager)
        End Sub


        ''' <summary>
        ''' 取得字串
        ''' </summary>
        ''' <param name="sName"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetValue(ByVal sName As String) As String
            Try
                Dim dr As DataRow

                dr = GetDataRow(Nothing, "NAME", sName)

                If IsNothing(dr) Then
                    Throw New SYException(
                        String.Format("無法取得內容(SY_CONFIGURATION.VALUE), SY_CONFIGURATION.NAME = {0} ", sName),
                                  SYMSG.SYCONFRATION_NAME_NOT_FOUND)
                End If

                Return CDbType(Of String)(dr("VALUE"), "")

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)

            End Try
        End Function


        Public Sub SetValue(ByVal sName As String, ByVal sValue As String)
            Try
                InsertUpdate("NAME", sName,
                             "VALUE", sValue)
            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)

            End Try
        End Sub

    End Class

End Namespace
