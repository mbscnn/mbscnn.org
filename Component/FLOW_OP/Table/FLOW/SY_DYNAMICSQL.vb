Option Explicit On
Option Strict On

Imports com.Azion.NET.VB

Namespace TABLE

    Public Class SY_DYNAMICSQL
        Inherits SY_TABLEBASE

        ''' <summary>
        ''' Constructor
        ''' </summary>
        ''' <param name="dbManager">DatabaseManager</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal dbManager As com.Azion.NET.VB.DatabaseManager)
            MyBase.new("SY_DYNAMICSQL", dbManager)
        End Sub

        ''' <summary>
        ''' 取得DynamicSQL
        ''' </summary>
        ''' <param name="name"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetDynamicSQL(ByVal name As String) As String
            Try
                Return CDbType(Of String)(ExecuteScalar("select VALUE from SY_DYNAMICSQL where NAME = @NAME@", "NAME", name), "")
            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try
        End Function

    End Class

End Namespace
