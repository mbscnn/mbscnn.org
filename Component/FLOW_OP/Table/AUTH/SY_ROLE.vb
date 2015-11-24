Option Explicit On
Option Strict On

Imports com.Azion.NET.VB
Imports FLOW_OP.TABLE

Namespace TABLE
    Public Class SY_ROLE
        Inherits SY_TABLEBASE

        ''' <summary>
        ''' Constructor
        ''' </summary>
        ''' <param name="dbManager">DatabaseManager</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal dbManager As com.Azion.NET.VB.DatabaseManager)
            MyBase.new("SY_ROLE", dbManager)
        End Sub


        ''' <summary>
        ''' 取得角色的TYPE
        ''' </summary>
        ''' <param name="nRoleId"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetRoleType(ByVal nRoleId As Integer) As String
            Try
                Return CDbType(Of String)(
                    ExecuteScalar(
                        "select ROLETYPE from SY_ROLE where ROLEID = @ROLEID@",
                        "ROLEID", nRoleId),
                    "")

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Function


    End Class

End Namespace
