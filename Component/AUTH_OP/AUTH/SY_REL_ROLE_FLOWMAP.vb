Option Explicit On
Option Strict On

Imports com.Azion.NET.VB

Namespace TABLE

    Public Class SY_REL_ROLE_FLOWMAP
        Inherits BosBase

        ''' <summary>
        ''' Constructor
        ''' </summary>
        ''' <param name="dbManager">DatabaseManager</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal dbManager As com.Azion.NET.VB.DatabaseManager)
            MyBase.new("SY_REL_ROLE_FLOWMAP", dbManager)
        End Sub



        ''' <summary>
        ''' 由 STEPNO取得ROLE列表
        ''' </summary>
        ''' <param name="sStepNo">若使用"041000%"表示，搜尋時會將041000??步驟的所有使用者列出</param>
        ''' <returns>return a datatable of rows which contains STEP_NO, ROLEID, ROLENAME and BRID</returns>
        ''' <remarks>可使用#_，</remarks>
        Public Function GetRoleListByStepNo(ByVal sStepNo As String) As DataTable

            Dim sSql As String

            Try
                If sStepNo.Contains("%") AndAlso sStepNo.Contains("_") Then
                    sSql = _
                        "select SY_REL_ROLE_FLOWMAP.STEP_NO, " & _
                        "       SY_ROLE.ROLEID, " & _
                        "       SY_ROLE.ROLENAME, " & _
                        "       SY_BRANCH.BRID " & _
                        "  from SY_REL_ROLE_FLOWMAP " & _
                        " inner join SY_ROLE " & _
                        "    on SY_ROLE.ROLEID = SY_REL_ROLE_FLOWMAP.ROLEID " & _
                        " inner join SY_REL_ROLE_USER " & _
                        "    on SY_ROLE.ROLEID = SY_REL_ROLE_USER.ROLEID " & _
                        " inner join SY_BRANCH " & _
                        "    on SY_BRANCH.BRA_DEPNO = SY_REL_ROLE_USER.BRA_DEPNO " & _
                        " where SY_REL_ROLE_FLOWMAP.STEP_NO like @STEP_NO@ "
                Else
                    sSql = _
                        "select SY_REL_ROLE_FLOWMAP.STEP_NO, " & _
                        "       SY_ROLE.ROLEID, " & _
                        "       SY_ROLE.ROLENAME, " & _
                        "       SY_BRANCH.brid " & _
                        "  from SY_REL_ROLE_FLOWMAP " & _
                        " inner join SY_ROLE " & _
                        "    on SY_ROLE.ROLEID = SY_REL_ROLE_FLOWMAP.ROLEID " & _
                        " inner join SY_REL_ROLE_USER " & _
                        "    on SY_ROLE.ROLEID = SY_REL_ROLE_USER.ROLEID " & _
                        " inner join SY_BRANCH " & _
                        "    on SY_BRANCH.BRA_DEPNO = SY_REL_ROLE_USER.BRA_DEPNO " & _
                        " where SY_REL_ROLE_FLOWMAP.STEP_NO = @STEP_NO@ "
                End If

                Return GetDataTable(sSql, "STEP_NO", sStepNo)

            Catch ex As Exception
                Throw
            End Try

        End Function


        ''' <summary>
        ''' 由 STEPNO取得USER列表
        ''' </summary>
        ''' <param name="sStepNo"></param>
        ''' <returns>return a datatable of rows which contains STAFFID and USERNAME</returns>
        ''' <remarks></remarks>
        Public Function GetUserListByStepNo(ByVal sStepNo As String) As DataTable

            Dim tbl As DataTable
            tbl = New DataTable("USERINFO")


            Dim column As DataColumn

            column = New DataColumn()
            column.DataType = GetType(String)
            column.ColumnName = "STAFFID"
            tbl.Columns.Add(column)

            column = New DataColumn()
            column.DataType = GetType(String)
            column.ColumnName = "USERNAME"
            tbl.Columns.Add(column)


            Dim row As DataRow
            row = tbl.NewRow()
            row("STAFFID") = "S004206"
            row("USERNAME") = "陳健渝"
            tbl.Rows.Add(row)

            Return tbl

            'Dim sSql As String

            'Try
            '    sSql = _
            '        "select SY_USER.STAFFID, " & _
            '        "       SY_USER.USERNAME " & _
            '        "  from SY_REL_ROLE_FLOWMAP " & _
            '        " inner join SY_ROLE " & _
            '        "    on SY_ROLE.ROLEID = SY_REL_ROLE_FLOWMAP.ROLEID " & _
            '        " inner join SY_REL_ROLE_USER " & _
            '        "    on SY_ROLE.ROLEID = SY_REL_ROLE_USER.ROLEID " & _
            '        " inner join SY_USER " & _
            '        "    on SY_USER.STAFFID = SY_REL_ROLE_USER.STAFFID " & _
            '        " where SY_REL_ROLE_FLOWMAP.STEP_NO = @STEP_NO@ "

            '    Return GetDataTable(sSql, New BosParameter() {PARAMETER("STEP_NO", sStepNo)})

            'Catch ex As Exception
            '    Throw
            'End Try
        End Function


    End Class

End Namespace
