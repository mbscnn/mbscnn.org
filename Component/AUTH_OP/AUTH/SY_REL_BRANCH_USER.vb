Option Explicit On
Option Strict On

Imports com.Azion.NET.VB

Namespace TABLE

    Public Class SY_REL_BRANCH_USER
        Inherits BosBase

        ''' <summary>
        ''' Constructor
        ''' </summary>
        ''' <param name="dbManager">DatabaseManager</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal dbManager As com.Azion.NET.VB.DatabaseManager)
            MyBase.new("SY_REL_BRANCH_USER", dbManager)
        End Sub


        '''' <summary>
        '''' 取得使用者的分行資料
        '''' </summary>
        '''' <param name="strStaffid"></param>
        '''' <returns></returns>
        '''' <remarks></remarks>
        'Public Function GetBranchInfoByUserId(ByVal strStaffid As String) As DataRow
        '    Try
        '        Return GetDataRow( _
        '            "SELECT SY_BRANCH.* " & vbNewLine & _
        '            "  FROM SY_BRANCH " & vbNewLine & _
        '            " INNER JOIN SY_REL_BRANCH_USER " & vbNewLine & _
        '            "    ON SY_BRANCH.BRA_DEPNO = SY_REL_BRANCH_USER.BRA_DEPNO " & vbNewLine & _
        '            " WHERE (SY_REL_BRANCH_USER.STAFFID = @STAFFID@) ",
        '            New BosParameter() {PARAMETER("STAFFID", strStaffid)})
        '    Catch ex As Exception
        '        Throw
        '    End Try
        'End Function

#Region "濟南昱勝添加"
#Region "Lake Function"
        ''' <summary>
        ''' 根據傳入的人員編號及部門代碼刪除該人員資料
        ''' </summary>
        ''' <param name="sBraDepNo">當前部門及其下屬部門</param>
        ''' <param name="sStaffId">人員編號</param>
        ''' <returns></returns>
        ''' <remarks>[Lake] 2012/06/26 Created</remarks>
        Public Function deleteStaffByBraDepNoAndStaffId(sBraDepNo As String, sStaffId As String) As Boolean
            Try
                Dim sSql As String = "DELETE SY_REL_BRANCH_USER WHERE BRA_DEPNO IN(" & sBraDepNo & ")" & _
                    "AND STAFFID = " & ProviderFactory.PositionPara & "STAFFID"

                Dim paras As IDbDataParameter = ProviderFactory.CreateDataParameter("STAFFID", sStaffId)

                If Me.getDatabaseManager.isTransaction Then
                    Return DBObject.ExecuteNonQuery(Me.getDatabaseManager.getTransaction, CommandType.Text, sSql, paras) > 0
                Else
                    Return DBObject.ExecuteNonQuery(Me.getDatabaseManager.getConnection, CommandType.Text, sSql, paras) > 0
                End If
            Catch ex As ProviderException
                Throw ex
            Catch ex As BosException
                Throw ex
            Catch ex As Exception
                Throw ex
            End Try
        End Function

        ''' <summary>
        ''' 根據部門代碼刪除人員
        ''' </summary>
        ''' <param name="sBraDepNo">部門代碼</param>
        ''' <returns></returns>
        ''' <remarks>[Lake] 2012/05/22</remarks>
        Public Function deleteByBraDepNo(ByVal sBraDepNo As String) As Boolean
            Try
                Dim sSql As String = "DELETE FROM SY_REL_BRANCH_USER WHERE BRA_DEPNO IN('" & sBraDepNo & "')"

                If Me.getDatabaseManager.isTransaction Then
                    Return DBObject.ExecuteNonQuery(Me.getDatabaseManager.getTransaction, CommandType.Text, sSql) > 0
                Else
                    Return DBObject.ExecuteNonQuery(Me.getDatabaseManager.getConnection, CommandType.Text, sSql) > 0
                End If
            Catch ex As ProviderException
                Throw ex
            Catch ex As BosException
                Throw ex
            Catch ex As Exception
                Throw ex
            End Try
        End Function

        ''' <summary>
        ''' 根據人員編號刪除人員
        ''' </summary>
        ''' <param name="sStaffId">人員編號</param>
        ''' <returns></returns>
        ''' <remarks>[Lake] 2012/05/22</remarks>
        Public Function deleteByStaffId(ByVal sStaffId As String) As Boolean
            Try
                Dim sSql As String = "DELETE FROM SY_REL_BRANCH_USER WHERE STAFFID IN(" & ProviderFactory.PositionPara & "STAFFID)"

                Dim paras As IDbDataParameter = ProviderFactory.CreateDataParameter("STAFFID", sStaffId)

                If Me.getDatabaseManager.isTransaction Then
                    Return DBObject.ExecuteNonQuery(Me.getDatabaseManager.getTransaction, CommandType.Text, sSql, paras) > 0
                Else
                    Return DBObject.ExecuteNonQuery(Me.getDatabaseManager.getConnection, CommandType.Text, sSql, paras) > 0
                End If
            Catch ex As ProviderException
                Throw ex
            Catch ex As BosException
                Throw ex
            Catch ex As Exception
                Throw ex
            End Try
        End Function

        ''' <summary>
        ''' 根據主鍵查詢
        ''' </summary>
        ''' <param name="sBraDepNo">部門代碼</param>
        ''' <param name="sStaffId">員工編號</param>
        ''' <returns></returns>
        ''' <remarks>[Lake] 2012/05/24 Created</remarks>
        Public Function loadByPk(ByVal sBraDepNo As String, ByVal sStaffId As String) As Boolean
            Try
                Dim paras(1) As IDbDataParameter
                paras(0) = ProviderFactory.CreateDataParameter("BRA_DEPNO", sBraDepNo)
                paras(1) = ProviderFactory.CreateDataParameter("STAFFID", sStaffId)

                Return Me.loadData(paras)
            Catch ex As ProviderException
                Throw ex
            Catch ex As BosException
                Throw ex
            Catch ex As Exception
                Throw ex
            End Try
        End Function
#End Region
#End Region
    End Class
End Namespace
