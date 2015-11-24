Option Explicit On
Option Strict On

Imports com.Azion.NET.VB

Public Class BRID_STAFFID
    Public Brid As String
    Public Brcname As String

    Public Staffid As String
    Public Staffname As String

    Public Shared Function Pair(ByVal Brid As String, ByVal Brcname As String,
                                ByVal Staffid As String, ByVal Staffname As String) As BRID_STAFFID
        Dim bs As New BRID_STAFFID
        bs.Brid = Brid
        bs.Brcname = Brcname
        bs.Staffid = Staffid
        bs.Staffname = Staffname
        Return bs
    End Function
End Class

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

        Public Shared Function getNewInstance(ByVal dbManager As com.Azion.NET.VB.DatabaseManager) As SY_REL_BRANCH_USER
            Return New SY_REL_BRANCH_USER(dbManager)
        End Function


        ''' <summary>
        ''' 取得使用者的分行資料
        ''' </summary>
        ''' <param name="strStaffid"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetBranchInfoByUserId(ByVal strStaffid As String) As DataRow
            Try
                Return GetDataRow( _
                    "select SY_BRANCH.* " & vbNewLine & _
                    "  from SY_BRANCH " & vbNewLine & _
                    " inner join SY_REL_BRANCH_USER " & vbNewLine & _
                    "    on SY_BRANCH.BRA_DEPNO = SY_REL_BRANCH_USER.BRA_DEPNO " & vbNewLine & _
                    " where SY_REL_BRANCH_USER.STAFFID = @STAFFID@ ",
                    "STAFFID", strStaffid)

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Function



        ''' <summary>
        ''' 取得使用者(STAFFID)的BRADEPNO
        ''' </summary>
        ''' <param name="strStaffid"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function getBraDepNoByUserId(ByVal strStaffid As String) As Integer
            Dim objBraDepNo As Object
            Try
                objBraDepNo = ExecuteScalar("select BRA_DEPNO " & vbNewLine & _
                                            "  from SY_REL_BRANCH_USER " & vbNewLine & _
                                            " where SY_REL_BRANCH_USER.STAFFID = @STAFFID@",
                                            "STAFFID", strStaffid)

                Dim nBraDepNo As Integer = CDbType(Of Integer)(objBraDepNo, 0)

                If nBraDepNo = 0 Then
                    Throw New SYException(
                        String.Format("無法取得使用者的BRA_DEPNO，STAFFID={0}", strStaffid),
                        SYMSG.SYRELBRANCHUSER_BRADEPNO_NOT_FOUND,
                        GetLastSQL)
                End If

                Return nBraDepNo

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Function

        ''' <summary>
        ''' 由BRA_DEPNO取得使用者(STAFFID)
        ''' </summary>
        ''' <param name="sBRA_DEPNO">BRA_DEPNO</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        ''' 2013/03/15 Add By [Nick]
        Public Function GetUserByBRA_DEPNO(ByVal sBRA_DEPNO As String) As DataTable
            Dim dtStaffid As DataTable
            Try
                dtStaffid = GetDataTable("select STAFFID " & vbNewLine & _
                                            "  from SY_REL_BRANCH_USER " & vbNewLine & _
                                            " where SY_REL_BRANCH_USER.BRA_DEPNO = @BRA_DEPNO@",
                                            "BRA_DEPNO", sBRA_DEPNO)

                Return dtStaffid

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Function

    End Class

End Namespace
