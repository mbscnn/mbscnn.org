Option Explicit On
Option Strict On

Imports com.Azion.NET.VB

Namespace TABLE

    Public Class SY_BRANCH
        Inherits BosBase

        ''' <summary>
        ''' Constructor
        ''' </summary>
        ''' <param name="dbManager">DatabaseManager</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal dbManager As com.Azion.NET.VB.DatabaseManager)
            MyBase.new("SY_BRANCH", dbManager)
        End Sub

        Public Shared Function getNewInstance(ByVal dbManager As com.Azion.NET.VB.DatabaseManager) As SY_BRANCH
            Return New SY_BRANCH(dbManager)
        End Function

        ''' <summary>
        ''' 取得BRID最上層的BRA_DEPNO
        ''' </summary>
        ''' <param name="sBRID"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetTopBraDepNo(ByVal sBRID As String) As Integer
            Try
                Dim data As Object

                data = ExecuteScalar( _
                    "select BRA_DEPNO " & vbNewLine & _
                    "  from SY_BRANCH " & vbNewLine & _
                    " where PARENT = 0 " & vbNewLine & _
                    "   and BRID = @BRID@ ",
                    "BRID", sBRID)

                'If IsNothing(data) Then
                '    Throw New Exception(String.Format("無法取得[BRID:{0}]最上層的BRA_DEPNO", sBRID))
                'End If

                Return CDbType(Of Integer)(data)

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try
        End Function

        Public Function GetBRID(ByVal nBraDepNo As Integer) As String
            Try
                Dim data As Object

                data = ExecuteScalar( _
                    "select BRID " & vbNewLine & _
                    "  from SY_BRANCH " & vbNewLine & _
                    " where BRA_DEPNO = @BRA_DEPNO@ ",
                   "BRA_DEPNO", nBraDepNo)

                Return CDbType(Of String)(data)

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try
        End Function

        Public Function GetBranchName(ByVal nBraDepNo As Integer) As String
            Try
                Return QueryString("[SY_BRANCH].[BRCNAME]", "BRA_DEPNO", nBraDepNo)

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try
        End Function



        Public Function GetBraDepNo(ByVal sBrid As String) As Integer
            Try
                Dim data As Object

                data = ExecuteScalar( _
                    "select BRA_DEPNO " & vbNewLine & _
                    "  from SY_BRANCH " & vbNewLine & _
                    " where BRID = @BRID@ " & vbNewLine & _
                    "   and PARENT = 0 " & vbNewLine,
                    "BRID", sBrid)

                If IsNothing(data) Then
                    data = ExecuteScalar( _
                        "select TOP 1 BRA_DEPNO " & vbNewLine & _
                        "  from SY_BRANCH " & vbNewLine & _
                        " where BRID = @BRID@ " & vbNewLine,
                        "BRID", sBrid)
                End If

                Dim nBraDepNo As Integer = CDbType(Of Integer)(data, 0)


                If nBraDepNo = 0 Then
                    Throw New SYException(
                        String.Format("無法取得[BRID:{0}]的BRA_DEPNO", sBrid),
                        SYMSG.SYBRANCH_BRADEPNO_NOT_FOUND,
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
        ''' 由FLOWNAME, STEP_NO及STAFFID取得BRID及BRA_DEPNO
        ''' </summary>
        ''' <param name="sFlowname"></param>
        ''' <param name="sStepNo"></param>
        ''' <param name="sStaffid"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetBridByFlownameStepno(ByVal sFlowname As String,
                                                ByVal sStepNo As String,
                                                Optional ByVal sStaffid As String = "") As DataTable

            Dim arrayParams As New BosParamsList
            Dim sSql As String

            Try
                sSql = _
                    "select distinct BRA.BRA_DEPNO, BRA.BRID " & vbNewLine & _
                    "  from SY_BRANCH BRA " & vbNewLine & _
                    " inner join SY_REL_ROLE_USER RRU " & vbNewLine & _
                    "    on BRA.BRA_DEPNO = RRU.BRA_DEPNO " & vbNewLine & _
                    " inner join SY_REL_ROLE_FLOWMAP RRF " & vbNewLine & _
                    "    on RRU.ROLEID = RRF.ROLEID " & vbNewLine & _
                    " inner join SY_FLOW_ID FI " & vbNewLine & _
                    "    on FI.FLOW_ID = RRF.FLOW_ID " & vbNewLine & _
                    " where FLOW_NAME = @FLOW_NAME@ " & vbNewLine & _
                    "   and STEP_NO = @STEP_NO@ " & vbNewLine

                arrayParams.Add("FLOW_NAME", sFlowname)
                arrayParams.Add("STEP_NO", sStepNo)

                If String.IsNullOrEmpty(sStaffid) = False Then
                    sSql &= "   and STAFFID = @STAFFID@ " & vbNewLine
                    arrayParams.Add("STAFFID", sStaffid)
                End If

                Return GetDataTable(sSql, arrayParams.ToArray)

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try
        End Function


        ''' <summary>
        ''' 取得分行下所有的科別代碼
        ''' </summary>
        ''' <param name="nBraDepno"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetChildFromParent(ByVal nBraDepno As Integer) As Integer()

            Dim drc As DataRowCollection
            Dim listBraDepno As New List(Of Integer)

            Try
                drc = GetDataRowCollection( _
                            "with RC(BRA_DEPNO, " & vbNewLine & _
                            "DISABLED, " & vbNewLine & _
                            "PARENT) as " & vbNewLine & _
                            " (select SY_BRANCH.BRA_DEPNO, " & vbNewLine & _
                            "         SY_BRANCH.DISABLED, " & vbNewLine & _
                            "         SY_BRANCH.PARENT " & vbNewLine & _
                            "    from SY_BRANCH " & vbNewLine & _
                            "   where SY_BRANCH.DISABLED = '0' " & vbNewLine & _
                            "     and SY_BRANCH.BRA_DEPNO = @BRA_DEPNO@ " & vbNewLine & _
                            "  union all " & vbNewLine & _
                            "  select a.BRA_DEPNO, a.DISABLED, a.PARENT " & vbNewLine & _
                            "    from SY_BRANCH as a, RC " & vbNewLine & _
                            "   where a.PARENT = RC.BRA_DEPNO " & vbNewLine & _
                            "     and RC.DISABLED = '0' " & vbNewLine & _
                            "     and a.DISABLED = '0') " & vbNewLine & _
                            "select * from RC " & vbNewLine,
                            "BRA_DEPNO", nBraDepno)

                If IsNothing(drc) Then
                    Throw New SYException(
                        String.Format("無法取得[BRADEPNO:{0}]內的所有部門", nBraDepno),
                        SYMSG.SYBRANCH_CHILDBRADEPNO_NOT_FOUND,
                        GetLastSQL)
                End If

                For Each dr As DataRow In drc
                    listBraDepno.Add(CInt(dr("BRA_DEPNO")))
                Next

                Return listBraDepno.ToArray()

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try
        End Function
    End Class

End Namespace
