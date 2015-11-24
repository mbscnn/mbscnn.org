Option Explicit On
Option Strict On

Imports com.Azion.NET.VB



Public Class ROLE_BRADEPNO_STAFFID
    Public RoleId As Integer
    Public BraDepNo As Integer
    Public StaffId As String

    Public Shared Function Pair(ByVal nRoleId As Integer,
                                ByVal nBraDepNo As Integer,
                                ByVal nStaffId As String) As ROLE_BRADEPNO_STAFFID
        Dim rbs As New ROLE_BRADEPNO_STAFFID
        rbs.RoleId = nRoleId
        rbs.BraDepNo = nBraDepNo
        rbs.StaffId = nStaffId
        Return rbs
    End Function

    Public Shared Function ToArray(ByVal ParamArray objs() As ROLE_BRADEPNO_STAFFID) As ROLE_BRADEPNO_STAFFID()
        Dim rbsList As New List(Of ROLE_BRADEPNO_STAFFID)

        For Each obj As ROLE_BRADEPNO_STAFFID In objs
            rbsList.Add(obj)
        Next

        Return rbsList.ToArray
    End Function

End Class



Namespace TABLE

    Public Class SY_REL_ROLE_USER
        Inherits BosBase


        ''' <summary>
        ''' Constructor
        ''' </summary>
        ''' <param name="dbManager">DatabaseManager</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal dbManager As com.Azion.NET.VB.DatabaseManager)
            MyBase.new("SY_REL_ROLE_USER", dbManager)
        End Sub


        Public Shared Function getNewInstance(ByVal dbManager As com.Azion.NET.VB.DatabaseManager) As SY_REL_ROLE_USER
            Return New SY_REL_ROLE_USER(dbManager)
        End Function



        ''' <summary>
        ''' 由FlowId, StepNo, BraDepNo取得使用者列表
        ''' </summary>
        ''' <param name="nFlowId"></param>
        ''' <param name="sStepNo"></param>
        ''' <param name="nBradepno"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetUserlistByFlowidStepnoBradepno(ByVal nFlowId As Integer, ByVal sStepNo As String, nBradepno As Integer) As DataRowCollection

            Dim strSql As String

            Try
                strSql = _
                    "select distinct SY_REL_ROLE_USER.STAFFID " & vbCrLf & _
                    "  from SY_REL_ROLE_FLOWMAP " & vbCrLf & _
                    " inner join SY_REL_ROLE_USER " & vbCrLf & _
                    "    on SY_REL_ROLE_USER.ROLEID = SY_REL_ROLE_FLOWMAP.ROLEID " & vbCrLf & _
                    " where SY_REL_ROLE_FLOWMAP.FLOW_ID = @FLOW_ID@ " & vbCrLf & _
                    "   and SY_REL_ROLE_FLOWMAP.STEP_NO = @STEP_NO@ " & vbCrLf & _
                    "   and SY_REL_ROLE_USER.BRA_DEPNO = @BRA_DEPNO@"

                Return GetDataRowCollection2(strSql,
                                            New BosParameter() {
                                                PARAMETER("FLOW_ID", nFlowId),
                                                PARAMETER("STEP_NO", sStepNo),
                                                PARAMETER("BRA_DEPNO", nBradepno)})

            Catch ex As Exception
                Throw
            End Try

            Return Nothing
        End Function


        ''' <summary>
        ''' 由使用者編號取得角色代碼
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetRolesByStaffid(ByVal sStaffid As String) As DataTable

            Dim dt As DataTable

            Try
                dt = GetDataTable( _
                        "select RRU.*, BR.BRID, BR.BRCNAME, US.USERNAME, BR.PARENT PARENTID" & vbCrLf & _
                        "  from SY_REL_ROLE_USER RRU " & vbCrLf & _
                        " inner join SY_ROLE RO " & vbCrLf & _
                        "    on RO.ROLEID = RRU.ROLEID " & vbCrLf & _
                        " inner join SY_USER US " & vbCrLf & _
                        "    on US.STAFFID = RRU.STAFFID " & vbCrLf & _
                        " inner join SY_BRANCH BR " & vbCrLf & _
                        "    on BR.BRA_DEPNO = RRU.BRA_DEPNO " & vbCrLf & _
                        " where RRU.STAFFID = @STAFFID@ " & vbCrLf & _
                        "   and RO.DISABLED = '0' " & vbCrLf & _
                        "   and US.STATUS = '0' " & vbCrLf,
                        "STAFFID", sStaffid)

                'If IsNothing(drc) Then
                '    Throw New Exception(String.Format("無法取得使用者的角色：STAFFID={0}",
                '                      sStaffid))
                'End If

                Return dt

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Function


        ''' <summary>
        ''' 由Brid及Roleid取得使用者列表
        ''' </summary>
        ''' <param name="sBrid"></param>
        ''' <param name="nRoleid"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetUserListByBridRoleid(ByVal sBrid As String, ByVal nRoleid As Integer) As DataTable

            Dim dtStaffid As DataTable

            Try
                dtStaffid = GetDataTable( _
                        "select STAFFID " & vbCrLf & _
                        "  from SY_REL_ROLE_USER RRU " & vbCrLf & _
                        " inner join SY_BRANCH BR " & vbCrLf & _
                        "    on BR.BRA_DEPNO = RRU.BRA_DEPNO " & vbCrLf & _
                        " where BR.BRID = @BRID@ " & vbCrLf & _
                        "   and RRU.ROLEID = @ROLEID@ ",
                        "BRID", sBrid,
                        "ROLEID", nRoleid)

                Return dtStaffid

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Function

        ''' <summary>
        ''' 由Brid及Roleid取得使用者列表
        ''' </summary>
        ''' <param name="sBrid"></param>
        ''' <param name="nRoleid1">先撈區督導</param>
        ''' <param name="nRoleid2">再排除法金總處長</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        '''   2013/03/15  Add By  [Nick]
        Public Function GetUserListByBridRoleidWithCONDID(ByVal sBrid As String, ByVal nRoleid1 As Integer, ByVal nRoleid2 As Integer) As DataTable

            Dim dtStaffid As DataTable

            Try
                dtStaffid = GetDataTable( _
                        "select STAFFID " & vbCrLf & _
                        "  from SY_REL_ROLE_USER RRU " & vbCrLf & _
                        " inner join SY_BRANCH BR " & vbCrLf & _
                        "    on BR.BRA_DEPNO = RRU.BRA_DEPNO " & vbCrLf & _
                        " where BR.BRID = @BRID@ " & vbCrLf & _
                        "   and RRU.ROLEID = @ROLEID1@ " & vbCrLf & _
                        "   and not exists(select * from SY_REL_ROLE_USER x where ROLEID = @ROLEID2@ and RRU.ROLEID = x.ROLEID)",
                        "BRID", sBrid,
                        "ROLEID1", nRoleid1,
                        "ROLEID2", nRoleid2)

                Return dtStaffid

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Function

        ''' <summary>
        ''' 由BraDepNo及Roleid取得使用者列表
        ''' </summary>
        ''' <param name="nBraDepNo"></param>
        ''' <param name="nRoleid"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetUserListByBradepnoRoleid(ByVal nBraDepNo As Integer, ByVal nRoleid As Integer) As DataTable

            Dim dtStaffid As DataTable

            Try
                dtStaffid = GetDataTable( _
                        "select STAFFID " & vbCrLf & _
                        "  from SY_REL_ROLE_USER RRU " & vbCrLf & _
                        " inner join SY_BRANCH BR " & vbCrLf & _
                        "    on BR.BRA_DEPNO = RRU.BRA_DEPNO " & vbCrLf & _
                        " where BR.BRA_DEPNO = @BRA_DEPNO@ " & vbCrLf & _
                        "   and RRU.ROLEID = @ROLEID@ ",
                        "BRA_DEPNO", nBraDepNo,
                        "ROLEID", nRoleid)

                Return dtStaffid

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Function

        ''' <summary>
        ''' 由BraDepNo及Roleid取得使用者列表
        ''' </summary>
        ''' <param name="nBraDepNo"></param>
        ''' <param name="nRoleid"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetUserListWithNameByBradepnoRoleid(ByVal nBraDepNo As Integer, ByVal nRoleid As Integer) As DataTable

            Dim dtStaffid As DataTable

            Try
                dtStaffid = GetDataTable( _
                        "select RRU.STAFFID, USR.USERNAME " & vbCrLf & _
                        "  from SY_REL_ROLE_USER RRU " & vbCrLf & _
                        " inner join SY_BRANCH BR " & vbCrLf & _
                        "    on BR.BRA_DEPNO = RRU.BRA_DEPNO " & vbCrLf & _
                        " inner join SY_USER USR " & vbCrLf & _
                        "    on RRU.STAFFID = USR.STAFFID " & vbCrLf & _
                        " where BR.BRA_DEPNO = @BRA_DEPNO@ " & vbCrLf & _
                        "   and RRU.ROLEID = @ROLEID@ ",
                        "BRA_DEPNO", nBraDepNo,
                        "ROLEID", nRoleid)

                Return dtStaffid

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Function

        ''' <summary>
        ''' 由Roleid取得使用者列表
        ''' </summary>
        ''' <param name="nRoleid"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetUserListByRoleid(ByVal nRoleid As Integer) As DataTable

            Dim dtStaffid As DataTable

            Try
                dtStaffid = GetDataTable( _
                        "select RU.STAFFID " & vbCrLf & _
                        "  from SY_REL_ROLE_USER RU " & vbCrLf & _
                        " inner join SY_USER UR " & vbCrLf & _
                        "    on RU.STAFFID = UR.STAFFID " & vbCrLf & _
                        " WHERE ROLEID = @ROLEID@ " & vbCrLf & _
                        "   AND STATUS = '0' ", _
                        "ROLEID", nRoleid)

                Return dtStaffid

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Function


    End Class

End Namespace
