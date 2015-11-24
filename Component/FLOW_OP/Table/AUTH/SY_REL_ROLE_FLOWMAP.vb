Option Explicit On
Option Strict On

Imports com.Azion.NET.VB
Imports System.Text

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
        ''' <param name="sBrids">指定哪些分行的USER，可以不設，不設表示全部的分行</param>
        ''' <param name="nBraDepNos">指定哪些BRA_DEPNO的USER，可以不設，不設表示全部的分行</param>        
        ''' <returns>return a datatable of rows which contains STEP_NO, ROLEID, ROLENAME and BRID</returns>
        ''' <remarks>可使用#_，</remarks>
        Public Function GetRoleListByStepNo(ByVal sStepNo As String,
                                            Optional ByVal sBrids() As String = Nothing,
                                            Optional ByVal nBraDepNos() As Integer = Nothing) As DataTable

            '取消，可同時指定BRID及BRADEPNO
            'Dbg.Assert(Not (String.IsNullOrEmpty(sBrid) = False AndAlso nBraDepNo <> 0),
            '    "不可同時指定sBrid及nBraDepNo,最多只能指定一個(sBrid或nBraDepNo)")
            Dbg.Assert(Not sStepNo.Contains("?"), "請使用'_'代替'?'")

            Dim sSql As String
            Try
                sSql = "select distinct SY_REL_ROLE_FLOWMAP.STEP_NO, " & vbCrLf & _
                        "       SY_ROLE.ROLEID, " & vbCrLf & _
                        "       SY_ROLE.ROLENAME, " & vbCrLf & _
                        "       SY_BRANCH.BRID, " & vbCrLf & _
                        "       SY_BRANCH.BRA_DEPNO, " & vbCrLf & _
                        "       SY_BRANCH.BRCNAME " & vbCrLf & _
                        "  from SY_REL_ROLE_FLOWMAP " & vbCrLf & _
                        " inner join SY_ROLE " & vbCrLf & _
                        "    on SY_ROLE.ROLEID = SY_REL_ROLE_FLOWMAP.ROLEID " & vbCrLf & _
                        " inner join SY_REL_ROLE_USER " & vbCrLf & _
                        "    on SY_ROLE.ROLEID = SY_REL_ROLE_USER.ROLEID " & vbCrLf & _
                        " inner join SY_BRANCH " & vbCrLf & _
                        "    on SY_BRANCH.BRA_DEPNO = SY_REL_ROLE_USER.BRA_DEPNO " & vbCrLf

                If sStepNo.Contains("%") OrElse sStepNo.Contains("_") Then
                    sSql &= " where SY_REL_ROLE_FLOWMAP.STEP_NO like @STEP_NO@ " & vbCrLf
                Else
                    sSql &= " where SY_REL_ROLE_FLOWMAP.STEP_NO = @STEP_NO@ " & vbCrLf
                End If
                sSql &= "   and SY_BRANCH.DISABLED = '0' " & vbCrLf

                If Not IsNothing(sBrids) Then
                    Dim sb As New StringBuilder

                    For Each sbird As String In sBrids
                        If sb.Length > 0 Then
                            sb.Append(", '" & sbird & "' ")
                        Else
                            sb.Append(" '" & sbird & "' ")
                        End If
                    Next
                    sSql &= "   and SY_BRANCH.BRID in (" & sb.ToString & ")" & vbCrLf
                End If

                If Not IsNothing(nBraDepNos) Then
                    Dim sb As New StringBuilder

                    For Each nBraDepNo As Integer In nBraDepNos
                        If sb.Length > 0 Then
                            sb.Append(", " & nBraDepNo & " ")
                        Else
                            sb.Append(" " & nBraDepNo & " ")
                        End If
                    Next
                    sSql &= "   and SY_BRANCH.BRA_DEPNO in (" & sb.ToString & ")" & vbCrLf
                End If


                Return GetDataTable(sSql, "STEP_NO", sStepNo)

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Function


        ''' <summary>
        ''' 由 STEPNO取得USER列表
        ''' </summary>
        ''' <param name="sStepNo">若使用"041000%"表示，搜尋時會將041000??步驟的所有使用者列出</param>
        ''' <param name="sBrid">指定哪一分行的USER，可以不設，不設表示全部的分行</param>
        ''' <param name="nBraDepNo">指定哪一BRA_DEPNO的USER，可以不設，不設表示全部的分行</param>
        ''' <returns>return a datatable of rows which contains STAFFID and USERNAME</returns>
        ''' <remarks>sBrid或nBraDepNo可以不設或只能設一個</remarks>
        Public Function GetUserListByStepNo(ByVal sStepNo As String,
                                            Optional ByVal sBrid As String = Nothing,
                                            Optional ByVal nBraDepNo As Integer = 0) As DataTable

            Dim sSql As String
            Dim arrayBosParameters As New BosParamsList

            Dbg.Assert(Not sStepNo.Contains("?"), "請使用'_'代替'?'")

            Try
                Dbg.Assert(Not (String.IsNullOrEmpty(sBrid) = False AndAlso nBraDepNo > 0),
                           "sBrid或nBraDepNo不可以同時指定，僅能指定一個或都不指定")
                'sBrid或nBraDepNo不可以同時指定，僅能指定一個或都不指定

                sSql = _
                    "select distinct US.STAFFID, right(US.STAFFID,5) + ' ' + US.USERNAME as IDName, US.USERNAME," & vbCrLf & _
                    "       RF.ROLEID, SR.ROLENAME,  BR.BRA_DEPNO, BR.BRID, RF.STEP_NO " & vbCrLf & _
                    "  from SY_REL_ROLE_FLOWMAP RF " & vbCrLf & _
                    " inner join SY_REL_ROLE_USER RU " & vbCrLf & _
                    "    on RF.ROLEID = RU.ROLEID " & vbCrLf & _
                    " inner join SY_USER US " & vbCrLf & _
                    "    on US.STAFFID = RU.STAFFID " & vbCrLf & _
                    " inner join SY_BRANCH BR " & vbCrLf & _
                    "    on RU.BRA_DEPNO = BR.BRA_DEPNO  " & vbCrLf & _
                    " inner join SY_ROLE SR" & vbCrLf & _
                    "    on SR.ROLEID = RF.ROLEID " & vbCrLf
                If sStepNo.Contains("%") OrElse sStepNo.Contains("_") Then
                    sSql &= " where RF.STEP_NO like @STEP_NO@ " & vbCrLf & _
                            "   and US.STATUS = '0' " & vbCrLf
                Else
                    sSql &= " where RF.STEP_NO = @STEP_NO@ " & vbCrLf & _
                            "   and US.STATUS = '0' " & vbCrLf
                End If

                arrayBosParameters.Add("STEP_NO", sStepNo)

                '有指定BRA_DEPNO
                If nBraDepNo > 0 Then
                    sSql &= _
                    "   and RU.BRA_DEPNO = @BRA_DEPNO@ " & vbCrLf

                    arrayBosParameters.Add("BRA_DEPNO", nBraDepNo)
                End If


                '有指定BRID
                If String.IsNullOrEmpty(sBrid) = False Then
                    sSql &= _
                    "   and BR.BRID = @BRID@ " & vbCrLf

                    arrayBosParameters.Add("BRID", sBrid)
                End If

                Return GetDataTable(sSql, arrayBosParameters.ToArray)

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Function



        ''' <summary>
        ''' 從流程名稱，步驟代碼及人員編號取得分行，角色，人員，流程及步驟的所有資訊
        ''' </summary>
        ''' <param name="sFlowName"></param>
        ''' <param name="sStepNo"></param>
        ''' <param name="sStaffId"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetBranchRoleUserFlowidStepnoInfoByFlownameStepnoStaffid(
                                            ByVal sFlowName As String,
                                            ByVal sStepNo As String,
                                            ByVal sStaffId As String) As DataRowCollection
            Dim sSql As String

            Try

                sSql = _
                    "select BRA.*, " & vbCrLf & _
                    "       USR.*, " & vbCrLf & _
                    "       ROL.DISABLED, " & vbCrLf & _
                    "       ROL.PARENT ROLE_ROLEPARENT, " & vbCrLf & _
                    "       ROL.ROLEID, " & vbCrLf & _
                    "       ROL.ROLENAME, " & vbCrLf & _
                    "       ROL.ROLETYPE, " & vbCrLf & _
                    "       FID.*, " & vbCrLf & _
                    "       SSN.* " & vbCrLf & _
                    "  from SY_FLOW_ID FID " & vbCrLf & _
                    " inner join SY_REL_ROLE_FLOWMAP RRF " & vbCrLf & _
                    "    on FID.FLOW_ID = RRF.FLOW_ID " & vbCrLf & _
                    " inner join SY_STEP_NO SSN " & vbCrLf & _
                    "    on RRF.STEP_NO = SSN.STEP_NO " & vbCrLf & _
                    " inner join SY_REL_ROLE_USER RRR " & vbCrLf & _
                    "    on RRR.ROLEID = RRF.ROLEID " & vbCrLf & _
                    " inner join SY_ROLE ROL " & vbCrLf & _
                    "    on RRR.ROLEID = ROL.ROLEID " & vbCrLf & _
                    " inner join SY_BRANCH BRA " & vbCrLf & _
                    "    on BRA.BRA_DEPNO = RRR.BRA_DEPNO " & vbCrLf & _
                    " inner join SY_USER USR " & vbCrLf & _
                    "    on USR.STAFFID = RRR.STAFFID " & vbCrLf & _
                    " where FLOW_NAME = @FLOW_NAME@ " & vbCrLf & _
                    "   and SSN.STEP_NO = @STEP_NO@ " & vbCrLf & _
                    "   and RRR.STAFFID = @STAFFID@ " & vbCrLf

                Return GetDataRowCollection(sSql,
                                            "FLOW_NAME", sFlowName,
                                            "STEP_NO", sStepNo,
                                            "STAFFID", sStaffId)
            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Function



        Public Function GetUserBradepno(ByVal nFlowId As Integer,
                                        ByVal sStepNo As String,
                                        ByVal sStaffId As String,
                                        Optional ByVal nBradepno As Integer = 0) As DataRowCollection

            Dim sSql As String
            Dim listParams As New BosParamsList


            Try
                sSql = _
                    "select STAFFID, BRA_DEPNO " & vbCrLf & _
                    "  from SY_REL_ROLE_USER RRU " & vbCrLf & _
                    " inner join SY_REL_ROLE_FLOWMAP RRM " & vbCrLf & _
                    "    on RRU.ROLEID = RRM.ROLEID " & vbCrLf & _
                    " where RRM.FLOW_ID = @FLOW_ID@ " & vbCrLf & _
                    "   and RRM.STEP_NO = @STEP_NO@ " & vbCrLf & _
                    "   and RRU.STAFFID = @STAFFID@ " & vbCrLf

                listParams.Add("FLOW_ID", nFlowId)
                listParams.Add("STEP_NO", sStepNo)
                listParams.Add("STAFFID", sStaffId)

                If nBradepno <> 0 Then
                    sSql &= "   and BRA_DEPNO = @BRA_DEPNO@ " & vbCrLf
                    listParams.Add("BRA_DEPNO", nBradepno)
                End If

                Return GetDataRowCollection(sSql, listParams.ToArray)

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Function


        ''' <summary>
        ''' 使用者是否擁有權限執行步驟
        ''' </summary>
        ''' <param name="strStaffid"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function IsValidUser(ByVal strStaffid As String, ByVal nFlowId As Integer, ByVal sStepNo As String) As Boolean
            Dim nValue As Object
            Try
                nValue = ExecuteScalar( _
                    "select count(*)" & vbCrLf & _
                    "  from SY_REL_ROLE_FLOWMAP RRF" & vbCrLf & _
                    " inner join SY_REL_ROLE_USER RRU" & vbCrLf & _
                    "    on RRF.ROLEID = RRU.ROLEID" & vbCrLf & _
                    " where FLOW_ID = @FLOW_ID@" & vbCrLf & _
                    "   and STEP_NO = @STEP_NO@" & vbCrLf & _
                    "   and STAFFID = @STAFFID@",
                    "FLOW_ID", nFlowId,
                    "STEP_NO", sStepNo,
                    "STAFFID", strStaffid)

                If CDbType(Of Integer)(nValue, 0) = 0 Then
                    Return False
                Else
                    Return True
                End If

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Function

    End Class

End Namespace
