Option Explicit On
Option Strict On

Imports com.Azion.NET.VB

Namespace TABLE

    Public Class SY_FLOW_DEF
        Inherits SY_TABLEBASE

        ''' <summary>
        ''' Constructor
        ''' </summary>
        ''' <param name="dbManager">DatabaseManager</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal dbManager As com.Azion.NET.VB.DatabaseManager)
            MyBase.new("SY_FLOW_DEF", dbManager)

        End Sub

        ''' <summary>
        ''' 取得下一步驟的所有資訊
        ''' </summary>
        ''' <param name="flowId"></param>
        ''' <param name="currentStepNo"></param>
        ''' <param name="pn"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetNextStep(ByVal flowId As Integer,
                                    ByVal currentStepNo As String,
                                    Optional ByVal pn As String = "N") As DataRowCollection
            Try
                Return GetNextStep(flowId, currentStepNo, pn, Nothing)
            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try
        End Function



        ''' <summary>
        ''' 取得下一步驟的所有資訊，或取得特定的下一步驟的資訊
        ''' </summary>
        ''' <param name="flowId"></param>
        ''' <param name="currentStepNo"></param>
        ''' <param name="pn"></param>
        ''' <param name="nextStepNo"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetNextStep(ByVal flowId As Integer,
                                    ByVal currentStepNo As String,
                                    ByVal pn As String,
                                    ByVal nextStepNo As String) As DataRowCollection

            Dim sSql As String
            Dim arrayBosParameters As New BosParamsList

            Dbg.Assert(String.IsNullOrEmpty(currentStepNo) = False)
            Dbg.Assert(String.IsNullOrEmpty(pn) = False)

            Try
                sSql = _
                    "select FD.*, FI.FLOW_NAME, FM.TYPE, NFM.TYPE NEXTTYPE " & vbCrLf & _
                    "  from SY_FLOW_DEF FD  " & vbCrLf & _
                    " inner join SY_FLOW_ID FI  " & vbCrLf & _
                    "    on FI.FLOW_ID = FD.FLOW_ID  " & vbCrLf & _
                    " inner join SY_FLOW_MAP FM  " & vbCrLf & _
                    "    on FM.FLOW_ID = FI.FLOW_ID  " & vbCrLf & _
                    "   and FM.STEP_NO = FD.CURR_STEP_NO  " & vbCrLf & _
                    "  left join SY_FLOW_MAP NFM  " & vbCrLf & _
                    "    on NFM.FLOW_ID = FI.FLOW_ID  " & vbCrLf & _
                    "   and NFM.STEP_NO = FD.NEXT_STEP_NO  " & vbCrLf & _
                    " where FD.FLOW_ID = @FLOW_ID@  " & vbCrLf & _
                    "   and (PN = @PN@ or FM.TYPE = 'P' or FM.TYPE = 'N')  " & vbCrLf & _
                    "   and CURR_STEP_NO = @CURR_STEP_NO@  " & vbCrLf

                arrayBosParameters.Add("FLOW_ID", flowId)
                arrayBosParameters.Add("PN", pn)
                arrayBosParameters.Add("CURR_STEP_NO", currentStepNo)

                If Not String.IsNullOrEmpty(nextStepNo) Then
                    sSql &=
                    "   and NEXT_STEP_NO = @NEXT_STEP_NO@  " & vbCrLf

                    arrayBosParameters.Add("NEXT_STEP_NO", nextStepNo)
                End If

                sSql &= " order by COND_ORDER " & vbCrLf

                'Dbg.Assert(nextStepNo <> "04101100")

                Return GetDataRowCollection(sSql, arrayBosParameters.ToArray)
            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try
        End Function


        ''' <summary>
        ''' 取得下一步驟的SPANBRADEPNO
        ''' </summary>
        ''' <param name="flowId"></param>
        ''' <param name="currentStepNo"></param>
        ''' <param name="pn"></param>
        ''' <param name="nextStepNo"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetNextStepSpanBraDepNo(ByVal flowId As Integer,
                            ByVal currentStepNo As String,
                            ByVal pn As String,
                            ByVal nextStepNo As String) As BRADEPNO_BRID()
            Dim drc As DataRowCollection
            Dim spanBraDepNo() As BRADEPNO_BRID

            Try
                drc = GetNextStep(flowId, currentStepNo, pn, nextStepNo)

                If IsNothing(drc) OrElse drc.Count = 0 Then
                    Return Nothing
                End If

                For Each dr As DataRow In drc
                    spanBraDepNo = getSYRelFlowMap_SpanBranch.GetSpanBranch(
                        CInt(dr("FLOW_ID")), CStr(dr("NEXT_STEP_NO")))

                    If IsNothing(spanBraDepNo) = False Then
                        Return spanBraDepNo
                    End If
                Next

                Return Nothing
            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try
        End Function


        ''' <summary>
        ''' 取得流程條件的所有資訊
        ''' </summary>
        ''' <param name="nCondId"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetConditionDef(ByVal nCondId As Integer) As DataRowCollection

            Dim drc As DataRowCollection

            Try
                drc = GetDataRowCollection( _
                    "select * " & vbCrLf & _
                    "  from SY_FLOW_DEF " & vbCrLf & _
                    " inner join SY_CONDITION_ID " & vbCrLf & _
                    "    on SY_FLOW_DEF.COND_ID = SY_CONDITION_ID.COND_ID " & vbCrLf & _
                    " where SY_FLOW_DEF.COND_ID = @COND_ID@ ",
                    "COND_ID", nCondId)

                Return drc
            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Function


        ''' <summary>
        ''' 上下步驟是否禁止為同一人
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function isProhibiteSameUser(ByVal nFlowId As Integer,
                                            ByVal sCurrentStepNo As String,
                                            ByVal sNextStepNo As String) As Boolean

            Dim obj As Object
            Try

                obj = ExecuteScalar( _
                    "select count(*) " & vbCrLf & _
                    "  from SY_FLOW_DEF FD " & vbCrLf & _
                    " inner join SY_FLOW_MAP FM1 " & vbCrLf & _
                    "    on FM1.FLOW_ID = FD.FLOW_ID " & vbCrLf & _
                    "   and FM1.STEP_NO = FD.CURR_STEP_NO " & vbCrLf & _
                    " inner join SY_FLOW_MAP FM2 " & vbCrLf & _
                    "    on FM2.FLOW_ID = FD.FLOW_ID " & vbCrLf & _
                    "   and FM2.STEP_NO = FD.NEXT_STEP_NO " & vbCrLf & _
                    " where FD.FLOW_ID = @FLOW_ID@ " & vbCrLf & _
                    "   and CURR_STEP_NO = @CURR_STEP_NO@ " & vbCrLf & _
                    "   and NEXT_STEP_NO = @NEXT_STEP_NO@ " & vbCrLf & _
                    "   and PROHIBITE_SAME = @PROHIBITE_SAME@ " & vbCrLf & _
                    "   and FM1.TYPE is null " & vbCrLf & _
                    "   and FM2.TYPE is null ",
                    "FLOW_ID", nFlowId,
                    "CURR_STEP_NO", sCurrentStepNo,
                    "NEXT_STEP_NO", sNextStepNo,
                    "PROHIBITE_SAME", "Y")

                If CInt(obj) = 0 Then
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




        ''' <summary>
        ''' 目前步驟到下一步驟的方向
        ''' </summary>
        ''' <param name="nFlowId"></param>
        ''' <param name="sCurrStepNo"></param>
        ''' <param name="nNextStepNo"></param>
        ''' <returns></returns>
        ''' <remarks>若反方向，SubFlow_Seq=MAX(SubFlow_Seq)+1</remarks>
        Public Function GetPN(ByVal nFlowId As Integer,
                              ByVal sCurrStepNo As String,
                              ByVal nNextStepNo As String) As String
            Dim dr As DataRow


            Try
                dr = GetDataRow(Nothing,
                                "FLOW_ID", nFlowId,
                                "CURR_STEP_NO", sCurrStepNo,
                                "NEXT_STEP_NO", nNextStepNo)

                If IsNothing(dr) Then
                    Throw New SYException(String.Format(
                                        "無法取得流程步驟的PN方向：SY_FLOW_DEF.FLOW_ID={0}，CURR_STEP_NO={1}，NEXT_STEP_NO={2}。",
                                        nFlowId, sCurrStepNo, nNextStepNo), SYMSG.SYFLOWDEF_DIRECTION_NOT_FOUND, GetLastSQL)
                End If

                Return CDbType(Of String)(dr("PN"), String.Empty)
            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try
        End Function


        ''' <summary>
        ''' 取得第一個步驟，但不包含START
        ''' </summary>
        ''' <param name="sFlowName"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetFirstStepExceptStart(ByVal sFlowName As String) As String

            Dim obj As Object

            Try
                obj = ExecuteScalar(
                     "select top 1 STEP_NO " & vbCrLf & _
                     "  from (select NEXT_STEP_NO as STEP_NO, " & vbCrLf & _
                     "               ROW_NUMBER() over(order by FD.NEXT_STEP_NO) _ORDER " & vbCrLf & _
                     "          from SY_FLOW_DEF FD " & vbCrLf & _
                     "         inner join SY_FLOW_ID FI " & vbCrLf & _
                     "            on FD.FLOW_ID = FI.FLOW_ID " & vbCrLf & _
                     "         where FLOW_NAME = @FLOW_NAME@ " & vbCrLf & _
                     "           and CURR_STEP_NO = 'START' " & vbCrLf & _
                     "        union " & vbCrLf & _
                     "        select *, " & vbCrLf & _
                     "               10000 + ROW_NUMBER() over(order by NON_START.STEP_NO) as _ORDER " & vbCrLf & _
                     "          from (select FD1.CURR_STEP_NO as STEP_NO " & vbCrLf & _
                     "                  from SY_FLOW_DEF FD1 " & vbCrLf & _
                     "                  left join SY_FLOW_DEF FD2 " & vbCrLf & _
                     "                    on FD1.FLOW_ID = FD2.FLOW_ID " & vbCrLf & _
                     "                   and FD2.PN = 'N' " & vbCrLf & _
                     "                   and FD1.CURR_STEP_NO = FD2.NEXT_STEP_NO " & vbCrLf & _
                     "                 inner join SY_FLOW_ID FI " & vbCrLf & _
                     "                    on FI.FLOW_ID = FD1.FLOW_ID " & vbCrLf & _
                     "                 where FD2.FLOW_ID is null " & vbCrLf & _
                     "                   and FD1.CURR_STEP_NO <> 'START' " & vbCrLf & _
                     "                   and FLOW_NAME = @FLOW_NAME@) NON_START) ALLREC " & vbCrLf & _
                     " order by _ORDER " & vbCrLf,
                     "FLOW_NAME", sFlowName)

                If IsDBNull(obj) OrElse IsNothing(obj) Then
                    Throw New SYException(String.Format(
                                        "無法取得流程步驟的第一個步驟，FLOW_NAME={0}",
                                        sFlowName), SYMSG.SYFLOWDEF_FIRSTSTEP_NOT_FOUND, GetLastSQL)
                End If

                Return CDbType(Of String)(obj, Nothing)

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Function


        ''' <summary>
        ''' 輸入Staffid, FlowName, StepNo取得分行代碼
        ''' </summary>
        ''' <param name="sStaffid"></param>
        ''' <param name="sFlowName"></param>
        ''' <param name="sStepNo"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetBraDepNoByUseridFlownameStepno(ByVal sStaffid As String,
                                                          ByVal sFlowName As String,
                                                          ByVal sStepNo As String) As Integer

            Dim obj As Object

            Try
                obj = ExecuteScalar( _
                        "with RC( ROLEID, ROLENAME, DISABLED, ROLETYPE, PARENT, FLOW_ID, STEP_NO) as " & vbCrLf & _
                        " (select RO.ROLEID, RO.ROLENAME, RO.DISABLED, RO.ROLETYPE, RO.PARENT, RRF.FLOW_ID, RRF.STEP_NO " & vbCrLf & _
                        "    from SY_ROLE RO " & vbCrLf & _
                        "   inner join SY_REL_ROLE_FLOWMAP RRF " & vbCrLf & _
                        "      on RO.ROLEID = RRF.ROLEID " & vbCrLf & _
                        "  union all " & vbCrLf & _
                        "  select A.ROLEID, A.ROLENAME, A.DISABLED, A.ROLETYPE, A.PARENT, RC.FLOW_ID, RC.STEP_NO " & vbCrLf & _
                        "    from SY_ROLE AS A, RC " & vbCrLf & _
                        "   where A.PARENT = RC.ROLEID " & vbCrLf & _
                        "     and RC.DISABLED = '0' " & vbCrLf & _
                        "     and A.DISABLED = '0') " & vbCrLf & _
                        "select distinct BRA_DEPNO " & vbCrLf & _
                        "  from RC " & vbCrLf & _
                        " inner join SY_REL_ROLE_USER RRU " & vbCrLf & _
                        "    on RRU.ROLEID = RC.ROLEID " & vbCrLf & _
                        " inner join SY_FLOW_ID FI " & vbCrLf & _
                        "    on FI.FLOW_ID = RC.FLOW_ID " & vbCrLf & _
                        " where STAFFID = @STAFFID@ and STEP_NO = @STEP_NO@ and FLOW_NAME = @FLOW_NAME@ " & vbCrLf,
                     "STAFFID", sStaffid,
                     "STEP_NO", sStepNo,
                     "FLOW_NAME", sFlowName)

                If IsDBNull(obj) OrElse IsNothing(obj) Then
                    Throw New SYException(String.Format(
                                          "流程步驟未設定關連的角色(未加入角色資料至SY_REL_ROLE_FLOWMAP)，" & vbCrLf & _
                                          "STAFFID={0}，STEP_NO={1}，FLOW_NAME={2}，SQL={3}",
                                          sStaffid, sStepNo, sFlowName, GetLastSQL),
                                      SYMSG.SYFLOWDEF_BRADEPNO_NOT_FOUND, GetLastSQL)
                End If

                Return CDbType(Of Integer)(obj, 0)

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Function


        ''' <summary>
        ''' 由FlowName, StepNo及PN取得下一步驟的步驟名稱
        ''' </summary>
        ''' <param name="sFlowName"></param>
        ''' <param name="sStepNo"></param>
        ''' <param name="sPN"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetNextStepByFlownameStepnoPn(ByVal sFlowName As String,
                                                      ByVal sStepNo As String,
                                                      sPN As String) As DataRowCollection
            Try
                Return GetDataRowCollection( _
                        "select SN.* " & vbCrLf & _
                        "  from SY_FLOW_DEF FD " & vbCrLf & _
                        " inner join SY_FLOW_ID FI " & vbCrLf & _
                        "    on FD.FLOW_ID = FI.FLOW_ID " & vbCrLf & _
                        " inner join SY_STEP_NO SN " & vbCrLf & _
                        "    on FD.NEXT_STEP_NO = SN.STEP_NO " & vbCrLf & _
                        " where FLOW_NAME = @FLOW_NAME@ " & vbCrLf & _
                        "   and CURR_STEP_NO = @CURR_STEP_NO@ " & vbCrLf & _
                        "   and PN = @PN@ " & vbCrLf,
                        "FLOW_NAME", sFlowName,
                        "CURR_STEP_NO", sStepNo,
                        "PN", sPN)
            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try
        End Function

    End Class

End Namespace

