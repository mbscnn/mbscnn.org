Option Explicit On
Option Strict On

Imports com.Azion.NET.VB

Namespace TABLE

    Public Class SY_FLOWINCIDENT
        Inherits SY_TABLEBASE

        ''' <summary>
        ''' Constructor
        ''' </summary>
        ''' <param name="dbManager">DatabaseManager</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal dbManager As com.Azion.NET.VB.DatabaseManager)
            MyBase.new("SY_FLOWINCIDENT", dbManager)
        End Sub


        ''' <summary>
        ''' 開始並初始一FlowIncident
        ''' </summary>
        ''' <param name="stepInfoItemExt"></param>
        ''' <remarks></remarks>
        Public Sub StartNewFlowIncident(ByVal stepInfoItemExt As StepInfoItemExt)

            Dim nResult As Integer
            Dim arrayBosParameters As New BosParamsList

            Try
                arrayBosParameters.Add("CASEID", stepInfoItemExt.caseId)
                arrayBosParameters.Add("SUBFLOW_SEQ", stepInfoItemExt.subflowSeq)
                arrayBosParameters.Add("STARTTIME", Now)
                arrayBosParameters.Add("ENDTIME", Nothing)
                arrayBosParameters.Add("STATUS", 1)

                If IsNothing(stepInfoItemExt.caseSender) = False AndAlso
                    String.IsNullOrEmpty(stepInfoItemExt.caseSender.userId) = False Then
                    arrayBosParameters.Add("LASTUPDATEUSER", stepInfoItemExt.caseSender.userId)
                End If

                arrayBosParameters.Add("LASTUPDATETIME", Now)
                'arrayBosParameters.Add(PARAMETER("FLOW_ID", stepInfoItemExt.flowId))

                If stepInfoItemExt.previousSubFlowSeq > 0 Then
                    arrayBosParameters.Add("PREV_SUBFLOWSEQ", stepInfoItemExt.previousSubFlowSeq)
                End If

                If stepInfoItemExt.subflowLevel > 0 Then
                    arrayBosParameters.Add("SUBFLOW_LEVEL", stepInfoItemExt.subflowLevel)
                End If

                If stepInfoItemExt.parallelNo > 0 Then
                    arrayBosParameters.Add("PARALLEL_NO", stepInfoItemExt.parallelNo)
                End If

                If stepInfoItemExt.parentSubFlowSeq > 0 Then
                    arrayBosParameters.Add("PARENT_SEQ", stepInfoItemExt.parentSubFlowSeq)
                End If


                nResult = Insert(arrayBosParameters.ToArray())
                Dbg.Assert(nResult = 1)
            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Sub


        Public Function GetLatestSubflowInfo(ByVal sCaseid As String, ByVal nSubflowseq As Integer) As DataRow
            Dim sSql As String

            Try
                sSql = _
                "select top 1 SY_FLOWINCIDENT.* " & vbCrLf & _
                "  from SY_FLOWINCIDENT " & vbCrLf & _
                " inner join (select * from SY_FLOWINCIDENT) PARENTFLOW " & vbCrLf & _
                "    on SY_FLOWINCIDENT.CASEID = PARENTFLOW.CASEID " & vbCrLf & _
                "   and SY_FLOWINCIDENT.PARENT_SEQ = PARENTFLOW.SUBFLOW_SEQ " & vbCrLf & _
                " where SY_FLOWINCIDENT.CASEID = @CASEID@ " & vbCrLf & _
                "   and PARENTFLOW.SUBFLOW_SEQ = @SUBFLOW_SEQ@ " & vbCrLf & _
                " order by SUBFLOW_SEQ DESC"

                Return GetDataRow(sSql,
                                  "CASEID", sCaseid,
                                  "SUBFLOW_SEQ", nSubflowseq)

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Function


        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="stepInfoItemExt"></param>
        ''' <remarks></remarks>
        Public Sub CloseFlowIncident(ByVal stepInfoItemExt As StepInfoItemExt)

            Dim nResult As Integer
            Dim arrayBosParameters As New BosParamsList

            Try
                arrayBosParameters.Add("CASEID", stepInfoItemExt.caseId)
                arrayBosParameters.Add("SUBFLOW_SEQ", stepInfoItemExt.subflowSeq)
                arrayBosParameters.Add("ENDTIME", Now)
                arrayBosParameters.Add("STATUS", 3)

                If IsNothing(stepInfoItemExt.caseSender) = False AndAlso
                    String.IsNullOrEmpty(stepInfoItemExt.caseSender.userId) = False Then
                    arrayBosParameters.Add("LASTUPDATEUSER", stepInfoItemExt.caseSender.userId)
                End If

                arrayBosParameters.Add("LASTUPDATETIME", Now)
                'arrayBosParameters.Add("FLOW_ID", stepInfoItemExt.flowId)

                If stepInfoItemExt.previousSubFlowSeq > 0 Then
                    arrayBosParameters.Add("PREV_SUBFLOWSEQ", stepInfoItemExt.previousSubFlowSeq)
                End If

                If stepInfoItemExt.subflowLevel > 0 Then
                    arrayBosParameters.Add("SUBFLOW_LEVEL", stepInfoItemExt.subflowLevel)
                End If

                If stepInfoItemExt.parallelNo > 0 Then
                    arrayBosParameters.Add("PARALLEL_NO", stepInfoItemExt.parallelNo)
                End If

                If stepInfoItemExt.parentSubFlowSeq > 0 Then
                    arrayBosParameters.Add("PARENT_SEQ", stepInfoItemExt.parentSubFlowSeq)
                End If


                nResult = Update(arrayBosParameters.ToArray())
                Dbg.Assert(nResult = 1)

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Sub

        ''' <summary>
        ''' 關閉所有的子流程
        ''' </summary>
        ''' <param name="stepInfoItemExt"></param>
        ''' <remarks></remarks>
        Public Sub CloseAllSubFlowIncident(ByVal stepInfoItemExt As StepInfoItemExt,
                                           Optional ByVal bIncludeParent As Boolean = False,
                                           Optional ByVal bUpdateSenderIfNothing As Boolean = True)

            Dim drc As DataRowCollection = Nothing
            Dim sSql As String = Nothing


            Dim loginUserId, workingUserId As USER_ID_NAME
            loginUserId = Nothing
            workingUserId = Nothing


            Try
                sSql = _
                    "with n(CASEID, SUBFLOW_SEQ, PARENT_SEQ, STATUS) as " & vbCrLf & _
                    " (select CASEID, SUBFLOW_SEQ, PARENT_SEQ, STATUS " & vbCrLf & _
                    "    from SY_FLOWINCIDENT " & vbCrLf & _
                    "   where CASEID = @CASEID@ " & vbCrLf & _
                    "     and SUBFLOW_SEQ = @SUBFLOW_SEQ@ " & vbCrLf & _
                    "   union all " & vbCrLf & _
                    "  select a.CASEID, a.SUBFLOW_SEQ, a.PARENT_SEQ, a.STATUS " & vbCrLf & _
                    "    from SY_FLOWINCIDENT as a, n " & vbCrLf & _
                    "   where ( n.SUBFLOW_SEQ = A.PARENT_SEQ or n.SUBFLOW_SEQ = A.PREV_SUBFLOWSEQ ) " & vbCrLf & _
                    "     and a.CASEID = @CASEID@) " & vbCrLf

                If bIncludeParent Then
                    sSql &= "select * from n " & vbCrLf
                Else
                    sSql &= "select * from n where n.SUBFLOW_SEQ <> @SUBFLOW_SEQ@ " & vbCrLf
                End If



                drc = GetDataRowCollection(sSql,
                                           "CASEID", stepInfoItemExt.caseId,
                                           "SUBFLOW_SEQ", stepInfoItemExt.subflowSeq)

                If bUpdateSenderIfNothing Then
                    FLOW_OP.FlowFacade.getNewInstance(getDatabaseManager).getELoanFlow().GetCurrentUserid(loginUserId, workingUserId)
                End If

                If Not IsNothing(drc) Then
                    For Each dr As DataRow In drc

                        Dim subFlowStepInfoItemExt As StepInfoItemExt
                        '關閉子流程中的所有子流程
                        subFlowStepInfoItemExt = stepInfoItemExt.clone
                        subFlowStepInfoItemExt.subflowSeq = CInt(dr("SUBFLOW_SEQ"))
                        'CloseAllSubFlowIncident(subFlowStepInfoItemExt)

                        '寫入3至SY_FLOWINCIDENT
                        Update("CASEID", stepInfoItemExt.caseId,
                               "SUBFLOW_SEQ", dr("SUBFLOW_SEQ"),
                               "STATUS", 3)

                        '寫入時間至SY_FLOWSTEP
                        getSYFlowStep.ExecuteNonQuery(
                            "update SY_FLOWSTEP" & vbCrLf & _
                            "   set PROCESSTIME = @PROCESSTIME@, " & vbCrLf & _
                            "            STATUS = 3 " & vbCrLf & _
                            " where CASEID = @CASEID@" & vbCrLf & _
                            "   and SUBFLOW_SEQ = @SUBFLOW_SEQ@" & vbCrLf & _
                            "   and PROCESSTIME is NULL",
                            "PROCESSTIME", Now,
                            "CASEID", stepInfoItemExt.caseId,
                            "SUBFLOW_SEQ", dr("SUBFLOW_SEQ"))

                        '寫入SENDER至SY_FLOWSTEP
                        If bUpdateSenderIfNothing = True Then
                            getSYFlowStep.ExecuteNonQuery(
                                "update SY_FLOWSTEP" & vbCrLf & _
                                "   set SENDER = @SENDER@ " & vbCrLf & _
                                " where CASEID = @CASEID@" & vbCrLf & _
                                "   and SUBFLOW_SEQ = @SUBFLOW_SEQ@" & vbCrLf & _
                                "   and SENDER is NULL",
                                "SENDER", loginUserId.userId,
                                "CASEID", stepInfoItemExt.caseId,
                                "SUBFLOW_SEQ", dr("SUBFLOW_SEQ"))
                        End If

                    Next
                End If

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Sub


        ''' <summary>
        ''' 取得流程案件內最大的SubFlowSeq
        ''' </summary>
        ''' <param name="strCaseid"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetMaxSubFlowSeq(ByVal strCaseid As String) As Integer
            Try
                Dim obj As Object

                obj = ExecuteScalar(
                    "select MAX(SUBFLOW_SEQ) as SUBFLOW_SEQ " & vbCrLf & _
                    "  from SY_FLOWINCIDENT " & vbCrLf & _
                    " where CASEID = @CASEID@",
                    "CASEID", strCaseid)

                If IsDBNull(obj) Then
                    Throw New SYException(
                        String.Format("無法從SY_FLOWINCIDENT.CASEID:{0}取得SUBFLOW_SEQ的最大值", strCaseid),
                        SYMSG.SYFLOWINCIDENT_CANNOT_GET_MAX_SUBFLOWSEQ,
                        GetLastSQL)
                End If

                Return CInt(obj)

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try
        End Function


        Public Function GetMaxSubFlowCount(ByVal strCaseid As String) As Integer
            Try
                Dim obj As Object

                obj = ExecuteScalar(
                    "select MAX(SUBFLOW_COUNT) as SUBFLOW_COUNT " & vbCrLf & _
                    "  from SY_FLOWSTEP " & vbCrLf & _
                    " where CASEID = @CASEID@",
                    "CASEID", strCaseid)

                If IsDBNull(obj) Then
                    Throw New SYException(
                        String.Format("無法從SY_FLOWSTEP.CASEID:{0}取得SUBFLOW_COUNT的最大值", strCaseid),
                        SYMSG.SYFLOWINCIDENT_CANNOT_GET_MAX_SUBFLOWCOUNT,
                        GetLastSQL)
                End If

                Return CInt(obj)

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try
        End Function


        ''' <summary>
        ''' 流程案件是否已結束
        ''' </summary>
        ''' <param name="sCaseid"></param>
        ''' <param name="nSubflowSeq"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function IsCaseCompleted(ByVal sCaseid As String, Optional ByVal nSubflowSeq As Integer = 0) As Boolean
            Try
                Dim obj As Object
                Dim nCount As Integer = 0

                If nSubflowSeq = 0 Then
                    obj = ExecuteScalar(
                        "select count(*) " & vbCrLf & _
                        "  from SY_FLOWINCIDENT " & vbCrLf & _
                        " where CASEID = @CASEID@ " & vbCrLf & _
                        "   and STATUS <> 3",
                        "CASEID", sCaseid)
                Else
                    obj = ExecuteScalar(
                        "select count(*) " & vbCrLf & _
                        "  from SY_FLOWINCIDENT " & vbCrLf & _
                        " where CASEID = @CASEID@ " & vbCrLf & _
                        "   and SUBFLOW_SEQ = @SUBFLOW_SEQ@ " & vbCrLf & _
                        "   and STATUS <> 3",
                        "CASEID", sCaseid,
                        "SUBFLOW_SEQ", nSubflowSeq)
                End If

                nCount = CInt(obj)

                If nCount = 0 Then
                    Return True
                End If

                Return False
            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Function


        ''' <summary>
        ''' 取得未完成的子流程數目
        ''' </summary>
        ''' <param name="sCaseid"></param>
        ''' <param name="nSubFlowSeq"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetInProgressSubFlowCount(ByVal sCaseid As String, ByVal nSubFlowSeq As Integer) As Integer

            Dim sSql As String = Nothing
            Dim obj As Object = Nothing

            Try

                sSql =
                    "with n(CASEID, SUBFLOW_SEQ, PARENT_SEQ, STATUS) as " & vbCrLf & _
                    " (select CASEID, SUBFLOW_SEQ, PARENT_SEQ, STATUS " & vbCrLf & _
                    "    from SY_FLOWINCIDENT " & vbCrLf & _
                    "   where CASEID = @CASEID@ " & vbCrLf & _
                    "     and SUBFLOW_SEQ = @SUBFLOW_SEQ@ " & vbCrLf & _
                    "   union all " & vbCrLf & _
                    "  select a.CASEID, " & vbCrLf & _
                    "         a.SUBFLOW_SEQ, " & vbCrLf & _
                    "         a.PARENT_SEQ, " & vbCrLf & _
                    "         a.STATUS " & vbCrLf & _
                    "    from SY_FLOWINCIDENT as a, n " & vbCrLf & _
                    "   where a.PARENT_SEQ = n.SUBFLOW_SEQ " & vbCrLf & _
                    "     and a.CASEID = @CASEID@) " & vbCrLf & _
                    "select COUNT(*) from n where SUBFLOW_SEQ <> @SUBFLOW_SEQ@ and STATUS <> 3 "

                obj = ExecuteScalar(sSql,
                                    "CASEID", sCaseid,
                                    "SUBFLOW_SEQ", nSubFlowSeq)

                Return CInt(obj)

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try
        End Function


        ''' <summary>
        ''' 設定案件的狀態
        ''' </summary>
        ''' <param name="sCaseid"></param>
        ''' <param name="nSubFlowSeq"></param>
        ''' <param name="sStatus">1 or 3</param>
        ''' <remarks></remarks>
        Public Sub SetCaseStatus(ByVal sCaseid As String, nSubFlowSeq As Integer, ByVal sStatus As String)
            Try
                Dbg.Assert(sStatus = "1" OrElse sStatus = "3", "只能設為1或3")

                If sStatus <> "1" AndAlso sStatus <> "3" Then
                    Throw New SYException(
                        "[SY_FLOWINCIDENT].[STATUS]只能設為1或3",
                        SYMSG.SYFLOWINCIDENT_SET_WRONG_STATUS,
                        GetLastSQL)
                End If


#If DEBUG Then
                Try
                    If sStatus = "1" Then
                        Dim sDbgValue As String
                        sDbgValue = QueryString("[SY_FLOWINCIDENT].[STATUS]", "CASEID", sCaseid,
                                                "SUBFLOW_SEQ", nSubFlowSeq)

                        Dbg.Assert(sDbgValue <> "3", "案件已關閉，不應該被更改", True)
                    End If
                Catch ex As Exception
                End Try
#End If

                InsertUpdate("CASEID", sCaseid,
                             "SUBFLOW_SEQ", nSubFlowSeq,
                             "STATUS", sStatus)
            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try
        End Sub



        ''' <summary>
        ''' 取得目前子流程編號的PARENT_SEQ
        ''' </summary>
        ''' <param name="sCaseid"></param>
        ''' <param name="nCurrentSubFlowSeq"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetParentSubFlowSeq(ByVal sCaseid As String, nCurrentSubFlowSeq As Integer) As Integer

            Dim obj As Object

            Try
                obj = ExecuteScalar("select PARENT_SEQ " & vbCrLf & _
                                     "  from SY_FLOWINCIDENT " & vbCrLf & _
                                     " where CASEID = @CASEID@ " & vbCrLf & _
                                     "   and SUBFLOW_LEVEL = @SUBFLOW_LEVEL@ " & vbCrLf,
                                     "CASEID", sCaseid,
                                     "SUBFLOW_LEVEL", nCurrentSubFlowSeq)

                If IsDBNull(obj) OrElse IsNothing(obj) Then
                    Throw New SYException(
                        String.Format("無法取得PARENT_SEQ的內容：CASEID={0}，SUBFLOW_SEQ={1}", sCaseid, nCurrentSubFlowSeq),
                        SYMSG.SYFLOWINCIDENT_CANNOT_GET_MAX_PARENTSEQ,
                        GetLastSQL)
                End If

                Return CInt(obj)

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Function

    End Class

End Namespace
