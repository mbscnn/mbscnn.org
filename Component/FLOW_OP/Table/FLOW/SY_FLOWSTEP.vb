Option Explicit On
Option Strict On

Imports com.Azion.NET.VB
Imports FLOW_OP.TABLE
Imports System.Text

Namespace TABLE

    Public Class SY_FLOWSTEP
        Inherits SY_TABLEBASE


        ''' <summary>
        ''' Constructor
        ''' </summary>
        ''' <param name="dbManager">DatabaseManager</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal dbManager As com.Azion.NET.VB.DatabaseManager)
            MyBase.new("SY_FLOWSTEP", dbManager)
        End Sub


        Public Shared Function getNewInstance(ByVal dbManager As com.Azion.NET.VB.DatabaseManager) As SY_FLOWSTEP
            Return New SY_FLOWSTEP(dbManager)
        End Function


        ''' <summary>
        ''' 加入或更新一FlowStep記錄，
        ''' </summary>
        ''' <param name="stepInfoItemExt"></param>
        ''' <remarks></remarks>
        Public Sub InsertUpdateRecord(ByVal stepInfoItemExt As StepInfoItemExt, Optional ByVal sType_IU As String = Nothing)

            Dim nResult As Integer
            Dim nCount As Integer
            Dim arrayBosParameters As New BosParamsList

            Try

                If String.IsNullOrEmpty(sType_IU) OrElse (sType_IU <> "I" AndAlso sType_IU <> "U") Then
                    nCount = GetCount(BosBase.PARAM_ARRAY("CASEID", stepInfoItemExt.caseId,
                                      "SUBFLOW_SEQ", stepInfoItemExt.subflowSeq,
                                      "SUBFLOW_COUNT", stepInfoItemExt.subflowCount,
                                      "STEP_NO", stepInfoItemExt.stepNo))
                    If nCount = 0 Then
                        sType_IU = "I"
                    Else
                        sType_IU = "U"
                    End If
                End If

                If String.IsNullOrEmpty(stepInfoItemExt.stepSummary) Then
                    stepInfoItemExt.stepSummary = GetSummary(stepInfoItemExt.braDepNo_brid(0).BraDepNo,
                                                             stepInfoItemExt.flowId,
                                                             stepInfoItemExt.stepNo,
                                                             stepInfoItemExt.caseOwnerList)
                End If

                Dbg.Assert(Not String.IsNullOrEmpty(stepInfoItemExt.stepSummary))

                arrayBosParameters.Add("CASEID", stepInfoItemExt.caseId)
                arrayBosParameters.Add("SUBFLOW_SEQ", stepInfoItemExt.subflowSeq)
                arrayBosParameters.Add("SUBFLOW_COUNT", stepInfoItemExt.subflowCount)
                arrayBosParameters.Add("STEP_NO", stepInfoItemExt.stepNo)
                arrayBosParameters.Add("SUMMARY", stepInfoItemExt.stepSummary)
                arrayBosParameters.Add("STATUS", stepInfoItemExt.flowstepStatus)
                arrayBosParameters.Add("BRA_DEPNO", stepInfoItemExt.braDepNo_brid(0).BraDepNo)

                If IsNothing(stepInfoItemExt.caseOwner) = False AndAlso
                    String.IsNullOrEmpty(stepInfoItemExt.caseOwner.userId) = False Then
                    arrayBosParameters.Add("OWNER", stepInfoItemExt.caseOwner.userId)
                End If

                If IsNothing(stepInfoItemExt.caseClient) = False AndAlso
                    String.IsNullOrEmpty(stepInfoItemExt.caseClient.userId) = False Then
                    arrayBosParameters.Add("CLIENT", stepInfoItemExt.caseClient.userId)
                End If

                If IsNothing(stepInfoItemExt.caseSender) = False AndAlso
                    String.IsNullOrEmpty(stepInfoItemExt.caseSender.userId) = False Then
                    arrayBosParameters.Add("SENDER", stepInfoItemExt.caseSender.userId)
                End If

                If stepInfoItemExt.processTime <> DateTime.MinValue Then
                    arrayBosParameters.Add("PROCESSTIME", stepInfoItemExt.processTime)
                End If

                If stepInfoItemExt.revisionSeqNo <> 0 Then
                    arrayBosParameters.Add("REVISION_SEQNO", stepInfoItemExt.revisionSeqNo)
                End If


                If sType_IU = "I" Then
                    arrayBosParameters.Add("STARTTIME", Now)
                    nResult = Insert(arrayBosParameters.ToArray())
                Else
                    nResult = Update(arrayBosParameters.ToArray())
                End If


            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

            Dbg.Assert(nResult = 1)

        End Sub



        ''' <summary>
        ''' 輸入分行，流程代號, STEPNO取得SUMMARY
        ''' </summary>
        ''' <param name="nBraDepNo"></param>
        ''' <param name="nFlowId"></param>
        ''' <param name="sStepNo"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetSummary(ByVal nBraDepNo As Integer,
                                   ByVal nFlowId As Integer,
                                   ByVal sStepNo As String,
                                   Optional ByVal listOwner As List(Of USER_ID_NAME) = Nothing) As String

            Dim sSql As String
            Dim obj As Object
            Dim sb As StringBuilder

            Try
                sSql = _
                "select SY_FLOW_ID.FLOW_CNAME + '[' + SY_STEP_NO.STEP_NAME + '][<USERLIST>]<+>' + " & vbCrLf & _
                "       SY_FLOW_ID.FLOW_CNAME + '<+>' + SY_BRANCH.BRCNAME + '<+>' + " & vbCrLf & _
                "       SY_FLOW_ID.FLOW_NAME " & vbCrLf & _
                "  from SY_FLOW_ID " & vbCrLf & _
                "  left join SY_STEP_NO" & vbCrLf & _
                "    on SY_STEP_NO.STEP_NO = @STEP_NO@" & vbCrLf & _
                "  left join SY_BRANCH " & vbCrLf & _
                "    on SY_BRANCH.BRA_DEPNO = @BRA_DEPNO@ " & vbCrLf & _
                " where SY_FLOW_ID.FLOW_ID = @FLOW_ID@ " & vbCrLf & _
                "   and SY_STEP_NO.STEP_NO = @STEP_NO@"

                obj = ExecuteScalar(sSql,
                                    "BRA_DEPNO", nBraDepNo,
                                    "FLOW_ID", nFlowId,
                                    "STEP_NO", sStepNo
                                    )

                sSql = CDbType(Of String)(obj)

                If IsNothing(listOwner) Then
                    Return sSql.Replace("[<USERLIST>]", "")
                End If

                sb = New StringBuilder()

                For Each item As USER_ID_NAME In listOwner
                    If sSql.Length + sb.Length + item.userId.Length < 200 Then
                        If sb.Length = 0 Then
                            sb.Append(item.userId)
                        Else
                            sb.Append("," & item.userId)
                        End If
                    Else
                        Exit For
                    End If
                Next

                Return sSql.Replace("[<USERLIST>]", "[" & sb.ToString & "]")

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Function



        ''' <summary>
        ''' 取得流程步驟的資訊
        ''' </summary>
        ''' <param name="sCaseid"></param>
        ''' <param name="nSubflowSeq"></param>
        ''' <param name="sStepNo"></param>
        ''' <param name="sCurrentUser"></param>
        ''' <param name="bInProgrssOnly"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetStepInfoItem(ByVal sCaseid As String,
                                        Optional ByVal nSubflowSeq As Integer = 0,
                                        Optional ByVal sStepNo As String = Nothing,
                                        Optional ByVal sCurrentUser As String = Nothing,
                                        Optional ByVal bInProgrssOnly As Boolean = True) As StepInfoItemExt
            Try
                Dim arrayBosParameters As New BosParamsList
                Dim drc As DataRowCollection
                Dim dr As DataRow = Nothing
                Dim sbSql As StringBuilder

                arrayBosParameters.Add("CASEID", sCaseid)

                sbSql = New StringBuilder(
                    "Select FS.*," & vbCrLf & _
                    "       FI.PREV_SUBFLOWSEQ," & vbCrLf & _
                    "       FI.SUBFLOW_LEVEL," & vbCrLf & _
                    "       FI.PARALLEL_NO," & vbCrLf & _
                    "       FI.PARENT_SEQ," & vbCrLf & _
                    "       CI.FLOW_ID, " & vbCrLf & _
                    "       FID.FLOW_NAME,  " & vbCrLf & _
                    "       FM.TYPE, " & vbCrLf & _
                    "       BR.BRID " & vbCrLf & _
                    "  from SY_FLOWSTEP FS" & vbCrLf & _
                    " inner join SY_FLOWINCIDENT FI" & vbCrLf & _
                    "    on FS.CASEID = FI.CASEID" & vbCrLf & _
                    "   and FS.SUBFLOW_SEQ = FI.SUBFLOW_SEQ" & vbCrLf & _
                    " inner join SY_CASEID CI" & vbCrLf & _
                    "    on FS.CASEID = CI.CASEID" & vbCrLf & _
                    " inner join SY_FLOW_MAP FM" & vbCrLf & _
                    "    on FM.FLOW_ID = CI.FLOW_ID" & vbCrLf & _
                    "   and FM.STEP_NO = FS.STEP_NO" & vbCrLf & _
                    "  left join SY_FLOW_ID FID" & vbCrLf & _
                    "    on FID.FLOW_ID = CI.FLOW_ID" & vbCrLf & _
                    "  left join SY_BRANCH BR" & vbCrLf & _
                    "    on BR.BRA_DEPNO = FS.BRA_DEPNO" & vbCrLf & _
                    " where FS.CASEID = @CASEID@" & vbCrLf)

                If bInProgrssOnly = True Then
                    sbSql.Append("   and FS.PROCESSTIME is NULL " & vbCrLf)
                End If

                If Not String.IsNullOrEmpty(sStepNo) Then
                    sbSql.Append("   and FS.STEP_NO=@STEP_NO@ " & vbCrLf)
                    arrayBosParameters.Add("STEP_NO", sStepNo)
                End If

                If nSubflowSeq <> 0 Then
                    sbSql.Append("   and FS.SUBFLOW_SEQ=@SUBFLOW_SEQ@ " & vbCrLf)
                    arrayBosParameters.Add("SUBFLOW_SEQ", nSubflowSeq)
                End If

                drc = GetDataRowCollection(sbSql.ToString(), arrayBosParameters.ToArray())
                If IsNothing(drc) OrElse drc.Count = 0 Then
                    Throw New SYException(
                        String.Format("無法從CASEID:{0}及CLIENT:{1}取得子流程編號", sCaseid, "" & sCurrentUser),
                        SYMSG.SYFLOWSTEP_SUBFLOWSEQ_NOT_FOUND,
                        GetLastSQL)
                End If

                If IsNothing(drc) = False AndAlso drc.Count > 1 Then
                    If Not String.IsNullOrEmpty(sCurrentUser) Then
                        For Each dr2 As DataRow In drc
                            If CDbType(Of String)(dr2("CLIENT")) = sCurrentUser Then
                                dr = dr2
                                Exit For
                            End If
                        Next
                    ElseIf nSubflowSeq <> 0 Then
                        For Each dr2 As DataRow In drc
                            If IsNothing(dr) = True OrElse CInt(dr2("SUBFLOW_COUNT")) > CInt(dr("SUBFLOW_COUNT")) Then
                                dr = dr2
                            End If
                        Next
                    End If

                    If IsNothing(dr) Then
                        Throw New SYException(
                            String.Format("無法從CASEID:{0}判斷子流程編號(子流程編號數目:{1} (子流程編號數目超過1))", sCaseid, drc.Count),
                        SYMSG.SYFLOWSTEP_SUBFLOWSEQ_NOT_FOUND_AMBIGUOUS,
                        GetLastSQL)
                    End If
                Else
                    dr = drc(0)
                End If

                Dim sii As New StepInfoItemExt
                sii.caseId = CDbType(Of String)(dr.Item("CASEID"))

                sii.braDepNo_brid = SY_TABLEBASE.ToArray(BRADEPNO_BRID.Pair(
                    CInt(dr.Item("BRA_DEPNO")),
                    CDbType(Of String)(dr.Item("BRID"))))

                sii.caseClient =
                    getSYUser.GetUserIdName(CDbType(Of String)(dr.Item("CLIENT")))
                sii.caseOwner =
                    getSYUser.GetUserIdName(CDbType(Of String)(dr.Item("OWNER")))
                sii.caseSender =
                    getSYUser.GetUserIdName(CDbType(Of String)(dr.Item("SENDER")))
                'If(IsDBNull(dr).Item("SENDER")), Nothing, CStr(dr.Item("SENDER")))
                sii.flowId = CInt(dr.Item("FLOW_ID"))
                sii.flowName = CDbType(Of String)(dr.Item("FLOW_NAME"))
                sii.flowstepStatus = CInt(dr.Item("STATUS"))
                sii.stepNo = CDbType(Of String)(dr.Item("STEP_NO"))
                sii.subflowCount = CInt(dr.Item("SUBFLOW_COUNT"))
                sii.subflowLevel = CInt(dr.Item("SUBFLOW_LEVEL"))
                sii.subflowSeq = CInt(dr.Item("SUBFLOW_SEQ"))
                sii.stepSummary = CDbType(Of String)(dr.Item("SUMMARY"))
                sii.stepType = CDbType(Of String)(dr.Item("TYPE"), Nothing)
                sii.startTime = CDate(dr.Item("STARTTIME"))

                sii.parentSubFlowSeq = CDbType(Of Integer)(dr.Item("PARENT_SEQ"), 0)
                sii.previousSubFlowSeq = CDbType(Of Integer)(dr.Item("PREV_SUBFLOWSEQ"), 0)
                sii.parallelNo = CDbType(Of Integer)(dr.Item("PARALLEL_NO"), 1)

                Return sii
            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Function



        ''' <summary>
        ''' 取得流程步驟的資訊
        ''' </summary>
        ''' <param name="sCaseId"></param>
        ''' <param name="nSubFlowSeq"></param>
        ''' <param name="nStatus"></param>
        ''' <param name="sStepNo"></param>
        ''' <param name="sTopN"></param>
        ''' <param name="sOrderBy"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetFlowStepInfo(ByVal sCaseId As String,
                              Optional ByVal nSubFlowSeq As Integer = 0,
                              Optional ByVal sStepNo As String = Nothing,
                              Optional ByVal nStatus As Integer = 0,
                              Optional ByVal sTopN As String = Nothing,
                              Optional ByVal sOrderBy As String = Nothing) As StepInfoItemExt()
            Dim sii As StepInfoItemExt = Nothing
            Dim arrayStepInfoItemExt As New List(Of StepInfoItemExt)
            Dim arrayBosParameters As New BosParamsList
            Dim sbSQL As StringBuilder
            Dim drc As DataRowCollection
            'Dim dt As DataTable

            If String.IsNullOrEmpty(sTopN) Then
                sTopN = String.Empty
            Else
                sTopN = " " & sTopN.Trim & " "
            End If

            Try
                sbSQL = New StringBuilder(
                    "select " & sTopN & "FS.*, BR.BRID, FID.FLOW_ID, FID.FLOW_NAME, FI.SUBFLOW_LEVEL, FI.PARALLEL_NO, UR1.USERNAME as CLIENTNAME, UR2.USERNAME as OWNERNAME, UR3.USERNAME as SENDERNAME, FI.PARENT_SEQ, FI.PREV_SUBFLOWSEQ " & vbCrLf & _
                    "  from SY_FLOWSTEP FS " & vbCrLf & _
                    " inner join SY_FLOWINCIDENT FI " & vbCrLf & _
                    "    on FI.CASEID = FS.CASEID " & vbCrLf & _
                    "   and FI.SUBFLOW_SEQ = FS.SUBFLOW_SEQ " & vbCrLf & _
                    " inner join SY_CASEID CI " & vbCrLf & _
                    "    on CI.CASEID = FS.CASEID " & vbCrLf & _
                    " inner join SY_FLOW_ID FID " & vbCrLf & _
                    "    on FID.FLOW_ID = CI.FLOW_ID " & vbCrLf & _
                    "  left join SY_USER UR1 " & vbCrLf & _
                    "    on FS.CLIENT = UR1.STAFFID " & vbCrLf & _
                    "  left join SY_USER UR2 " & vbCrLf & _
                    "    on FS.OWNER = UR2.STAFFID " & vbCrLf & _
                    "  left join SY_USER UR3 " & vbCrLf & _
                    "    on FS.SENDER = UR3.STAFFID " & vbCrLf & _
                    "  left join SY_BRANCH BR " & vbCrLf & _
                    "    on BR.BRA_DEPNO = FS.BRA_DEPNO " & vbCrLf & _
                    " where FS.CASEID = @CASEID@ " & vbCrLf)

                arrayBosParameters.Add("CASEID", sCaseId)

                If nSubFlowSeq <> 0 Then
                    sbSQL.Append("   and FS.SUBFLOW_SEQ = @SUBFLOW_SEQ@ " & vbCrLf)
                    arrayBosParameters.Add("SUBFLOW_SEQ", nSubFlowSeq)
                End If

                If nStatus <> 0 Then
                    If nStatus = 3 Then
                        sbSQL.Append("   and FS.PROCESSTIME is not NULL " & vbCrLf)
                    Else
                        sbSQL.Append("   and FS.PROCESSTIME is NULL " & vbCrLf)
                    End If
                End If

                If Not String.IsNullOrEmpty(sStepNo) Then
                    sbSQL.Append("   and STEP_NO = @STEP_NO@ " & vbCrLf)
                    arrayBosParameters.Add("STEP_NO", sStepNo)
                End If

                If Not String.IsNullOrEmpty(sOrderBy) Then
                    sbSQL.Append(" order by " & sOrderBy & vbCrLf)
                End If

                drc = GetDataRowCollection(sbSQL.ToString, arrayBosParameters.ToArray)

                If Not IsNothing(drc) Then
                    For Each dr As DataRow In drc
                        sii = New StepInfoItemExt

                        sii.caseId = CStr(dr("CASEID"))
                        sii.stepNo = CStr(dr("STEP_NO"))
                        sii.subflowSeq = CInt(dr("SUBFLOW_SEQ"))
                        sii.subflowCount = CInt(dr("SUBFLOW_COUNT"))

                        sii.parentSubFlowSeq = CDbType(Of Integer)(dr("PARENT_SEQ"), 0)
                        sii.previousSubFlowSeq = CDbType(Of Integer)(dr("PREV_SUBFLOWSEQ"), 0)
                        sii.flowId = CInt(dr("FLOW_ID"))
                        sii.flowName = CStr(dr("FLOW_NAME"))
                        'sii.braDepNo = CInt(dr("BRA_DEPNO"))
                        sii.braDepNo_brid = SY_TABLEBASE.ToArray(BRADEPNO_BRID.Pair(CInt(dr("BRA_DEPNO")), CStr(dr("BRID"))))
                        'Dbg.Assert(False)
                        sii.caseOwner = USER_ID_NAME.USER(CDbType(Of String)(dr("OWNER"), String.Empty), CDbType(Of String)(dr("OWNERNAME"), String.Empty))
                        sii.caseClient = USER_ID_NAME.USER(CDbType(Of String)(dr("CLIENT"), String.Empty), CDbType(Of String)(dr("CLIENTNAME"), String.Empty))
                        sii.caseSender = USER_ID_NAME.USER(CDbType(Of String)(dr("SENDER"), String.Empty), CDbType(Of String)(dr("SENDERNAME"), String.Empty))
                        sii.subflowLevel = CInt(dr("SUBFLOW_LEVEL"))
                        sii.parallelNo = CInt(dr("PARALLEL_NO"))
                        sii.flowstepStatus = CInt(dr("STATUS"))
                        sii.stepSummary = CStr(dr("SUMMARY"))

                        arrayStepInfoItemExt.Add(sii)
                    Next
                End If

                Return arrayStepInfoItemExt.ToArray

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Function




        ''' <summary>
        ''' 取得目前正在進行中流程階段的StepInfoItem
        ''' </summary>
        ''' <param name="sCaseid"></param>
        ''' <param name="nSubflowSeq">可以為0，表示只會從Caseid取得SubFLowSeq</param>
        ''' <returns></returns>
        ''' <remarks>只能運用在同時只有一個步驟在流程裡進行(沒有兩個以上的步驟在進行),
        ''' 若多個子流程必須指定nSubflowSeq</remarks>
        Public Function GetInProgressStepInfoItem(ByVal sCaseid As String,
                                        Optional ByVal nSubflowSeq As Integer = 0,
                                        Optional ByVal sCurrentUser As String = Nothing) As StepInfoItemExt

            Return GetStepInfoItem(sCaseid, nSubflowSeq, Nothing, sCurrentUser, True)
        End Function


        ''' <summary>
        ''' 取得案件的已執行流程步驟的CLIENT
        ''' </summary>
        ''' <param name="sCaseid"></param>
        ''' <param name="sStepNo"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        'Public Function GetLatestClientByCaseidSubflowseqStepno(ByVal sCaseid As String, ByVal sStepNo As String) As String
        '    Return GetLatestClientByCaseidSubflowseqStepno(sCaseid, 0, sStepNo)
        'End Function

        ''' <summary>
        ''' 取得案件的已執行流程步驟的CLIENT
        ''' </summary>
        ''' <param name="sCaseid"></param>
        ''' <param name="nSubflowSeq">nSubflowSeq=0表示不指定nSubflowSeq</param>
        ''' <param name="sStepNo"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetLatestInfoByCaseidSubflowseqStepno(ByVal sCaseid As String,
                                                        ByVal nSubflowSeq As Integer,
                                                        ByVal sStepNo As String) As DataRowCollection

            Dim strSql As String

            Try
                strSql = _
                    "with n (CASEID , STARTTIME , ENDTIME , SUBFLOW_SEQ, PREV_SUBFLOWSEQ, PARENT_SEQ ) as" & vbCrLf & _
                    " (select CASEID, STARTTIME, ENDTIME, SUBFLOW_SEQ, PREV_SUBFLOWSEQ, PARENT_SEQ " & vbCrLf & _
                    "    from SY_FLOWINCIDENT " & vbCrLf & _
                    "   where CASEID = @CASEID@ " & vbCrLf

                If nSubflowSeq <> 0 Then
                    strSql &=
                    "     and SUBFLOW_SEQ = @SUBFLOW_SEQ@ " & vbCrLf
                End If

                strSql &= _
                    "   union all " & vbCrLf & _
                    "  select a.CASEID, a.STARTTIME, a.ENDTIME, a.SUBFLOW_SEQ, a.PREV_SUBFLOWSEQ, a.PARENT_SEQ " & vbCrLf & _
                    "    from SY_FLOWINCIDENT as a, n " & vbCrLf & _
                    "   where (a.SUBFLOW_SEQ = n.PREV_SUBFLOWSEQ and a.SUBFLOW_SEQ < n.SUBFLOW_SEQ and a.CASEID = n.CASEID) " & vbCrLf & _
                    "      or (a.SUBFLOW_SEQ = n.PARENT_SEQ and a.SUBFLOW_SEQ < n.SUBFLOW_SEQ and a.CASEID = n.CASEID) )" & vbCrLf & _
                    "select top 1 s.CLIENT, s.BRA_DEPNO from n  " & vbCrLf & _
                    " inner join SY_FLOWSTEP s " & vbCrLf & _
                    "    on n.CASEID = s.CASEID " & vbCrLf & _
                    "   and n.SUBFLOW_SEQ = s.SUBFLOW_SEQ " & vbCrLf & _
                    " where s.PROCESSTIME is not null" & vbCrLf & _
                    "   and s.STEP_NO = @STEP_NO@ " & vbCrLf & _
                    " order by s.SUBFLOW_SEQ desc "

                Dim arrayBosParameters As New BosParamsList
                arrayBosParameters.Add("CASEID", sCaseid)
                arrayBosParameters.Add("STEP_NO", sStepNo)

                If nSubflowSeq <> 0 Then
                    arrayBosParameters.Add("SUBFLOW_SEQ", nSubflowSeq)
                End If

                Return GetDataRowCollection2(strSql, arrayBosParameters.ToArray())

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Function



        ''' <summary>
        ''' 取得案件的第一個執行步驟
        ''' </summary>
        ''' <param name="sCaseid"></param>
        ''' <param name="nSubFlowSeq"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetStartStepNo(ByVal sCaseid As String, Optional ByVal nSubFlowSeq As Integer = 0) As String
            Dim sSql As String
            Dim arrayBosParameters As BosParamsList
            Dim obj As Object = Nothing

            Try
                arrayBosParameters = New BosParamsList
                arrayBosParameters.Add("CASEID", sCaseid)

                sSql = _
                    "select STEP_NO " & vbCrLf & _
                    "  from SY_FLOWSTEP " & vbCrLf & _
                    " where CASEID = @CASEID@ " & vbCrLf

                If nSubFlowSeq <> 0 Then
                    sSql &= "   and SUBFLOW_SEQ = @SUBFLOW_SEQ@ " & vbCrLf
                    arrayBosParameters.Add("SUBFLOW_SEQ", nSubFlowSeq)
                End If

                sSql &= " order by STARTTIME "

                obj = ExecuteScalar(sSql, arrayBosParameters.ToArray())

                If IsDBNull(obj) Then
                    Throw New SYException(
                        String.Format("無法取得CASEID:{0}及子流程代碼:{1}的最先一個步驟代碼", sCaseid, nSubFlowSeq),
                        SYMSG.SYFLOWSTEP_FIRSTSTEP_NOT_FOUND,
                        GetLastSQL)
                End If

                Return CStr(obj)

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Function




        ''' <summary>
        ''' 取得流程中最新一筆記錄
        ''' </summary>
        ''' <param name="sCaseid"></param>
        ''' <param name="nSubFlowSeq"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetLatestStepNo(ByVal sCaseid As String, Optional ByVal nSubFlowSeq As Integer = 0) As String
            Dim sSql As String
            Dim arrayBosParameters As BosParamsList
            Dim obj As Object = Nothing

            Try
                arrayBosParameters = New BosParamsList
                arrayBosParameters.Add("CASEID", sCaseid)

                sSql = _
                    "select top 1 STEP_NO " & vbCrLf & _
                    "  from (select STEP_NO, ISNULL(PROCESSTIME, STARTTIME) as LATESTDATE " & vbCrLf & _
                    "          from SY_FLOWSTEP " & vbCrLf & _
                    "         where CASEID = @CASEID@ " & vbCrLf

                If nSubFlowSeq <> 0 Then
                    sSql &= "         and SUBFLOW_SEQ = @SUBFLOW_SEQ@ " & vbCrLf
                    arrayBosParameters.Add("SUBFLOW_SEQ", nSubFlowSeq)
                End If

                sSql &= _
                    " ) VIEW1 " & vbCrLf & _
                    " order by VIEW1.LATESTDATE desc "

                obj = ExecuteScalar(sSql, arrayBosParameters.ToArray())

                If IsDBNull(obj) Then
                    Throw New SYException(
                        String.Format("無法取得CASEID:{0}及子流程代碼:{1}的最後一個步驟代碼", sCaseid, nSubFlowSeq),
                        SYMSG.SYFLOWSTEP_LASTSTEP_NOT_FOUND,
                        GetLastSQL)
                End If

                Return CStr(obj)

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Function



        ''' <summary>
        ''' 取得已記錄的流程中，SUBFLOW_SEQ及SUBFLOW_COUNT的最大值
        ''' </summary>
        ''' <param name="sCaseid"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetSubFlowInfoMax(ByVal sCaseid As String, Optional ByVal sStepNo As String = Nothing) As DataRow
            Dim sSql As String
            Dim arrayBosParameters As New BosParamsList

            Try
                sSql = _
                    "select MAX(SUBFLOW_SEQ) as MAX_SUBFLOWSEQ, " & vbCrLf & _
                    "       MAX(SUBFLOW_COUNT) as MAX_SUBFLOW_COUNT " & vbCrLf & _
                    "  from SY_FLOWSTEP " & vbCrLf & _
                    " where CASEID = @CASEID@ " & vbCrLf

                arrayBosParameters.Add("CASEID", sCaseid)

                If Not String.IsNullOrEmpty(sStepNo) Then
                    sSql &= " and STEP_NO = @STEP_NO@ " & vbCrLf
                    arrayBosParameters.Add("STEP_NO", sStepNo)
                End If

                Return GetDataRow(sSql, arrayBosParameters.ToArray)

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Function




        ''' <summary>
        ''' 取得正在運作中的案件的流程步驟
        ''' </summary>
        ''' <param name="sCaseid"></param>
        ''' <param name="nSubFlowSeq">若為0表示所有的子流程都列出</param>
        ''' <param name="sStepNo">若為NOTHING表示所有的步驟都列出</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetIncompletedStepNoByCaseid(ByVal sCaseid As String,
                                                     Optional ByVal nSubFlowSeq As Integer = 0,
                                                     Optional ByVal sStepNo As String = Nothing) As DataRowCollection

            'ByVal sCurrentCaseid As String,
            '                            ByVal sNewClient As String,
            '                            Optional ByVal nNewBraDepNo As Integer = 0,
            '                            Optional ByVal nCurrentSubFlowSeq As Integer = 0,
            '                            Optional ByVal sCurrentStepNo As String = Nothing

            Dim sSql As String
            Dim arrayBosParameters As New BosParamsList

            Try
                sSql = _
                    "select SY_FLOWSTEP.*, SY_CASEID.FLOW_ID, SY_FLOW_MAP.TYPE  " & vbCrLf & _
                    "  from SY_FLOWSTEP " & vbCrLf & _
                    " inner join SY_FLOWINCIDENT " & vbCrLf & _
                    "    on SY_FLOWSTEP.CASEID = SY_FLOWINCIDENT.CASEID " & vbCrLf & _
                    "   and SY_FLOWSTEP.SUBFLOW_SEQ = SY_FLOWINCIDENT.SUBFLOW_SEQ " & vbCrLf & _
                    " inner join SY_CASEID " & vbCrLf & _
                    "    on SY_FLOWSTEP.CASEID = SY_CASEID.CASEID " & vbCrLf & _
                    " inner join SY_FLOW_MAP " & vbCrLf & _
                    "    on SY_FLOW_MAP.STEP_NO = SY_FLOWSTEP.STEP_NO " & vbCrLf & _
                    "   and SY_FLOW_MAP.FLOW_ID = SY_CASEID.FLOW_ID " & vbCrLf & _
                    " where PROCESSTIME is null " & vbCrLf & _
                    "   and SY_FLOWSTEP.CASEID = @CASEID@ "

                arrayBosParameters.Add("CASEID", sCaseid)


                If nSubFlowSeq <> 0 Then
                    sSql = sSql & vbCrLf & _
                        "   and SY_FLOWSTEP.SUBFLOW_SEQ = @SUBFLOW_SEQ@"
                    arrayBosParameters.Add("SUBFLOW_SEQ", nSubFlowSeq)
                End If

                If String.IsNullOrEmpty(sStepNo) = False Then
                    sSql = sSql & vbCrLf & _
                        "   and SY_FLOWSTEP.STEP_NO = @STEP_NO@"
                    arrayBosParameters.Add("STEP_NO", sStepNo)
                End If

                Return GetDataRowCollection(sSql, arrayBosParameters.ToArray())

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Function


        ''' <summary>
        ''' 取得待處理及己處理案件內容
        ''' </summary>
        ''' <param name="sUserid"></param>
        ''' <param name="sBRID"></param>
        ''' <param name="nStatus">1:待處理, 3:已處理</param>
        ''' <param name="nStartRow">開始顯示案件的ROWNUM</param>
        ''' <param name="nEndRow">結束顯示案件的ROWNUM</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetCaselistByUseridStatus(ByVal sUserid As String,
                                                  ByVal sBRID As String,
                                                  ByVal nStatus As Integer,
                                                  ByVal nStartRow As Integer,
                                                  ByVal nEndRow As Integer,
                                                  ByVal sOrderBy As String,
                                                  Optional ByVal sOptionalFilter As String = "") As DataTable

            Dim sSql As String

            Try
                sSql = _
                    "select * " & vbCrLf & _
                    "  from (select *, " & vbCrLf & _
                    "               ROW_NUMBER() over(order by SUBQUERY1.STARTTIME desc) as ROWID, " & vbCrLf & _
                    "               COUNT(*) over(partition by NULL) as ROW_COUNT " & vbCrLf & _
                    "          from (select distinct SY_CASEID.CASEID, " & vbCrLf & _
                    "                                FS.SUBFLOW_SEQ, " & vbCrLf & _
                    "                                FS.CLIENT        AS CLIENT_ID, " & vbCrLf & _
                    "                                FS.BRA_DEPNO     AS CLIENT_BRADEPNO, " & vbCrLf & _
                    "                                SY_CASEID.APV_CAS_ID, " & vbCrLf & _
                    "                                SY_FLOW_ID.FLOW_CNAME, " & vbCrLf & _
                    "                                SY_CASEID.CPL_APL_ID, " & vbCrLf & _
                    "                                SY_CASEID.CPL_APL_NAM, " & vbCrLf & _
                    "                                SY_STEP_NO.STEP_NO, " & vbCrLf & _
                    "                                SY_STEP_NO.STEP_NAME, " & vbCrLf & _
                    "                                SY_FLOW_MAP.FORMURL, " & vbCrLf & _
                    "                                FIRSTSUBFLOW.STARTTIME    as CASESTTIME, " & vbCrLf & _
                    "                                FS.STARTTIME, " & vbCrLf & _
                    "                                APPBRANCH.BRCNAME " & vbCrLf & _
                    "                  from SY_BRANCH BRA" & vbCrLf & _
                    "                 inner join SY_FLOWSTEP FS" & vbCrLf & _
                    "                    on FS.BRA_DEPNO = BRA.BRA_DEPNO " & vbCrLf & _
                    "                 inner join SY_CASEID " & vbCrLf & _
                    "                    on SY_CASEID.CASEID = FS.CASEID " & vbCrLf & _
                    "                 inner join SY_FLOW_MAP " & vbCrLf & _
                    "                    on SY_FLOW_MAP.STEP_NO = FS.STEP_NO " & vbCrLf & _
                    "                   and SY_FLOW_MAP.FLOW_ID = SY_CASEID.FLOW_ID " & vbCrLf & _
                    "                 inner join SY_STEP_NO " & vbCrLf & _
                    "                    on SY_STEP_NO.STEP_NO = FS.STEP_NO " & vbCrLf & _
                    "                 inner join SY_FLOWINCIDENT " & vbCrLf & _
                    "                    on SY_FLOWINCIDENT.CASEID = FS.CASEID " & vbCrLf & _
                    "                   and SY_FLOWINCIDENT.SUBFLOW_SEQ = FS.SUBFLOW_SEQ " & vbCrLf & _
                    "                 inner join SY_FLOW_ID " & vbCrLf & _
                    "                    on SY_FLOW_ID.FLOW_ID = SY_CASEID.FLOW_ID " & vbCrLf & _
                    "                  left join SY_FLOWINCIDENT FIRSTSUBFLOW " & vbCrLf & _
                    "                    on FIRSTSUBFLOW.CASEID = FS.CASEID " & vbCrLf & _
                    "                  left join SY_BRANCH APPBRANCH " & vbCrLf & _
                    "                    on APPBRANCH.BRA_DEPNO = SY_CASEID.APP_BRADEPNO " & vbCrLf

                If nStatus = 3 Then
                    sSql &= _
                    "                 where FS.CLIENT = @CLIENT@ " & vbCrLf & _
                    "                   and FIRSTSUBFLOW.SUBFLOW_SEQ = 1 " & vbCrLf & _
                    "                   and FS.PROCESSTIME is not null " & vbCrLf
                Else
                    sSql &= _
                    "                 where FS.CLIENT = @CLIENT@ " & vbCrLf & _
                    "                   and SY_FLOWINCIDENT.STATUS <>3 " & vbCrLf & _
                    "                   and FIRSTSUBFLOW.SUBFLOW_SEQ = 1 " & vbCrLf & _
                    "                   and FS.PROCESSTIME is null " & vbCrLf
                End If

                If String.IsNullOrEmpty(sOptionalFilter) = False Then
                    sSql &= "                   and " & sOptionalFilter & vbCrLf
                End If


                sSql &= _
                    "                   and SY_FLOW_MAP.TYPE is NULL " & vbCrLf & _
                    "                   and BRA.BRID = @BRID@ ) SUBQUERY1) SUBQUERY " & vbCrLf & _
                    " where ROWID >= @START@ " & vbCrLf & _
                    "   and ROWID <= @END@ " & vbCrLf & _
                    "--取得待處理及己處理案件內容" & vbCrLf

                Return GetDataTable(sSql,
                                    "CLIENT", sUserid,
                                    "BRID", sBRID,
                                    "START", nStartRow,
                                    "END", nEndRow)

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Function




        ''' <summary>
        ''' 取得上一步驟的使用者，僅包含父流程或修正補充前變更代號的流程
        ''' </summary>
        ''' <param name="sCaseid"></param>
        ''' <param name="nSubFlowSeq"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetFormerStepStaff(ByVal sCaseid As String, ByVal nSubFlowSeq As Integer) As String

            'Dim sSql As String
            Dim dr As DataRow

            Try
                dr = GetFormerFlowStepInfo(sCaseid, nSubFlowSeq)
                Return CDbType(Of String)(dr("SENDER"), Nothing)

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Function

        ''' <summary>
        ''' 取得上一步驟的步驟資訊，僅包含父流程或修正補充前變更代號的流程
        ''' </summary>
        ''' <param name="sCaseid"></param>
        ''' <param name="nSubFlowSeq"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetFormerFlowStepInfo(ByVal sCaseid As String, ByVal nSubFlowSeq As Integer) As DataRow

            Dim sSql As String

            Try
                sSql = _
                "with n(CASEID, SUBFLOW_SEQ, PREV_SUBFLOWSEQ, PARENT_SEQ) as" & vbCrLf & _
                " (select CASEID, SUBFLOW_SEQ, PREV_SUBFLOWSEQ, PARENT_SEQ" & vbCrLf & _
                "    from SY_FLOWINCIDENT" & vbCrLf & _
                "   where CASEID = @CASEID@" & vbCrLf & _
                "     and SUBFLOW_SEQ = @SUBFLOW_SEQ@" & vbCrLf & _
                "  union all" & vbCrLf & _
                "  select a.CASEID, a.SUBFLOW_SEQ, a.PREV_SUBFLOWSEQ, a.PARENT_SEQ" & vbCrLf & _
                "    from SY_FLOWINCIDENT as a, n" & vbCrLf & _
                "   where (a.SUBFLOW_SEQ = n.PREV_SUBFLOWSEQ and a.SUBFLOW_SEQ < n.SUBFLOW_SEQ and a.CASEID = n.CASEID)" & vbCrLf & _
                "      or (a.SUBFLOW_SEQ = n.PARENT_SEQ and a.SUBFLOW_SEQ < n.SUBFLOW_SEQ and a.CASEID = n.CASEID))" & vbCrLf & _
                "select TOP 1 fs.*, SN.STEP_NAME, BR.BRCNAME" & vbCrLf & _
                "  from n" & vbCrLf & _
                " inner join SY_FLOWSTEP fs" & vbCrLf & _
                "    on n.CASEID = fs.CASEID and n.SUBFLOW_SEQ = fs.SUBFLOW_SEQ" & vbCrLf & _
                " inner join SY_CASEID ca" & vbCrLf & _
                "    on ca.CASEID = n.CASEID" & vbCrLf & _
                " inner join SY_FLOW_MAP fm" & vbCrLf & _
                "    on ca.FLOW_ID = fm.FLOW_ID and fs.STEP_NO = fm.STEP_NO" & vbCrLf & _
                " inner join SY_FLOW_DEF fd" & vbCrLf & _
                "    on ca.FLOW_ID = fd.FLOW_ID and fs.STEP_NO = fd.CURR_STEP_NO" & vbCrLf & _
                " inner join SY_STEP_NO SN" & vbCrLf & _
                "    on SN.STEP_NO = fm.STEP_NO" & vbCrLf & _
                " inner join SY_BRANCH BR" & vbCrLf & _
                "    on BR.BRA_DEPNO = fs.BRA_DEPNO  " & vbCrLf & _
                " where PROCESSTIME is not NULL and TYPE is null and STATUS <> 9" & vbCrLf & _
                "   and ca.CASEID = @CASEID@" & vbCrLf & _
                "   and PROCESSTIME = (select max(PROCESSTIME) from SY_FLOWSTEP where PROCESSTIME is not NULL and TYPE is null and STATUS <> 9 and CASEID = @CASEID@) "

                Return GetDataRow(sSql,
                                  "CASEID", sCaseid,
                                  "SUBFLOW_SEQ", nSubFlowSeq)

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Function



        ''' <summary>
        ''' 寫入結束(PROCESSTIME is not NULL)至FLOWSTEP
        ''' </summary>
        ''' <param name="sCaseid"></param>
        ''' <param name="nSubFlowSeq"></param>
        ''' <remarks></remarks>
        Public Sub CloseByCaseidSubflowseq(ByVal sCaseid As String, ByVal nSubFlowSeq As Integer, ByVal sSender As String)
            Dim sSql As String
            Dim nValue As Integer

            Try
                sSql =
                    "update SY_FLOWSTEP " & vbCrLf & _
                    "   set PROCESSTIME = @PROCESSTIME@, SENDER = @SENDER@ " & vbCrLf & _
                    " where CASEID = @CASEID@ " & vbCrLf & _
                    "   and SUBFLOW_SEQ = @SUBFLOW_SEQ@ " & vbCrLf & _
                    "   and PROCESSTIME is NULL"

                nValue = ExecuteNonQuery(sSql,
                                         "PROCESSTIME", Now,
                                         "CASEID", sCaseid,
                                         "SUBFLOW_SEQ", nSubFlowSeq,
                                         "SENDER", sSender)
                'Dbg.Assert(nValue = 1)

                If nValue = 0 Then
                    nValue = GetCount(BosBase.PARAM_ARRAY("CASEID", sCaseid,
                                      "SUBFLOW_SEQ", nSubFlowSeq))
                    If nValue = 0 Then
                        Throw New SYException(
                            String.Format("無法變更進行中的流程狀態，CASEID:{0}及子流程代碼:{1}", sCaseid, nSubFlowSeq),
                            SYMSG.SYCASEID_CANNOT_CHANGE,
                            GetLastSQL)
                    End If
                End If

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Sub

        ''' <summary>
        ''' 取得某一使用者佇列取件內的案件
        ''' </summary>
        ''' <param name="sUserid"></param>
        ''' <param name="sBrid"></param>
        ''' <param name="sOrderBy">排序的欄位，欄位後加入 DESC為反向排序</param>
        ''' <param name="sCaseid">額外過濾條件，只顯示CASEID為sCaseid的案件</param>
        ''' <param name="nSubFlowSeq">額外過濾條件，只顯示Subflow_Seq為nSubFlowSeq的案件</param>
        ''' <returns>若上一步驟跟此步驟的使用者不可以是同一使用者，則不列出</returns>
        ''' <remarks></remarks>
        Public Function GetCaseListFromQueue(ByVal sUserid As String,
                                             Optional ByVal sBrid As String = Nothing,
                                             Optional ByVal sOrderBy As String = Nothing,
                                             Optional ByVal sCaseid As String = Nothing,
                                             Optional ByVal nSubFlowSeq As Integer = 0,
                                             Optional ByVal sStepNo As String = Nothing,
                                             Optional ByVal bNotGetVirtualStep As Boolean = False) As DataTable

            Dim sSql As String
            Dim dt As DataTable
            'Dim drList As New List(Of DataRow)
            Dim paramList As New BosParamsList
            Dim bSameUserDeleted As Boolean = False

            'Dim arrayBraDepNo() As Integer = Nothing

            Try
                'If nBraDepNo <> 0 Then
                '    arrayBraDepNo = getSYBranch.GetChildFromParent(nBraDepNo)
                'End If

                paramList.Add("STAFFID", sUserid)

                sSql = _
                    "with RC( ROLEID, ROLENAME, DISABLED, ROLETYPE, PARENT, STAFFID) as " & vbCrLf & _
                    " (select SY_ROLE.ROLEID, SY_ROLE.ROLENAME, SY_ROLE.DISABLED, SY_ROLE.ROLETYPE, SY_ROLE.PARENT, RU.STAFFID " & vbCrLf & _
                    "    from SY_ROLE " & vbCrLf & _
                    "   inner join SY_REL_ROLE_USER RU " & vbCrLf & _
                    "      on RU.ROLEID = SY_ROLE.ROLEID " & vbCrLf & _
                    "  where SY_ROLE.DISABLED = '0' " & vbCrLf & _
                    "  union all " & vbCrLf & _
                    "  select a.ROLEID, a.ROLENAME, a.DISABLED, a.ROLETYPE, a.PARENT, RC.STAFFID " & vbCrLf & _
                    "    from SY_ROLE as a, RC " & vbCrLf & _
                    "   where a.PARENT = RC.ROLEID " & vbCrLf & _
                    "     and RC.DISABLED = '0' " & vbCrLf & _
                    "     and a.DISABLED = '0') " & vbCrLf & _
                    "select distinct CA.CASEID, CA.APV_CAS_ID, FID.FLOW_CNAME, CA.CPL_APL_ID, CA.CPL_APL_NAM, SN.STEP_NAME, FIRSTSUBFLOW.STARTTIME as CASESTTIME, FI.STARTTIME, BRA_APP.BRCNAME, FS.STEP_NO, FS.SUBFLOW_SEQ, CA.FLOW_ID " & vbCrLf & _
                    "  from RC " & vbCrLf & _
                    " inner join SY_REL_ROLE_FLOWMAP RRF " & vbCrLf & _
                    "    on RRF.ROLEID = RC.ROLEID " & vbCrLf & _
                    " inner join SY_CASEID CA " & vbCrLf & _
                    "    on CA.FLOW_ID = RRF.FLOW_ID " & vbCrLf & _
                    " inner join SY_FLOW_ID FID " & vbCrLf & _
                    "    on FID.FLOW_ID = CA.FLOW_ID " & vbCrLf & _
                    " inner join SY_FLOWINCIDENT FI " & vbCrLf & _
                    "    on FI.CASEID = CA.CASEID " & vbCrLf & _
                    " inner join SY_FLOWINCIDENT FIRSTSUBFLOW " & vbCrLf & _
                    "    on FIRSTSUBFLOW.CASEID = CA.CASEID " & vbCrLf & _
                    " inner join SY_FLOWSTEP FS " & vbCrLf & _
                    "    on FS.CASEID = FI.CASEID and FS.SUBFLOW_SEQ = FI.SUBFLOW_SEQ and FS.STEP_NO = RRF.STEP_NO " & vbCrLf & _
                    " inner join SY_STEP_NO SN " & vbCrLf & _
                    "    on SN.STEP_NO = FS.STEP_NO " & vbCrLf & _
                    " inner join SY_BRANCH BRA " & vbCrLf & _
                    "    on BRA.BRA_DEPNO = FS.BRA_DEPNO " & vbCrLf & _
                    "  left join SY_BRANCH BRA_APP " & vbCrLf & _
                    "    on BRA_APP.BRA_DEPNO = CA.APP_BRADEPNO " & vbCrLf & _
                    " outer apply( " & vbCrLf & _
                    "select TOP 1 SENDER PR_SENDER, PR_FS.STEP_NO PR_STEPNO, PR_FM.TYPE PR_TYPE " & vbCrLf & _
                    "  from SY_FLOWSTEP PR_FS " & vbCrLf & _
                    " inner join SY_CASEID PR_CA " & vbCrLf & _
                    "    on PR_FS.CASEID = PR_CA.CASEID " & vbCrLf & _
                    " inner join SY_FLOW_MAP PR_FM " & vbCrLf & _
                    "    on PR_FM.FLOW_ID = PR_CA.FLOW_ID and PR_FM.STEP_NO = PR_FS.STEP_NO " & vbCrLf & _
                    " where PR_FM.TYPE is null and PR_FS.CASEID = FS.CASEID and PR_FS.SUBFLOW_SEQ = FS.SUBFLOW_SEQ and PROCESSTIME is not null " & vbCrLf &
                    " order by PROCESSTIME desc) PRE " & vbCrLf & _
                    "  left join SY_FLOW_DEF FD " & vbCrLf & _
                    "    on FD.FLOW_ID = CA.FLOW_ID and FD.NEXT_STEP_NO = FS.STEP_NO and FD.CURR_STEP_NO = PRE.PR_STEPNO " & vbCrLf & _
                    "  left join SY_FLOW_MAP FM " & vbCrLf & _
                    "    on FM.FLOW_ID = FD.FLOW_ID and FM.STEP_NO = FD.CURR_STEP_NO " & vbCrLf & _
                    " where FIRSTSUBFLOW.SUBFLOW_SEQ = 1 " & vbCrLf & _
                    "   and FS.CLIENT is null " & vbCrLf & _
                    "   and FS.PROCESSTIME is null " & vbCrLf & _
                    "   and RC.STAFFID = @STAFFID@ " & vbCrLf & _
                    "   and FI.STATUS = 1 " & vbCrLf & _
                    "   and (FD.PROHIBITE_SAME is null or FD.PROHIBITE_SAME = 'N' or (FD.PROHIBITE_SAME = 'Y' and PRE.PR_SENDER <> RC.STAFFID and PRE.PR_TYPE is null and FM.TYPE is NULL))" & vbCrLf

                If bNotGetVirtualStep Then
                    sSql &= "   and FM.TYPE is NULL " & vbCrLf
                End If

                If String.IsNullOrEmpty(sBrid) = False Then

                    sSql &= "   and BRA.BRID = @BRID@ and BRA.DISABLED='0' " & vbCrLf
                    paramList.Add("BRID", sBrid)
                End If

                If Not String.IsNullOrEmpty(sCaseid) Then
                    sSql &= "   and FI.CASEID = @CASEID@ " & vbCrLf
                    paramList.Add("CASEID", sCaseid)
                End If

                If nSubFlowSeq <> 0 Then
                    sSql &= "   and FI.SUBFLOW_SEQ = @SUBFLOW_SEQ@ " & vbCrLf
                    paramList.Add("SUBFLOW_SEQ", nSubFlowSeq)
                End If

                If Not String.IsNullOrEmpty(sStepNo) Then
                    sSql &= "   and FS.STEP_NO = @STEP_NO@ " & vbCrLf
                    paramList.Add("STEP_NO", sStepNo)
                End If

                If Not String.IsNullOrEmpty(sOrderBy) Then
                    sSql &= " order by " & sOrderBy.Trim & vbCrLf
                End If

                If String.IsNullOrEmpty(sCaseid) Then
                    sSql &= "--佇列取件"
                End If

                dt = GetDataTable(sSql, paramList.ToArray())

                If dt.Rows.Count = 0 Then
                    If bSameUserDeleted = True Then
                        Throw New SYException(
                            String.Format("佇列取件內沒有可分派給目前使用者的案件：OWNER={0}，BRID={1}，CASEID={2}，SUBFLOW_SEQ={3}。", sUserid, sBrid, sCaseid, nSubFlowSeq) & vbCrLf & "有些案件的上一步驟案件送出者與目前使用者相同，所以被禁止列出。",
                            SYMSG.SYFLOW_CASENOTFOUND_PROHIBITESAMEUSER, GetLastSQL, dt)
                    Else
                        Throw New SYException(
                            String.Format("佇列取件內沒有可分派給目前使用者的案件：OWNER={0}，BRID={1}，CASEID={2}，SUBFLOW_SEQ={3}", sUserid, sBrid, sCaseid, nSubFlowSeq),
                            SYMSG.SYFLOW_CASE_NOT_FOUND, GetLastSQL, dt)
                    End If
                End If

                Return dt

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Function


        ''' <summary>
        ''' 指定某一步驟的Owner
        ''' </summary>
        ''' <param name="sCaseid"></param>
        ''' <param name="nCurrentSubFlowSeq"></param>
        ''' <param name="nextStepUser"></param>
        ''' <remarks></remarks>
        Public Sub AssignNextStepOwner(ByVal sCaseid As String, ByVal nCurrentSubFlowSeq As Integer,
                                       ByVal nextStepUser As STEPNO_USERID)

            Dim obj As Object

            Try
                obj = getSYRelBranchUser.ExecuteScalar(
                    "select TOP 1 BRA_DEPNO " & vbCrLf & _
                    "  from SY_REL_BRANCH_USER " & vbCrLf & _
                    " where STAFFID=@STAFFID@ " & vbCrLf,
                    "STAFFID", nextStepUser.sUserId)

                If IsDBNull(obj) Then
                    Throw New SYException(
                        String.Format("無法找到使用者的分行代碼，STAFFID:{0}", nextStepUser.sUserId),
                        SYMSG.SYFLOWSTEP_BRID_NOT_FOUND,
                        GetLastSQL)
                End If

                InsertUpdate("CASEID", sCaseid,
                             "SUBFLOW_SEQ", nCurrentSubFlowSeq,
                             "SUBFLOW_COUNT", 0,
                             "STEP_NO", nextStepUser.sStepNo,
                             "SUMMARY", "指定步驟人員",
                             "STATUS", 9,
                             "BRA_DEPNO", obj,
                             "OWNER", nextStepUser.sUserId,
                             "CLIENT", nextStepUser.sUserId,
                             "SENDER", nextStepUser.sUserId,
                             "PROCESSTIME", Now,
                             "STARTTIME", Now
                             )

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Sub



        ''' <summary>
        ''' 找出目前流程的SPLITTER
        ''' </summary>
        ''' <param name="sCaseid"></param>
        ''' <param name="nSubflowSeq"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetStepNameByCaseidSubflowseq(ByVal sCaseid As String, ByVal nSubflowSeq As Integer) As String

            Dim obj As Object

            Try

                obj = ExecuteScalar( _
                    "select distinct FS.STEP_NO " & vbCrLf & _
                    "  from SY_FLOWSTEP FS " & vbCrLf & _
                    " inner join SY_FLOW_MAP FM " & vbCrLf & _
                    "    on FS.STEP_NO = FM.STEP_NO " & vbCrLf & _
                    "   and FM.TYPE = 'S' " & vbCrLf & _
                    " where CASEID = @CASEID@ " & vbCrLf & _
                    "   and SUBFLOW_SEQ = @SUBFLOW_SEQ@ " & vbCrLf,
                    "CASEID", sCaseid,
                    "SUBFLOW_SEQ", nSubflowSeq)

                If IsDBNull(obj) Then
                    Throw New SYException(
                        String.Format("無法找到平行流程的分支點：CASEID={0}，SUBFLOW_SEQ={1}", sCaseid, nSubflowSeq),
                        SYMSG.SYFLOWSTEP_SPLITTER_NOT_FOUND,
                        GetLastSQL)
                End If

                Return CStr(obj)

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Function


        ''' <summary>
        ''' 輸入CASEID及STEPNO，取得取近一筆的記錄
        ''' </summary>
        ''' <param name="sCaseId"></param>
        ''' <param name="sStepNo"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function GetLatestSubFlowNoByStepNo(ByVal sCaseId As String, ByVal sStepNo As String) As Integer
            Dim sSQL As String
            Dim obj As Object

            Try
                sSQL = _
                    "select top 1 SUBFLOW_SEQ from SY_FLOWSTEP " & vbCrLf & _
                    " where CASEID = @CASEID@ " & vbCrLf & _
                    "   and STEP_NO = @STEP_NO@" & vbCrLf & _
                    " order by PROCESSTIME DESC" & vbCrLf

                obj = ExecuteScalar("select top 1 SUBFLOW_SEQ from SY_FLOWSTEP " & vbCrLf & _
                                    " where CASEID = @CASEID@ " & vbCrLf & _
                                    "   and STEP_NO = @STEP_NO@" & vbCrLf & _
                                    " order by PROCESSTIME DESC" & vbCrLf,
                                    "CASEID", sCaseId,
                                    "STEP_NO", sStepNo)

                If IsNothing(obj) OrElse IsDBNull(obj) Then
                    Throw New SYException(
                        String.Format("無法找到流程的子流程編號：CASEID={0}，STEP_NO={1}", sCaseId, sStepNo),
                        SYMSG.SYFLOWSTEP_SUBFLOWSEQ_NOT_FOUND,
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
        ''' 取得案件步驟的開始時間
        ''' </summary>
        ''' <param name="sCaseid"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        'Function GetStartTime(ByVal sCaseid As String, ByVal nSubFlowSeq As Integer,
        '                      ByVal sStepNo As String) As DateTime

        '    Dim obj As Object

        '    Try
        '        obj = ExecuteScalar(
        '            "select top 1 STARTTIME " & vbCrLf & _
        '            "  from SY_FLOWSTEP " & vbCrLf & _
        '            " where CASEID = @CASEID@ " & vbCrLf & _
        '            "   and SUBFLOW_SEQ = @SUBFLOW_SEQ@ " & vbCrLf & _
        '            "   and STEP_NO = @STEP_NO@ " & vbCrLf & _
        '            " order by SUBFLOW_COUNT desc",
        '            New BosParameter() {
        '                PARAMETER("CASEID", sCaseid),
        '                PARAMETER("SUBFLOW_SEQ", nSubFlowSeq),
        '                PARAMETER("STEP_NO", sStepNo)})

        '        Return CDbType(Of DateTime)(obj)

        '    Catch ex As Exception
        '        Throw
        '    End Try

        'End Function

        ''' <summary>
        ''' 取得最近的步驟
        ''' </summary>
        ''' <param name="sCaseid"></param>
        ''' <param name="bExcludeEnd">不包含END</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetLatestStepInfo(ByVal sCaseid As String, Optional ByVal bExcludeEnd As Boolean = False) As DataRow
            Try
                If bExcludeEnd Then
                    Return GetDataRow( _
                        "select top 1 * " & vbCrLf & _
                        "  from SY_FLOWSTEP " & vbCrLf & _
                        " where CASEID = @CASEID@ " & vbCrLf & _
                        "   and STEP_NO <> 'END' " & vbCrLf & _
                        " order by STARTTIME desc " & vbCrLf, "CASEID", sCaseid)
                Else
                    Return GetDataRow( _
                        "select top 1 * " & vbCrLf & _
                        "  from SY_FLOWSTEP " & vbCrLf & _
                        " where CASEID = @CASEID@ " & vbCrLf & _
                        " order by STARTTIME desc " & vbCrLf, "CASEID", sCaseid)
                End If

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Function


        ''' <summary>
        ''' 取得案件可被重新指派的使用者
        ''' </summary>
        ''' <param name="sCaseId"></param>
        ''' <param name="nSubFlowSeq">子流程編號</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetAlternativeUserList(ByVal sCaseId As String,
                                               Optional ByVal nSubFlowSeq As Integer = 0) As DataTable
            Dim sSql As String = Nothing
            Dim dt As DataTable

            Try

                dt = GetDataTable(
                    "with RC( ROLEID, ROLENAME, DISABLED, ROLETYPE, PARENT, STAFFID) as " & vbCrLf & _
                    " (select SY_ROLE.ROLEID, SY_ROLE.ROLENAME, SY_ROLE.DISABLED, SY_ROLE.ROLETYPE, SY_ROLE.PARENT, RU.STAFFID " & vbCrLf & _
                    "    from SY_ROLE " & vbCrLf & _
                    "   inner join SY_REL_ROLE_USER RU " & vbCrLf & _
                    "      on RU.ROLEID = SY_ROLE.ROLEID " & vbCrLf & _
                    "   inner join SY_BRANCH BR" & vbCrLf & _
                    "      on BR.BRA_DEPNO = RU.BRA_DEPNO   " & vbCrLf & _
                    "   where SY_ROLE.DISABLED = '0' " & vbCrLf & _
                    "   union all " & vbCrLf & _
                    "  select a.ROLEID, a.ROLENAME, a.DISABLED, a.ROLETYPE, a.PARENT, RC.STAFFID " & vbCrLf & _
                    "    from SY_ROLE a" & vbCrLf & _
                    "   inner join RC " & vbCrLf & _
                    "      on a.PARENT = RC.ROLEID " & vbCrLf & _
                    "     and RC.DISABLED = '0' " & vbCrLf & _
                    "     and a.DISABLED = '0') " & vbCrLf & _
                    "select distinct RC.STAFFID, UR.USERNAME " & vbCrLf & _
                    "  from RC " & vbCrLf & _
                    " inner join SY_REL_ROLE_FLOWMAP RRF " & vbCrLf & _
                    "    on RRF.ROLEID = RC.ROLEID " & vbCrLf & _
                    " inner join SY_CASEID CA " & vbCrLf & _
                    "    on CA.FLOW_ID = RRF.FLOW_ID " & vbCrLf & _
                    " inner join SY_FLOW_ID FID " & vbCrLf & _
                    "    on FID.FLOW_ID = CA.FLOW_ID " & vbCrLf & _
                    " inner join SY_FLOWINCIDENT FI " & vbCrLf & _
                    "    on FI.CASEID = CA.CASEID " & vbCrLf & _
                    " inner join SY_FLOWINCIDENT FIRSTSUBFLOW " & vbCrLf & _
                    "    on FIRSTSUBFLOW.CASEID = CA.CASEID " & vbCrLf & _
                    " inner join SY_FLOWSTEP FS " & vbCrLf & _
                    "    on FS.CASEID = FI.CASEID and FS.SUBFLOW_SEQ = FI.SUBFLOW_SEQ and FS.STEP_NO = RRF.STEP_NO " & vbCrLf & _
                    " inner join SY_BRANCH BRA " & vbCrLf & _
                    "    on BRA.BRA_DEPNO = FS.BRA_DEPNO " & vbCrLf & _
                    " inner join SY_REL_ROLE_USER RRU " & vbCrLf & _
                    "    on RRU.ROLEID = RC.ROLEID " & vbCrLf & _
                    "   and RRU.BRA_DEPNO = FS.BRA_DEPNO " & vbCrLf & _
                    "   and RRU.STAFFID = RC.STAFFID " & vbCrLf & _
                    " outer apply( " & vbCrLf & _
                    "select TOP 1 SENDER PR_SENDER, PR_FS.STEP_NO PR_STEPNO, PR_FM.TYPE PR_TYPE " & vbCrLf & _
                    "  from SY_FLOWSTEP PR_FS " & vbCrLf & _
                    " inner join SY_CASEID PR_CA " & vbCrLf & _
                    "    on PR_FS.CASEID = PR_CA.CASEID " & vbCrLf & _
                    " inner join SY_FLOW_MAP PR_FM " & vbCrLf & _
                    "    on PR_FM.FLOW_ID = PR_CA.FLOW_ID and PR_FM.STEP_NO = PR_FS.STEP_NO " & vbCrLf & _
                    " where PR_FM.TYPE is null and PR_FS.CASEID = FS.CASEID and PR_FS.SUBFLOW_SEQ = FS.SUBFLOW_SEQ and PROCESSTIME is not null " & vbCrLf & _
                    " order by PROCESSTIME desc) PRE " & vbCrLf & _
                    "  left join SY_FLOW_DEF FD " & vbCrLf & _
                    "    on FD.FLOW_ID = CA.FLOW_ID and FD.NEXT_STEP_NO = FS.STEP_NO and FD.CURR_STEP_NO = PRE.PR_STEPNO " & vbCrLf & _
                    "  left join SY_FLOW_MAP FM " & vbCrLf & _
                    "    on FM.FLOW_ID = FD.FLOW_ID and FM.STEP_NO = FD.CURR_STEP_NO " & vbCrLf & _
                    " inner join SY_USER UR " & vbCrLf & _
                    "    on UR.STAFFID = RC.STAFFID " & vbCrLf & _
                    "   and UR.STATUS = '0' " & vbCrLf & _
                    " where FIRSTSUBFLOW.SUBFLOW_SEQ = 1 " & vbCrLf & _
                    "   and FS.CASEID = @CASEID@" & vbCrLf & _
                    "   and (FS.SUBFLOW_SEQ = @SUBFLOW_SEQ@ or 0 = @SUBFLOW_SEQ@)" & vbCrLf & _
                    "   and FS.PROCESSTIME is null " & vbCrLf & _
                    "   and FI.STATUS = 1 " & vbCrLf & _
                    "   and (FD.PROHIBITE_SAME is null or FD.PROHIBITE_SAME = 'N' or (FD.PROHIBITE_SAME = 'Y' and PRE.PR_SENDER <> RC.STAFFID and PRE.PR_TYPE is null and FM.TYPE is NULL))" & vbCrLf & _
                    "   and FM.TYPE is NULL " & vbCrLf & _
                    "--取得案件可被重新指派的使用者" & vbCrLf,
                    "CASEID", sCaseId,
                    "SUBFLOW_SEQ", nSubFlowSeq)

                Return dt

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Function



        ''' <summary>
        ''' 取得分行內特定流程及步驟的待處理案件
        ''' </summary>
        ''' <param name="sBrid"></param>
        ''' <param name="nFlowId"></param>
        ''' <param name="sStepNo"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetPendingCaseList(ByVal sBrid As String,
                                           ByVal nFlowId As Integer,
                                           ByVal sStepNo As String) As DataTable
            Dim sSql As String = Nothing
            Dim dt As DataTable

            Try

                dt = GetDataTable(
                    "select distinct SY_CASEID.CASEID," & vbCrLf & _
                    "                FS.SUBFLOW_SEQ," & vbCrLf & _
                    "                SY_CASEID.CPL_APL_ID," & vbCrLf & _
                    "                SY_CASEID.CPL_APL_NAM," & vbCrLf & _
                    "                FS.STARTTIME           as CASERECVTIME," & vbCrLf & _
                    "                FIRSTSUBFLOW.STARTTIME as CASESTTIME," & vbCrLf & _
                    "                FS.OWNER," & vbCrLf & _
                    "                OUSER.USERNAME as OWNERNAME," & vbCrLf & _
                    "                FS.CLIENT, " & vbCrLf & _
                    "                CUSER.USERNAME as CLIENTNAME" & vbCrLf & _
                    "  from SY_BRANCH BRA" & vbCrLf & _
                    " inner join SY_FLOWSTEP FS" & vbCrLf & _
                    "    on FS.BRA_DEPNO = BRA.BRA_DEPNO" & vbCrLf & _
                    " inner join SY_CASEID" & vbCrLf & _
                    "    on SY_CASEID.CASEID = FS.CASEID" & vbCrLf & _
                    " inner join SY_FLOW_MAP" & vbCrLf & _
                    "    on SY_FLOW_MAP.STEP_NO = FS.STEP_NO" & vbCrLf & _
                    "   and SY_FLOW_MAP.FLOW_ID = SY_CASEID.FLOW_ID" & vbCrLf & _
                    " inner join SY_STEP_NO" & vbCrLf & _
                    "    on SY_STEP_NO.STEP_NO = FS.STEP_NO" & vbCrLf & _
                    " inner join SY_FLOWINCIDENT" & vbCrLf & _
                    "    on SY_FLOWINCIDENT.CASEID = FS.CASEID" & vbCrLf & _
                    "   and SY_FLOWINCIDENT.SUBFLOW_SEQ = FS.SUBFLOW_SEQ" & vbCrLf & _
                    " inner join SY_FLOW_ID" & vbCrLf & _
                    "    on SY_FLOW_ID.FLOW_ID = SY_CASEID.FLOW_ID" & vbCrLf & _
                    "  left join SY_FLOWINCIDENT FIRSTSUBFLOW" & vbCrLf & _
                    "    on FIRSTSUBFLOW.CASEID = FS.CASEID" & vbCrLf & _
                    "  left join SY_USER OUSER " & vbCrLf & _
                    "    on OUSER.STAFFID = FS.OWNER" & vbCrLf & _
                    "  left join SY_USER CUSER" & vbCrLf & _
                    "    on CUSER.STAFFID = FS.CLIENT" & vbCrLf & _
                    " where SY_FLOWINCIDENT.STATUS <> 3" & vbCrLf & _
                    "   and FIRSTSUBFLOW.SUBFLOW_SEQ = 1" & vbCrLf & _
                    "   and SY_FLOW_MAP.TYPE is NULL" & vbCrLf & _
                    "   and BRA.BRID = @BRID@" & vbCrLf & _
                    "   and SY_FLOW_ID.FLOW_ID = @FLOW_ID@" & vbCrLf & _
                    "   and FS.STEP_NO = @STEP_NO@" & vbCrLf & _
                    "   and FS.CLIENT is not null" & vbCrLf & _
                    "   and FS.PROCESSTIME is null" & vbCrLf,
                    "BRID", sBrid,
                    "FLOW_ID", nFlowId,
                    "STEP_NO", sStepNo)

                Return dt

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Function

    End Class

End Namespace


