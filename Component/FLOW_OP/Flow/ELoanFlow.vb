Option Explicit On
Option Strict On

Imports com.Azion.NET.VB
Imports FLOW_OP.TABLE
Imports System.Text



Public Class StepInfoItem
    Public caseId As String
    Public stepNo As String
    Public subflowSeq As Integer
    Public subflowCount As Integer

    Public ReadOnly Property clone As StepInfoItem
        Get
            Dim sii As New StepInfoItem
            sii.caseId = caseId
            sii.stepNo = stepNo
            sii.subflowSeq = subflowSeq
            sii.subflowCount = subflowCount
            Return sii
        End Get
    End Property

End Class

Public Class StepInfoItemExt
    Inherits StepInfoItem

    Public flowId As Integer
    Public flowName As String
    Public braDepNo_brid() As BRADEPNO_BRID

    Public caseOwner As USER_ID_NAME          ' 案件的擁有者，若有多個OWNER，則caseOwner為NULL，但caseOwnerList會存放所有的OWNER
    Public caseClient As USER_ID_NAME
    Public caseSender As USER_ID_NAME

    Public subflowLevel As Integer
    Public parentSubFlowSeq As Integer
    Public previousSubFlowSeq As Integer
    Public parallelNo As Integer
    Public flowstepStatus As Integer
    Public stepSummary As String
    Public stepType As String
    Public caseOwnerList As List(Of USER_ID_NAME)

    Public revisionSeqNo As Integer '修正補充
    Public processTime As DateTime
    Public startTime As Date


    Public Overloads ReadOnly Property clone As StepInfoItemExt
        Get
            Dim sii As New StepInfoItemExt
            sii.caseId = caseId
            sii.stepNo = stepNo
            sii.subflowSeq = subflowSeq
            sii.subflowCount = subflowCount

            sii.parentSubFlowSeq = parentSubFlowSeq
            sii.previousSubFlowSeq = previousSubFlowSeq
            sii.flowId = flowId
            sii.flowName = flowName
            sii.braDepNo_brid = braDepNo_brid
            sii.caseOwner = caseOwner
            sii.caseClient = caseClient
            sii.caseSender = caseSender
            sii.subflowLevel = subflowLevel
            sii.parallelNo = parallelNo
            sii.flowstepStatus = flowstepStatus
            sii.stepType = stepType
            sii.startTime = startTime

            Return sii
        End Get
    End Property

    Public Function ToArray() As StepInfoItemExt()
        Dim siie(0) As StepInfoItemExt
        siie(0) = Me
        Return siie
    End Function

End Class


''' <summary>
''' 
''' </summary>
''' <remarks>
''' caseid放在currentStepInfo內
''' </remarks>
Public Class StepInfo
    Public currentStepInfo As StepInfoItemExt
    Public nextStepInfo() As StepInfoItemExt    '<StepInfoItem>
End Class



Public Class STEPNO_BRADEPNO
    Public BraDepNo As Integer
    Public StepNo As String

    Public Shared Function Pair(ByVal nBraDepNo As Integer, ByVal sStepNo As String) As STEPNO_BRADEPNO
        Dim sb As New STEPNO_BRADEPNO
        sb.BraDepNo = nBraDepNo
        sb.StepNo = sStepNo
        Return sb
    End Function

    Public Shared Function arrayStepBrano(ByVal nBraDepNo As Integer, ByVal sStepNo As String) As STEPNO_BRADEPNO()

        If nBraDepNo = 0 AndAlso String.IsNullOrEmpty(sStepNo) Then
            Return Nothing
        End If

        Dim sb(0) As STEPNO_BRADEPNO
        sb(0) = New STEPNO_BRADEPNO
        sb(0).BraDepNo = nBraDepNo
        sb(0).StepNo = sStepNo
        Return sb
    End Function

End Class


Public Class BRADEPNO_BRID
    Public BraDepNo As Integer
    Public Brid As String

    Public Shared Function Pair(ByVal nBraDepNo As Integer, ByVal sBrid As String) As BRADEPNO_BRID
        Dim bb As New BRADEPNO_BRID
        bb.BraDepNo = nBraDepNo
        bb.Brid = sBrid
        Return bb
    End Function

    Public Shared Function arrayBradepnoBrid(ByVal nBraDepNo As Integer, ByVal sBrid As String) As BRADEPNO_BRID()

        If nBraDepNo = 0 AndAlso String.IsNullOrEmpty(sBrid) Then
            Return Nothing
        End If

        Dim bb(0) As BRADEPNO_BRID
        bb(0) = New BRADEPNO_BRID
        bb(0).BraDepNo = nBraDepNo
        bb(0).Brid = sBrid
        Return bb
    End Function

End Class


Public Class ELoanFlow
    Inherits SY_TABLEBASE


    Public Shared m_callbackUserInfo As IFlowCallBack
    Public Shared m_callbackFlowMail As IFlowSentNotification

    ''' <summary>
    ''' Constructor
    ''' </summary>
    ''' <param name="dbManager"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal dbManager As com.Azion.NET.VB.DatabaseManager)
        MyBase.new(dbManager)
    End Sub



    ' ''' <summary>
    ' ''' 
    ' ''' </summary>
    ' ''' <param name="nBraDepNo"></param>
    ' ''' <param name="sFlowName"></param>
    ' ''' <param name="bAutoSendToNext">起案，並自動送至下一步驟</param>
    ' ''' <param name="assignNextStepUser"></param>
    ' ''' <returns></returns>
    ' ''' <remarks></remarks>
    ' ''' 

    ''' <summary>
    ''' 起案
    ''' </summary>
    ''' <param name="nBraDepNo">分行代碼</param>
    ''' <param name="sFlowName">流程名稱</param>
    ''' <param name="bAutoSendToNext">起案後自動送至下一步驟，例如，當bAutoSendToNext=True時START->0000->0300</param>
    ''' <param name="assignNextStepUser">指定下一步驟的USER</param>
    ''' <param name="bSendToEnd">自動送至END，例如，START->0000->END，適用只有一個步驟的案件(起案按確認送出後結案)</param>
    ''' <param name="sApvCasId">核准編號</param>
    ''' <param name="sAppBrid">申請單位</param>
    ''' <param name="sCplAplId">授信戶編號</param>
    ''' <param name="sCplAplNam">授信戶名稱</param>
    ''' <param name="bSendMail">是否要送信至下個步驟的USER</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function StartFlow(ByVal nBraDepNo As Integer,
                              ByVal sFlowName As String,
                              Optional ByVal bAutoSendToNext As Boolean = False,
                              Optional ByVal assignNextStepUser() As STEPNO_USERID = Nothing,
                              Optional ByVal bSendToEnd As Boolean = False,
                              Optional ByVal sApvCasId As String = "",
                              Optional ByVal sAppBrid As String = "",
                              Optional ByVal sCplAplId As String = "",
                              Optional ByVal sCplAplNam As String = "",
                              Optional ByVal bSendMail As Boolean = True) As StepInfo

        Dim newCaseid As String
        Dim stepInfo As StepInfo

        Try
            newCaseid = CreateNewCaseId(Nothing, nBraDepNo, sFlowName)

            '使用新案號起案
            stepInfo = StartFlowByCaseId(newCaseid, bAutoSendToNext, assignNextStepUser, bSendToEnd, sApvCasId, sAppBrid, sCplAplId, sCplAplNam, bSendMail)
            Dbg.Assert(Not IsNothing(stepInfo))
            If IsNothing(stepInfo) Then
                Return Nothing
            End If

            Return stepInfo

        Catch ex As SYException
            Throw
        Catch ex As Exception
            Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
        End Try

    End Function



    ''' <summary>
    ''' 建立新CASEID並寫入至SY_CASEID
    ''' </summary>
    ''' <param name="sCaseid"></param>
    ''' <param name="nBraDepNo"></param>
    ''' <param name="sFlowName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CreateNewCaseId(ByVal sCaseid As String,
                                    ByVal nBraDepNo As Integer,
                                    ByVal sFlowName As String) As String
        Dim newCaseid As String = Nothing
        Dim sSysid, sSubsysid As String
        Dim nFlowId As Integer
        Dim dr As DataRow
        Dim bResult As Boolean

        Try

            dr = getSYFlowId().GetInfo(sFlowName)

            If Not IsNothing(dr) Then
                sSysid = CStr(dr("SYSID"))
                sSubsysid = CStr(dr("SUBSYSID"))
                nFlowId = CInt(dr("FLOW_ID"))
            Else
                Throw New SYException(
                    "無法取得(FLOWNAME = " & sFlowName & " )的流程資訊",
                    SYMSG.FLOW_FLOWINFO_NOT_FOUND,
                    getSYFlowId().GetLastSQL)
            End If

            '若不指定BraDepNo，則從Session的WorkingUser取得BRA_DEPNO
            If (nBraDepNo = 0) Then

                Dim loginUserId, workingUserId As USER_ID_NAME
                loginUserId = Nothing
                workingUserId = Nothing
                GetCurrentUserid(loginUserId, workingUserId)

                Dim sFirstStep = getSYFlowDef.GetFirstStepExceptStart(sFlowName)
                nBraDepNo = getSYFlowDef.GetBraDepNoByUseridFlownameStepno(workingUserId.userId, sFlowName, sFirstStep)
            End If

            '取得新的Caseid
            If String.IsNullOrEmpty(sCaseid) = False Then
                newCaseid = sCaseid
            Else
                newCaseid = getSYCaseId().GetNewCaseId(sSubsysid, nBraDepNo, "0")
            End If


            '將新caseid寫入至SY_CASEID內
            bResult = getSYCaseId().Insert(sSysid, sSubsysid, nFlowId, newCaseid, nBraDepNo, nBraDepNo)
            Dbg.Assert(bResult)
            If bResult = False Then
                Throw New SYException(
                    "caseid無法寫入至SY_CASEID",
                    SYMSG.FLOW_CASEID_CANNOT_WRITTEN,
                    getSYCaseId().GetLastSQL)
            End If

            Return newCaseid

        Catch ex As SYException
            Throw
        Catch ex As Exception
            Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
        End Try

    End Function



    ''' <summary>
    ''' 傳入案號==>起案
    ''' </summary>
    ''' <param name="sCaseId">案號</param>
    ''' <param name="bAutoSendToNext"></param>
    ''' <param name="assignNextStepUser">可指定多個(分支)下個步驟的負責人員</param>
    ''' <param name="bSendToEnd">自動送至END，例如，START->0000->END，適用只有一個步驟的案件(起案按確認送出後結案)</param>
    ''' <param name="sApvCasId">核准編號</param>
    ''' <param name="sAppBrid">申請單位</param>
    ''' <param name="sCplAplId">授信戶編號</param>
    ''' <param name="sCplAplNam">授信戶名稱</param>
    ''' <param name="bSendMail">sWorkingUserId及sLoginUserId若沒有特別的理由，應設為""由FLOW程式自行判斷</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function StartFlowByCaseId(ByVal sCaseId As String,
                                      Optional ByVal bAutoSendToNext As Boolean = False,
                                      Optional ByVal assignNextStepUser() As STEPNO_USERID = Nothing,
                                      Optional ByVal bSendToEnd As Boolean = False,
                                      Optional ByVal sApvCasId As String = "",
                                      Optional ByVal sAppBrid As String = "",
                                      Optional ByVal sCplAplId As String = "",
                                      Optional ByVal sCplAplNam As String = "",
                                      Optional ByVal bSendMail As Boolean = True) As StepInfo

        Dim stepInfo As StepInfo
        Dim stepInfoItemExt As New StepInfoItemExt
        Dim drCaseid As DataRow
        Dim drFlowid As DataRow
        Dim nBraDepNo As Integer


        Dim sLoginUser As New USER_ID_NAME
        Dim sWorkingUser As New USER_ID_NAME

        Try

            '設核准編號, 申請單位, 授信戶編號, 授信戶名稱

            If String.IsNullOrEmpty(sApvCasId) = False OrElse
                String.IsNullOrEmpty(sAppBrid) = False OrElse
                String.IsNullOrEmpty(sCplAplId) = False OrElse
                String.IsNullOrEmpty(sCplAplNam) = False Then
                getSYCaseId.SetCaseInfo(sCaseId, sApvCasId, sAppBrid, sCplAplId, sCplAplNam)
            End If

            stepInfoItemExt.caseId = sCaseId

            '從案號取BRA_DEPNO
            drCaseid = getSYCaseId().GetInfo(sCaseId)
            If IsNothing(drCaseid) Then
                Throw New SYException(
                    String.Format("無法取得案件[{0}]的記錄，SQL={1}", stepInfoItemExt.caseId, getSYCaseId.GetLastSQL),
                    SYMSG.FLOW_CASEINFO_NOT_FOUND,
                     getSYCaseId.GetLastSQL)
            End If

            nBraDepNo = CInt(drCaseid("BRA_DEPNO"))
            stepInfoItemExt.braDepNo_brid = SY_TABLEBASE.ToArray(
                BRADEPNO_BRID.Pair(nBraDepNo, getSYBranch.GetBRID(nBraDepNo)))
            'stepInfoItemExt.bridDepNo = CInt(drCaseid("BRA_DEPNO"))

            '取得WORKINGUSER及LOGINUSER，或檢查傳入的WORKINGUSER及LOGINUSER是否是正確的使用者
            GetCurrentUserid(sLoginUser, sWorkingUser)

            stepInfoItemExt.caseOwner = sWorkingUser
            stepInfoItemExt.caseClient = sLoginUser
            stepInfoItemExt.caseSender = sLoginUser


            stepInfoItemExt.flowId = CInt(drCaseid("FLOW_ID"))

            '由FLOWID 從SY_FLOW_ID取得FLOWNAME
            drFlowid = getSYFlowId.GetDataRow(BosBase.PARAMETER("FLOW_ID", stepInfoItemExt.flowId))
            If IsNothing(drFlowid) Then
                Throw New SYException(
                    String.Format("無法取得SY_FLOW_ID的資料[FLOW_ID={0}]", stepInfoItemExt.flowId),
                    SYMSG.FLOW_FLOWINFO_NOT_FOUND,
                    getSYFlowId.GetLastSQL)
            End If

            stepInfoItemExt.flowName = CStr(drFlowid("FLOW_NAME"))

            '取得開始的流程步驟
            stepInfoItemExt.stepNo = getSYFlowMap().GetFirstStepNo(PARAMETER("FLOW_ID", stepInfoItemExt.flowId))

            '初始流程
            stepInfoItemExt.subflowCount = 1
            stepInfoItemExt.subflowSeq = 1
            stepInfoItemExt.subflowLevel = 1
            stepInfoItemExt.parallelNo = 1
            stepInfoItemExt.flowstepStatus = 1
            stepInfoItemExt.stepSummary = String.Empty
            stepInfoItemExt.processTime = DateTime.MinValue
            stepInfoItemExt.parentSubFlowSeq = 0
            stepInfoItemExt.previousSubFlowSeq = 0

            '檢查是否重啟案件
            If getSYFlowIncident.GetCount(BosBase.PARAM_ARRAY("CASEID", sCaseId)) > 0 Then

                '案件是否已結束
                If getSYFlowIncident.IsCaseCompleted(sCaseId) = False Then
                    Throw New SYException(
                        String.Format("案件[{0}]未結束，不可以重新開啟案件!", sCaseId),
                        SYMSG.FLOW_CASE_CANNOT_RESET,
                        getSYFlowIncident.GetLastSQL)
                End If

                stepInfoItemExt.subflowCount = getSYFlowIncident.GetMaxSubFlowCount(sCaseId) + 1
                stepInfoItemExt.subflowSeq = getSYFlowIncident.GetMaxSubFlowSeq(sCaseId) + 1

            End If


            '寫入新案件至FlowIncident()
            getSYFlowIncident().StartNewFlowIncident(stepInfoItemExt)

            '寫入新案件至FlowStep()
            getSYFlowStep().InsertUpdateRecord(stepInfoItemExt, "I")

            '送下一關

            If bSendToEnd Then
                stepInfo = SendFlow(stepInfoItemExt.caseId,
                                    "N",
                                    stepInfoItemExt.subflowSeq,
                                    SY_TABLEBASE.ToArray(STEPNO_BRADEPNO.Pair(0, SY_STEP_NO.SystemStep_End)),
                                    assignNextStepUser,
                                    "",
                                    False,
                                    True,
                                    True,
                                    False)

            Else
                stepInfo = SendFlow(stepInfoItemExt.caseId,
                                    "N",
                                    stepInfoItemExt.subflowSeq,
                                    Nothing,
                                    assignNextStepUser,
                                    "",
                                    False,
                                    bAutoSendToNext,
                                    True,
                                    False)
            End If

            '起案後寄發郵件
            If bSendMail = True Then
                Try
                    If IsNothing(m_callbackFlowMail) = False AndAlso bSendMail = True Then
                        m_callbackFlowMail.Notify(getDatabaseManager(), stepInfo.nextStepInfo)
                    End If
                Catch ex As Exception
                End Try
            End If

            Return stepInfo

        Catch ex As SYException
            Throw
        Catch ex As Exception
            Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
        End Try

    End Function


    ''' <summary>
    ''' 送至下一步驟
    ''' </summary>
    ''' <param name="sCaseId">案件編號</param>
    ''' <param name="sPN">方向，前進及後退</param>
    ''' <param name="nSubFlowSeq">子流程序號，用來區分子流程及退回流程</param>
    ''' <param name="assignNextStepBradepNo">指派流程送至下一步驟</param>
    ''' <param name="assignNextStepUser">指派接下來流程至哪一步驟時的使用者是誰</param>
    ''' <param name="sRevisionNote">修正補充序號</param>
    ''' <param name="bOnEvent">是否是流程事件，流程事件</param>
    ''' <param name="bSendToNext">是否可送至下一個實際步驟，若否則表示只能送至虛擬步驟</param>
    ''' <param name="bSendMail">是否要送信至下個步驟的USER</param> 
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SendFlow(ByVal sCaseId As String,
                             Optional ByVal sPN As String = "N",
                             Optional ByVal nSubFlowSeq As Integer = 0,
                             Optional ByVal assignNextStepBradepNo() As STEPNO_BRADEPNO = Nothing,
                             Optional ByVal assignNextStepUser() As STEPNO_USERID = Nothing,
                             Optional ByVal sRevisionNote As String = "",
                             Optional ByVal bOnEvent As Boolean = False,
                             Optional ByVal bSendToNext As Boolean = True,
                             Optional ByVal bSendtoSubOnSplitter As Boolean = True,
                             Optional ByVal bSendMail As Boolean = True) As StepInfo

        Dim returnStepInfo As StepInfo


        ''移除離職員工
        If IsNothing(assignNextStepUser) = False Then
            Dim listNextStepUser As New List(Of STEPNO_USERID)

            For Each item In assignNextStepUser
                If getSYUser.IsValidUser(item.sUserId) Then
                    listNextStepUser.Add(item)
                End If
            Next

            assignNextStepUser = listNextStepUser.ToArray
        End If


        returnStepInfo = InnerSendFlow(sCaseId,
                             sPN,
                             nSubFlowSeq,
                             assignNextStepBradepNo,
                             assignNextStepUser,
                             sRevisionNote,
                             bOnEvent,
                             bSendToNext,
                             bSendtoSubOnSplitter
                             )

        Try
            If IsNothing(m_callbackFlowMail) = False AndAlso bSendMail = True Then
                m_callbackFlowMail.Notify(getDatabaseManager(), returnStepInfo.nextStepInfo)
            End If
        Catch ex As Exception
        End Try

        Return returnStepInfo

    End Function


    ''' <summary>
    ''' 送至下一步驟
    ''' </summary>
    ''' <param name="sCaseId">案件編號</param>
    ''' <param name="sPN">方向，前進及後退</param>
    ''' <param name="nSubFlowSeq">子流程序號，用來區分子流程及退回流程</param>
    ''' <param name="assignNextStepBradepNo">指派流程送至下一步驟</param>
    ''' <param name="assignNextStepUser">指派接下來流程至哪一步驟時的使用者是誰</param>
    ''' <param name="sRevisionNote">修正補充序號</param>
    ''' <param name="bOnEvent">是否是流程事件，流程事件</param>
    ''' <param name="bSendToNext">是否可送至下一個實際步驟，若否則表示只能送至虛擬步驟</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function InnerSendFlow(ByVal sCaseId As String,
                             Optional ByVal sPN As String = "N",
                             Optional ByVal nSubFlowSeq As Integer = 0,
                             Optional ByVal assignNextStepBradepNo() As STEPNO_BRADEPNO = Nothing,
                             Optional ByVal assignNextStepUser() As STEPNO_USERID = Nothing,
                             Optional ByVal sRevisionNote As String = "",
                             Optional ByVal bOnEvent As Boolean = False,
                             Optional ByVal bSendToNext As Boolean = True,
                             Optional ByVal bSendtoSubOnSplitter As Boolean = True) As StepInfo

        Dim nEventCount As Integer = 0

RESTART:
        Try
            Dim stepInfo As New StepInfo
            Dim currentStepInfoItemExt As StepInfoItemExt
            Dim nextStepInfoItemExtArray() As StepInfoItemExt = Nothing

            Dim sCurrentFlowStepType As String = Nothing

            Dim eventStepInfo As StepInfo = Nothing

            Dim loginUser As New USER_ID_NAME
            Dim workingUser As New USER_ID_NAME
            Dim drCaseRevision As DataRow = Nothing

            '流程事件觸發不超過10次
            nEventCount = nEventCount + 1
            If nEventCount > 10 Then
                Throw New SYException(
                    "流程事件觸發不超過10次",
                    SYMSG.FLOW_TOO_MANY_EVENTS_FIRED,
                    "N/A")
            End If


            Dbg.Assert(sPN = "N" OrElse sPN = "P")

            '流程案件是否已結束，若結束則無法再送出案件
            If getSYFlowIncident().IsCaseCompleted(sCaseId, nSubFlowSeq) Then
                Throw New SYException(
                    String.Format("案件已結案，無法再送出案件: CASEID={0}, SUBFLOW_SEQ={1}", sCaseId, nSubFlowSeq),
                    SYMSG.FLOW_CASE_CANNOT_SEND_BECAUSE_CASE_CLOSED,
                    getSYFlowIncident.GetLastSQL())
            End If

            GetCurrentUserid(loginUser, workingUser)

            '取得目前流程步驟的相關資訊
            currentStepInfoItemExt = getSYFlowStep().GetInProgressStepInfoItem(sCaseId, nSubFlowSeq, workingUser.userId)

            'Dbg.Assert(currentStepInfoItemExt.stepNo <> "04100601")

            '如果是虛擬步驟，將OWNER及CLIENT設為登入者
            If IsNothing(currentStepInfoItemExt.caseClient) AndAlso
                 (getSYFlowMap().GetStepType(currentStepInfoItemExt.flowId, currentStepInfoItemExt.stepNo) = "N" OrElse
                 getSYFlowMap().GetStepType(currentStepInfoItemExt.flowId, currentStepInfoItemExt.stepNo) = "P" OrElse
                 getSYFlowMap().GetStepType(currentStepInfoItemExt.flowId, currentStepInfoItemExt.stepNo) = "V") Then
                currentStepInfoItemExt.caseOwner = loginUser
                currentStepInfoItemExt.caseClient = loginUser
            End If

            '依步驟類別重設往前或往後
            If currentStepInfoItemExt.stepType = "N" AndAlso sPN = "P" Then
                sPN = "N"
            End If

            If currentStepInfoItemExt.stepType = "P" AndAlso sPN = "N" Then
                sPN = "P"
            End If


            '如果目前步驟不是END，取得下一流程流程步驟的相關資訊
            If String.Compare(currentStepInfoItemExt.stepNo, SY_STEP_NO.SystemStep_End, True) <> 0 Then

                Dim stepInfoItem As StepInfoItemExt
                stepInfoItem = currentStepInfoItemExt.clone()
                stepInfoItem.caseSender = loginUser    '為了檢查下一步驟的使用者跟目前步驟的使用者不可相同


                '下一步驟為SPLITTER，且未指定任何下一步驟的分行及使用者，先送至SPLITEER再繼續。
                If (IsNothing(assignNextStepBradepNo) = False OrElse IsNothing(assignNextStepUser) = False) AndAlso
                    isNextStep_a_Splitter(stepInfoItem, sPN) = True Then

                    InnerSendFlow(sCaseId,
                             sPN,
                             nSubFlowSeq,
                             Nothing,
                             Nothing,
                             sRevisionNote,
                             False,
                             True,
                             False)

                    GoTo RESTART
                End If

                'Dbg.Assert(currentStepInfoItemExt.stepNo <> "04100000")

                If bSendToNext = True AndAlso
                    String.Compare(currentStepInfoItemExt.stepNo, SY_STEP_NO.SystemStep_Start, True) = 0 AndAlso
                    IsNothing(nextStepInfoItemExtArray) Then
                    '起案即第一步驟並送至第二步驟，第一步驟使用者設為登入者
                    nextStepInfoItemExtArray = GetNextStep(stepInfoItem, sPN, assignNextStepBradepNo, STEPNO_USERID.arrayStepUser(Nothing, loginUser.userId))
                Else
                    nextStepInfoItemExtArray = GetNextStep(stepInfoItem, sPN, assignNextStepBradepNo, assignNextStepUser)
                End If

                'If nextStepInfoItemExtArray.Count = 1 Then
                '    Dbg.Assert(nextStepInfoItemExtArray(0).stepNo <> "SPLITTER2")
                'End If

                'Dbg.Assert(nextStepInfoItemExtArray.Count() = 1)

                '如果不是步驟不是END且沒有下一步驟，則顯示錯誤
                If IsNothing(nextStepInfoItemExtArray) OrElse
                    nextStepInfoItemExtArray.Count = 0 Then

                    Dbg.Assert(False, "下一步會產生EXCEPTION")
                    Throw New SYException(
                        String.Format("無法取得流程:{0} 步驟:{1} 的下一步驟，SQL={2}", currentStepInfoItemExt.flowId, currentStepInfoItemExt.stepNo, BosBase.GetGlobalLastSQL),
                        SYMSG.FLOW_NEXTSTEP_NOT_FOUND,
                        BosBase.GetGlobalLastSQL)
                End If

                If nextStepInfoItemExtArray.Count > 1 AndAlso _
                    getSYFlowMap().GetStepType(currentStepInfoItemExt.flowId, currentStepInfoItemExt.stepNo) <> "S" Then
                    Throw New SYException(
                        String.Format("目前步驟不能有兩個以上的下一步驟，目前CASIED：{0}，目前步驟：{1}", currentStepInfoItemExt.caseId, currentStepInfoItemExt.stepNo),
                        SYMSG.FLOW_NEXTSTEP_NOT_FOUND_AMBIGUOUS,
                        BosBase.GetGlobalLastSQL)
                End If

            End If

            stepInfo.currentStepInfo = currentStepInfoItemExt
            stepInfo.nextStepInfo = nextStepInfoItemExtArray


            '觸發流程事件，若已經觸發，則不再重複
            If bOnEvent = False Then
                sCurrentFlowStepType = getSYFlowMap().GetStepType(currentStepInfoItemExt.flowId, currentStepInfoItemExt.stepNo)

                If String.Compare(currentStepInfoItemExt.stepNo, SY_STEP_NO.SystemStep_Start, True) = 0 OrElse
                    sCurrentFlowStepType = "N" OrElse
                    sCurrentFlowStepType = "P" Then

                    'Dim nextStepUser() As STEPNO_USERID
                    'nextStepUser = STEPNO_USERID.arrayStepUser("", sWorkingUser)

                    eventStepInfo = OnFlowEvent(stepInfo, Nothing, assignNextStepUser, bSendToNext)

                    If IsNothing(eventStepInfo) = False AndAlso bSendToNext = True Then
                        GoTo RESTART
                    End If
                End If
            End If


            '若不送至下一步驟，返回。若是虛擬步驟，仍會繼續執行
            If bSendToNext = False Then
                sCurrentFlowStepType = getSYFlowMap().GetStepType(currentStepInfoItemExt.flowId, currentStepInfoItemExt.stepNo)

                If String.IsNullOrEmpty(sCurrentFlowStepType) = True AndAlso
                    String.Compare(currentStepInfoItemExt.stepNo, SY_STEP_NO.SystemStep_Start, True) <> 0 Then

                    Array.Resize(stepInfo.nextStepInfo, 1)
                    stepInfo.nextStepInfo(0) = stepInfo.currentStepInfo

                    Return stepInfo
                End If
            End If


            '檢查是否有caseClient
            Dbg.Assert(Not IsNothing(currentStepInfoItemExt.caseClient))

            If IsNothing(currentStepInfoItemExt.caseClient) Then
                Throw New SYException(
                    String.Format("目前步驟沒有可以送件的使用者，目前CASIED：{0}，目前步驟：{1}", currentStepInfoItemExt.caseId, currentStepInfoItemExt.stepNo),
                    SYMSG.FLOW_CLIENT_NOT_FOUND_IN_CURRENT_STEP,
                    BosBase.GetGlobalLastSQL)
            End If

            If String.IsNullOrEmpty(currentStepInfoItemExt.caseOwner.userId) Then
                currentStepInfoItemExt.caseOwner = loginUser '虛擬步驟才會出現caseOwner是空的，將重指定caseOwner=sLoginUser
            End If


            '更新現在的步驟
            currentStepInfoItemExt.flowstepStatus = 3
            currentStepInfoItemExt.processTime = Now
            currentStepInfoItemExt.caseSender = loginUser
            currentStepInfoItemExt.stepSummary = Nothing
            'currentStepInfoItemExt.caseClient = sLoginUser

            'Dbg.Assert(Not (currentStepInfoItemExt.subflowSeq = 1 And
            '                currentStepInfoItemExt.subflowCount = 4 And
            '                currentStepInfoItemExt.stepNo = "S6"))

            getSYFlowStep().InsertUpdateRecord(currentStepInfoItemExt, "U")

            '如果已經是最後一個步驟，表示流程已經完成，寫入3至FlowIncident
            If String.Compare(currentStepInfoItemExt.stepNo, SY_STEP_NO.SystemStep_End, True) = 0 Then
                getSYFlowIncident().CloseFlowIncident(currentStepInfoItemExt)
                Return stepInfo
            End If

            '如果下一步驟是JOINER，結束目前子流程，等待所有子流程結束後繼續父流程
            If (Not IsNothing(nextStepInfoItemExtArray)) AndAlso
               nextStepInfoItemExtArray.Count > 0 Then

                Dim drNextFlowStep As DataRow =
                    getSYFlowMap().GetStepInfo(nextStepInfoItemExtArray(0).flowId, nextStepInfoItemExtArray(0).stepNo)

                If SY_TABLEBASE.CDbType(Of String)(drNextFlowStep("TYPE"), "") = "J" Then
                    getSYFlowIncident().CloseFlowIncident(currentStepInfoItemExt)

                    Dim parentFlowInfo As StepInfoItemExt = Nothing
                    Dim bActivateParentFlow As Boolean = False

                    '取得父流程的資料
                    parentFlowInfo = getSYFlowStep().GetStepInfoItem(
                        currentStepInfoItemExt.caseId,
                        currentStepInfoItemExt.parentSubFlowSeq,
                        getSYFlowStep().GetLatestStepNo(currentStepInfoItemExt.caseId,
                                                                      currentStepInfoItemExt.parentSubFlowSeq),
                        Nothing,
                        False)

                    '如果JOINER的TYPE是OR，關閉其它的流程並繼續
                    If SY_TABLEBASE.CDbType(Of String)(drNextFlowStep("SPLITER_TYPE"), "") = "OR" Then

                        getSYFlowIncident().CloseAllSubFlowIncident(parentFlowInfo)
                        bActivateParentFlow = True
                    Else    '否則都是AND，等待其它流程結束才繼續
                        If getSYFlowIncident().GetInProgressSubFlowCount(parentFlowInfo.caseId, parentFlowInfo.subflowSeq) = 0 Then
                            bActivateParentFlow = True
                        Else
                            bActivateParentFlow = False
                        End If
                    End If

                    '完成所有子流程後繼續
                    If bActivateParentFlow = True Then
                        If IsNothing(parentFlowInfo) Then
                            parentFlowInfo = getSYFlowStep().GetStepInfoItem(
                                currentStepInfoItemExt.caseId,
                                currentStepInfoItemExt.parentSubFlowSeq,
                                getSYFlowStep().GetLatestStepNo(currentStepInfoItemExt.caseId,
                                                                              currentStepInfoItemExt.parentSubFlowSeq),
                                Nothing,
                                False)
                        End If

                        Dbg.Assert(nextStepInfoItemExtArray.Count = 1)
                        If nextStepInfoItemExtArray.Count <> 1 Then
                            Throw New SYException(
                                String.Format("目前步驟不能有兩個以上的下一步驟，目前CASIED：{0}，目前步驟：{1}", currentStepInfoItemExt.caseId, currentStepInfoItemExt.stepNo),
                                SYMSG.FLOW_NEXTSTEP_NOT_FOUND_AMBIGUOUS,
                                BosBase.GetGlobalLastSQL)
                        End If

                        Dim oldStepInfoItemext As StepInfoItemExt = Nothing
                        oldStepInfoItemext = nextStepInfoItemExtArray(0)
                        nextStepInfoItemExtArray(0) = parentFlowInfo.clone

                        nextStepInfoItemExtArray(0).stepNo = oldStepInfoItemext.stepNo
                        nextStepInfoItemExtArray(0).stepType = oldStepInfoItemext.stepType
                        nextStepInfoItemExtArray(0).caseSender = Nothing

                        Dim dr As DataRow = getSYFlowStep().GetSubFlowInfoMax(sCaseId)
                        nextStepInfoItemExtArray(0).subflowCount = CInt(dr("MAX_SUBFLOW_COUNT"))

                    Else
                        GoTo LAST_EVENT
                    End If

                End If
            End If

            '如果已經是PN='P"，表示流程已經重新，重新一子流程 SUBFLOW_SEQ=MAX(SUBFLOW_SEQ)+1
            If sPN = "P" Then
                'Dim currentDate As Date = Now

                '修正補充
                If Not String.IsNullOrEmpty(sRevisionNote) Then

                    'Dbg.Assert(IsNothing(nextStepInfoItemExtArray(0).caseOwner))

                    Dim sRefundOwnerId As String = Nothing
                    If IsNothing(nextStepInfoItemExtArray(0).caseOwner) = False Then
                        sRefundOwnerId = nextStepInfoItemExtArray(0).caseOwner.userId
                    End If

                    drCaseRevision = getSYCaseRevision().InsertRevision(
                        currentStepInfoItemExt.caseId,
                        currentStepInfoItemExt.stepNo,
                        sRevisionNote,
                        nextStepInfoItemExtArray(0).stepNo,
                        currentStepInfoItemExt.caseSender.userId,
                        currentStepInfoItemExt.braDepNo_brid(0).BraDepNo,
                        sRefundOwnerId)

                    If IsNothing(drCaseRevision) OrElse IsDBNull(drCaseRevision("SEQNO")) Then
                        Throw New SYException(
                            String.Format("無法從SY_CASEREVISION取得資料，CASEID={0}", sCaseId),
                            SYMSG.FLOW_CASEREVISION_NOT_FOUND,
                            BosBase.GetGlobalLastSQL)
                    End If

                    currentStepInfoItemExt.revisionSeqNo = CInt(drCaseRevision("SEQNO"))

                    getSYFlowStep().InsertUpdateRecord(currentStepInfoItemExt, "U")
                End If


                '結束舊CASEID, SUBFLOWSEQ的案件
                getSYFlowIncident().CloseFlowIncident(currentStepInfoItemExt)

                '建立新CASEID, SUBFLOWSEQ的案件
                Dim nOldSubFlowSeq As Integer = currentStepInfoItemExt.subflowSeq
                Dim nNewSubFlowSeq As Integer = getSYFlowIncident().GetMaxSubFlowSeq(currentStepInfoItemExt.caseId) + 1

                Dim newStepInfoItemExt As StepInfoItemExt = currentStepInfoItemExt.clone
                newStepInfoItemExt.caseSender = currentStepInfoItemExt.caseSender
                newStepInfoItemExt.previousSubFlowSeq = nOldSubFlowSeq
                newStepInfoItemExt.subflowSeq = nNewSubFlowSeq      '使用新SubFlowSeq起新子流程
                getSYFlowIncident().StartNewFlowIncident(newStepInfoItemExt)
                'currentStepInfoItemExt.subflowSeq = nOldSubFlowSeq      '復原SubFlowSeq

                For Each item As StepInfoItemExt In nextStepInfoItemExtArray
                    item.subflowSeq = nNewSubFlowSeq
                    item.previousSubFlowSeq = nOldSubFlowSeq
                Next
            End If

            '更新下一步驟，可能有多筆同時進行
            'Dbg.Assert(nextStepInfoItemExtArray.Count = 1)
            Dim nCount As Integer = 0
            For Each item As StepInfoItemExt In nextStepInfoItemExtArray
                item.flowstepStatus = 1
                nCount = nCount + 1

                sCurrentFlowStepType = getSYFlowMap().GetStepType(currentStepInfoItemExt.flowId, currentStepInfoItemExt.stepNo)

                '如果目前是分支，開始多個子流程
                If sCurrentFlowStepType = "S" Then
                    Dim dr As DataRow
                    dr = getSYFlowStep().GetSubFlowInfoMax(currentStepInfoItemExt.caseId)

                    item.subflowSeq = CInt(dr("MAX_SUBFLOWSEQ")) + 1
                    item.subflowCount = CInt(dr("MAX_SUBFLOW_COUNT")) + 1
                    item.subflowLevel = currentStepInfoItemExt.subflowLevel + 1
                    item.parallelNo = nCount
                    item.parentSubFlowSeq = currentStepInfoItemExt.subflowSeq
                    item.previousSubFlowSeq = 0

                    'item.caseOwner = item.caseOwner
                    item.caseSender = loginUser
                    getSYFlowIncident.StartNewFlowIncident(item)
                    item.caseSender = Nothing
                End If

                Dim sNextStepType As String
                sNextStepType = getSYFlowMap().GetStepType(item.flowId, item.stepNo)

                '若是下一步驟分支
                If sNextStepType = "S" Then
                    item.caseOwner = currentStepInfoItemExt.caseOwner
                    item.caseClient = currentStepInfoItemExt.caseClient

                    GoTo INSERT_UPDATE_RECORD
                End If

                '由角色查出的人是否只有一位，若只有一位，就不需使用佇列取件分派案件
                If (Not IsNothing(item.caseOwnerList)) AndAlso item.caseOwnerList.Count = 1 Then
                    item.caseOwner = item.caseOwnerList(0)
                    item.caseClient = item.caseOwnerList(0)

                    GoTo INSERT_UPDATE_RECORD
                End If

                '若下一步驟是最後一步驟(END)
                If String.Compare(nextStepInfoItemExtArray(0).stepNo, SY_STEP_NO.SystemStep_End, True) = 0 Then
                    item.caseOwner = currentStepInfoItemExt.caseOwner
                    item.caseClient = currentStepInfoItemExt.caseClient

                    If IsNothing(nextStepInfoItemExtArray(0).braDepNo_brid) Then
                        nextStepInfoItemExtArray(0).braDepNo_brid = currentStepInfoItemExt.braDepNo_brid
                    End If

                    GoTo INSERT_UPDATE_RECORD
                End If

                '若目前步驟是最先一步驟(START)
                If String.Compare(currentStepInfoItemExt.stepNo, SY_STEP_NO.SystemStep_Start, True) = 0 Then

                    If bSendToNext Then
                        item.caseOwner = currentStepInfoItemExt.caseOwner
                        item.caseClient = currentStepInfoItemExt.caseClient
                    Else
                        item.caseOwner = Nothing
                        item.caseClient = Nothing

                        Dim arrayNextStepInfoItemExt2() As StepInfoItemExt
                        Dim arrayNextStep(0) As String

                        arrayNextStep(0) = item.stepNo
                        arrayNextStepInfoItemExt2 = GetNextStep(currentStepInfoItemExt,
                                                                sPN,
                                                                assignNextStepBradepNo,
                                                                assignNextStepUser)

                        If Not IsNothing(arrayNextStepInfoItemExt2) Then
                            item.caseClient = arrayNextStepInfoItemExt2(0).caseClient
                            item.caseOwner = arrayNextStepInfoItemExt2(0).caseOwner
                            item.caseOwnerList = arrayNextStepInfoItemExt2(0).caseOwnerList
                        End If
                    End If

                    GoTo INSERT_UPDATE_RECORD
                End If

INSERT_UPDATE_RECORD:

                '如果下一步驟的USER不在自己的分行內，則顯示錯誤
                If IsNothing(item.caseClient) = False Then
                    If getSYRelBranchUser.GetCount(PARAM_ARRAY("BRA_DEPNO", item.braDepNo_brid(0).BraDepNo, "STAFFID", item.caseClient.userId)) = 0 AndAlso
                        String.IsNullOrEmpty(item.stepType) = True Then
                        Dbg.Assert(False, "下一步驟的USER不在自己的分行內")
                        Throw New SYException(
                            String.Format("下一步驟的USER不在自己的分行內，CASIED：{0}，步驟：{1}，使用者：{2}，分行代碼：{3}", item.caseId, item.stepNo, item.caseClient.userId, item.braDepNo_brid(0).BraDepNo),
                            SYMSG.FLOW_NEXTSTEP_CLIENT_NOT_IN_HIS_BRANCH,
                            BosBase.GetGlobalLastSQL)
                    End If
                End If

                getSYFlowStep().InsertUpdateRecord(item)
            Next

LAST_EVENT:
            '觸發流程事件

            If bSendtoSubOnSplitter = False AndAlso IsNothing(nextStepInfoItemExtArray) = False AndAlso
                nextStepInfoItemExtArray.Count = 1 AndAlso nextStepInfoItemExtArray(0).stepType = "S" Then
                Return stepInfo
            End If

            'Dbg.Assert(stepInfo.currentStepInfo.stepNo <> "JOINER2")
            eventStepInfo = OnFlowEvent(stepInfo, assignNextStepBradepNo, assignNextStepUser, bSendToNext)

            If Not IsNothing(eventStepInfo) Then
                stepInfo.nextStepInfo = eventStepInfo.nextStepInfo
            End If
            'End If

            'Dbg.Assert(IsNothing(stepInfo.nextStepInfo) = False)

            Return stepInfo

        Catch ex As SYException
            Throw
        Catch ex As Exception
            Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
        End Try

    End Function

    ''' <summary>
    ''' 子流程完成後再追加執行子流程
    ''' </summary>
    ''' <param name="sCaseid"></param>
    ''' <param name="nParentSubFlowSeq"></param>
    ''' <param name="sNextStep"></param>
    ''' <param name="nNextBraDepNo"></param>
    ''' <param name="sPreviousStep"></param>
    ''' <param name="assignUser"></param>
    ''' <param name="bStartFromSplitter"></param>
    ''' <param name="sRevisionNote">修正補充</param>
    ''' <param name="sRevisionSrcStep"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ExecuteAdditionalSubflow(ByVal sCaseid As String,
                                             ByVal nParentSubFlowSeq As Integer,
                                             ByVal sNextStep As String,
                                             Optional ByVal nNextBraDepNo As Integer = 0,
                                             Optional ByVal sPreviousStep As String = Nothing,
                                             Optional ByVal assignUser As String = Nothing,
                                             Optional ByVal bStartFromSplitter As Boolean = True,
                                             Optional ByVal sRevisionNote As String = "",
                                             Optional ByVal sRevisionSrcStep As String = "") As StepInfoItemExt
        Dim dr1 As DataRow = Nothing
        Dim dr2 As DataRow = Nothing
        Dim sStartNo As String = Nothing
        Dim stepItemInfoExt As StepInfoItemExt
        Dim loginUser, workingUser As USER_ID_NAME
        Dim drcBradepnoUser As DataRowCollection

        'Dim assignedUserInfo As USER_ID_NAME

        Try
            'getSYFlowIncident().ReActivateClosedFlowIncident(sCaseid, nSubFlowSeq)

            'assignedUserInfo = getSYUser.GetUserIdName(assignUser)

            '取得WORKINGUSER及LOGINUSER，或檢查傳入的WORKINGUSER及LOGINUSER是否是正確的使用者
            loginUser = Nothing
            workingUser = Nothing
            GetCurrentUserid(loginUser, workingUser)

            getSYFlowStep.CloseByCaseidSubflowseq(sCaseid, nParentSubFlowSeq, loginUser.userId)

            'If getSYFlowIncident.IsCaseCompleted(sCaseid, nParentSubFlowSeq) Then
            '    'Throw New Exception(String.Format("案件已關閉，無法再追加執行一子流程(流程:{0}，子流程序號:{1})",
            '    '                                  sCaseid, nParentSubFlowSeq))
            'End If
            '將父流程狀態重設為1
            getSYFlowIncident.SetCaseStatus(sCaseid, nParentSubFlowSeq, "1")


            dr1 = getSYFlowIncident.GetLatestSubflowInfo(sCaseid, nParentSubFlowSeq)
            If IsNothing(dr1) = False Then
                '取得案件第一個步驟
                sStartNo = getSYFlowStep.GetStartStepNo(sCaseid, CInt(dr1("SUBFLOW_SEQ")))

                '取得案件的StepItemInfoExt
                stepItemInfoExt = getSYFlowStep.GetStepInfoItem(sCaseid, CInt(dr1("SUBFLOW_SEQ")), sStartNo, Nothing, False)
            Else
                stepItemInfoExt = getSYFlowStep.GetStepInfoItem(sCaseid, nParentSubFlowSeq, Nothing, Nothing, False)
                stepItemInfoExt.parallelNo = 0
                stepItemInfoExt.previousSubFlowSeq = 0
                stepItemInfoExt.subflowLevel = stepItemInfoExt.subflowLevel + 1
                stepItemInfoExt.parentSubFlowSeq = stepItemInfoExt.subflowSeq
                stepItemInfoExt.stepType = ""
                'Throw New Exception(String.Format("無法再追加執行一子流程(流程:{0}，子流程序號:{1})", sCaseid, nParentSubFlowSeq))
            End If

            If String.IsNullOrEmpty(assignUser) = False Then
                stepItemInfoExt.caseOwner = getSYUser.GetUserIdName(assignUser)
            End If

            stepItemInfoExt.stepNo = sNextStep

            '不指定分行或員編時，從SY_REL_ROLE_FLOWMAP找出第一個可能的分行
            If nNextBraDepNo = 0 AndAlso String.IsNullOrEmpty(assignUser) = True Then
                nNextBraDepNo = QueryValue(Of Integer)(
                    "select BRA_DEPNO from SY_REL_ROLE_FLOWMAP A inner join SY_REL_ROLE_USER B on A.ROLEID=B.ROLEID where STEP_NO=@STEP_NO@",
                    0,
                    "STEP_NO", sNextStep)
            End If

            If nNextBraDepNo <> 0 Then
                stepItemInfoExt.braDepNo_brid = SY_TABLEBASE.ToArray(
                    BRADEPNO_BRID.Pair(nNextBraDepNo, getSYBranch.GetBRID(nNextBraDepNo)))
            Else
                '跨單位
                Dim spanBranch() As BRADEPNO_BRID
                spanBranch = getSYRelFlowMap_SpanBranch.GetSpanBranch(stepItemInfoExt.flowId, sNextStep)

                If Not IsNothing(spanBranch) Then
                    stepItemInfoExt.braDepNo_brid = spanBranch
                End If
            End If

            dr2 = getSYFlowStep().GetSubFlowInfoMax(sCaseid)

            stepItemInfoExt.subflowSeq = CInt(dr2("MAX_SUBFLOWSEQ")) + 1
            stepItemInfoExt.subflowCount = CInt(dr2("MAX_SUBFLOW_COUNT")) + 1
            'stepItemInfoExt.subflowLevel = CInt(dr1("SUBFLOW_LEVEL")) 'currentStepInfoItemExt.subflowLevel + 1
            stepItemInfoExt.parallelNo = stepItemInfoExt.parallelNo + 1


            If String.IsNullOrEmpty(assignUser) = True Then
                Dim sii As StepInfoItemExt
                Dim siis() As StepInfoItemExt
                Dim astepNo_braDepNo(0) As STEPNO_BRADEPNO
                sii = GetCurrentStepInfoOfCaseid(sCaseid, nParentSubFlowSeq)

                If String.IsNullOrEmpty(sPreviousStep) Then
                    sii.stepNo = getSYFlowStep.GetStepNameByCaseidSubflowseq(sCaseid, nParentSubFlowSeq)
                Else
                    sii.stepNo = sPreviousStep
                End If

                astepNo_braDepNo(0) = STEPNO_BRADEPNO.Pair(0, sNextStep)
                siis = GetNextStep(sii, "N", astepNo_braDepNo)

                If IsNothing(siis) = False Then
                    'Dim drcBradepnoUser As DataRowCollection = Nothing

                    For Each sii In siis
                        If IsNothing(sii.caseOwner) = False Then
                            assignUser = sii.caseOwner.userId

                            drcBradepnoUser = getSYRelRoleFlowMap.GetUserBradepno(
                                sii.flowId, sii.stepNo,
                                sii.caseOwner.userId, sii.braDepNo_brid(0).BraDepNo)

                            '目前案件所屬的分行不對，重設分行
                            Debug.Assert(IsNothing(drcBradepnoUser) = False AndAlso drcBradepnoUser.Count > 0,
                                         "案件改分派至其它分行時需指定分行代碼")

                            If IsNothing(drcBradepnoUser) = True OrElse drcBradepnoUser.Count = 0 Then
                                drcBradepnoUser = getSYRelRoleFlowMap.GetUserBradepno(
                                    sii.flowId, sii.stepNo,
                                    sii.caseOwner.userId)

                                Dbg.Assert(IsNothing(drcBradepnoUser) = False)
                                '若找出新的分行，重設，若沒有，不重設
                                If IsNothing(drcBradepnoUser) = False AndAlso drcBradepnoUser.Count > 0 Then
                                    stepItemInfoExt.braDepNo_brid = SY_TABLEBASE.ToArray(
                                        BRADEPNO_BRID.Pair(CInt(drcBradepnoUser(0)("BRA_DEPNO")), getSYBranch.GetBRID(CInt(drcBradepnoUser(0)("BRA_DEPNO")))))
                                End If
                            End If

                            GoTo EXIT_IF
                        End If

                        If IsNothing(sii.caseOwnerList) = False Then
                            For Each userIdName As USER_ID_NAME In sii.caseOwnerList
                                If IsNothing(userIdName) = False Then
                                    assignUser = userIdName.userId
                                    GoTo EXIT_IF
                                End If
                            Next
                        End If
                    Next
                End If
            Else
                'Dim drcBradepnoUser As DataRowCollection = Nothing
                drcBradepnoUser = getSYRelRoleFlowMap.GetUserBradepno(
                     stepItemInfoExt.flowId, stepItemInfoExt.stepNo,
                     stepItemInfoExt.caseOwner.userId, stepItemInfoExt.braDepNo_brid(0).BraDepNo)

                '目前案件所屬的分行不對，重設分行
                Debug.Assert(IsNothing(drcBradepnoUser) = False AndAlso drcBradepnoUser.Count > 0,
                             "案件改分派至其它分行時需指定分行代碼")

                If IsNothing(drcBradepnoUser) = True OrElse drcBradepnoUser.Count = 0 Then
                    drcBradepnoUser = getSYRelRoleFlowMap.GetUserBradepno(
                        stepItemInfoExt.flowId, stepItemInfoExt.stepNo,
                        stepItemInfoExt.caseOwner.userId)

                    Dbg.Assert(IsNothing(drcBradepnoUser) = False)
                    '若找出新的分行，重設，若沒有，不重設
                    If IsNothing(drcBradepnoUser) = False AndAlso drcBradepnoUser.Count > 0 Then
                        stepItemInfoExt.braDepNo_brid = SY_TABLEBASE.ToArray(
                            BRADEPNO_BRID.Pair(CInt(drcBradepnoUser(0)("BRA_DEPNO")), getSYBranch.GetBRID(CInt(drcBradepnoUser(0)("BRA_DEPNO")))))
                    End If
                End If

            End If

EXIT_IF:
            '檢查OWNER是否在BRID內

            'If 


            '    drcBradepnoUser = getSYRelRoleFlowMap.GetUserBradepno(
            '         stepItemInfoExt.flowId, stepItemInfoExt.stepNo,
            '         stepItemInfoExt.caseOwner.userId, stepItemInfoExt.braDepNo_brid(0).BraDepNo)


            stepItemInfoExt.caseOwner = getSYUser.GetUserIdName(assignUser)
            stepItemInfoExt.caseClient = stepItemInfoExt.caseOwner
            stepItemInfoExt.caseSender = loginUser
            stepItemInfoExt.caseOwnerList = Nothing
            stepItemInfoExt.stepSummary = String.Empty
            stepItemInfoExt.processTime = DateTime.MinValue
            stepItemInfoExt.stepNo = sNextStep
            stepItemInfoExt.flowstepStatus = 1


            '修正補充
            If Not String.IsNullOrEmpty(sRevisionNote) Then
                Dim drCaseRevision As DataRow

                Dbg.Assert(IsNothing(sRevisionSrcStep) = False)

                drCaseRevision = getSYCaseRevision().InsertRevision(
                    stepItemInfoExt.caseId,
                    sRevisionSrcStep,
                    sRevisionNote,
                    stepItemInfoExt.stepNo,
                    stepItemInfoExt.caseSender.userId,
                    stepItemInfoExt.braDepNo_brid(0).BraDepNo,
                    stepItemInfoExt.caseOwner.userId)

                If IsNothing(drCaseRevision) OrElse IsDBNull(drCaseRevision("SEQNO")) Then
                    Throw New SYException(
                        String.Format("無法從SY_CASEREVISION取得資料，CASEID={0}", sCaseid),
                        SYMSG.FLOW_CASEREVISION_NOT_FOUND,
                        getSYCaseRevision.GetLastSQL)
                End If

                stepItemInfoExt.revisionSeqNo = CInt(drCaseRevision("SEQNO"))
            End If

            '寫入新案件至FlowIncident()
            getSYFlowIncident().StartNewFlowIncident(stepItemInfoExt)

            '寫入新案件至FlowStep()
            stepItemInfoExt.caseSender = Nothing
            getSYFlowStep().InsertUpdateRecord(stepItemInfoExt, "I")

            Try
                If IsNothing(m_callbackFlowMail) = False Then
                    Dim arrayStepItemInfoExt(0) As FLOW_OP.StepInfoItemExt
                    arrayStepItemInfoExt(0) = stepItemInfoExt
                    m_callbackFlowMail.Notify(getDatabaseManager(), arrayStepItemInfoExt)
                End If
            Catch ex As Exception
            End Try

            Return stepItemInfoExt

        Catch ex As SYException
            Throw
        Catch ex As Exception
            Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
        End Try

    End Function


    ''' <summary>
    ''' 流程特殊事件觸發
    ''' </summary>
    ''' <param name="currentStepInfo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Function OnFlowEvent(ByVal currentStepInfo As StepInfo,
                                   Optional ByVal assignNextStepBradepNo() As STEPNO_BRADEPNO = Nothing,
                                   Optional ByVal assignNextStepUser() As STEPNO_USERID = Nothing,
                                   Optional ByVal bSendToNext As Boolean = True) As StepInfo

        Dim stepinfo As StepInfo = Nothing
        Dim tempStepInfo As StepInfo
        Dim drc As DataRowCollection

        Try
            drc = getSYFlowStep().GetIncompletedStepNoByCaseid(currentStepInfo.currentStepInfo.caseId)

            Dbg.Assert(Not IsNothing(drc))

            For Each dr As DataRow In drc

                '如果目前步驟有最後的步驟(END)
                If String.Compare(SY_TABLEBASE.CDbType(Of String)(dr("STEP_NO")), SY_STEP_NO.SystemStep_End, True) = 0 Then
                    tempStepInfo = InnerSendFlow(SY_TABLEBASE.CDbType(Of String)(dr("CASEID")), "N",
                                        CInt(dr("SUBFLOW_SEQ")), Nothing, Nothing, "", True)

                    If tempStepInfo.currentStepInfo.subflowSeq = currentStepInfo.currentStepInfo.subflowSeq Then
                        stepinfo = tempStepInfo
                    End If

                    Continue For
                End If

                '如果目前步驟有最初的步驟(START)
                If String.Compare(SY_TABLEBASE.CDbType(Of String)(dr("STEP_NO")), SY_STEP_NO.SystemStep_Start, True) = 0 Then
                    tempStepInfo = InnerSendFlow(SY_TABLEBASE.CDbType(Of String)(dr("CASEID")),
                                            "N",
                                            CInt(dr("SUBFLOW_SEQ")),
                                            Nothing,
                                            assignNextStepUser,
                                            "",
                                            True,
                                            bSendToNext)

                    If tempStepInfo.currentStepInfo.subflowSeq = currentStepInfo.currentStepInfo.subflowSeq Then
                        stepinfo = tempStepInfo
                    End If

                    Continue For
                End If

                Dim sNextStepType As String
                sNextStepType = getSYFlowMap().GetStepType(CInt(dr("FLOW_ID")), SY_TABLEBASE.CDbType(Of String)(dr("STEP_NO")))

                '如果下一個步驟有子流程的分支
                If sNextStepType = "S" Then
                    tempStepInfo = InnerSendFlow(SY_TABLEBASE.CDbType(Of String)(dr("CASEID")), "N",
                                        CInt(dr("SUBFLOW_SEQ")), Nothing, assignNextStepUser, "", True)

                    If tempStepInfo.currentStepInfo.subflowSeq = currentStepInfo.currentStepInfo.subflowSeq Then
                        stepinfo = tempStepInfo
                    End If

                    Continue For
                End If

                '如果下一個步驟有子流程的合併
                If sNextStepType = "J" OrElse sNextStepType = "N" OrElse sNextStepType = "P" Then

                    tempStepInfo = InnerSendFlow(SY_TABLEBASE.CDbType(Of String)(dr("CASEID")), "N",
                                        CInt(dr("SUBFLOW_SEQ")), Nothing, assignNextStepUser, "", True)

                    If tempStepInfo.currentStepInfo.subflowSeq = currentStepInfo.currentStepInfo.subflowSeq OrElse sNextStepType = "P" OrElse sNextStepType = "J" Then
                        stepinfo = tempStepInfo
                    End If

                    Continue For
                End If
            Next

            Return stepinfo

        Catch ex As SYException
            Throw
        Catch ex As Exception
            Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
        End Try

        Return Nothing
    End Function


    ''' <summary>
    ''' 取得下一步驟
    ''' </summary>
    ''' <param name="currStep"></param>
    ''' <param name="sPN"></param>
    ''' <param name="assignNextStepBradepNo">指定要送至下一步驟或分行，若沒指定表示由流程定義自行決定</param>
    ''' <param name="assignNextStepUser">指定下一步要使用的STEPUSER，若沒指定表示由流程定義自行決定</param>
    ''' <param name="bGetStepNoOnly">只取得下一步驟的流程步驟，不取得其它資訊(員編，分行等)</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetNextStep(ByVal currStep As StepInfoItemExt,
                                   ByVal sPN As String,
                                   Optional ByVal assignNextStepBradepNo() As STEPNO_BRADEPNO = Nothing,
                                   Optional ByVal assignNextStepUser() As STEPNO_USERID = Nothing,
                                   Optional ByVal bGetStepNoOnly As Boolean = False) As StepInfoItemExt()

        Dim listStepInfoItemExt1, listStepInfoItemExt2 As New List(Of StepInfoItemExt)
        Dim nextStepInfoItemExt As StepInfoItemExt
        'Dim bStepNoSpecified As Boolean = False
        'Dim sNextUser As String = String.Empty
        Dim bDuplicatedUserRemoved As Boolean = False

        'Dbg.Assert(currStep.stepNo <> "SPLITTER2")

        'If IsNothing(assignNextStepBradepNo) = False Then
        '    Dbg.Assert(assignNextStepBradepNo(0).StepNo <> "04100710")
        'End If

        Dim nConditionCount As Integer = 0

        Try
            '若指定下一步要使用的STEPNO, 設為使用該STEPNO
            '若沒指定則由流程定義取得


            '若不是條件，可送至指定的步驟及分行，若是條件，可送至符合步驟的分行
            nConditionCount = CDbType(Of Integer)(ExecuteScalar(
                "select count(*) from SY_FLOW_DEF where FLOW_ID=@FLOW_ID@ and CURR_STEP_NO=@CURR_STEP_NO@ and PN=@PN@ and COND_ID is not null",
                "FLOW_ID", currStep.flowId,
                "CURR_STEP_NO", currStep.stepNo,
                "PN", sPN), 0)

            If IsNothing(assignNextStepBradepNo) = False AndAlso nConditionCount = 0 Then
                For Each itemNextStep As STEPNO_BRADEPNO In assignNextStepBradepNo
                    If Not IsNothing(itemNextStep) Then
                        nextStepInfoItemExt = currStep.clone

                        nextStepInfoItemExt.stepNo = itemNextStep.StepNo
                        nextStepInfoItemExt.stepType = getSYFlowMap.GetStepType(nextStepInfoItemExt.flowId, nextStepInfoItemExt.stepNo)

                        If itemNextStep.BraDepNo <> 0 Then
                            nextStepInfoItemExt.braDepNo_brid = ToArray(Of BRADEPNO_BRID)(BRADEPNO_BRID.Pair(itemNextStep.BraDepNo, getSYBranch.GetBRID(itemNextStep.BraDepNo)))
                        Else
                            nextStepInfoItemExt.braDepNo_brid = getSYFlowDef.GetNextStepSpanBraDepNo(
                                currStep.flowId, currStep.stepNo, sPN, nextStepInfoItemExt.stepNo)

                            If IsNothing(nextStepInfoItemExt.braDepNo_brid) Then
                                nextStepInfoItemExt.braDepNo_brid = currStep.braDepNo_brid
                            End If
                        End If

                        '只指定給特定分行及角色
                        nextStepInfoItemExt.braDepNo_brid = getSYNextFlowStepRule.FilterBrid(
                                nextStepInfoItemExt.caseId,
                                nextStepInfoItemExt.stepNo,
                                nextStepInfoItemExt.braDepNo_brid)

                        '設定指派給誰
                        nextStepInfoItemExt.caseOwner = Nothing
                        nextStepInfoItemExt.caseClient = Nothing
                        nextStepInfoItemExt.caseSender = Nothing
                        nextStepInfoItemExt.caseOwnerList = New List(Of USER_ID_NAME)

                        If Not IsNothing(assignNextStepUser) Then
                            For Each stepitem In assignNextStepUser
                                If stepitem.sStepNo = nextStepInfoItemExt.stepNo OrElse String.IsNullOrEmpty(stepitem.sStepNo) Then
                                    If getSYUser.IsValidUser(stepitem.sUserId) Then
                                        nextStepInfoItemExt.caseOwner = getSYUser.GetUserIdName(stepitem.sUserId)
                                    End If
                                    Exit For
                                End If
                            Next
                        End If

                        '只指定給特定分行及角色
                        nextStepInfoItemExt.caseOwner = getSYNextFlowStepRule.FilterUser(
                                nextStepInfoItemExt.caseId,
                                nextStepInfoItemExt.stepNo,
                                nextStepInfoItemExt.caseOwner)

                        listStepInfoItemExt1.Add(nextStepInfoItemExt)
                    End If
                Next
            Else
                '先取得目前步驟的所有的下一步驟
                Dim drcFlowDef As DataRowCollection
                drcFlowDef = getSYFlowDef().GetNextStep(currStep.flowId, currStep.stepNo, sPN)

                If IsNothing(drcFlowDef) OrElse drcFlowDef.Count = 0 Then
                    Return Nothing
                    'Throw New Exception(String.Format("無法取得下一步驟(流程:{0}，步驟:{1}，PN:{2})",
                    '                                  currStep.flowId, currStep.stepNo, sPN))
                End If

                Dim bCondition As Boolean = False

                For Each drFlowDef As DataRow In drcFlowDef
                    nextStepInfoItemExt = currStep.clone

                    'Dim nSpanBraDepNo As Integer
                    'nSpanBraDepNo = SY_TABLEBASE.CDbType(Of Integer)(drFlowDef("SPANBRID"), 0)

                    'If nSpanBraDepNo <> 0 Then
                    '    'nextStepInfoItemExt.braDepNo = nSpanBraDepNo
                    '    'nextStepInfoItemExt.brid = SY_TABLEBASE.CDbType(Of String)(drFlowDef("BRID"), Nothing)
                    '    nextStepInfoItemExt.braDepNo_brid = BRADEPNO_BRID.ToArray(
                    '        BRADEPNO_BRID.Pair(nSpanBraDepNo, SY_TABLEBASE.CDbType(Of String)(drFlowDef("BRID"), Nothing)))
                    'End If

                    nextStepInfoItemExt.flowId = CInt(drFlowDef("FLOW_ID"))
                    nextStepInfoItemExt.flowName = CStr(drFlowDef("FLOW_NAME"))
                    nextStepInfoItemExt.stepNo = CStr(drFlowDef("NEXT_STEP_NO"))
                    nextStepInfoItemExt.stepType = CDbType(Of String)(drFlowDef("NEXTTYPE"), "")

                    '跨單位
                    Dim spanBranch() As BRADEPNO_BRID
                    spanBranch = getSYRelFlowMap_SpanBranch.GetSpanBranch(nextStepInfoItemExt.flowId,
                                                                              nextStepInfoItemExt.stepNo)
                    If Not IsNothing(spanBranch) Then
                        nextStepInfoItemExt.braDepNo_brid = spanBranch
                    Else
                        If IsNothing(assignNextStepBradepNo) = False Then
                            For Each currStepnoBradepno As STEPNO_BRADEPNO In assignNextStepBradepNo
                                If currStepnoBradepno.StepNo = nextStepInfoItemExt.stepNo AndAlso
                                    currStepnoBradepno.BraDepNo <> 0 Then
                                    nextStepInfoItemExt.braDepNo_brid = BRADEPNO_BRID.arrayBradepnoBrid(currStepnoBradepno.BraDepNo, getSYBranch.GetBRID(currStepnoBradepno.BraDepNo))
                                End If
                            Next
                        End If
                    End If

                    '只指定給特定分行及角色
                    nextStepInfoItemExt.braDepNo_brid = getSYNextFlowStepRule.FilterBrid(
                            nextStepInfoItemExt.caseId,
                            nextStepInfoItemExt.stepNo,
                            nextStepInfoItemExt.braDepNo_brid)

                    nextStepInfoItemExt.caseOwner = Nothing
                    nextStepInfoItemExt.caseClient = Nothing
                    nextStepInfoItemExt.caseSender = Nothing

                    If Not IsNothing(assignNextStepUser) Then
                        For Each stepitem In assignNextStepUser
                            If stepitem.sStepNo = nextStepInfoItemExt.stepNo OrElse
                                String.IsNullOrEmpty(stepitem.sStepNo) Then

                                If getSYUser.IsValidUser(stepitem.sUserId) Then
                                    nextStepInfoItemExt.caseOwner = getSYUser.GetUserIdName(stepitem.sUserId)

                                    '只指定給特定分行及角色
                                    nextStepInfoItemExt.caseOwner = getSYNextFlowStepRule.FilterUser(
                                            nextStepInfoItemExt.caseId,
                                            nextStepInfoItemExt.stepNo,
                                            nextStepInfoItemExt.caseOwner)

                                    nextStepInfoItemExt.caseClient = nextStepInfoItemExt.caseOwner
                                    Exit For
                                End If
                            End If
                        Next
                    End If

                    '如果有條件，檢查條件是否符合
                    If Not IsDBNull(drFlowDef("COND_ID")) Then
                        Dim nCondId As Integer = CInt(drFlowDef("COND_ID"))
                        bCondition = True

                        If getSYConditionId().LogicalOperation(currStep.caseId, nCondId) = False Then
                            Continue For
                        End If
                    End If

                    listStepInfoItemExt1.Add(nextStepInfoItemExt)

                    '如果是個條件且不是Splitter，取第一個符合的步驟
                    If bCondition = True AndAlso currStep.stepType <> "S" Then
                        Exit For
                    End If
                Next
            End If

            If bGetStepNoOnly = True Then
                Return listStepInfoItemExt1.ToArray
            End If


            For Each nextStepInfoItemExt In listStepInfoItemExt1

                'Dbg.Assert(nextStepInfoItemExt.stepNo <> "04100721" OrElse nextStepInfoItemExt.stepNo <> "04100751")

                '如果下一個步驟不是END，找出下一步驟的OWNER
                If String.Compare(nextStepInfoItemExt.stepNo, SY_STEP_NO.SystemStep_End, True) = 0 Then
                    GoTo SKIP_OWNER
                End If


                '如果預設的USER是空的且此步驟已經被重複執行(修正補充)，找出原本的使用者
                If IsNothing(nextStepInfoItemExt.caseOwner) Then

                    Dim previousInfo As DataRowCollection
                    Dim sStepType As String
                    previousInfo = getSYFlowStep.GetLatestInfoByCaseidSubflowseqStepno(nextStepInfoItemExt.caseId, 0, nextStepInfoItemExt.stepNo)
                    sStepType = getSYFlowMap.GetStepType(nextStepInfoItemExt.flowId, nextStepInfoItemExt.stepNo)
                    If previousInfo.Count > 0 AndAlso String.IsNullOrEmpty(sStepType) = True Then
                        Dim previousUser As String = Nothing
                        previousUser = CDbType(Of String)(previousInfo(0)("CLIENT"), "")

                        Dim previousBraDepNo As Integer = 0
                        previousBraDepNo = CDbType(Of Integer)(previousInfo(0)("BRA_DEPNO"), 0)

                        If previousBraDepNo <> 0 Then
                            nextStepInfoItemExt.braDepNo_brid =
                                SY_TABLEBASE.ToArray(BRADEPNO_BRID.Pair(previousBraDepNo, getSYBranch.GetBRID(previousBraDepNo)))
                        End If

                        If getSYRelBranchUser.GetCount(PARAM_ARRAY("BRA_DEPNO", previousBraDepNo, "STAFFID", previousUser)) > 0 Then
                            nextStepInfoItemExt.caseOwner = getSYUser.GetUserIdName(previousUser)
                        Else
                            nextStepInfoItemExt.caseOwner = Nothing
                        End If

                    End If

                    'Dim userId As String =
                    'getSYFlowStep.GetLatestClientByCaseidSubflowseqStepno(nextStepInfoItemExt.caseId, nextStepInfoItemExt.stepNo)

                    'getSYFlowStep.GetLatestClientByCaseidSubflowseqStepno(
                    '    nextStepInfoItemExt.caseId, nextStepInfoItemExt.subflowSeq, nextStepInfoItemExt.stepNo)


                    '只指定給特定分行及角色
                    nextStepInfoItemExt.caseOwner = getSYNextFlowStepRule.FilterUser(
                            nextStepInfoItemExt.caseId,
                            nextStepInfoItemExt.stepNo,
                            nextStepInfoItemExt.caseOwner)

                    nextStepInfoItemExt.caseClient = nextStepInfoItemExt.caseOwner
                End If

                '如果使用者離職，將下一步驟的使用者設為空的
                '如果使用者的角色被移除，將下一步驟的使用者設為空的
                If IsNothing(nextStepInfoItemExt.caseOwner) OrElse
                    getSYUser().IsValidUser(nextStepInfoItemExt.caseOwner.userId) = False OrElse
                    getSYRelRoleFlowMap.IsValidUser(nextStepInfoItemExt.caseOwner.userId, nextStepInfoItemExt.flowId, nextStepInfoItemExt.stepNo) = False Then
                    nextStepInfoItemExt.caseOwner = Nothing
                    nextStepInfoItemExt.caseClient = Nothing
                End If

                '如果是虛擬步驟
                'If IsNothing(nextStepInfoItemExt.caseOwner) AndAlso
                '    getSYFlowMap().GetStepType(currStep.flowId, currStep.stepNo) = "N" Then
                '    nextStepInfoItemExt.caseOwner = currStep.caseSender
                '    nextStepInfoItemExt.caseClient = currStep.caseSender
                'End If

                '檢查上下步驟是否可為同一人,
                If (Not IsNothing(nextStepInfoItemExt.caseOwner)) AndAlso
                    getSYFlowDef.isProhibiteSameUser(
                        currStep.flowId, currStep.stepNo, nextStepInfoItemExt.stepNo) = True AndAlso
                    String.Compare(currStep.caseSender.userId, nextStepInfoItemExt.caseOwner.userId) = 0 Then

                    '若同一人則設為空的
                    nextStepInfoItemExt.caseOwner = Nothing
                    nextStepInfoItemExt.caseClient = nextStepInfoItemExt.caseOwner
                    bDuplicatedUserRemoved = True
                End If

                '如果預設的使用者是空的，則查詢符合該步驟角色的使用者
                nextStepInfoItemExt.caseOwnerList = New List(Of USER_ID_NAME)

                If IsNothing(nextStepInfoItemExt.caseOwner) Then
                    Dim drc As DataRowCollection
                    'drc = getSYRelRoleUser().GetUserlistByFlowidStepnoBradepno(nextStepInfoItemExt.flowId, nextStepInfoItemExt.stepNo, nextStepInfoItemExt.braDepNo)

                    'Dbg.Assert(currStep.stepNo <> "SPLITTER2" OrElse nextStepInfoItemExt.stepNo <> "04101100")

                    Dim nBraDepNo As Integer = 0
                    If Not IsNothing(nextStepInfoItemExt.braDepNo_brid) Then
                        nBraDepNo = nextStepInfoItemExt.braDepNo_brid(0).BraDepNo
                    End If

                    If nBraDepNo <> 0 Then
                        Dim bradepnoBrid As BRADEPNO_BRID
                        bradepnoBrid = BRADEPNO_BRID.Pair(nBraDepNo, String.Empty)

                        bradepnoBrid = getSYNextFlowStepRule.FilterBrid(
                                nextStepInfoItemExt.caseId,
                                nextStepInfoItemExt.stepNo,
                                bradepnoBrid)

                        If IsNothing(bradepnoBrid) Then
                            Continue For
                        End If
                    End If


Find_Next_User:
                    '找出下一步驟的使用者
                    drc = getSYFlowMap.GetNextStepUser(
                        currStep.flowId,
                        currStep.stepNo,
                        nextStepInfoItemExt.stepNo,
                        nBraDepNo,
                        currStep.caseSender.userId)

                    'Dbg.Assert(nextStepInfoItemExt.stepNo <> "04100710")

                    If IsNothing(drc) Then

                        'Dim sStepType As String = getSYFlowMap().GetStepType(nextStepInfoItemExt.flowId, nextStepInfoItemExt.stepNo)
                        'nextStepInfoItemExt.stepType = sStepType

                        If nextStepInfoItemExt.stepType = "S" OrElse nextStepInfoItemExt.stepType = "J" OrElse nextStepInfoItemExt.stepType = "N" OrElse nextStepInfoItemExt.stepType = "P" Then   '不需要理會使用者是誰
                            GoTo SKIP_OWNER
                        End If

                        '如果找不到，但符合的只有一間分行，改nBraDepNo再找(可能是分行沒設，只要只有一家分行，可自行重新指派分行，兩家或以上不可自行重新指派分行)
                        Dim drcNextBraDepNO As DataRowCollection
                        drcNextBraDepNO = getSYFlowMap.GetNextStepUser(
                            currStep.flowId,
                            currStep.stepNo,
                            nextStepInfoItemExt.stepNo,
                            0,
                            currStep.caseSender.userId,
                            True)

                        '如果找不到，但符合的只有一間分行，改nBraDepNo再找(可能是分行沒設，只要只有一家分行，可自行重新指派分行，兩家或以上不可自行重新指派分行)
                        If IsNothing(drcNextBraDepNO) = False AndAlso drcNextBraDepNO.Count = 1 Then
                            If CDbType(Of Integer)(drcNextBraDepNO(0)("BRA_DEPNO")) <> nBraDepNo Then
                                nBraDepNo = CDbType(Of Integer)(drcNextBraDepNO(0)("BRA_DEPNO"))
                                nextStepInfoItemExt.braDepNo_brid = ToArray(Of BRADEPNO_BRID)(BRADEPNO_BRID.Pair(nBraDepNo, getSYBranch.GetBRID(nBraDepNo)))
                                GoTo Find_Next_User
                            End If
                        End If


                        If bDuplicatedUserRemoved Then
                            Dim sException As String
                            sException = String.Format("無法取得下一步驟的使用者(流程:{0}，步驟:{1}，單位代碼:{2})。",
                                                               nextStepInfoItemExt.flowId,
                                                               nextStepInfoItemExt.stepNo,
                                                               nextStepInfoItemExt.braDepNo_brid(0).BraDepNo) & vbCrLf & _
                                         "因為下一步驟的唯一的使用者跟此步驟的使用者相同，所以無法案件無法送出。"
                            Throw New SYException(sException,
                                                  SYMSG.FLOW_CLIENT_NOT_FOUND_IN_NEXTSTEP_SAMEPROHIBITED,
                                                  GetGlobalLastSQL)
                        Else
                            Throw New SYException(
                                String.Format("無法取得下一步驟的使用者(流程:{0}，步驟:{1}，單位代碼:{2})", nextStepInfoItemExt.flowId, nextStepInfoItemExt.stepNo, nextStepInfoItemExt.braDepNo_brid(0).BraDepNo),
                                SYMSG.FLOW_CLIENT_NOT_FOUND_IN_NEXTSTEP,
                                GetGlobalLastSQL)
                        End If

                    End If

                    nextStepInfoItemExt.caseOwnerList = New List(Of USER_ID_NAME)

                    For Each row As DataRow In drc
                        nextStepInfoItemExt.caseOwnerList.Add(
                            getSYUser.GetUserIdName(CStr(row("STAFFID"))))
                    Next

                    nextStepInfoItemExt.caseOwner = getSYNextFlowStepRule.FilterUser(
                            nextStepInfoItemExt.caseId,
                            nextStepInfoItemExt.stepNo,
                            nextStepInfoItemExt.caseOwner)

                    If nextStepInfoItemExt.caseOwnerList.Count = 1 Then
                        nextStepInfoItemExt.caseOwner = nextStepInfoItemExt.caseOwnerList(0)
                        nextStepInfoItemExt.caseClient = nextStepInfoItemExt.caseOwner
                        nextStepInfoItemExt.caseOwnerList = Nothing
                    End If

                Else
                    'nextStepInfoItemExt.caseOwnerList.Add(nextStepInfoItemExt.caseOwner)
                    nextStepInfoItemExt.caseOwnerList = Nothing
                    nextStepInfoItemExt.caseClient = nextStepInfoItemExt.caseOwner
                    'nextStepInfoItemExt.caseOwner = String.Empty
                End If
SKIP_OWNER:
                nextStepInfoItemExt.caseClient = nextStepInfoItemExt.caseOwner

                'Dbg.Assert(currStep.stepNo <> "04100300")


                '如果案件所屬的分行不對，重設分行
                If String.IsNullOrEmpty(nextStepInfoItemExt.stepType) Then
                    Dim drcBradepnoUser As DataRowCollection = Nothing
                    Dim newBraDepNo As Integer

                    If IsNothing(nextStepInfoItemExt.caseOwner) = False Then
                        drcBradepnoUser = getSYRelRoleFlowMap.GetUserBradepno(
                            nextStepInfoItemExt.flowId, nextStepInfoItemExt.stepNo,
                            nextStepInfoItemExt.caseOwner.userId, nextStepInfoItemExt.braDepNo_brid(0).BraDepNo)

                        '目前案件所屬的分行不對，重設分行
                        Debug.Assert(IsNothing(drcBradepnoUser) = False AndAlso drcBradepnoUser.Count > 0,
                                     "案件改分派至其它分行時需指定分行代碼")

                        If IsNothing(drcBradepnoUser) = True OrElse drcBradepnoUser.Count = 0 Then
                            drcBradepnoUser = getSYRelRoleFlowMap.GetUserBradepno(
                                nextStepInfoItemExt.flowId, nextStepInfoItemExt.stepNo,
                                nextStepInfoItemExt.caseOwner.userId)

                            Dbg.Assert(IsNothing(drcBradepnoUser) = False)
                            '若找出新的分行，重設，若沒有，不重設
                            If IsNothing(drcBradepnoUser) = False AndAlso drcBradepnoUser.Count > 0 Then
                                newBraDepNo = CInt(drcBradepnoUser(0)("BRA_DEPNO"))
                                nextStepInfoItemExt.braDepNo_brid = ToArray(Of BRADEPNO_BRID)(BRADEPNO_BRID.Pair(newBraDepNo, getSYBranch.GetBRID(newBraDepNo)))
                            Else
                                Throw New SYException(String.Format("使用者沒有權限執行下一個流程步驟(案件的下一步驟資訊:{0}, 流程:{1}, 步驟:{2}, 子流程代碼:{3}, 使用者:{4})",
                                                              nextStepInfoItemExt.caseId,
                                                              nextStepInfoItemExt.flowId,
                                                              nextStepInfoItemExt.stepNo,
                                                              nextStepInfoItemExt.subflowSeq,
                                                              nextStepInfoItemExt.caseOwner.userId),
                                                          SYMSG.FLOW_NO_RIGHTS_TO_PERFORM_NEXTSTEP,
                                                          GetGlobalLastSQL)
                            End If
                        End If
                    Else
                        If IsNothing(nextStepInfoItemExt.caseOwnerList) = False Then
                            For Each caseUser As USER_ID_NAME In nextStepInfoItemExt.caseOwnerList
                                drcBradepnoUser = getSYRelRoleFlowMap.GetUserBradepno(
                                    nextStepInfoItemExt.flowId, nextStepInfoItemExt.stepNo,
                                    caseUser.userId, nextStepInfoItemExt.braDepNo_brid(0).BraDepNo)

                                '目前案件所屬的分行不對，重設分行
                                Debug.Assert(IsNothing(drcBradepnoUser) = False AndAlso drcBradepnoUser.Count > 0,
                                             "案件改分派至其它分行時需指定分行代碼")
                                If IsNothing(drcBradepnoUser) = True OrElse drcBradepnoUser.Count = 0 Then
                                    drcBradepnoUser = getSYRelRoleFlowMap.GetUserBradepno(
                                        nextStepInfoItemExt.flowId, nextStepInfoItemExt.stepNo,
                                        caseUser.userId)

                                    Dbg.Assert(IsNothing(drcBradepnoUser) = False)
                                    '若找出新的分行，重設，若沒有，不重設
                                    If IsNothing(drcBradepnoUser) = False AndAlso drcBradepnoUser.Count > 0 Then
                                        newBraDepNo = CInt(drcBradepnoUser(0)("BRA_DEPNO"))
                                        nextStepInfoItemExt.braDepNo_brid = ToArray(Of BRADEPNO_BRID)(BRADEPNO_BRID.Pair(newBraDepNo, getSYBranch.GetBRID(newBraDepNo)))
                                    Else
                                        Throw New SYException(String.Format("使用者沒有權限執行下一個流程步驟(案件:{0}, 流程:{1}, 步驟:{2}, 子流程代碼:{3}, 使用者:{4})",
                                                                      nextStepInfoItemExt.caseId,
                                                                      nextStepInfoItemExt.flowId,
                                                                      nextStepInfoItemExt.stepNo,
                                                                      nextStepInfoItemExt.subflowSeq,
                                                                      caseUser.userId),
                                                                  SYMSG.FLOW_NO_RIGHTS_TO_PERFORM_NEXTSTEP,
                                                                  GetGlobalLastSQL)
                                    End If
                                End If
                            Next
                        End If
                    End If
                End If

                listStepInfoItemExt2.Add(nextStepInfoItemExt)
            Next

            '如果不是分支，只取第一個成立的條件
            If currStep.stepType <> "S" AndAlso listStepInfoItemExt2.Count > 1 Then
                listStepInfoItemExt2.RemoveRange(1, listStepInfoItemExt2.Count - 1)
            End If

            Return listStepInfoItemExt2.ToArray()

        Catch ex As SYException
            Throw
        Catch ex As Exception
            Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
        End Try

    End Function


    Protected Function isNextStep_a_Splitter(ByVal currStep As StepInfoItemExt,
                               ByVal sPN As String) As Boolean
        Dim sii() As StepInfoItemExt
        'Dim listNextStepUser As New List(Of STEPNO_USERID)
        'listNextStepUser.Add(currentUser)

        Try
            sii = GetNextStep(currStep, sPN, Nothing, Nothing, True)

            If IsNothing(sii) OrElse sii.Count <> 1 Then
                Return False
            End If

            If sii(0).stepType = "S" Then
                Return True
            End If


        Catch ex As Exception
            '不回傳Exception
        End Try

        Return False
    End Function


    ''' <summary>
    ''' 取得目前案件未完成的步驟資訊，若流程已完成，取得最後一步驟的流程步驟
    ''' </summary>
    ''' <param name="sCaseId"></param>
    ''' <param name="nSubFlowSeq"></param>
    ''' <returns></returns>
    ''' <remarks>t</remarks>
    Public Function GetCurrentStepInfoOfCaseid(ByVal sCaseId As String,
                                  Optional ByVal nSubFlowSeq As Integer = 0) As StepInfoItemExt

        Dim arraySII() As StepInfoItemExt

        Try
            arraySII = getSYFlowStep.GetFlowStepInfo(sCaseId,
                                               nSubFlowSeq,
                                               Nothing,
                                               1,
                                               "TOP 1",
                                               "STARTTIME DESC")

            If IsNothing(arraySII) OrElse arraySII.Count = 0 Then
                arraySII = getSYFlowStep.GetFlowStepInfo(sCaseId,
                                                   nSubFlowSeq,
                                                   Nothing,
                                                   3,
                                                   "TOP 1",
                                                   "PROCESSTIME DESC")
            End If

            If IsNothing(arraySII) OrElse arraySII.Count = 0 Then
                Return Nothing
            End If

            If IsNothing(arraySII(0).caseOwner) = False AndAlso
                String.IsNullOrEmpty(arraySII(0).caseOwner.userId) = False Then
                Return arraySII(0)
            End If

            '若案件沒有設為OWNER，找出所有的OWNER

            Dim drFlowStep As DataRow = Nothing
            '取得上一步驟
            drFlowStep = getSYFlowStep.GetFormerFlowStepInfo(arraySII(0).caseId, arraySII(0).subflowSeq)

            '取得目前步驟的使用者
            Dim drcOwner As DataRowCollection = Nothing
            drcOwner = getSYFlowMap.GetNextStepUser(arraySII(0).flowId,
                                                                CStr(drFlowStep("STEP_NO")),
                                                                arraySII(0).stepNo,
                                                                CInt(drFlowStep("BRA_DEPNO")),
                                                                CStr(drFlowStep("CLIENT")))

            If IsNothing(drcOwner) OrElse drcOwner.Count = 0 Then
                Return arraySII(0)
            End If

            If drcOwner.Count = 1 Then
                arraySII(0).caseOwner = USER_ID_NAME.USER(CStr(drcOwner(0)("STAFFID")), CStr(drcOwner(0)("USERNAME")))
            Else
                Dim listUserIdName As New List(Of USER_ID_NAME)

                For Each drOwner As DataRow In drcOwner
                    listUserIdName.Add(USER_ID_NAME.USER(CStr(drOwner("STAFFID")), CStr(drOwner("USERNAME"))))
                Next

                arraySII(0).caseOwnerList = listUserIdName
            End If

            Return arraySII(0)

        Catch ex As SYException
            Throw
        Catch ex As Exception
            Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
        End Try

    End Function


    ''' <summary>
    ''' 取得STEPINFO內WORKINGUSER及LOGINUSER
    ''' </summary>
    ''' <param name="loginUserid"></param>
    ''' <param name="workingUserid"></param>
    ''' <remarks>如果已設定WORKINGUSER及LOGINUSER，則會檢查是否真得存在SY_USERINFO內</remarks>
    Public Sub GetCurrentUserid(ByRef loginUserid As USER_ID_NAME, ByRef workingUserid As USER_ID_NAME)

        Dim strLoginUser As String = Nothing
        Dim strWorkingUser As String = Nothing

        Try
            Dbg.Assert(IsNothing(m_callbackUserInfo) = False, "程式需要設用何種方式取得目前使用者")

            If IsNothing(m_callbackUserInfo) Then
                Throw New SYException("使用者登入資訊遺失，可能是系統更新元件時重置不同元件內的物件參考，請重新登入或回到待處理工作重新點選案件", SYMSG.FLOW_USERINFO_NOT_FOUND)
            End If

            m_callbackUserInfo.GetCurrentUserid(strLoginUser, strWorkingUser)
            '若測試程式發生錯誤，請加入下列的程式，表示WEB程式的USER可從SESSION取得
            'FLOW_OP.ELoanFlow.m_callbackUserInfo = New com.Azion.UITools.ExportUserInfo

            '設定目前的工作使用者
            If IsNothing(workingUserid) OrElse
                String.IsNullOrEmpty(workingUserid.userId) OrElse
                String.IsNullOrEmpty(workingUserid.userName) Then
                workingUserid = getSYUser.GetUserIdName(strWorkingUser)

                Dbg.Assert(workingUserid.userId.Length > 0 AndAlso workingUserid.userName.Length > 0)
            End If

            '設定目前的登入使用者
            If IsNothing(loginUserid) OrElse
                String.IsNullOrEmpty(loginUserid.userId) OrElse
                String.IsNullOrEmpty(loginUserid.userName) Then
                loginUserid = getSYUser.GetUserIdName(strLoginUser)

                Dbg.Assert(loginUserid.userId.Length > 0 AndAlso loginUserid.userName.Length > 0)
            End If

            '檢查使用者是否有效
            IsUserValid(workingUserid.userId)
            If String.Compare(workingUserid.userId, loginUserid.userId, True) <> 0 Then     '若同一使用者則不重覆檢查
                IsUserValid(loginUserid.userId)
            End If

        Catch ex As SYException
            Throw
        Catch ex As Exception
            Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
        End Try

    End Sub


    ''' <summary>
    ''' 檢查是否是有效的使用者，(SY_USER.STATUS=0)
    ''' </summary>
    ''' <param name="sUserId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function IsUserValid(ByVal sUserId As String) As Boolean

        Dim dataRow As DataRow
        dataRow = getSYUser().GetUserInfo(sUserId)

        '找不到使用者
        If IsNothing(dataRow) Then
            Throw New SYException(String.Format("使用者({0})不存在", sUserId), SYMSG.FLOW_INVALID_USER, GetGlobalLastSQL)
        End If

        '使用者不在職
        If CType(dataRow("STATUS"), Integer) <> 0 Then
            Throw New SYException(String.Format("使用者({0})已離職", sUserId), SYMSG.FLOW_USER_LEFT, GetGlobalLastSQL)
        End If

        Return True
    End Function



    Public Sub EMailNotification(ByVal sii() As StepInfoItemExt)
        Try
            If IsNothing(m_callbackFlowMail) = False Then
                m_callbackFlowMail.Notify(getDatabaseManager(), sii)
            End If
        Catch ex As Exception
        End Try
    End Sub


End Class
