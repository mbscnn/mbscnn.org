Option Explicit On
Option Strict On

Imports com.Azion.NET.VB
'Imports com.Azion.EloanUtility
Imports FLOW_OP.TABLE


Public Class STEPNO_USERID
    Public sStepNo As String
    Public sUserId As String

    Public Shared Function arrayStepUser(ByVal sStepNo As String, ByVal sUserId As String) As STEPNO_USERID()

        If String.IsNullOrEmpty(sStepNo) AndAlso String.IsNullOrEmpty(sUserId) Then
            Return Nothing
        End If

        Dim stepUser(0) As STEPNO_USERID
        stepUser(0) = New STEPNO_USERID
        stepUser(0).sStepNo = sStepNo
        stepUser(0).sUserId = sUserId
        Return stepUser
    End Function

    Public Shared Function StepUser(ByVal sStepNo As String, ByVal sUserId As String) As STEPNO_USERID

        If String.IsNullOrEmpty(sStepNo) AndAlso String.IsNullOrEmpty(sUserId) Then
            Return Nothing
        End If

        Dim objStepUser As New STEPNO_USERID
        objStepUser.sStepNo = sStepNo
        objStepUser.sUserId = sUserId
        Return objStepUser
    End Function
End Class


''' <summary>
''' a facade pattern of FLOW_OP
''' </summary>
''' <remarks></remarks>
Public Class FlowFacade
    Inherits SY_TABLEBASE

    Protected m_EloanFlow As ELoanFlow
    'Protected m_SyTables As SY_TABLEBASE


    ''' <summary>
    ''' Constructor
    ''' </summary>
    ''' <param name="dbManager"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal dbManager As com.Azion.NET.VB.DatabaseManager)
        MyBase.new(dbManager)
        'm_dbManager = dbManager
    End Sub

    Public Shared Function getNewInstance(ByVal dbManager As com.Azion.NET.VB.DatabaseManager) As FlowFacade
        Return New FlowFacade(dbManager)
    End Function

    ReadOnly Property getELoanFlow() As ELoanFlow
        Get
            If IsNothing(m_EloanFlow) Then m_EloanFlow = New ELoanFlow(getDatabaseManager())
            Return m_EloanFlow
        End Get
    End Property

    'ReadOnly Property getSYTables() As SY_TABLEBASE
    '    Get
    '        If IsNothing(m_SyTables) Then m_SyTables = New SY_TABLEBASE(m_dbManager)
    '        Return m_SyTables
    '    End Get
    'End Property


    ReadOnly Property getBraDepNo(ByVal sBrid As String) As Integer
        Get
            If String.IsNullOrEmpty(sBrid) Then
                Return 0
            End If

            Return getSYBranch().GetBraDepNo(sBrid)
        End Get
    End Property


    ''' <summary>
    ''' 直接起案並回傳new caseid
    ''' </summary>
    ''' <param name="nBraDepNo">分行別</param>
    ''' <param name="sFlowName">Flow Name</param>
    ''' <param name="bAutoSendToNext">起案後，自動送至下一步驟</param>
    ''' <param name="sNextUserid">下一步驟的員工編號。預設值為NULL，由FLOW程式自行判斷</param> 
    ''' <param name="assignNextStepUser">下一步驟的員工編號。預設值為NULL，由FLOW程式自行判斷</param>
    ''' <param name="bSendToEnd">直接結案</param>
    ''' <returns></returns>
    ''' <remarks>sWorkingUserId及sLoginUserId若沒有特別的理由，應設為""由FLOW程式自行判斷</remarks>
    'Public Function StartFlow(ByVal sFlowName As String,
    '                  ByVal bAutoSendToNext As Boolean,
    '                  Optional ByVal bSendToEnd As Boolean = False) As StepInfo
    '    Return getELoanFlow().StartFlow(0, sFlowName, bAutoSendToNext, Nothing, bSendToEnd)
    'End Function

    ''' <summary>
    ''' 直接起案並回傳new caseid
    ''' </summary>
    ''' <param name="nBraDepNo">分行別</param>
    ''' <param name="sFlowName">Flow Name</param>
    ''' <param name="bAutoSendToNext">起案後，自動送至下一步驟</param>
    ''' <param name="sNextUserid">下一步驟的員工編號。預設值為NULL，由FLOW程式自行判斷</param> 
    ''' <param name="assignNextStepUser">下一步驟的員工編號。預設值為NULL，由FLOW程式自行判斷</param>
    ''' <param name="bSendToEnd">直接結案</param>
    ''' <returns></returns>
    ''' <remarks>sWorkingUserId及sLoginUserId若沒有特別的理由，應設為""由FLOW程式自行判斷</remarks>
    'Public Function StartFlow(ByVal sFlowName As String,
    '                      ByVal bAutoSendToNext As Boolean,
    '                      ByVal assignNextStepUser() As STEPNO_USERID,
    '                      Optional ByVal bSendToEnd As Boolean = False) As StepInfo

    '    Return getELoanFlow().StartFlow(0, sFlowName, bAutoSendToNext, assignNextStepUser, bSendToEnd)
    'End Function


    ''' <summary>
    ''' 直接起案並回傳new caseid
    ''' </summary>
    ''' <param name="nBraDepNo">分行別</param>
    ''' <param name="sFlowName">Flow Name</param>
    ''' <param name="bAutoSendToNext">起案後，自動送至下一步驟</param>
    ''' <param name="sNextUserid">下一步驟的員工編號。預設值為NULL，由FLOW程式自行判斷</param> 
    ''' <param name="assignNextStepUser">下一步驟的員工編號。預設值為NULL，由FLOW程式自行判斷</param>
    ''' <param name="bSendToEnd">直接結案</param>
    ''' <returns></returns>
    ''' <remarks>sWorkingUserId及sLoginUserId若沒有特別的理由，應設為""由FLOW程式自行判斷</remarks>
    'Public Function StartFlow(ByVal sFlowName As String,
    '                          ByVal bAutoSendToNext As Boolean,
    '                          ByVal sNextUserid As String,
    '                          Optional ByVal bSendToEnd As Boolean = False) As StepInfo

    '    Dim stepUser() As STEPNO_USERID = Nothing

    '    If (Not String.IsNullOrEmpty(sNextUserid)) Then
    '        stepUser = STEPNO_USERID.arrayStepUser("", sNextUserid)
    '    End If

    '    Return getELoanFlow().StartFlow(0, sFlowName, bAutoSendToNext, stepUser, bSendToEnd)
    'End Function


    ''' <summary>
    ''' 直接起案並回傳new caseid
    ''' </summary>
    ''' <param name="nBraDepNo">分行別</param>
    ''' <param name="sFlowName">Flow Name</param>
    ''' <param name="bAutoSendToNext">起案後，自動送至下一步驟</param>
    ''' <param name="sNextUserid">下一步驟的員工編號。預設值為NULL，由FLOW程式自行判斷</param> 
    ''' <param name="assignNextStepUser">下一步驟的員工編號。預設值為NULL，由FLOW程式自行判斷</param>
    ''' <param name="bSendToEnd">直接結案</param>
    ''' <returns></returns>
    ''' <remarks>sWorkingUserId及sLoginUserId若沒有特別的理由，應設為""由FLOW程式自行判斷</remarks>
    'Public Function StartFlow(ByVal nBraDepNo As Integer,
    '                          ByVal sFlowName As String,
    '                          ByVal bAutoSendToNext As Boolean,
    '                          ByVal assignNextStepUser() As STEPNO_USERID,
    '                          Optional ByVal bSendToEnd As Boolean = False) As StepInfo

    '    Return getELoanFlow().StartFlow(nBraDepNo, sFlowName, bAutoSendToNext, assignNextStepUser, bSendToEnd)
    'End Function

    ''' <summary>
    ''' 直接起案並回傳new caseid
    ''' </summary>
    ''' <param name="nBraDepNo">分行別</param>
    ''' <param name="sFlowName">Flow Name</param>
    ''' <param name="bAutoSendToNext">起案後，自動送至下一步驟</param>
    ''' <param name="sNextUserid">下一步驟的員工編號。預設值為NULL，由FLOW程式自行判斷</param> 
    ''' <param name="assignNextStepUser">下一步驟的員工編號。預設值為NULL，由FLOW程式自行判斷</param>
    ''' <param name="bSendToEnd">直接結案</param>
    ''' <returns></returns>
    ''' <remarks>sWorkingUserId及sLoginUserId若沒有特別的理由，應設為""由FLOW程式自行判斷</remarks>
    'Public Function StartFlow(ByVal nBraDepNo As Integer,
    '                          ByVal sFlowName As String,
    '                          Optional ByVal bAutoSendToNext As Boolean = False) As StepInfo
    '    Dim sUser As String = Nothing
    '    Return StartFlow(nBraDepNo, sFlowName, bAutoSendToNext, sUser)
    'End Function

    ''' <summary>
    ''' 直接起案並回傳new caseid
    ''' </summary>
    ''' <param name="nBraDepNo">分行別</param>
    ''' <param name="sFlowName">Flow Name</param>
    ''' <param name="bAutoSendToNext">起案後，自動送至下一步驟</param>
    ''' <param name="sNextUserid">下一步驟的員工編號。預設值為NULL，由FLOW程式自行判斷</param> 
    ''' <param name="assignNextStepUser">下一步驟的員工編號。預設值為NULL，由FLOW程式自行判斷</param>
    ''' <param name="bSendToEnd">直接結案</param>
    ''' <returns></returns>
    ''' <remarks>sWorkingUserId及sLoginUserId若沒有特別的理由，應設為""由FLOW程式自行判斷</remarks>
    'Public Function StartFlow(ByVal nBraDepNo As Integer,
    '                          ByVal sFlowName As String,
    '                          ByVal bAutoSendToNext As Boolean,
    '                          ByVal sNextUserid As String) As StepInfo

    '    Dim stepUser() As STEPNO_USERID = Nothing

    '    If (Not String.IsNullOrEmpty(sNextUserid)) Then
    '        stepUser = STEPNO_USERID.arrayStepUser("", sNextUserid)
    '    End If

    '    Return getELoanFlow().StartFlow(nBraDepNo, sFlowName, bAutoSendToNext, stepUser)
    'End Function

    ''' <summary>
    ''' 直接起案並回傳new caseid
    ''' </summary>
    ''' <param name="sBRID">分行別</param>
    ''' <param name="sFlowName">Flow Name</param>
    ''' <param name="bAutoSendToNext">起案後，自動送至下一步驟</param>
    ''' <param name="sNextUserid">下一步驟的員工編號。預設值為NULL，由FLOW程式自行判斷</param> 
    ''' <param name="bSendToEnd">直接結案</param>
    ''' <returns></returns>
    ''' <remarks>sWorkingUserId及sLoginUserId若沒有特別的理由，應設為""由FLOW程式自行判斷</remarks>
    Public Function StartFlow(ByVal sBRID As String,
                              ByVal sFlowName As String,
                              Optional ByVal bAutoSendToNext As Boolean = False,
                              Optional ByVal sNextUserid As String = "",
                              Optional ByVal bSendToEnd As Boolean = False,
                              Optional ByVal sApvCasId As String = "",
                              Optional ByVal sAppBrid As String = "",
                              Optional ByVal sCplAplId As String = "",
                              Optional ByVal sCplAplNam As String = "",
                              Optional ByVal bSendMail As Boolean = True) As StepInfo

        Dim stepUser() As STEPNO_USERID = Nothing

        If (Not String.IsNullOrEmpty(sNextUserid)) Then
            stepUser = STEPNO_USERID.arrayStepUser("", sNextUserid)
        End If

        Return getELoanFlow().StartFlow(getBraDepNo(sBRID), sFlowName, bAutoSendToNext, stepUser, bSendToEnd, sApvCasId, sAppBrid, sCplAplId, sCplAplNam, bSendMail)
    End Function
    ''' <summary>
    ''' 傳入案號==>起案
    ''' </summary>
    ''' <param name="sCaseId">案號</param>
    ''' <param name="bAutoSendToNext">起案，自動送至下一步驟</param>
    ''' <param name="assignNextStepUser">可指定多個(分支)下個步驟的負責人員, 可為Nothing, 表示不指定</param>
    ''' <returns></returns>
    ''' <remarks>sWorkingUserId及sLoginUserId若沒有特別的理由，應設為""由FLOW程式自行判斷</remarks>
    Public Function StartFlowByCaseId(ByVal sCaseId As String,
                                      ByVal bAutoSendToNext As Boolean,
                                      Optional ByVal assignNextStepUser() As STEPNO_USERID = Nothing,
                                      Optional ByVal sApvCasId As String = "",
                                      Optional ByVal sAppBrid As String = "",
                                      Optional ByVal sCplAplId As String = "",
                                      Optional ByVal sCplAplNam As String = "",
                                      Optional ByVal bSendMail As Boolean = True) As StepInfo

        Return getELoanFlow().StartFlowByCaseId(sCaseId, bAutoSendToNext, assignNextStepUser, False, sApvCasId, sAppBrid, sCplAplId, sCplAplNam, bSendMail)
    End Function

    Public Function StartFlowByCaseId(ByVal sCaseId As String,
                                      ByVal bAutoSendToNext As Boolean,
                                      ByVal sNextUserid As String,
                                      Optional ByVal sApvCasId As String = "",
                                      Optional ByVal sAppBrid As String = "",
                                      Optional ByVal sCplAplId As String = "",
                                      Optional ByVal sCplAplNam As String = "",
                                      Optional ByVal bSendMail As Boolean = True) As StepInfo

        Dim stepUser() As STEPNO_USERID = Nothing

        If (Not String.IsNullOrEmpty(sNextUserid)) Then
            stepUser = STEPNO_USERID.arrayStepUser("", sNextUserid)
        End If

        Return getELoanFlow().StartFlowByCaseId(sCaseId, bAutoSendToNext, stepUser, False, sApvCasId, sAppBrid, sCplAplId, sCplAplNam, bSendMail)
    End Function


    Public Function StartFlowByCaseId(ByVal sCaseId As String,
                                      ByVal sFirstStepNo As String,
                                      Optional ByVal bAutoSendToNext As Boolean = False,
                                      Optional ByVal sNextUserid As String = Nothing,
                                      Optional ByVal sApvCasId As String = "",
                                      Optional ByVal sAppBrid As String = "",
                                      Optional ByVal sCplAplId As String = "",
                                      Optional ByVal sCplAplNam As String = "",
                                      Optional ByVal bSendMail As Boolean = True) As StepInfo

        Dim stepUser() As STEPNO_USERID = Nothing

        If (Not String.IsNullOrEmpty(sFirstStepNo)) AndAlso
            (Not String.IsNullOrEmpty(sNextUserid)) Then
            stepUser = STEPNO_USERID.arrayStepUser(sFirstStepNo, sNextUserid)
        End If

        Return getELoanFlow().StartFlowByCaseId(sCaseId, bAutoSendToNext, stepUser)
    End Function


    ''' <summary>
    ''' 取得new caseid
    ''' </summary>
    ''' <param name="sSubSysId">子系統代號ex:06,05,04</param>
    ''' <param name="nBraDepNo">分行別</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetNewCaseId(ByVal sSubSysId As String, ByVal nBraDepNo As Integer) As String
        Return getSYCaseId().GetNewCaseId(sSubSysId, nBraDepNo)
    End Function

    Public Function GetNewCaseId(ByVal sSubSysId As String, ByVal sBrid As String) As String
        Return getSYCaseId().GetNewCaseId(sSubSysId, getBraDepNo(sBrid))
    End Function

    Public Function GetNewCaseIdByFlowName(ByVal sFlowName As String, ByVal sBrid As String) As String
        Return getSYCaseId().GetNewCaseId(
            getSYFlowId.GetSubSysId(sFlowName), getBraDepNo(sBrid))
    End Function


    ''' <summary>
    ''' 送關元件，並返回訊息
    ''' </summary>
    ''' <param name="sCaseId">案號</param>
    ''' <param name="nSubFlowSeq">Optional nSubFlowSeq, 0表示同一CASEID沒有同時運作的兩個流程步驟 </param>
    ''' <param name="sNextUserid">Optional sNextUserid </param>
    ''' <remarks></remarks>
    Public Function SendFlow(ByVal sCaseId As String, _
                             Optional ByVal nSubFlowSeq As Integer = 0,
                             Optional ByVal sNextStep As String = "",
                             Optional ByVal sNextUserid As String = "",
                             Optional ByVal nNextBraDepNo As Integer = 0) As StepInfo

        Dim nextStepUser() As STEPNO_USERID = Nothing
        Dim nextStepNoBraDepNo(0) As STEPNO_BRADEPNO

        If Not String.IsNullOrEmpty(sNextUserid) Then
            nextStepUser = STEPNO_USERID.arrayStepUser(sNextStep, sNextUserid)
        End If

        If Not String.IsNullOrEmpty(sNextStep) Then
            nextStepNoBraDepNo(0) = STEPNO_BRADEPNO.Pair(nNextBraDepNo, sNextStep)
            Return getELoanFlow().SendFlow(sCaseId, "N", nSubFlowSeq, nextStepNoBraDepNo, nextStepUser)
        Else
            If nNextBraDepNo <> 0 Then
                nextStepNoBraDepNo(0) = STEPNO_BRADEPNO.Pair(nNextBraDepNo, sNextStep)
                Return getELoanFlow().SendFlow(sCaseId, "N", nSubFlowSeq, nextStepNoBraDepNo, nextStepUser)
            End If
        End If

        Return getELoanFlow().SendFlow(sCaseId, "N", nSubFlowSeq, Nothing, nextStepUser)
    End Function


    ''' <summary>
    ''' 送關元件，並返回訊息,
    ''' 可指定多個(分支)下個步驟的負責人員
    ''' </summary>
    ''' <param name="sCaseId"></param>
    ''' <param name="nSubFlowSeq">0表示由流程決定</param>
    ''' <param name="assignNextStepBradepNo">可送至多個分支的下個步驟</param>
    ''' <param name="assignNextStepUser">可指定多個(分支)下個步驟的負責人員</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SendFlow(ByVal sCaseId As String, _
                             ByVal nSubFlowSeq As Integer,
                             ByVal assignNextStepBradepNo() As STEPNO_BRADEPNO,
                             ByVal assignNextStepUser() As STEPNO_USERID) As StepInfo

        Return getELoanFlow().SendFlow(sCaseId, "N", nSubFlowSeq, assignNextStepBradepNo, assignNextStepUser)
    End Function



    ''' <summary>
    ''' 送關元件(分流時使用)，可指定下一步驟(多條子流程)的使用者
    ''' </summary>
    ''' <param name="sCaseId">案號</param>
    ''' <param name="nSubFlowSeq">Optional nSubFlowSeq, 0表示同一CASEID沒有同時運作的兩個流程步驟 </param>
    ''' <remarks></remarks>
    Public Function SendFlow(ByVal sCaseId As String,
                             ByVal sPN As String,
                             ByVal nSubFlowSeq As Integer,
                             ByVal assignNextStepBradepNo() As STEPNO_BRADEPNO,
                             Optional ByVal assignNextStepUser() As STEPNO_USERID = Nothing,
                             Optional ByVal sRevisionNote As String = "") As StepInfo
        Return getELoanFlow().SendFlow(sCaseId, sPN, nSubFlowSeq, assignNextStepBradepNo, assignNextStepUser, sRevisionNote)
    End Function



    ''' <summary>
    ''' 退關元件(修正補充)
    ''' </summary>
    ''' <param name="sCaseId">案號</param>
    ''' <param name="nSubFlowSeq">理由</param>
    ''' <param name="sRevisionNotes">理由</param>
    ''' <param name="destStpNoBradepno">Optional 指定某一分行的某一步驟代碼</param>
    ''' <param name="sNextUserid">Optional 指定某一user，員工編號ex:S064933</param>
    ''' <remarks></remarks>
    Public Function JumpRollBack(ByVal sCaseId As String,
                                 ByVal nSubFlowSeq As Integer,
                                 ByVal sRevisionNotes As String,
                                 ByVal destStpNoBradepno As STEPNO_BRADEPNO,
                                 Optional ByVal sNextUserid As String = "") As StepInfo

        Dim nextStepUser() As STEPNO_USERID = Nothing

        If (Not String.IsNullOrEmpty(sNextUserid)) AndAlso (Not String.IsNullOrEmpty(sNextUserid)) Then
            nextStepUser = STEPNO_USERID.arrayStepUser("", sNextUserid)
        End If


        If IsNothing(destStpNoBradepno) Then
            Return getELoanFlow().SendFlow(sCaseId, "P", nSubFlowSeq, Nothing, nextStepUser, sRevisionNotes)
        Else
            Dim nextStepBradepno(0) As STEPNO_BRADEPNO
            nextStepBradepno(0) = destStpNoBradepno

            Return getELoanFlow().SendFlow(sCaseId, "P", nSubFlowSeq, nextStepBradepno, nextStepUser, sRevisionNotes)
        End If

    End Function


    Public Function JumpRollBack(ByVal sCaseId As String,
                             Optional ByVal nSubFlowSeq As Integer = 0,
                             Optional ByVal sRevisionNotes As String = "",
                             Optional ByVal sDestStpNo As String = "",
                             Optional ByVal sNextUserid As String = "") As StepInfo

        Dim nextStepBradepno As STEPNO_BRADEPNO = Nothing

        If Not String.IsNullOrEmpty(sDestStpNo) Then
            nextStepBradepno = STEPNO_BRADEPNO.Pair(0, sDestStpNo)
        End If

        Return JumpRollBack(sCaseId, nSubFlowSeq, sRevisionNotes, nextStepBradepno, sNextUserid)
    End Function


    ''' <summary>
    ''' 子流程完成後再追加執行子流程
    ''' </summary>
    ''' <param name="sCaseid"></param>
    ''' <param name="nParentSubFlowSeq"></param>
    ''' <param name="sNextStep"></param>
    ''' <param name="sNextBraDepNo"></param>
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
                                             Optional ByVal sNextBraDepNo As Integer = 0,
                                             Optional ByVal sPreviousStep As String = Nothing,
                                             Optional ByVal assignUser As String = Nothing,
                                             Optional ByVal bStartFromSplitter As Boolean = True,
                                             Optional ByVal sRevisionNote As String = "",
                                             Optional ByVal sRevisionSrcStep As String = "") As StepInfoItemExt

        Return getELoanFlow().ExecuteAdditionalSubflow(sCaseid,
                                                       nParentSubFlowSeq,
                                                       sNextStep,
                                                       sNextBraDepNo,
                                                       sPreviousStep,
                                                       assignUser,
                                                       bStartFromSplitter,
                                                       sRevisionNote,
                                                       sRevisionSrcStep)
    End Function


    ''' <summary>
    ''' 預設某一步驟的OWNER
    ''' </summary>
    ''' <param name="sCaseid"></param>
    ''' <param name="nCurrentSubFlowSeq"></param>
    ''' <param name="nextStepUser"></param>
    ''' <remarks>當流程經過此步驟時，會使用預設的OWNER</remarks>
    Public Sub AssignNextStepOwner(ByVal sCaseid As String, ByVal nCurrentSubFlowSeq As Integer,
                                   ByVal nextStepUser As STEPNO_USERID)
        getSYFlowStep.AssignNextStepOwner(sCaseid, nCurrentSubFlowSeq, nextStepUser)
    End Sub


    ''' <summary>
    ''' 取消案件，流程留下紀錄
    ''' </summary>
    ''' <param name="sCaseId">案號</param>
    ''' <remarks>
    ''' 須注意各流程取消案件後，自身業務的flag是否需要回復回原本狀態
    ''' </remarks>
    Public Sub DisableFlow(ByVal sCaseId As String)

        Dim sSenderId As New USER_ID_NAME
        Dim sWorkingUserId As New USER_ID_NAME
        getELoanFlow.GetCurrentUserid(sSenderId, sWorkingUserId)

        getSYCaseId().DisableFlow(sCaseId, sSenderId.userId)
    End Sub



    ''' <summary>
    ''' 刪除流程，流程不留紀錄
    ''' </summary>
    ''' <param name="sCaseId">案號</param>
    ''' <remarks>
    ''' 須注意各流程刪除後，自身業務的flag是否需要回復回原本狀態 
    ''' </remarks>
    Public Sub DeleteFlow(ByVal sCaseId As String)
        getSYCaseId().DeleteFlow(sCaseId)
    End Sub


    ''' <summary>
    ''' 取得流程名稱
    ''' </summary>
    ''' <param name="sCaseId">案號</param>
    ''' <returns>流程名稱</returns> 
    Public Function GetFlowInfo(ByVal sCaseId As String) As DataRow
        Return Nothing
    End Function


    ''' <summary>
    ''' 取得目前步驟
    ''' </summary>
    ''' <param name="sCaseId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetCurrStepNO(ByVal sCaseId As String) As String
        Return Nothing
    End Function

    ''' <summary>
    ''' 依子系統、流程名稱、流程步驟、分行別取得使用者
    ''' </summary>
    ''' <param name="sSubSysId">子系統</param>
    ''' <param name="sFlowName">流程名稱</param>
    ''' <param name="sStepNO">流程步驟</param>
    ''' <param name="sBrid">分行別</param>
    ''' <param name="sBrDivid"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetGrpUserList(ByVal sSubSysId As String, ByVal sFlowName As String, ByVal sStepNO As String, _
                                                      ByVal sBrid As String, Optional ByVal sBrDivid As String = "NA") As ArrayList
        Return Nothing
    End Function



    ''' <summary>
    ''' 取得己處理工作列表
    ''' </summary>
    ''' <param name="sStaffid"></param>
    ''' <param name="sBrid"></param>
    ''' <param name="nStartRow">要取得幾筆之後資料</param>
    ''' <param name="nEndRow">要取得幾筆之前資料</param>
    ''' <returns>GetDoneCaseList("S001024",1,10) 取得第1-10筆的資料</returns>
    Public Function GetDoneCaseList(ByVal sStaffid As String,
                                    ByVal sBrid As String,
                                    ByVal nStartRow As Integer,
                                    ByVal nEndRow As Integer,
                                    ByVal sOrderBy As String) As DataTable
        Return getSYFlowStep().GetCaselistByUseridStatus(sStaffid, sBrid, 3, nStartRow, nEndRow, sOrderBy)
    End Function


    ''' <summary>
    '''  取得待處理工作列表
    ''' </summary>
    ''' <param name="sStaffid"></param>
    ''' <param name="sBrid"></param>
    ''' <param name="nStartRow">要取得幾筆之後資料</param>
    ''' <param name="nEndRow">要取得幾筆之前資料</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetPendingCaseList(ByVal sStaffid As String,
                                       ByVal sBrid As String,
                                       ByVal nStartRow As Integer,
                                       ByVal nEndRow As Integer,
                                       Optional ByVal sOrderBy As String = "CASEID") As DataTable
        Return getSYFlowStep().GetCaselistByUseridStatus(sStaffid, sBrid, 1, nStartRow, nEndRow, sOrderBy)
    End Function

    ''' <summary>
    ''' 取得佇列取件內的案件
    ''' </summary>
    ''' <param name="sUserid"></param>
    ''' <param name="sBrid"></param>
    ''' <param name="sOrderBy">排序的欄位，欄位後加入 DESC為反向排序</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetCaseListFromQueue(ByVal sUserid As String,
                                         ByVal sBrid As String,
                                         Optional ByVal sOrderBy As String = Nothing) As DataTable

        Try
            Return getSYFlowStep().GetCaseListFromQueue(sUserid, sBrid, sOrderBy, Nothing, 0, Nothing, True)
        Catch ex As SYException
            If ex.Code = SYMSG.SYFLOW_CASE_NOT_FOUND OrElse
                ex.Code = SYMSG.SYFLOW_CASENOTFOUND_PROHIBITESAMEUSER Then
                Return CType(ex.Obj, DataTable)
            End If

            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function


    ''' <summary>
    ''' 由 STEPNO取得ROLE列表
    ''' </summary>
    ''' <param name="sStepNo">若使用"041000%"表示，搜尋時會將041000??步驟的所有使用者列出</param>
    ''' <param name="sBrids">指定哪些分行的USER，可以不設，不設表示全部的分行</param>
    ''' <param name="nBraDepNos">指定哪些BRA_DEPNO的USER，可以不設，不設表示全部的分行</param>
    ''' <returns>return a datatable of rows which contains CASEID, ROLEID and ROLENAME</returns>
    ''' <remarks>可使用#_，</remarks>
    Public Function GetRoleListByStepNo(ByVal sStepNo As String,
                                        Optional ByVal sBrids() As String = Nothing,
                                        Optional ByVal nBraDepNos() As Integer = Nothing) As DataTable
        '取消，可同時指定BRID及BRADEPNO
        'Dbg.Assert(Not (String.IsNullOrEmpty(sBrid) = False AndAlso nBraDepNo <> 0),
        '    "不可同時指定sBrid及nBraDepNo,最多只能指定一個(sBrid或nBraDepNo)")

        Return getSYRelRoleFlowMap().GetRoleListByStepNo(sStepNo, sBrids, nBraDepNos)
    End Function


    ''' <summary>
    ''' 由 STEPNO取得USER列表
    ''' </summary>
    ''' <param name="sStepNo"></param>
    ''' <param name="sBrid">指定哪一分行的USER，可以不設，不設表示全部的分行</param>
    ''' <param name="nBraDepNo">指定哪一BRA_DEPNO的USER，可以不設，不設表示全部的分行</param>
    ''' <returns>return a datatable of rows which contains STAFFID and USERNAME</returns>
    ''' <remarks>sBrid或nBraDepNo可以不設或只能設一個</remarks>
    Public Function GetUserListByStepNo(ByVal sStepNo As String,
                                        Optional ByVal sBrid As String = Nothing,
                                        Optional ByVal nBraDepNo As Integer = 0) As DataTable
        Return getSYRelRoleFlowMap().GetUserListByStepNo(sStepNo, sBrid, nBraDepNo)
    End Function


    ''' <summary>
    ''' 將佇列取件內的案件分派給哪一使用者　　
    ''' </summary>
    ''' <param name="sCaseid"></param>
    ''' <param name="sOwner">指定將案件分派給哪一使用者</param>
    ''' <param name="sBrid">分行代碼，可為0，表示由程式判斷</param>
    ''' <param name="nSubFlowSeq">子流程編號，可為0，表示由程式判斷</param>
    ''' <param name="sStepNo">指定流程步驟，可為NOTHING</param>
    ''' <remarks></remarks>
    Public Function SetOwnerOfCaseFromQueue(ByVal sCaseid As String,
                                            ByVal sOwner As String,
                                            Optional ByVal sBrid As String = Nothing,
                                            Optional ByVal nSubFlowSeq As Integer = 0,
                                            Optional ByVal sStepNo As String = Nothing) As StepInfoItemExt
        Return getSYCaseId.SetOwnerOfCaseFromQueue(sCaseid, sOwner, sBrid, nSubFlowSeq, sStepNo)
    End Function


    ''' <summary>
    ''' 案件改分派
    ''' </summary>
    ''' <param name="sCurrentCaseid">目前的案件編號</param>
    ''' <param name="sNewClient">改分派給誰</param>
    ''' <param name="nNewBraDepNo">新CLIENT所在的BRADEPNO</param>
    ''' <param name="nCurrentSubFlowSeq">目前案件的子流程編號</param>
    ''' <param name="sCurrentStepNo">目前案件的步驟代碼</param>
    ''' <param name="sChangeUserId">案件步驟的變更人員</param>
    ''' <param name="dtChangeDateTime">案件步驟的變更日期</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function SetClientOfCase(ByVal sCurrentCaseid As String,
                                    ByVal sNewClient As String,
                                    Optional ByVal nNewBraDepNo As Integer = 0,
                                    Optional ByVal nCurrentSubFlowSeq As Integer = 0,
                                    Optional ByVal sCurrentStepNo As String = Nothing,
                                    Optional ByVal sChangeUserId As String = Nothing,
                                    Optional ByVal dtChangeDateTime As DateTime = Nothing
                                    ) As StepInfoItemExt
        Return getSYCaseId.SetClientOfCase(sCurrentCaseid,
                                               sNewClient,
                                               nNewBraDepNo,
                                               nCurrentSubFlowSeq,
                                               sCurrentStepNo,
                                               sChangeUserId,
                                               dtChangeDateTime)

    End Function



    ''' <summary>
    ''' 取得目前案件最新的步驟資訊
    ''' </summary>
    ''' <param name="sCaseId"></param>
    ''' <param name="nSubFlowSeq"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetCurrentStepInfoOfCaseid(ByVal sCaseId As String,
                                               Optional ByVal nSubFlowSeq As Integer = 0) As StepInfoItemExt

        Return getELoanFlow().GetCurrentStepInfoOfCaseid(sCaseId, nSubFlowSeq)
    End Function



    ''' <summary>
    ''' 設定案件的核准編號, 申請單位, 授信戶編號, 授信戶名稱
    ''' </summary>
    ''' <param name="sCaseId">案件編號</param>
    ''' <param name="sApvCasId">核准編號</param>
    ''' <param name="sAppBrid">申請單位</param>
    ''' <param name="sCplAplId">授信戶編號</param>
    ''' <param name="sCplAplNam">授信戶名稱</param>
    ''' <remarks></remarks>
    Public Sub SetCaseInfo(ByVal sCaseId As String,
                           ByVal sApvCasId As String,
                           ByVal sAppBrid As String,
                           ByVal sCplAplId As String,
                           ByVal sCplAplNam As String)
        getSYCaseId.SetCaseInfo(sCaseId, sApvCasId, sAppBrid, sCplAplId, sCplAplNam)
    End Sub



    ''' <summary>
    ''' 流程案件是否已結束
    ''' </summary>
    ''' <param name="sCaseid"></param>
    ''' <param name="nSubflowSeq"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function IsCaseCompleted(ByVal sCaseid As String, Optional ByVal nSubflowSeq As Integer = 0) As Boolean
        Return getSYFlowIncident.IsCaseCompleted(sCaseid, nSubflowSeq)

    End Function



    ''' <summary>
    ''' 輸入CASEID及STEPNO，取得取近一筆的記錄
    ''' </summary>
    ''' <param name="sCaseId"></param>
    ''' <param name="sStepNo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function GetLatestSubFlowNoByStepNo(ByVal sCaseId As String, ByVal sStepNo As String) As Integer
        Try
            Return getSYFlowStep.GetLatestSubFlowNoByStepNo(sCaseId, sStepNo)
        Catch ex As Exception
            Throw
        End Try
    End Function



    ''' <summary>
    ''' 取得目前子流程編號的PARENT_SEQ
    ''' </summary>
    ''' <param name="sCaseid"></param>
    ''' <param name="nCurrentSubFlowSeq"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetParentSubFlowSeq(ByVal sCaseid As String, nCurrentSubFlowSeq As Integer) As Integer
        Return getSYFlowIncident.GetParentSubFlowSeq(sCaseid, nCurrentSubFlowSeq)
    End Function



    ''' <summary>
    ''' 關閉所有子流程
    ''' </summary>
    ''' <param name="caseid"></param>
    ''' <param name="nSubFlowSeq"></param>
    ''' <remarks></remarks>
    Public Sub CloseAllSubFlowIncident(ByVal caseid As String,
                                       ByVal nSubFlowSeq As Integer,
                                       Optional ByVal bIncludeParent As Boolean = False,
                                       Optional ByVal bUpdateSenderIfNothing As Boolean = True)

        Dim sii As New StepInfoItemExt
        sii.caseId = caseid
        sii.subflowSeq = nSubFlowSeq

        getSYFlowIncident.CloseAllSubFlowIncident(sii, bIncludeParent, bUpdateSenderIfNothing)
    End Sub




    ''' <summary>
    ''' 取得已記錄的流程中，SUBFLOW_SEQ及SUBFLOW_COUNT的最大值
    ''' </summary>
    ''' <param name="sCaseid"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetSubFlowInfoMax(ByVal sCaseid As String, Optional ByVal sStepNo As String = Nothing) As DataRow
        Return getSYFlowStep.GetSubFlowInfoMax(sCaseid, sStepNo)
    End Function


    ''' <summary>
    ''' 由sCaseid取得Flowname
    ''' </summary>
    ''' <param name="sCaseid"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetFlownameByCaseid(ByVal sCaseid As String) As String
        Return getSYFlowId.GetFlownameByCaseid(sCaseid)
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
        Return getSYFlowDef.GetNextStepByFlownameStepnoPn(sFlowName, sStepNo, sPN)
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
        Return getSYRelRoleFlowMap.GetBranchRoleUserFlowidStepnoInfoByFlownameStepnoStaffid(sFlowName, sStepNo, sStaffId)
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
        Return getSYBranch.GetBridByFlownameStepno(sFlowname, sStepNo, sStaffid)
    End Function




    ''' <summary>
    ''' 由使用者編號取得角色代碼
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetRolesByStaffid(ByVal sStaffid As String) As DataTable
        Return getSYRelRoleUser.GetRolesByStaffid(sStaffid)
    End Function

    ''' <summary>
    ''' 取得案件下一步驟的資訊
    ''' </summary>
    ''' <param name="sCaseId">案號</param>
    ''' <param name="sPN">往前或往後(送件或修正補充)</param>
    ''' <param name="nSubFlowSeq">子流程代碼</param>
    ''' <param name="assignNextStepBradepNo">指定要送至下一步驟或分行，若沒指定表示由流程定義自行決定</param>
    ''' <param name="assignNextStepUser">指定下一步要使用的STEPUSER，若沒指定表示由流程定義自行決定</param>
    ''' <param name="bGetStepNoOnly">只取得下一步驟的流程步驟，不取得其它資訊(員編，分行，日期，SUMMARY等)</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetNextStepInfoByCaseid(ByVal sCaseId As String,
                                ByVal sPN As String,
                                Optional ByVal nSubFlowSeq As Integer = 0,
                                Optional ByVal assignNextStepBradepNo() As STEPNO_BRADEPNO = Nothing,
                                Optional ByVal assignNextStepUser() As STEPNO_USERID = Nothing,
                                Optional ByVal bGetStepNoOnly As Boolean = True) As StepInfoItemExt()

        Dim siie As StepInfoItemExt

        siie = GetCurrentStepInfoOfCaseid(sCaseId, nSubFlowSeq)
        Return getELoanFlow().GetNextStep(siie, sPN, assignNextStepBradepNo, assignNextStepUser, bGetStepNoOnly)

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
        Return getSYFlowStep().GetAlternativeUserList(sCaseId, nSubFlowSeq)
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
        Return getSYFlowStep.GetPendingCaseList(sBrid, nFlowId, sStepNo)
    End Function


End Class
