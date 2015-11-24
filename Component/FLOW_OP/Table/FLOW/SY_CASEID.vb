Option Explicit On
Option Strict On

Imports com.Azion.NET.VB
Imports System.Data.SqlClient


Namespace TABLE

    Public Class SY_CASEID
        Inherits SY_TABLEBASE

        ''' <summary>
        ''' Constructor
        ''' </summary>
        ''' <param name="dbManager">DatabaseManager</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal dbManager As com.Azion.NET.VB.DatabaseManager)
            MyBase.new("SY_CASEID", dbManager)
        End Sub

        Public Shared Function getNewInstance(ByVal dbManager As com.Azion.NET.VB.DatabaseManager) As SY_CASEID
            Return New SY_CASEID(dbManager)
        End Function

        ''' <summary>
        ''' 取新的案號
        ''' </summary>
        ''' <param name="sSubsysId"></param>
        ''' <param name="nBraDepNo"></param>
        ''' <param name="sType">0:非八碼案號 ,1:8碼案件編號 ,2:核准編號</param>
        ''' <returns></returns>
        ''' <remarks>TALBE內的SUBSYSID存放的是BRID</remarks>
        Public Function GetNewCaseId(ByVal sSubsysId As String, ByVal nBraDepNo As Integer, Optional ByVal sType As String = "0") As String

            Dim nRepeatCount As Integer = 0
            Dim dbMaxId As BosBase
            Dim nMaxId As Integer
            Dim nYear100 As Integer

            Try
                nYear100 = Year(DateTime.Now) - 1911

                '輸入檢查
                Dbg.Assert(IsNothing(sSubsysId) = False, "參數空值錯誤")

                If IsNothing(sSubsysId) Then
                    Throw New SYException("參數空值錯誤", SYMSG.ERROR_INVALID_PARAMETER)
                End If

                If sSubsysId.Length <> 2 Then
                    Throw New SYException("參數錯誤", SYMSG.ERROR_INVALID_PARAMETER)
                End If

                Dbg.Assert(sType = "0" OrElse sType = "1" OrElse sType = "2")
                If sType <> "0" AndAlso sType <> "1" AndAlso sType <> "2" Then
                    Throw New SYException("參數錯誤", SYMSG.ERROR_INVALID_PARAMETER)
                End If


                '"取號時須使用Transaction
                Dbg.Assert(IsNothing(getDatabaseManager().getTransaction()) = False, "取號時須使用Transaction", False)

                dbMaxId = New BosBase("MAXID", getDatabaseManager())
                Dim sBrid As String = getSYBranch().GetBRID(nBraDepNo)

                If String.IsNullOrEmpty(sBrid) Then
                    Throw New SYException("無法取得分行代碼", SYMSG.SYCASEID_BRID_NOT_FOUND)
                End If

REPEAT:
                nRepeatCount = nRepeatCount + 1

                If dbMaxId.GetCount(BosBase.PARAM_ARRAY("SUBSYSID", sBrid, "IDTYPE", sType, "YEAR", nYear100)) = 0 Then
                    dbMaxId.Insert("SUBSYSID", sBrid,
                                   "IDTYPE", sType,
                                   "YEAR", nYear100,
                                   "MAXID", 0,
                                   "PURPOSE", IIf(sType = "0", "非八碼案號", IIf(sType = "1", "8碼案件編號用", "核准編號")))
                End If

                nMaxId = CInt(dbMaxId.ExecuteScalar( _
                        "select MAXID " & vbCrLf & _
                        "  from MAXID " & vbCrLf & _
                        " where SUBSYSID = @SUBSYSID@ " & vbCrLf & _
                        "   and YEAR = @YEAR@ " & vbCrLf & _
                        "   and IDTYPE = @IDTYPE@ " & vbCrLf,
                        "SUBSYSID", sBrid,
                        "YEAR", nYear100,
                        "IDTYPE", sType))

                '上線後之案件編號，2012年度各區均從5001開始
                'If Year(Now) = 2012 Then
                '    nMaxId = Math.Max(5000, nMaxId)
                'End If

                nMaxId = nMaxId + 1

                dbMaxId.ExecuteNonQuery( _
                        "update MAXID " & vbCrLf & _
                        "   set MAXID = @MAXID@ " & vbCrLf & _
                        " where SUBSYSID = @SUBSYSID@ " & vbCrLf & _
                        "   and YEAR = @YEAR@ " & vbCrLf & _
                        "   and IDTYPE = @IDTYPE@ " & vbCrLf,
                        "SUBSYSID", sBrid,
                        "YEAR", nYear100,
                        "IDTYPE", sType,
                        "MAXID", nMaxId)


                Dim TheMaxID As String = nMaxId.ToString

                '組案號
                'Dim sYear As String = StrUtility.FillZero(CStr(CInt(Year(Today)) - 1911), 3)
                'Dim sResultCaseId As String = sSubsysId & sBrid & sYear & "9" & StrUtility.FillZero(TheMaxID, 6)
                Dim sResultCaseId As String = String.Empty

                If sType = "0" Then
                    sResultCaseId = sSubsysId & sBrid & nYear100.ToString("D3") & "9" & Right(TheMaxID.PadLeft(6, "0"c), 6)

                    'String.Format("{0:D3}9{1:D6}", CInt(Year(Today)) - 1911, TheMaxID)

                    If sResultCaseId.Length <> 15 Then
                        Throw New SYException("取號的號碼格式錯誤", SYMSG.SYCASEID_WRONG_FORMAT)
                    End If

                    If nRepeatCount <= 100 Then
                        If GetCount(BosBase.PARAM_ARRAY("CASEID", sResultCaseId)) <> 0 Then
                            GoTo REPEAT
                        End If
                    Else
                        Throw New SYException("取號的號碼重複而且重新取號100次仍然無法取得新案號", SYMSG.SYCASEID_WRONG_DUPLICATION)
                    End If


                ElseIf sType = "2" Then
                    sResultCaseId = sBrid & nYear100.ToString("D3") & Right(TheMaxID.PadLeft(4, "0"c), 4)

                    If sResultCaseId.Length <> 10 Then
                        Throw New SYException("取號的號碼格式錯誤", SYMSG.SYCASEID_WRONG_FORMAT)
                    End If

                    If nRepeatCount <= 100 Then
                        If GetCount(BosBase.PARAM_ARRAY("APV_CAS_ID", sResultCaseId)) <> 0 Then
                            GoTo REPEAT
                        End If
                    Else
                        Throw New SYException("取號的號碼重複而且重新取號100次仍然無法取得新案號", SYMSG.SYCASEID_WRONG_DUPLICATION)
                    End If


                Else
                    Throw New SYException("取號錯誤。", SYMSG.UNDEFINE, dbMaxId.GetLastSQL)
                End If


                Return sResultCaseId

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Function


        ''' <summary>
        ''' 將案號寫入至SY_CASEID內
        ''' </summary>
        ''' <param name="sSysid"></param>
        ''' <param name="sSubSysId"></param>
        ''' <param name="sCaseid"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Function Insert(ByVal sSysid As String,
                                         ByVal sSubSysId As String,
                                         ByVal nFlowId As Integer,
                                         ByVal sCaseid As String,
                                         ByVal nBraDepNo As Integer,
                                         Optional ByVal nAppBraDepNo As Integer = 0) As Boolean
            Try
                Dim nResult As Integer
                Dim arrayBosParameter As New BosParamsList

                'sCaseid的前兩位應為sSubsysid
                Dbg.Assert(String.Compare(sCaseid.Substring(0, 2), sSubSysId, True) = 0)

                arrayBosParameter.Add("CASEID", sCaseid)
                arrayBosParameter.Add("SYSID", sSysid)
                arrayBosParameter.Add("SUBSYSID", sSubSysId)
                arrayBosParameter.Add("FLOW_ID", nFlowId)
                arrayBosParameter.Add("BRA_DEPNO", nBraDepNo)

                If nAppBraDepNo <> 0 Then
                    arrayBosParameter.Add("APP_BRADEPNO", nAppBraDepNo)
                End If


                '寫入至DB內
                nResult = Insert(arrayBosParameter.ToArray)

                '應寫入一筆記錄
                Dbg.Assert(nResult = 1)

                '如果寫入一筆記錄，回傳TRUE
                If nResult = 1 Then
                    Return True
                End If

                Return False

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try
        End Function


        Public Function GetInfo(ByVal sCaseid As String) As DataRow
            Try
                Dim dataRow As DataRow
                dataRow = GetDataRow(Nothing, "CASEID", sCaseid)

                Return dataRow
            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try
        End Function


        ''' <summary>
        ''' 取消案件，流程留下紀錄
        ''' </summary>
        ''' <param name="sCaseId">案號</param>
        ''' <remarks>
        ''' 須注意各流程取消案件後，自身業務的flag是否需要回復回原本狀態
        ''' </remarks>
        Public Sub DisableFlow(ByVal sCaseId As String, ByVal sSenderId As String)

            Dim nResult As Integer
            Dim drc As DataRowCollection
            Dim dateNow As DateTime = Now

            Try
                nResult = ExecuteNonQuery(
                    "update SY_FLOWINCIDENT " & vbCrLf & _
                    "   set STATUS = 3 " & vbCrLf & _
                    " where STATUS <> 3 " & vbCrLf & _
                    "   and CASEID = @CASEID@ ",
                    "CASEID", sCaseId)

                If nResult = 0 Then
                    Throw New SYException(String.Format("案件:{0} 已結束，不可取消。", sCaseId),
                                          SYMSG.SYCASEID_ISCLOSE)
                End If

                drc = GetDataRowCollection(
                    "select SY_FLOWSTEP.*, SY_CASEID.FLOW_ID " & vbCrLf & _
                    "  from SY_FLOWSTEP " & vbCrLf & _
                    " inner join SY_CASEID " & vbCrLf & _
                    "    on SY_CASEID.CASEID=SY_FLOWSTEP.CASEID  " & vbCrLf & _
                    " where SY_FLOWSTEP.CASEID = @CASEID@ " & vbCrLf & _
                    "   and PROCESSTIME is null ",
                    "CASEID", sCaseId)

                For Each dr As DataRow In drc
                    dr("PROCESSTIME") = dateNow
                    'dr("STATUS") = 3
                    dr("SENDER") = sSenderId
                    UpdateFlowStep(dr)

                    dr("STARTTIME") = dateNow
                    dr("PROCESSTIME") = dateNow
                    dr("SUMMARY") = getSYFlowStep().GetSummary(CInt(dr("BRA_DEPNO")), CInt(dr("FLOW_ID")), SY_STEP_NO.SystemStep_Cancel)
                    dr("STATUS") = 3
                    dr("STEP_NO") = SY_STEP_NO.SystemStep_Cancel
                    dr("CLIENT") = sSenderId
                    dr("OWNER") = sSenderId
                    dr("SENDER") = sSenderId
                    InsertFlowStep(dr)
                Next

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)

            End Try
        End Sub



        Protected Function UpdateFlowStep(ByVal datarow As DataRow) As Integer

            Dim arrayBosParameters As New BosParamsList

            Try
                For Each column As DataColumn In datarow.Table.Columns
                    If IsNothing(datarow(column.ColumnName)) = False AndAlso column.ColumnName <> "FLOW_ID" Then
                        arrayBosParameters.Add(column.ColumnName, datarow(column.ColumnName))
                    End If
                Next
                Return getSYFlowStep.Update(arrayBosParameters.ToArray())

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Function


        Protected Function InsertFlowStep(ByVal datarow As DataRow) As Integer

            Dim arrayBosParameters As New BosParamsList

            Try
                For Each column As DataColumn In datarow.Table.Columns
                    If IsNothing(datarow(column.ColumnName)) = False AndAlso column.ColumnName <> "FLOW_ID" Then
                        arrayBosParameters.Add(column.ColumnName, datarow(column.ColumnName))
                    End If
                Next
                Return getSYFlowStep.Insert(arrayBosParameters.ToArray())
            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Function


        ''' <summary>
        ''' 刪除流程，流程不留紀錄
        ''' </summary>
        ''' <param name="sCaseId">案號</param>
        ''' <remarks>
        ''' 須注意各流程刪除後，自身業務的flag是否需要回復回原本狀態 
        ''' </remarks>
        Public Sub DeleteFlow(ByVal sCaseId As String)

            Try
                getSYCaseRevision.Delete("CASEID", sCaseId)

                '刪除FLOWSTEP
                getSYFlowStep().Delete("CASEID", sCaseId)

                '刪除FLOWINCIDENT
                getSYFlowIncident().Delete("CASEID", sCaseId)

                '刪除CASEID
                Delete("CASEID", sCaseId)

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Sub


        ''' <summary>
        ''' 將佇列取件內的案件分派給哪一使用者　　
        ''' </summary>
        ''' <param name="sCaseid"></param>
        ''' <param name="sOwner">指定將案件分派給哪一使用者</param>
        ''' <param name="sBrid">分行代碼</param>
        ''' <param name="nSubFlowSeq">可為0，表示由程式判斷</param>
        ''' <remarks></remarks>
        Public Function SetOwnerOfCaseFromQueue(ByVal sCaseid As String,
                                                ByVal sOwner As String,
                                                Optional ByVal sBrid As String = Nothing,
                                                Optional ByVal nSubFlowSeq As Integer = 0,
                                                Optional ByVal sStepNo As String = Nothing) As StepInfoItemExt
            Dim dt As DataTable = Nothing
            Dim dr As DataRow = Nothing
            Dim arrayBosParameter As New BosParamsList

            Dim sii() As StepInfoItemExt

            Try
                dt = getSYFlowStep.GetCaseListFromQueue(sOwner, sBrid, Nothing, sCaseid, nSubFlowSeq, sStepNo)

                Dbg.Assert(dt.Rows.Count = 1)
                dr = dt.Rows(0)

                arrayBosParameter.Add("CASEID", dr("CASEID"))
                arrayBosParameter.Add("SUBFLOW_SEQ", dr("SUBFLOW_SEQ"))
                arrayBosParameter.Add("STEP_NO", dr("STEP_NO"))
                arrayBosParameter.Add("OWNER", sOwner)

                '更新DB的內容
                getSYFlowStep.ExecuteNonQuery(
                    "update SY_FLOWSTEP " & vbCrLf & _
                    "   set CLIENT = @OWNER@ " & vbCrLf & _
                    " where CASEID = @CASEID@ " & vbCrLf & _
                    "   and SUBFLOW_SEQ = @SUBFLOW_SEQ@ " & vbCrLf & _
                    "   and STEP_NO = @STEP_NO@; " & vbCrLf & _
                    "update SY_FLOWSTEP " & vbCrLf & _
                    "   set OWNER = @OWNER@ " & vbCrLf & _
                    " where CASEID = @CASEID@ " & vbCrLf & _
                    "   and SUBFLOW_SEQ = @SUBFLOW_SEQ@ " & vbCrLf & _
                    "   and STEP_NO = @STEP_NO@ " & vbCrLf & _
                    "   and OWNER is null ",
                    arrayBosParameter.ToArray)

                sii = getSYFlowStep.GetFlowStepInfo(CStr(dr("CASEID")),
                                                     CInt(dr("SUBFLOW_SEQ")),
                                                     CStr(dr("STEP_NO")),
                                                     0)

                If IsNothing(sii) OrElse sii.Count = 0 Then
                    Return Nothing
                End If

                Return sii(0)

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

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

            Dim drc As DataRowCollection
            Dim dr As DataRow
            Dim sii As StepInfoItemExt
            Dim arrayBosParameter As New BosParamsList
            Dim sSql As String
            Dim nValue As Integer

            Try

                drc = getSYFlowStep.GetIncompletedStepNoByCaseid(sCurrentCaseid,
                                              nCurrentSubFlowSeq,
                                              sCurrentStepNo)

                If IsNothing(drc) = True OrElse drc.Count = 0 Then
                    Throw New SYException(String.Format("案件沒有任何步驟可以改分派：CASEID={0}，SUBFLOW_SEQ={1}，STEP_NO={2}",
                                                        sCurrentCaseid, nCurrentSubFlowSeq, "" & sCurrentStepNo),
                                          SYMSG.SYCASEID_CANNOT_CHANGECLIENT_CSAEUNFOUND)
                End If

                If drc.Count > 1 Then
                    Throw New SYException(String.Format("案件有二個或二個以上的步驟可以改分派：CASEID={0}，SUBFLOW_SEQ={1}，STEP_NO={2}",
                                                      sCurrentCaseid, nCurrentSubFlowSeq, "" & sCurrentStepNo),
                                                  SYMSG.SYCASEID_CANNOT_CHANGECLIENT_AMBIGUOUSSTEP)
                End If

                dr = drc(0)

                arrayBosParameter.Add("CASEID", dr("CASEID"))
                arrayBosParameter.Add("SUBFLOW_SEQ", dr("SUBFLOW_SEQ"))
                arrayBosParameter.Add("STEP_NO", dr("STEP_NO"))
                arrayBosParameter.Add("OWNER", sNewClient)

                sSql = _
                    "update SY_FLOWSTEP " & vbCrLf & _
                    "   set CLIENT = @OWNER@, " & vbCrLf & _
                    "      OCLIENT = case " & vbCrLf & _
                    "                  when OCLIENT is null then CLIENT " & vbCrLf & _
                    "                  else OCLIENT " & vbCrLf & _
                    "                 end " & vbCrLf

                If nNewBraDepNo <> 0 Then
                    sSql &= "   ,BRA_DEPNO = @BRA_DEPNO@ " & vbCrLf
                    arrayBosParameter.Add("BRA_DEPNO", nNewBraDepNo)
                End If

                If String.IsNullOrEmpty(sChangeUserId) = False Then
                    sSql &= "   ,CHGUID = @CHGUID@ " & vbCrLf & _
                            "   ,CHGDATE = @CHGDATE@ " & vbCrLf

                    arrayBosParameter.Add("CHGUID", sChangeUserId)
                    arrayBosParameter.Add("CHGDATE", dtChangeDateTime)
                End If


                sSql &= _
                    " where CASEID = @CASEID@ " & vbCrLf & _
                    "   and SUBFLOW_SEQ = @SUBFLOW_SEQ@ " & vbCrLf & _
                    "   and STEP_NO = @STEP_NO@; " & vbCrLf & _
                    "Update SY_FLOWSTEP " & vbCrLf & _
                    "   set OWNER = @OWNER@ " & vbCrLf & _
                    " where CASEID = @CASEID@ " & vbCrLf & _
                    "   and SUBFLOW_SEQ = @SUBFLOW_SEQ@ " & vbCrLf & _
                    "   and STEP_NO = @STEP_NO@ " & vbCrLf & _
                    "   and OWNER is null "

                '更新DB的內容
                nValue = getSYFlowStep.ExecuteNonQuery(sSql, arrayBosParameter.ToArray)


                If nValue <> 1 Then
                    Throw New SYException(String.Format("案件無法改分配：CASEID={0}，CLIENT={1}，BRA_DEPNO={2}，SUBFLOW_SEQ={3}，STEP_NO={3}", sCurrentCaseid, sNewClient, nNewBraDepNo, nCurrentSubFlowSeq, "" & sCurrentStepNo),
                                          SYMSG.SYCASEID_CANNOT_CHANGECLIENT)
                End If


                sii = getSYFlowStep.GetStepInfoItem(
                    sCurrentCaseid,
                    CInt(dr("SUBFLOW_SEQ")),
                    CStr(dr("STEP_NO")),
                    Nothing,
                    True)

                Return sii

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)

            End Try

        End Function



        ',[APV_CAS_ID]
        ',[APP_BRADEPNO]
        ',[CPL_APL_ID]
        ',[CPL_APL_NAM]

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

            Dim arrayBosParameter As New BosParamsList

            Try
                arrayBosParameter.Add("CASEID", sCaseId)

                If Not String.IsNullOrEmpty(sApvCasId) Then
                    arrayBosParameter.Add("APV_CAS_ID", sApvCasId)
                End If

                If Not String.IsNullOrEmpty(sAppBrid) Then
                    arrayBosParameter.Add("APP_BRADEPNO", getSYBranch.GetBraDepNo(sAppBrid))
                End If

                If Not String.IsNullOrEmpty(sCplAplId) Then
                    arrayBosParameter.Add("CPL_APL_ID", sCplAplId)
                End If

                If Not String.IsNullOrEmpty(sCplAplNam) Then
                    arrayBosParameter.Add("CPL_APL_NAM", sCplAplNam)
                End If

                Update(arrayBosParameter.ToArray)

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)

            End Try
        End Sub

    End Class

End Namespace
