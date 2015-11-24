Option Explicit On
Option Strict On

Imports com.Azion.NET.VB
Imports System.Text

Namespace TABLE

    Public Class SY_FLOW_MAP
        Inherits SY_TABLEBASE


        ''' <summary>
        ''' Constructor
        ''' </summary>
        ''' <param name="dbManager">DatabaseManager</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal dbManager As com.Azion.NET.VB.DatabaseManager)
            MyBase.new("SY_FLOW_MAP", dbManager)
        End Sub


        ''' <summary>
        ''' 取得FLOW的第一個STEP_NO
        ''' </summary>
        ''' <param name="flowName"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        ''' 
        Public Function GetFirstStepNo(ByVal flowName As String) As String
            Try
                Return GetFirstStepNo(PARAMETER("FLOW_NAME", flowName))
            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try
        End Function


        ''' <summary>
        ''' 取得流程的第一個步驟
        ''' </summary>
        ''' <param name="bosParameter">可以使用FLOW_ID或FLOW_NAME做為過濾條件，
        ''' 例如，PARAMETER("FLOW_ID", 1))或PARAMETER("FLOW_NAME", "FLOW_NAME")) </param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetFirstStepNo(ByVal bosParameter As BosParameter) As String
            Dim dataRow As DataRow
            Dim sField As String = bosParameter.parameterName

            Try
                '是否有"START"，有的話傳回START
                '如果沒有"START"，傳回最小值(但不包括END)
                dataRow = GetDataRow( _
                    "select top 1 ALLSTEP.STEP_NO, ALLSTEP.FLOW_ID, ALLSTEP.SPLITER_TYPE" & vbCrLf & _
                    "  from (" & _
                    "        select SFM.STEP_NO, SFM.FLOW_ID, SFM.SPLITER_TYPE, 1 AS R_O_W" & vbCrLf & _
                    "          from SY_FLOW_ID SFI" & vbCrLf & _
                    "         inner join SY_FLOW_MAP SFM" & vbCrLf & _
                    "            on SFI.FLOW_ID = SFM.FLOW_ID" & vbCrLf & _
                    "         where SFI." & sField & " = @" & sField & "@" & vbCrLf & _
                    "           and SFM.STEP_NO = @FLOW_START@" & vbCrLf & _
                    "        union" & vbCrLf & _
                    "        select min(SFM.STEP_NO), SFM.FLOW_ID, SFM.SPLITER_TYPE, 2 AS R_O_W" & vbCrLf & _
                    "          from SY_FLOW_ID SFI" & vbCrLf & _
                    "         inner join SY_FLOW_MAP SFM" & vbCrLf & _
                    "            on SFI.FLOW_ID = SFM.FLOW_ID" & vbCrLf & _
                    "         where SFI." & sField & " = @" & sField & "@" & vbCrLf & _
                    "           and SFM.STEP_NO <> @FLOW_START@" & vbCrLf & _
                    "           and SFM.STEP_NO <> @FLOW_END@" & vbCrLf & _
                    "         group by SFM.FLOW_ID, SPLITER_TYPE) ALLSTEP" & vbCrLf & _
                    "         order by R_O_W",
                    bosParameter.parameterName, bosParameter.parameterValue,
                    "FLOW_START", SY_STEP_NO.SystemStep_Start,
                    "FLOW_END", SY_STEP_NO.SystemStep_End)

                '如果找不到最先的步驟
                If IsNothing(dataRow) Then
                    Throw New SYException("無法取得SY_FLOW_MAP [" & bosParameter.parameterName & " = " & bosParameter.parameterValue.ToString & "]的第一個STEP_NO",
                                          SYMSG.SY_FLOWMAP_FIRSTSTEPNO_NOT_FOUND,
                                          GetLastSQL)
                End If

                '如果最先的步驟是'START'，去SY_FLOWDEF找下一步驟
                'If dataRow("STEP_NO").ToString().ToUpper().CompareTo("START") = 0 Then
                '    Dim drc As DataRowCollection
                '    drc = getSYFlowDef().GetNextStep(CInt(dataRow("FLOW_ID")), CStr(dataRow("STEP_NO")), "N")

                '    If IsNothing(drc) OrElse drc.Count <> 1 Then
                '        Throw New Exception("無法取得 [" & flowName & "]的第一個STEP_NO")
                '    End If

                '    Return CDbType(Of String)(drc(0).Item("NEXT_STEP_NO"))
                'End If

                Return CDbType(Of String)(dataRow("STEP_NO"))

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Function



        Public Function GetStepType(ByVal nFlowId As Integer, ByVal sStepNo As String) As String

            Dim dr As DataRow

            Try
                dr = GetStepInfo(nFlowId, sStepNo)

                If IsNothing(dr) Then
                    Throw New SYException(String.Format("流程步驟的未定義：FLOW_ID={0}，STEP_NO={1} ", nFlowId, sStepNo),
                                          SYMSG.SY_FLOWMAP_RECORD_NOT_FOUND)
                End If

                If IsDBNull(dr("TYPE")) Then
                    Return ""
                End If

                Return CStr(dr("TYPE"))

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Function


        Public Function GetStepInfo(ByVal nFlowId As Integer, ByVal sStepNo As String) As DataRow

            Try
                Return GetDataRow(Nothing,
                                  "FLOW_ID", nFlowId,
                                  "STEP_NO", sStepNo)
            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Function



        ''' <summary>
        ''' 取得下一步驟的使用者
        ''' </summary>
        ''' <param name="nFlowId"></param>
        ''' <param name="sCurrentStepNo"></param>
        ''' <param name="sNextStepNo"></param>
        ''' <param name="nBraDepNo"></param>
        ''' <param name="sCurrentUser"></param>
        ''' <param name="bGetBraDepNoOnly">只取下一步驟的使用者所在的分行代碼</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetNextStepUser(ByVal nFlowId As Integer,
                                        ByVal sCurrentStepNo As String,
                                        ByVal sNextStepNo As String,
                                        ByVal nBraDepNo As Integer,
                                        ByVal sCurrentUser As String,
                                        Optional ByVal bGetBraDepNoOnly As Boolean = False) As DataRowCollection
            Dim sSql As String = Nothing
            Dim drc As DataRowCollection
            Dim nAllChilds() As Integer
            Dim sbAllChilds As New StringBuilder
            Dim sbParamStr As New StringBuilder
            Dim paramList As New BosParamsList
            Dim nIndex As Integer = 0

            Try

                If nBraDepNo <> 0 Then
                    nAllChilds = getSYBranch.GetChildFromParent(nBraDepNo)

                    For Each nChild As Integer In nAllChilds
                        nIndex = nIndex + 1

                        If sbParamStr.Length <> 0 Then
                            sbParamStr.Append(", @BRADEPNO" & nIndex.ToString & "@")
                        Else
                            sbParamStr.Append("@BRADEPNO" & nIndex.ToString & "@")
                        End If

                        paramList.Add("BRADEPNO" & nIndex.ToString, nChild)
                    Next

                End If

                sSql = _
                        "with RC(ROLEID,ROLENAME,DISABLED,ROLETYPE,PARENT,STAFFID) as " & vbCrLf & _
                        " (select SY_ROLE.ROLEID,SY_ROLE.ROLENAME,SY_ROLE.DISABLED,SY_ROLE.ROLETYPE,SY_ROLE.PARENT, RU.STAFFID " & vbCrLf & _
                        "    from SY_ROLE " & vbCrLf & _
                        "   inner join SY_REL_ROLE_USER RU " & vbCrLf & _
                        "      on RU.ROLEID = SY_ROLE.ROLEID " & vbCrLf & _
                        "   where SY_ROLE.DISABLED = '0' " & vbCrLf & _
                        "  union all " & vbCrLf & _
                        "  select a.ROLEID,a.ROLENAME,a.DISABLED,a.ROLETYPE,a.PARENT, RC.STAFFID " & vbCrLf & _
                        "    from SY_ROLE as a, RC " & vbCrLf & _
                        "   where a.PARENT = RC.ROLEID " & vbCrLf & _
                        "     and RC.DISABLED = '0' " & vbCrLf & _
                        "     and a.DISABLED = '0') " & vbCrLf

                If bGetBraDepNoOnly = False Then
                    sSql &= _
                            "select distinct RU.STAFFID, UR.USERNAME from RC " & vbCrLf
                Else
                    sSql &= _
                            "select distinct BRA_DEPNO from RC " & vbCrLf
                End If

                sSql &= _
                        " inner join SY_REL_ROLE_USER RU " & vbCrLf & _
                        "    on RU.ROLEID = RC.ROLEID " & vbCrLf & _
                        "   and RU.STAFFID = RC.STAFFID " & vbCrLf & _
                        " inner join SY_USER UR " & vbCrLf & _
                        "    on RU.STAFFID = UR.STAFFID " & vbCrLf & _
                        " inner join SY_REL_ROLE_FLOWMAP RF " & vbCrLf & _
                        "    on RU.ROLEID = RF.ROLEID " & vbCrLf & _
                        " inner join SY_FLOW_DEF FD " & vbCrLf & _
                        "    on RF.FLOW_ID = FD.FLOW_ID " & vbCrLf & _
                        "   and RF.STEP_NO = FD.NEXT_STEP_NO " & vbCrLf & _
                        " inner join SY_FLOW_MAP FM " & vbCrLf & _
                        "    on FM.FLOW_ID = FD.FLOW_ID " & vbCrLf & _
                        "   and FM.STEP_NO = FD.CURR_STEP_NO " & vbCrLf & _
                        " inner join SY_FLOW_MAP FM2 " & vbCrLf & _
                        "    on FM2.FLOW_ID = FD.FLOW_ID " & vbCrLf & _
                        "   and FM2.STEP_NO = FD.NEXT_STEP_NO " & vbCrLf & _
                        " where FD.FLOW_ID = @FLOW_ID@ " & vbCrLf & _
                        "   and UR.STATUS = 0 " & vbCrLf & _
                        "   and FD.CURR_STEP_NO = @CURR_STEP_NO@ " & vbCrLf & _
                        "   and FD.NEXT_STEP_NO = @NEXT_STEP_NO@ " & vbCrLf & _
                        "   and (FD.PROHIBITE_SAME is null or FD.PROHIBITE_SAME<>'N' or (FD.PROHIBITE_SAME = 'Y' and RU.STAFFID <>  @STAFFID@ and FM.TYPE is null and FM2.TYPE is NULL)) "

                If nBraDepNo <> 0 Then
                    sSql &= _
                        "   and RU.BRA_DEPNO in ( " & sbParamStr.ToString & " )" & vbCrLf
                End If



                paramList.Add("FLOW_ID", nFlowId)
                paramList.Add("CURR_STEP_NO", sCurrentStepNo)
                paramList.Add("NEXT_STEP_NO", sNextStepNo)
                paramList.Add("STAFFID", sCurrentUser)

                drc = GetDataRowCollection(sSql, paramList.ToArray)

                Return drc

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try
        End Function


        ''' <summary>
        ''' 取得Splitter相對應的Joiner
        ''' </summary>
        ''' <param name="nFlowId"></param>
        ''' <param name="sSplitterStepNo"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetJoinnerBySplitter(ByVal nFlowId As Integer, ByVal sSplitterStepNo As String) As String

            Dim objSplitter As Object

            Try
                objSplitter = ExecuteScalar( _
                                "select JOINER_STEPNO " & vbCrLf & _
                                "  from SY_FLOW_MAP " & vbCrLf & _
                                " where FLOW_ID = @FLOW_ID@ " & vbCrLf & _
                                "   and STEP_NO = @STEP_NO@ " & vbCrLf,
                                "FLOW_ID", nFlowId,
                                "STEP_NO", sSplitterStepNo)

                Return CDbType(Of String)(objSplitter, Nothing)
            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try
        End Function


    End Class

End Namespace

