Imports com.Azion.NET.VB
Imports AUTH_OP.TABLE
Imports System.Text

Public Class SY_USERROLEAGENTList
    : Inherits BosList

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.New("SY_USERROLEAGENT", dbManager)
    End Sub

    Overrides Function newBos() As BosBase
        Dim bos As SY_USERROLEAGENT = New SY_USERROLEAGENT(MyBase.getDatabaseManager)
        Return bos
    End Function

#Region "濟南昱勝添加"
#Region "Lake Function"
    ''' <summary>
    ''' 代理人設定功能
    ''' 根據“分行單位代碼”查詢代理信息
    ''' </summary>
    ''' <param name="sBrid">分行單位代碼</param>
    ''' <param name="sParent">上級部門代碼</param>
    ''' <param name="sHoFlag">屬性</param>
    ''' <returns></returns>
    ''' <remarks>[Lake] 2012/05/10 Created</remarks>
    Public Function getAgentInfo(ByVal sBrid As String, ByVal sParent As String, ByVal sHoFlag As String) As Boolean
        Try
            Dim sSql As String = "Select " & _
                                 "    SY_USERROLEAGENT.STAFFID " & _
                                 "   ,SY_USERROLEAGENT.BRA_DEPNO AS BRADEPNO " & _
                                 "   ,SY_USERROLEAGENT.AGENT_BRADEPNO AS AGENTBRADEPNO " & _
                                 "   ,SY_USERROLEAGENT.CREATETIME " & _
                                 "   ,SY_USER.USERNAME " & _
                                 "   ,SY_USERROLEAGENT.AGENT_STAFFID AS AGENTSTAFFID " & _
                                 "   ,Agent.USERNAME AS AGENTUSERNAME " & _
                                 "   ,ANGENTBRANCH.BRCNAME " & _
                                 "   ,SY_USERROLEAGENT.STARTTIME AS STARTDATETIME " & _
                                 "   ,SY_USERROLEAGENT.ENDTIME AS ENDDATETIME " & _
                                 "From " & _
                                 "   SY_USERROLEAGENT " & _
                                 "   JOIN SY_USER ON SY_USERROLEAGENT.STAFFID = SY_USER.STAFFID AND SY_USER.STATUS = '0' " & _
                                 "   JOIN SY_USER AS Agent ON SY_USERROLEAGENT.AGENT_STAFFID = Agent.STAFFID AND agent.STATUS = '0' " & _
                                 "   JOIN SY_BRANCH ON SY_USERROLEAGENT.BRA_DEPNO = SY_BRANCH.BRA_DEPNO " & _
                                 "   JOIN SY_BRANCH ANGENTBRANCH ON SY_USERROLEAGENT.AGENT_BRADEPNO = ANGENTBRANCH.BRA_DEPNO " & _
                                 "Where " & _
                                 "   CANCELTIME IS NULL "

            Dim paras(-1) As IDbDataParameter

            If sHoFlag <> "1" Then
                sSql = sSql & "AND SY_BRANCH.BRID= " & ProviderFactory.PositionPara & "BRID " & _
                    "AND SY_BRANCH.PARENT = " & ProviderFactory.PositionPara & "PARENT"

                ReDim Preserve paras(1)
                paras(0) = ProviderFactory.CreateDataParameter("BRID", sBrid)
                paras(1) = ProviderFactory.CreateDataParameter("PARENT", sParent)
            End If

            sSql = sSql & " ORDER BY CREATETIME "

            Return Me.loadBySQL(sSql, paras) > 0
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 代理人查詢功能
    ''' 查詢目前的代理情況
    ''' </summary>
    ''' <param name="sBraDepNo">被代理人單位代碼</param>
    ''' <param name="sFlag">"C":目前代理情況；"H":歷史代理情況</param>
    ''' <param name="sAgentBraDepNo">代理人單位代碼</param>
    ''' <param name="sStaffId">被代理人員編號</param>
    ''' <param name="sAgentStaffId">代理人員編號</param>
    ''' <param name="sStartDateTime">開始時間</param>
    ''' <param name="sEndDateTIme">結束時間</param>
    ''' <returns></returns>
    ''' <remarks>[Lake] 2012/06/26 Created</remarks>
    Public Function getCurrentAgentInfo(ByVal sBraDepNo As String, _
                                        Optional ByVal sFlag As String = Nothing, _
                                        Optional ByVal sAgentBraDepNo As String = Nothing, _
                                        Optional ByVal sStaffId As String = Nothing, _
                                        Optional ByVal sAgentStaffId As String = Nothing, _
                                        Optional ByVal sStartDateTime As String = Nothing, _
                                        Optional ByVal sEndDateTIme As String = Nothing) As Boolean
        Try
            Dim sb As New StringBuilder()
            sb.Append("Select " & _
                      "    SY_USERROLEAGENT.STAFFID+SY_USER.USERNAME AS STAFFID " & _
                      "   ,SY_USERROLEAGENT.AGENT_STAFFID+Agent.USERNAME AS AGENTSTAFFID " & _
                      "   ,SY_BRANCH.BRCNAME " & _
                      "   ,SY_USERROLEAGENT.STARTTIME AS STARTDATETIME" & _
                      "   ,SY_USERROLEAGENT.ENDTIME AS ENDDATETIME " & _
                      "   ,SY_USERROLEAGENT.CREATEUSER + Createuser.USERNAME AS CREATEUSER " & _
                      "   ,SY_USERROLEAGENT.CREATETIME " & _
                      "   ,SY_USERROLEAGENT.CANCELTIME " & _
                      "From " & _
                      "    SY_USERROLEAGENT LEFT OUTER JOIN SY_USER on SY_USERROLEAGENT. STAFFID = SY_USER.STAFFID " & _
                      "    LEFT OUTER JOIN SY_USER as Agent on SY_USERROLEAGENT.AGENT_STAFFID = Agent.STAFFID " & _
                      "    LEFT OUTER JOIN SY_USER as Createuser on SY_USERROLEAGENT.CREATEUSER = Createuser.STAFFID " & _
                      "    JOIN SY_BRANCH on SY_USERROLEAGENT.AGENT_BRADEPNO = SY_BRANCH.BRA_DEPNO " & _
                      "WHERE "
                      )

            'Dim paras(6) As IDbDataParameter
            Dim paras As New List(Of IDbDataParameter)

            If sFlag = "C" Then
                sb.Append("GETDATE() BETWEEN SY_USERROLEAGENT.STARTTIME AND SY_USERROLEAGENT.ENDTIME " & _
                          "AND (SY_USER.STATUS = 0 OR (GETDATE() BETWEEN SY_USERROLEAGENT.STARTTIME AND SY_USERROLEAGENT.ENDTIME AND CANCELTIME IS NULL)) " & _
                          "AND (Agent.STATUS = 0 OR (GETDATE() BETWEEN SY_USERROLEAGENT.STARTTIME AND SY_USERROLEAGENT.ENDTIME AND CANCELTIME IS NULL)) ")
            ElseIf sFlag = "H" Then
                sb.Append("(SY_USER.STATUS = 0 OR (GETDATE() BETWEEN SY_USERROLEAGENT.STARTTIME AND SY_USERROLEAGENT.ENDTIME AND CANCELTIME IS NULL)) " & _
                          "AND (Agent.STATUS = 0 OR (GETDATE() BETWEEN SY_USERROLEAGENT.STARTTIME AND SY_USERROLEAGENT.ENDTIME AND CANCELTIME IS NULL)) ")
            End If

            If sStartDateTime <> Nothing AndAlso sEndDateTIme <> Nothing Then
                sb.Append("And ((SY_USERROLEAGENT.STARTTIME BETWEEN " & ProviderFactory.PositionPara & "STARTTIME " & _
                          "AND " & ProviderFactory.PositionPara & "ENDTIME) " & _
                          "Or (SY_USERROLEAGENT.ENDTIME BETWEEN " & ProviderFactory.PositionPara & "STARTTIME " & _
                          "AND " & ProviderFactory.PositionPara & "ENDTIME) " & _
                          "Or (SY_USERROLEAGENT.STARTTIME <=" & ProviderFactory.PositionPara & "STARTTIME " & _
                          "AND SY_USERROLEAGENT.ENDTIME >=" & ProviderFactory.PositionPara & "ENDTIME)) ")

                paras.Add(ProviderFactory.CreateDataParameter("STARTTIME", sStartDateTime))
                paras.Add(ProviderFactory.CreateDataParameter("ENDTIME", sEndDateTIme))
            End If

            ' 被代理人單位
            If sBraDepNo <> "0" Then
                sb.Append("AND SY_USERROLEAGENT.BRA_DEPNO = " & ProviderFactory.PositionPara & "BRADEPNO ")
                paras.Add(ProviderFactory.CreateDataParameter("BRADEPNO", sBraDepNo))
            End If

            ' 代理人單位
            If sAgentBraDepNo <> Nothing AndAlso sAgentBraDepNo <> "0" Then
                sb.Append("AND SY_USERROLEAGENT.AGENT_BRADEPNO = " & ProviderFactory.PositionPara & "AGENTBRADEPNO ")

                paras.Add(ProviderFactory.CreateDataParameter("AGENTBRADEPNO", sAgentBraDepNo))
            End If

            If sStaffId <> Nothing Then
                sb.Append("AND SY_USERROLEAGENT.STAFFID = " & ProviderFactory.PositionPara & "STAFFID ")

                paras.Add(ProviderFactory.CreateDataParameter("STAFFID", sStaffId))
            End If

            If sAgentStaffId <> Nothing Then
                sb.Append("AND SY_USERROLEAGENT.AGENT_STAFFID = " & ProviderFactory.PositionPara & "AGENTSTAFFID ")

                paras.Add(ProviderFactory.CreateDataParameter("AGENTSTAFFID", sAgentStaffId))
            End If

            sb.Append(" ORDER BY SY_USERROLEAGENT.STAFFID, SY_USERROLEAGENT.BRA_DEPNO")

            Return Me.loadBySQL(sb.ToString(), paras.ToArray)
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 查詢待刪除人員是否已存在代理關係
    ''' </summary>
    ''' <param name="sStaffId">待刪除人員</param>
    ''' <param name="sBraDepNo">當前點選部門</param>
    ''' <returns></returns>
    ''' <remarks>[Lake] 2012/06/20 Created</remarks>
    Public Function getToDeleteStaff(ByVal sBraDepNo As String, ByVal sStaffId As String) As Boolean
        Try
            Dim sSql As String = "SELECT * FROM SY_USERROLEAGENT " & _
                    "WHERE STARTTIME <= GETDATE() AND ENDTIME >=GETDATE() AND CANCELTIME IS NULL " & _
                    "AND (BRA_DEPNO =" & ProviderFactory.PositionPara & "BRADEPNO AND STAFFID IN ('" & sStaffId & "') " & _
                    "OR AGENT_BRADEPNO =" & ProviderFactory.PositionPara & "BRADEPNO AND AGENT_STAFFID IN ('" & sStaffId & "'))"

            Dim paras As IDbDataParameter = ProviderFactory.CreateDataParameter("BRADEPNO", sBraDepNo)

            Return Me.loadBySQL(sSql, paras) > 0
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

#Region "Avril Function"

    ''' <summary>
    ''' 根據登入者查詢代理人資料
    ''' </summary>
    ''' <param name="sStaffId">登入者</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Add by Avril 2012/06/11
    ''' </remarks>
    Function loadAgentData(ByVal sStaffId As String) As Boolean
        Try
            '  BRCNAME
            Dim sSql As String = "with rc(bra_depno,brid,parentid,brcname,brename," & _
                                    "hoflag,disabled,serialnum,serialstr,level) as" & _
                                    " (select bra_depno,brid,parent," & _
                                    "         brcname,brename,hoflag,disabled," & _
                                    "         cast(row_number() over(order by bra_depno) as int) serialnum," & _
                                    "         cast(row_number() over(order by bra_depno) as varchar(255)) serialstr," & _
                                    "         0 level" & _
                                    "    from sy_branch a" & _
                                    "   where parent = 0" & _
                                    "  union all" & _
                                    "  select p.bra_depno,p.brid,p.parent," & _
                                    "         p.brcname,p.brename,p.hoflag,p.disabled," & _
                                    "         cast(rc.serialnum + 1 as int)," & _
                                    "         cast(rtrim(rc.serialstr) + '-' +" & _
                                    "              cast(row_number() over(order by p.bra_depno) as varchar(10)) as" & _
                                    "              varchar(255))," & _
                                    "         rc.level + 1" & _
                                    "    from sy_branch p, rc" & _
                                    "   where rc.bra_depno = p.parent" & _
                                    "     and p.parent > 0)" & _
                                    "select a.STAFFID AGENT_STAFFID, rc.brid, brcname = rc.brid + ' ' + '(' + cast(rc.bra_depno as varchar) + ')' + '  ' + rtrim (rc.brcname), rc.bra_depno AGENT_BRADEPNO, null STARTTIME, null ENDTIME" & _
                                    "  from rc, SY_REL_BRANCH_USER a,(" & _
                                    "select rc.brid, max(level) level" & _
                                    "  from rc, SY_REL_BRANCH_USER a" & _
                                    " where a. BRA_DEPNO = rc.BRA_DEPNO and a.STAFFID = " & ProviderFactory.PositionPara & "STAFFID and disabled = '0'" & _
                                    " group by rc.brid) b" & _
                                    " where a. BRA_DEPNO = rc.BRA_DEPNO and a.STAFFID = " & ProviderFactory.PositionPara & "STAFFID and disabled = '0' and rc.level = b.level and rc.brid = b.brid and a.STAFFID = " & ProviderFactory.PositionPara & "STAFFID and disabled = '0'" & _
                                    " UNION " & _
                                    "select a.STAFFID AGENT_STAFFID, rc.brid, brcname = rc.brid + ' ' + '(' + cast(rc.bra_depno as varchar) + ')' + '  ' + rtrim (rc.brcname), rc.bra_depno, null STARTTIME, null ENDTIME" & _
                                    "  from rc, SY_REL_BRANCH_USER a,(" & _
                                    "select rc.brid, max(level) level" & _
                                    "  from rc, SY_REL_BRANCH_USER a" & _
                                    " where a. BRA_DEPNO = rc.BRA_DEPNO and a.STAFFID = " & ProviderFactory.PositionPara & "STAFFID and disabled = '0'" & _
                                    " group by rc.brid) b,(" & _
                                    " Select SY_USERROLEAGENT.AGENT_STAFFID, SY_USERROLEAGENT.AGENT_BRADEPNO, SY_USERROLEAGENT. STARTTIME, SY_USERROLEAGENT.ENDTIME" & _
                                    "  FROM SY_USERROLEAGENT" & _
                                    "  JOIN SY_USER" & _
                                    "    ON SY_USERROLEAGENT.AGENT_STAFFID = SY_USER.STAFFID" & _
                                    " WHERE SY_USERROLEAGENT.CANCELTIME IS NULL AND SY_USER.status = 0 AND" & _
                                    " (SY_USERROLEAGENT.STARTTIME <= GETDATE()" & _
                                    " AND SY_USERROLEAGENT.ENDTIME >= GETDATE()" & _
                                    " OR ( SY_USERROLEAGENT.STARTTIME <= GETDATE()" & _
                                    "  AND SY_USERROLEAGENT.ENDTIME IS NULL))" & _
                                    "  AND SY_USERROLEAGENT.STAFFID = " & ProviderFactory.PositionPara & "STAFFID) c" & _
                                    " where a. BRA_DEPNO = rc.BRA_DEPNO and a.STAFFID = " & ProviderFactory.PositionPara & "STAFFID" & _
                                    " and disabled = '0'" & _
                                    " and rc.level = b.level" & _
                                    " and rc.brid = b.brid" & _
                                    " and a.BRA_DEPNO = c.AGENT_BRADEPNO" & _
                                    " and a.STAFFID = " & ProviderFactory.PositionPara & "STAFFID" & _
                                    " and disabled = '0'"



            Dim paras(0) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("STAFFID", sStaffId)

            Return Me.loadBySQL(sSql, paras)
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 根據登入者信息查詢其代理人信息
    ''' </summary>
    ''' <param name="sStaffId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function loadDataByCon(ByVal sStaffId As String) As Boolean
        Try
            Dim sSql As String = "SELECT SY_BRANCH.BRID LoginBrId," & _
                                    "       SY_BRANCH.BRCNAME LoginBrName," & _
                                    "       sy_rel_role_user.staffid LoginStaffid," & _
                                    "       sy_user.USERNAME LoginUserName," & _
                                    "       sy_rel_role_user.BRA_DEPNO LoginDepNo," & _
                                    "       '' as WorkingBrid," & _
                                    "       '' as WorkingBrName," & _
                                    "       '' as WorkingStaffid," & _
                                    "       '' as WorkingUserName," & _
                                    "       '' as WorkingDepNo," & _
                                    "       null as STARTTIME," & _
                                    "       null as ENDTIME" & _
                                    "  FROM sy_branch" & _
                                    "  join sy_rel_role_user" & _
                                    "    on sy_branch.bra_depno = sy_rel_role_user.bra_depno" & _
                                    "  left join sy_user" & _
                                    "    on sy_rel_role_user.staffid = sy_user.staffid" & _
                                    " where sy_branch.parent = 0" & _
                                    "   and sy_branch.disabled = 0" & _
                                    "   and sy_rel_role_user.staffid = " & ProviderFactory.PositionPara & "STAFFID " & _
                                    " union " & _
                                    "Select BRANCHAGENT.BRID                LoginBrId," & _
                                    "       BRANCHAGENT.BRCNAME             LoginBrName," & _
                                    "       SY_USERROLEAGENT.AGENT_STAFFID  LoginStaffid," & _
                                    "       USERAGENT.USERNAME              LoginUserName," & _
                                    "       SY_USERROLEAGENT.AGENT_BRADEPNO LoginDepNo," & _
                                    "       SY_BRANCH.BRID                  as WorkingBrid," & _
                                    "       SY_BRANCH.BRCNAME               as WorkingBrName," & _
                                    "       SY_USERROLEAGENT.STAFFID        as WorkingStaffid," & _
                                    "       SY_USER.USERNAME                as WorkingUserName," & _
                                    "       SY_USERROLEAGENT.BRA_DEPNO      as WorkingDepNo," & _
                                    "       SY_USERROLEAGENT.STARTTIME      as STARTTIME," & _
                                    "       SY_USERROLEAGENT.ENDTIME        as ENDTIME" & _
                                    "  FROM SY_USERROLEAGENT" & _
                                    "  JOIN SY_USER USERAGENT" & _
                                    "    ON SY_USERROLEAGENT.AGENT_STAFFID = USERAGENT.STAFFID" & _
                                    "  JOIN SY_BRANCH" & _
                                    "    ON SY_BRANCH.BRA_DEPNO = SY_USERROLEAGENT.BRA_DEPNO" & _
                                    "  JOIN SY_BRANCH BRANCHAGENT" & _
                                    "    ON BRANCHAGENT.BRA_DEPNO = SY_USERROLEAGENT.AGENT_BRADEPNO" & _
                                    "  JOIN SY_USER" & _
                                    "    ON SY_USERROLEAGENT.STAFFID = SY_USER.STAFFID" & _
                                    " WHERE SY_USERROLEAGENT.CANCELTIME IS NULL" & _
                                    "   AND USERAGENT.STATUS = 0" & _
                                    "   AND SY_USER.status = 0" & _
                                    "   AND SY_BRANCH.DISABLED = 0" & _
                                    "   AND BRANCHAGENT.DISABLED = 0" & _
                                    "   AND (SY_USERROLEAGENT.STARTTIME <= GETDATE() AND" & _
                                    "       SY_USERROLEAGENT.ENDTIME >= GETDATE() OR" & _
                                    "       (SY_USERROLEAGENT.STARTTIME <= GETDATE() AND" & _
                                    "       SY_USERROLEAGENT.ENDTIME IS NULL))" & _
                                    "   AND SY_USERROLEAGENT.AGENT_STAFFID = " & ProviderFactory.PositionPara & "STAFFID"


            Dim paras(0) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("STAFFID", sStaffId)


            Return Me.loadBySQL(sSql, paras)
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region
#End Region
End Class
