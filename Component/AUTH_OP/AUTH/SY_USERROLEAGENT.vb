Option Explicit On
Option Strict On

Imports com.Azion.NET.VB
Imports System.Text

Namespace TABLE

    Public Class SY_USERROLEAGENT
        Inherits BosBase

        Sub New()
            MyBase.New()
            Me.setPrimaryKeys()
        End Sub

        ''' <summary>
        ''' Constructor
        ''' </summary>
        ''' <param name="dbManager">DatabaseManager</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal dbManager As com.Azion.NET.VB.DatabaseManager)
            MyBase.new("SY_USERROLEAGENT", dbManager)
        End Sub

#Region "濟南昱勝添加"
#Region "Lake Function"
        ''' <summary>
        ''' 根據部門及人員編號刪除代理關係
        ''' </summary>
        ''' <param name="sBraDepNo">當前部門及其下屬部門</param>
        ''' <param name="sStaffId">人員編號</param>
        ''' <returns></returns>
        ''' <remarks>[Lake] 2012/06/26 Created</remarks>
        Public Function deleteStaffByBraDepNoAndStaffId(sBraDepNo As String, sStaffId As String) As Boolean
            Try
                Dim sSql As String = "DELETE SY_USERROLEAGENT " & _
                    "WHERE STARTTIME <= GETDATE() AND ENDTIME >=GETDATE() AND CANCELTIME IS NULL " & _
                    "AND (BRA_DEPNO =" & ProviderFactory.PositionPara & "BRADEPNO  AND STAFFID = " & ProviderFactory.PositionPara & "STAFFID " & _
                    "OR AGENT_BRADEPNO =" & ProviderFactory.PositionPara & "BRADEPNO  AND AGENT_STAFFID = " & ProviderFactory.PositionPara & "STAFFID)"

                Dim paras(1) As IDbDataParameter
                paras(0) = ProviderFactory.CreateDataParameter("STAFFID", sStaffId)
                paras(1) = ProviderFactory.CreateDataParameter("BRADEPNO", sBraDepNo)

                If Me.getDatabaseManager.isTransaction Then
                    Return DBObject.ExecuteNonQuery(Me.getDatabaseManager.getTransaction, CommandType.Text, sSql, paras) > 0
                Else
                    Return DBObject.ExecuteNonQuery(Me.getDatabaseManager.getConnection, CommandType.Text, sSql, paras) > 0
                End If
            Catch ex As ProviderException
                Throw ex
            Catch ex As BosException
                Throw ex
            Catch ex As Exception
                Throw ex
            End Try
        End Function

        ''' <summary>
        ''' 根據主鍵查詢使用者角色代理檔
        ''' </summary>
        ''' <param name="sStaffId">員工編號</param>
        ''' <param name="sBraDepNo">部門代碼</param>
        ''' <param name="sAgentStaffId">代理人員工編號</param>
        ''' <param name="sAgentBraDepNo">代理人部門代碼</param>
        ''' <param name="sCreateTime">建立時間</param>
        ''' <returns></returns>
        ''' <remarks>[Lake] 2012/05/10 Created</remarks>
        Public Function loadByPK(ByVal sStaffId As String, ByVal sBraDepNo As String, ByVal sAgentStaffId As String, ByVal sAgentBraDepNo As String, ByVal sCreateTime As Date) As Boolean
            Try
                Dim paras(4) As IDbDataParameter
                paras(0) = ProviderFactory.CreateDataParameter("STAFFID", sStaffId)
                paras(1) = ProviderFactory.CreateDataParameter("BRA_DEPNO", sBraDepNo)
                paras(2) = ProviderFactory.CreateDataParameter("AGENT_STAFFID", sAgentStaffId)
                paras(3) = ProviderFactory.CreateDataParameter("AGENT_BRADEPNO", sAgentBraDepNo)
                paras(4) = ProviderFactory.CreateDataParameter("CREATETIME", sCreateTime)

                Return Me.loadData(paras)
            Catch ex As ProviderException
                Throw ex
            Catch ex As BosException
                Throw ex
            Catch ex As Exception
                Throw ex
            End Try
        End Function

        ''' <summary>
        ''' 根據代理人編號，被代理人編號以及代理起訖時間檢核代理關係
        ''' </summary>
        ''' <param name="sStaffId">被代理人編號</param>
        ''' <param name="sAgentStaffId">代理人編號</param>
        ''' <param name="tStartDateTime">代理起時間</param>
        ''' <param name="sEndDateTime">代理訖時間</param>
        ''' <param name="sBraDepNo">被代理人所在部門</param>
        ''' <param name="sAgentBraDepNo">代理人所在部門</param>
        ''' <returns></returns>
        ''' <remarks>[Lake] 2012/05/11 Cretaed</remarks>
        Public Function checkAgentRelation(ByVal sStaffId As String, ByVal sAgentStaffId As String, _
                                           ByVal tStartDateTime As String, ByVal sEndDateTime As String, _
                                           ByVal sBraDepNo As String, ByVal sAgentBraDepNo As String) As Boolean
            Try
                Dim sSql As String = "SELECT " & _
                                     "   * " & _
                                     "FROM " & _
                                     "   SY_USERROLEAGENT " & _
                                     "WHERE " & _
                                     "       STAFFID = " & ProviderFactory.PositionPara & "STAFFID " & _
                                     "   AND AGENT_STAFFID = " & ProviderFactory.PositionPara & "AGENTSTAFFID " & _
                                     "   AND CANCELTIME IS NULL " & _
                                     "   AND BRA_DEPNO = " & ProviderFactory.PositionPara & "BRADEPNO " & _
                                     "   AND AGENT_BRADEPNO = " & ProviderFactory.PositionPara & "AGENTBRADEPNO " & _
                                     "   AND ((STARTTIME BETWEEN " & ProviderFactory.PositionPara & "STARTTIME " & _
                                     "   AND " & ProviderFactory.PositionPara & "ENDTIME) " & _
                                     "   OR (ENDTIME BETWEEN " & ProviderFactory.PositionPara & "STARTTIME " & _
                                     "   AND " & ProviderFactory.PositionPara & "ENDTIME) " & _
                                     "   OR (STARTTIME <= " & ProviderFactory.PositionPara & "STARTTIME " & _
                                     "   AND ENDTIME >= " & ProviderFactory.PositionPara & "ENDTIME))"

                Dim paras(5) As IDbDataParameter
                paras(0) = ProviderFactory.CreateDataParameter("STAFFID", sStaffId)
                paras(1) = ProviderFactory.CreateDataParameter("AGENTSTAFFID", sAgentStaffId)
                paras(2) = ProviderFactory.CreateDataParameter("STARTTIME", tStartDateTime)
                paras(3) = ProviderFactory.CreateDataParameter("ENDTIME", sEndDateTime)
                paras(4) = ProviderFactory.CreateDataParameter("BRADEPNO", sBraDepNo)
                paras(5) = ProviderFactory.CreateDataParameter("AGENTBRADEPNO", sAgentBraDepNo)

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
        ''' 根據部門代碼及人員編號刪除代理關係
        ''' </summary>
        ''' <param name="sCurrStaffId"></param>
        ''' <param name="sCurrBraDepNo"></param>
        ''' <param name="sFunccode"></param>
        ''' <param name="sChangeBraDepNo"></param>
        ''' <param name="sCaseId"></param>
        ''' <returns></returns>
        ''' <remarks>[Lake] 2012/06/20 Created</remarks>
        Public Function deleteChangeStaffAgent(sCurrStaffId As String, sCurrBraDepNo As String, sFunccode As String, sChangeBraDepNo As String, sCaseId As String) As Boolean
            Try
                Dim sb As New StringBuilder()
                sb.Append("DELETE FROM SY_USERROLEAGENT WHERE (BRA_DEPNO IN('" & sChangeBraDepNo & "')")
                sb.Append("AND STAFFID IN(DECLARE @tempData XML ")
                sb.Append("SELECT @tempData = TEMPDATA FROM SY_TEMPINFO ")
                sb.Append("WHERE ")
                If sCaseId <> "" Then
                    sb.Append("CASEID = " & ProviderFactory.PositionPara & "CASEID ")
                Else
                    sb.Append("STAFFID=" & ProviderFactory.PositionPara & "STAFFID ")
                    sb.Append("AND BRA_DEPNO=" & ProviderFactory.PositionPara & "BRADEPNO ")
                    sb.Append("AND FUNCCODE=" & ProviderFactory.PositionPara & "FUNCCODE ")
                End If
                sb.Append("SELECT " & _
                "sy_rel_branch_user.staffid as STAFFID " & _
                "FROM @tempData.nodes('/SY/SY_REL_BRANCH_USER') AS temp(item) RIGHT JOIN  sy_rel_branch_user " & _
                "ON sy_rel_branch_user.BRA_DEPNO = temp.item.query('BRA_DEPNO').value('.[1]','nvarchar(20)') " & _
                "and sy_rel_branch_user.staffid = temp.item.query('STAFFID').value('.[1]','nvarchar(20)') " & _
                "and temp.item.query('OPERATION').value('.[1]','nvarchar(20)')<>'D' " & _
                "join sy_user on SY_USER. STAFFID = SY_REL_BRANCH_USER. STAFFID " & _
                "where " & _
                "temp.item.query('STAFFID').value('.[1]','nvarchar(20)') is null " & _
                "and " & _
                "sy_rel_branch_user.BRA_DEPNO in ('" & sChangeBraDepNo & "') " & _
                "and SY_USER.STATUS=0))")
                sb.Append(" OR  (AGENT_BRADEPNO IN('" & sChangeBraDepNo & "')")
                sb.Append("AND AGENT_STAFFID IN(DECLARE @tempData XML ")
                sb.Append("SELECT @tempData = TEMPDATA FROM SY_TEMPINFO ")
                sb.Append("WHERE ")
                If sCaseId <> "" Then
                    sb.Append("CASEID = " & ProviderFactory.PositionPara & "CASEID ")
                Else
                    sb.Append("STAFFID=" & ProviderFactory.PositionPara & "STAFFID ")
                    sb.Append("AND BRA_DEPNO=" & ProviderFactory.PositionPara & "BRADEPNO ")
                    sb.Append("AND FUNCCODE=" & ProviderFactory.PositionPara & "FUNCCODE ")
                End If
                sb.Append("SELECT " & _
                "sy_rel_branch_user.staffid as STAFFID " & _
                "FROM @tempData.nodes('/SY/SY_REL_BRANCH_USER') AS temp(item) RIGHT JOIN  sy_rel_branch_user " & _
                "ON sy_rel_branch_user.BRA_DEPNO = temp.item.query('BRA_DEPNO').value('.[1]','nvarchar(20)') " & _
                "and sy_rel_branch_user.staffid = temp.item.query('STAFFID').value('.[1]','nvarchar(20)') " & _
                "and temp.item.query('OPERATION').value('.[1]','nvarchar(20)')<>'D' " & _
                "join sy_user on SY_USER. STAFFID = SY_REL_BRANCH_USER. STAFFID " & _
                "where " & _
                "temp.item.query('STAFFID').value('.[1]','nvarchar(20)') is null " & _
                "and " & _
                "sy_rel_branch_user.BRA_DEPNO in ('" & sChangeBraDepNo & "') " & _
                "and SY_USER.STATUS=0))")

                Dim paras(2) As IDataParameter
                If sCaseId <> "" Then
                    paras(0) = ProviderFactory.CreateDataParameter("CASEID", sCaseId)
                Else
                    paras(0) = ProviderFactory.CreateDataParameter("STAFFID", sCurrStaffId)
                    paras(1) = ProviderFactory.CreateDataParameter("BRADEPNO", sCurrBraDepNo)
                    paras(2) = ProviderFactory.CreateDataParameter("FUNCCODE", sFunccode)
                End If

                If Me.getDatabaseManager.isTransaction Then
                    Return DBObject.ExecuteNonQuery(Me.getDatabaseManager.getTransaction, CommandType.Text, sb.ToString(), paras) > 0
                Else
                    Return DBObject.ExecuteNonQuery(Me.getDatabaseManager.getConnection, CommandType.Text, sb.ToString(), paras) > 0
                End If

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
End Namespace
