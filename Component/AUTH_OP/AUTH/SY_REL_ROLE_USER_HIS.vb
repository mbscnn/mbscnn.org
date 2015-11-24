Imports com.Azion.NET.VB

Public Class SY_REL_ROLE_USER_HIS
    Inherits BosBase

    Sub New()
        MyBase.New()
        Me.setPrimaryKeys()
    End Sub

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.new("SY_REL_ROLE_USER_HIS", dbManager)
    End Sub

#Region "濟南昱勝添加"
#Region "Lake Function"
    ''' <summary>
    ''' 組織人員分派
    ''' 刪除人員時，若有刪除其角色，
    ''' 新增人員角色歷史檔
    ''' </summary>
    ''' <param name="sCaseId"></param>
    ''' <param name="sStepNo"></param>
    ''' <param name="sSubFlowSeq"></param>
    ''' <param name="sSubFlowCount"></param>
    ''' <param name="sStaffId"></param>
    ''' <param name="sBraDepNo"></param>
    ''' <returns></returns>
    ''' <remarks>[Lake] 2012/07/20 Created</remarks>
    Public Function insertSyRelRoleUserHis(ByVal sCaseId As String, ByVal sStepNo As String, ByVal sSubFlowSeq As String, _
                                           ByVal sSubFlowCount As String, ByVal sStaffId As String, ByVal sBraDepNo As String) As Boolean
        Try
            Dim sSql As String = "insert into sy_rel_role_user_his " & _
                "select roleid " & _
                "," & ProviderFactory.PositionPara & "CASEID " & _
                "," & ProviderFactory.PositionPara & "STEPNO " & _
                "," & ProviderFactory.PositionPara & "SUBFLOWSEQ " & _
                "," & ProviderFactory.PositionPara & "SUBFLOWCOUNT " & _
                ",'Y' " & _
                ",STAFFID,BRA_DEPNO,NULL,'D' from " & _
                "sy_rel_role_user " & _
                "where staffid = " & ProviderFactory.PositionPara & "STAFFID " & _
                "and bra_depno in(" & sBraDepNo & ")"

            Dim paras(4) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("CASEID", sCaseId)
            paras(1) = ProviderFactory.CreateDataParameter("STEPNO", sStepNo)
            paras(2) = ProviderFactory.CreateDataParameter("SUBFLOWSEQ", sSubFlowSeq)
            paras(3) = ProviderFactory.CreateDataParameter("SUBFLOWCOUNT", sSubFlowCount)
            paras(4) = ProviderFactory.CreateDataParameter("STAFFID", sStaffId)

            If Me.getDatabaseManager.isTransaction Then
                Return DBObject.ExecuteNonQuery(Me.getDatabaseManager.getTransaction, CommandType.Text, sSql, paras) > 0
            Else
                Return DBObject.ExecuteNonQuery(Me.getDatabaseManager.getConnection, CommandType.Text, sSql, paras) > 0
            End If
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function
#End Region

#Region "Zack Function"

    ''' <summary>
    ''' 根據主鍵查詢數據
    ''' </summary>
    ''' <param name="sCASEID">案號</param>
    ''' <param name="sSUBFLOW_SEQ">子流程序號</param>
    ''' <param name="sSUBFLOW_COUNT">已執行的子流程數量</param>
    ''' <param name="sSTEP_NO">流程工作步驟編號</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Zack 2012-05-15 Create
    ''' </remarks>
    Public Function laodByPK(ByVal sROLEID As String, ByVal sCASEID As String, ByVal sSTEP_NO As String _
                             , ByVal sSUBFLOW_SEQ As String, ByVal sSUBFLOW_COUNT As String, ByVal sSTAFFID As String, ByVal sBRA_DEPNO As String) As Boolean
        Try
            Dim paras(6) As IDbDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("ROLEID", sROLEID)
            paras(1) = ProviderFactory.CreateDataParameter("CASEID", sCASEID)
            paras(2) = ProviderFactory.CreateDataParameter("STEP_NO", sSTEP_NO)
            paras(3) = ProviderFactory.CreateDataParameter("SUBFLOW_SEQ", sSUBFLOW_SEQ)
            paras(4) = ProviderFactory.CreateDataParameter("SUBFLOW_COUNT", sSUBFLOW_COUNT)
            paras(5) = ProviderFactory.CreateDataParameter("STAFFID", sSTAFFID)
            paras(6) = ProviderFactory.CreateDataParameter("BRA_DEPNO", sBRA_DEPNO)

            Return Me.loadBySQL(paras)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' 根據角色編號，人員編號，案件編號查詢資料
    ''' </summary>
    ''' <param name="sStepNO">步驟號</param>
    ''' <param name="sCaseID">案件編號</param>
    ''' <param name="sROLEID">角色編號</param>
    ''' <param name="sSTAFFID">人員編號</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Zack 2012-05-18 Create
    ''' </remarks>
    Public Function loadByRoleIDStaffid(ByVal sStepNO As String, ByVal sCaseID As String, ByVal sROLEID As String, ByVal sSTAFFID As String) As Boolean
        Try
            Dim sSQL As String = "SELECT * FROM SY_REL_ROLE_user_HIS " & _
                                " JOIN SY_ROLE ON SY_REL_ROLE_user_HIS.ROLEID = SY_ROLE.ROLEID" & _
                                " WHERE SY_REL_ROLE_user_HIS.ROLEID IN " & _
                                " (" & sROLEID & ")" & _
                                " AND SY_REL_ROLE_user_HIS.STAFFID in " & _
                                " (" & ProviderFactory.PositionPara & "STAFFID)" & _
                                " AND SY_REL_ROLE_user_HIS.CASEID IN " & _
                                " (SELECT DISTINCT caseid FROM SY_FLOWSTEP WHERE STATUS != '3')"

            If sStepNO <> "" Then
                sSQL = sSQL + " AND CASEID !=" & ProviderFactory.PositionPara & "CASEID"

                Dim paras(1) As IDbDataParameter

                paras(0) = ProviderFactory.CreateDataParameter("STAFFID", sSTAFFID)
                paras(1) = ProviderFactory.CreateDataParameter("CASEID", sCaseID)

                Return Me.loadBySQL(sSQL, paras)
            Else
                Dim paras(0) As IDbDataParameter

                paras(0) = ProviderFactory.CreateDataParameter("STAFFID", sSTAFFID)

                Return Me.loadBySQL(sSQL, paras)
            End If
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' 根據案件編號刪除歷史檔信息
    ''' </summary>
    ''' <param name="sCASEID"></param>
    ''' <remarks>
    ''' Zack 2012-06-04 Create
    ''' </remarks>
    Public Sub deleteByCaseID(ByVal sCASEID As String)
        Try
            Dim sSQL As String = "DELETE FROM SY_REL_ROLE_USER_HIS WHERE CASEID=" & ProviderFactory.PositionPara & "CASEID"

            Dim paras(0) As IDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("CASEID", sCASEID)

            If Me.getDatabaseManager.isTransaction Then
                DBObject.ExecuteNonQuery(Me.getDatabaseManager.getTransaction, CommandType.Text, sSQL, paras)
            Else
                DBObject.ExecuteNonQuery(Me.getDatabaseManager.getConnection, CommandType.Text, sSQL, paras)
            End If
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' 根據主鍵取得最大的SUBFLOW_SEQ值
    ''' </summary>
    ''' <param name="sRoleID">角色編號</param>
    ''' <param name="sCaseID">案件號</param>
    ''' <param name="sStepNO">步驟號</param>
    ''' <param name="sSubFlowCount"></param>
    ''' <param name="sStaffid">人員編號</param>
    ''' <param name="sBraDepno">部門編號</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getMaxSubFlowSeq(ByVal sRoleID As String, ByVal sCaseID As String, ByVal sStepNO As String _
                                     , ByVal sSubFlowCount As String, ByVal sStaffid As String, ByVal sBraDepno As String) As Integer
        Try
            Dim syRelRoleUserHis As New SY_REL_ROLE_USER_HIS(Me.getDatabaseManager)

            Dim sSQL As String = "SELECT MAX(SUBFLOW_SEQ) AS SUBFLOW_SEQ FROM SY_REL_ROLE_USER_HIS " & _
                                " WHERE ROLEID =" & ProviderFactory.PositionPara & "ROLEID" & _
                                " AND CASEID =" & ProviderFactory.PositionPara & "CASEID" & _
                                " AND STEP_NO =" & ProviderFactory.PositionPara & "STEP_NO" & _
                                " AND SUBFLOW_COUNT =" & ProviderFactory.PositionPara & "SUBFLOW_COUNT" & _
                                " AND STAFFID =" & ProviderFactory.PositionPara & "STAFFID" & _
                                " AND BRA_DEPNO =" & ProviderFactory.PositionPara & "BRA_DEPNO"

            Dim para(5) As IDbDataParameter
            para(0) = ProviderFactory.CreateDataParameter("ROLEID", sRoleID)
            para(1) = ProviderFactory.CreateDataParameter("CASEID", sCaseID)
            para(2) = ProviderFactory.CreateDataParameter("STEP_NO", sStepNO)
            para(3) = ProviderFactory.CreateDataParameter("SUBFLOW_COUNT", sSubFlowCount)
            para(4) = ProviderFactory.CreateDataParameter("STAFFID", sStaffid)
            para(5) = ProviderFactory.CreateDataParameter("BRA_DEPNO", sBraDepno)


            If (syRelRoleUserHis.loadBySQL(sSQL, para)) Then

                ' 如果“Roleid”不為Nothing
                If Not syRelRoleUserHis.getAttribute("SUBFLOW_SEQ") Is Nothing Then
                    If Convert.ToInt32(syRelRoleUserHis.getAttribute("SUBFLOW_SEQ").ToString()) > 0 Then
                        Return Convert.ToInt32(syRelRoleUserHis.getAttribute("SUBFLOW_SEQ").ToString())
                    Else
                        Return 0
                    End If
                Else
                    Return 0
                End If
            End If

            Return 0
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function
#End Region
#End Region
End Class
