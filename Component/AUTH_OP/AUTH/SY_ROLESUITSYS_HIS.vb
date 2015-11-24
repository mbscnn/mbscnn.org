Imports com.Azion.NET.VB

Public Class SY_ROLESUITSYS_HIS
    Inherits BosBase

    Public Sub New(ByVal dbManager As com.Azion.NET.VB.DatabaseManager)
        MyBase.new("SY_ROLESUITSYS_HIS", dbManager)
    End Sub

    Sub New()
        MyBase.New()
        Me.setPrimaryKeys()
    End Sub

#Region "濟南昱勝添加"
#Region "Avril Function"
    ''' <summary>
    ''' 根據主鍵查詢資料
    ''' </summary>
    ''' <param name="sRoleId">角色編號</param>
    ''' <param name="sCaseId">案號</param>
    ''' <param name="sStepNo">步驟號</param>
    ''' <param name="iSubFlowSeq">流程序號</param>
    ''' <param name="iSubFlowCount">資料筆數</param>
    ''' <param name="sSubSysId">子系統編號</param>
    ''' <param name="sSysId">系統編號</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Add by Avril 2012/05/09
    ''' </remarks>
    Function loadByPK(ByVal sRoleId As String, ByVal sCaseId As String, ByVal sStepNo As String, ByVal iSubFlowSeq As Integer, ByVal iSubFlowCount As Integer, ByVal sSubSysId As String, ByVal sSysId As String) As Boolean
        Try
            Dim para(6) As IDbDataParameter
            para(0) = ProviderFactory.CreateDataParameter("ROLEID", sRoleId)
            para(1) = ProviderFactory.CreateDataParameter("CASEID", sCaseId)
            para(2) = ProviderFactory.CreateDataParameter("STEP_NO", sStepNo)
            para(3) = ProviderFactory.CreateDataParameter("SUBFLOW_SEQ", iSubFlowSeq)
            para(4) = ProviderFactory.CreateDataParameter("SUBFLOW_COUNT", iSubFlowCount)
            para(5) = ProviderFactory.CreateDataParameter("SUBSYSID", sSubSysId)
            para(6) = ProviderFactory.CreateDataParameter("SYSID", sSysId)

            Return MyBase.loadBySQL(para)
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 根據CASEID刪除資料
    ''' </summary>
    ''' <param name="sCaseId">案號</param>
    ''' <remarks>
    ''' Add by Avril 2012/04/26
    ''' </remarks>
    Sub delHisDataByCaseId(ByVal sCaseId As String)
        Try
            Dim strSql As String = "DELETE " & _
                                   "FROM " & _
                                   "   SY_ROLESUITSYS_HIS" & _
                                   " WHERE SY_ROLESUITSYS_HIS.CASEID = " & ProviderFactory.PositionPara & "CASEID"

            Dim para(0) As IDbDataParameter
            para(0) = ProviderFactory.CreateDataParameter("CASEID", sCaseId)

            If Me.getDatabaseManager.isTransaction Then
                DBObject.ExecuteNonQuery(Me.getDatabaseManager.getTransaction, CommandType.Text, strSql, para)
            Else
                DBObject.ExecuteNonQuery(Me.getDatabaseManager.getConnection, CommandType.Text, strSql, para)
            End If
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ''' <summary>
    ''' 取得最大Seq
    ''' </summary>
    ''' <param name="sRoleId">角色編號</param>
    ''' <param name="sStepNo">步驟號</param>
    ''' <param name="sCaseId">案號</param>
    ''' <param name="sSubFlowCount">資料筆數</param>
    ''' <param name="sSubSysId">子系統</param>
    ''' <param name="sSysId">系統</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Add by 2012/06/29
    ''' </remarks>
    Function getMaxSubFlowSeq(ByVal sRoleId As String, ByVal sStepNo As String, ByVal sCaseId As String, ByVal sSubFlowCount As String, ByVal sSubSysId As String, ByVal sSysId As String) As Integer
        Try
            Dim syRoleSuitSysHis As New SY_ROLESUITSYS_HIS(Me.getDatabaseManager)
            Dim sSQL As String = "SELECT MAX(SUBFLOW_SEQ) SUBFLOW_SEQ  FROM SY_ROLESUITSYS_HIS  " & _
                " WHERE ROLEID = " & ProviderFactory.PositionPara & "ROLEID " & _
                " AND   STEP_NO = " & ProviderFactory.PositionPara & "STEP_NO " & _
                " AND   CASEID = " & ProviderFactory.PositionPara & "CASEID " & _
                " AND SUBFLOW_COUNT = " & ProviderFactory.PositionPara & "SUBFLOW_COUNT " & _
                " AND SUBSYSID = " & ProviderFactory.PositionPara & "SUBSYSID " & _
                " AND SYSID = " & ProviderFactory.PositionPara & "SYSID "

            Dim para(5) As IDbDataParameter
            para(0) = ProviderFactory.CreateDataParameter("ROLEID", sRoleId)
            para(1) = ProviderFactory.CreateDataParameter("STEP_NO", sStepNo)
            para(2) = ProviderFactory.CreateDataParameter("CASEID", sCaseId)
            para(3) = ProviderFactory.CreateDataParameter("SUBFLOW_COUNT", sSubFlowCount)
            para(4) = ProviderFactory.CreateDataParameter("SUBSYSID", sSubSysId)
            para(5) = ProviderFactory.CreateDataParameter("SYSID", sSysId)


            If (syRoleSuitSysHis.loadBySQL(sSQL, para)) Then

                ' 如果“Roleid”不為Nothing
                If Not syRoleSuitSysHis.getAttribute("SUBFLOW_SEQ") Is Nothing Then
                    If Convert.ToInt32(syRoleSuitSysHis.getAttribute("SUBFLOW_SEQ").ToString()) > 0 Then
                        Return Convert.ToInt32(syRoleSuitSysHis.getAttribute("SUBFLOW_SEQ").ToString())
                    Else
                        Return 0
                    End If
                Else
                    Return 0
                End If
            End If

            Return 0
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 查詢刪除的歷史檔資料
    ''' </summary>
    ''' <param name="sRoleId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function checkOperation(ByVal sRoleId As String) As Boolean
        Try
            Dim syRoleSuitSysHis As New SY_ROLESUITSYS_HIS(Me.getDatabaseManager)
            Dim sSQL As String = "SELECT COUNT(*)   FROM SY_ROLESUITSYS_HIS  " & _
                " WHERE ROLEID = " & ProviderFactory.PositionPara & "ROLEID " & _
                " AND OPERATION = 'D' "

            Dim paras As IDbDataParameter = ProviderFactory.CreateDataParameter("ROLEID", sRoleId)
            Dim oRowCount As Object = com.Azion.NET.VB.DBObject.ExecuteScalar(Me.getDatabaseManager().getConnection(), CommandType.Text, sSQL, paras)

            Return Convert.ToInt32(oRowCount) > 0
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
