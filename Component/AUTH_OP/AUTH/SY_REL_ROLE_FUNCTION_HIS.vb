Imports com.Azion.NET.VB

Public Class SY_REL_ROLE_FUNCTION_HIS
    Inherits BosBase

    Public Sub New(ByVal dbManager As com.Azion.NET.VB.DatabaseManager)
        MyBase.new("SY_REL_ROLE_FUNCTION_HIS", dbManager)
    End Sub

    Sub New()
        MyBase.New()
        Me.setPrimaryKeys()
    End Sub

#Region "濟南昱勝添加"

    ''' <summary>
    ''' 查詢資料是否在送簽中
    ''' </summary>
    ''' <param name="sRoleIdList">角色編號集合</param>
    ''' <param name="sStepNo">步驟號</param>
    ''' <param name="sCaseId">案號</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Add by Avril 2012/06/06
    ''' </remarks>
    Function loadDataByCon(ByVal sRoleIdList As String, ByVal sStepNo As String, ByVal sCaseId As String) As Integer
        Try
            Dim sSQL As String = "SELECT  * FROM  SY_REL_ROLE_FUNCTION_HIS join SY_ROLE ON SY_REL_ROLE_FUNCTION_HIS.ROLEID = SY_ROLE.ROLEID " & _
                                 " WHERE  1=1 "

            If sRoleIdList.Length > 0 Then
                sSQL = sSQL & " AND SY_REL_ROLE_FUNCTION_HIS.ROLEID IN( " & sRoleIdList & ") "
            End If

            sSQL = sSQL & " AND CASEID IN (SELECT CASEID FROM SY_FLOWSTEP WHERE STATUS<>3)"

            If sStepNo <> "" Then
                sSQL = sSQL & " AND CASEID<> " & ProviderFactory.PositionPara & "CASEID"
            End If

            If sStepNo <> "" Then
                Dim paras(0) As System.Data.IDbDataParameter
                paras(0) = ProviderFactory.CreateDataParameter("CASEID", sCaseId)

                Return MyBase.loadBySQL(sSQL, paras)
            Else
                Return MyBase.loadBySQL(sSQL)
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
    ''' 根據主鍵查詢資料
    ''' </summary>
    ''' <param name="sFuncCode">功能編號</param>
    ''' <param name="sSubSysId">子系統</param>
    ''' <param name="sSysId">系統</param>
    ''' <param name="sCaseId">案號</param>
    ''' <param name="sStepNo">步驟號</param>
    ''' <param name="iSubFlowSeq">序號</param>
    ''' <param name="iSubFlowCount">資料筆數</param>
    ''' <param name="iRoleId">角色編號</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function loadByPK(ByVal sFuncCode As String, ByVal sSubSysId As String, ByVal sSysId As String,
                      ByVal sCaseId As String, ByVal sStepNo As String, ByVal iSubFlowSeq As Integer,
                      ByVal iSubFlowCount As Integer, ByVal iRoleId As Integer) As Boolean
        Try
            Dim para(7) As IDbDataParameter
            para(0) = ProviderFactory.CreateDataParameter("FUNCCODE", sFuncCode)
            para(1) = ProviderFactory.CreateDataParameter("SUBSYSID", sSubSysId)
            para(2) = ProviderFactory.CreateDataParameter("SYSID", sSysId)
            para(3) = ProviderFactory.CreateDataParameter("CASEID", sCaseId)
            para(4) = ProviderFactory.CreateDataParameter("STEP_NO", sStepNo)
            para(5) = ProviderFactory.CreateDataParameter("SUBFLOW_SEQ", iSubFlowSeq)
            para(6) = ProviderFactory.CreateDataParameter("SUBFLOW_COUNT", iSubFlowCount)
            para(7) = ProviderFactory.CreateDataParameter("ROLEID", iRoleId)

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
    ''' 取得最大的序號
    ''' </summary>
    ''' <param name="sFuncCode">交易編號</param>
    ''' <param name="sCaseId">案號</param>
    ''' <param name="sStepNo">步驟數</param>
    ''' <param name="sSubFlowCount">資料筆數</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Add by Avril 2012/07/02
    ''' </remarks>
    Function getMaxSubFlowSeq(ByVal sFuncCode As String, ByVal sSubSysId As String, ByVal sSysId As String, ByVal sCaseId As String, ByVal sStepNo As String, ByVal sSubFlowCount As String, ByVal sRoleId As String) As Integer
        Try
            Dim syRoleHis As New SY_ROLE_HIS(Me.getDatabaseManager)
            Dim sSQL As String = "SELECT MAX(SUBFLOW_SEQ) SUBFLOW_SEQ  FROM SY_REL_ROLE_FUNCTION_HIS  " & _
                " WHERE FUNCCODE = " & ProviderFactory.PositionPara & "FUNCCODE " & _
                " AND   SUBSYSID = " & ProviderFactory.PositionPara & "SUBSYSID " & _
                " AND   SYSID = " & ProviderFactory.PositionPara & "SYSID " & _
                " AND   CASEID = " & ProviderFactory.PositionPara & "CASEID " & _
                " AND   STEP_NO = " & ProviderFactory.PositionPara & "STEP_NO " & _
                " AND SUBFLOW_COUNT = " & ProviderFactory.PositionPara & "SUBFLOW_COUNT " & _
                " AND   ROLEID = " & ProviderFactory.PositionPara & "ROLEID "

            Dim para(6) As IDbDataParameter
            para(0) = ProviderFactory.CreateDataParameter("FUNCCODE", sFuncCode)
            para(1) = ProviderFactory.CreateDataParameter("SUBSYSID", sSubSysId)
            para(2) = ProviderFactory.CreateDataParameter("SYSID", sSysId)
            para(3) = ProviderFactory.CreateDataParameter("CASEID", sCaseId)
            para(4) = ProviderFactory.CreateDataParameter("STEP_NO", sStepNo)
            para(5) = ProviderFactory.CreateDataParameter("SUBFLOW_COUNT", sSubFlowCount)
            para(6) = ProviderFactory.CreateDataParameter("ROLEID", sRoleId)

            If (syRoleHis.loadBySQL(sSQL, para)) Then

                ' 如果“Roleid”不為Nothing
                If Not syRoleHis.getAttribute("SUBFLOW_SEQ") Is Nothing Then
                    If Convert.ToInt32(syRoleHis.getAttribute("SUBFLOW_SEQ").ToString()) > 0 Then
                        Return Convert.ToInt32(syRoleHis.getAttribute("SUBFLOW_SEQ").ToString())
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
                                   "   SY_REL_ROLE_FUNCTION_HIS" & _
                                   " WHERE SY_REL_ROLE_FUNCTION_HIS.CASEID = " & ProviderFactory.PositionPara & "CASEID"

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
#End Region

End Class
