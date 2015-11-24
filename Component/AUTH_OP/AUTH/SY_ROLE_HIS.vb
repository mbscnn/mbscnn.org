Imports com.Azion.NET.VB

Public Class SY_ROLE_HIS
    Inherits BosBase

    Public Sub New(ByVal dbManager As com.Azion.NET.VB.DatabaseManager)
        MyBase.new("SY_ROLE_HIS", dbManager)
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
    ''' <param name="iSubFlowSeq">序號</param>
    ''' <param name="iSubFlowCount">流程數</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Add by Avril 2012/05/09
    ''' </remarks>
    Function loadByPK(ByVal sRoleId As String, ByVal sCaseId As String, ByVal sStepNo As String, ByVal iSubFlowSeq As Integer, ByVal iSubFlowCount As Integer) As Boolean
        Try
            Dim para(4) As IDbDataParameter
            para(0) = ProviderFactory.CreateDataParameter("ROLEID", sRoleId)
            para(1) = ProviderFactory.CreateDataParameter("CASEID", sCaseId)
            para(2) = ProviderFactory.CreateDataParameter("STEP_NO", sStepNo)
            para(3) = ProviderFactory.CreateDataParameter("SUBFLOW_SEQ", iSubFlowSeq)
            para(4) = ProviderFactory.CreateDataParameter("SUBFLOW_COUNT", iSubFlowCount)

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
                                   "   SY_ROLE_HIS" & _
                                   " WHERE SY_ROLE_HIS.CASEID = " & ProviderFactory.PositionPara & "CASEID"

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
    ''' 根據角色編號查詢資料
    ''' </summary>
    ''' <param name="sRoleId">角色編號</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Add by Avril 2012/05/10
    ''' </remarks>
    Function loadDataByRoleId(ByVal sRoleId As String) As Integer
        Try

            Dim syRole As New SY_ROLE(Me.getDatabaseManager)
            Dim sSQL As String = "SELECT *  FROM SY_ROLE_HIS WHERE  ROLEID = " & ProviderFactory.PositionPara & "ROLEID "

            Dim paras(0) As System.Data.IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("ROLEID", sRoleId)

            Return MyBase.loadBySQL(sSQL, paras)
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 查詢資料是否在送簽中
    ''' </summary>
    ''' <param name="sFuncNameList">功能名稱集合</param>
    ''' <param name="sRoleIdList">角色編號集合</param>
    ''' <param name="sStepNo">步驟號</param>
    ''' <param name="sCaseId">案號</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Add by Avril 2012/05/10
    ''' </remarks>
    Function loadDataByCon(ByVal sFuncNameList As String, ByVal sRoleIdList As String, ByVal sStepNo As String, ByVal sCaseId As String) As Integer
        Try

            Dim syRole As New SY_ROLE(Me.getDatabaseManager)
            Dim sSQL As String = "SELECT  roleid FROM SY_ROLE_HIS " & _
                                 " WHERE  1=1 "

            If sFuncNameList.Length > 0 Then
                sSQL = sSQL & " AND ROLENAME  IN( " & sFuncNameList & ") "
            End If

            sSQL = sSQL & " AND CASEID IN (SELECT CASEID FROM SY_FLOWSTEP WHERE STATUS<>3)"

            If sStepNo <> "" Then
                sSQL = sSQL & " AND CASEID<> " & ProviderFactory.PositionPara & "CASEID1 "
            End If

            sSQL = sSQL & " UNION"
            sSQL = sSQL & " SELECT roleid FROM SY_ROLE_HIS WHERE 1=1 "
            If sRoleIdList.Length > 0 Then
                sSQL = sSQL & " AND ROLEID  IN( " & sRoleIdList & ") "
            End If
            sSQL = sSQL & " AND CASEID IN (SELECT CASEID FROM SY_FLOWSTEP WHERE STATUS<>3)"

            If sStepNo <> "" Then
                sSQL = sSQL & " AND CASEID<> " & ProviderFactory.PositionPara & "CASEID2 "
            End If

            If sStepNo <> "" Then
                Dim paras(1) As System.Data.IDbDataParameter
                paras(0) = ProviderFactory.CreateDataParameter("CASEID1", sCaseId)
                paras(1) = ProviderFactory.CreateDataParameter("CASEID2", sCaseId)

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
    ''' 取得最大的序號
    ''' </summary>
    ''' <param name="sRoleId">角色編號</param>
    ''' <param name="sCaseId">案號</param>
    ''' <param name="sStepNo">步驟數</param>
    ''' <param name="sSubFlowCount">資料筆數</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Add by Avril 2012/06/28
    ''' </remarks>
    Function getMaxSubFlowSeq(ByVal sRoleId As String, ByVal sCaseId As String, ByVal sStepNo As String, ByVal sSubFlowCount As String) As Integer
        Try
            Dim syRoleHis As New SY_ROLE_HIS(Me.getDatabaseManager)
            Dim sSQL As String = "SELECT MAX(SUBFLOW_SEQ) SUBFLOW_SEQ  FROM SY_ROLE_HIS  " & _
                " WHERE ROLEID = " & ProviderFactory.PositionPara & "ROLEID " & _
                " AND   CASEID = " & ProviderFactory.PositionPara & "CASEID " & _
                " AND   STEP_NO = " & ProviderFactory.PositionPara & "STEP_NO " & _
                " AND SUBFLOW_COUNT = " & ProviderFactory.PositionPara & "SUBFLOW_COUNT "

            Dim para(4) As IDbDataParameter
            para(0) = ProviderFactory.CreateDataParameter("ROLEID", sRoleId)
            para(1) = ProviderFactory.CreateDataParameter("CASEID", sCaseId)
            para(2) = ProviderFactory.CreateDataParameter("STEP_NO", sStepNo)
            para(3) = ProviderFactory.CreateDataParameter("SUBFLOW_COUNT", sSubFlowCount)

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
#End Region
#End Region
End Class
