Imports com.Azion.NET.VB
Public Class SY_FLOWSTEP
    Inherits BosBase

    Public Sub New(ByVal dbManager As com.Azion.NET.VB.DatabaseManager)
        MyBase.new("SY_FLOWSTEP", dbManager)
    End Sub

    Sub New()
        MyBase.New()
        Me.setPrimaryKeys()
    End Sub

#Region "濟南昱勝添加"
#Region "Avril Function"

    ''' <summary>
    ''' 根據案號查詢資料
    ''' </summary>
    ''' <param name="sCaseId">案號</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Add by Avril 2012/05/09
    ''' </remarks>
    Function loadByCaseId(ByVal sCaseId As String) As Boolean
        Try
            Dim sSQL As String = "SELECT TOP 1 *  FROM SY_FLOWSTEP WHERE STATUS=3 AND CASEID = " & ProviderFactory.PositionPara & "CASEID " & _
                                " order by STARTTIME  desc "

            Dim paras(0) As System.Data.IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("CASEID", sCaseId)

            Return MyBase.loadBySQL(sSQL, paras)
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

#Region "Lake Function"
    ''' <summary>
    ''' 根據“案件編號”更新“案件所有人”為登錄人員
    ''' </summary>
    ''' <param name="sWorkingId">代理使用者</param>
    ''' <param name="sWorkingUserId">登入使用者</param>
    ''' <param name="sCaseId">案件編號</param>
    ''' <returns></returns>
    ''' <remarks>[Lake] 2012/05/14 Cretaed</remarks>
    Public Function updateByCaseId(ByVal sWorkingId As String, ByVal sWorkingUserId As String, ByVal sCaseId As String) As Boolean
        Try
            Dim sSql As String = "UPDATE " & _
                                 "   SY_FLOWSTEP " & _
                                 "SET " & _
                                 "    OWNER = " & ProviderFactory.PositionPara & "WORKINGID " & _
                                 "   ,CLIENT = " & ProviderFactory.PositionPara & "CLIENT " & _
                                 "WHERE " & _
                                 "   CASEID IN('" & sCaseId & "')"

            Dim paras(1) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("WORKINGID", sWorkingId)
            paras(1) = ProviderFactory.CreateDataParameter("CLIENT", sWorkingUserId)

            Return Me.ExecuteNonQuery(CommandType.Text, sSql, paras) > 0
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

#Region "Zack Function"

    ''' <summary>
    ''' 根據案件編號，狀態查詢信息
    ''' </summary>
    ''' <param name="sStatus">狀態</param>
    ''' <param name="sCaseID">案件編號</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Zack  2012-05-15 Create
    ''' </remarks>
    Public Function loadByCaseIDStatus(ByVal sStatus As String, ByVal sCaseID As String) As Boolean
        Try
            Dim sSql As String = "SELECT * FROM sy_flowstep " & _
                                 " WHERE status=" & ProviderFactory.PositionPara & "status" & _
                                 " AND caseid=" & ProviderFactory.PositionPara & "caseid"

            Dim paras(1) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("caseid", sCaseID)
            paras(1) = ProviderFactory.CreateDataParameter("status", sStatus)

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
    ''' 根據案件編號，步驟號查詢CLIENT
    ''' </summary>
    ''' <param name="sCaseID">案件編號</param>
    ''' <param name="sStepNO">步驟號</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Zack 2012-05-24 Create
    ''' </remarks>
    Public Function loadByCaseIDStepNO(ByVal sCaseID As String, ByVal sStepNO As String) As Boolean
        Try
            Dim sSql As String = "SELECT DISTINCT CLIENT FROM sy_flowstep " & _
                                 " WHERE caseid=" & ProviderFactory.PositionPara & "caseid" & _
                                 " AND STEP_NO != " & ProviderFactory.PositionPara & "STEP_NO"

            Dim paras(1) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("caseid", sCaseID)
            paras(1) = ProviderFactory.CreateDataParameter("STEP_NO", sStepNO)

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
    ''' 與臨時檔關聯查詢流程步驟的資料
    ''' </summary>
    ''' <param name="sStaffid">員工編號</param>
    ''' <param name="sBraDepno">部門編號</param>
    ''' <param name="sFuncCode"></param>
    ''' <param name="sStatus">狀態</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Zack 2012-06-12 Create
    ''' </remarks>
    Public Function loadRelSYTEMPINFO(ByVal sStaffid As String, ByVal sBraDepno As String, ByVal sFuncCode As String, ByVal sStatus As String) As Boolean
        Try
            Dim sSql As String = "SELECT SY_FLOWSTEP.* FROM SY_FLOWSTEP " & _
                                " JOIN SY_TEMPINFO ON SY_FLOWSTEP.CASEID = SY_TEMPINFO.CASEID " & _
                                " WHERE SY_TEMPINFO.STAFFID =" & ProviderFactory.PositionPara & "STAFFID" & _
                                " AND SY_TEMPINFO.BRA_DEPNO =" & ProviderFactory.PositionPara & "BRA_DEPNO" & _
                                " AND SY_TEMPINFO.FUNCCODE =" & ProviderFactory.PositionPara & "FUNCCODE" & _
                                " AND SY_FLOWSTEP.STATUS =" & ProviderFactory.PositionPara & "STATUS"

            Dim paras(3) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("STAFFID", sStaffid)
            paras(1) = ProviderFactory.CreateDataParameter("BRA_DEPNO", sBraDepno)
            paras(2) = ProviderFactory.CreateDataParameter("FUNCCODE", sFuncCode)
            paras(3) = ProviderFactory.CreateDataParameter("STATUS", sStatus)

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
