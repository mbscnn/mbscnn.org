Imports com.Azion.NET.VB

Public Class SY_TEMPINFO
    Inherits BosBase

    Sub New()
        MyBase.New()
        Me.setPrimaryKeys()
    End Sub

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.new("SY_TEMPINFO", dbManager)
    End Sub

    Sub New(ByVal dbManager As DatabaseManager, ByVal sTableName As String)
        MyBase.new(sTableName, dbManager)
    End Sub


#Region "濟南昱勝添加"
#Region "Avril Function"

    ''' <summary>
    ''' 根據登入者編號，部門編號，程式碼查詢資料
    ''' </summary>
    ''' <param name="sStaffid">根據登入者編號</param>
    ''' <param name="sDepNo">部門編號</param>
    ''' <param name="sFunccode">程式碼</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Add by Avril 2012/04/18
    ''' </remarks>
    Function loadTempData(ByVal sStaffid As String, ByVal sDepNo As String, ByVal sFunccode As String) As Boolean
        Try
            Dim strSql = "SELECT " & _
                         "    * FROM " & getTableName() & " " & _
                         " WHERE " & _
                         "  STAFFID = " & ProviderFactory.PositionPara & "STAFFID" & _
                         " AND BRA_DEPNO = " & ProviderFactory.PositionPara & "BRA_DEPNO" & _
                         " AND FUNCCODE = " & ProviderFactory.PositionPara & "FUNCCODE" & _
                         " ORDER BY FUNCCODE DESC"


            Dim para(2) As IDbDataParameter
            para(0) = ProviderFactory.CreateDataParameter("STAFFID", sStaffid)
            para(1) = ProviderFactory.CreateDataParameter("BRA_DEPNO", sDepNo)
            para(2) = ProviderFactory.CreateDataParameter("FUNCCODE", sFunccode)

            Return MyBase.loadBySQL(strSql, para)

        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' 根據案號，流程號，部門編號，功能編號查詢資料
    ''' </summary>
    ''' <param name="sCaseId">案號</param>
    ''' <param name="sStepNo">流程號</param>
    ''' <param name="sDepNo">部門編號</param>
    ''' <param name="sFunccode">功能編號</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Add by Avril 2012/05/16
    ''' </remarks>
    Function loadTempDataByCon(ByVal sCaseId As String, ByVal sStepNo As String, ByVal sDepNo As String, ByVal sFunccode As String) As Boolean
        Try
            Dim strSql = "SELECT " & _
                         "    * FROM " & getTableName() & " " & _
                         " WHERE " & _
                         "  STAFFID = ( select distinct client from sy_flowstep where  caseid= " & ProviderFactory.PositionPara & "CASEID" & _
                         " and STEP_NO <>" & ProviderFactory.PositionPara & "STEPNO" & ")" &
                         " AND BRA_DEPNO = " & ProviderFactory.PositionPara & "BRA_DEPNO" & _
                         " AND FUNCCODE = " & ProviderFactory.PositionPara & "FUNCCODE"

            Dim para(3) As IDbDataParameter
            para(0) = ProviderFactory.CreateDataParameter("CASEID", sCaseId)
            para(1) = ProviderFactory.CreateDataParameter("STEPNO", sStepNo)
            para(2) = ProviderFactory.CreateDataParameter("BRA_DEPNO", sDepNo)
            para(3) = ProviderFactory.CreateDataParameter("FUNCCODE", sFunccode)

            Return MyBase.loadBySQL(strSql, para)


        Catch ex As Exception
            Throw
        End Try
    End Function
    ''' <summary>
    ''' 根據主鍵查找資料
    ''' </summary>
    ''' <param name="sStaffid">登入者</param>
    ''' <param name="sDepNo">部門編號</param>
    ''' <param name="sFuncCode">程式代碼</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Add by Avril 2012/04/19
    ''' </remarks>
    Function loadByPK(ByVal sStaffid As String, ByVal sDepNo As String, ByVal sFuncCode As String) As Boolean
        Try
            Dim para(2) As IDbDataParameter

            para(0) = ProviderFactory.CreateDataParameter("STAFFID", sStaffid)
            para(1) = ProviderFactory.CreateDataParameter("BRA_DEPNO", sDepNo)
            para(2) = ProviderFactory.CreateDataParameter("FUNCCODE", sFuncCode)

            Return MyBase.loadBySQL(para)

        Catch ex As Exception
            Throw
        End Try
    End Function


    Function loadBySQL_NullCaseid(ByVal sStaffid As String, ByVal sDepNo As String, ByVal sFuncCode As String) As Boolean
        Try
            Dim para(2) As IDbDataParameter

            Dim strSql = _
                "select * from " & getTableName() & vbCrLf & _
                " where STAFFID = " & ProviderFactory.PositionPara & "STAFFID" & vbCrLf & _
                "   and BRA_DEPNO = " & ProviderFactory.PositionPara & "BRA_DEPNO" & vbCrLf & _
                "   and FUNCCODE = " & ProviderFactory.PositionPara & "FUNCCODE" & vbCrLf & _
                "   and CASEID is null"

            para(0) = ProviderFactory.CreateDataParameter("STAFFID", sStaffid)
            para(1) = ProviderFactory.CreateDataParameter("BRA_DEPNO", sDepNo)
            para(2) = ProviderFactory.CreateDataParameter("FUNCCODE", sFuncCode)

            Return MyBase.loadBySQL(strSql, para)

        Catch ex As Exception
            Throw
        End Try
    End Function


#End Region

#Region "Lake Function"

    ''' <summary>
    ''' 根據案件編號取得TempInfo資料
    ''' </summary>
    ''' <param name="sCaseId">案件編號</param>
    ''' <returns></returns>
    ''' <remarks>[Lake] 2012/06/06 Created</remarks>
    Public Function loadByCaseId(ByVal sCaseId As String) As Boolean
        Try
            Dim sSql As String = "SELECT * FROM " & getTableName() & " WHERE CASEID = " & ProviderFactory.PositionPara & "CASEID"

            Dim paras As IDbDataParameter = ProviderFactory.CreateDataParameter("CASEID", sCaseId)

            Return Me.loadBySQL(sSql, paras)


        Catch ex As Exception
            Throw
        End Try
    End Function
#End Region

#Region "Zack Function"

    ''' <summary>
    ''' 根據主鍵刪除資料
    ''' </summary>
    ''' <param name="sSTAFFID">人員編號</param>
    ''' <param name="sBRA_DEPNO">部門編號</param>
    ''' <param name="sFUNCCODE"></param>
    ''' <remarks>
    ''' Zack 2012-05-24 Create
    ''' </remarks>
    Public Sub DeleteByPK(ByVal sSTAFFID As String, ByVal sBRA_DEPNO As String, ByVal sFUNCCODE As String)
        Try
            Dim sSQL As String = " DELETE FROM " & getTableName() & " " & _
                                " WHERE STAFFID =" & ProviderFactory.PositionPara & "STAFFID" & _
                                " AND BRA_DEPNO = " & ProviderFactory.PositionPara & "BRA_DEPNO" & _
                                " AND FUNCCODE = " & ProviderFactory.PositionPara & "FUNCCODE"

            Dim para(2) As IDataParameter
            para(0) = ProviderFactory.CreateDataParameter("STAFFID", sSTAFFID)
            para(1) = ProviderFactory.CreateDataParameter("BRA_DEPNO", sBRA_DEPNO)
            para(2) = ProviderFactory.CreateDataParameter("FUNCCODE", sFUNCCODE)

            If Me.getDatabaseManager.isTransaction Then
                DBObject.ExecuteNonQuery(Me.getDatabaseManager.getTransaction, CommandType.Text, sSQL, para)
            Else
                DBObject.ExecuteNonQuery(Me.getDatabaseManager.getConnection, CommandType.Text, sSQL, para)
            End If


        Catch ex As Exception
            Throw
        End Try
    End Sub

    ''' <summary>
    ''' 根據案件編號刪除資料
    ''' </summary>
    ''' <param name="sCaseID">案件編號</param>
    ''' <remarks>
    ''' Zack 2012-06-06 Create
    ''' </remarks>
    Public Sub deleteByCaseID(ByVal sCaseID As String)
        Try
            Dim sSQL As String = "DELETE FROM " & getTableName() & " WHERE CASEID =" & ProviderFactory.PositionPara & "CASEID"

            Dim para(0) As IDataParameter
            para(0) = ProviderFactory.CreateDataParameter("CASEID", sCaseID)

            If Me.getDatabaseManager.isTransaction Then
                DBObject.ExecuteNonQuery(Me.getDatabaseManager.getTransaction, CommandType.Text, sSQL, para)
            Else
                DBObject.ExecuteNonQuery(Me.getDatabaseManager.getConnection, CommandType.Text, sSQL, para)
            End If


        Catch ex As Exception
            Throw
        End Try
    End Sub
#End Region
#End Region
End Class

