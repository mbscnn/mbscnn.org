Imports com.Azion.NET.VB

Public Class SY_REL_ROLE_USER
    Inherits BosBase

    Sub New()
        MyBase.New()
        Me.setPrimaryKeys()
    End Sub

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.new("SY_REL_ROLE_USER", dbManager)
    End Sub

#Region "濟南昱勝添加"
#Region "Lake Function"
    ''' <summary>
    ''' 根據部門及人員刪除角色資料
    ''' </summary>
    ''' <param name="sBraDepNo">當前部門</param>
    ''' <param name="sStaffId">人員編號</param>
    ''' <returns></returns>
    ''' <remarks>[Lake] 2012/06/26 Created</remarks>
    Public Function deleteStaffByBraDepNoAndStaffId(sBraDepNo As String, sStaffId As String) As Boolean
        Dim sSql As String = "DELETE SY_REL_ROLE_USER WHERE BRA_DEPNO IN(" & sBraDepNo & ") " & _
                    "AND STAFFID = " & ProviderFactory.PositionPara & "STAFFID"

        Dim paras As IDbDataParameter = ProviderFactory.CreateDataParameter("STAFFID", sStaffId)

        If Me.getDatabaseManager.isTransaction Then
            Return DBObject.ExecuteNonQuery(Me.getDatabaseManager.getTransaction, CommandType.Text, sSql, paras) > 0
        Else
            Return DBObject.ExecuteNonQuery(Me.getDatabaseManager.getConnection, CommandType.Text, sSql, paras) > 0
        End If
    End Function
#End Region

#Region "Zack Function"

    ''' <summary>
    ''' 根據使用者員工編號，角色代碼查詢信息
    ''' </summary>
    ''' <param name="sSTAFFID">使用者員工編號</param>
    ''' <param name="sROLEID">角色代碼</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Zack 2012-05-15 Create
    ''' </remarks>
    Public Function loadByPK(ByVal sSTAFFID As String, ByVal sROLEID As String, ByVal sBRA_DEPNO As String) As Boolean
        Try
            Dim paras(2) As IDbDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("STAFFID", sSTAFFID)
            paras(1) = ProviderFactory.CreateDataParameter("ROLEID", sROLEID)
            paras(2) = ProviderFactory.CreateDataParameter("BRA_DEPNO", sBRA_DEPNO)

            Return Me.loadBySQL(paras)


        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' 根據分行單位代碼，部門代碼刪除信息
    ''' </summary>
    ''' <param name="sROLEID">分行單位代碼</param>
    ''' <param name="sBRID">部門代碼</param>
    ''' <remarks>
    ''' Zack 2012-05-15 Create
    ''' </remarks>
    Public Sub deleteByRoleIDStaffid(ByVal sROLEID As String, ByVal sBRID As String)
        Try
            Dim sSQL As String = " DELETE FROM SY_REL_ROLE_USER " & _
                                 " WHERE ROLEID =" & ProviderFactory.PositionPara & "ROLEID" & _
                                 " AND STAFFID IN(SELECT STAFFID FROM SY_REL_ROLE_USER WHERE BRA_DEPNO = " & _
                                 " (SELECT BRA_DEPNO FROM SY_BRANCH WHERE DISABLED =0 AND BRID=" & ProviderFactory.PositionPara & "BRID AND PARENT = 0))"

            Dim para(1) As IDataParameter
            para(0) = ProviderFactory.CreateDataParameter("ROLEID", sROLEID)
            para(1) = ProviderFactory.CreateDataParameter("BRID", sBRID)

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
    ''' 根據角色編號，部門編號刪除信息
    ''' </summary>
    ''' <param name="sROLEID">角色編號</param>
    ''' <param name="sBRA_DEPNO">部門編號</param>
    ''' <remarks>
    ''' Zack  2012-05-24 Create
    ''' </remarks>
    Public Sub deleteByRoleIDDepNO(ByVal sROLEID As String, ByVal sBRA_DEPNO As String)
        Try
            Dim sSQL As String = " DELETE FROM SY_REL_ROLE_USER " & _
                                 " WHERE  ROLEID =" & ProviderFactory.PositionPara & "ROLEID" & _
                                 " AND BRA_DEPNO =" & ProviderFactory.PositionPara & "BRA_DEPNO"

            Dim para(1) As IDataParameter
            para(0) = ProviderFactory.CreateDataParameter("ROLEID", sROLEID)
            para(1) = ProviderFactory.CreateDataParameter("BRA_DEPNO", sBRA_DEPNO)

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
    ''' 根據員工編號，部門編號刪除信息
    ''' </summary>
    ''' <param name="sSTAFFID">員工編號</param>
    ''' <param name="sBRA_DEPNO">部門編號</param>
    ''' <remarks>
    ''' Zack 2012-05-24 Create
    ''' </remarks>
    Public Sub deleteBySTAFFIDDepNO(ByVal sSTAFFID As String, ByVal sBRA_DEPNO As String, ByVal sRoleid As Integer)
        Try
            Dim sSQL As String = String.Empty

            sSQL = " DELETE FROM SY_REL_ROLE_USER WHERE ROLEID =" & ProviderFactory.PositionPara & "ROLEID" & _
                   " AND BRA_DEPNO=" & ProviderFactory.PositionPara & "BRA_DEPNO" & _
                   " AND staffid=" & ProviderFactory.PositionPara & "staffid"

            Dim para(2) As IDataParameter
            para(0) = ProviderFactory.CreateDataParameter("STAFFID", sSTAFFID)
            para(1) = ProviderFactory.CreateDataParameter("BRA_DEPNO", sBRA_DEPNO)
            para(2) = ProviderFactory.CreateDataParameter("ROLEID", sRoleid)

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
    ''' 根據主鍵刪除數據
    ''' </summary>
    ''' <param name="dtData">鍵值集合</param>
    ''' <remarks>
    ''' Zack  2012-07-02 Create
    ''' </remarks>
    Public Sub deleteByPK(ByVal dtData As DataTable)
        Try
            For i As Integer = 0 To dtData.Rows.Count - 1

                Dim sSQL As String = " DELETE FROM SY_REL_ROLE_USER " & _
                                     " WHERE ROLEID =7 AND STAFFID  =" & ProviderFactory.PositionPara & "STAFFID" & _
                                     " AND BRA_DEPNO =" & ProviderFactory.PositionPara & "BRA_DEPNO"

                Dim para(1) As IDataParameter
                para(0) = ProviderFactory.CreateDataParameter("STAFFID", dtData.Rows(i)("STAFFID"))
                para(1) = ProviderFactory.CreateDataParameter("BRA_DEPNO", dtData.Rows(i)("BRA_DEPNO"))

                If Me.getDatabaseManager.isTransaction Then
                    DBObject.ExecuteNonQuery(Me.getDatabaseManager.getTransaction, CommandType.Text, sSQL, para)
                Else
                    DBObject.ExecuteNonQuery(Me.getDatabaseManager.getConnection, CommandType.Text, sSQL, para)
                End If
            Next


        Catch ex As Exception
            Throw
        End Try
    End Sub
#End Region
#End Region

    ''' <summary>
    ''' 根據當前登錄人員所在部門及其下屬部門查詢【變動人員角色清單】
    ''' TODO:日期部份條件未添加
    ''' </summary>
    ''' <param name="sBraDepNo"></param>
    ''' <returns></returns>
    ''' <remarks>[Lake] 2012/06/05 Created</remarks>
    Public Function getChangeStaffRoleList(ByVal sBraDepNo As Integer,
                                           Optional ByVal sStartYear As String = Nothing,
                                           Optional ByVal sStartMonth As String = Nothing,
                                           Optional ByVal sStartDay As String = Nothing,
                                           Optional ByVal sEndYear As String = Nothing,
                                           Optional ByVal sEndMonth As String = Nothing,
                                           Optional ByVal sEndDay As String = Nothing) As DataTable

        Dim sSql As String
        Dim paramList As New BosParamsList

        Try
            sSql = _
                "select distinct '[' + RIGHT(SY_USER.STAFFID, 5) + '] ' + SY_USER.USERNAME as STAFF," & vbCrLf & _
                "                SY_ROLE.ROLENAME," & vbCrLf & _
                "                HIS.STEP_NO," & vbCrLf & _
                "                case HIS.OPERATION" & vbCrLf & _
                "                  when 'I' then" & vbCrLf & _
                "                   '新增'" & vbCrLf & _
                "                  when 'D' then" & vbCrLf & _
                "                   '刪除'" & vbCrLf & _
                "                  when 'N' then" & vbCrLf & _
                "                   '無異動'" & vbCrLf & _
                "                end as OPERATION," & vbCrLf & _
                "                FS.PROCESSTIME as ENDTIME," & vbCrLf & _
                "                (select top 1 '[' + RIGHT(SY_USER.STAFFID, 5) + '] ' + SY_USER.USERNAME" & vbCrLf & _
                "                   from SY_FLOWSTEP FS1" & vbCrLf & _
                "                  inner join SY_USER" & vbCrLf & _
                "                     on FS1.SENDER = SY_USER.STAFFID" & vbCrLf & _
                "                  where FS1.CASEID = FS.CASEID" & vbCrLf & _
                "                    and FS1.STEP_NO = 'SY000300'" & vbCrLf & _
                "                    and FS1.SUBFLOW_SEQ = FS.SUBFLOW_SEQ) AS STAFFONE," & vbCrLf & _
                "                (select top 1 '[' + RIGHT(SY_USER.STAFFID, 5) + '] ' + SY_USER.USERNAME" & vbCrLf & _
                "                   from SY_FLOWSTEP FS2" & vbCrLf & _
                "                  inner join SY_USER" & vbCrLf & _
                "                     on FS2.SENDER = SY_USER.STAFFID" & vbCrLf & _
                "                  where FS2.CASEID = FS.CASEID" & vbCrLf & _
                "                    and FS2.STEP_NO = 'SY000400'" & vbCrLf & _
                "                    and FS2.SUBFLOW_SEQ = FS.SUBFLOW_SEQ) AS STAFFTWO" & vbCrLf & _
                "  from SY_REL_ROLE_USER_HIS HIS" & vbCrLf & _
                " inner join SY_ROLE" & vbCrLf & _
                "    on HIS.ROLEID = SY_ROLE.ROLEID" & vbCrLf & _
                " inner join SY_USER" & vbCrLf & _
                "    on HIS.STAFFID = SY_USER.STAFFID" & vbCrLf & _
                " inner join SY_FLOWINCIDENT FI" & vbCrLf & _
                "    on FI.CASEID = HIS.CASEID" & vbCrLf & _
                "   and FI.SUBFLOW_SEQ = HIS.SUBFLOW_SEQ" & vbCrLf & _
                " inner join SY_FLOWSTEP FS" & vbCrLf & _
                "    on FS.CASEID = HIS.CASEID" & vbCrLf & _
                "   and FS.SUBFLOW_SEQ = HIS.SUBFLOW_SEQ" & vbCrLf & _
                "   and FS.STEP_NO = HIS.STEP_NO" & vbCrLf & _
                "   and FS.STEP_NO = 'SY000400'" & vbCrLf & _
                " where SY_USER.STATUS = 0" & vbCrLf & _
                "   and SY_ROLE.DISABLED = 0" & vbCrLf & _
                "   and HIS.BRA_DEPNO = @BRA_DEPNO@" & vbCrLf & _
                "   and HIS.OPERATION <> 'N'" & vbCrLf & _
                "   and HIS.APPROVED = 'Y'" & vbCrLf

            paramList.Add("BRA_DEPNO", sBraDepNo)


            If Not String.IsNullOrEmpty(sStartYear) Then
                Dim dtStart As New DateTime(CInt(sStartYear), CInt(sStartMonth), CInt(sStartDay))
                Dim dtEnd As New DateTime(CInt(sEndYear), CInt(sEndMonth), CInt(sEndDay))

                sSql &= _
                "   and convert(date,ENDTIME) >= @STARTTIME@" & vbCrLf & _
                "   and convert(date,ENDTIME) <= @ENDTIME@" & vbCrLf

                paramList.Add("STARTTIME", dtStart)
                paramList.Add("ENDTIME", dtEnd)
            End If

            Return GetDataTable(sSql, paramList.ToArray)

        Catch ex As Exception
            Throw
        End Try

    End Function


End Class
