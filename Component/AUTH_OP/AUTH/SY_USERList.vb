Imports com.Azion.NET.VB
Imports AUTH_OP.TABLE
Imports System.Text

Public Class SY_USERList
    Inherits BosList

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.New("SY_USER", dbManager)
    End Sub

    Overrides Function newBos() As BosBase
        Dim bos As SY_USER = New SY_USER(MyBase.getDatabaseManager)
        Return bos
    End Function

#Region "濟南昱勝添加"
#Region "Lake Function"

    ''' <summary>
    ''' 正常模塊
    ''' 點選一級單位，
    ''' 查詢左側人員列表
    ''' </summary>
    ''' <param name="sStaffId">要排除的右側人員組合</param>
    ''' <returns></returns>
    ''' <remarks>[Lake] 2012/06/25 Created</remarks>
    Public Function getAllStaff(ByVal sStaffId As String) As Boolean
        Try
            Dim sSql As String = "SELECT " & _
                "STAFFID,SUBSTRING(SY_USER.STAFFID,3,5) + ' ' + USERNAME AS NAME FROM SY_USER " & _
                "WHERE STATUS=0 " & _
                "AND STAFFID NOT IN('" & sStaffId & "') ORDER BY STAFFID"

            Return Me.loadBySQL(sSql) > 0
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 根據部門代碼查詢人員資料
    ''' 如果sStaffId不為空，則要排除掉傳入的人員
    ''' </summary>
    ''' <param name="sBraDepNo"></param>
    ''' <param name="sStaffId"></param>
    ''' <returns></returns>
    ''' <remarks>[Lake] 2012/06/25 Created</remarks>
    Public Function getStaffByBraDepNo(ByVal sBraDepNo As String, ByVal sStaffId As String) As Boolean
        Try
            Dim sb As New StringBuilder()

            sb.Append("SELECT " & _
                      " SY_USER.STAFFID AS STAFFID " & _
                      ",SUBSTRING(SY_USER.STAFFID,3,5) + ' ' + SY_USER.USERNAME AS NAME " & _
                      "FROM SY_USER " & _
                      "INNER JOIN SY_REL_BRANCH_USER ON SY_USER.STAFFID = SY_REL_BRANCH_USER.STAFFID " & _
                      "WHERE SY_USER.STATUS = 0 AND " & _
                      "SY_REL_BRANCH_USER.BRA_DEPNO = " & ProviderFactory.PositionPara & "BRADEPNO ")

            If sStaffId <> "" Then
                sb.Append("AND SY_USER.STAFFID NOT IN('" & sStaffId & "')")
            End If

            Dim paras As IDbDataParameter = ProviderFactory.CreateDataParameter("BRADEPNO", sBraDepNo)

            Return Me.loadBySQL(sb.ToString(), paras)
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function


    ''' <summary>
    ''' 根據STAFFID取得個人資訊
    ''' </summary>
    ''' <param name="sStaffId"></param>
    ''' <returns></returns>
    ''' <remarks>[Lake] 2012/05/07 Created</remarks>
    Public Function loadByStaffId(ByVal sStaffId As String) As Boolean
        Try
            Dim sSql As String = "SELECT " & _
                                 "    SY_USER.USERNAME " & _
                                 "   ,SY_USER.STAFFID " & _
                                 "   ,SY_USER.JOBTITLE " & _
                                 "   ,SY_BRANCH. BRCNAME " & _
                                 "   ,SY_USER.OFFICETEL " & _
                                 "   ,SY_USER.Email " & _
                                 "FROM " & _
                                 "      SY_USER LEFT OUTER JOIN SY_REL_BRANCH_USER " & _
                                 "   ON SY_REL_BRANCH_USER.STAFFID = SY_USER.STAFFID " & _
                                 "      LEFT OUTER JOIN SY_BRANCH " & _
                                 "   ON SY_REL_BRANCH_USER.BRA_DEPNO = SY_BRANCH.BRA_DEPNO " & _
                                 "WHERE " & _
                                 "   SY_USER.STAFFID = " & ProviderFactory.PositionPara & "STAFFID"

            Dim paras As IDbDataParameter = ProviderFactory.CreateDataParameter("STAFFID", sStaffId)

            Return (Me.loadBySQL(sSql, paras))
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 根據“部門代碼”查詢被代理人列表
    ''' </summary>
    ''' <param name="sBraDepNo">部門代碼</param>
    ''' <param name="sStaffid">人員代碼</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' [Lake] 2012/05/11 Created
    ''' </remarks>
    Public Function getStaffList(ByVal sBraDepNo As String, Optional ByVal sStaffid As String = "") As Boolean
        Try
            Dim sSql As String = "Select " & _
                                 "    SY_USER.STAFFID + ';' + CONVERT(VARCHAR,SY_REL_BRANCH_USER.BRA_DEPNO) AS BRADEPNO " & _
                                 "   ,SY_USER.STAFFID+' '+ SY_USER.USERNAME AS USERINFO " & _
                                 "From " & _
                                 "    SY_USER JOIN SY_REL_BRANCH_USER " & _
                                 "on SY_REL_BRANCH_USER.STAFFID = SY_USER.STAFFID and SY_USER.STATUS = 0 " & _
                                 "Join SY_BRANCH " & _
                                 "on SY_REL_BRANCH_USER.BRA_DEPNO = SY_BRANCH.BRA_DEPNO " & _
                                 "Where SY_BRANCH.BRA_DEPNO = " & ProviderFactory.PositionPara & "BRADEPNO"

            Dim paras As New List(Of IDbDataParameter)
            paras.Add(ProviderFactory.CreateDataParameter("BRADEPNO", sBraDepNo))

            If String.IsNullOrEmpty(sStaffid) = False Then
                sSql &= "  and SY_USER.STAFFID = " & ProviderFactory.PositionPara & "STAFFID"
                paras.Add(ProviderFactory.CreateDataParameter("STAFFID", sStaffid))
            End If

            Return Me.loadBySQL(sSql, paras.ToArray) > 0
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 取得代理人列表
    ''' </summary>
    ''' <param name="sBraDepNo">部門代碼</param>
    ''' <returns></returns>
    ''' <remarks>[Lake] 2012/05/11 Created</remarks>
    Public Function getAgentStaffList(ByVal sBraDepNo As String) As Boolean
        Try
            Dim sSql As String = "Select " & _
                                 "    SY_USER.STAFFID + ';' + CONVERT(VARCHAR,SY_REL_BRANCH_USER.BRA_DEPNO) AS AGENTBRADEPNO " & _
                                 "   ,SY_USER.STAFFID+' '+ SY_USER.USERNAME AS AGENTUSERINFO " & _
                                 "From SY_USER " & _
                                 "join SY_REL_BRANCH_USER on SY_REL_BRANCH_USER.STAFFID = SY_USER.STAFFID and SY_USER.STATUS=0 " & _
                                 "join SY_BRANCH on SY_REL_BRANCH_USER.BRA_DEPNO = SY_BRANCH.BRA_DEPNO " & _
                                 "Where SY_USER.STATUS=0 " & _
                                 "AND SY_BRANCH. BRA_DEPNO = " & ProviderFactory.PositionPara & "BRADEPNO"

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

    ''' <summary>
    ''' 當前點選單位有過人員變動，取得變動后人員列表
    ''' </summary>
    ''' <param name="sBraDepNo"></param>
    ''' <returns></returns>
    ''' <remarks>[Lake] 2012/05/23 Created</remarks>
    Public Function getEditLeftStaff(ByVal sBraDepNo As String, ByVal sStaffId As String) As Boolean
        Try
            Dim sSql As String = "SELECT " & _
                                 "    SY_USER.STAFFID+' '+SY_USER.USERNAME AS USERNAME " & _
                                 "   ,SY_USER.STAFFID +';L' AS STAFFID " & _
                                 "FROM " & _
                                 "   SY_USER " & _
                                 "   INNER JOIN SY_REL_BRANCH_USER ON SY_USER.STAFFID = SY_REL_BRANCH_USER.STAFFID AND SY_USER.STATUS = 0 " & _
                                 "   AND SY_REL_BRANCH_USER.BRA_DEPNO IN('" & sBraDepNo & "') " & _
                                 "   INNER JOIN SY_BRANCH ON SY_BRANCH.BRA_DEPNO = SY_REL_BRANCH_USER.BRA_DEPNO AND SY_BRANCH.DISABLED = 0 " & _
                                 "WHERE " & _
                                 "   SY_USER.STAFFID NOT IN('" & sStaffId & "') ORDER BY SY_USER.STAFFID"

            Return Me.loadBySQL(sSql) > 0
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 根據部門查詢人員
    ''' </summary>
    ''' <param name="sBraDepNo">點選部門</param>
    ''' <returns></returns>
    ''' <remarks>[Lake] 2012/05/18 Created</remarks>
    Public Function getRightStaff(ByVal sBraDepNo As String) As Boolean
        Try
            Dim sSql As String = "SELECT " & _
                                 "    SY_USER.STAFFID+' '+SY_USER.USERNAME AS USERNAME " & _
                                 "   ,SY_USER.STAFFID +';R' AS STAFFID " & _
                                 "FROM " & _
                                 "   SY_USER " & _
                                 "   INNER JOIN SY_REL_BRANCH_USER ON SY_USER.STAFFID = SY_REL_BRANCH_USER.STAFFID AND SY_USER.STATUS = 0 " & _
                                 "   INNER JOIN SY_BRANCH ON SY_BRANCH.BRA_DEPNO = SY_REL_BRANCH_USER.BRA_DEPNO AND SY_BRANCH.DISABLED = 0 " & _
                                 "WHERE " & _
                                 "   SY_REL_BRANCH_USER.BRA_DEPNO IN(" & ProviderFactory.PositionPara & "BRADEPNO)  ORDER BY SY_USER.STAFFID"

            Dim paras As IDbDataParameter = ProviderFactory.CreateDataParameter("BRADEPNO", sBraDepNo)

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
    ''' 存在XML時，
    ''' 點選一級單位時，
    ''' 查詢【正常模塊】左側人員列表（排除點選單位或者XML中"I,N"人員）
    ''' </summary>
    ''' <param name="sStaffIf">排除人員</param>
    ''' <returns></returns>
    ''' <remarks>[Lake] 2012/05/29 Created</remarks>
    Public Function getLeftFirstLevelStaffId(ByVal sStaffIf As String) As Boolean
        Try
            Dim sSql As String = "SELECT " & _
                                 "    DISTINCT SY_USER.STAFFID AS ID, " & _
                                 "    SY_USER.STAFFID+' '+SY_USER.USERNAME AS USERNAME " & _
                                 "   ,SY_USER.STAFFID + ';L' AS STAFFID " & _
                                 "FROM SY_USER " & _
                                 "   INNER JOIN SY_REL_BRANCH_USER ON SY_USER.STAFFID = SY_REL_BRANCH_USER.STAFFID AND SY_USER.STATUS = 0 " & _
                                 "   INNER JOIN SY_BRANCH ON SY_BRANCH.BRA_DEPNO = SY_REL_BRANCH_USER.BRA_DEPNO AND SY_BRANCH.DISABLED = 0 " & _
                                 "WHERE SY_USER.STATUS=0 " & _
                                 "AND SY_USER.STAFFID NOT IN('" & sStaffIf & "') ORDER BY SY_USER.STAFFID"

            Return Me.loadBySQL(sSql) > 0
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 存在XML時，
    ''' 點選一級單位的下級單位時，
    ''' 查詢【正常模塊】左側人員列表
    ''' </summary>
    ''' <param name="sWorkingId">黨前登錄人員編號</param>
    ''' <param name="sWorkingDepNo">黨前登錄人員部門代碼</param>
    ''' <param name="sFunccode">funccode</param>
    ''' <param name="sAllBraDepNo">點選部門上級部門及所有下屬部門代碼集合</param>
    ''' <param name="sXmlAllBraDepNo">XML中所有部門代碼集合</param>
    ''' <param name="sXmlINStaffId">XML中標記為I,N員工編號集合</param>
    ''' <param name="sXmlINBraDepNo">屬於黨前選擇單位上級單位的下屬單位，且存在于XML中標記為I,N的單位代碼集合</param>
    ''' <returns></returns>
    ''' <remarks>[Lake] 2012/05/29 Created</remarks>
    Public Function getLeftNotFirstLevelStaffId(ByVal sWorkingId As String, ByVal sWorkingDepNo As String, ByVal sFunccode As String _
                                                , ByVal sAllBraDepNo As String, ByVal sXmlAllBraDepNo As String, ByVal sXmlINStaffId As String, _
                                                ByVal sXmlINBraDepNo As String) As Boolean
        Try
            Dim sSql As New StringBuilder()
            sSql.Append("DECLARE @tempData XML ")
            sSql.Append("SELECT @tempData = TEMPDATA FROM SY_TEMPINFO WHERE STAFFID=" & ProviderFactory.PositionPara & "STAFFID ")
            sSql.Append("AND BRA_DEPNO=" & ProviderFactory.PositionPara & "BRA_DEPNO ")
            sSql.Append("AND FUNCCODE=" & ProviderFactory.PositionPara & "FUNCCODE ")
            sSql.Append("SELECT ")
            sSql.Append("     SY_USER.STAFFID+' '+SY_USER.USERNAME AS USERNAME ")
            sSql.Append("    ,temp.item.query('STAFFID').value('.[1]','nvarchar(20)')+';L' AS STAFFID ")
            sSql.Append("FROM ")
            sSql.Append("    @tempData.nodes('/SY/SY_REL_BRANCH_USER') AS temp(item) ")
            sSql.Append("LEFT OUTER JOIN SY_USER on temp.item.query('STAFFID').value('.[1]','nvarchar(20)') = SY_USER.STAFFID ")
            sSql.Append("WHERE ")
            sSql.Append("        (temp.item.query('OPERATION').value('.[1]','nvarchar(20)') = 'I' or temp.item.query('OPERATION').value('.[1]','nvarchar(20)') = 'N') ")
            sSql.Append("    AND temp.item.query('BRA_DEPNO').value('.[1]','nvarchar(20)') IN('" & sAllBraDepNo & "') ")
            sSql.Append("    AND temp.item.query('STAFFID').value('.[1]','nvarchar(20)') NOT IN('" & sXmlINStaffId & "') ")
            sSql.Append("UNION ")
            sSql.Append("SELECT ")
            sSql.Append("     SY_USER.STAFFID+' '+SY_USER.USERNAME AS USERNAME ")
            sSql.Append("    ,SY_USER.STAFFID +';L' AS STAFFID  ")
            sSql.Append("FROM ")
            sSql.Append("    SY_USER ")
            sSql.Append("INNER JOIN SY_REL_BRANCH_USER ON SY_USER.STAFFID = SY_REL_BRANCH_USER.STAFFID ")
            sSql.Append("INNER JOIN SY_BRANCH ON SY_BRANCH.BRA_DEPNO = SY_REL_BRANCH_USER.BRA_DEPNO ")
            sSql.Append("WHERE ")
            sSql.Append("        SY_USER.STATUS=0 AND SY_BRANCH.DISABLED =0 ")
            sSql.Append("    AND SY_BRANCH.BRA_DEPNO IN('" & sAllBraDepNo & "') ")
            sSql.Append("    AND SY_BRANCH.BRA_DEPNO NOT IN('" & sXmlAllBraDepNo & "') ")
            sSql.Append("    And SY_USER.STAFFID NOT IN('" & sXmlINStaffId & "')  ORDER BY STAFFID")

            Dim paras(2) As IDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("STAFFID", sWorkingId)
            paras(1) = ProviderFactory.CreateDataParameter("BRA_DEPNO", sWorkingDepNo)
            paras(2) = ProviderFactory.CreateDataParameter("FUNCCODE", sFunccode)

            Return Me.loadBySQL(sSql.ToString(), paras) > 0
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 取得右側人員列表
    ''' </summary>
    ''' <param name="sWorkingId">登錄人員編號</param>
    ''' <param name="sWorkingDepNo">登錄人員所在部門代碼</param>
    ''' <param name="sFunccode">功能編號</param>
    ''' <param name="sBraDepNoInXml">有過人員變動的單位</param>
    ''' <param name="sBraDepNoNotInXml">沒有人員變動的單位</param>
    ''' <returns></returns>
    ''' <remarks>[Lake] 2012/06/19 Created</remarks>
    Public Function getRightStaffList(ByVal sWorkingId As String, ByVal sWorkingDepNo As String, ByVal sFunccode As String, _
                                      ByVal sBraDepNoInXml As String, ByVal sBraDepNoNotInXml As String) As Boolean
        Try
            Dim sSql As New StringBuilder()
            sSql.Append("DECLARE @tempData XML ")
            sSql.Append("SELECT @tempData = TEMPDATA FROM SY_TEMPINFO WHERE STAFFID=" & ProviderFactory.PositionPara & "STAFFID ")
            sSql.Append("AND BRA_DEPNO=" & ProviderFactory.PositionPara & "BRA_DEPNO ")
            sSql.Append("AND FUNCCODE=" & ProviderFactory.PositionPara & "FUNCCODE ")
            sSql.Append("SELECT ")
            sSql.Append("     SY_USER.STAFFID+' '+SY_USER.USERNAME AS USERNAME ")
            sSql.Append("    ,temp.item.query('STAFFID').value('.[1]','nvarchar(20)')+';R' AS STAFFID ")
            sSql.Append("FROM ")
            sSql.Append("    @tempData.nodes('/SY/SY_REL_BRANCH_USER') AS temp(item) ")
            sSql.Append("LEFT OUTER JOIN SY_USER on temp.item.query('STAFFID').value('.[1]','nvarchar(20)') = SY_USER.STAFFID ")
            sSql.Append("WHERE ")
            sSql.Append("        (temp.item.query('OPERATION').value('.[1]','nvarchar(20)') = 'I' or temp.item.query('OPERATION').value('.[1]','nvarchar(20)') = 'N') ")
            sSql.Append("    AND temp.item.query('BRA_DEPNO').value('.[1]','nvarchar(20)') IN('" & sBraDepNoInXml & "') ")
            sSql.Append("UNION ")
            sSql.Append("SELECT ")
            sSql.Append("     SY_USER.STAFFID+' '+SY_USER.USERNAME AS USERNAME ")
            sSql.Append("    ,SY_USER.STAFFID +';R' AS STAFFID  ")
            sSql.Append("FROM ")
            sSql.Append("    SY_USER ")
            sSql.Append("INNER JOIN SY_REL_BRANCH_USER ON SY_USER.STAFFID = SY_REL_BRANCH_USER.STAFFID ")
            sSql.Append("INNER JOIN SY_BRANCH ON SY_BRANCH.BRA_DEPNO = SY_REL_BRANCH_USER.BRA_DEPNO ")
            sSql.Append("WHERE ")
            sSql.Append("        SY_USER.STATUS=0 AND SY_BRANCH.DISABLED =0 ")
            sSql.Append("    AND SY_BRANCH.BRA_DEPNO IN('" & sBraDepNoNotInXml & "') ORDER BY STAFFID")

            Dim paras(2) As IDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("STAFFID", sWorkingId)
            paras(1) = ProviderFactory.CreateDataParameter("BRA_DEPNO", sWorkingDepNo)
            paras(2) = ProviderFactory.CreateDataParameter("FUNCCODE", sFunccode)

            Return Me.loadBySQL(sSql.ToString(), paras) > 0
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 不存在XML時，
    ''' 【正常模塊】左側人員加載
    ''' </summary>
    ''' <param name="sBraDepNo">黨前選擇人員編號</param>
    ''' <returns></returns>
    ''' <remarks>[Lake] 2012/05/29 Created</remarks>
    Public Function getLeftStaffNotSelectBranch(ByVal sBraDepNo As String) As Boolean
        Try
            Dim sSql As String = "SELECT " & _
                                 "    SY_USER.STAFFID+' '+SY_USER.USERNAME AS USERNAME " & _
                                 "   ,SY_USER.STAFFID +';L' AS STAFFID " & _
                                 "FROM SY_USER " & _
                                 "   INNER JOIN SY_REL_BRANCH_USER ON SY_USER.STAFFID = SY_REL_BRANCH_USER.STAFFID AND SY_USER.STATUS = 0 " & _
                                 "   INNER JOIN SY_BRANCH ON SY_BRANCH.BRA_DEPNO = SY_REL_BRANCH_USER.BRA_DEPNO AND SY_BRANCH.DISABLED = 0 " & _
                                 "WHERE SY_USER.STATUS=0 " & _
                                 "AND SY_BRANCH.BRA_DEPNO NOT IN('" & sBraDepNo & "')  ORDER BY SY_USER.STAFFID"


            Return Me.loadBySQL(sSql) > 0
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
    ''' 根據角色編號，部門編號查詢安控待選擇人員信息
    ''' </summary>
    ''' <param name="sRoleID">角色編號</param>
    ''' <param name="sBraDepno">單位編號</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Zack  2012-06-25 Create
    ''' </remarks>
    Public Function loadByBridRoleID(ByVal sRoleID As String, ByVal sBraDepno As String, ByVal sFlag As String) As Integer
        Try
            Dim sSql As String = "SELECT " & _
                                    "SY_USER.STAFFID + ' ' + USERNAME AS LText" & _
                                    ",SY_USER.STAFFID +';' + CONVERT(VARCHAR(10),SY_REL_BRANCH_USER.BRA_DEPNO) AS LValue " & _
                                " FROM SY_USER " & _
                                    "INNER JOIN SY_REL_BRANCH_USER ON SY_USER. STAFFID = SY_REL_BRANCH_USER. STAFFID" & _
                                    " INNER JOIN SY_BRANCH ON SY_BRANCH.BRA_DEPNO = SY_REL_BRANCH_USER.BRA_DEPNO" & _
                                    " LEFT JOIN SY_REL_ROLE_USER on SY_USER. STAFFID = SY_REL_ROLE_USER. STAFFID " & _
                                    " AND SY_BRANCH.BRA_DEPNO = SY_REL_ROLE_USER.BRA_DEPNO and ROLEID=" & ProviderFactory.PositionPara & "ROLEID" & _
                                " WHERE SY_USER.STATUS=0 AND SY_BRANCH.BRA_DEPNO =" & ProviderFactory.PositionPara & "BRA_DEPNO" & _
                                " AND SY_BRANCH.parent=0 "

            ' 若是查詢待選擇人員信息
            If sFlag = "Y" Then
                sSql = sSql & " AND SY_REL_ROLE_USER. STAFFID is null"
            Else
                sSql = sSql & " AND SY_REL_ROLE_USER. STAFFID is not null"
            End If

            Dim para(1) As IDbDataParameter
            para(0) = ProviderFactory.CreateDataParameter("ROLEID", sRoleID)
            para(1) = ProviderFactory.CreateDataParameter("BRA_DEPNO", sBraDepno)

            Return Me.loadBySQL(sSql, para)
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 根據分行單位代碼，角色代碼查詢待選擇人員
    ''' </summary>
    ''' <param name="sBraDepno">單位代碼</param>
    ''' <param name="sROLEID">角色代碼</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Zack 2012-05-14 Create
    ''' Zack 2012-05-14 Update (修改查詢條件)
    ''' </remarks>
    Public Function loadByDepnoRoleID(ByVal sBraDepno As String, ByVal sROLEID As String) As Integer
        Try
            Dim sSql As String = "SELECT SY_USER.STAFFID + ' ' + USERNAME AS LText,SY_USER.STAFFID +';' + CONVERT(VARCHAR(10),SY_REL_BRANCH_USER.BRA_DEPNO) AS LValue FROM SY_USER  " & _
                                " INNER JOIN SY_REL_BRANCH_USER ON SY_USER.STAFFID = SY_REL_BRANCH_USER.STAFFID " & _
                                " WHERE SY_USER.STATUS = 0" & _
                                " AND SY_REL_BRANCH_USER.BRA_DEPNO =" & ProviderFactory.PositionPara & "BRA_DEPNO" & _
                                " AND SY_USER.STAFFID NOT IN " & _
                                " (SELECT STAFFID FROM SY_REL_ROLE_USER " & _
                                " WHERE SY_REL_ROLE_USER.ROLEID=" & ProviderFactory.PositionPara & "ROLEID" & _
                                " AND SY_REL_ROLE_USER.BRA_DEPNO =" & ProviderFactory.PositionPara & "BRA_DEPNO1)" & _
                                " ORDER BY SY_USER.STAFFID"

            Dim para(2) As IDbDataParameter
            para(0) = ProviderFactory.CreateDataParameter("BRA_DEPNO", sBraDepno)
            para(1) = ProviderFactory.CreateDataParameter("BRA_DEPNO1", sBraDepno)
            para(2) = ProviderFactory.CreateDataParameter("ROLEID", sROLEID)

            Return Me.loadBySQL(sSql, para)
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 根據部門分行代號，角色編號查詢信息
    ''' </summary>
    ''' <param name="sBraDepno">部門分行代號</param>
    ''' <param name="sROLEID">角色編號</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Zack 2012-05-15 Create
    ''' Zack 2012-05-14 Update (修改查詢條件)
    ''' </remarks>
    Public Function loadInFlow(ByVal sBraDepno As String, ByVal sROLEID As String) As Integer
        Try
            Dim sSql As String = "SELECT SY_USER.STAFFID + ' ' + USERNAME AS LText,SY_USER.STAFFID +';' + CONVERT(VARCHAR(10),SY_REL_BRANCH_USER.BRA_DEPNO) AS LValue FROM SY_USER" & _
                                " INNER JOIN SY_REL_BRANCH_USER ON SY_USER. STAFFID = SY_REL_BRANCH_USER. STAFFID" & _
                                " WHERE SY_USER.STATUS=0 AND SY_REL_BRANCH_USER.BRA_DEPNO =" & ProviderFactory.PositionPara & "BRA_DEPNO" & _
                                " AND SY_User.STAFFID IN (SELECT STAFFID FROM SY_REL_ROLE_USER " & _
                                " WHERE SY_REL_ROLE_USER.ROLEID=" & ProviderFactory.PositionPara & "ROLEID" & _
                                " AND SY_REL_ROLE_USER.BRA_DEPNO =" & ProviderFactory.PositionPara & "BRA_DEPNO1)" & _
                                " ORDER BY SY_USER.STAFFID"

            Dim para(2) As IDbDataParameter
            para(0) = ProviderFactory.CreateDataParameter("BRA_DEPNO", sBraDepno)
            para(1) = ProviderFactory.CreateDataParameter("BRA_DEPNO1", sBraDepno)
            para(2) = ProviderFactory.CreateDataParameter("ROLEID", sROLEID)

            Return Me.loadBySQL(sSql, para)
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 根據部門代碼查詢人員信息
    ''' </summary>
    ''' <param name="sBraDepno">分行單位代碼</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Zack 2012-05-17 Create
    ''' Zack 2012-06-25 Update
    ''' </remarks>
    Public Function loadByBRA_DEPNO(ByVal sBraDepno As String) As Integer
        Try
            Dim sSql As String = "SELECT SY_USER.STAFFID + ' ' + SY_USER.USERNAME AS USERNAME,SY_USER.STAFFID + ';' +CONVERT(VARCHAR(10),SY_REL_BRANCH_USER.BRA_DEPNO) AS BRA_DEPNO" & _
                                " FROM SY_USER JOIN SY_REL_BRANCH_USER" & _
                                " ON SY_USER. STAFFID = SY_REL_BRANCH_USER. STAFFID" & _
                                " WHERE STATUS =0 AND BRA_DEPNO =" & ProviderFactory.PositionPara & "BRA_DEPNO"

            Dim para(0) As IDbDataParameter
            para(0) = ProviderFactory.CreateDataParameter("BRA_DEPNO", sBraDepno)

            Return Me.loadBySQL(sSql, para)
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 根據人員編號，分行號查詢人員信息
    ''' </summary>
    ''' <param name="sSTAFFID">人員編號</param>
    ''' <param name="sBraDepno">部門代號</param>
    '''  <param name="sBrid">分行號</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Zack 2012-05-18 Create
    ''' </remarks>
    Public Function loadByRoleEdit(ByVal sSTAFFID As String, ByVal sBraDepno As String) As Integer
        Try
            Dim sSql As String = "SELECT SY_USER.STAFFID + ' ' + USERNAME AS LText,SY_USER.STAFFID +';' + CONVERT(VARCHAR(10),SY_REL_BRANCH_USER.BRA_DEPNO) AS LValue FROM SY_USER " & _
                                " INNER JOIN SY_REL_BRANCH_USER " & _
                                " ON SY_USER.STAFFID = SY_REL_BRANCH_USER.STAFFID" & _
                                " WHERE  SY_USER.STATUS=0 " & _
                                " AND SY_REL_BRANCH_USER.BRA_DEPNO =" & ProviderFactory.PositionPara & "BRA_DEPNO" & _
                                " AND SY_USER.STAFFID not in (" + sSTAFFID + ")" & _
                                " ORDER BY SY_USER.STAFFID"

            Dim para(0) As IDbDataParameter
            para(0) = ProviderFactory.CreateDataParameter("BRA_DEPNO", sBraDepno)

            Return Me.loadBySQL(sSql, para)
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
