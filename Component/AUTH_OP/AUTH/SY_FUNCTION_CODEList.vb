Imports com.Azion.NET.VB
Imports AUTH_OP.TABLE
Public Class SY_FUNCTION_CODEList
    : Inherits BosList

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.new("SY_FUNCTION_CODE", dbManager)
    End Sub

    Overrides Function newBos() As BosBase
        Return New SY_FUNCTION_CODE(MyBase.getDatabaseManager)
    End Function

    ''' <summary>
    ''' 依這個員工所登錄的DepNo取得某一子系統的左側列表
    ''' </summary>
    ''' <param name="sStaffid">Sxxxxxx</param>
    ''' <param name="iDepNo"></param>
    ''' <param name="sSysId">D、F、Z</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function genFunList(ByVal sStaffid As String, ByVal iDepNo As Integer, ByVal sSysId As String) As DataTable
        Try
            Dim strSql As String = "WITH rc(FUCNAME, parentid, FUNCCODE, FUCURL, SORTCTRL, HOFLAG, DISABLED,Level) AS" & _
                        " (SELECT FUCNAME, parent, FUNCCODE, FUCURL," & _
                        "         CAST(row_number() over(order by a.SORTCTRL) as varchar(255))," & _
                        "         HOFLAG, DISABLED, 0 Level" & _
                        "    FROM SY_FUNCTION_CODE a" & _
                        "   WHERE parent = 0 and hoflag <> '' " & _
                        "  UNION ALL" & _
                        "  SELECT p.FUCNAME, p.parent, p.FUNCCODE, p.FUCURL," & _
                        "         CAST(RTrim(rc.SORTCTRL) + '-' + CAST(row_number() over(order by p.SORTCTRL) as varchar(10)) as varchar(255))," & _
                        "         p.HOFLAG, p.DISABLED, Level + 1 Level" & _
                        "    FROM SY_FUNCTION_CODE P, rc" & _
                        "   WHERE rc.FUNCCODE = P.PARENT" & _
                        "     and p.PARENT > 0) " & _
                        " select distinct Level, rc.FUCNAME FUNCTIONNAME, rc.parentid, b.SYSID, b.SUBSYSID, b.FUNCCODE, " & _
                        " rc.FUCURL, rc.SORTCTRL, rc.DISABLED, " & ProviderFactory.PositionPara & "STAFFID STAFFID, null CheckFlag, SUBSYSNAME, HOFLAG " & _
                        "  from SY_REL_ROLE_FUNCTION b, rc, SY_SUBSYSID c " & _
                        " where exists( " & _
                        "   select m.roleid, m.* from SY_REL_ROLE_USER m   left join sy_role   on  m.ROLEID = sy_role.ROLEID left join sy_branch  ON m.BRA_DEPNO = SY_BRANCH.BRA_DEPNO  where m.STAFFID = " & ProviderFactory.PositionPara & "STAFFID " & _
                        "    and sy_branch.brid = " & ProviderFactory.PositionPara & "BRID and m.ROLEID = b.ROLEID and sy_role.disabled='0' )" & _
                        " and b.SYSID =" & ProviderFactory.PositionPara & "SYSID" & _
                        " and rc.FUNCCODE = b.FUNCCODE" & _
                        " and rc.DISABLED = 0" & _
                        " and c.SUBSYSID = b.SUBSYSID" & _
                        " order by  SORTCTRL OPTION(MAXRECURSION 5)"

            Dim para(2) As IDbDataParameter
            para(0) = ProviderFactory.CreateDataParameter("STAFFID", sStaffid)
            para(1) = ProviderFactory.CreateDataParameter("BRID", iDepNo)
            para(2) = ProviderFactory.CreateDataParameter("SYSID", sSysId)

            If Me.loadBySQLOnlyDs(strSql, para) > 0 Then
                Return Me.getCurrentDataSet.Tables(0)
            End If
            Return Nothing
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Function genFunListList(ByVal sStaffid As String, ByVal sDepNoList As String, ByVal sSysId As String) As DataTable
        Try
            Dim strSql As String = "WITH rc(FUCNAME, parentid, FUNCCODE, FUCURL, SORTCTRL, HOFLAG, DISABLED,Level) AS" & _
                        " (SELECT FUCNAME, parent, FUNCCODE, FUCURL," & _
                        "         CAST(row_number() over(order by a.SORTCTRL) as varchar(255))," & _
                        "         HOFLAG, DISABLED, 0 Level" & _
                        "    FROM SY_FUNCTION_CODE a" & _
                        "   WHERE parent = 0 and hoflag <> ''  " & _
                        "  UNION ALL" & _
                        "  SELECT p.FUCNAME, p.parent, p.FUNCCODE, p.FUCURL," & _
                        "         CAST(RTrim(rc.SORTCTRL) + '-' + CAST(row_number() over(order by p.SORTCTRL) as varchar(10)) as varchar(255))," & _
                        "         p.HOFLAG, p.DISABLED, Level + 1 Level" & _
                        "    FROM SY_FUNCTION_CODE P, rc" & _
                        "   WHERE rc.FUNCCODE = P.PARENT" & _
                        "     and p.PARENT > 0) " & _
                        " select distinct Level, rc.FUCNAME FUNCTIONNAME, rc.parentid, b.SYSID, b.SUBSYSID, b.FUNCCODE, " & _
                        " rc.FUCURL, rc.SORTCTRL, rc.DISABLED, " & ProviderFactory.PositionPara & "STAFFID STAFFID, null CheckFlag, SUBSYSNAME, HOFLAG " & _
                        "  from SY_REL_ROLE_FUNCTION b, rc, SY_SUBSYSID c " & _
                        " where exists( " & _
                        "   select m.roleid, m.* from SY_REL_ROLE_USER m   left join sy_role   on  m.ROLEID = sy_role.ROLEID  where m.STAFFID = " & ProviderFactory.PositionPara & "STAFFID " & _
                        "    and m.BRA_DEPNO in (" & sDepNoList & ")  and m.ROLEID = b.ROLEID and sy_role.disabled='0' )" & _
                        " and b.SYSID =" & ProviderFactory.PositionPara & "SYSID" & _
                        " and rc.FUNCCODE = b.FUNCCODE" & _
                        " and rc.DISABLED = 0" & _
                        " and c.SUBSYSID = b.SUBSYSID" & _
                        " order by  SORTCTRL OPTION(MAXRECURSION 5)"

            Dim para(1) As IDbDataParameter
            para(0) = ProviderFactory.CreateDataParameter("STAFFID", sStaffid)
            para(1) = ProviderFactory.CreateDataParameter("SYSID", sSysId)

            If Me.loadBySQLOnlyDs(strSql, para) > 0 Then
                Return Me.getCurrentDataSet.Tables(0)
            End If
            Return Nothing
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function


    Function genSYSFunList() As DataTable
        Try 'FUCNAME,parentid,FUNCCODE,FUCURL,SORTCTRL,HOFLAG,DISABLED,Level
            Dim strSql As String = "WITH rc(FUNCTIONNAME," & _
                                                "parentid," & _
                                                "FUNCCODE," & _
                                                "FUCURL," & _
                                                "SORTCTRL," & _
                                                "HOFLAG," & _
                                                "DISABLED," & _
                                                "Level) AS" & _
                                                " (" & _
                                                "" & _
                                                "" & _
                                                " SELECT FUCNAME," & _
                                                "         parent," & _
                                                "         FUNCCODE," & _
                                                "         FUCURL," & _
                                                "         CAST(row_number() over(order by a.SORTCTRL) as varchar(255))," & _
                                                "         HOFLAG," & _
                                                "         DISABLED," & _
                                                "         0 Level" & _
                                                "    FROM SY_FUNCTION_CODE a" & _
                                                "   WHERE parent = 0 and hoflag <> '' " & _
                                                "  UNION ALL" & _
                                                "  SELECT p.FUCNAME," & _
                                                "         p.parent," & _
                                                "         p.FUNCCODE," & _
                                                "         p.FUCURL," & _
                                                "         CAST(RTrim(rc.SORTCTRL) + '-' + CAST(row_number() over (order by p.SORTCTRL) as varchar(10)) as varchar(255))," & _
                                                "         p.HOFLAG," & _
                                                "         p.DISABLED," & _
                                                "         Level + 1 Level" & _
                                                "    FROM SY_FUNCTION_CODE P, rc" & _
                                                "   WHERE rc.FUNCCODE = P.PARENT" & _
                                                "     and p.PARENT > 0)" & _
                                                "select  rc.*" & _
                                                "  from  rc where HOFLAG=0 and rc.DISABLED=0" & _
                                                " order by SORTCTRL OPTION(MAXRECURSION 5)"




            If Me.loadBySQLOnlyDs(strSql) > 0 Then
                Return Me.getCurrentDataSet.Tables(0)
            End If
            Return Nothing
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

#Region "濟南昱勝添加"
#Region "Avril Function"
    ''' <summary>
    ''' 根據角色查詢資料
    ''' </summary>
    ''' <param name="sRoleId">角色</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Avril 2012/04/11 Add
    ''' </remarks>
    Function loadFunctionCodeListById(ByVal sRoleId As String) As Boolean
        Try
            Dim strSql As String = "SELECT  DISTINCT" & _
                                   "    SY_FUNCTION_CODE.FUCNAME,TempFunctionList.ID ," & _
                                   "   TempFunctionList.SYSID, TempFunctionList.SUBSYSID, " & _
                                   "    SY_FUNCTION_CODE.FUNCCODE,, " & _
                                   "  SY_FUNCTION_CODE.FUCURL +'?FUNCCODE'+ Convert(varchar, SY_FUNCTION_CODE.FUNCCODE) +'&HOFLAG='+ Convert(varchar,HOFLAG) as FUCURL " & _
                                   "  SY_FUNCTION_CODE.SORTCTRL " & _
                                   " from  SY_FUNCTION_CODE join TempFunctionList " & _
                                   "  on SY_FUNCTION_CODE.PARENT = TempFunctionList.FUNCCODE " & _
                                   " Join SY_REL_ROLE_FUNCTION on SY_REL_ROLE_FUNCTION. FUNCCODE = SY_FUNCTION_CODE. FUNCCODE   " & _
                                   " WHERE SY_FUNCTION_CODE .DISABLED=0 " & _
                                   " AND SY_REL_ROLE_FUNCTION.ROLEID in (" & ProviderFactory.PositionPara & "ROLEID" & ")" & _
                                   " and SY_FUNCTION_CODE. FUNCCODE not in(select FUNCCODE from TempFunctionList) "

            Dim para(0) As IDbDataParameter
            para(0) = ProviderFactory.CreateDataParameter("ROLEID", sRoleId)

            Return MyBase.loadBySQL(strSql, para)
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function


    ''' <summary>
    ''' 根據功能編號查詢資料
    ''' </summary>
    ''' <param name="sFunccode">功能編號</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Avril 2012/05/08 Add
    ''' </remarks>
    Function loadDataByFunccode(ByVal sFunccode As String) As Boolean
        Try
            Dim strSql As String = "SELECT  " & _
                                   "   *" & _
                                   " FROM  SY_FUNCTION_CODE " & _
                                   " WHERE FUNCCODE IN( " & sFunccode & ") "

            Return MyBase.loadBySQL(strSql)
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 查詢資料
    ''' </summary>
    ''' <param name="sParent">父節點</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Add by Avril 2012/05/23
    ''' </remarks>
    Function loadDataByType(ByVal sParent As String) As Boolean
        Try
            Dim strSql = "SELECT " & _
                         "    * FROM SY_FUNCTION_CODE " & _
                         " WHERE " & _
                         "   HOFLAG <> '0'  " & _
                         " AND PARENT = " & ProviderFactory.PositionPara & "PARENT"

            Dim para(0) As IDbDataParameter
            para(0) = ProviderFactory.CreateDataParameter("PARENT", sParent)

            Return MyBase.loadBySQL(strSql, para)
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 查詢資料
    ''' </summary>
    ''' <param name="sParent">父節點</param>
    ''' <param name="sFuncCodeList">交易集合</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Add by Avril 2012/06/07
    ''' </remarks>
    Function loadDataByParentAndCode(ByVal sParent As String, ByVal sFuncCodeList As String) As Boolean
        Try
            Dim strSql = "SELECT " & _
                         "    * FROM SY_FUNCTION_CODE " & _
                         " WHERE " & _
                         "    1=1 " & _
                         " AND PARENT = " & ProviderFactory.PositionPara & "PARENT" & _
                         " AND FUNCCODE IN (" & sFuncCodeList & ")"

            Dim para(0) As IDbDataParameter
            para(0) = ProviderFactory.CreateDataParameter("PARENT", sParent)

            Return MyBase.loadBySQL(strSql, para)
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 根據父節點查詢交易
    ''' </summary>
    ''' <param name="sParentList">父節點集合</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Add by Avril 2012/06/07
    ''' </remarks>
    Function loadDataByParentList(ByVal sParentList As String) As Boolean
        Try
            Dim strSql = "SELECT " & _
                         "    * FROM SY_FUNCTION_CODE " & _
                         " WHERE " & _
                         "    1=1 " & _
                         " AND FUNCCODE IN (" & sParentList & ")"

            Return MyBase.loadBySQL(strSql)
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 取得交易建立左邊目錄樹的資料
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function genAllFunctionCodeList() As DataTable
        Try
            Dim strSql As String = "WITH rc(FUNCCODE,FUCNAME,FUCURL,PARENTID,DISABLED,SORTCTRL,HOFLAG,SerialNum ,SerialStr,Level) AS" & _
                                    "(SELECT FUNCCODE,FUCNAME,FUCURL,PARENT,DISABLED,SORTCTRL,HOFLAG," & _
                                    "     CAST(row_number() over (order by FUNCCODE) as int) SerialNum ," & _
                                    "     CAST(row_number() over (order by FUNCCODE) as varchar(255)) SerialStr," & _
                                    "     0 Level" & _
                                    "    FROM SY_FUNCTION_CODE a" & _
                                    "   WHERE PARENT = 0" & _
                                    "  UNION ALL" & _
                                    "  SELECT p.FUNCCODE," & _
                                    "         p.FUCNAME," & _
                                    "         p.FUCURL," & _
                                    "         p.PARENT," & _
                                    "         p.DISABLED," & _
                                    "         p.SORTCTRL," & _
                                    "         p.HOFLAG," & _
                                    "         CAST(rc.SerialNum + 1 as int)," & _
                                    "         CAST(RTrim(rc.SerialStr) + '-' + CAST(row_number() over (order by p.FUNCCODE) as varchar(10)) as varchar(255))," & _
                                    "              rc.Level + 1" & _
                                    "    FROM SY_FUNCTION_CODE P, rc" & _
                                    "   WHERE rc.FUNCCODE = P.PARENT" & _
                                    "     and p.PARENT > 0)" & _
                                    "select  *" & _
                                    "  from rc " & _
                                    " order by FUNCCODE,SerialStr  OPTION(MAXRECURSION 5)"



            If Me.loadBySQLOnlyDs(strSql) > 0 Then
                Return Me.getCurrentDataSet.Tables(0)
            End If
            Return Nothing
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Function genAllFunctionNonHoFlag0() As DataTable
        Try
            Dim strSql As String = "WITH rc(FUNCCODE,FUCNAME,FUCURL,PARENTID,DISABLED,SORTCTRL,HOFLAG,SerialNum ,SerialStr,Level) AS" & _
                                    "(SELECT FUNCCODE,FUCNAME,FUCURL,PARENT,DISABLED,SORTCTRL,HOFLAG," & _
                                    "     CAST(row_number() over (order by FUNCCODE) as int) SerialNum ," & _
                                    "     CAST(row_number() over (order by FUNCCODE) as varchar(255)) SerialStr," & _
                                    "     0 Level" & _
                                    "    FROM SY_FUNCTION_CODE a" & _
                                    "   WHERE PARENT = 0 AND [DISABLED] =0" & _
                                    "  UNION ALL" & _
                                    "  SELECT p.FUNCCODE," & _
                                    "         p.FUCNAME," & _
                                    "         p.FUCURL," & _
                                    "         p.PARENT," & _
                                    "         p.DISABLED," & _
                                    "         p.SORTCTRL," & _
                                    "         p.HOFLAG," & _
                                    "         CAST(rc.SerialNum + 1 as int)," & _
                                    "         CAST(RTrim(rc.SerialStr) + '-' + CAST(row_number() over (order by p.FUNCCODE) as varchar(10)) as varchar(255))," & _
                                    "              rc.Level + 1" & _
                                    "    FROM SY_FUNCTION_CODE P, rc" & _
                                    "   WHERE rc.FUNCCODE = P.PARENT" & _
                                    "     and p.PARENT > 0 AND P.DISABLED = 0)" & _
                                    "select  *" & _
                                    "  from rc where hoflag<>'0' " & _
                                    " order by FUNCCODE,SerialStr  OPTION(MAXRECURSION 5)"



            If Me.loadBySQLOnlyDs(strSql) > 0 Then
                Return Me.getCurrentDataSet.Tables(0)
            End If
            Return Nothing
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
