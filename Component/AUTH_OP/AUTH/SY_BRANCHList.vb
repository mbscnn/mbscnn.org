Imports com.Azion.NET.VB

Public Class SY_BRANCHList
    Inherits BosList

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.New("SY_BRANCH", dbManager)
    End Sub

    Overrides Function newBos() As BosBase
        Dim bos As SY_BRANCH = New SY_BRANCH(MyBase.getDatabaseManager)
        Return bos
    End Function


#Region "濟南昱勝添加"
#Region "Lake Function"
    ''' <summary>
    ''' 根據上級部門編號查詢子部門
    ''' </summary>
    ''' <param name="sParentId">上級部門編號</param>
    ''' <param name="sBrid">分行單位代碼；“T”:組織作業維護顯示所有資料</param>
    ''' <param name="sHoFlag">"3":單位；"1":總管理處</param>
    ''' <returns></returns>
    ''' <remarks>[Lake] 2012/05/16 Created</remarks>
    Public Function getBrancByParent(ByVal sParentId As String, ByVal sBrid As String, ByVal sHoFlag As String) As Boolean
        Try
            Dim sSql As String = "SELECT " & _
                                 "    (CONVERT(VARCHAR,a.BRA_DEPNO) + ';' + CONVERT(VARCHAR,BRID) + ';' + BRCNAME + ';' " & _
                                 "    + CASE WHEN EPSDEP IS NULL THEN 'WRITE' ELSE EPSDEP END) AS BRA_DEPNO " & _
                                 "   ,(CONVERT(VARCHAR,BRID) + '(' + CONVERT(VARCHAR,a.BRA_DEPNO) + ')' + '  ' + BRCNAME + " & _
                                 "    CASE " & _
                                 "    WHEN EPSDEP IS NULL OR EPSDEP = '' THEN '' " & _
                                 "    ELSE '('+EPSDEP+')' END) AS BRCNAME " & _
                                 "   ,PARENT " & _
                                 "FROM " & _
                                 "   SY_BRANCH a " & _
                                 "WHERE  " & _
                                 "   PARENT = " & ProviderFactory.PositionPara & "PARENT "

            '“T”:組織作業維護顯示所有資料
            If sBrid <> "T" Then
                sSql = sSql & "AND DISABLED = 0"
            End If

            Dim paras(0) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("PARENT", sParentId)

            If sHoFlag <> "1" Then
                sSql = sSql & "AND BRID = " & ProviderFactory.PositionPara & "BRID"

                ReDim Preserve paras(1)
                paras(1) = ProviderFactory.CreateDataParameter("BRID", sBrid)
            End If

            sSql &= " order by brid,BRA_DEPNO"
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
    ''' 根據上級部門編號查詢子部門
    ''' </summary>
    ''' <param name="sParentId">上級部門編號</param>
    ''' <param name="sBrid">分行單位代碼</param>
    ''' <param name="sHoFlag">"3":單位；"1":總管理處</param>
    ''' <param name="sShowAll">“T”:組織作業維護顯示所有資料</param>
    ''' <returns></returns>
    ''' <remarks>[Lake] 2012/05/16 Created</remarks>
    Public Function getBrancByParentForDdl(ByVal sParentId As String, ByVal sBrid As String, ByVal sHoFlag As String, Optional ByVal sShowAll As String = "") As Boolean
        Try
            Dim sSql As String = "SELECT " & _
                                 "    (CONVERT(VARCHAR,a.BRA_DEPNO) + ';' + CONVERT(VARCHAR,BRID) + ';' + BRCNAME + ';' " & _
                                 "    + CASE WHEN EPSDEP IS NULL THEN 'WRITE' ELSE EPSDEP END) AS BRA_DEPNO " & _
                                 "   ,(CONVERT(VARCHAR,BRID) + ' ' + '(' + CONVERT(VARCHAR,a.BRA_DEPNO) + ')' + '  ' + BRCNAME) AS BRCNAME " & _
                                 "   ,PARENT " & _
                                 "FROM " & _
                                 "   SY_BRANCH a " & _
                                 "WHERE " & _
                                 "   a.PARENT = " & ProviderFactory.PositionPara & "PARENT "

            '“T”:組織作業維護顯示所有資料
            If sShowAll <> "T" Then
                sSql = sSql & "AND DISABLED = 0"
            End If

            Dim paras(0) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("PARENT", sParentId)

            If sHoFlag <> "1" Then
                ReDim Preserve paras(1)
                sSql = sSql & "AND BRID = " & ProviderFactory.PositionPara & "BRID"
                paras(1) = ProviderFactory.CreateDataParameter("BRID", sBrid)
            End If

            sSql &= " order by brid,BRA_DEPNO"
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
    ''' 取得一級單位下所有下屬單位集合
    ''' </summary>
    ''' <param name="sBrid">單位代碼</param>
    ''' <param name="sBraDepNo">部門代碼</param>
    ''' <returns></returns>
    ''' <remarks>[Lake] 2012/06/26 Created</remarks>
    Public Function getAllSubBraDepNo(ByVal sBrid As String, ByVal sBraDepNo As String) As Boolean
        Try
            Dim sSql As String = "SELECT " & _
                                 "    (CONVERT(VARCHAR,a.BRA_DEPNO) + ';' + CONVERT(VARCHAR,BRID) + ';' + BRCNAME + ';' " & _
                                 "    + CASE WHEN EPSDEP IS NULL THEN 'WRITE' ELSE EPSDEP END) AS BRA_DEPNO " & _
                                 "   ,(CONVERT(VARCHAR,BRID) + ' (' + CONVERT(VARCHAR,a.BRA_DEPNO) + ')' + '  ' + BRCNAME + " & _
                                 "    CASE " & _
                                 "    WHEN EPSDEP IS NULL OR EPSDEP = '' THEN '' " & _
                                 "    ELSE '('+EPSDEP+')' END) AS BRCNAME " & _
                                 "   ,PARENT " & _
                                 "FROM SY_BRANCH a " & _
                                 "WHERE BRID = " & ProviderFactory.PositionPara & "BRID " & _
                                 "AND a.BRA_DEPNO != " & ProviderFactory.PositionPara & "BRADEPNO"

            Dim paras(1) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("BRID", sBrid)
            paras(1) = ProviderFactory.CreateDataParameter("BRADEPNO", sBraDepNo)

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
    ''' 根據上級單位代碼取得部門清單
    ''' </summary>
    ''' <param name="sParent">上級單位代碼</param>
    ''' <returns></returns>
    ''' <remarks>[lake] 2012/06/28 Created</remarks>
    Public Function getBranch(ByVal sParent As String) As Boolean
        Try
            Dim sSql As String = "SELECT * FROM SY_BRANCH WHERE PARENT = " & ProviderFactory.PositionPara & "PARENT"

            Dim paras As IDbDataParameter = ProviderFactory.CreateDataParameter("PARENT", sParent)

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
    ''' 根據傳入的上級部門編號及啟用狀態，屬性查詢部門集合
    ''' </summary>
    ''' <param name="sFlag">代理人單位或被代理人單位</param>
    ''' <param name="sBrid">單位代碼</param>
    ''' <param name="sHoflag">屬性</param>
    ''' <returns></returns>
    ''' <remarks>[Lake] 2012/06/26 Created</remarks>
    Public Function getBranchByParas(ByVal sFlag As String, ByVal sHoflag As String, ByVal sBrid As String) As Boolean
        Try
            Dim sSql As String = "SELECT BRID + ' ' + '(' + CAST(BRA_DEPNO AS VARCHAR) + ')' + '  ' + BRCNAME AS BRCNAME,BRA_DEPNO " & _
                                 " FROM SY_BRANCH WHERE PARENT = 0 AND DISABLED = 0 "


            ' 若查詢被代理人部門
            If sFlag = "Y" AndAlso sHoflag <> "1" Then
                sSql = sSql & "AND BRID = " & ProviderFactory.PositionPara & "BRID"

                Dim paras(0) As IDbDataParameter
                paras(0) = ProviderFactory.CreateDataParameter("BRID", sBrid)

                Return Me.loadBySQL(sSql, paras) > 0
            Else
                Return Me.loadBySQL(sSql) > 0
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

#Region "Titan Function"
    Function loadBranchInfo(ByVal sBrId As String, ByVal sHoflag As String, Optional ByRef sDisable As String = "0", Optional ByRef iLevel As Integer = 5) As DataTable
        Dim sSQL As String = "WITH rc(BRA_DEPNO, BRID, PARENTID, BRCNAME, BRCADDR, BRATEL, BRTEL," & _
                        "BRFAX, BRENAME, BREADDR, DISABLED, BRCCITY, BRCAREA, HOFLAG," & _
                        "EPSDEP, SerialNum, SerialStr, Level) AS" & _
                        " (SELECT BRA_DEPNO," & _
                        "         BRID," & _
                        "         PARENT," & _
                        "         BRCNAME," & _
                        "         BRCADDR," & _
                        "         BRATEL," & _
                        "         BRTEL," & _
                        "         BRFAX," & _
                        "         BRENAME," & _
                        "         BREADDR," & _
                        "         DISABLED," & _
                        "         BRCCITY," & _
                        "         BRCAREA," & _
                        "         HOFLAG," & _
                        "         EPSDEP," & _
                        "         CAST(row_number() over(order by BRA_DEPNO) as int) SerialNum," & _
                        "         CAST(row_number() over(order by BRA_DEPNO) as varchar(255)) SerialStr," & _
                        "         0 Level" & _
                        "    FROM SY_BRANCH a" & _
                        "   WHERE PARENT = 0" & _
                        "  UNION ALL" & _
                        "  SELECT p.BRA_DEPNO," & _
                        "         p.BRID," & _
                        "         p.PARENT," & _
                        "         p.BRCNAME," & _
                        "         p.BRCADDR," & _
                        "         p.BRATEL," & _
                        "         p.BRTEL," & _
                        "         p.BRFAX," & _
                        "         p.BRENAME," & _
                        "         p.BREADDR," & _
                        "         p.DISABLED," & _
                        "         p.BRCCITY," & _
                        "         p.BRCAREA," & _
                        "         p.HOFLAG," & _
                        "         p.EPSDEP," & _
                        "         CAST(rc.SerialNum + 1 as int)," & _
                        "         CAST(RTrim(rc.SerialStr) + '-' +" & _
                        "              CAST(row_number() over(order by p.BRA_DEPNO) as varchar(10)) as" & _
                        "              varchar(255))," & _
                        "         rc.Level + 1" & _
                        "    FROM SY_BRANCH P, rc" & _
                        "   WHERE rc.BRA_DEPNO = P.PARENT" & _
                        "     and p.PARENT > 0)" & _
                        "select * , name = replicate('　', level) +  brid + ' ' +'(' + cast(bra_depno as varchar) + ')' + '  ' + rtrim (rc.brcname)" & _
                        "  from rc " & _
                        " where BRID = " & ProviderFactory.PositionPara & "BRID" & _
                        " and DISABLED =" & ProviderFactory.PositionPara & "DISABLED" & _
                        " and LEVEL <=" & ProviderFactory.PositionPara & "LEVEL" & _
                        " order by SerialStr OPTION (MAXRECURSION 5)"

        Dim paras(2) As IDbDataParameter
        paras(0) = ProviderFactory.CreateDataParameter("BRID", sBrId)
        paras(1) = ProviderFactory.CreateDataParameter("DISABLED", sDisable)
        paras(2) = ProviderFactory.CreateDataParameter("LEVEL", iLevel)

        If sHoflag = "1" Then
            Return Me.loadAllBranch(sDisable, iLevel)
        Else
            If Me.loadBySQLOnlyDs(sSQL, paras) > 0 Then
                Return Me.getCurrentDataSet.Tables(0)
            End If
        End If
        Return Nothing
    End Function


    Function loadAllBranch(Optional ByRef sDisable As String = "0", Optional ByRef iLevel As Integer = 5) As DataTable

        Dim sSQL As String = "with rc(bra_depno, brid, parentid, brcname, brcaddr, bratel," & _
                                "brtel, brfax, brename, breaddr, disabled, brccity," & _
                                "brcarea, hoflag, epsdep, serialnum, serialstr, level) as" & _
                                " (select bra_depno," & _
                                "         brid," & _
                                "         parent," & _
                                "         brcname," & _
                                "         brcaddr," & _
                                "         bratel," & _
                                "         brtel," & _
                                "         brfax," & _
                                "         brename," & _
                                "         breaddr," & _
                                "         disabled," & _
                                "         brccity," & _
                                "         brcarea," & _
                                "         hoflag," & _
                                "         epsdep," & _
                                "         cast(row_number() over(order by bra_depno) as int) serialnum," & _
                                "         cast(row_number() over(order by bra_depno) as varchar(255)) serialstr," & _
                                "         0 level" & _
                                "    from sy_branch a" & _
                                "   where parent = 0" & _
                                "  union all" & _
                                "  select p.bra_depno," & _
                                "         p.brid," & _
                                "         p.parent," & _
                                "         p.brcname," & _
                                "         p.brcaddr," & _
                                "         p.bratel," & _
                                "         p.brtel," & _
                                "         p.brfax," & _
                                "         p.brename," & _
                                "         p.breaddr," & _
                                "         p.disabled," & _
                                "         p.brccity," & _
                                "         p.brcarea," & _
                                "         p.hoflag," & _
                                "         p.epsdep," & _
                                "         cast(rc.serialnum + 1 as int)," & _
                                "         cast(rtrim(rc.serialstr) + '-' +" & _
                                "              cast(row_number() over(order by p.bra_depno) as varchar(10)) as" & _
                                "              varchar(255))," & _
                                "         rc.level + 1" & _
                                "    from sy_branch p, rc" & _
                                "   where rc.bra_depno = p.parent" & _
                                "     and p.parent > 0)" & _
                                "select *, name = replicate('　', level) + brid + ' ' + '(' + cast(bra_depno as varchar) + ')' + '  ' + rtrim (rc.brcname)" & _
                                "  from rc " & _
                                " where disabled =" & ProviderFactory.PositionPara & "DISABLED" & _
                                " and level <=" & ProviderFactory.PositionPara & "LEVEL" & _
                                " order by brid,SerialStr option (maxrecursion 5)"


        Dim paras(1) As IDbDataParameter
        paras(0) = ProviderFactory.CreateDataParameter("DISABLED", sDisable)
        paras(1) = ProviderFactory.CreateDataParameter("LEVEL", iLevel)

        If Me.loadBySQLOnlyDs(sSQL, paras) > 0 Then
            Return Me.getCurrentDataSet.Tables(0)
        End If

        Return Nothing

    End Function

    ''' <summary>
    ''' 取得目前單位的管理單位
    ''' </summary>
    ''' <param name="sBrid"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getTopDepInfo(ByVal sBrid As String) As Boolean
        Try
            Dim sSQL As String = ""
            Dim paras(1) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("PARENT", 0)
            paras(1) = ProviderFactory.CreateDataParameter("BRID", sBrid)
            Return Me.loadBySQL(paras)
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

#End Region
End Class

