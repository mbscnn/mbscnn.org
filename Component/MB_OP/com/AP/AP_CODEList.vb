Imports com.Azion.NET.VB
Public Class AP_CODEList
    Inherits BosList

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.New("AP_CODE", dbManager)
    End Sub

    Overrides Function newBos() As BosBase
        Dim bos As AP_CODE = New AP_CODE(MyBase.getDatabaseManager)
        Return bos
    End Function

    ''' <summary>
    ''' 根據[子系統代號]取得資料
    ''' </summary>
    ''' <param name="sSubSysID">子系統代號</param>
    ''' <returns>Integer 取得筆數</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' 	[Willy]	2009/11/4	Created
    ''' 	[Willy]	2009/11/12	Modified
    ''' 	[Willy]	2009/11/30	Modified
    ''' </history>
    Function loadBySubSysID(ByVal sSubSysID As String) As Integer

        Try
            APinitTable.initApCode(Me.getDatabaseManager)

            Dim rows() As DataRow
            rows = APinitTable.dtApCodes.Select(String.Format("SUBSYSID='{0}'", sSubSysID))

            For Each row As DataRow In rows
                Dim bosbsse As AP_CODE = Me.newBos
                For Each col As DataColumn In row.Table.Columns
                    bosbsse.setAttribute(col.ColumnName, row(col.ColumnName))
                Next
                add(bosbsse)
            Next

            Return size()
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 根據[父類別代號][子系統代號]取得資料
    ''' </summary>
    ''' <param name="sSubSysID">子系統代號</param>
    ''' <param name="iUpCode">父類別代號</param>
    ''' <returns>Integer 取得筆數</returns>
    ''' <remarks>
    ''' sSubSysID="99" 代號全部的子系統
    ''' </remarks> 
    ''' <history>
    ''' 	[Willy]	2009/11/5 Created
    ''' 	[Willy]	2009/11/12	Modified
    ''' 	[Willy]	2009/11/30	Modified
    ''' </history>
    Function loadByUpCode(ByVal iUpCode As Integer, Optional ByVal sSubSysID As String = "99", Optional ByVal sDisabled As String = "") As Integer

        Try
            APinitTable.initApCode(Me.getDatabaseManager)
            Dim rows() As DataRow
            Dim dt As DataTable = APinitTable.dtApCodes.Clone

            If sSubSysID.Equals("99") Then

                rows = APinitTable.dtApCodes.Select(String.Format("UPCODE={0}", iUpCode))

                For Each row As DataRow In rows
                    Dim bosbsse As AP_CODE = Me.newBos
                    For Each col As DataColumn In row.Table.Columns
                        bosbsse.setAttribute(col.ColumnName, row(col.ColumnName))
                    Next
                    add(bosbsse)
                    dt.ImportRow(row)
                Next

            ElseIf sDisabled = "" Then
                rows = APinitTable.dtApCodes.Select(String.Format("SUBSYSID='{0}' and UPCODE={1}", sSubSysID, iUpCode))

                For Each row As DataRow In rows
                    Dim bosbsse As AP_CODE = Me.newBos
                    For Each col As DataColumn In row.Table.Columns
                        bosbsse.setAttribute(col.ColumnName, row(col.ColumnName))
                    Next
                    add(bosbsse)
                    dt.ImportRow(row)
                Next

            Else
                rows = APinitTable.dtApCodes.Select(String.Format("SUBSYSID='{0}' and UPCODE={1} and DISABLED={2}", sSubSysID, iUpCode, sDisabled))

                For Each row As DataRow In rows
                    Dim bosbsse As AP_CODE = Me.newBos
                    For Each col As DataColumn In row.Table.Columns
                        bosbsse.setAttribute(col.ColumnName, row(col.ColumnName))
                    Next
                    add(bosbsse)
                    dt.ImportRow(row)
                Next
            End If

            Dim ds As New DataSet
            ds.Tables.Add(dt)
            Me.setCurrentDataSet(ds)

            Return Me.size
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 根據[文字]取得資料
    ''' </summary>
    ''' <param name="sText">文字</param>
    ''' <returns>Integer 取得筆數</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' 	[Willy]	2009/11/5   Created
    ''' 	[Willy]	2009/11/9   Modified
    ''' 	[Willy]	2009/11/12	Modified
    ''' 	[Willy]	2009/11/30	Modified
    ''' </history>
    Function loadByText(ByVal sText As String) As Integer

        Try
            APinitTable.initApCode(Me.getDatabaseManager)

            Dim rows() As DataRow
            rows = APinitTable.dtApCodes.Select(String.Format("TEXT='{0}'", sText))

            For Each row As DataRow In rows
                Dim bosbsse As AP_CODE = Me.newBos
                For Each col As DataColumn In row.Table.Columns
                    bosbsse.setAttribute(col.ColumnName, row(col.ColumnName))
                Next
                add(bosbsse)
            Next

            Return size()

        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 根據[子系統代號][項目序號]取得資料
    ''' </summary>
    ''' <param name="_SubSysID">子系統代號</param>
    ''' <param name="sCodeID">項目序號</param>
    ''' <returns>Integer 取得筆數</returns>
    ''' <remarks>
    ''' for CoMntItemCode_01_V01's DataGrid
    ''' </remarks> 
    ''' <history>
    ''' 	[Willy]	2009/11/6   Created
    ''' 	[Willy]	2009/11/9   Modified
    ''' 	[Willy]	2009/11/12	Modified
    ''' 	[Willy]	2009/11/16	Modified
    ''' 	[Willy]	2009/11/30	Modified
    ''' </history>
    Function loadByCoMntItemCode_01_V01(ByVal _SubSysID As String, ByVal sCodeID As String) As Integer
        Dim sSubSysID As String
        'Dim sSql As String = " SELECT * FROM AP_CODE WHERE " & _
        '                    " SUBSYSID=" & ProviderFactory.PositionPara & "SUBSYSID           " & _
        '                    " AND CODEID=" & ProviderFactory.PositionPara & "CODEID           " & _
        '                    " UNION                       " & _
        '                    " SELECT * FROM AP_CODE WHERE " & _
        '                    " SUBSYSID=" & ProviderFactory.PositionPara & "SUBSYSID           " & _
        '                    " AND UPCODE=" & ProviderFactory.PositionPara & "CODEID           " & _
        '                    "  ORDER BY 7,6,2             "
        Try
            APinitTable.initApCode(Me.getDatabaseManager)
            '當子系統代號等於全系統時，選擇sCodeID的子系統當條件
            If _SubSysID.Equals("99") Then
                Dim apCode As New AP_CODE(Me.getDatabaseManager)
                apCode.loadByPK(sCodeID)
                sSubSysID = apCode.getString("SUBSYSID")
            Else
                sSubSysID = _SubSysID
            End If

            Dim rows() As DataRow
            rows = APinitTable.dtApCodes.Select(String.Format("SUBSYSID='{0}' and CODEID={1}", _SubSysID, sCodeID))

            For Each row As DataRow In rows
                Dim bosbsse As AP_CODE = Me.newBos
                For Each col As DataColumn In row.Table.Columns
                    bosbsse.setAttribute(col.ColumnName, row(col.ColumnName))
                Next
                add(bosbsse)
            Next


            rows = APinitTable.dtApCodes.Select(String.Format("SUBSYSID='{0}' and UPCODE={1}", _SubSysID, sCodeID))

            For Each row As DataRow In rows
                Dim bosbsse As AP_CODE = Me.newBos
                For Each col As DataColumn In row.Table.Columns
                    bosbsse.setAttribute(col.ColumnName, row(col.ColumnName))
                Next
                add(bosbsse)
            Next

            Return Me.size
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' 根據[搜尋字串]比對[TEXT],[VALUE],[NOTE]欄位中是否有該字串取得資料
    ''' </summary>
    ''' <param name="sQuery">搜尋字串</param>
    ''' <returns>Integer 取得筆數</returns>Throw
    ''' <remarks>
    ''' for CoMntItemCode_01_V01's Advance Search
    ''' </remarks> 
    ''' <history>
    ''' 	[Willy]	2009/11/13   Created
    ''' 	[Willy]	2009/11/30	Modified
    ''' </history>
    Function loadForAdvanceSearch(ByVal sQuery As String) As Integer
        Try
            APinitTable.initApCode(Me.getDatabaseManager)

            Dim rows() As DataRow
            rows = APinitTable.dtApCodes.Select(String.Format("TEXT like '%{0}%' or VALUE like '%{1}%' or NOTE like '%{2}%'", sQuery, sQuery, sQuery))

            For Each row As DataRow In rows
                Dim bosbsse As AP_CODE = Me.newBos
                For Each col As DataColumn In row.Table.Columns
                    bosbsse.setAttribute(col.ColumnName, row(col.ColumnName))
                Next
                add(bosbsse)
            Next

            Return Me.size
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' 根據"UPCODE","會員帳號"取得有權使用的頁籤資料
    ''' </summary>
    ''' <param name="iUPCODE">UPCODE</param>
    ''' <param name="sMB_ACCT">會員帳號</param>
    ''' <returns>Integer 取得筆數</returns>Throw
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' 	Ted 2014/9/18   Created
    ''' </history>
    Public Function Load_TAB_AUTH(ByVal iUPCODE As Integer, ByVal sMB_ACCT As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT A.* " &
                     "  FROM AP_CODE A " &
                     " WHERE     A.UPCODE = " & ProviderFactory.PositionPara & "UPCODE " &
                     "       AND (   A.LEVEL <=1 " &
                     "            OR (    A.LEVEL > 1 " &
                     "                AND EXISTS " &
                     "                       (SELECT * " &
                     "                          FROM MB_FNCAUTH " &
                     "                         WHERE MB_ACCT = " & ProviderFactory.PositionPara & "MB_ACCT " &
                     "                              AND MB_UPCODE=A.UPCODE AND MB_FUCCODE = A.VALUE)))"

            Dim paras(1) As IDbDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("UPCODE", iUPCODE)

            paras(1) = ProviderFactory.CreateDataParameter("MB_ACCT", sMB_ACCT)

            Return Me.loadBySQLOnlyDs(sqlStr, paras)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

#Region "Rachel Function"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 根據[父類別代號]取得資料
    ''' </summary>
    ''' <param name="sUpCode">父類別代號</param>
    ''' <returns>Integer 取得筆數</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' 	[Rachel]	2009/11/16 Created
    ''' </history>

    Function loadByUpCodeOrderBySortno(ByVal sUpcode As String) As Integer

        If Utility.isValidateData(sUpcode) Then
            Try
                Dim sSQL As String = "select * from AP_CODE where UPCODE=" & ProviderFactory.PositionPara & "UPCODE and DISABLED=1 order by SORTNO "

                Dim paras(0) As System.Data.IDbDataParameter
                paras(0) = ProviderFactory.CreateDataParameter("UPCODE", sUpcode)

                Return MyBase.loadBySQL(sSQL, paras)

            Catch ex As ProviderException
                Throw
            Catch ex As BosException
                Throw
            Catch ex As Exception
                Throw
            End Try

        End If


    End Function


#End Region

#Region "nick Function"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 根據[父類別代號]取得資料
    ''' </summary>
    ''' <param name="sUpCode">父類別代號</param>
    ''' <returns>Integer 取得筆數</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' 	[Nick]	2010/08/13 Created
    ''' </history>
    Function getVALUEandTEXTbyUPCODE(ByVal sUpcode As String) As Integer

        If Utility.isValidateData(sUpcode) Then
            Try
                Dim sSQL As String = "select VALUE,TEXT from AP_CODE where UPCODE=" & ProviderFactory.PositionPara & "UPCODE and DISABLED = 1 order by SORTNO "

                Dim paras(0) As System.Data.IDbDataParameter
                paras(0) = ProviderFactory.CreateDataParameter("UPCODE", sUpcode)

                Return MyBase.loadBySQL(sSQL, paras)

            Catch ex As ProviderException
                Throw
            Catch ex As BosException
                Throw
            Catch ex As Exception
                Throw
            End Try

        End If
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 根據[父類別代號],[Value]取得資料
    ''' </summary>
    ''' <param name="sUpCode">父類別代號</param>
    ''' <param name="sValue">Value</param>
    ''' <returns>Integer 取得筆數</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' 	[Nick]	2010/11/05 Created
    ''' </history>
    Function LoadbyUPCODEandVALUE(ByVal sUpcode As String, ByVal sValue As String) As Integer

        If Utility.isValidateData(sUpcode) Then
            Try
                Dim sSQL As String = "select * from AP_CODE where UPCODE=" & ProviderFactory.PositionPara & "UPCODE "
                sSQL &= " and value = " & ProviderFactory.PositionPara & "VALUE "

                Dim paras(1) As System.Data.IDbDataParameter
                paras(0) = ProviderFactory.CreateDataParameter("UPCODE", sUpcode)
                paras(1) = ProviderFactory.CreateDataParameter("VALUE", sValue)

                Return MyBase.loadBySQL(sSQL, paras)

            Catch ex As ProviderException
                Throw
            Catch ex As BosException
                Throw
            Catch ex As Exception
                Throw
            End Try

        End If
    End Function

#End Region

#Region "Thera Function"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 以父類別(Upcode)查詢此項目值與文字
    ''' </summary>
    ''' <param name="sUpCode">父類別代號</param>
    ''' <returns>Integer 取得筆數</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    '''  2010/10/14 CREATE
    ''' </history>
    Function getVALUIRTEXTbyUPCODE(ByVal sUpcode As String) As Integer

        If Utility.isValidateData(sUpcode) Then
            Try
                Dim sSQL As String = "select VALUE,TEXT from AP_CODE where UPCODE=" & ProviderFactory.PositionPara & "UPCODE order by SORTNO ASC  "

                Dim paras(0) As System.Data.IDbDataParameter
                paras(0) = ProviderFactory.CreateDataParameter("UPCODE", sUpcode)

                Return MyBase.loadBySQL(sSQL, paras)

            Catch ex As ProviderException
                Throw
            Catch ex As BosException
                Throw
            Catch ex As Exception
                Throw
            End Try

        End If
    End Function

    ''' <summary>
    ''' 傳入arraylist查詢出此範圍的項目資料
    ''' </summary>
    ''' <param name="aUpCode"></param>
    ''' <returns></returns>
    ''' 
    ''' <remarks></remarks>
    Function getUpCode(ByVal aUpCode As ArrayList)
        Try
            Dim paras(14) As System.Data.IDbDataParameter '先設全部12個參數
            Dim i As Integer = 0 '記錄位置
            Dim m_sSql As String = ""
            Dim sb As New System.Text.StringBuilder
            Dim strSql As String = ""
            Dim sb1 As New System.Text.StringBuilder

            m_sSql = " SELECT  A.*  " & _
                             "    FROM AP_CODE A "

            If Not aUpCode.Count = 0 Then
                For k As Integer = 0 To aUpCode.Count - 1
                    If aUpCode.Item(k) <> "" Then
                        sb1.Append(ProviderFactory.PositionPara).Append("UPCODE").Append(k).Append(",")
                        paras(i) = ProviderFactory.CreateDataParameter(ProviderFactory.PositionPara + "UPCODE" & k.ToString, aUpCode.Item(k))
                        i += 1
                    End If
                Next

                sb1.Remove(sb1.Length - 1, 1)
                strSql = sb1.ToString
                m_sSql += " WHERE  A.UPCODE IN ( " & strSql & ")"
            End If

            Dim paras1(i) As System.Data.IDbDataParameter '正式要送出的參數

            For j As Integer = 0 To i
                paras1(j) = paras(j)
            Next

            Return MyBase.loadBySQLOnlyDs(m_sSql, paras1)

        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function
#End Region

#Region "Joe Function"

    Public Function loadDataByUpCode(sUpCode As String) As Integer
        If Not Utility.isValidateData(sUpCode) Then Return 0
        '=========================================================================================
        Dim paras(0) As IDbDataParameter
        Dim sSQL As String

        sSQL = " select VALUE, TEXT from AP_CODE" & _
               " where DISABLED='1' and UPCODE=" & ProviderFactory.PositionPara & "UPCODE" & _
               " order by SORTNO asc"
        paras(0) = ProviderFactory.CreateDataParameter("UPCODE", sUpCode)

        Return MyBase.loadBySQLOnlyDs(sSQL, paras)
    End Function
#End Region

#Region "濟南昱勝添加"
    ''' <summary>
    ''' 獲取財報幣別單位
    ''' </summary>
    ''' <param name="sUPCODE">1759</param>
    ''' <returns>Boolean</returns>
    ''' <remarks></remarks>
    '''  <history>
    ''' 	[Alicia]	2011/12/29 Created
    ''' </history>
    Public Function getCurUnit(ByVal sUPCODE As Integer) As Integer
        Try
            Dim sqlStr As String = "Select VALUE,TEXT,SORTNO" &
                                    " FROM AP_CODE" &
                                    " WHERE UPCODE =" & ProviderFactory.PositionPara & "UPCODE "

            Dim para(0) As System.Data.IDbDataParameter
            para(0) = ProviderFactory.CreateDataParameter("UPCODE", sUPCODE)

            Return MyBase.loadBySQL(sqlStr, para)
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 取得[父類別代號]及不等於[Value]的資料
    ''' </summary>
    ''' <param name="sUpCode">父類別代號</param>
    ''' <param name="sValue">Value</param>
    ''' <returns>Integer 取得筆數</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' 	[Nick]	2010/11/05 Created
    ''' </history>
    Function loadByUpcodeNotValue(ByVal sUpcode As String, ByVal sValue As String) As Integer

        If Utility.isValidateData(sUpcode) Then
            Try
                Dim sSQL As String = "select * from AP_CODE where UPCODE=" & ProviderFactory.PositionPara & "UPCODE "
                sSQL &= " and value <> " & ProviderFactory.PositionPara & "VALUE "

                Dim paras(1) As System.Data.IDbDataParameter
                paras(0) = ProviderFactory.CreateDataParameter("UPCODE", sUpcode)
                paras(1) = ProviderFactory.CreateDataParameter("VALUE", sValue)

                Return MyBase.loadBySQL(sSQL, paras)

            Catch ex As ProviderException
                Throw ex
            Catch ex As BosException
                Throw ex
            Catch ex As Exception
                Throw ex
            End Try

        End If
    End Function

    ''' <summary>
    ''' 根據codeid查詢相關資料
    ''' </summary>
    ''' <param name="sCodeId">codeid</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Add by Avril 2012/03/28
    ''' </remarks>
    Function loadDataByCodeId(ByVal sCodeId As String) As Boolean
        Try
            Dim sSQL As String = "SELECT * FROM AP_CODE" &
                                   " WHERE CODEID = " & ProviderFactory.PositionPara & "CODEID "


            Dim para(0) As System.Data.IDbDataParameter
            para(0) = ProviderFactory.CreateDataParameter("CODEID", sCodeId)

            Return MyBase.loadBySQL(sSQL, para)
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 根據Upcode查詢維護角色信息
    ''' </summary>
    ''' <param name="sUpcode">2283</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Zack  2012-06-01 Create
    ''' </remarks>
    Function loadBuUpCodeValue(ByVal sUpcode As String) As Integer
        Try
            Dim sSQL As String = "SELECT * FROM AP_CODE WHERE UPCODE =" & ProviderFactory.PositionPara & "UPCODE" &
                                 " AND VALUE != 1"

            Dim para(0) As System.Data.IDbDataParameter
            para(0) = ProviderFactory.CreateDataParameter("UPCODE", sUpcode)

            Return MyBase.loadBySQL(sSQL, para)
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

End Class
