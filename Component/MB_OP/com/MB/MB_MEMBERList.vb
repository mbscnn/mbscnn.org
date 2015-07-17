Imports com.Azion.NET.VB

Public Class MB_MEMBERList
    Inherits BosList

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.New("MB_MEMBER", dbManager)
    End Sub

    Overrides Function newBos() As BosBase
        Dim bos As MB_MEMBER = New MB_MEMBER(MyBase.getDatabaseManager)
        Return bos
    End Function

#Region "Ted Function"
    ''' <summary>
    ''' 根據[手機]取得資料
    ''' </summary>
    ''' <param name="sMB_MOBIL">MB_MOBIL手機</param>
    ''' <returns>integer取得筆數</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Function loadByMB_MOBIL(ByVal sMB_MOBIL As String) As Integer
        Try
            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_MOBIL", sMB_MOBIL)

            Return Me.loadBySQLOnlyDs(para)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' 根據[e-Mail]取得資料
    ''' </summary>
    ''' <param name="sMB_EMAIL">e-Mail</param>
    ''' <returns>integer取得筆數</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Function loadByMB_EMAIL(ByVal sMB_EMAIL As String) As Integer
        Try
            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_EMAIL", sMB_EMAIL)

            Return Me.loadBySQLOnlyDs(para)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' 根據會員編號取得家族資料
    ''' </summary>
    ''' <param name="iMB_MEMSEQ">會員編號</param>
    ''' <returns>integer取得筆數</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Function loadByMB_FAMILY(ByVal iMB_MEMSEQ As Decimal) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT B.* " & _
                     "  FROM MB_FAMILY A, MB_MEMBER B" & _
                     " WHERE A.MB_MEMSEQ = " & ProviderFactory.PositionPara & "MB_MEMSEQ AND A.MB_FAMSEQ = B.MB_MEMSEQ"

            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_MEMSEQ", iMB_MEMSEQ)

            Return Me.loadBySQLOnlyDs(sqlStr, para)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' 根據[電話]取得資料
    ''' </summary>
    ''' <param name="sMB_TEL">MB_TEL電話</param>
    ''' <returns>integer取得筆數</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Function loadByMB_TEL(ByVal sMB_TEL As String) As Integer
        Try
            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_TEL", sMB_TEL)

            Return Me.loadBySQLOnlyDs(para)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' 根據[姓名]取得資料
    ''' </summary>
    ''' <param name="sMB_NAME">MB_NAME姓名</param>
    ''' <returns>integer取得筆數</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Function loadByMB_NAME(ByVal sMB_NAME As String) As Integer
        Try
            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_NAME", sMB_NAME)

            Return Me.loadBySQLOnlyDs(para)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' 根據[姓名]模糊搜尋取得資料
    ''' </summary>
    ''' <param name="sMB_NAME">MB_NAME姓名</param>
    ''' <returns>integer取得筆數</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Function loadByMB_NAME_Like(ByVal sMB_NAME As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "select * from mb_member where mb_name like " & ProviderFactory.PositionPara & "MB_NAME "

            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_NAME", "%" & sMB_NAME & "%")

            Return Me.loadBySQLOnlyDs(sqlStr, para)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' 取得存在EXCEL匯入暫存檔資料
    ''' </summary>
    ''' <returns>integer取得筆數</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Function loadExistsIMP() As DataTable
        Dim DT_MB_MEMBER As New DataTable
        Dim sqlAda As System.Data.Common.DbDataAdapter
        Dim sqlCmd As IDbCommand = ProviderFactory.CreateCommand
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT a.*                                 " & _
                     "  FROM mb_member a                         " & _
                     " WHERE EXISTS                              " & _
                     "          (SELECT *                        " & _
                     "             FROM mb_mbimp                 " & _
                     "            WHERE MB_MEMSEQ = a.MB_MEMSEQ) " & _
                     " ORDER BY MB_MEMSEQ "

            'sqlStr = "SELECT X.*, @x := ifnull(@x, 0) + 1 AS rownum" & _
            '         "  FROM (SELECT a.*" & _
            '         "          FROM mb_member a" & _
            '         "         WHERE EXISTS (SELECT * FROM mb_mbimp WHERE MB_MEMSEQ = a.MB_MEMSEQ)) x," & _
            '         "       (SELECT @x := 0) y" & _
            '         " ORDER BY MB_MEMSEQ"

            'Return Me.loadBySQLOnlyDs(sqlStr)

            Dim objROWNUM As New DataColumn("ROWNUM")
            objROWNUM.DataType = Type.GetType("System.Decimal")
            objROWNUM.AutoIncrement = True
            objROWNUM.AutoIncrementStep = 1
            DT_MB_MEMBER.Columns.Add(objROWNUM)

            sqlCmd.Connection = Me.getDatabaseManager.getConnection
            sqlCmd.CommandText = sqlStr

            sqlAda = ProviderFactory.CreateDataAdapter(sqlCmd)

            sqlAda.Fill(DT_MB_MEMBER)

            Return DT_MB_MEMBER
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' 根據[會員編號]取得資料
    ''' </summary>
    ''' <param name="iMB_MEMSEQ">會員編號</param>
    ''' <returns>integer取得筆數</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Function loadByMB_MEMSEQ(ByVal iMB_MEMSEQ As Decimal) As Integer
        Try
            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_MEMSEQ", iMB_MEMSEQ)

            Return Me.loadBySQLOnlyDs(para)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function Load_MB_NAME_MB_MOBIL(ByVal sMB_NAME As String, ByVal sMB_MOBIL As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT * FROM MB_MEMBER WHERE MB_NAME = " & ProviderFactory.PositionPara & "MB_NAME " & _
                     "  AND MB_MOBIL = " & ProviderFactory.PositionPara & "MB_MOBIL "

            Dim paras(1) As IDbDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("MB_NAME", sMB_NAME)

            paras(1) = ProviderFactory.CreateDataParameter("MB_MOBIL", sMB_MOBIL)

            Return Me.loadBySQLOnlyDs(sqlStr, paras)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function Load_MB_NAME_MB_TEL(ByVal sMB_NAME As String, ByVal sMB_TEL As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT * FROM MB_MEMBER WHERE MB_NAME = " & ProviderFactory.PositionPara & "MB_NAME " & _
                     "  AND MB_TEL = " & ProviderFactory.PositionPara & "MB_TEL "

            Dim paras(1) As IDbDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("MB_NAME", sMB_NAME)

            paras(1) = ProviderFactory.CreateDataParameter("MB_TEL", sMB_TEL)

            Return Me.loadBySQLOnlyDs(sqlStr, paras)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function Load_MB_NAME_MB_EMAIL(ByVal sMB_NAME As String, ByVal sMB_EMAIL As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT * FROM MB_MEMBER WHERE MB_NAME = " & ProviderFactory.PositionPara & "MB_NAME " & _
                     "  AND MB_EMAIL = " & ProviderFactory.PositionPara & "MB_EMAIL "

            Dim paras(1) As IDbDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("MB_NAME", sMB_NAME)

            paras(1) = ProviderFactory.CreateDataParameter("MB_EMAIL", sMB_EMAIL)

            Return Me.loadBySQLOnlyDs(sqlStr, paras)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

#Region "報表查詢"
    Public Function QRY_1(ByVal iSTART As Decimal, ByVal iPageSize As Decimal) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT * FROM MB_MEMBER ORDER BY CREDATE LIMIT " & iSTART & "," & iPageSize

            Return Me.loadBySQLOnlyDs(sqlStr)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function QRY_2(ByVal iSTART As Decimal, ByVal iPageSize As Decimal, ByVal D_BDay As Date, ByVal D_EDay As Date) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT * FROM MB_MEMBER WHERE CREDATE BETWEEN " & ProviderFactory.PositionPara & "D_BDay " & _
                     "  AND " & ProviderFactory.PositionPara & "D_EDay ORDER BY CREDATE LIMIT " & iSTART & "," & iPageSize

            Dim paras(1) As IDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("D_BDay", D_BDay)
            paras(1) = ProviderFactory.CreateDataParameter("D_EDay", D_EDay)

            Return Me.loadBySQLOnlyDs(sqlStr, paras)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function QRY_3(ByVal iSTART As Decimal, ByVal iPageSize As Decimal, ByVal D_BDay As Date, ByVal D_EDay As Date, ByVal sMB_AREA As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT * FROM MB_MEMBER WHERE CREDATE BETWEEN " & ProviderFactory.PositionPara & "D_BDay " & _
                     "  AND " & ProviderFactory.PositionPara & "D_EDay " & _
                     "  AND MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA ORDER BY CREDATE LIMIT " & iSTART & "," & iPageSize

            Dim paras(2) As IDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("D_BDay", D_BDay)
            paras(1) = ProviderFactory.CreateDataParameter("D_EDay", D_EDay)
            paras(2) = ProviderFactory.CreateDataParameter("MB_AREA", sMB_AREA)

            Return Me.loadBySQLOnlyDs(sqlStr, paras)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function QRY_4(ByVal iSTART As Decimal, ByVal iPageSize As Decimal, ByVal D_BDay As Date, ByVal D_EDay As Date, ByVal sMB_LEADER As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT * FROM MB_MEMBER WHERE CREDATE BETWEEN " & ProviderFactory.PositionPara & "D_BDay " & _
                     "  AND " & ProviderFactory.PositionPara & "D_EDay " & _
                     "  AND MB_LEADER = " & ProviderFactory.PositionPara & "MB_LEADER ORDER BY CREDATE LIMIT " & iSTART & "," & iPageSize

            Dim paras(2) As IDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("D_BDay", D_BDay)
            paras(1) = ProviderFactory.CreateDataParameter("D_EDay", D_EDay)
            paras(2) = ProviderFactory.CreateDataParameter("MB_LEADER", sMB_LEADER)

            Return Me.loadBySQLOnlyDs(sqlStr, paras)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function QRY_5(ByVal iSTART As Decimal, ByVal iPageSize As Decimal, ByVal D_BDay As Date, ByVal D_EDay As Date, ByVal sMB_AREA As String, ByVal sMB_LEADER As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT * FROM MB_MEMBER WHERE CREDATE BETWEEN " & ProviderFactory.PositionPara & "D_BDay " & _
                     "  AND " & ProviderFactory.PositionPara & "D_EDay " & _
                     "  AND MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA " & _
                     "  AND MB_LEADER = " & ProviderFactory.PositionPara & "MB_LEADER ORDER BY CREDATE LIMIT " & iSTART & "," & iPageSize

            Dim paras(3) As IDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("D_BDay", D_BDay)
            paras(1) = ProviderFactory.CreateDataParameter("D_EDay", D_EDay)
            paras(2) = ProviderFactory.CreateDataParameter("MB_AREA", sMB_AREA)
            paras(3) = ProviderFactory.CreateDataParameter("MB_LEADER", sMB_LEADER)

            Return Me.loadBySQLOnlyDs(sqlStr, paras)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function QRY_6(ByVal iSTART As Decimal, ByVal iPageSize As Decimal, ByVal sMB_AREA As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT * FROM MB_MEMBER WHERE MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA ORDER BY CREDATE LIMIT " & iSTART & "," & iPageSize

            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_AREA", sMB_AREA)

            Return Me.loadBySQLOnlyDs(sqlStr, para)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function QRY_7(ByVal iSTART As Decimal, ByVal iPageSize As Decimal, ByVal sMB_AREA As String, ByVal sMB_LEADER As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT * FROM MB_MEMBER WHERE MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA " & _
                     "  AND MB_LEADER = " & ProviderFactory.PositionPara & "MB_LEADER ORDER BY CREDATE LIMIT " & iSTART & "," & iPageSize

            Dim paras(1) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("MB_AREA", sMB_AREA)
            paras(1) = ProviderFactory.CreateDataParameter("MB_LEADER", sMB_LEADER)

            Return Me.loadBySQLOnlyDs(sqlStr, paras)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function
#End Region

#Region "報表查詢(EXCEL)"
    Public Function QRY_EXCEL_1() As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT * FROM MB_MEMBER ORDER BY CREDATE "

            Return Me.loadBySQLOnlyDs(sqlStr)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function QRY_EXCEL_2(ByVal D_BDay As Date, ByVal D_EDay As Date) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT * FROM MB_MEMBER WHERE CREDATE BETWEEN " & ProviderFactory.PositionPara & "D_BDay " & _
                     "  AND " & ProviderFactory.PositionPara & "D_EDay ORDER BY CREDATE "

            Dim paras(1) As IDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("D_BDay", D_BDay)
            paras(1) = ProviderFactory.CreateDataParameter("D_EDay", D_EDay)

            Return Me.loadBySQLOnlyDs(sqlStr, paras)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function QRY_EXCEL_3(ByVal D_BDay As Date, ByVal D_EDay As Date, ByVal sMB_AREA As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT * FROM MB_MEMBER WHERE CREDATE BETWEEN " & ProviderFactory.PositionPara & "D_BDay " & _
                     "  AND " & ProviderFactory.PositionPara & "D_EDay " & _
                     "  AND MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA ORDER BY CREDATE "

            Dim paras(2) As IDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("D_BDay", D_BDay)
            paras(1) = ProviderFactory.CreateDataParameter("D_EDay", D_EDay)
            paras(2) = ProviderFactory.CreateDataParameter("MB_AREA", sMB_AREA)

            Return Me.loadBySQLOnlyDs(sqlStr, paras)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function QRY_EXCEL_4(ByVal D_BDay As Date, ByVal D_EDay As Date, ByVal sMB_LEADER As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT * FROM MB_MEMBER WHERE CREDATE BETWEEN " & ProviderFactory.PositionPara & "D_BDay " & _
                     "  AND " & ProviderFactory.PositionPara & "D_EDay " & _
                     "  AND MB_LEADER = " & ProviderFactory.PositionPara & "MB_LEADER ORDER BY CREDATE "

            Dim paras(2) As IDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("D_BDay", D_BDay)
            paras(1) = ProviderFactory.CreateDataParameter("D_EDay", D_EDay)
            paras(2) = ProviderFactory.CreateDataParameter("MB_LEADER", sMB_LEADER)

            Return Me.loadBySQLOnlyDs(sqlStr, paras)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function QRY_EXCEL_5(ByVal D_BDay As Date, ByVal D_EDay As Date, ByVal sMB_AREA As String, ByVal sMB_LEADER As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT * FROM MB_MEMBER WHERE CREDATE BETWEEN " & ProviderFactory.PositionPara & "D_BDay " & _
                     "  AND " & ProviderFactory.PositionPara & "D_EDay " & _
                     "  AND MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA " & _
                     "  AND MB_LEADER = " & ProviderFactory.PositionPara & "MB_LEADER ORDER BY CREDATE "

            Dim paras(3) As IDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("D_BDay", D_BDay)
            paras(1) = ProviderFactory.CreateDataParameter("D_EDay", D_EDay)
            paras(2) = ProviderFactory.CreateDataParameter("MB_AREA", sMB_AREA)
            paras(3) = ProviderFactory.CreateDataParameter("MB_LEADER", sMB_LEADER)

            Return Me.loadBySQLOnlyDs(sqlStr, paras)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function QRY_EXCEL_6(ByVal sMB_AREA As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT * FROM MB_MEMBER WHERE MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA ORDER BY CREDATE "

            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_AREA", sMB_AREA)

            Return Me.loadBySQLOnlyDs(sqlStr, para)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function QRY_EXCEL_7(ByVal sMB_AREA As String, ByVal sMB_LEADER As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT * FROM MB_MEMBER WHERE MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA " & _
                     "  AND MB_LEADER = " & ProviderFactory.PositionPara & "MB_LEADER ORDER BY CREDATE "

            Dim paras(1) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("MB_AREA", sMB_AREA)
            paras(1) = ProviderFactory.CreateDataParameter("MB_LEADER", sMB_LEADER)

            Return Me.loadBySQLOnlyDs(sqlStr, paras)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function
#End Region

#Region "報表統計"
    Public Function COUNT_1() As Decimal
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT IFNULL(COUNT(*),0) FROM MB_MEMBER "

            If Me.getDatabaseManager.isTransaction Then
                Return DBObject.ExecuteScalar(Me.getDatabaseManager.getTransaction, CommandType.Text, sqlStr)
            Else
                Return DBObject.ExecuteScalar(Me.getDatabaseManager.getConnection, CommandType.Text, sqlStr)
            End If
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function COUNT_2(ByVal D_BDay As Date, ByVal D_EDay As Date) As Decimal
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT IFNULL(COUNT(*),0) FROM MB_MEMBER WHERE CREDATE BETWEEN " & ProviderFactory.PositionPara & "D_BDay " & _
                     "  AND " & ProviderFactory.PositionPara & "D_EDay "

            Dim paras(1) As IDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("D_BDay", D_BDay)
            paras(1) = ProviderFactory.CreateDataParameter("D_EDay", D_EDay)

            If Me.getDatabaseManager.isTransaction Then
                Return DBObject.ExecuteScalar(Me.getDatabaseManager.getTransaction, CommandType.Text, sqlStr, paras)
            Else
                Return DBObject.ExecuteScalar(Me.getDatabaseManager.getConnection, CommandType.Text, sqlStr, paras)
            End If
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function COUNT_3(ByVal D_BDay As Date, ByVal D_EDay As Date, ByVal sMB_AREA As String) As Decimal
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT IFNULL(COUNT(*),0) FROM MB_MEMBER WHERE CREDATE BETWEEN " & ProviderFactory.PositionPara & "D_BDay " & _
                     "  AND " & ProviderFactory.PositionPara & "D_EDay " & _
                     "  AND MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA "

            Dim paras(2) As IDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("D_BDay", D_BDay)
            paras(1) = ProviderFactory.CreateDataParameter("D_EDay", D_EDay)
            paras(2) = ProviderFactory.CreateDataParameter("MB_AREA", sMB_AREA)

            If Me.getDatabaseManager.isTransaction Then
                Return DBObject.ExecuteScalar(Me.getDatabaseManager.getTransaction, CommandType.Text, sqlStr, paras)
            Else
                Return DBObject.ExecuteScalar(Me.getDatabaseManager.getConnection, CommandType.Text, sqlStr, paras)
            End If
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function COUNT_4(ByVal D_BDay As Date, ByVal D_EDay As Date, ByVal sMB_LEADER As String) As Decimal
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT IFNULL(COUNT(*),0) FROM MB_MEMBER WHERE CREDATE BETWEEN " & ProviderFactory.PositionPara & "D_BDay " & _
                     "  AND " & ProviderFactory.PositionPara & "D_EDay " & _
                     "  AND MB_LEADER = " & ProviderFactory.PositionPara & "MB_LEADER "

            Dim paras(2) As IDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("D_BDay", D_BDay)
            paras(1) = ProviderFactory.CreateDataParameter("D_EDay", D_EDay)
            paras(2) = ProviderFactory.CreateDataParameter("MB_LEADER", sMB_LEADER)

            If Me.getDatabaseManager.isTransaction Then
                Return DBObject.ExecuteScalar(Me.getDatabaseManager.getTransaction, CommandType.Text, sqlStr, paras)
            Else
                Return DBObject.ExecuteScalar(Me.getDatabaseManager.getConnection, CommandType.Text, sqlStr, paras)
            End If
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function COUNT_5(ByVal D_BDay As Date, ByVal D_EDay As Date, ByVal sMB_AREA As String, ByVal sMB_LEADER As String) As Decimal
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT IFNULL(COUNT(*),0) FROM MB_MEMBER WHERE CREDATE BETWEEN " & ProviderFactory.PositionPara & "D_BDay " & _
                     "  AND " & ProviderFactory.PositionPara & "D_EDay " & _
                     "  AND MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA " & _
                     "  AND MB_LEADER = " & ProviderFactory.PositionPara & "MB_LEADER "

            Dim paras(3) As IDataParameter

            paras(0) = ProviderFactory.CreateDataParameter("D_BDay", D_BDay)
            paras(1) = ProviderFactory.CreateDataParameter("D_EDay", D_EDay)
            paras(2) = ProviderFactory.CreateDataParameter("MB_AREA", sMB_AREA)
            paras(3) = ProviderFactory.CreateDataParameter("MB_LEADER", sMB_LEADER)

            If Me.getDatabaseManager.isTransaction Then
                Return DBObject.ExecuteScalar(Me.getDatabaseManager.getTransaction, CommandType.Text, sqlStr, paras)
            Else
                Return DBObject.ExecuteScalar(Me.getDatabaseManager.getConnection, CommandType.Text, sqlStr, paras)
            End If
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function COUNT_6(ByVal sMB_AREA As String) As Decimal
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT IFNULL(COUNT(*),0) FROM MB_MEMBER WHERE MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA "

            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_AREA", sMB_AREA)

            If Me.getDatabaseManager.isTransaction Then
                Return DBObject.ExecuteScalar(Me.getDatabaseManager.getTransaction, CommandType.Text, sqlStr, para)
            Else
                Return DBObject.ExecuteScalar(Me.getDatabaseManager.getConnection, CommandType.Text, sqlStr, para)
            End If
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function COUNT_7(ByVal sMB_AREA As String, ByVal sMB_LEADER As String) As Decimal
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT IFNULL(COUNT(*),0) FROM MB_MEMBER WHERE MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA " & _
                     "  AND MB_LEADER = " & ProviderFactory.PositionPara & "MB_LEADER "

            Dim paras(1) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("MB_AREA", sMB_AREA)
            paras(1) = ProviderFactory.CreateDataParameter("MB_LEADER", sMB_LEADER)

            If Me.getDatabaseManager.isTransaction Then
                Return DBObject.ExecuteScalar(Me.getDatabaseManager.getTransaction, CommandType.Text, sqlStr, paras)
            Else
                Return DBObject.ExecuteScalar(Me.getDatabaseManager.getConnection, CommandType.Text, sqlStr, paras)
            End If
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function
#End Region

#End Region

End Class
