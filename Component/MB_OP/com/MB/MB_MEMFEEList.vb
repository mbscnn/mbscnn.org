Imports com.Azion.NET.VB

Public Class MB_MEMFEEList
    Inherits BosList

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.New("MB_MEMFEE", dbManager)
    End Sub

    Overrides Function newBos() As BosBase
        Dim bos As MB_MEMFEE = New MB_MEMFEE(MyBase.getDatabaseManager)
        Return bos
    End Function

#Region "Ted Function"

#Region "報表查詢"
    Public Function QRY_1(ByVal iSTART As Decimal, ByVal iPageSize As Decimal) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT C.*, D.*" & _
                     "  FROM (SELECT A.MB_MEMSEQ," & _
                     "               B.MB_NAME," & _
                     "               B.MB_AREA," & _
                     "               B.MB_LEADER" & _
                     "          FROM MB_VIP A, MB_MEMBER B" & _
                     "         WHERE A.MB_MEMSEQ = B.MB_MEMSEQ AND A.VIP = '1') C" & _
                     "       LEFT JOIN MB_MEMFEE D ON C.MB_MEMSEQ = D.MB_MEMSEQ" & _
                     " ORDER BY D.APVDATE " & _
                     " LIMIT " & iSTART & "," & iPageSize

            Return Me.loadBySQLOnlyDs(sqlStr)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function QRY_2(ByVal iB_MB_YYYY As Decimal, ByVal iE_MB_YYYY As Decimal, ByVal iSTART As Decimal, ByVal iPageSize As Decimal) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT C.*, D.*" & _
                     "  FROM (SELECT A.MB_MEMSEQ," & _
                     "               B.MB_NAME," & _
                     "               B.MB_AREA," & _
                     "               B.MB_LEADER" & _
                     "          FROM MB_VIP A, MB_MEMBER B" & _
                     "         WHERE A.MB_MEMSEQ = B.MB_MEMSEQ AND A.VIP = '1') C" & _
                     "       LEFT JOIN MB_MEMFEE D" & _
                     "          ON C.MB_MEMSEQ = D.MB_MEMSEQ AND D.MB_YYYY BETWEEN " & ProviderFactory.PositionPara & "B_MB_YYYY AND " & ProviderFactory.PositionPara & "E_MB_YYYY" & _
                     " ORDER BY D.APVDATE " & _
                     " LIMIT " & iSTART & "," & iPageSize

            Dim paras(1) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("B_MB_YYYY", iB_MB_YYYY)
            paras(1) = ProviderFactory.CreateDataParameter("E_MB_YYYY", iE_MB_YYYY)

            Return Me.loadBySQLOnlyDs(sqlStr, paras)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function QRY_3(ByVal iB_MB_YYYY As Decimal, ByVal iE_MB_YYYY As Decimal, ByVal sMB_AREA As String, ByVal iSTART As Decimal, ByVal iPageSize As Decimal) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT C.*, D.*" & _
                     "  FROM (SELECT A.MB_MEMSEQ," & _
                     "               B.MB_NAME," & _
                     "               B.MB_AREA," & _
                     "               B.MB_LEADER" & _
                     "          FROM MB_VIP A, MB_MEMBER B" & _
                     "         WHERE A.MB_MEMSEQ = B.MB_MEMSEQ AND A.VIP = '1' AND B.MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA)" & _
                     "       C" & _
                     "       LEFT JOIN MB_MEMFEE D" & _
                     "          ON C.MB_MEMSEQ = D.MB_MEMSEQ AND D.MB_YYYY BETWEEN " & ProviderFactory.PositionPara & "B_MB_YYYY AND " & ProviderFactory.PositionPara & "E_MB_YYYY " & _
                     " ORDER BY D.APVDATE " & _
                     " LIMIT " & iSTART & "," & iPageSize

            Dim paras(2) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("MB_AREA", sMB_AREA)
            paras(1) = ProviderFactory.CreateDataParameter("B_MB_YYYY", iB_MB_YYYY)
            paras(2) = ProviderFactory.CreateDataParameter("E_MB_YYYY", iE_MB_YYYY)

            Return Me.loadBySQLOnlyDs(sqlStr, paras)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function QRY_4(ByVal iB_MB_YYYY As Decimal, ByVal iE_MB_YYYY As Decimal, ByVal sMB_AREA As String, ByVal sMB_LEADER As String, ByVal iSTART As Decimal, ByVal iPageSize As Decimal) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT C.*, D.*" & _
                     "  FROM (SELECT A.MB_MEMSEQ," & _
                     "               B.MB_NAME," & _
                     "               B.MB_AREA," & _
                     "               B.MB_LEADER" & _
                     "          FROM MB_VIP A, MB_MEMBER B" & _
                     "         WHERE A.MB_MEMSEQ = B.MB_MEMSEQ AND A.VIP = '1' AND B.MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA" & _
                     "              AND B.MB_LEADER = " & ProviderFactory.PositionPara & "MB_LEADER) C " & _
                     "       LEFT JOIN MB_MEMFEE D" & _
                     "          ON C.MB_MEMSEQ = D.MB_MEMSEQ AND D.MB_YYYY BETWEEN " & ProviderFactory.PositionPara & "B_MB_YYYY AND " & ProviderFactory.PositionPara & "E_MB_YYYY " & _
                     " ORDER BY D.APVDATE " & _
                     " LIMIT " & iSTART & "," & iPageSize

            Dim paras(3) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("MB_AREA", sMB_AREA)
            paras(1) = ProviderFactory.CreateDataParameter("MB_LEADER", sMB_LEADER)
            paras(2) = ProviderFactory.CreateDataParameter("B_MB_YYYY", iB_MB_YYYY)
            paras(3) = ProviderFactory.CreateDataParameter("E_MB_YYYY", iE_MB_YYYY)

            Return Me.loadBySQLOnlyDs(sqlStr, paras)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function QRY_5(ByVal sMB_AREA As String, ByVal sMB_LEADER As String, ByVal iSTART As Decimal, ByVal iPageSize As Decimal) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT C.*, D.*" & _
                     "  FROM (SELECT A.MB_MEMSEQ," & _
                     "               B.MB_NAME," & _
                     "               B.MB_AREA," & _
                     "               B.MB_LEADER" & _
                     "          FROM MB_VIP A, MB_MEMBER B" & _
                     "         WHERE A.MB_MEMSEQ = B.MB_MEMSEQ AND A.VIP = '1' AND B.MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA" & _
                     "              AND B.MB_LEADER = " & ProviderFactory.PositionPara & "MB_LEADER) C " & _
                     "       LEFT JOIN MB_MEMFEE D" & _
                     "          ON C.MB_MEMSEQ = D.MB_MEMSEQ " & _
                     " ORDER BY D.APVDATE " & _
                     " LIMIT " & iSTART & "," & iPageSize

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

    Public Function QRY_6(ByVal sMB_AREA As String, ByVal iSTART As Decimal, ByVal iPageSize As Decimal) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT C.*, D.*" & _
                     "  FROM (SELECT A.MB_MEMSEQ," & _
                     "               B.MB_NAME," & _
                     "               B.MB_AREA," & _
                     "               B.MB_LEADER" & _
                     "          FROM MB_VIP A, MB_MEMBER B" & _
                     "         WHERE A.MB_MEMSEQ = B.MB_MEMSEQ AND A.VIP = '1' AND B.MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA)" & _
                     "       C" & _
                     "       LEFT JOIN MB_MEMFEE D" & _
                     "          ON C.MB_MEMSEQ = D.MB_MEMSEQ " & _
                     " ORDER BY D.APVDATE " & _
                     " LIMIT " & iSTART & "," & iPageSize

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

    Public Function QRY_7(ByVal sMB_LEADER As String, ByVal iSTART As Decimal, ByVal iPageSize As Decimal) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT C.*, D.*" & _
                     "  FROM (SELECT A.MB_MEMSEQ," & _
                     "               B.MB_NAME," & _
                     "               B.MB_AREA," & _
                     "               B.MB_LEADER" & _
                     "          FROM MB_VIP A, MB_MEMBER B" & _
                     "         WHERE A.MB_MEMSEQ = B.MB_MEMSEQ AND A.VIP = '1' AND B.MB_LEADER = " & ProviderFactory.PositionPara & "MB_LEADER)" & _
                     "       C" & _
                     "       LEFT JOIN MB_MEMFEE D" & _
                     "          ON C.MB_MEMSEQ = D.MB_MEMSEQ " & _
                     " ORDER BY D.APVDATE " & _
                     " LIMIT " & iSTART & "," & iPageSize

            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_LEADER", sMB_LEADER)

            Return Me.loadBySQLOnlyDs(sqlStr, para)
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

            sqlStr = "SELECT C.*, D.*" & _
                     "  FROM (SELECT A.MB_MEMSEQ," & _
                     "               B.MB_NAME," & _
                     "               B.MB_AREA," & _
                     "               B.MB_LEADER" & _
                     "          FROM MB_VIP A, MB_MEMBER B" & _
                     "         WHERE A.MB_MEMSEQ = B.MB_MEMSEQ AND A.VIP = '1') C" & _
                     "       LEFT JOIN MB_MEMFEE D ON C.MB_MEMSEQ = D.MB_MEMSEQ" & _
                     " ORDER BY D.APVDATE "

            Return Me.loadBySQLOnlyDs(sqlStr)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function QRY_EXCEL_2(ByVal iB_MB_YYYY As Decimal, ByVal iE_MB_YYYY As Decimal) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT C.*, D.*" & _
                     "  FROM (SELECT A.MB_MEMSEQ," & _
                     "               B.MB_NAME," & _
                     "               B.MB_AREA," & _
                     "               B.MB_LEADER" & _
                     "          FROM MB_VIP A, MB_MEMBER B" & _
                     "         WHERE A.MB_MEMSEQ = B.MB_MEMSEQ AND A.VIP = '1') C" & _
                     "       LEFT JOIN MB_MEMFEE D" & _
                     "          ON C.MB_MEMSEQ = D.MB_MEMSEQ AND D.MB_YYYY BETWEEN " & ProviderFactory.PositionPara & "B_MB_YYYY AND " & ProviderFactory.PositionPara & "E_MB_YYYY" & _
                     " ORDER BY D.APVDATE "

            Dim paras(1) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("B_MB_YYYY", iB_MB_YYYY)
            paras(1) = ProviderFactory.CreateDataParameter("E_MB_YYYY", iE_MB_YYYY)

            Return Me.loadBySQLOnlyDs(sqlStr, paras)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function QRY_EXCEL_3(ByVal iB_MB_YYYY As Decimal, ByVal iE_MB_YYYY As Decimal, ByVal sMB_AREA As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT C.*, D.*" & _
                     "  FROM (SELECT A.MB_MEMSEQ," & _
                     "               B.MB_NAME," & _
                     "               B.MB_AREA," & _
                     "               B.MB_LEADER" & _
                     "          FROM MB_VIP A, MB_MEMBER B" & _
                     "         WHERE A.MB_MEMSEQ = B.MB_MEMSEQ AND A.VIP = '1' AND B.MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA)" & _
                     "       C" & _
                     "       LEFT JOIN MB_MEMFEE D" & _
                     "          ON C.MB_MEMSEQ = D.MB_MEMSEQ AND D.MB_YYYY BETWEEN " & ProviderFactory.PositionPara & "B_MB_YYYY AND " & ProviderFactory.PositionPara & "E_MB_YYYY " & _
                     " ORDER BY D.APVDATE "

            Dim paras(2) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("MB_AREA", sMB_AREA)
            paras(1) = ProviderFactory.CreateDataParameter("B_MB_YYYY", iB_MB_YYYY)
            paras(2) = ProviderFactory.CreateDataParameter("E_MB_YYYY", iE_MB_YYYY)

            Return Me.loadBySQLOnlyDs(sqlStr, paras)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function QRY_EXCEL_4(ByVal iB_MB_YYYY As Decimal, ByVal iE_MB_YYYY As Decimal, ByVal sMB_AREA As String, ByVal sMB_LEADER As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT C.*, D.*" & _
                     "  FROM (SELECT A.MB_MEMSEQ," & _
                     "               B.MB_NAME," & _
                     "               B.MB_AREA," & _
                     "               B.MB_LEADER" & _
                     "          FROM MB_VIP A, MB_MEMBER B" & _
                     "         WHERE A.MB_MEMSEQ = B.MB_MEMSEQ AND A.VIP = '1' AND B.MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA" & _
                     "              AND B.MB_LEADER = " & ProviderFactory.PositionPara & "MB_LEADER) C " & _
                     "       LEFT JOIN MB_MEMFEE D" & _
                     "          ON C.MB_MEMSEQ = D.MB_MEMSEQ AND D.MB_YYYY BETWEEN " & ProviderFactory.PositionPara & "B_MB_YYYY AND " & ProviderFactory.PositionPara & "E_MB_YYYY " & _
                     " ORDER BY D.APVDATE "

            Dim paras(3) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("MB_AREA", sMB_AREA)
            paras(1) = ProviderFactory.CreateDataParameter("MB_LEADER", sMB_LEADER)
            paras(2) = ProviderFactory.CreateDataParameter("B_MB_YYYY", iB_MB_YYYY)
            paras(3) = ProviderFactory.CreateDataParameter("E_MB_YYYY", iE_MB_YYYY)

            Return Me.loadBySQLOnlyDs(sqlStr, paras)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function QRY_EXCEL_5(ByVal sMB_AREA As String, ByVal sMB_LEADER As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT C.*, D.*" & _
                     "  FROM (SELECT A.MB_MEMSEQ," & _
                     "               B.MB_NAME," & _
                     "               B.MB_AREA," & _
                     "               B.MB_LEADER" & _
                     "          FROM MB_VIP A, MB_MEMBER B" & _
                     "         WHERE A.MB_MEMSEQ = B.MB_MEMSEQ AND A.VIP = '1' AND B.MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA" & _
                     "              AND B.MB_LEADER = " & ProviderFactory.PositionPara & "MB_LEADER) C " & _
                     "       LEFT JOIN MB_MEMFEE D" & _
                     "          ON C.MB_MEMSEQ = D.MB_MEMSEQ " & _
                     " ORDER BY D.APVDATE "

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

    Public Function QRY_EXCEL_6(ByVal sMB_AREA As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT C.*, D.*" & _
                     "  FROM (SELECT A.MB_MEMSEQ," & _
                     "               B.MB_NAME," & _
                     "               B.MB_AREA," & _
                     "               B.MB_LEADER" & _
                     "          FROM MB_VIP A, MB_MEMBER B" & _
                     "         WHERE A.MB_MEMSEQ = B.MB_MEMSEQ AND A.VIP = '1' AND B.MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA)" & _
                     "       C" & _
                     "       LEFT JOIN MB_MEMFEE D" & _
                     "          ON C.MB_MEMSEQ = D.MB_MEMSEQ " & _
                     " ORDER BY D.APVDATE "

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

    Public Function QRY_EXCEL_7(ByVal sMB_LEADER As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT C.*, D.*" & _
                     "  FROM (SELECT A.MB_MEMSEQ," & _
                     "               B.MB_NAME," & _
                     "               B.MB_AREA," & _
                     "               B.MB_LEADER" & _
                     "          FROM MB_VIP A, MB_MEMBER B" & _
                     "         WHERE A.MB_MEMSEQ = B.MB_MEMSEQ AND A.VIP = '1' AND B.MB_LEADER = " & ProviderFactory.PositionPara & "MB_LEADER)" & _
                     "       C" & _
                     "       LEFT JOIN MB_MEMFEE D" & _
                     "          ON C.MB_MEMSEQ = D.MB_MEMSEQ " & _
                     " ORDER BY D.APVDATE "

            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_LEADER", sMB_LEADER)

            Return Me.loadBySQLOnlyDs(sqlStr, para)
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

            sqlStr = "SELECT IFNULL(COUNT(*),0)" & _
                     "  FROM (SELECT A.MB_MEMSEQ," & _
                     "               B.MB_NAME," & _
                     "               B.MB_AREA," & _
                     "               B.MB_LEADER" & _
                     "          FROM MB_VIP A, MB_MEMBER B" & _
                     "         WHERE A.MB_MEMSEQ = B.MB_MEMSEQ AND A.VIP = '1') C" & _
                     "       LEFT JOIN MB_MEMFEE D ON C.MB_MEMSEQ = D.MB_MEMSEQ"

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

    Public Function COUNT_2(ByVal iB_MB_YYYY As Decimal, ByVal iE_MB_YYYY As Decimal) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT IFNULL(COUNT(*), 0)" & _
                     "  FROM (SELECT A.MB_MEMSEQ," & _
                     "               B.MB_NAME," & _
                     "               B.MB_AREA," & _
                     "               B.MB_LEADER" & _
                     "          FROM MB_VIP A, MB_MEMBER B" & _
                     "         WHERE A.MB_MEMSEQ = B.MB_MEMSEQ AND A.VIP = '1') C" & _
                     "       LEFT JOIN MB_MEMFEE D" & _
                     "          ON C.MB_MEMSEQ = D.MB_MEMSEQ AND D.MB_YYYY BETWEEN " & ProviderFactory.PositionPara & "B_MB_YYYY AND " & ProviderFactory.PositionPara & "E_MB_YYYY " & _
                     "ORDER BY D.APVDATE"

            Dim paras(1) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("B_MB_YYYY", iB_MB_YYYY)
            paras(1) = ProviderFactory.CreateDataParameter("E_MB_YYYY", iE_MB_YYYY)

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

    Public Function COUNT_3(ByVal iB_MB_YYYY As Decimal, ByVal iE_MB_YYYY As Decimal, ByVal sMB_AREA As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT IFNULL(COUNT(*),0) " & _
                     "  FROM (SELECT A.MB_MEMSEQ," & _
                     "               B.MB_NAME," & _
                     "               B.MB_AREA," & _
                     "               B.MB_LEADER" & _
                     "          FROM MB_VIP A, MB_MEMBER B" & _
                     "         WHERE A.MB_MEMSEQ = B.MB_MEMSEQ AND A.VIP = '1' AND B.MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA)" & _
                     "       C" & _
                     "       LEFT JOIN MB_MEMFEE D" & _
                     "          ON C.MB_MEMSEQ = D.MB_MEMSEQ AND D.MB_YYYY BETWEEN " & ProviderFactory.PositionPara & "B_MB_YYYY AND " & ProviderFactory.PositionPara & "E_MB_YYYY " & _
                     " ORDER BY D.APVDATE "

            Dim paras(2) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("MB_AREA", sMB_AREA)
            paras(1) = ProviderFactory.CreateDataParameter("B_MB_YYYY", iB_MB_YYYY)
            paras(2) = ProviderFactory.CreateDataParameter("E_MB_YYYY", iE_MB_YYYY)

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

    Public Function COUNT_4(ByVal iB_MB_YYYY As Decimal, ByVal iE_MB_YYYY As Decimal, ByVal sMB_AREA As String, ByVal sMB_LEADER As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT IFNULL(COUNT(*),0) " & _
                     "  FROM (SELECT A.MB_MEMSEQ," & _
                     "               B.MB_NAME," & _
                     "               B.MB_AREA," & _
                     "               B.MB_LEADER" & _
                     "          FROM MB_VIP A, MB_MEMBER B" & _
                     "         WHERE A.MB_MEMSEQ = B.MB_MEMSEQ AND A.VIP = '1' AND B.MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA" & _
                     "              AND B.MB_LEADER = " & ProviderFactory.PositionPara & "MB_LEADER) C " & _
                     "       LEFT JOIN MB_MEMFEE D" & _
                     "          ON C.MB_MEMSEQ = D.MB_MEMSEQ AND D.MB_YYYY BETWEEN " & ProviderFactory.PositionPara & "B_MB_YYYY AND " & ProviderFactory.PositionPara & "E_MB_YYYY " & _
                     " ORDER BY D.APVDATE "

            Dim paras(3) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("MB_AREA", sMB_AREA)
            paras(1) = ProviderFactory.CreateDataParameter("MB_LEADER", sMB_LEADER)
            paras(2) = ProviderFactory.CreateDataParameter("B_MB_YYYY", iB_MB_YYYY)
            paras(3) = ProviderFactory.CreateDataParameter("E_MB_YYYY", iE_MB_YYYY)

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

    Public Function COUNT_5(ByVal sMB_AREA As String, ByVal sMB_LEADER As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT IFNULL(COUNT(*),0) " & _
                     "  FROM (SELECT A.MB_MEMSEQ," & _
                     "               B.MB_NAME," & _
                     "               B.MB_AREA," & _
                     "               B.MB_LEADER" & _
                     "          FROM MB_VIP A, MB_MEMBER B" & _
                     "         WHERE A.MB_MEMSEQ = B.MB_MEMSEQ AND A.VIP = '1' AND B.MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA" & _
                     "              AND B.MB_LEADER = " & ProviderFactory.PositionPara & "MB_LEADER) C " & _
                     "       LEFT JOIN MB_MEMFEE D" & _
                     "          ON C.MB_MEMSEQ = D.MB_MEMSEQ " & _
                     " ORDER BY D.APVDATE "

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

    Public Function COUNT_6(ByVal sMB_AREA As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT C.*, D.*" & _
                     "  FROM (SELECT A.MB_MEMSEQ," & _
                     "               B.MB_NAME," & _
                     "               B.MB_AREA," & _
                     "               B.MB_LEADER" & _
                     "          FROM MB_VIP A, MB_MEMBER B" & _
                     "         WHERE A.MB_MEMSEQ = B.MB_MEMSEQ AND A.VIP = '1' AND B.MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA)" & _
                     "       C" & _
                     "       LEFT JOIN MB_MEMFEE D" & _
                     "          ON C.MB_MEMSEQ = D.MB_MEMSEQ " & _
                     " ORDER BY D.APVDATE "

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

    Public Function COUNT_7(ByVal sMB_LEADER As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT C.*, D.*" & _
                     "  FROM (SELECT A.MB_MEMSEQ," & _
                     "               B.MB_NAME," & _
                     "               B.MB_AREA," & _
                     "               B.MB_LEADER" & _
                     "          FROM MB_VIP A, MB_MEMBER B" & _
                     "         WHERE A.MB_MEMSEQ = B.MB_MEMSEQ AND A.VIP = '1' AND B.MB_LEADER = " & ProviderFactory.PositionPara & "MB_LEADER)" & _
                     "       C" & _
                     "       LEFT JOIN MB_MEMFEE D" & _
                     "          ON C.MB_MEMSEQ = D.MB_MEMSEQ " & _
                     " ORDER BY D.APVDATE "

            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_LEADER", sMB_LEADER)

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
#End Region

#End Region

End Class
