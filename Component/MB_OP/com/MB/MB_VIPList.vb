Imports com.Azion.NET.VB

Public Class MB_VIPList
    Inherits BosList

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.New("MB_VIP", dbManager)
    End Sub

    Overrides Function newBos() As BosBase
        Dim bos As MB_VIP = New MB_VIP(MyBase.getDatabaseManager)
        Return bos
    End Function

#Region "Ted Function"
    ''' <summary>
    ''' 根據"會員編號"取得資料
    ''' </summary>
    ''' <param name="iMB_MEMSEQ">會員編號</param>
    ''' <returns>Boolean 是否有取得</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' </history>
    Public Function LoadByMB_MEMSEQ(ByVal iMB_MEMSEQ As Decimal) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT * FROM MB_VIP WHERE MB_MEMSEQ = " & ProviderFactory.PositionPara & "MB_MEMSEQ "

            Dim para As IDbDataParameter = ProviderFactory.CreateDataParameter("MB_MEMSEQ", iMB_MEMSEQ)

            Return Me.loadBySQL(sqlStr, para)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

#Region "種籽會員繳款報表"
    Public Function QRY_1(ByVal iSTART As Decimal, ByVal iPageSize As Decimal) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT C.*" & _
                     "  FROM (SELECT A.MB_MEMSEQ," & _
                     "               B.MB_NAME," & _
                     "               B.MB_AREA," & _
                     "               B.MB_LEADER" & _
                     "          FROM MB_VIP A, MB_MEMBER B" & _
                     "         WHERE A.VIP = '2' AND A.MB_MEMSEQ = B.MB_MEMSEQ) C" & _
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

    Public Function QRY_2(ByVal sMB_AREA As String, ByVal iSTART As Decimal, ByVal iPageSize As Decimal) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT C.*" & _
                     "  FROM (SELECT A.MB_MEMSEQ," & _
                     "               B.MB_NAME," & _
                     "               B.MB_AREA," & _
                     "               B.MB_LEADER" & _
                     "          FROM MB_VIP A, MB_MEMBER B" & _
                     "         WHERE A.VIP = '2' AND B.MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA AND A.MB_MEMSEQ = B.MB_MEMSEQ)" & _
                     "       C" & _
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

    Public Function QRY_3(ByVal sMB_AREA As String, ByVal sMB_LEADER As String, ByVal iSTART As Decimal, ByVal iPageSize As Decimal) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT C.*" & _
                     "  FROM (SELECT A.MB_MEMSEQ," & _
                     "               B.MB_NAME," & _
                     "               B.MB_AREA," & _
                     "               B.MB_LEADER" & _
                     "          FROM MB_VIP A, MB_MEMBER B" & _
                     "         WHERE A.VIP = '2' AND B.MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA " & _
                     "          AND B.MB_LEADER = " & ProviderFactory.PositionPara & "MB_LEADER " & _
                     "          AND A.MB_MEMSEQ = B.MB_MEMSEQ)" & _
                     "       C " & _
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

    Public Function QRY_4(ByVal sMB_LEADER As String, ByVal iSTART As Decimal, ByVal iPageSize As Decimal) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT C.*" & _
                     "  FROM (SELECT A.MB_MEMSEQ," & _
                     "               B.MB_NAME," & _
                     "               B.MB_AREA," & _
                     "               B.MB_LEADER" & _
                     "          FROM MB_VIP A, MB_MEMBER B" & _
                     "         WHERE A.VIP = '2' " & _
                     "          AND B.MB_LEADER = " & ProviderFactory.PositionPara & "MB_LEADER " & _
                     "          AND A.MB_MEMSEQ = B.MB_MEMSEQ)" & _
                     "       C " & _
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

#Region "報表統計"
    Public Function COUNT_1() As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT IFNULL(COUNT(*),0) " & _
                     "  FROM (SELECT A.MB_MEMSEQ," & _
                     "               B.MB_NAME," & _
                     "               B.MB_AREA," & _
                     "               B.MB_LEADER" & _
                     "          FROM MB_VIP A, MB_MEMBER B" & _
                     "         WHERE A.VIP = '2' AND A.MB_MEMSEQ = B.MB_MEMSEQ) C"

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

    Public Function COUNT_2(ByVal sMB_AREA As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT IFNULL(COUNT(*),0) " & _
                     "  FROM (SELECT A.MB_MEMSEQ," & _
                     "               B.MB_NAME," & _
                     "               B.MB_AREA," & _
                     "               B.MB_LEADER" & _
                     "          FROM MB_VIP A, MB_MEMBER B" & _
                     "         WHERE A.VIP = '2' AND B.MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA AND A.MB_MEMSEQ = A.MB_MEMSEQ)" & _
                     "       C "

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

    Public Function COUNT_3(ByVal sMB_AREA As String, ByVal sMB_LEADER As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT IFNULL(COUNT(*),0) " & _
                     "  FROM (SELECT A.MB_MEMSEQ," & _
                     "               B.MB_NAME," & _
                     "               B.MB_AREA," & _
                     "               B.MB_LEADER" & _
                     "          FROM MB_VIP A, MB_MEMBER B" & _
                     "         WHERE A.VIP = '2' AND B.MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA " & _
                     "          AND B.MB_LEADER = " & ProviderFactory.PositionPara & "MB_LEADER " & _
                     "          AND A.MB_MEMSEQ = A.MB_MEMSEQ)" & _
                     "       C "

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

    Public Function COUNT_4(ByVal sMB_LEADER As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT IFNULL(COUNT(*),0) " & _
                     "  FROM (SELECT A.MB_MEMSEQ," & _
                     "               B.MB_NAME," & _
                     "               B.MB_AREA," & _
                     "               B.MB_LEADER" & _
                     "          FROM MB_VIP A, MB_MEMBER B" & _
                     "         WHERE A.VIP = '2' " & _
                     "          AND B.MB_LEADER = " & ProviderFactory.PositionPara & "MB_LEADER " & _
                     "          AND A.MB_MEMSEQ = B.MB_MEMSEQ " & _
                     "          AND EXISTS (SELECT * FROM MB_MEMREV WHERE MB_MEMSEQ=A.MB_MEMSEQ AND MB_ITEMID = 'A' and  MB_MEMTYP = '2'))" & _
                     "       C "

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

#Region "種籽會員繳款報表EXCEL"
    Public Function EXCEL_1() As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT C.*" & _
                     "  FROM (SELECT A.MB_MEMSEQ," & _
                     "               B.MB_NAME," & _
                     "               B.MB_AREA," & _
                     "               B.MB_LEADER" & _
                     "          FROM MB_VIP A, MB_MEMBER B" & _
                     "         WHERE A.VIP = '2' AND A.MB_MEMSEQ = B.MB_MEMSEQ) C"

            Return Me.loadBySQLOnlyDs(sqlStr)
        Catch ex As ProviderException
            Throw
        Catch ex As BosException
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function EXCEL_2(ByVal sMB_AREA As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT C.*" & _
                     "  FROM (SELECT A.MB_MEMSEQ," & _
                     "               B.MB_NAME," & _
                     "               B.MB_AREA," & _
                     "               B.MB_LEADER" & _
                     "          FROM MB_VIP A, MB_MEMBER B" & _
                     "         WHERE A.VIP = '2' AND B.MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA AND A.MB_MEMSEQ = B.MB_MEMSEQ)" & _
                     "       C"

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

    Public Function EXCEL_3(ByVal sMB_AREA As String, ByVal sMB_LEADER As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT C.*" & _
                     "  FROM (SELECT A.MB_MEMSEQ," & _
                     "               B.MB_NAME," & _
                     "               B.MB_AREA," & _
                     "               B.MB_LEADER" & _
                     "          FROM MB_VIP A, MB_MEMBER B" & _
                     "         WHERE A.VIP = '2' AND B.MB_AREA = " & ProviderFactory.PositionPara & "MB_AREA " & _
                     "          AND B.MB_LEADER = " & ProviderFactory.PositionPara & "MB_LEADER " & _
                     "          AND A.MB_MEMSEQ = B.MB_MEMSEQ)" & _
                     "       C "

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

    Public Function EXCEL_4(ByVal sMB_LEADER As String) As Integer
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT C.*" & _
                     "  FROM (SELECT A.MB_MEMSEQ," & _
                     "               B.MB_NAME," & _
                     "               B.MB_AREA," & _
                     "               B.MB_LEADER" & _
                     "          FROM MB_VIP A, MB_MEMBER B" & _
                     "         WHERE A.VIP = '2' " & _
                     "          AND B.MB_LEADER = " & ProviderFactory.PositionPara & "MB_LEADER " & _
                     "          AND A.MB_MEMSEQ = B.MB_MEMSEQ)" & _
                     "       C "

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


#End Region

End Class
