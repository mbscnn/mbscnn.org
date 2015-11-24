Option Explicit On
Option Strict On

Imports com.Azion.NET.VB


Public Class SY_BRANCH
    Inherits BosBase

    ''' <summary>
    ''' Constructor
    ''' </summary>
    ''' <param name="dbManager">DatabaseManager</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal dbManager As com.Azion.NET.VB.DatabaseManager)
        MyBase.new("SY_BRANCH", dbManager)
    End Sub


#Region "濟南昱勝添加"
#Region "Avril Function"

    ''' <summary>
    ''' 根據主鍵查找資料
    ''' </summary>
    ''' <param name="sBraDepNo">單位編號</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Add by Avril 2012/05/11
    ''' </remarks>
    Function loadByPK(ByVal sBraDepNo As String) As Boolean
        Try
            Dim para(0) As IDbDataParameter
            para(0) = ProviderFactory.CreateDataParameter("BRA_DEPNO", sBraDepNo)

            Return MyBase.loadBySQL(para)
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 根據案件編號查詢資料
    ''' </summary>
    ''' <param name="sCaseId">案件編號</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Add by Avril 2012/04/24
    ''' </remarks>
    Function loadByCaseId(ByVal sCaseId As String) As Boolean
        Try
            Dim sSQL As String = "SELECT * FROM SY_BRANCH JOIN SY_CASEID " & _
                                "  ON  SY_BRANCH. BRA_DEPNO  = SY_CASEID. APP_BRADEPNO " & _
                                 " WHERE SY_CASEID.CASEID = " & ProviderFactory.PositionPara & "CASEID "

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
    ''' 根據主鍵刪除部門
    ''' </summary>
    ''' <param name="sBraDepNo">部門代碼</param>
    ''' <returns></returns>
    ''' <remarks>[Lake] 2012/05/16 Created</remarks>
    Public Function deleteByPk(ByVal sBraDepNo As String) As Boolean
        Try
           

            Dim sSql As String = "DELETE " & _
                                "FROM " & _
                                "   SY_BRANCHMGR " & _
                                "WHERE " & _
                                "   BRA_DEPNO = " & ProviderFactory.PositionPara & "BRA_DEPNO"

            Dim paras As IDbDataParameter = ProviderFactory.CreateDataParameter("BRA_DEPNO", sBraDepNo)


            DBObject.ExecuteNonQuery(Me.getDatabaseManager, CommandType.Text, sSql, paras)


            sSql = "DELETE " & _
                                "FROM " & _
                                "   SY_BRANCH " & _
                                "WHERE " & _
                                "   BRA_DEPNO = " & ProviderFactory.PositionPara & "BRA_DEPNO"

            Return DBObject.ExecuteNonQuery(Me.getDatabaseManager, CommandType.Text, sSql, paras) > 0

        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 根據部門代碼查詢組織資料代碼為NULL部門資料條數
    ''' </summary>
    ''' <param name="sBraDepNo">部門代碼</param>
    ''' <returns></returns>
    ''' <remarks>[Lake] 2012/05/17 Created</remarks>
    Public Function getBranchCountByPK(ByVal sBraDepNo As String) As Boolean
        Try
            Dim sSql As String = "SELECT " & _
                                 "   COUNT(*) " & _
                                 "FROM " & _
                                 "   SY_BRANCH " & _
                                 "WHERE " & _
                                 "       BRA_DEPNO = " & ProviderFactory.PositionPara & "BRA_DEPNO " & _
                                 "   AND EPSDEP IS NOT NULL"

            Dim paras As IDbDataParameter = ProviderFactory.CreateDataParameter("BRA_DEPNO", sBraDepNo)
            Dim oRowCount As Object = com.Azion.NET.VB.DBObject.ExecuteScalar(Me.getDatabaseManager().getConnection(), CommandType.Text, sSql, paras)

            Return Convert.ToInt32(oRowCount) > 0
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 取得最大部門代碼加1
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>[Lake] 2012/05/17 Created</remarks>
    Public Function getMaxBraDepNo() As String
        Try
            Dim sSql As String = "SELECT ISNULL(MAX(BRA_DEPNO),0)+1 FROM SY_BRANCH"

            Dim oBraDepNo As Object = com.Azion.NET.VB.DBObject.ExecuteScalar(Me.getDatabaseManager.getConnection(), CommandType.Text, sSql)

            Return oBraDepNo.ToString()
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 取得最大BRID加1
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>[Lake] 2012/05/17 Created</remarks>
    Public Function getMaxBrid() As String
        Try
            Dim sSql As String = "SELECT ISNULL(MAX(BRID),0)+1 FROM SY_BRANCH"

            Dim oBraDepNo As Object = com.Azion.NET.VB.DBObject.ExecuteScalar(Me.getDatabaseManager.getConnection(), CommandType.Text, sSql)

            Return oBraDepNo.ToString()
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 取得當前登錄人員所在單位
    ''' </summary>
    ''' <param name="sBrid"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getFirstLevelBraDepNo(ByVal sBrid As String) As Boolean
        Try
            Dim sSql As String = "select * from sy_branch where parent =0 and brid=" & ProviderFactory.PositionPara & "BRID"

            Dim paras As IDbDataParameter = ProviderFactory.CreateDataParameter("BRID", sBrid)

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
    ''' 根據BRID或者MGR_BRID修改DB資料
    ''' </summary>
    ''' <param name="sNewBrid">新單位代碼或者管理單位代碼</param>
    ''' <param name="sOldBrid">舊單位代碼或者管理單位代碼</param>
    ''' <returns></returns>
    ''' <remarks>[Lake] 2012/07/05 Created</remarks>
    Public Function updateByBrid(ByVal sNewBrid As String, ByVal sOldBrid As String) As Boolean
        Try
            Dim sSql As String = String.Empty

            sSql = "UPDATE SY_BRANCH SET BRID = " & ProviderFactory.PositionPara & "NEWBRID " & _
                "WHERE BRID = " & ProviderFactory.PositionPara & "OLDBRID"

            Dim paras(1) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("NEWBRID", sNewBrid)
            paras(1) = ProviderFactory.CreateDataParameter("OLDBRID", sOldBrid)

            If Me.getDatabaseManager.isTransaction Then
                Return DBObject.ExecuteNonQuery(Me.getDatabaseManager.getTransaction, CommandType.Text, sSql, paras) > 0
            Else
                Return DBObject.ExecuteNonQuery(Me.getDatabaseManager.getConnection, CommandType.Text, sSql, paras) > 0
            End If
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 取得部門資料
    ''' </summary>
    ''' <param name="sBraDepNo">部門編號</param>
    ''' <returns></returns>
    ''' <remarks>[Lake] 2012/05/16 Created</remarks>
    Public Function getBrancByBraDepNo(ByVal sBraDepNo As String) As Boolean
        Try
            Dim sSql As String = "SELECT " & _
                                 "    BRA_DEPNO " & _
                                 "   ,(CONVERT(VARCHAR,BRID) + '(' + CONVERT(VARCHAR,BRA_DEPNO) + ')' + '  ' + BRCNAME + " & _
                                 "    CASE " & _
                                 "    WHEN EPSDEP IS NULL OR EPSDEP = '' THEN '' " & _
                                 "    ELSE '('+EPSDEP+')' END) AS BRCNAME " & _
                                 "   ,PARENT " & _
                                 "FROM " & _
                                 "   SY_BRANCH " & _
                                 "WHERE " & _
                                 "       BRA_DEPNO = " & ProviderFactory.PositionPara & "BRADEPNO"

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
    ''' 判斷Brid是否重複
    ''' </summary>
    ''' <param name="sBrid">新增的分行代碼</param>
    ''' <returns></returns>
    ''' <remarks>[Lake] 2012/07/05 Created</remarks>
    Public Function existBrid(ByVal sBrid As String) As Boolean
        Try
            Dim sSql As String = "SELECT TOP(1) * FROM SY_BRANCH WHERE BRID = " & ProviderFactory.PositionPara & "BRID"

            Dim paras As IDbDataParameter = ProviderFactory.CreateDataParameter("BRID", sBrid)

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
    ''' 判斷輸入的管理單位是否存在，
    ''' 返回有效管理單位個數
    ''' 通過前臺比較判斷
    ''' </summary>
    ''' <param name="sMgrBrid">管理單位</param>
    ''' <returns></returns>
    ''' <remarks>[Lake] 2012/07/13 Created</remarks>
    Public Function existMgrBrid(ByVal sMgrBrid As String) As String
        Try
            Dim sSql As String = "select count(bra_depno) " & _
                "from sy_branch where bra_depno in('" & sMgrBrid & "')"

            Return com.Azion.NET.VB.DBObject.ExecuteScalar(Me.getDatabaseManager().getConnection(), CommandType.Text, sSql).ToString()
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 上級單位Disabled狀態變更時，更新部門Disabled狀態
    ''' </summary>
    ''' <param name="sBraDepNo"></param>
    ''' <param name="sDisabled"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function updateDisabled(ByVal sBraDepNo As String, ByVal sDisabled As String) As Boolean
        Try
            Dim sSql As String = "update sy_branch " & _
                "set disabled= " & ProviderFactory.PositionPara & "DISABLED " & _
                "where bra_depno in('" & sBraDepNo & "')"

            Dim paras As IDbDataParameter = ProviderFactory.CreateDataParameter("DISABLED", sDisabled)

            If Me.getDatabaseManager.isTransaction Then
                Return DBObject.ExecuteNonQuery(Me.getDatabaseManager.getTransaction, CommandType.Text, sSql, paras) > 0
            Else
                Return DBObject.ExecuteNonQuery(Me.getDatabaseManager.getConnection, CommandType.Text, sSql, paras) > 0
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

#Region "Zack Function"

    ''' <summary>
    ''' 判斷單位名稱是否重複
    ''' </summary>
    ''' <param name="sParent">父節點Value</param>
    ''' <param name="sBranchName">單位名稱</param>
    ''' <param name="sBraDepno">單位編號</param>
    ''' <param name="sBrid">分行代號</param>
    ''' <param name="bIsEdit">是否為修改單位信息</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Zack  2012-06-26 Create
    ''' </remarks>
    Public Function loadBranchName(ByVal sParent As String, ByVal sBranchName As String, ByVal sBraDepno As String, ByVal sBrid As String, ByVal bIsEdit As Boolean) As Boolean
        Try
            Dim sSql As String = "SELECT * FROM SY_BRANCH WHERE BRCNAME =" & ProviderFactory.PositionPara & "BRCNAME"

            ' 若沒有父節點
            If sParent = "0" Then
                sSql = sSql & " AND PARENT = 0"
            Else
                sSql = sSql & " AND BRID =" & ProviderFactory.PositionPara & "BRID"
            End If

            ' 若是修改單位
            If bIsEdit Then
                sSql = sSql & " AND BRA_DEPNO !=" & ProviderFactory.PositionPara & "BRA_DEPNO"
            End If

            Dim paras(2) As System.Data.IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("BRCNAME", sBranchName)
            paras(1) = ProviderFactory.CreateDataParameter("BRA_DEPNO", sBraDepno)

            ' 若有父節點
            If sParent <> "0" Then
                paras(2) = ProviderFactory.CreateDataParameter("BRID", sBrid)
            End If

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


