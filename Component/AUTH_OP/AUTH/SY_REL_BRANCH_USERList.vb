Imports com.Azion.NET.VB
Imports AUTH_OP.TABLE

Public Class SY_REL_BRANCH_USERList
    : Inherits BosList

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.new("SY_REL_BRANCH_USER", dbManager)
    End Sub

    Overrides Function newBos() As BosBase
        Return New SY_REL_BRANCH_USER(MyBase.getDatabaseManager)
    End Function

#Region "濟南昱勝添加"
#Region "Lake Function"

    ''' <summary>
    ''' 取得登入者對應部門及單位
    ''' </summary>
    ''' <param name="sStaffId">登入者編號</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' <history>
    ''' [Lake] 2012/04/06 Created
    ''' Avril 2012/06/27 添加當前登入者的部門
    '''</history>
    Function loadByStaffId(ByVal sStaffId As String, ByVal sBraDepno As String) As Boolean
        Try
            Dim strSql As String = "SELECT " _
                                 & "    syb.BRID " _
                                 & "   ,syb.BRA_DEPNO " _
                                 & "   ,syb.BRCNAME " _
                                 & "FROM" _
                                 & "   SY_REL_BRANCH_USER syrbu JOIN SY_BRANCH syb " _
                                 & "ON " _
                                 & "   syrbu.BRA_DEPNO = syb.BRA_DEPNO " _
                                 & "WHERE" _
                                 & "   syrbu.STAFFID = " & ProviderFactory.PositionPara & "STAFFID " _
                                 & " AND  syrbu.BRA_DEPNO = " & ProviderFactory.PositionPara & "BRA_DEPNO " _
                                 & "ORDER BY syb.BRID"

            Dim para(1) As IDataParameter
            para(0) = ProviderFactory.CreateDataParameter("STAFFID", sStaffId)
            para(1) = ProviderFactory.CreateDataParameter("BRA_DEPNO", sBraDepno)

            If MyBase.loadBySQL(strSql, para) > 0 Then
                Return True
            Else
                Return False
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
    ''' 根據部門代碼判斷是否該部門下還有員工
    ''' </summary>
    ''' <param name="sBraDepNo">部門代碼</param>
    ''' <returns></returns>
    ''' <remarks>[Lake] 2012/05/17 Created</remarks>
    Public Function getUserCountByBraDepNo(ByVal sBraDepNo As String) As Boolean
        Try
            Dim sSql As String = "SELECT " & _
                                 "   COUNT(*) " & _
                                 "FROM " & _
                                 "   SY_REL_BRANCH_USER " & _
                                 "WHERE " & _
                                 "   BRA_DEPNO = " & ProviderFactory.PositionPara & "BRA_DEPNO"

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
#End Region

#Region "Avril Function"
    ''' <summary>
    ''' 查詢登入者組織資料
    ''' </summary>
    ''' <param name="SStaffId">登入者</param>
    ''' <remarks>
    ''' Add by Avril 2012/06/27
    ''' </remarks>
    Function GetBranchCount(ByVal SStaffId As String) As Integer
        Try
            Dim sSql As String = "SELECT " & _
                              "   BRID, SY_REL_BRANCH_USER.BRA_DEPNO, BRCNAME " & _
                              "FROM " & _
                              "   SY_REL_BRANCH_USER " & _
                              "   JOIN SY_BRANCH ON SY_REL_BRANCH_USER. BRA_DEPNO= SY_BRANCH.BRA_DEPNO  AND PARENT='0' " & _
                              "WHERE " & _
                              "   SY_REL_BRANCH_USER.STAFFID = " & ProviderFactory.PositionPara & "STAFFID" & _
                              " AND SY_BRANCH.DISABLED='0' AND SY_BRANCH.PARENT = '0' "

            Dim paras As IDbDataParameter = ProviderFactory.CreateDataParameter("STAFFID", SStaffId)

            Return MyBase.loadBySQL(sSql, paras)
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
