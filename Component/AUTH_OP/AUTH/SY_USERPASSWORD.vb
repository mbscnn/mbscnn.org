Option Explicit On
Option Strict On

Imports com.Azion.NET.VB

Namespace TABLE

    Public Class SY_USERPASSWORD
        Inherits BosBase

        Sub New()
            MyBase.New()
            Me.setPrimaryKeys()
        End Sub

        ''' <summary>
        ''' Constructor
        ''' </summary>
        ''' <param name="dbManager">DatabaseManager</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal dbManager As com.Azion.NET.VB.DatabaseManager)
            MyBase.new("SY_USERPASSWORD", dbManager)
        End Sub

#Region "濟南昱勝添加"
#Region "Zack Function"

        ''' <summary>
        ''' 根據主鍵查詢
        ''' </summary>
        ''' <param name="sStaffId">員工編號</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' Zack 2012-06-05 Create
        ''' </remarks>
        Function loadByPK(ByVal sStaffId As String) As Boolean
            Try
                Dim paras(0) As System.Data.IDbDataParameter
                paras(0) = ProviderFactory.CreateDataParameter("STAFFID", sStaffId.Trim)
                Return MyBase.loadData(paras)
            Catch ex As ProviderException
                Throw ex
            Catch ex As BosException
                Throw ex
            Catch ex As Exception
                Throw ex
            End Try
        End Function

        ''' <summary>
        ''' 根據人員編號修改密碼錯誤次數
        ''' </summary>
        ''' <param name="sStaffId">人員編號</param>
        ''' <param name="sPwdWrongTimes">密碼錯誤次數</param>
        ''' <remarks>
        ''' Zack 2012-06-05 Create
        ''' </remarks>
        Sub updatePWD_WRONG_TIMES(ByVal sStaffId As String, ByVal sPwdWrongTimes As String)
            Try
                Dim sSQL As String = "UPDATE SY_USERPASSWORD SET PWD_WRONG_TIMES =" & ProviderFactory.PositionPara & "PWD_WRONG_TIMES" & _
                    " WHERE STAFFID =" & ProviderFactory.PositionPara & "STAFFID"

                Dim para(1) As IDataParameter

                para(0) = ProviderFactory.CreateDataParameter("STAFFID", sStaffId)
                para(1) = ProviderFactory.CreateDataParameter("PWD_WRONG_TIMES", sPwdWrongTimes)

                If Me.getDatabaseManager.isTransaction Then
                    DBObject.ExecuteNonQuery(Me.getDatabaseManager.getTransaction, CommandType.Text, sSQL, para)
                Else
                    DBObject.ExecuteNonQuery(Me.getDatabaseManager.getConnection, CommandType.Text, sSQL, para)
                End If
            Catch ex As ProviderException
                Throw ex
            Catch ex As BosException
                Throw ex
            Catch ex As Exception
                Throw ex
            End Try
        End Sub

        ''' <summary>
        ''' 根據主鍵修改密碼,密碼變更時間
        ''' </summary>
        ''' <param name="sStaffid">人員編號</param>
        ''' <param name="sPASSWORD">密碼</param>
        ''' <remarks>
        ''' Zack 2012-06-05 Create
        ''' </remarks>
        Sub updatePWDDATETIME(ByVal sStaffid As String, ByVal sPASSWORD As String)
            Try
                Dim sSQL As String = "UPDATE SY_USERPASSWORD" & _
                                     " SET [PASSWORD] =" & ProviderFactory.PositionPara & "PASSWORD" & _
                                     ",PWD_CHANGE_DATE=getdate()" & _
                                     ",PWD_WRONG_TIMES= 0" & _
                                     " WHERE STAFFID =" & ProviderFactory.PositionPara & "STAFFID"

                Dim para(1) As IDataParameter
                para(0) = ProviderFactory.CreateDataParameter("STAFFID", sStaffid)
                para(1) = ProviderFactory.CreateDataParameter("PASSWORD", sPASSWORD)

                If Me.getDatabaseManager.isTransaction Then
                    DBObject.ExecuteNonQuery(Me.getDatabaseManager.getTransaction, CommandType.Text, sSQL, para)
                Else
                    DBObject.ExecuteNonQuery(Me.getDatabaseManager.getConnection, CommandType.Text, sSQL, para)
                End If
            Catch ex As ProviderException
                Throw ex
            Catch ex As BosException
                Throw ex
            Catch ex As Exception
                Throw ex
            End Try
        End Sub
#End Region
#End Region

    End Class

End Namespace
