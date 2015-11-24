Option Explicit On
Option Strict On

Imports com.Azion.NET.VB



' ''' <summary>
' ''' 使用者狀態 0在職 9離職
' ''' </summary>
' ''' <remarks></remarks>
'Public Enum ENUM_USER_STATUS
'    ONJOB = 0
'    LEFT = 9
'    ERROR_EXCEPTION = -1
'    ERROR_UNKNOWNSTATUS = -2
'End Enum

Namespace TABLE

    ''' <summary>
    ''' 使用者物件
    ''' </summary>
    ''' <remarks></remarks>
    Public Class SY_USER
        Inherits BosBase

        ''' <summary>
        ''' Constructor
        ''' </summary>
        ''' <param name="dbManager">DatabaseManager</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal dbManager As com.Azion.NET.VB.DatabaseManager)
            MyBase.new("SY_USER", dbManager)
        End Sub


#Region "濟南昱勝添加"
#Region "Lake Function"
        Sub New()
            MyBase.New()
            Me.setPrimaryKeys()
        End Sub

        ''' <summary>
        ''' 檢核用戶輸入編號是否存在
        ''' </summary>
        ''' <param name="sStaffId"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        ''' <history>
        ''' [Lake] 2012/04/11 Created
        '''</history>
        Function checkExsitUserId(ByVal sStaffId As String) As Boolean
            Try
                Dim strSql = "SELECT" _
                           & "    STAFFID" _
                           & "   ,USERNAME " _
                           & "FROM" _
                           & "   SY_USER " _
                           & "WHERE" _
                           & "       STATUS = 0" _
                           & "   AND STAFFID = " & ProviderFactory.PositionPara & "STAFFID"

                If sStaffId.Trim.Length = 5 Then
                    sStaffId = com.Azion.EloanUtility.UIUtility.convRMCD(sStaffId.Trim)
                End If

                Dim para As IDataParameter = ProviderFactory.CreateDataParameter("STAFFID", sStaffId)

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
        ''' 根據主鍵查詢用戶資料
        ''' </summary>
        ''' <param name="sStaffId"></param>
        ''' <returns></returns>
        ''' <remarks>[Lake] 2012/05/07 Created</remarks>
        Public Function loadByPK(ByVal sStaffId As String) As Boolean
            Try
                If sStaffId.Trim.Length = 5 Then
                    sStaffId = com.Azion.EloanUtility.UIUtility.convRMCD(sStaffId.Trim)
                End If
                Dim para As IDataParameter = ProviderFactory.CreateDataParameter("STAFFID", sStaffId)

                Return Me.loadData(para)
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
End Namespace
