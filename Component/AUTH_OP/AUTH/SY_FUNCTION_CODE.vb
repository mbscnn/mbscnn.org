Option Explicit On
Option Strict On

Imports com.Azion.NET.VB

Public Class SY_FUNCTION_CODE
    Inherits BosBase

    ''' <summary>
    ''' Constructor
    ''' </summary>
    ''' <param name="dbManager">DatabaseManager</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal dbManager As com.Azion.NET.VB.DatabaseManager)
        MyBase.new("SY_FUNCTION_CODE", dbManager)
    End Sub


    Sub New()
        MyBase.New()
        Me.setPrimaryKeys()
    End Sub


#Region "濟南昱勝添加"
#Region "Avril Function"

    ''' <summary>
    ''' 根據主鍵查找資料
    ''' </summary>
    ''' <param name="sFuncCode">交易編號</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Add by Avril 2012/04/24
    ''' </remarks>
    Function loadByPK(ByVal sFuncCode As String) As Boolean
        Try
            Dim para(0) As IDbDataParameter
            para(0) = ProviderFactory.CreateDataParameter("FUNCCODE", sFuncCode)

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
    ''' 取得最大的FUNCCODE
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' Add by Avril 2012/04/24
    ''' </remarks>
    Function getMinFuncCode() As Integer
        Try
            If Not IsNothing(Me.getDatabaseManager) Then
                Dim syFunctionCode As New SY_FUNCTION_CODE(Me.getDatabaseManager)
                Dim sSQL As String = "SELECT MIN(FUNCCODE) FUNCCODE  FROM SY_FUNCTION_CODE  "

                If (syFunctionCode.loadBySQL(sSQL)) Then

                    ' 如果“FUNCCODE”不為Nothing
                    If Not syFunctionCode.getAttribute("FUNCCODE") Is Nothing Then
                        If Convert.ToInt32(syFunctionCode.getAttribute("FUNCCODE").ToString()) > 0 Then
                            Return -1
                        Else
                            Return Convert.ToInt32(syFunctionCode.getAttribute("FUNCCODE").ToString()) - 1
                        End If
                    Else
                        Return -1
                    End If
                End If
            End If

            Return -1
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 根據功能名稱查詢資料
    ''' </summary>
    ''' <param name="sFuncName">功能名稱</param>
    ''' <param name="sFuncCode">功能編號</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Add by Avril 2012/04/26
    ''' </remarks>
    Function loadByFuncName(ByVal sFuncName As String, ByVal sFuncCode As String) As Boolean
        Try
            Dim sSQL As String = "SELECT *  FROM SY_FUNCTION_CODE WHERE FUCNAME = " & ProviderFactory.PositionPara & "FUCNAME " & _
                                    " AND FUNCCODE <> " & ProviderFactory.PositionPara & "FUNCCODE "

            Dim paras(1) As System.Data.IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("FUCNAME", sFuncName)
            paras(1) = ProviderFactory.CreateDataParameter("FUNCCODE", sFuncCode)

            Return MyBase.loadBySQL(sSQL, paras)
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 取得最大的排序控制
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function getMaxSortCtrl() As Integer
        Try
            If Not IsNothing(Me.getDatabaseManager) Then
                Dim syFunctionCode As New SY_FUNCTION_CODE(Me.getDatabaseManager)
                Dim sSQL As String = "SELECT MAX(SORTCTRL) SORTCTRL  FROM SY_FUNCTION_CODE  "

                If (syFunctionCode.loadBySQL(sSQL)) Then

                    ' 如果“FUNCCODE”不為Nothing
                    If Not syFunctionCode.getAttribute("SORTCTRL") Is Nothing Then
                        Return Convert.ToInt32(syFunctionCode.getAttribute("SORTCTRL")) + 10
                    End If
                End If
            End If

            Return 1
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 取得最大的RoleId
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    '''  Add by Avril 2012/04/24
    ''' </remarks>
    Function getMaxFuncCode() As Integer
        Try
            If Not IsNothing(Me.getDatabaseManager) Then
                Dim syFunctionCode As New SY_FUNCTION_CODE(Me.getDatabaseManager)
                Dim sSQL As String = "SELECT MAX(FUNCCODE) FUNCCODE  FROM SY_FUNCTION_CODE  "

                If (syFunctionCode.loadBySQL(sSQL)) Then

                    ' 如果“Roleid”不為Nothing
                    If Not syFunctionCode.getAttribute("FUNCCODE") Is Nothing Then
                        If Convert.ToInt32(syFunctionCode.getAttribute("FUNCCODE").ToString()) > 0 Then
                            Return Convert.ToInt32(syFunctionCode.getAttribute("FUNCCODE").ToString())
                        Else
                            Return 1
                        End If
                    Else
                        Return 1
                    End If
                End If
            End If

            Return 1
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 根據資料查詢角色名稱是否存在
    ''' </summary>
    ''' <param name="sFuncName">功能名稱</param>
    ''' <param name="sFuncCode">功能編號</param>
    ''' <param name="sFuncCodeList">臨時表中的功能編號集合</param>
    ''' <returns></returns>
    ''' <remarks>
    '''  Add by Avril 2012/05/15
    ''' </remarks>
    Function loadDataByCon(ByVal sFuncName As String, ByVal sFuncCode As String, ByVal sFuncCodeList As String) As Boolean
        Try
            Dim sSQL As String = "SELECT *  FROM SY_FUNCTION_CODE WHERE FUCNAME = " & ProviderFactory.PositionPara & "FUCNAME " & _
                                " AND FUNCCODE  <> " & ProviderFactory.PositionPara & "FUNCCODE "

            If sFuncCodeList.Length > 0 Then
                sSQL = sSQL & " AND FUNCCODE   IN( " & sFuncCodeList & ") "
            End If

            Dim paras(1) As System.Data.IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("FUCNAME", sFuncName)
            paras(1) = ProviderFactory.CreateDataParameter("FUNCCODE", sFuncCode)

            Return MyBase.loadBySQL(sSQL, paras)
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Function updateDisable(ByVal sFuncCodeList As String, ByVal sDisabled As String) As Boolean
        Try
            Dim sSql As String = String.Empty

            sSql = "UPDATE SY_FUNCTION_CODE SET DISABLED = " & ProviderFactory.PositionPara & "DISABLED " & _
              " WHERE FUNCCODE   IN( " & sFuncCodeList & ") "

            Dim paras(0) As IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("DISABLED", sDisabled)

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
    ''' 根據案件編號查詢區域標識
    ''' </summary>
    ''' <param name="sCaseID">案件編號</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Zack 2012-06-15 Create
    ''' </remarks>
    Function getHoflag(ByVal sCaseID As String) As Boolean
        Try
            Dim sSQL As String = "SELECT SY_FUNCTION_CODE.HOFLAG FROM SY_FUNCTION_CODE  " & _
                                " JOIN  SY_TEMPINFO ON SY_FUNCTION_CODE.FUNCCODE = SY_TEMPINFO .FUNCCODE " & _
                                " WHERE SY_TEMPINFO .CASEID =" & ProviderFactory.PositionPara & "CASEID"

            Dim paras(0) As System.Data.IDbDataParameter
            paras(0) = ProviderFactory.CreateDataParameter("CASEID", sCaseID)

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
#End Region
End Class


