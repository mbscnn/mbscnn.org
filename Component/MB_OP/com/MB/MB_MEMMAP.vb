﻿Imports com.Azion.NET.VB

Public Class MB_MEMMAP
    Inherits BosBase

    Sub New()
        MyBase.New()
        Me.setPrimaryKeys()
    End Sub

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.New("MB_MEMMAP", dbManager)
    End Sub

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
    Public Function LoadByPK(ByVal iMB_MEMSEQ As Decimal) As Boolean
        Try
            Dim sqlStr As String = String.Empty

            sqlStr = "SELECT * FROM MB_MEMMAP " &
                     "  WHERE MB_MEMSEQ = " & ProviderFactory.PositionPara & "MB_MEMSEQ "

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
#End Region

End Class
