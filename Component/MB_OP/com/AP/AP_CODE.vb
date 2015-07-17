Imports com.Azion.NET.VB
Public Class AP_CODE
    Inherits BosBase

    Sub New()
        MyBase.New()
        Me.setPrimaryKeys()
    End Sub

    Sub New(ByVal dbManager As DatabaseManager)
        MyBase.new("AP_CODE", dbManager)
    End Sub

    ''' ----------------------------------------------------------------------------
    ''' <summary>
    ''' 根據[項目序號]取得資料
    ''' </summary>
    ''' <param name="sCode">項目序號</param>
    ''' <returns>Boolean 是否有取得</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' 	[Willy]	2009/11/4	Created
    ''' 	[Willy]	2009/11/12	Modified
    ''' 	[Willy]	2009/11/30	Modified
    ''' </history>
    Function loadByPK(ByVal sCode As String) As Boolean
        Try
            APinitTable.initApCode(Me.getDatabaseManager)

            Dim rows() As DataRow = APinitTable.dtApCodes.Select(String.Format("CODEID='{0}'", sCode))

            For Each row As DataRow In rows
                For Each col As DataColumn In row.Table.Columns
                    Me.setAttribute(col.ColumnName, row(col.ColumnName))
                Next
                Me.setIsLoaded(True)
                Return True
            Next

            Return False
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 根據[父類別]取得層級關係
    ''' </summary>
    ''' <param name="sUpCode">父類別</param>
    ''' <returns>String 取得層級關係</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' 	[Willy]	2009/11/5	Created
    ''' 	[Willy]	2009/11/12	Modified
    ''' </history>
    Function getHeirLevel(ByVal sUpCode As String) As String
        If sUpCode.Equals("0") OrElse Not Utility.isValidateData(sUpCode) Then Return ""

        Dim sb As New Text.StringBuilder
        Dim sText As String

        Try
            APinitTable.initApCode(Me.getDatabaseManager)
            If Me.loadByPK(sUpCode) Then
                sText = Me.getString("TEXT")
                Return sb.Append(getHeirLevel(Me.getString("UPCODE"))).Append(sText).Append(">>").ToString()
            End If
            Return ""
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 取得該[CODEID]最大值
    ''' </summary>
    ''' <returns>Integer 最大值</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' 	[Willy]	2009/11/6 Created
    ''' 	[Willy]	2009/11/12	Modified
    ''' 	[Willy]	2009/11/30	Modified
    ''' </history>
    Function getMaxCodeID() As Integer
        Dim iMax As Integer = 0
        Try
            Dim sqlStr As String = ""
            sqlStr = "select IFNULL(max(CODEID),0) AS CODEID  from AP_CODE "
            iMax = MyBase.loadBySQLScalar(sqlStr)
            Return iMax

        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 根據[UPCODE],[TEXT]取得資料
    ''' </summary>
    ''' <param name="iUpCode">父類別</param>
    ''' <param name="sText">文字</param>
    ''' <returns>Boolean 是否有無取得資料</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' 	[Willy]	2009/11/9 Created
    ''' 	[Willy]	2009/11/12	Modified
    ''' 	[Willy]	2009/11/30	Modified
    ''' 	[Willy]	2009/12/22	Modified
    ''' </history>
    Function loadByText(ByVal iUpCode As Integer, ByVal sText As String) As Boolean
        Try
            APinitTable.initApCode(Me.getDatabaseManager)

            Dim rows() As DataRow = APinitTable.dtEnLinkTabs.Select(String.Format("UPCODE={0} and TEXT='{1}'", iUpCode, sText))

            For Each row As DataRow In rows
                For Each col As DataColumn In row.Table.Columns
                    Me.setAttribute(col.ColumnName, row(col.ColumnName))
                Next
                Me.setIsLoaded(True)
                Return True
            Next
            Return False
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 根據[SUBSYSID],[UPCODE],[TEXT]取得資料
    ''' </summary>
    ''' <param name="sSubSys">子系統</param>
    ''' <param name="iUpCode">父類別</param>
    ''' <param name="sText">文字</param>
    ''' <returns>Boolean 是否有無取得資料</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' 	[Willy]	2009/12/22	Created
    ''' </history>
    Function loadByText(ByVal sSubSys As String, ByVal iUpCode As Integer, ByVal sText As String) As Boolean
        Try
            APinitTable.initApCode(Me.getDatabaseManager)

            Dim rows() As DataRow = APinitTable.dtApCodes.Select(String.Format("SUBSYSID='{0}' and UPCODE={1}  and TEXT='{2}'", sSubSys, iUpCode, sText))

            For Each row As DataRow In rows
                For Each col As DataColumn In row.Table.Columns
                    Me.setAttribute(col.ColumnName, row(col.ColumnName))
                Next
                Me.setIsLoaded(True)
                Return True
            Next
            Return False

        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 儲存後清空Static Object
    ''' </summary>
    ''' <returns>Integer </returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' 	[Willy]	2009/11/18  Created
    ''' 	[Willy]	2009/11/30	Modified
    ''' </history>
    Shadows Function save() As Integer
        Dim iReturn As Integer = MyBase.save()
        'APinitTable.clean()
        APinitTable.dtApCodes = Nothing
        Return iReturn
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 刪除後清空Static Object
    ''' </summary>
    ''' <returns>Integer </returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' 	[Willy]	2009/11/18  Created
    ''' 	[Willy]	2009/11/30	Modified
    ''' </history>
    Shadows Function remove() As Boolean
        Dim bReturn As Boolean = MyBase.remove()
        APinitTable.dtApCodes = Nothing
        Return bReturn
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 根據[CODEID]取得資料
    ''' 更新、插入、刪除專用
    ''' </summary>
    ''' <param name="sCodeID">CodeID</param>
    ''' <returns>Boolean 是否有無取得資料</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' 	[Willy]	2009/11/18 Created
    ''' </history>
    Function loadByPkForUpdate(ByVal sCodeID As String) As Boolean
        Dim paras(0) As IDataParameter

        Try
            paras(0) = ProviderFactory.CreateDataParameter("CODEID", sCodeID)
            Return MyBase.loadData(paras)
        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

#Region "Rachel Function"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 取得該[UPCODE]SORTNO=1
    ''' </summary>
    ''' <returns>Integer 第一順位值</returns>
    ''' <remarks>
    ''' </remarks> 
    ''' <history>
    ''' 	[Rachel]	2009/11/16 Created
    ''' 	[Willy]	    2009/11/18 Modified
    ''' 	[Willy]	    2009/11/30	Modified
    ''' </history>
    Function loadFirstUpcode(ByVal iUpCode As Integer) As Boolean

        Try
            APinitTable.initApCode(Me.getDatabaseManager)

            Dim rows() As DataRow = APinitTable.dtApCodes.Select(String.Format("UPCODE={0}  and SORTNO=1", iUpCode))

            For Each row As DataRow In rows
                For Each col As DataColumn In row.Table.Columns
                    Me.setAttribute(col.ColumnName, row(col.ColumnName))
                Next
                Return True
            Next
            Return False

        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try

    End Function


#End Region

#Region "MikeKuan Function"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 查詢海外分行的PK
    ''' </summary>
    ''' <param name="sUPCODE">444</param>
    ''' <param name="sBRID">分行代碼</param>
    ''' <returns>Boolean 是否有取得</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[MikeKuan]	2010/11/18 Created
    ''' </history>
    '''

    Function getDATEbyUPCODEandVALUE(ByVal sUPCODE As String, ByVal sBRID As String) As Boolean

        If Utility.isValidateData(sUPCODE) Then
            Try
                Dim sSQL As String = "SELECT *" & vbNewLine & _
                                    "  FROM AP_CODE" & vbNewLine & _
                                    " WHERE UPCODE = " & ProviderFactory.PositionPara & "UPCODE " & _
                                    "   AND VALUE = " & ProviderFactory.PositionPara & "VALUE "

                Dim paras(1) As System.Data.IDbDataParameter
                paras(0) = ProviderFactory.CreateDataParameter("UPCODE", sUPCODE)
                paras(1) = ProviderFactory.CreateDataParameter("VALUE", sBRID)


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

#End Region

#Region "Dora"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 取得TEXT
    ''' </summary>
    ''' <param name="iUPCODE">父類別</param>
    ''' <param name="sValue">值</param>
    ''' <returns>String TEXT</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Dora]	2011/07/28 Created
    ''' </history>
    '''
    Function getTEXT(ByVal iUPCODE As Integer, ByVal sValue As String) As String
        Try
           

            APinitTable.initApCode(Me.getDatabaseManager)

            Dim rows() As DataRow = APinitTable.dtApCodes.Select(String.Format("UPCODE={0}  and VALUE='{1}'", iUPCODE, sValue))

            If rows.Length() > 0 Then
                For Each col As DataColumn In rows(0).Table.Columns
                    Me.setAttribute(col.ColumnName, rows(0)(col.ColumnName))
                Next 
                Return rows(0)("TEXT")
            End If
            Return "" 

        Catch ex As ProviderException
            Throw ex
        Catch ex As BosException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Function

#End Region

End Class
