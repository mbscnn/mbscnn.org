Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Data
Imports System.Diagnostics

'#If __MSSQL Then
Imports System.Data.SqlClient
Imports com.Azion.EloanUtility

'#End If


Namespace TinyDb

    Partial Public Class DbTable

        ''' <summary>
        ''' 若已有記錄，則更新記錄。否則新增記錄。
        ''' </summary>
        ''' <param name="parameters"></param>
        ''' <param name="deferredExecute">
        ''' 要排入佇列時，將deferredExecute設為TRUE
        ''' 要執行所有的SQL時，將deferredExecute設為FALSE
        ''' </param>
        ''' <returns>受影響的資料列數目。</returns>
        ''' <remarks>會以PK來判斷是否是新記錄還是舊記錄</remarks>
        Public Function InsertUpdate(ByVal parameters As CmdParameter(), Optional ByVal deferredExecute As Boolean = False) As Integer
#If DEBUG Then
            Dbg.Assert(IsNothing(m_strTableName) = False AndAlso m_strTableName.Length > 0)
#End If
            Return InsertUpdate(m_strTableName, parameters, deferredExecute)
        End Function


        ''' <summary>
        ''' 若已有記錄，則更新記錄。否則新增記錄。
        ''' </summary>
        ''' <param name="tableName"></param>
        ''' <param name="parameters"></param>
        ''' <param name="deferredExecute">延後處理，可一次處理多個SQL語法，
        ''' 要排入佇列時，將deferredExecute設為TRUE
        ''' 要執行所有的SQL時，將deferredExecute設為FALSE
        ''' </param>
        ''' <returns>受影響的資料列數目。</returns>
        ''' <remarks>會以PK來判斷是否是新記錄還是舊記錄</remarks>
        ''' <SAMPLE>
        ''' </SAMPLE>
        Public Function InsertUpdate(tableName As String, parameters As CmdParameter(), Optional deferredExecute As Boolean = False) As Integer
            Dim nValue As Integer = 0

            '不需延後處理且沒有任何的延後執行的SQL命令存放在QUEUE內
            If deferredExecute = False AndAlso m_queueDeferredSQL.Count = 0 Then

                nValue = Update(tableName, parameters, deferredExecute)

                If nValue = 0 Then
                    nValue = Insert(tableName, parameters, deferredExecute)
                End If

#If DEBUG Then
                Dbg.Assert(nValue > 0)
                '應該至少會有一行被更新或新增
#End If
                Return nValue
            End If


            Dim listPk As New List(Of CmdParameter)()


            For Each param In parameters
                If param.parameterIsPK Then
                    listPk.Add(param)
                End If
            Next

            If listPk.Count = 0 Then        ' 找不到PK，只能新增資料
                nValue = Insert(tableName, parameters, deferredExecute)
#If DEBUG Then
                Dbg.Assert(deferredExecute = False And nValue > 0, "應該至少會有一行被更新或新增") '應該至少會有一行被更新或新增
#End If

                Return nValue
            End If

            '判斷是否已有資料
            Dim strWhere As New StringBuilder
            Dim parametersWhere As New List(Of CmdParameter)

            For Each param In listPk
                If Not IsNothing(param.parameterValue) Then
                    If strWhere.Length > 0 Then
                        strWhere.Append(" AND ")
                    End If
                    strWhere.Append(" " & param.parameterName & " = @" & param.parameterName & "@ ")
                    parametersWhere.Add(param)
                End If
            Next

            nValue = ExecuteScalar("SELECT COUNT(*) FROM " & tableName & _
                                   " WHERE" & strWhere.ToString(), _
                                   parametersWhere.ToArray())

            If nValue = 0 Then
                nValue = Insert(tableName, parameters, deferredExecute)
            Else
                nValue = Update(tableName, parameters, deferredExecute)
            End If

            Return nValue

        End Function

        ''' <summary>
        ''' 新增記錄
        ''' </summary>
        ''' <param name="parameters"></param>
        ''' <param name="deferredExecute"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Insert(ByVal parameters As CmdParameter(), _
                               Optional ByVal deferredExecute As Boolean = False) As Integer

#If DEBUG Then
            Dbg.Assert(IsNothing(m_strTableName) = False AndAlso m_strTableName.Length > 0)
#End If
            Return Insert(m_strTableName, parameters, deferredExecute)
        End Function


        ''' <summary>
        ''' 新增記錄
        ''' </summary>
        ''' <param name="tableName"></param>
        ''' <param name="parameters"></param>
        ''' <param name="deferredExecute">延後處理，可一次處理多個SQL語法，
        ''' 要排入佇列時，將deferredExecute設為TRUE
        ''' 要執行所有的SQL時，將deferredExecute設為FALSE
        ''' </param>
        ''' <returns>受影響的資料列數目</returns>
        ''' <remarks></remarks>
        Public Function Insert(ByVal tableName As String, _
                                     ByVal parameters As CmdParameter(), _
                                     Optional ByVal deferredExecute As Boolean = False) As Integer

            Dim strInsertUpdate1, strInsertUpdate2 As New StringBuilder
            Dim parametersInsertUpdate As New List(Of CmdParameter)
            Dim strInsertUpdate As String
            Dim nValue As Integer

            ' 新增資料, 組SQL
            'strInsertUpdate1.Clear()
            'strInsertUpdate2.Clear()
            'parametersInsertUpdate.Clear()
            'strInsertUpdate = ""

            For Each param In parameters
                If strInsertUpdate1.Length <> 0 Then
                    strInsertUpdate1.Append(" , ")
                    strInsertUpdate2.Append(" , ")
                End If
                strInsertUpdate1.Append(" " & param.parameterName & " ")
                strInsertUpdate2.Append(" @" & param.parameterName & "@ ")
                parametersInsertUpdate.Add(param)
            Next

#If DEBUG Then
            Dbg.Assert(strInsertUpdate1.Length > 0)
            Dbg.Assert(strInsertUpdate2.Length > 0)
#End If

            strInsertUpdate = _
                " INSERT INTO " & tableName & " ( " & strInsertUpdate1.ToString() & _
                " ) VALUES ( " & strInsertUpdate2.ToString() & " )"

            nValue = ExecuteNonQuery(strInsertUpdate, parametersInsertUpdate.ToArray(), Nothing, deferredExecute)

            'Dbg.Assert(nValue > 0)    '應該至少會有一行被更新或新增
            Return nValue

        End Function

        ''' <summary>
        ''' 更新記錄。
        ''' </summary>
        ''' <param name="parameters"></param>
        ''' <param name="deferredExecute"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Update(ByVal parameters As CmdParameter(), _
                               Optional ByVal deferredExecute As Boolean = False) As Integer
#If DEBUG Then
            Dbg.Assert(IsNothing(m_strTableName) = False AndAlso m_strTableName.Length > 0)
#End If
            Return Update(m_strTableName, parameters, deferredExecute)
        End Function

        ''' <summary>
        ''' 更新記錄。
        ''' </summary>
        ''' <param name="tableName"></param>
        ''' <param name="parameters"></param>
        ''' <param name="deferredExecute">延後處理，可一次處理多個SQL語法，
        ''' 要排入佇列時，將deferredExecute設為TRUE
        ''' 要執行所有的SQL時，將deferredExecute設為FALSE
        ''' </param>
        ''' <returns>受影響的資料列數目。</returns>
        ''' <remarks></remarks>
        ''' <SAMPLE>
        ''' </SAMPLE>
        Public Function Update(ByVal tableName As String, _
                                     ByVal parameters As CmdParameter(), _
                                     Optional ByVal deferredExecute As Boolean = False) As Integer
            Dim listPk, listNotPk As New List(Of CmdParameter)
            ''Dim strInsert, strUpdate As String

            '區分PK及非PK
            For Each param In parameters
                If param.parameterIsPK Then
                    listPk.Add(param)
                Else
                    listNotPk.Add(param)
                End If
            Next

            Dim strInsertUpdate1, strInsertUpdate2, strInsertUpdate3, strInsertUpdate4 As New StringBuilder
            Dim parametersInsertUpdate As New List(Of CmdParameter)
            Dim strInsertUpdate As String
            Dim nValue As Integer

            ' 更新資料, 組SQL
            For Each param In listNotPk
                If strInsertUpdate1.Length <> 0 Then
                    strInsertUpdate1.Append(" , ")
                End If
                strInsertUpdate1.Append(" " & param.parameterName & " = @" & param.parameterName & "@ ")
                parametersInsertUpdate.Add(param)
            Next

            If listPk.Count = 0 Then
                Throw New Exception("找沒有Primary Key，需檢查傳入的parameters是否有設定PK。")
            End If

            For Each param In listPk
                If IsNothing(param.parameterValue) Then
#If DEBUG Then
                    Dbg.Assert(IsNothing(param.parameterValue) = False, "PK欄位的值不可為NULL")
#End If
                    Throw New Exception("PK欄位的值不可為NULL: [" & param.parameterName & "]")
                End If

                If strInsertUpdate2.Length > 0 Then
                    strInsertUpdate2.Append(" AND ")
                End If
                strInsertUpdate2.Append(" " & param.parameterName & " = @" & param.parameterName & "@ ")
                parametersInsertUpdate.Add(param)
            Next

#If DEBUG Then
            Dbg.Assert(strInsertUpdate1.Length > 0)
            Dbg.Assert(strInsertUpdate2.Length > 0)
#End If
            strInsertUpdate = _
                " UPDATE " & tableName & " SET " & strInsertUpdate1.ToString() & _
                " WHERE " & strInsertUpdate2.ToString()

            Try
                nValue = ExecuteNonQuery(strInsertUpdate, parametersInsertUpdate.ToArray(), Nothing, deferredExecute)
            Catch ex As Exception
                Dim sSQLErr As String = String.Empty
                For Each param In listNotPk
                    If strInsertUpdate3.Length <> 0 Then
                        strInsertUpdate3.Append(" , ")
                    End If
                    strInsertUpdate3.Append(" " & param.parameterName & " = " & param.parameterValue & " ")
                Next

                For Each param In listPk
                    If strInsertUpdate4.Length > 0 Then
                        strInsertUpdate4.Append(" AND ")
                    End If
                    strInsertUpdate4.Append(" " & param.parameterName & " = " & param.parameterValue & " ")
                Next

                sSQLErr = " UPDATE " & tableName & " SET " & strInsertUpdate3.ToString() & _
                          " WHERE " & strInsertUpdate4.ToString()

                Throw New Exception(Me.getMsg(ex) & "【" & sSQLErr & "】")
            End Try

            Return nValue
        End Function

        Private Function getMsg(ByVal ex As Exception) As String
            Dim sb As New System.Text.StringBuilder
            Do
                If Not IsNothing(ex.Message) Then
                    sb.Append(ex.Message.ToString)
                End If

                If Not IsNothing(ex.StackTrace) Then
                    If Titan.Utility.isValidateData(sb.ToString) Then
                        sb.Append("<br>")
                    End If

                    sb.Append(ex.StackTrace.ToString & "。<br>")
                End If

                ex = ex.InnerException
            Loop Until (ex Is Nothing)

            Return sb.ToString
        End Function


        ''' <summary>
        ''' 取得記錄筆數。
        ''' </summary>
        ''' <param name="parameters"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetCount(ByVal parameters As CmdParameter()) As Integer

#If DEBUG Then
            Dbg.Assert(IsNothing(m_strTableName) = False AndAlso m_strTableName.Length > 0)
#End If

            Return GetCount(m_strTableName, parameters)
        End Function



        ''' <summary>
        ''' 刪除記錄
        ''' </summary>
        ''' <param name="parameters"></param>
        ''' <param name="deferredExecute">延後處理，可一次處理多個SQL語法，
        ''' 要排入佇列時，將deferredExecute設為TRUE
        ''' 要執行所有的SQL時，將deferredExecute設為FALSE
        ''' </param>
        ''' <returns>受影響的資料列數目</returns>
        ''' <remarks></remarks>
        Public Function Delete(ByVal parameters As CmdParameter(), _
                       Optional ByVal deferredExecute As Boolean = False) As Integer
#If DEBUG Then
            Dbg.Assert(IsNothing(m_strTableName) = False AndAlso m_strTableName.Length > 0)
#End If
            Return Delete(m_strTableName, parameters, deferredExecute)
        End Function



        ''' <summary>
        ''' 刪除記錄
        ''' </summary>
        ''' <param name="tableName"></param>
        ''' <param name="parameters"></param>
        ''' <param name="deferredExecute">延後處理，可一次處理多個SQL語法，
        ''' 要排入佇列時，將deferredExecute設為TRUE
        ''' 要執行所有的SQL時，將deferredExecute設為FALSE
        ''' </param>
        ''' <returns>受影響的資料列數目</returns>
        ''' <remarks></remarks>
        Public Function Delete(ByVal tableName As String, _
                                     ByVal parameters As CmdParameter(), _
                                     Optional ByVal deferredExecute As Boolean = False) As Integer

            Dim sbDelete As New StringBuilder
            Dim parametersDelete As New List(Of CmdParameter)
            Dim strDelete As String
            Dim nValue As Integer

            ' 新增資料, 組SQL

            For Each param In parameters
                If Not IsNothing(param.parameterValue) Then
                    If sbDelete.Length <> 0 Then
                        sbDelete.Append(" , ")
                    End If

                    sbDelete.Append(" " & param.parameterName & " = @" & param.parameterName & "@ ")
                    parametersDelete.Add(param)
                End If
            Next

#If DEBUG Then
            Dbg.Assert(sbDelete.Length > 0)
#End If
            If sbDelete.Length = 0 Then
                Throw New Exception("刪除必須要有條件")
            End If

            strDelete = _
                " DELETE FROM " & tableName & " WHERE " & sbDelete.ToString()

            nValue = ExecuteNonQuery(strDelete, parametersDelete.ToArray(), Nothing, deferredExecute)

            Return nValue

        End Function


        ''' <summary>
        ''' 取得記錄筆數。
        ''' </summary>
        ''' <param name="tableName"></param>
        ''' <param name="parameters"></param>
        ''' <returns>資料列數目。</returns>
        ''' <remarks></remarks>
        ''' <SAMPLE>
        ''' </SAMPLE>
        Public Function GetCount(ByVal tableName As String, _
                                     ByVal parameters As CmdParameter()) As Integer
            Dim listFilter As New List(Of CmdParameter)
            ''Dim strInsert, strUpdate As String

            For Each param In parameters
                listFilter.Add(param)
            Next

            Dim sbFilter As New StringBuilder
            Dim parametersFilter As New List(Of CmdParameter)
            Dim nValue As Integer

            ' 更新資料, 組SQL

            sbFilter.Append(" SELECT COUNT(*) FROM " & tableName)

            If listFilter.Count > 0 Then
                'sbFilter.Append(" WHERE ")

                Dim sbFilterWhere As New StringBuilder

                For Each param In listFilter
                    If Not IsNothing(param.parameterValue) Then
                        If sbFilterWhere.Length <> 0 Then
                            sbFilterWhere.Append(" AND ")
                        End If
                        sbFilterWhere.Append(" " & param.parameterName & " = @" & param.parameterName & "@ ")
                        parametersFilter.Add(param)
                    End If
                Next

                If (sbFilterWhere.Length > 0) Then
                    sbFilter.Append(" WHERE ")
                    sbFilter.Append(sbFilterWhere)
                End If
            End If

            nValue = ExecuteScalar(sbFilter.ToString(), parametersFilter.ToArray())
            Return nValue
        End Function


        ''' <summary>
        ''' 取得DataRowCollection。若不為NOTHING，則一定至少有一筆資料
        ''' </summary>
        ''' <param name="strSQL"></param>
        ''' <param name="parameters"></param>
        ''' <param name="commandType"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetDataRowCollection(ByVal strSQL As String, _
                                     ByVal parameters As CmdParameter(), _
                                     Optional ByVal commandType As CommandType = Nothing) As DataRowCollection

            Dim dt As DataTable
            dt = GetDataTable(strSQL, parameters, commandType)

            Dbg.Assert(Not IsNothing(dt))
            If IsNothing(dt) Then
                Return Nothing
            End If

            If IsNothing(dt.Rows) Then
                Return Nothing
            End If

            If dt.Rows.Count = 0 Then
                Return Nothing
            End If

            Return dt.Rows

        End Function


        ''' <summary>
        ''' 取得DataRow。若不為NOTHING，則一定只有一筆資料
        ''' </summary>
        ''' <param name="strSQL"></param>
        ''' <param name="parameters"></param>
        ''' <param name="commandType"></param>
        ''' <returns>若不為NOTHING，則傳回一筆資料</returns>
        ''' <remarks></remarks>
        Public Function GetDataRow(ByVal strSQL As String, _
                                     ByVal parameters As CmdParameter(), _
                                     Optional ByVal commandType As CommandType = Nothing) As DataRow

            Dim drc As DataRowCollection
            drc = GetDataRowCollection(strSQL, parameters, commandType)

            Dbg.Assert(Not IsNothing(drc))
            If IsNothing(drc) Then
                Return Nothing
            End If

#If DEBUG Then
            '若有多筆資料則警告，(可能QUERY語法不正確導致多筆資料)
            Debug.Assert(drc.Count = 1, "資料筆數應等於1")
#End If
            Return (drc(0))

        End Function


        ''' <summary>
        ''' 立即執行所有延後執行的SQL命令
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ExecuteDeferedSql() As Integer
            If DeferredSQLCount > 0 Then
                Return ExecuteNonQuery(Nothing, Nothing, CommandType.Text, False)
            Else
                Return 0
            End If
        End Function


        ''' <summary>
        ''' 清除所有延後執行的SQL命令
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub ClearDeferedSql()
            m_queueDeferredSQL.Clear()
        End Sub

    End Class

End Namespace