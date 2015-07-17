Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Data
Imports System.Diagnostics

'#If __MSSQL Then
Imports System.Data.SqlClient
'#End If


Namespace TinyDb

    Partial Public Class DbTable
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

            '不需延後處理
            If deferredExecute = False Then
                Dim nDeferredValue As Integer
                nDeferredValue = ExecuteDeferedSql()

                nValue = Update(tableName, parameters, deferredExecute)

                If nValue = 0 Then
                    nValue = Insert(tableName, parameters, deferredExecute)
                End If

                Debug.Assert(nValue > 0)
                '應該至少會有一行被更新或新增
                Return nValue + nDeferredValue
            End If


            Dim listPk As New List(Of CmdParameter)()


            For Each param In parameters
                If param.parameterIsPK Then
                    listPk.Add(param)
                End If
            Next

            If listPk.Count = 0 Then        ' 找不到PK，只能新增資料
                nValue = Insert(tableName, parameters, deferredExecute)
                Debug.Assert(deferredExecute = False And nValue > 0) '應該至少會有一行被更新或新增
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
            strInsertUpdate1.Clear()
            strInsertUpdate2.Clear()
            parametersInsertUpdate.Clear()
            strInsertUpdate = ""

            For Each param In parameters
                If Not IsNothing(param.parameterValue) Then
                    If strInsertUpdate1.Length <> 0 Then
                        strInsertUpdate1.Append(" , ")
                        strInsertUpdate2.Append(" , ")
                    End If
                    strInsertUpdate1.Append(" " & param.parameterName & " ")
                    strInsertUpdate2.Append(" @" & param.parameterName & "@ ")
                    parametersInsertUpdate.Add(param)
                End If
            Next

            Debug.Assert(strInsertUpdate1.Length > 0)
            Debug.Assert(strInsertUpdate2.Length > 0)

            strInsertUpdate = _
                " INSERT INTO " & tableName & " ( " & strInsertUpdate1.ToString() & _
                " ) VALUE ( " & strInsertUpdate2.ToString() & " )"

            nValue = ExecuteNonQuery(strInsertUpdate, parametersInsertUpdate.ToArray(), Nothing, deferredExecute)

            'Debug.Assert(nValue > 0)    '應該至少會有一行被更新或新增
            Return nValue

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

            Dim strInsertUpdate1, strInsertUpdate2 As New StringBuilder
            Dim parametersInsertUpdate As New List(Of CmdParameter)
            Dim strInsertUpdate As String
            Dim nValue As Integer

            ' 更新資料, 組SQL
            For Each param In listNotPk
                If Not IsNothing(param.parameterValue) Then
                    If strInsertUpdate1.Length <> 0 Then
                        strInsertUpdate1.Append(" , ")
                    End If
                    strInsertUpdate1.Append(" " & param.parameterName & " = @" & param.parameterName & "@ ")
                    parametersInsertUpdate.Add(param)
                End If
            Next

            If listPk.Count = 0 Then
                Throw New Exception("找沒有Primary Key，需檢查傳入的parameters是否有設定PK。")
            End If

            For Each param In listPk
                If Not IsNothing(param.parameterValue) Then
                    If strInsertUpdate2.Length > 0 Then
                        strInsertUpdate2.Append(" AND ")
                    End If
                    strInsertUpdate2.Append(" " & param.parameterName & " = @" & param.parameterName & "@ ")
                    parametersInsertUpdate.Add(param)
                End If
            Next

            Debug.Assert(strInsertUpdate1.Length > 0)
            Debug.Assert(strInsertUpdate2.Length > 0)

            strInsertUpdate = _
                " UPDATE " & tableName & " SET " & strInsertUpdate1.ToString() & _
                " WHERE " & strInsertUpdate2.ToString()

            nValue = ExecuteNonQuery(strInsertUpdate, parametersInsertUpdate.ToArray(), Nothing, deferredExecute)
            Return nValue
        End Function

        Public Function ExecuteDeferedSql() As Integer
            If DeferredSQLCount > 0 Then
                Return ExecuteNonQuery(Nothing, Nothing, CommandType.Text, False)
            Else
                Return 0
            End If
        End Function

    End Class

End Namespace