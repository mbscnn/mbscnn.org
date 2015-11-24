Option Explicit On
Option Strict On

Imports com.Azion.NET.VB
Imports System.Text
Imports System.Text.RegularExpressions

Namespace TABLE

    Public Class SY_CONDITION_ID
        Inherits SY_TABLEBASE


        ''' <summary>
        ''' Constructor
        ''' </summary>
        ''' <param name="dbManager">DatabaseManager</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal dbManager As com.Azion.NET.VB.DatabaseManager)
            MyBase.new("SY_CONDITION_ID", dbManager)
        End Sub

        ''' <summary>
        ''' 運算條件是否符合
        ''' </summary>
        ''' <param name="sCaseid"></param>
        ''' <param name="nCondId"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function LogicalOperation(ByVal sCaseid As String, ByVal nCondId As Integer) As Boolean

            Dim sCondition As StringBuilder

            Try
                sCondition = New StringBuilder(GetConditionExpression(nCondId))

#If DEBUG Then
                System.Diagnostics.Trace.WriteLine("--- SY_CONDITION ID = " & nCondId & " : " & vbCrLf & sCondition.ToString() & vbCrLf & vbCrLf)
#End If

                Dim bChanged As Boolean = True
                Dim rgx1 As New Regex("\[COND_ID:(?<COND_ID>\d*)\]", RegexOptions.IgnoreCase)
                Dim rgx2 As New Regex("\[(?<TABLE>\w*)\]\.\[(?<FIELD>\w*)\]", RegexOptions.IgnoreCase)

                Do While bChanged = True
                    bChanged = False

                    Dim matches1 As MatchCollection = rgx1.Matches(sCondition.ToString())
                    'Dim listParameters As New BosParamsList

                    For Each match As Match In matches1
                        Dim groups As GroupCollection = match.Groups
                        Dim nInnerCondition As Integer = CInt(groups.Item("COND_ID").Value)

                        sCondition.Replace(match.Value, GetConditionExpression(nInnerCondition))
                        bChanged = True
                    Next

                    Dim matches2 As MatchCollection = rgx2.Matches(sCondition.ToString())
                    'Dim listParameters As New BosParamsList

                    For Each match As Match In matches2
                        Dim groups As GroupCollection = match.Groups

                        '取得TABLE及FILED
                        Dim tableName = groups.Item("TABLE").Value
                        Dim fieldName = groups.Item("FIELD").Value

                        Dim sCondSql As String
                        sCondSql = vbCrLf & _
                            "       (select top 1 " & fieldName & vbCrLf & _
                            "          from " & tableName & vbCrLf & _
                            "         where CASEID = @CASEID@)"

                        sCondition.Replace(match.Value, sCondSql)
                        bChanged = True
                    Next
                Loop

                '將[%!CASEID!%]轉換成sCaseid
                sCondition.Replace("[%!CASEID!%]", sCaseid)

                Dim ObjValue As Object
                ObjValue = ExecuteScalar("select 1 where " & sCondition.ToString() & vbCrLf,
                              "CASEID", sCaseid)

                '如果是空的，表示條件不符合
                If IsDBNull(ObjValue) OrElse IsNothing(ObjValue) Then
                    Return False
                End If

                Return True

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

            Return False
        End Function


        ''' <summary>
        ''' 輸入TABLE, 欄位及CASEID取得TABLE.FIELD內容
        ''' </summary>
        ''' <param name="sTable"></param>
        ''' <param name="sField"></param>
        ''' <param name="sCaseid"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Protected Function GetValueFromTableFieldName(ByVal sTable As String, ByVal sField As String, ByVal sCaseid As String) As Object
            Try
                Dim obj As Object

                obj = ExecuteScalar(
                    "select distinct " & sField & vbCrLf & _
                    "  from " & sTable & vbCrLf & _
                    " where CASEID = @CASEID@" & vbCrLf,
                    "CASEID", sCaseid)

                If IsDBNull(obj) Then
                    Return Nothing
                End If

                Return obj

                Return Nothing
            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)

            End Try
        End Function


        ''' <summary>
        ''' 取得條件字串
        ''' </summary>
        ''' <param name="nCondId"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Protected Function GetConditionExpression(ByVal nCondId As Integer) As String
            Try
                Dim obj As Object

                obj = ExecuteScalar(
                    "select distinct CONDITION " & vbCrLf & _
                    "  from SY_CONDITION_ID " & vbCrLf & _
                    " where SY_CONDITION_ID.COND_ID = @COND_ID@",
                    "COND_ID", nCondId)

                If IsDBNull(obj) Then
                    Throw New SYException(
                        String.Format("無法取得流程步驟的條件內容(SY_CONDITION_ID.CONDITION), SY_CONDITION_ID.COND_ID = {0} ", nCondId),
                                  SYMSG.SYCONDITIONID_CANNOT_GETINFO)
                End If

                Return CStr(obj)
            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)

            End Try
        End Function
    End Class

End Namespace
