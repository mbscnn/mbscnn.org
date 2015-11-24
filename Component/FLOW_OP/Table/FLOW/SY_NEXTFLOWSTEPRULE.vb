Option Explicit On
Option Strict On

Imports com.Azion.NET.VB


Public Class BRADEPNO_ROLEID
    Public braDepNo As Integer
    Public roleId As Integer

    Public Shared Function Pair(ByVal nBraDepNo As Integer, ByVal nRoleId As Integer) As BRADEPNO_ROLEID
        Dim br As New BRADEPNO_ROLEID
        br.braDepNo = nBraDepNo
        br.roleId = nRoleId
        Return br
    End Function
End Class


Namespace TABLE

    Public Class SY_NEXTFLOWSTEPRULE
        Inherits SY_TABLEBASE

        ''' <summary>
        ''' Constructor
        ''' </summary>
        ''' <param name="dbManager">DatabaseManager</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal dbManager As com.Azion.NET.VB.DatabaseManager)
            MyBase.new("SY_NEXTFLOWSTEPRULE", dbManager)
        End Sub




        ''' <summary>
        ''' 設定案件步驟只能指定給哪一個部門及角色
        ''' </summary>
        ''' <param name="sCaseid"></param>
        ''' <param name="sStepNo"></param>
        ''' <param name="nBraDepNo">可為0，表示可以指定所有的部門</param>
        ''' <param name="nRoleId">可為0，表示可以指定所有的角色</param>
        ''' <remarks>nBraDepNo及nRoleId不可同時為0</remarks>
        Public Function AddRule(ByVal sCaseid As String,
                           ByVal sStepNo As String,
                           ByVal nBraDepNo As Integer,
                           ByVal nRoleId As Integer) As Integer
            Dbg.Assert(nBraDepNo <> 0 OrElse nRoleId <> 0, "nBraDepNo及nRoleId不可同時為0")

            Dim arrayBosParameters As BosParamsList

            Try
                Static Dim nCount As Integer = 0
                nCount = nCount + 1

                '每執行一百次新增清理一次所有已結案的條件()
                If nCount Mod 100 = 0 Then
                    DeleteAllPossibleRules()
                End If
            Catch ex As Exception
            End Try

            Try
                arrayBosParameters = New BosParamsList
                arrayBosParameters.Add("CASEID", sCaseid)
                arrayBosParameters.Add("STEP_NO", sStepNo)

                If nBraDepNo <> 0 Then
                    arrayBosParameters.Add("BRA_DEPNO", nBraDepNo)
                End If

                If nRoleId <> 0 Then
                    arrayBosParameters.Add("ROLEID", nRoleId)
                End If

                Return InsertUpdate(arrayBosParameters.ToArray)
            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try
        End Function

        ''' <summary>
        ''' 刪除 設定案件步驟只能指定給哪一個部門及角色
        ''' </summary>
        ''' <param name="sCaseid"></param>
        ''' <param name="sStepNo">可為nothing</param>
        ''' <param name="nBraDepNo">可為0</param>
        ''' <param name="nRoleId">可為0</param>
        ''' <remarks></remarks>
        Public Function DelRule(ByVal sCaseid As String,
                   Optional ByVal sStepNo As String = Nothing,
                   Optional ByVal nBraDepNo As Integer = 0,
                   Optional ByVal nRoleId As Integer = 0) As Integer
            Dim arrayBosParameters As BosParamsList

            Try
                arrayBosParameters = New BosParamsList
                arrayBosParameters.Add("CASEID", sCaseid)

                If String.IsNullOrEmpty(sStepNo) = False Then
                    arrayBosParameters.Add("STEP_NO", sStepNo)
                End If

                If nBraDepNo <> 0 Then
                    arrayBosParameters.Add("BRA_DEPNO", nBraDepNo)
                End If

                If nRoleId <> 0 Then
                    arrayBosParameters.Add("ROLEID", nRoleId)
                End If

                Return Delete(arrayBosParameters.ToArray)
            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try
        End Function


        Public Function GetRules(ByVal sCaseid As String, ByVal sStepNo As String) As BRADEPNO_ROLEID()

            Dim drc As DataRowCollection
            Dim listBradepnoRoleid As New List(Of BRADEPNO_ROLEID)

            Try
                drc = GetDataRowCollection( _
                                        "select BRA_DEPNO, ROLEID " & vbCrLf & _
                                        "  from SY_NEXTFLOWSTEPRULE " & vbCrLf & _
                                        " where CASEID = @CASEID@ " & vbCrLf & _
                                        "   and STEP_NO = @STEP_NO@ " & vbCrLf,
                                        "CASEID", sCaseid,
                                        "STEP_NO", sStepNo)

                If Not IsNothing(drc) Then
                    For Each dr As DataRow In drc
                        listBradepnoRoleid.Add(BRADEPNO_ROLEID.Pair(
                                               CDbType(Of Integer)(dr("BRA_DEPNO"), 0),
                                               CDbType(Of Integer)(dr("ROLEID"), 0)))
                    Next
                End If

                Return listBradepnoRoleid.ToArray()

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Function



        ''' <summary>
        ''' 取
        ''' </summary>
        ''' <param name="sCaseid"></param>
        ''' <param name="sStepNo"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetUsersByRule(ByVal sCaseid As String, ByVal sStepNo As String) As List(Of ROLE_BRADEPNO_STAFFID)

            Dim drc As DataRowCollection
            Dim listRoleidBradepnoRoleid As New List(Of ROLE_BRADEPNO_STAFFID)

            Try
                drc = GetDataRowCollection( _
                                        "select NFR.ROLEID, NFR.BRA_DEPNO, RRU.STAFFID " & vbCrLf & _
                                        "  from SY_NEXTFLOWSTEPRULE NFR " & vbCrLf & _
                                        "  left join SY_REL_ROLE_USER RRU " & vbCrLf & _
                                        "    on NFR.ROLEID = RRU.ROLEID " & vbCrLf & _
                                        "   and NFR.BRA_DEPNO = RRU.BRA_DEPNO " & vbCrLf & _
                                        " where CASEID = @CASEID@ " & vbCrLf & _
                                        "   and STEP_NO = @STEP_NO@ " & vbCrLf,
                                        "CASEID", sCaseid,
                                        "STEP_NO", sStepNo)

                If Not IsNothing(drc) Then
                    For Each dr As DataRow In drc
                        listRoleidBradepnoRoleid.Add(ROLE_BRADEPNO_STAFFID.Pair(
                                                CDbType(Of Integer)(dr("ROLEID"), 0),
                                                CDbType(Of Integer)(dr("BRA_DEPNO"), 0),
                                                CDbType(Of String)(dr("STAFFID"), Nothing)
                                               ))
                    Next
                End If

                Return listRoleidBradepnoRoleid

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Function


        Public Function FilterBrid(ByVal sCaseid As String, ByVal sStepNo As String, inputBrid As BRADEPNO_BRID) As BRADEPNO_BRID
            Dim inputBrids() As BRADEPNO_BRID

            inputBrids = ToArray(inputBrid)
            inputBrids = FilterBrid(sCaseid, sStepNo, inputBrids)

            If IsNothing(inputBrids) OrElse inputBrids.Count = 0 Then
                Return Nothing
            End If

            Return inputBrids(0)
        End Function


        Public Function FilterBrid(ByVal sCaseid As String, ByVal sStepNo As String, inputBrids() As BRADEPNO_BRID) As BRADEPNO_BRID()

            Dim listRoleBradepnoStaffid As List(Of ROLE_BRADEPNO_STAFFID)
            Dim listBrapdeno_Brid1 As BRADEPNO_BRID()
            Dim listBrapdeno_Brid2 As List(Of BRADEPNO_BRID)

            Try
                If IsNothing(inputBrids) Then
                    Return inputBrids
                End If

                listRoleBradepnoStaffid = GetUsersByRule(sCaseid, sStepNo).ToList
                If listRoleBradepnoStaffid.Count = 0 Then
                    Return inputBrids
                End If


                listBrapdeno_Brid1 = inputBrids
                listBrapdeno_Brid2 = New List(Of BRADEPNO_BRID)

                For Each item1 As BRADEPNO_BRID In listBrapdeno_Brid1
                    For Each item2 As ROLE_BRADEPNO_STAFFID In listRoleBradepnoStaffid
                        If item1.BraDepNo = item2.BraDepNo OrElse item2.BraDepNo = 0 Then
                            listBrapdeno_Brid2.Add(item1)
                            Exit For
                        End If
                    Next
                Next

                Return listBrapdeno_Brid2.ToArray

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Function


        Public Function FilterUser(ByVal sCaseid As String, ByVal sStepNo As String, inputUser As USER_ID_NAME) As USER_ID_NAME

            Dim listRoleBradepnoStaffid As List(Of ROLE_BRADEPNO_STAFFID)

            Try
                If IsNothing(inputUser) Then
                    Return inputUser
                End If

                If String.IsNullOrEmpty(inputUser.userId) Then
                    Return inputUser
                End If

                listRoleBradepnoStaffid = GetUsersByRule(sCaseid, sStepNo).ToList
                If listRoleBradepnoStaffid.Count = 0 Then
                    Return inputUser
                End If

                For Each item As ROLE_BRADEPNO_STAFFID In listRoleBradepnoStaffid

                    If String.IsNullOrEmpty(item.StaffId) = False AndAlso
                        item.StaffId = inputUser.userId Then
                        Return inputUser
                    End If
                Next

                Return Nothing

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Function


        ''' <summary>
        ''' 刪除所有可能可被刪除的條件
        ''' </summary>
        ''' <remarks>當案件已結束，相關連的條件可以被刪除</remarks>
        Protected Sub DeleteAllPossibleRules()

            Try
                ExecuteNonQuery(
                    "delete SY_NEXTFLOWSTEPRULE " & vbCrLf & _
                    " where not exists (select distinct NFS2.CASEID " & vbCrLf & _
                    "          from SY_NEXTFLOWSTEPRULE NFS2 " & vbCrLf & _
                    "         inner join SY_FLOWINCIDENT FI " & vbCrLf & _
                    "            on NFS2.CASEID = FI.CASEID " & vbCrLf & _
                    "           and STATUS = '1' " & vbCrLf & _
                    "           and NFS2.CASEID = SY_NEXTFLOWSTEPRULE.CASEID) " & vbCrLf,
                    Nothing)

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Sub

    End Class

End Namespace
