Option Explicit On
Option Strict On

Imports com.Azion.NET.VB

Namespace TABLE

    Public Class SY_REL_FLOWMAP_SPANBRANCH
        Inherits SY_TABLEBASE

        ''' <summary>
        ''' Constructor
        ''' </summary>
        ''' <param name="dbManager">DatabaseManager</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal dbManager As com.Azion.NET.VB.DatabaseManager)
            MyBase.new("SY_REL_FLOWMAP_SPANBRANCH", dbManager)
        End Sub



        ''' <summary>
        ''' 取得流程部驟的表列單位列表
        ''' </summary>
        ''' <param name="nflowId"></param>
        ''' <param name="sStepNo"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetSpanBranch(ByVal nflowId As Integer, ByVal sStepNo As String) As BRADEPNO_BRID()

            Dim listBradepnoBrid As New List(Of BRADEPNO_BRID)
            Dim drc As DataRowCollection

            Try
                drc = GetDataRowCollection( _
                    "select distinct BR.BRA_DEPNO, BR.BRID " & vbCrLf & _
                    "  from SY_REL_FLOWMAP_SPANBRANCH FS " & vbCrLf & _
                    " inner join SY_BRANCH BR " & vbCrLf & _
                    "    on FS.SPAN_BRADEPNO = BR.BRA_DEPNO " & vbCrLf & _
                    " where FS.FLOW_ID = @FLOW_ID@ " & vbCrLf & _
                    "   and FS.STEP_NO = @STEP_NO@ " & vbCrLf,
                    "FLOW_ID", nflowId,
                    "STEP_NO", sStepNo)


                'Dbg.Assert(sStepNo <> "SPLITTER1")

                If IsNothing(drc) Then
                    Return Nothing
                End If

                For Each dr As DataRow In drc
                    listBradepnoBrid.Add(BRADEPNO_BRID.Pair(
                                         CInt(dr("BRA_DEPNO")), CStr(dr("BRID"))))
                Next
 
                Return listBradepnoBrid.ToArray

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Function


    End Class

End Namespace

