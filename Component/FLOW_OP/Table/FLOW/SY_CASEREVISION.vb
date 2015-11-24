Option Explicit On
Option Strict On

Imports com.Azion.NET.VB
Imports System.Data.SqlClient




Namespace TABLE

    Public Class SY_CASEREVISION
        Inherits SY_TABLEBASE

        ''' <summary>
        ''' Constructor
        ''' </summary>
        ''' <param name="dbManager">DatabaseManager</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal dbManager As com.Azion.NET.VB.DatabaseManager)
            MyBase.new("SY_CASEREVISION", dbManager)
        End Sub


        ''' <summary>
        ''' 新增修正補充理由
        ''' </summary>
        ''' <param name="sCaseid">案件</param>
        ''' <param name="sSrcStepNo">同案件修正補充的序號</param>
        ''' <param name="sNote">修正補充內容</param>
        ''' <param name="sRefundStepNo">退回的步驟代碼</param>
        ''' <param name="sChgUid">目前的USER</param>
        ''' <param name="nBraDepNo">目前的BRA_DEPNO</param>
        ''' <param name="sRefundUid">退回給哪一個USER</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function InsertRevision(ByVal sCaseid As String,
                                       ByVal sSrcStepNo As String,
                                       ByVal sNote As String,
                                       ByVal sRefundStepNo As String,
                                       ByVal sChgUid As String,
                                       ByVal nBraDepNo As Integer,
                                       ByVal sRefundUid As String) As DataRow

            Dim sSql As String
            Dim nValue As Integer
            Dim dr As DataRow

            Try
                sSql = _
                "insert into SY_CASEREVISION " & vbCrLf & _
                "           (CASEID " & vbCrLf & _
                "           ,SEQNO " & vbCrLf & _
                "           ,SRCSTEPNO " & vbCrLf & _
                "           ,VERNO " & vbCrLf & _
                "           ,NOTE " & vbCrLf & _
                "           ,REFUNDSTEPNO " & vbCrLf & _
                "           ,CHGUID " & vbCrLf & _
                "           ,CHGDATE " & vbCrLf & _
                "           ,BRA_DEPNO " & vbCrLf & _
                "           ,REFUNDUID) " & vbCrLf & _
                "     values " & vbCrLf & _
                "           (@CASEID@ " & vbCrLf & _
                "           ,(select ISNULL(MAX(SEQNO)+1,1) from SY_CASEREVISION where CASEID=@CASEID@) " & vbCrLf & _
                "           ,@SRCSTEPNO@ " & vbCrLf & _
                "           ,(select ISNULL(MAX(VERNO)+1,1) from SY_CASEREVISION where CASEID=@CASEID@ and SRCSTEPNO=@SRCSTEPNO@) " & vbCrLf & _
                "           ,@NOTE@ " & vbCrLf & _
                "           ,@REFUNDSTEPNO@ " & vbCrLf & _
                "           ,@CHGUID@ " & vbCrLf & _
                "           ,GETDATE() " & vbCrLf & _
                "           ,@BRA_DEPNO@ " & vbCrLf & _
                "           ,@REFUNDUID@) "

                nValue = ExecuteNonQuery(sSql,
                                          "CASEID", sCaseid,
                                          "SRCSTEPNO", sSrcStepNo,
                                          "NOTE", sNote,
                                          "REFUNDSTEPNO", sRefundStepNo,
                                          "CHGUID", sChgUid,
                                          "BRA_DEPNO", nBraDepNo,
                                          "REFUNDUID", sRefundUid)

                If nValue = 0 Then
                    Throw New SYException(
                        String.Format("無法寫入至SY_CASEREVISION，CASEID={0}, SRCSTEPNO={1} ", sCaseid, sSrcStepNo),
                        SYMSG.SYCASEREVISION_CANNOT_INSERTREVISION)
                End If

                dr = GetDataRow( _
                    "select * " & vbCrLf & _
                    "  from SY_CASEREVISION " & vbCrLf & _
                    " where CASEID = @CASEID@ " & vbCrLf & _
                    "   and VERNO = (select MAX(VERNO) " & vbCrLf & _
                    "                  from SY_CASEREVISION " & vbCrLf & _
                    "                 WHERE CASEID = @CASEID@) ",
                    "CASEID", sCaseid)

                Return dr

            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Function



    End Class


End Namespace
