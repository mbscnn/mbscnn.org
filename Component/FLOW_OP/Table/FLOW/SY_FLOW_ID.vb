Option Explicit On
Option Strict On

Imports com.Azion.NET.VB

Namespace TABLE

    Public Class SY_FLOW_ID
        Inherits SY_TABLEBASE


        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="dbManager">DatabaseManager</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal dbManager As com.Azion.NET.VB.DatabaseManager)
            MyBase.new("SY_FLOW_ID", dbManager)
        End Sub


        ''' <summary>
        ''' 傳入Flowname，取得FLOW的所有資料
        ''' </summary>
        ''' <param name="sFlowName"></param>
        ''' <returns></returns>
        ''' <remarks>FLOWNAME為UNIQUE屬性，因此可從FLOW_NAME取得FLOW_ID, SUBSYSID, SYSID</remarks>
        Public Function GetInfo(ByVal sFlowName As String) As DataRow
            Try
                Return GetDataRow(PARAMETER("FLOW_NAME", sFlowName))
            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try

        End Function



        ''' <summary>
        ''' 取得SY_FLOW_ID的SUBSYSID
        ''' </summary>
        ''' <param name="sFlowName"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetSubSysId(ByVal sFlowName As String) As String
            Dim dr As DataRow

            Try
                dr = GetInfo(sFlowName)

                If IsNothing(dr) Then
                    Throw New SYException(
                        String.Format("無法從取得SY_FLOW_ID的內容，FLOWNAME={0}", sFlowName),
                        SYMSG.SYFLOWID_CANNOT_GET_SUBSYSID, GetLastSQL)
                End If

                Return CDbType(Of String)(dr("SUBSYSID"), Nothing)
            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try
        End Function

        ''' <summary>
        ''' 取得SY_FLOW_ID的FLOWCNAME
        ''' </summary>
        ''' <param name="nFlowId"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetFlow_CName(ByVal nFlowId As Integer) As String
            Dim dr As DataRow

            Try
                dr = GetDataRow(Nothing, "FLOW_ID", nFlowId)

                If IsNothing(dr) Then
                    Throw New SYException(
                        String.Format("無法從取得SY_FLOW_ID的內容，FlowId={0}", nFlowId),
                        SYMSG.SYFLOWID_CANNOT_GET_FLOWCNAME, GetLastSQL)
                End If

                Return CDbType(Of String)(dr("FLOW_CNAME"), "")
            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try
        End Function




        ''' <summary>
        ''' 由sCaseid取得Flowname
        ''' </summary>
        ''' <param name="sCaseid"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetFlownameByCaseid(ByVal sCaseid As String) As String
            Try
                Return CDbType(Of String)(ExecuteScalar( _
                    "select SY_FLOW_ID.FLOW_NAME " & vbCrLf & _
                    "  from SY_CASEID " & vbCrLf & _
                    " inner join SY_FLOW_ID " & vbCrLf & _
                    "    on SY_CASEID.FLOW_ID = SY_FLOW_ID.FLOW_ID " & vbCrLf & _
                    " where CASEID = @CASEID@ " & vbCrLf,
                    "CASEID", sCaseid), "")
            Catch ex As SYException
                Throw
            Catch ex As Exception
                Throw New SYException(ex.Message, ex, SYMSG.UNDEFINE, GetGlobalLastSQL)
            End Try
        End Function



    End Class

End Namespace
