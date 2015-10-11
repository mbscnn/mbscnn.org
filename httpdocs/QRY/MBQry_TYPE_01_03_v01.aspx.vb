Imports MBSC.MB_OP
Imports MBSC.UICtl
Imports com.Azion.NET.VB

Public Class MBQry_TYPE_01_03_v01
    Inherits System.Web.UI.Page

    Dim m_sMB_SEQ As String = String.Empty

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.m_sMB_SEQ = "" & Request.QueryString("MB_SEQ")

            Me.Bind_TYPE_3()
        Catch ex As System.Threading.ThreadAbortException
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

    Sub Bind_TYPE_3()
        Dim DBManager As DatabaseManager = UIShareFun.getDataBaseManager
        Dim objStreamReader As System.IO.StreamReader = Nothing
        Dim DT_MB_MEMCLASS As DataTable = Nothing
        Try
            If System.IO.File.Exists(Server.MapPath(com.Azion.EloanUtility.UIUtility.getRootPath & "/QRY/MBQry_TYPE_01_03_v01.xml")) Then
                objStreamReader = New System.IO.StreamReader(Server.MapPath(com.Azion.EloanUtility.UIUtility.getRootPath & "/QRY/MBQry_TYPE_01_03_v01.xml"))
            Else
                Throw New Exception("EXCEL範本檔案(MBQry_TYPE_01_03_v01)不存在")
            End If

            Dim mbMEMCLASSList As New MB_MEMCLASSList(DBManager)
            mbMEMCLASSList.LoadQEXCELData(CDec(Me.m_sMB_SEQ))
            DT_MB_MEMCLASS = mbMEMCLASSList.getCurrentDataSet.Tables(0)

            Dim sbCtlContent As New System.Text.StringBuilder(objStreamReader.ReadToEnd())

            Dim objStringBuilder As New System.Text.StringBuilder

            For Each ROW As DataRow In DT_MB_MEMCLASS.Rows
                objStringBuilder.Append("<Row ss:AutoFitHeight=""1"" ss:Height=""50"" >")

                '報課會員
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & ROW("MB_NAME").ToString & "</Data></Cell>")

                '性別
                If ROW("MB_SEX").ToString = "1" Then
                    objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & "男" & "</Data></Cell>")
                ElseIf ROW("MB_SEX").ToString = "2" Then
                    objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & "女" & "</Data></Cell>")
                Else
                    objStringBuilder.Append("<Cell><Data ss:Type=""String""></Data></Cell>")
                End If

                '會員信箱
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & ROW("MB_EMAIL").ToString & "</Data></Cell>")

                '會員電話
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & ROW("MB_TEL").ToString & "</Data></Cell>")

                '會員手機
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & ROW("MB_MOBIL").ToString & "</Data></Cell>")

                '課程編號
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & ROW("MB_SEQ").ToString & "</Data></Cell>")

                '課程名稱
                Dim sMB_CLASS_NAME As String = String.Empty
                sMB_CLASS_NAME = ROW("MB_CLASS_NAME").ToString
                If IsNumeric(ROW("MB_BATCH")) AndAlso CDec(ROW("MB_BATCH")) > 0 Then
                    sMB_CLASS_NAME &= "　第" & ROW("MB_BATCH").ToString & "梯次"
                End If
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & sMB_CLASS_NAME & "</Data></Cell>")

                '錄取
                'If ROW("MB_CHKFLAG").ToString = "1" Then
                '    objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & "正取" & "</Data></Cell>")
                'ElseIf ROW("MB_CHKFLAG").ToString = "2" Then
                '    objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & "備取" & "</Data></Cell>")
                'Else
                '    objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & "" & "</Data></Cell>")
                'End If
                Dim sMB_CHKFLAG As String = String.Empty
                sMB_CHKFLAG = Me.getMB_CHKFLAG(DBManager, ROW("MB_SEQ").ToString, ROW("MB_BATCH"), ROW("MB_MEMSEQ").ToString, ROW("MB_FWMK").ToString, ROW("MB_SORTNO").ToString, ROW("MB_CHKFLAG").ToString, ROW("MB_CDATE"))
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & sMB_CHKFLAG & "</Data></Cell>")

                '回信出席
                If ROW("MB_RESP").ToString = "Y" Then
                    objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & "是" & "</Data></Cell>")
                ElseIf ROW("MB_RESP").ToString = "N" Then
                    objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & "否" & "</Data></Cell>")
                Else
                    objStringBuilder.Append("<Cell><Data ss:Type=""String""></Data></Cell>")
                End If

                '未出席
                If ROW("MB_FWMK").ToString = "1" Then
                    If ROW("MB_RESP").ToString = "N" Then
                        objStringBuilder.Append("<Cell><Data ss:Type=""String""></Data></Cell>")
                    Else
                        objStringBuilder.Append("<Cell><Data ss:Type=""String""></Data></Cell>")
                    End If
                Else
                    If ROW("MB_FWMK").ToString = "2" Then
                        objStringBuilder.Append("<Cell><Data ss:Type=""String"">是</Data></Cell>")
                    Else
                        objStringBuilder.Append("<Cell><Data ss:Type=""String""></Data></Cell>")
                    End If
                End If

                objStringBuilder.Append("</Row>")
            Next

            sbCtlContent.Replace("{{DATAROW}}", objStringBuilder.ToString)

            Dim iYYYYMMDD As Decimal = 0
            iYYYYMMDD = Now.Year * 10000 + Now.Month * 100 + Now.Day
            UIShareFun.downLoadFile(Me.Page, iYYYYMMDD.ToString & ".xls", sbCtlContent.ToString)
        Catch ex As System.Threading.ThreadAbortException
        Catch ex As Exception
            Throw
        Finally
            UIShareFun.releaseConnection(DBManager)

            If Not IsNothing(objStreamReader) Then
                objStreamReader.Close()
            End If

            If Not IsNothing(DT_MB_MEMCLASS) Then
                DT_MB_MEMCLASS.Dispose()
            End If
        End Try
    End Sub

    Public Function getMB_CHKFLAG(ByVal dbManager As DatabaseManager, ByVal iMB_SEQ As Object, ByVal iMB_BATCH As Object, ByVal iMB_MEMSEQ As Object, ByVal sMB_FWMK As Object, ByVal iMB_SORTNO As Object, sMB_CHKFLAG As Object, ByVal MB_CDATE As Object) As String
        Dim DT_MB_MEMCLASS As DataTable = Nothing
        Dim DT_SEQ As DataTable = Nothing
        Dim DT_WAIT As DataTable = Nothing
        Dim DT_FULL As DataTable = Nothing
        Try
            'If Utility.isValidateData(sMB_CHKFLAG) Then
            '    If sMB_CHKFLAG = "1" Then
            '        Return "正取"
            '    ElseIf sMB_CHKFLAG = "2" Then
            '        Return "備取"
            '    End If
            'End If

            If Utility.isValidateData(sMB_FWMK) AndAlso sMB_FWMK = "3" Then
                Return "取消報名" & MB_CDATE.ToString
            ElseIf Utility.isValidateData(sMB_FWMK) AndAlso sMB_FWMK = "4" Then
                Dim sSECCancel As String = String.Empty
                sSECCancel = "第二次取消報名"
                If IsDate(MB_CDATE) Then
                    sSECCancel &= CDate(MB_CDATE).Year & "/" & CDate(MB_CDATE).Month & "/" & CDate(MB_CDATE).Day
                End If
                Return sSECCancel
            Else
                If IsNumeric(iMB_SEQ) AndAlso IsNumeric(iMB_BATCH) AndAlso IsNumeric(iMB_MEMSEQ) Then
                    Dim mbCLASS As New MB_CLASS(dbManager)
                    If mbCLASS.loadByPK(iMB_SEQ, iMB_BATCH) Then
                        Dim iMB_FULL As Decimal = 0
                        iMB_FULL = mbCLASS.getDecimal("MB_FULL")
                        Dim iMB_WAIT As Decimal = 0
                        iMB_WAIT = mbCLASS.getDecimal("MB_WAIT")

                        Dim mbMEMCLASSList As New MB_MEMCLASSList(dbManager)
                        mbMEMCLASSList.loadByMB_SEQ(iMB_SEQ, iMB_BATCH)
                        DT_MB_MEMCLASS = mbMEMCLASSList.getCurrentDataSet.Tables(0)

                        Dim DV_SEQ As New DataView(DT_MB_MEMCLASS, "ISNULL(MB_FWMK,' ') NOT IN ('3','4')", "MB_CREDATETIME", DataViewRowState.CurrentRows)
                        DT_SEQ = DV_SEQ.ToTable
                        DT_FULL = DT_SEQ.Clone
                        Dim iFULL As Decimal = 0
                        For i As Integer = 0 To DT_SEQ.Rows.Count - 1
                            If iFULL <= iMB_FULL Then
                                DT_FULL.ImportRow(DT_SEQ.Rows(i))
                            End If

                            iFULL += 1
                        Next
                        Dim ROW_FULL() As DataRow = Nothing
                        ROW_FULL = DT_FULL.Select("MB_MEMSEQ=" & iMB_MEMSEQ)
                        If Not IsNothing(ROW_FULL) AndAlso ROW_FULL.Length > 0 Then
                            Return "正取"
                        Else
                            Return "備取"
                        End If
                    End If
                End If
            End If

            Return String.Empty
        Catch ex As Exception
            Throw
        Finally
            If Not IsNothing(DT_MB_MEMCLASS) Then
                DT_MB_MEMCLASS.Dispose()
            End If

            If Not IsNothing(DT_SEQ) Then
                DT_SEQ.Dispose()
            End If

            If Not IsNothing(DT_WAIT) Then
                DT_WAIT.Dispose()
            End If

            If Not IsNothing(DT_FULL) Then
                DT_FULL.Dispose()
            End If
        End Try
    End Function
End Class