Imports MBSC.MB_OP
Imports MBSC.UICtl
Imports com.Azion.NET.VB
Imports com.Azion.EloanUtility

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
        Dim DT_MB_MEMCLASS As DataTable = Nothing
        Dim DT_F_DT_MB_MEMCLASS As DataTable = Nothing
        Dim DT_UK_MB_BATCH As DataTable = Nothing
        Dim DT_FF_DT_MB_MEMCLASS As DataTable = Nothing
        Try
            Dim MB_CLASSList As New MB_CLASSList(DBManager)
            MB_CLASSList.LoadByMB_SEQ(CDec(Me.m_sMB_SEQ))
            Dim sMB_APV As String = String.Empty
            If MB_CLASSList.getCurrentDataSet.Tables(0).Rows.Count > 0 Then
                sMB_APV = MB_CLASSList.getCurrentDataSet.Tables(0).Rows(0)("MB_APV").ToString
            End If

            Dim mbMEMCLASSList As New MB_MEMCLASSList(DBManager)
            mbMEMCLASSList.LoadQEXCELData(CDec(Me.m_sMB_SEQ))
            DT_MB_MEMCLASS = mbMEMCLASSList.getCurrentDataSet.Tables(0)
            If Not IsNothing(DT_MB_MEMCLASS) AndAlso DT_MB_MEMCLASS.Rows.Count > 0 Then
                Dim DV_MB_MEMCLASS As New DataView(DT_MB_MEMCLASS, "ISNULL(MB_FWMK,' ') NOT IN ('3','4','5') AND ISNULL(MB_RESP,' ')<>'N' ", String.Empty, DataViewRowState.CurrentRows)
                DT_F_DT_MB_MEMCLASS = DV_MB_MEMCLASS.ToTable
            Else
                DT_F_DT_MB_MEMCLASS = DT_MB_MEMCLASS
            End If

            Dim sSORT As String = String.Empty
            If Not IsNothing(DT_F_DT_MB_MEMCLASS) AndAlso DT_F_DT_MB_MEMCLASS.Rows.Count > 0 Then
                Dim DV_UK_MB_BATCH As New DataView(DT_F_DT_MB_MEMCLASS)
                DT_UK_MB_BATCH = DV_UK_MB_BATCH.ToTable(True, "MB_BATCH")
                If DT_UK_MB_BATCH.Rows.Count > 1 Then
                    sSORT = "MB_MEMSEQ, MB_BATCH"
                Else
                    sSORT = "MB_CREDATETIME"
                End If

                DT_FF_DT_MB_MEMCLASS = New DataView(DT_F_DT_MB_MEMCLASS, String.Empty, sSORT, DataViewRowState.CurrentRows).ToTable
            Else
                DT_FF_DT_MB_MEMCLASS = DT_F_DT_MB_MEMCLASS
            End If

            Dim sContent As String = String.Empty
            If sMB_APV = "1" OrElse sMB_APV = "3" Then
                sContent = Me.getMB_APV_Y(DBManager, DT_FF_DT_MB_MEMCLASS)
            Else
                sContent = Me.getMB_APV_N(DBManager, DT_FF_DT_MB_MEMCLASS)
            End If

            Dim iYYYYMMDD As Decimal = 0
            iYYYYMMDD = Now.Year * 10000 + Now.Month * 100 + Now.Day
            UIShareFun.downLoadFile(Me.Page, iYYYYMMDD.ToString & ".xls", sContent)
        Catch ex As System.Threading.ThreadAbortException
        Catch ex As Exception
            Throw
        Finally
            UIShareFun.releaseConnection(DBManager)

            If Not IsNothing(DT_MB_MEMCLASS) Then
                DT_MB_MEMCLASS.Dispose()
            End If

            If Not IsNothing(DT_F_DT_MB_MEMCLASS) Then
                DT_F_DT_MB_MEMCLASS.Dispose()
            End If
        End Try
    End Sub

    Function getMB_APV_Y(ByVal DBManager As DatabaseManager, ByVal DT_MB_MEMCLASS As DataTable) As String
        Dim objStreamReader As System.IO.StreamReader = Nothing
        Dim DT_MB_SICK As DataTable = Nothing
        Try
            If System.IO.File.Exists(Server.MapPath(com.Azion.EloanUtility.UIUtility.getRootPath & "/QRY/MBQry_TYPE_01_03_v01.xml")) Then
                objStreamReader = New System.IO.StreamReader(Server.MapPath(com.Azion.EloanUtility.UIUtility.getRootPath & "/QRY/MBQry_TYPE_01_03_v01.xml"))
            Else
                Throw New Exception("EXCEL範本檔案(MBQry_TYPE_01_03_v01)不存在")
            End If

            Dim sbCtlContent As New System.Text.StringBuilder(objStreamReader.ReadToEnd())

            Dim objStringBuilder As New System.Text.StringBuilder

            Dim AP_CODEList As New AP_CODEList(DBManager)
            AP_CODEList.getVALUEandTEXTbyUPCODE(CodeList.getAppSettings("UPCODE136"))
            DT_MB_SICK = AP_CODEList.getCurrentDataSet.Tables(0)

            For Each ROW As DataRow In DT_MB_MEMCLASS.Rows
                objStringBuilder.Append("<Row ss:AutoFitHeight=""1"" ss:Height=""50"" >")

                '姓名
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & ROW("MB_NAME").ToString & "</Data></Cell>")

                '法名
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & ROW("MB_MONKNAME").ToString & "</Data></Cell>")

                '性別
                If ROW("MB_SEX").ToString = "1" Then
                    objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & "男" & "</Data></Cell>")
                ElseIf ROW("MB_SEX").ToString = "2" Then
                    objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & "女" & "</Data></Cell>")
                Else
                    objStringBuilder.Append("<Cell><Data ss:Type=""String""></Data></Cell>")
                End If

                '出生年月日
                Dim sMB_BIRTH As String = String.Empty
                If IsDate(ROW("MB_BIRTH").ToString) Then
                    Dim D_MB_BIRTH As Date = CDate(ROW("MB_BIRTH").ToString)
                    sMB_BIRTH = D_MB_BIRTH.Year & "/" & D_MB_BIRTH.Month & "/" & D_MB_BIRTH.Day
                End If
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & sMB_BIRTH & "</Data></Cell>")

                '會員手機
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & ROW("MB_MOBIL").ToString & "</Data></Cell>")

                '會員電話
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & ROW("MB_TEL").ToString & "</Data></Cell>")

                '會員信箱
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & ROW("MB_EMAIL").ToString & "</Data></Cell>")

                '通訊住址
                Dim sMB_Addr As String = String.Empty
                sMB_Addr = ROW("MB_CITY").ToString & ROW("MB_VLG").ToString & ROW("MB_ADDR").ToString
                objStringBuilder.Append("<Cell ss:StyleID=""s65""><Data ss:Type=""String"">" & sMB_Addr & "</Data></Cell>")

                '身心狀況
                objStringBuilder.Append("<Cell ss:StyleID=""s65""><Data ss:Type=""String"">" & Me.getMB_SICK(DT_MB_SICK, ROW) & "</Data></Cell>")

                '我是初學者
                If ROW("MB_BEGIN").ToString = "Y" Then
                    objStringBuilder.Append("<Cell><Data ss:Type=""String"">是</Data></Cell>")
                ElseIf ROW("MB_BEGIN").ToString = "N" Then
                    objStringBuilder.Append("<Cell><Data ss:Type=""String"">否</Data></Cell>")
                Else
                    objStringBuilder.Append("<Cell><Data ss:Type=""String""></Data></Cell>")
                End If

                '每次禪坐時間
                If IsNumeric(ROW("MB_SITIME")) Then
                    objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & ROW("MB_SITIME").ToString & "分鐘</Data></Cell>")
                Else
                    objStringBuilder.Append("<Cell><Data ss:Type=""String""></Data></Cell>")
                End If

                '参加梯次
                '課程名稱
                'Dim sMB_CLASS_NAME As String = String.Empty
                'sMB_CLASS_NAME = ROW("MB_CLASS_NAME").ToString
                'If IsNumeric(ROW("MB_BATCH")) AndAlso CDec(ROW("MB_BATCH")) > 0 Then
                '    sMB_CLASS_NAME &= "　第" & ROW("MB_BATCH").ToString & "梯次"
                'End If
                'objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & sMB_CLASS_NAME & "</Data></Cell>")
                Dim sMB_BATCH As String = String.Empty
                'If IsNumeric(ROW("MB_BATCH")) AndAlso CDec(ROW("MB_BATCH")) > 0 Then
                '    sMB_BATCH = ROW("MB_BATCH").ToString
                'End If
                Dim MB_MEMBATCHList As New MB_MEMBATCHList(DBManager)
                MB_MEMBATCHList.setSQLCondition(" ORDER BY MB_BATCH ")
                MB_MEMBATCHList.LoadBySEQ(ROW("MB_MEMSEQ"), ROW("MB_SEQ"))
                For Each sqlRow As DataRow In MB_MEMBATCHList.getCurrentDataSet.Tables(0).Rows
                    sMB_BATCH &= sqlRow("MB_BATCH").ToString & ","
                Next
                If Utility.isValidateData(sMB_BATCH) Then
                    sMB_BATCH = Left(sMB_BATCH, sMB_BATCH.Length - 1)
                End If
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & sMB_BATCH & "</Data></Cell>")

                '錄取
                'Dim sMB_CHKFLAG As String = String.Empty
                'sMB_CHKFLAG = Me.getMB_CHKFLAG(DBManager, ROW("MB_SEQ"), ROW("MB_BATCH"), ROW("MB_MEMSEQ"), ROW("MB_FWMK"), ROW("MB_SORTNO"), ROW("MB_CHKFLAG"), ROW("MB_CDATE"))
                'objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & sMB_CHKFLAG & "</Data></Cell>")

                objStringBuilder.Append("</Row>")
            Next

            sbCtlContent.Replace("{{DATAROW}}", objStringBuilder.ToString)

            Return sbCtlContent.ToString
        Catch ex As Exception
            Throw
        Finally
            If Not IsNothing(objStreamReader) Then
                objStreamReader.Close()
            End If

            If Not IsNothing(DT_MB_SICK) Then
                DT_MB_SICK.Dispose()
            End If
        End Try
    End Function

    Function getMB_SICK(ByVal DT_MB_SICK As DataTable, ByVal ROW As DataRow) As String
        Dim sMB_SICK As New System.Text.StringBuilder
        If Utility.isValidateData(ROW("MB_SICK")) Then
            Dim sTMP() As String = Split(ROW("MB_SICK").ToString, ",")
            If Not IsNothing(sTMP) AndAlso sTMP.Length > 0 Then
                For Each sSICK As String In sTMP
                    If Utility.isValidateData(sSICK) Then
                        Dim ROW_SICK() As DataRow = Nothing
                        ROW_SICK = DT_MB_SICK.Select("VALUE='" & sSICK & "'")
                        If Not IsNothing(ROW_SICK) AndAlso ROW_SICK.Length > 0 Then
                            Dim sTEXT As String = String.Empty
                            sTEXT = ROW_SICK(0)("TEXT").ToString

                            sMB_SICK.Append(sTEXT)
                            Select Case sTEXT
                                Case "過敏"
                                    If Utility.isValidateData(ROW("MB_ALLERGY")) Then
                                        sMB_SICK.Append("(").Append(ROW("MB_ALLERGY").ToString).Append(")")
                                    End If
                                Case "開刀"
                                    If Utility.isValidateData(ROW("MB_OPERATE")) Then
                                        sMB_SICK.Append("(").Append(ROW("MB_OPERATE").ToString).Append(")")
                                    End If
                                Case "其他"
                                    If Utility.isValidateData(ROW("MB_OSICK")) Then
                                        sMB_SICK.Append("(").Append(ROW("MB_OSICK").ToString).Append(")")
                                    End If
                            End Select

                            sMB_SICK.Append("、")
                        End If
                    End If
                Next
            End If
        End If

        Dim sReturn As String = String.Empty
        sReturn = sMB_SICK.ToString
        If Utility.isValidateData(sReturn) Then
            sReturn = Left(sReturn, sReturn.Length - 1)
        End If

        Return sReturn
    End Function

    Function getMB_APV_N(ByVal DBManager As DatabaseManager, ByVal DT_MB_MEMCLASS As DataTable) As String
        Dim objStreamReader As System.IO.StreamReader = Nothing
        Dim DT_MB_RELIGION As DataTable = Nothing
        Try
            If System.IO.File.Exists(Server.MapPath(com.Azion.EloanUtility.UIUtility.getRootPath & "/QRY/MBQry_TYPE_01_03_v02.xml")) Then
                objStreamReader = New System.IO.StreamReader(Server.MapPath(com.Azion.EloanUtility.UIUtility.getRootPath & "/QRY/MBQry_TYPE_01_03_v02.xml"))
            Else
                Throw New Exception("EXCEL範本檔案(MBQry_TYPE_01_03_v02)不存在")
            End If

            Dim AP_CODEList As New AP_CODEList(DBManager)
            AP_CODEList.getVALUEandTEXTbyUPCODE(CodeList.getAppSettings("UPCODE127"))
            DT_MB_RELIGION = AP_CODEList.getCurrentDataSet.Tables(0)

            Dim sbCtlContent As New System.Text.StringBuilder(objStreamReader.ReadToEnd())

            Dim objStringBuilder As New System.Text.StringBuilder

            For Each ROW As DataRow In DT_MB_MEMCLASS.Rows
                objStringBuilder.Append("<Row ss:AutoFitHeight=""1"" ss:Height=""50"" >")

                '姓名
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & ROW("MB_NAME").ToString & "</Data></Cell>")

                '性別
                If ROW("MB_SEX").ToString = "1" Then
                    objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & "男" & "</Data></Cell>")
                ElseIf ROW("MB_SEX").ToString = "2" Then
                    objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & "女" & "</Data></Cell>")
                Else
                    objStringBuilder.Append("<Cell><Data ss:Type=""String""></Data></Cell>")
                End If

                '年齡
                Dim sMB_BIRTH As String = String.Empty
                If IsDate(ROW("MB_BIRTH").ToString) Then
                    sMB_BIRTH = Me.getAge(CDate(ROW("MB_BIRTH").ToString))
                End If
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & sMB_BIRTH & "</Data></Cell>")

                '電話
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & ROW("MB_TEL").ToString & "</Data></Cell>")

                '手機
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & ROW("MB_MOBIL").ToString & "</Data></Cell>")

                '電子郵件
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & ROW("MB_EMAIL").ToString & "</Data></Cell>")

                '參加動機或目的
                objStringBuilder.Append("<Cell ss:StyleID=""s65""><Data ss:Type=""String"">" & ROW("MB_OBJECT").ToString & "</Data></Cell>")

                '我是初學者
                If ROW("MB_BEGIN").ToString = "Y" Then
                    objStringBuilder.Append("<Cell><Data ss:Type=""String"">是</Data></Cell>")
                ElseIf ROW("MB_BEGIN").ToString = "N" Then
                    objStringBuilder.Append("<Cell><Data ss:Type=""String"">否</Data></Cell>")
                Else
                    objStringBuilder.Append("<Cell><Data ss:Type=""String""></Data></Cell>")
                End If

                '我的宗教信仰
                Dim sMB_RELIGION As String = String.Empty
                If Utility.isValidateData(ROW("MB_RELIGION")) Then
                    Dim ROW_MB_RELIGION() As DataRow = Nothing
                    ROW_MB_RELIGION = DT_MB_RELIGION.Select("VALUE='" & ROW("MB_RELIGION").ToString & "'")
                    If Not IsNothing(ROW_MB_RELIGION) AndAlso ROW_MB_RELIGION.Length > 0 Then
                        sMB_RELIGION = ROW_MB_RELIGION(0)("TEXT").ToString
                    End If
                End If
                objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & sMB_RELIGION & "</Data></Cell>")

                '每次禪坐時間
                If IsNumeric(ROW("MB_SITIME")) Then
                    objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & ROW("MB_SITIME").ToString & "分鐘</Data></Cell>")
                Else
                    objStringBuilder.Append("<Cell><Data ss:Type=""String""></Data></Cell>")
                End If

                '錄取
                'Dim sMB_CHKFLAG As String = String.Empty
                'sMB_CHKFLAG = Me.getMB_CHKFLAG(DBManager, ROW("MB_SEQ"), ROW("MB_BATCH"), ROW("MB_MEMSEQ"), ROW("MB_FWMK"), ROW("MB_SORTNO"), ROW("MB_CHKFLAG"), ROW("MB_CDATE"))
                'objStringBuilder.Append("<Cell><Data ss:Type=""String"">" & sMB_CHKFLAG & "</Data></Cell>")

                objStringBuilder.Append("</Row>")
            Next

            sbCtlContent.Replace("{{DATAROW}}", objStringBuilder.ToString)

            Return sbCtlContent.ToString
        Catch ex As Exception
            Throw
        Finally
            If Not IsNothing(objStreamReader) Then
                objStreamReader.Close()
            End If
        End Try
    End Function

    Function getAge(D_BIRTH As Date) As Integer
        Dim intAge As Integer
        Dim D_NOW As Date = Now
        intAge = DateDiff("yyyy", D_BIRTH, D_NOW)
        If D_NOW < DateSerial(Year(D_NOW), Month(D_BIRTH), Day(D_BIRTH)) Then
            intAge = intAge - 1
        End If
        Return intAge
    End Function

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

                        Dim DV_SEQ As New DataView(DT_MB_MEMCLASS, "ISNULL(MB_FWMK,' ') NOT IN ('3','4','5')", "MB_CREDATETIME", DataViewRowState.CurrentRows)
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