Imports MBSC.MB_OP
Imports com.Azion.NET.VB

Public Class MBMnt_ATO_MEMBER
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Private Sub btnATO_Click(sender As Object, e As EventArgs) Handles btnATO.Click
        Try
            'MB_ACCT有資料，MB_MEMBER無資料，自動補會員檔(會員編號)
            'MB_MEMSEQ = 取號(會員編號)
            Using dbManager As com.Azion.NET.VB.DatabaseManager = com.Azion.NET.VB.DatabaseManager.getInstance
                Try
                    dbManager.beginTran()

                    Dim MB_ACCTList As New MB_ACCTList(dbManager)
                    MB_ACCTList.Load_BATCH()

                    Dim MB_ACCT As New MB_ACCT(dbManager)
                    Dim MB_MEMBER As New MB_MEMBER(dbManager)
                    Dim MB_MEMBER_BATCH As New MB_MEMBER_BATCH(dbManager)
                    For Each ROW As DataRow In MB_ACCTList.getCurrentDataSet.Tables(0).Rows
                        '由MB_ACCT有資料但MB_MEMBER無資料
                        'MB_MEMSEQ = 取號(會員編號)
                        'MB_NAME= MB_ACCT.MB_NAME
                        'MB_SEX= MB_ACCT.MB_SEX
                        'MB_MOBIL= MB_ACCT.MB_MOBIL
                        'MB_TEL= MB_ACCT.MB_TEL
                        'MB_EMAIL= MB_ACCT.MB_ACCT
                        'Insert into MB_MEMBER
                        Dim sProcName As String = String.Empty
                        sProcName = LCase(com.Azion.NET.VB.Properties.getSchemaName) & ".MB_GET_MAXID"
                        Dim inParaAL As New ArrayList
                        Dim outParaAL As New ArrayList
                        inParaAL.Add("01")
                        inParaAL.Add("1")
                        outParaAL.Add(7)
                        Dim HT_Return As Hashtable = com.Azion.NET.VB.DBObject.getProcedureData(dbManager, sProcName, inParaAL, outParaAL)
                        Dim iMAXID As Decimal = HT_Return.Item("@IMAXID")
                        Dim sMB_MEMSEQ As String = String.Empty
                        sMB_MEMSEQ = com.Azion.EloanUtility.StrUtility.FillZero(iMAXID, 7)

                        MB_MEMBER.clear()
                        MB_MEMBER.loadByPK(CDec(sMB_MEMSEQ))
                        'MB_MEMSEQ	decimal(7,0)		NO	PRI	0		select,insert,update,references	會員編號
                        MB_MEMBER.setAttribute("MB_MEMSEQ", CDec(sMB_MEMSEQ))
                        'MB_NAME	varchar(50)	utf8_general_ci	NO	MUL			select,insert,update,references	姓名
                        MB_MEMBER.setAttribute("MB_NAME", ROW("MB_NAME").ToString)
                        'MB_SEX	char(1)	latin1_swedish_ci	YES				select,insert,update,references	性別 1:男 2:女
                        MB_MEMBER.setAttribute("MB_SEX", ROW("MB_SEX").ToString)
                        'MB_MOBIL	varchar(20)	utf8_general_ci	YES	MUL			select,insert,update,references	手機
                        MB_MEMBER.setAttribute("MB_MOBIL", ROW("MB_MOBIL").ToString)
                        'MB_TEL	varchar(40)	utf8_general_ci	YES	MUL			select,insert,update,references	電話
                        MB_MEMBER.setAttribute("MB_TEL", ROW("MB_TEL").ToString)
                        'MB_EMAIL	varchar(40)	utf8_general_ci	YES				select,insert,update,references	e-mail
                        MB_MEMBER.setAttribute("MB_EMAIL", ROW("MB_ACCT").ToString)
                        MB_MEMBER.setAttribute("CHGDATE", Now)
                        MB_MEMBER.setAttribute("CHGUID", "BATCH")
                        MB_MEMBER.save()

                        MB_ACCT.clear()
                        If MB_ACCT.LoadByPK(ROW("MB_ACCT").ToString) Then
                            MB_ACCT.setAttribute("MB_MEMSEQ", CDec(sMB_MEMSEQ))
                            MB_ACCT.save()
                        End If

                        MB_MEMBER_BATCH.clear()
                        MB_MEMBER_BATCH.loadByPK(CDec(sMB_MEMSEQ))
                        Dim BosAttributeList As com.Azion.NET.VB.BosAttributeList = MB_MEMBER.getBosAttributes
                        For i As Integer = 0 To BosAttributeList.size - 1
                            MB_MEMBER_BATCH.setAttribute(BosAttributeList.item(i).getColName(), MB_MEMBER.getAttribute(BosAttributeList.item(i).getColName()))
                        Next
                        MB_MEMBER_BATCH.save()
                    Next

                    Dim MB_ACCT_BATCH As New MB_ACCT_BATCH(dbManager)
                    Dim MB_MEMBERList As New MB_MEMBERList(dbManager)
                    MB_MEMBERList.Load_BATCH()
                    Dim MB_MEMMAP As New MB_MEMMAP(dbManager)
                    For Each ROW As DataRow In MB_MEMBERList.getCurrentDataSet.Tables(0).Rows
                        Dim sMB_ACCT As String = String.Empty
                        If Utility.isValidateData(ROW("MB_EMAIL")) Then
                            Dim Regex As New Regex("^[_a-z0-9-]+(.[a-z0-9-]+)@[a-z0-9-]+(.[a-z0-9-]+)*(.[a-z]{2,4})$")
                            If Regex.IsMatch(ROW("MB_EMAIL").ToString) Then
                                sMB_ACCT = ROW("MB_EMAIL").ToString
                            End If
                        End If
                        If Not Utility.isValidateData(sMB_ACCT) Then
                            MB_MEMMAP.clear()
                            If MB_MEMMAP.LoadByPK(ROW("MB_MEMSEQ")) Then
                                sMB_ACCT = MB_MEMMAP.getString("MB_BKSEQ")
                            End If
                        End If
                        If Utility.isValidateData(sMB_ACCT) Then
                            MB_ACCT.clear()
                            MB_ACCT.LoadByPK(sMB_ACCT)

                            MB_ACCT.setAttribute("MB_ACCT", sMB_ACCT)

                            Dim sMB_PSW As String = String.Empty
                            If Utility.isValidateData(ROW("MB_MOBIL")) Then
                                sMB_PSW = "MBSC" & Right(Trim(ROW("MB_MOBIL").ToString), 6)
                            ElseIf Utility.isValidateData(ROW("MB_TEL")) Then
                                sMB_PSW = "MBSC" & Right(Trim(ROW("MB_TEL").ToString), 6)
                            Else
                                sMB_PSW = sMB_ACCT
                            End If
                            MB_ACCT.setAttribute("MB_PSW", sMB_PSW)

                            MB_ACCT.setAttribute("MB_APV", "Y")

                            MB_ACCT.setAttribute("MB_NAME", ROW("MB_NAME").ToString)

                            MB_ACCT.setAttribute("MB_SEX", ROW("MB_SEX").ToString)

                            MB_ACCT.setAttribute("MB_CRE_DATE", Now)

                            MB_ACCT.setAttribute("MB_IDTIFY", "2")

                            MB_ACCT.setAttribute("MB_APVID", "BATCH")

                            MB_ACCT.setAttribute("MB_PASSVID", Nothing)

                            MB_ACCT.setAttribute("MB_MOBIL", ROW("MB_MOBIL").ToString)

                            MB_ACCT.setAttribute("MB_TEL", ROW("MB_TEL").ToString)

                            MB_ACCT.setAttribute("MB_MEMSEQ", ROW("MB_MEMSEQ"))

                            MB_ACCT.save()

                            MB_ACCT_BATCH.clear()
                            MB_ACCT_BATCH.LoadByPK(sMB_ACCT)
                            Dim BosAttributeList As com.Azion.NET.VB.BosAttributeList = MB_ACCT.getBosAttributes
                            For i As Integer = 0 To BosAttributeList.size - 1
                                MB_ACCT_BATCH.setAttribute(BosAttributeList.item(i).getColName(), MB_ACCT.getAttribute(BosAttributeList.item(i).getColName()))
                            Next
                            MB_ACCT_BATCH.save()
                        End If
                    Next

                    dbManager.commit()
                Catch ex As Exception
                    dbManager.Rollback()
                    Throw
                End Try
            End Using

            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "轉檔成功")
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me,ex)
        End Try
    End Sub
End Class