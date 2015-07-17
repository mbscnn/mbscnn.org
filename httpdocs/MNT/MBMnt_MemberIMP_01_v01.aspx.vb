Imports NPOI.HSSF.UserModel
Imports NPOI.SS.UserModel
Imports MBSC.MB_OP
Imports com.Azion.NET.VB
Imports MBSC.UICtl

Public Class MBMnt_MemberIMP_01_v01
    Inherits System.Web.UI.Page

    Dim m_DT_ALLCITY As DataTable

    Dim m_PageSize As Decimal = 15

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Sub btIMPORT_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles btIMPORT.Click
        Dim dbManager As com.Azion.NET.VB.DatabaseManager = UIShareFun.getDataBaseManager
        Try
            Dim objHSSFWorkbook As New HSSFWorkbook(Me.FUEXCEL.PostedFile.InputStream)

            Dim sheet As ISheet = objHSSFWorkbook.GetSheetAt(0)

            Dim rows As System.Collections.IEnumerator = sheet.GetRowEnumerator()

            dbManager.beginTran()

            Dim mbMEMBER As New MB_MEMBER(dbManager)

            Dim mbMBIMP As New MB_MBIMP(dbManager)

            Dim mbMEMBERList As New MB_MEMBERList(dbManager)

            'Dim apROADSECList As New AP_ROADSECList(dbManager)
            'apROADSECList.loadCityNOT_D()
            'Me.m_DT_ALLCITY = apROADSECList.getCurrentDataSet.Tables(0)

            While rows.MoveNext()
                Dim DT_MB_MEMBER As DataTable = Nothing
                Try
                    Dim row As IRow = DirectCast(rows.Current, HSSFRow)

                    '姓名
                    Dim sMB_NAME As String = String.Empty
                    If Utility.isValidateData(row.GetCell(0)) Then
                        sMB_NAME = Trim(row.GetCell(0).ToString())
                    End If

                    'MB_MONKNAME
                    If Utility.isValidateData(row.GetCell(19)) Then
                        sMB_NAME = Trim(row.GetCell(19).ToString)
                    End If

                    'MB_EMAIL
                    If Not Utility.isValidateData(sMB_NAME) Then
                        If Utility.isValidateData(row.GetCell(6)) Then
                            sMB_NAME = Trim(row.GetCell(6).ToString)
                        End If
                    End If

                    'MB_MOBIL
                    If Not Utility.isValidateData(sMB_NAME) Then
                        If Utility.isValidateData(row.GetCell(4)) Then
                            sMB_NAME = Trim(row.GetCell(4).ToString)
                        End If
                    End If

                    'MB_TEL
                    If Not Utility.isValidateData(sMB_NAME) Then
                        If Utility.isValidateData(row.GetCell(5)) Then
                            sMB_NAME = Trim(row.GetCell(5).ToString)
                        End If
                    End If

                    'MB_ID
                    If Not Utility.isValidateData(sMB_NAME) Then
                        If Utility.isValidateData(row.GetCell(7)) Then
                            sMB_NAME = Trim(row.GetCell(7).ToString)
                        End If
                    End If

                    If Utility.isValidateData(sMB_NAME) Then
                        '手機
                        Dim sMB_MOBIL As String = String.Empty
                        If Utility.isValidateData(row.GetCell(4)) Then
                            sMB_MOBIL = Trim(row.GetCell(4).ToString)
                        End If

                        '電話
                        Dim sMB_TEL As String = String.Empty
                        If Utility.isValidateData(row.GetCell(5)) Then
                            sMB_TEL = Trim(row.GetCell(5).ToString)
                            If Left(sMB_TEL, 1) = "（" AndAlso sMB_TEL.Substring(3, 1) = "）" Then
                                sMB_TEL = Replace(sMB_TEL, "（", "")
                                sMB_TEL = Replace(sMB_TEL, "）", "-")
                            End If
                        End If

                        '身分證字號
                        Dim sMB_ID As String = String.Empty
                        If Utility.isValidateData(row.GetCell(7)) Then
                            sMB_ID = Trim(row.GetCell(7).ToString)
                        End If

                        mbMEMBERList.clear()
                        mbMEMBERList.loadByMB_NAME(sMB_NAME)
                        DT_MB_MEMBER = mbMEMBERList.getCurrentDataSet.Tables(0)

                        mbMEMBER.clear()

                        '取得會員編號
                        Dim iMB_MEMSEQ As Decimal = 0

                        If DT_MB_MEMBER.Rows.Count = 0 Then
                            iMB_MEMSEQ = Me.getMB_MEMSEQ(dbManager)
                        Else
                            Dim isHaveGetMB_MEMSEQ As Boolean = False
                            If Utility.isValidateData(sMB_ID) Then
                                Dim ROW_MB_ID() As DataRow = DT_MB_MEMBER.Select("MB_ID='" & sMB_ID & "'")
                                If Not IsNothing(ROW_MB_ID) AndAlso ROW_MB_ID.Length > 0 Then
                                    iMB_MEMSEQ = ROW_MB_ID(0)("MB_MEMSEQ")

                                    isHaveGetMB_MEMSEQ = True
                                End If
                            End If

                            If Not isHaveGetMB_MEMSEQ Then
                                If Utility.isValidateData(sMB_MOBIL) Then
                                    Dim ROW_MB_MOBIL() As DataRow = DT_MB_MEMBER.Select("MB_MOBIL='" & sMB_MOBIL & "'")
                                    If Not IsNothing(ROW_MB_MOBIL) AndAlso ROW_MB_MOBIL.Length > 0 Then
                                        iMB_MEMSEQ = ROW_MB_MOBIL(0)("MB_MEMSEQ")

                                        isHaveGetMB_MEMSEQ = True
                                    End If
                                End If
                            End If

                            If Not isHaveGetMB_MEMSEQ Then
                                If Utility.isValidateData(sMB_TEL) Then
                                    Dim ROW_MB_TEL() As DataRow = DT_MB_MEMBER.Select("MB_TEL='" & sMB_TEL & "'")
                                    If Not IsNothing(ROW_MB_TEL) AndAlso ROW_MB_TEL.Length > 0 Then
                                        iMB_MEMSEQ = ROW_MB_TEL(0)("MB_MEMSEQ")

                                        isHaveGetMB_MEMSEQ = True
                                    End If
                                End If
                            End If

                            If Not isHaveGetMB_MEMSEQ Then
                                If DT_MB_MEMBER.Rows.Count = 1 Then
                                    iMB_MEMSEQ = DT_MB_MEMBER.Rows(0)("MB_MEMSEQ")
                                Else
                                    iMB_MEMSEQ = Me.getMB_MEMSEQ(dbManager)
                                End If
                            End If
                        End If

                        mbMEMBER.loadByPK(iMB_MEMSEQ)

                        '會員編號
                        mbMEMBER.setAttribute("MB_MEMSEQ", iMB_MEMSEQ)

                        '姓名
                        mbMEMBER.setAttribute("MB_NAME", sMB_NAME)

                        '性別 1:男 2:女
                        'If Utility.isValidateData(Trim(row.GetCell(3))) Then
                        '    If UCase(com.Azion.EloanUtility.ValidateUtility.ChPersonalID1(Trim(row.GetCell(3)))) = "TRUE" Then
                        '        If Trim(row.GetCell(3)).Substring(1, 1) = "1" Then
                        '            mbMEMBER.setAttribute("MB_SEX", "1")
                        '        Else
                        '            mbMEMBER.setAttribute("MB_SEX", "2")
                        '        End If
                        '    End If
                        'End If
                        If Utility.isValidateData(row.GetCell(1)) Then
                            mbMEMBER.setAttribute("MB_SEX", Trim(row.GetCell(1).ToString))
                        Else
                            mbMEMBER.setAttribute("MB_SEX", Nothing)
                        End If


                        '出生日期
                        Dim sMB_BIRTH As String = String.Empty
                        If Utility.isValidateData(row.GetCell(2)) Then
                            sMB_BIRTH = Trim(row.GetCell(2).ToString)
                        End If
                        Dim D_MB_BIRTH As Object = Convert.DBNull
                        If Utility.isValidateData(sMB_BIRTH) Then
                            Dim sTEMP() As String = Nothing
                            sTEMP = Split(sMB_BIRTH, "/")
                            If Not IsNothing(sTEMP) AndAlso sTEMP.Length >= 3 Then
                                Dim iYYY As Integer = 0
                                If IsNumeric(sTEMP(0)) Then
                                    iYYY = CInt(sTEMP(0))
                                End If
                                Dim iMM As Integer = 0
                                If IsNumeric(sTEMP(1)) Then
                                    iMM = CInt(sTEMP(1))
                                End If
                                Dim iDD As Integer = 0
                                If IsNumeric(sTEMP(2)) Then
                                    iDD = CInt(sTEMP(2))
                                End If

                                Try
                                    D_MB_BIRTH = New Date(iYYY, iMM, iDD)
                                Catch ex As Exception
                                    D_MB_BIRTH = Convert.DBNull
                                End Try
                            End If
                        End If
                        If IsDate(D_MB_BIRTH) AndAlso CDate(D_MB_BIRTH).Year > 1911 Then
                            mbMEMBER.setAttribute("MB_BIRTH", D_MB_BIRTH)
                        Else
                            mbMEMBER.setAttribute("MB_BIRTH", Nothing)
                        End If

                        '身分別 勾選多種以，隔開 A:會員B:委員C:學員D:法工
                        If Utility.isValidateData(row.GetCell(3)) Then
                            mbMEMBER.setAttribute("MB_IDENTIFY", Trim(row.GetCell(3).ToString))
                        Else
                            mbMEMBER.setAttribute("MB_IDENTIFY", Nothing)
                        End If

                        '手機
                        If Utility.isValidateData(row.GetCell(4)) Then
                            mbMEMBER.setAttribute("MB_MOBIL", Trim(row.GetCell(4).ToString))
                        Else
                            mbMEMBER.setAttribute("MB_MOBIL", Nothing)
                        End If

                        '市話
                        If Utility.isValidateData(row.GetCell(5)) Then
                            mbMEMBER.setAttribute("MB_TEL", Trim(row.GetCell(5).ToString))
                        Else
                            mbMEMBER.setAttribute("MB_TEL", Nothing)
                        End If

                        'e-mail
                        If Utility.isValidateData(row.GetCell(6)) Then
                            mbMEMBER.setAttribute("MB_EMAIL", Trim(row.GetCell(6).ToString))
                        Else
                            mbMEMBER.setAttribute("MB_EMAIL", Nothing)
                        End If


                        '身分證字號
                        If Utility.isValidateData(row.GetCell(7)) Then
                            mbMEMBER.setAttribute("MB_ID", Trim(row.GetCell(7).ToString))
                        Else
                            mbMEMBER.setAttribute("MB_ID", Nothing)
                        End If

                        'MB_EDU
                        If Utility.isValidateData(row.GetCell(8)) Then
                            mbMEMBER.setAttribute("MB_EDU", Trim(row.GetCell(8).ToString))
                        Else
                            mbMEMBER.setAttribute("MB_EDU", Nothing)
                        End If

                        'MB_REFER
                        If Utility.isValidateData(row.GetCell(9)) Then
                            mbMEMBER.setAttribute("MB_REFER", Trim(row.GetCell(9).ToString))
                        Else
                            mbMEMBER.setAttribute("MB_REFER", Nothing)
                        End If

                        'MB_CITY
                        If Utility.isValidateData(row.GetCell(10)) Then
                            mbMEMBER.setAttribute("MB_CITY", Trim(row.GetCell(10).ToString))
                        Else
                            mbMEMBER.setAttribute("MB_CITY", Nothing)
                        End If

                        'MB_VLG
                        If Utility.isValidateData(row.GetCell(11)) Then
                            mbMEMBER.setAttribute("MB_VLG", Trim(row.GetCell(11).ToString))
                        Else
                            mbMEMBER.setAttribute("MB_VLG", Nothing)
                        End If

                        'MB_ADDR
                        If Utility.isValidateData(row.GetCell(12)) Then
                            mbMEMBER.setAttribute("MB_ADDR", Trim(row.GetCell(12).ToString))
                        Else
                            mbMEMBER.setAttribute("MB_ADDR", Nothing)
                        End If

                        'MB_CITY1
                        If Utility.isValidateData(row.GetCell(13)) Then
                            mbMEMBER.setAttribute("MB_CITY1", Trim(row.GetCell(13).ToString))
                        Else
                            mbMEMBER.setAttribute("MB_CITY1", Nothing)
                        End If

                        'MB_VLG1
                        If Utility.isValidateData(row.GetCell(14)) Then
                            mbMEMBER.setAttribute("MB_VLG1", Trim(row.GetCell(14).ToString))
                        Else
                            mbMEMBER.setAttribute("MB_VLG1", Nothing)
                        End If

                        'MB_ADDR1
                        If Utility.isValidateData(row.GetCell(15)) Then
                            mbMEMBER.setAttribute("MB_ADDR1", Trim(row.GetCell(15).ToString))
                        Else
                            mbMEMBER.setAttribute("MB_ADDR1", Nothing)
                        End If

                        'MB_AREA
                        If Utility.isValidateData(row.GetCell(16)) Then
                            mbMEMBER.setAttribute("MB_AREA", Trim(row.GetCell(16).ToString))
                        Else
                            mbMEMBER.setAttribute("MB_AREA", "A")
                        End If

                        'MB_LEADER
                        If Utility.isValidateData(row.GetCell(17)) Then
                            mbMEMBER.setAttribute("MB_LEADER", Trim(row.GetCell(17).ToString))
                        Else
                            mbMEMBER.setAttribute("MB_LEADER", Nothing)
                        End If

                        'MB_MONK
                        If Utility.isValidateData(row.GetCell(18)) Then
                            mbMEMBER.setAttribute("MB_MONK", Trim(row.GetCell(18).ToString))
                        Else
                            mbMEMBER.setAttribute("MB_MONK", Nothing)
                        End If

                        'MB_MONKNAME
                        If Utility.isValidateData(row.GetCell(19)) Then
                            mbMEMBER.setAttribute("MB_MONKNAME", Trim(row.GetCell(19).ToString))
                        Else
                            mbMEMBER.setAttribute("MB_MONKNAME", Nothing)
                        End If

                        'MB_MONKTECH
                        If Utility.isValidateData(row.GetCell(20)) Then
                            mbMEMBER.setAttribute("MB_MONKTECH", Trim(row.GetCell(20).ToString))
                        Else
                            mbMEMBER.setAttribute("MB_MONKTECH", Nothing)
                        End If

                        'MB_EDUTYPE
                        If Utility.isValidateData(row.GetCell(21)) Then
                            mbMEMBER.setAttribute("MB_EDUTYPE", Trim(row.GetCell(21).ToString))
                        Else
                            mbMEMBER.setAttribute("MB_EDUTYPE", Nothing)
                        End If

                        'MB_MONKTYPE
                        If Utility.isValidateData(row.GetCell(22)) Then
                            mbMEMBER.setAttribute("MB_MONKTYPE", Trim(row.GetCell(22).ToString))
                        Else
                            mbMEMBER.setAttribute("MB_MONKTYPE", Nothing)
                        End If

                        'MB_MONKPLACE
                        If Utility.isValidateData(row.GetCell(23)) Then
                            mbMEMBER.setAttribute("MB_MONKPLACE", Trim(row.GetCell(23).ToString))
                        Else
                            mbMEMBER.setAttribute("MB_MONKPLACE", Nothing)
                        End If

                        'MB_MONKDATE
                        Dim sMB_MONKDATE As String = String.Empty
                        If Utility.isValidateData(row.GetCell(24)) Then
                            sMB_MONKDATE = Trim(row.GetCell(24).ToString)
                        End If

                        If IsDate(sMB_MONKDATE) Then
                            mbMEMBER.setAttribute("MB_MONKDATE", CDate(sMB_MONKDATE))
                        Else
                            mbMEMBER.setAttribute("MB_MONKDATE", Nothing)
                        End If

                        'MB_MONKPLACE1
                        If Utility.isValidateData(row.GetCell(25)) Then
                            mbMEMBER.setAttribute("MB_MONKPLACE", Trim(row.GetCell(25).ToString))
                        Else
                            mbMEMBER.setAttribute("MB_MONKPLACE", Nothing)
                        End If

                        'MB_LANG
                        If Utility.isValidateData(row.GetCell(26)) Then
                            mbMEMBER.setAttribute("MB_LANG", Trim(row.GetCell(26).ToString))
                        Else
                            mbMEMBER.setAttribute("MB_LANG", Nothing)
                        End If

                        'MB_OLANG
                        If Utility.isValidateData(row.GetCell(27)) Then
                            mbMEMBER.setAttribute("MB_OLANG", Trim(row.GetCell(27).ToString))
                        Else
                            mbMEMBER.setAttribute("MB_OLANG", Nothing)
                        End If

                        'MB_SPECIAL
                        If Utility.isValidateData(row.GetCell(28)) Then
                            mbMEMBER.setAttribute("MB_SPECIAL", Trim(row.GetCell(28).ToString))
                        Else
                            mbMEMBER.setAttribute("MB_SPECIAL", Nothing)
                        End If

                        'MB_PROFESSION
                        If Utility.isValidateData(row.GetCell(29)) Then
                            mbMEMBER.setAttribute("MB_PROFESSION", Trim(row.GetCell(29).ToString))
                        Else
                            mbMEMBER.setAttribute("MB_PROFESSION", Nothing)
                        End If

                        'MB_SICK
                        If Utility.isValidateData(row.GetCell(30)) Then
                            mbMEMBER.setAttribute("MB_SICK", Trim(row.GetCell(30).ToString))
                        Else
                            mbMEMBER.setAttribute("MB_SICK", Nothing)
                        End If

                        'MB_ALLERGY
                        If Utility.isValidateData(row.GetCell(31)) Then
                            mbMEMBER.setAttribute("MB_ALLERGY", Trim(row.GetCell(31).ToString))
                        Else
                            mbMEMBER.setAttribute("MB_ALLERGY", Nothing)
                        End If

                        'MB_OPERATE
                        If Utility.isValidateData(row.GetCell(32)) Then
                            mbMEMBER.setAttribute("MB_OPERATE", Trim(row.GetCell(32).ToString))
                        Else
                            mbMEMBER.setAttribute("MB_OPERATE", Nothing)
                        End If

                        'MB_OSICK
                        If Utility.isValidateData(row.GetCell(33)) Then
                            mbMEMBER.setAttribute("MB_OSICK", Trim(row.GetCell(33).ToString))
                        Else
                            mbMEMBER.setAttribute("MB_OSICK", Nothing)
                        End If

                        'MB_PIPOSHENA
                        If Utility.isValidateData(row.GetCell(34)) Then
                            mbMEMBER.setAttribute("MB_PIPOSHENA", Trim(row.GetCell(34).ToString))
                        Else
                            mbMEMBER.setAttribute("MB_PIPOSHENA", Nothing)
                        End If

                        'MB_TEACH
                        If Utility.isValidateData(row.GetCell(35)) Then
                            mbMEMBER.setAttribute("MB_TEACH", Trim(row.GetCell(35).ToString))
                        Else
                            mbMEMBER.setAttribute("MB_TEACH", Nothing)
                        End If

                        'MB_FAMENNIAN
                        If Utility.isValidateData(row.GetCell(36)) Then
                            mbMEMBER.setAttribute("MB_FAMENNIAN", Trim(row.GetCell(36).ToString))
                        Else
                            mbMEMBER.setAttribute("MB_FAMENNIAN", Nothing)
                        End If

                        'MB_OVER7DAY
                        If Utility.isValidateData(row.GetCell(37)) Then
                            mbMEMBER.setAttribute("MB_OVER7DAY", Trim(row.GetCell(37).ToString))
                        Else
                            mbMEMBER.setAttribute("MB_OVER7DAY", Nothing)
                        End If

                        'MB_PLACE
                        If Utility.isValidateData(row.GetCell(38)) Then
                            mbMEMBER.setAttribute("MB_PLACE", Trim(row.GetCell(38).ToString))
                        Else
                            mbMEMBER.setAttribute("MB_PLACE", Nothing)
                        End If

                        'MB_EMGCONT
                        If Utility.isValidateData(row.GetCell(39)) Then
                            mbMEMBER.setAttribute("MB_EMGCONT", Trim(row.GetCell(39).ToString))
                        Else
                            mbMEMBER.setAttribute("MB_EMGCONT", Nothing)
                        End If

                        'MB_CONTID
                        If Utility.isValidateData(row.GetCell(40)) Then
                            mbMEMBER.setAttribute("MB_CONTID", Trim(row.GetCell(40).ToString))
                        Else
                            mbMEMBER.setAttribute("MB_CONTID", Nothing)
                        End If

                        'MB_RELATE
                        If Utility.isValidateData(row.GetCell(41)) Then
                            mbMEMBER.setAttribute("MB_RELATE", Trim(row.GetCell(41).ToString))
                        Else
                            mbMEMBER.setAttribute("MB_RELATE", Nothing)
                        End If

                        'MB_CONTTEL_D
                        If Utility.isValidateData(row.GetCell(42)) Then
                            mbMEMBER.setAttribute("MB_CONTTEL_D", Trim(row.GetCell(42).ToString))
                        Else
                            mbMEMBER.setAttribute("MB_CONTTEL_D", Nothing)
                        End If

                        'MB_CONTTEL_N
                        If Utility.isValidateData(row.GetCell(43)) Then
                            mbMEMBER.setAttribute("MB_CONTTEL_N", Trim(row.GetCell(43).ToString))
                        Else
                            mbMEMBER.setAttribute("MB_CONTTEL_N", Nothing)
                        End If

                        'MB_CONTMOBIL
                        If Utility.isValidateData(row.GetCell(44)) Then
                            mbMEMBER.setAttribute("MB_CONTMOBIL", Trim(row.GetCell(44).ToString))
                        Else
                            mbMEMBER.setAttribute("MB_CONTMOBIL", Nothing)
                        End If

                        'MB_CONTFAX
                        If Utility.isValidateData(row.GetCell(45)) Then
                            mbMEMBER.setAttribute("MB_CONTFAX", Trim(row.GetCell(45).ToString))
                        Else
                            mbMEMBER.setAttribute("MB_CONTFAX", Nothing)
                        End If

                        'MB_CONTFAX
                        If Utility.isValidateData(row.GetCell(46)) Then
                            mbMEMBER.setAttribute("MB_CONTFAX", Trim(row.GetCell(46).ToString))
                        Else
                            mbMEMBER.setAttribute("MB_CONTFAX", Nothing)
                        End If

                        'MB_CONT_VLG
                        If Utility.isValidateData(row.GetCell(47)) Then
                            mbMEMBER.setAttribute("MB_CONT_VLG", Trim(row.GetCell(47).ToString))
                        Else
                            mbMEMBER.setAttribute("MB_CONT_VLG", Nothing)
                        End If

                        'MB_CONT_ADDR
                        If Utility.isValidateData(row.GetCell(48)) Then
                            mbMEMBER.setAttribute("MB_CONT_ADDR", Trim(row.GetCell(48).ToString))
                        Else
                            mbMEMBER.setAttribute("MB_CONT_ADDR", Nothing)
                        End If

                        'MB_JOIN_ITEM
                        If Utility.isValidateData(row.GetCell(49)) Then
                            mbMEMBER.setAttribute("MB_JOIN_ITEM", Trim(row.GetCell(49).ToString))
                        Else
                            mbMEMBER.setAttribute("MB_JOIN_ITEM", Nothing)
                        End If

                        'MB_JOINOTH
                        If Utility.isValidateData(row.GetCell(50)) Then
                            mbMEMBER.setAttribute("MB_JOINOTH", Trim(row.GetCell(50).ToString))
                        Else
                            mbMEMBER.setAttribute("MB_JOINOTH", Nothing)
                        End If

                        'MB_FAMILY
                        If Utility.isValidateData(row.GetCell(51)) Then
                            mbMEMBER.setAttribute("MB_FAMILY", Trim(row.GetCell(51).ToString))
                        Else
                            mbMEMBER.setAttribute("MB_FAMILY", Nothing)
                        End If

                        '修改員工編號
                        mbMEMBER.setAttribute("CHGUID", "SYSTEM")

                        '修改日期
                        mbMEMBER.setAttribute("CHGDATE", Now)

                        'MB_SNORE
                        If Utility.isValidateData(row.GetCell(54)) Then
                            mbMEMBER.setAttribute("MB_SNORE", Trim(row.GetCell(54).ToString))
                        Else
                            mbMEMBER.setAttribute("MB_SNORE", Nothing)
                        End If

                        'MB_JOBTITLE
                        If Utility.isValidateData(row.GetCell(55)) Then
                            mbMEMBER.setAttribute("MB_JOBTITLE", Trim(row.GetCell(55).ToString))
                        Else
                            mbMEMBER.setAttribute("MB_JOBTITLE", Nothing)
                        End If

                        mbMEMBER.save()

                        mbMBIMP.clear()
                        mbMBIMP.LoadByPK(iMB_MEMSEQ)
                        mbMBIMP.setAttribute("MB_MEMSEQ", iMB_MEMSEQ)
                        mbMBIMP.setAttribute("IMPDATE", Now)
                        mbMBIMP.setAttribute("CHGUID", "SYSTEM")
                        mbMBIMP.save()
                    End If
                Catch ex As Exception
                    Throw
                Finally
                    If Not IsNothing(DT_MB_MEMBER) Then DT_MB_MEMBER.Dispose()
                End Try
            End While

            dbManager.commit()

            Me.Bind_RP_RESULT(1)
        Catch ex As Exception
            'Rollback後，取號檔MB_MAXID會還原到最初的號碼
            dbManager.Rollback()
            UIShareFun.showErrMsg(Me, ex)
        Finally
            UIShareFun.releaseConnection(dbManager)
        End Try
    End Sub

    Sub btLoadIMP_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles btLoadIMP.Click
        Try
            Me.Bind_RP_RESULT(1)
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    Sub Bind_RP_RESULT(ByVal iPageNUM As Integer)
        Dim DT_MB_MEMBER As DataTable = Nothing
        Dim DT_MB_SEQ As DataTable = Nothing
        Dim dbManager As DatabaseManager = UIShareFun.getDataBaseManager
        Try
            Me.initCHGPage(iPageNUM)

            Dim mbMEMBERList As New MB_MEMBERList(dbManager)
            DT_MB_MEMBER = mbMEMBERList.loadExistsIMP()

            Dim iFSTNUM As Integer = Me.m_PageSize * (iPageNUM - 1)
            Dim iENDNUM As Integer = Me.m_PageSize * iPageNUM

            Dim ROW_SELECT() As DataRow = DT_MB_MEMBER.Select("ROWNUM>=" & iFSTNUM & " AND ROWNUM<=" & iENDNUM)

            Me.RP_RESULT.DataSource = ROW_SELECT
            Me.RP_RESULT.DataBind()
            Me.PLH_RESULT.Visible = True
        Catch ex As Exception
            Throw
        Finally
            If Not IsNothing(DT_MB_MEMBER) Then DT_MB_MEMBER.Dispose()
            If Not IsNothing(DT_MB_SEQ) Then DT_MB_SEQ.Dispose()
            UIShareFun.releaseConnection(dbManager)
        End Try
    End Sub

    Sub RP_RESULT_ItemDataBound(ByVal Sender As Object, ByVal e As RepeaterItemEventArgs) Handles RP_RESULT.ItemDataBound
        Try
            If e.Item.ItemType = ListItemType.Item OrElse e.Item.ItemType = ListItemType.AlternatingItem Then
                Dim IMG_EDIT As ImageButton = e.Item.FindControl("IMG_EDIT")
                IMG_EDIT.ImageUrl = com.Azion.EloanUtility.UIUtility.getImgPath & "imgEdit.gif"

                Dim IMG_DEL As ImageButton = e.Item.FindControl("IMG_DEL")
                IMG_DEL.ImageUrl = com.Azion.EloanUtility.UIUtility.getImgPath & "imgDelete.gif"

                Dim LTL_ARRD As Literal = e.Item.FindControl("LTL_ARRD")
                Dim DRV As DataRow = CType(e.Item.DataItem, DataRow)
                LTL_ARRD.Text = DRV("MB_CITY").ToString & DRV("MB_VLG").ToString & DRV("MB_ADDR").ToString

                If IsNothing(Me.m_DT_ALLCITY) Then
                    Dim dbManager As com.Azion.NET.VB.DatabaseManager = UIShareFun.getDataBaseManager
                    Try
                        Dim apROADSECList As New AP_ROADSECList(dbManager)
                        apROADSECList.loadCityNOT_D()
                        Me.m_DT_ALLCITY = apROADSECList.getCurrentDataSet.Tables(0)
                    Catch ex As Exception
                        Throw
                    Finally
                        UIShareFun.releaseConnection(dbManager)
                    End Try
                End If
                Dim DDL_MB_CITY As DropDownList = e.Item.FindControl("DDL_MB_CITY")
                DDL_MB_CITY.DataSource = Me.m_DT_ALLCITY
                DDL_MB_CITY.DataBind()
                DDL_MB_CITY.Items.Insert(0, New ListItem("請選擇", ""))

                If Utility.isValidateData(DRV("MB_CITY")) Then
                    DDL_MB_CITY.SelectedIndex = -1
                    If Not IsNothing(DDL_MB_CITY.Items.FindByText(DRV("MB_CITY"))) Then
                        DDL_MB_CITY.Items.FindByText(DRV("MB_CITY")).Selected = True
                    End If
                End If
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Sub RP_RESULT_ItemCommand(ByVal Sender As Object, ByVal e As RepeaterCommandEventArgs) Handles RP_RESULT.ItemCommand
        Try
            If UCase(e.CommandName) = "EDIT" Then
                Me.ShowMode(e.Item, False, True)

                Dim HID_MB_CITY As HtmlInputHidden = e.Item.FindControl("HID_MB_CITY")
                If Utility.isValidateData(HID_MB_CITY.Value) Then
                    Dim dbManager As DatabaseManager = UIShareFun.getDataBaseManager
                    Try
                        Dim apROADSECList As New AP_ROADSECList(dbManager)
                        apROADSECList.loadByCityNOT_D(HID_MB_CITY.Value)
                        Dim DDL_MB_VLG As DropDownList = e.Item.FindControl("DDL_MB_VLG")
                        DDL_MB_VLG.DataSource = apROADSECList.getCurrentDataSet.Tables(0)
                        DDL_MB_VLG.DataBind()
                        DDL_MB_VLG.Items.Insert(0, New ListItem("請選擇", ""))
                        Dim HID_MB_VLG As HtmlInputHidden = e.Item.FindControl("HID_MB_VLG")
                        If Utility.isValidateData(HID_MB_VLG.Value) Then
                            DDL_MB_VLG.SelectedIndex = -1
                            If Not IsNothing(DDL_MB_VLG.Items.FindByText(HID_MB_VLG.Value)) Then
                                DDL_MB_VLG.Items.FindByText(HID_MB_VLG.Value).Selected = True
                            End If
                        End If
                    Catch ex As Exception
                        Throw
                    Finally
                        UIShareFun.releaseConnection(dbManager)
                    End Try

                    Dim btYES As Button = e.Item.FindControl("btYES")
                    Utility.setObjFocus(btYES.ClientID, Me)
                End If
            ElseIf UCase(e.CommandName) = "YES" Then
                '姓名
                Dim TXT_MB_NAME As System.Web.UI.WebControls.TextBox = e.Item.FindControl("TXT_MB_NAME")
                If Not Utility.isValidateData(Trim(TXT_MB_NAME.Text)) Then
                    com.Azion.EloanUtility.UIUtility.alert("請輸入姓名!")
                    Return
                End If

                '出生年月日
                Dim TXT_MB_BIRTH As System.Web.UI.WebControls.TextBox = e.Item.FindControl("TXT_MB_BIRTH")
                If Utility.isValidateData(TXT_MB_BIRTH) AndAlso Trim(TXT_MB_BIRTH.Text).Length <> 7 Then
                    com.Azion.EloanUtility.UIUtility.alert("出生年月日格式應為西元年YYYYMMDD!")
                    Return
                ElseIf Utility.isValidateData(TXT_MB_BIRTH) AndAlso Trim(TXT_MB_BIRTH.Text).Length = 7 Then
                    Dim sYEAR As String = CDec(Left(Trim(TXT_MB_BIRTH.Text), 3))
                    Dim sMonth As String = Trim(TXT_MB_BIRTH.Text).Substring(3, 2)
                    Dim sDay As String = Right(Trim(TXT_MB_BIRTH.Text), 2)

                    If Not IsDate(sYEAR & "/" & sMonth & "/" & sDay) Then
                        com.Azion.EloanUtility.UIUtility.alert("出生年月日輸入錯誤!")
                        Return
                    End If
                End If

                Dim dbManager As DatabaseManager = UIShareFun.getDataBaseManager
                Try
                    Dim mbMEMBER As New MB_MEMBER(dbManager)
                    If mbMEMBER.loadByPK(e.CommandArgument) Then
                        '出生年月日
                        If Utility.isValidateData(TXT_MB_BIRTH) AndAlso Trim(TXT_MB_BIRTH.Text).Length = 7 Then
                            Dim sYEAR As String = Left(Trim(TXT_MB_BIRTH.Text), 3)
                            Dim sMonth As String = Trim(TXT_MB_BIRTH.Text).Substring(3, 2)
                            Dim sDay As String = Right(Trim(TXT_MB_BIRTH.Text), 2)

                            mbMEMBER.setAttribute("MB_BIRTH", New Date(CDec(sYEAR), CDec(sMonth), CDec(sDay)))
                        End If

                        '身分證字號
                        Dim TXT_MB_ID As System.Web.UI.WebControls.TextBox = e.Item.FindControl("TXT_MB_ID")
                        mbMEMBER.setAttribute("MB_ID", Trim(TXT_MB_ID.Text))

                        '手機
                        Dim TXT_MB_MOBIL As System.Web.UI.WebControls.TextBox = e.Item.FindControl("TXT_MB_MOBIL")
                        mbMEMBER.setAttribute("MB_MOBIL", Trim(TXT_MB_MOBIL.Text))

                        '電話
                        Dim TXT_MB_TEL As System.Web.UI.WebControls.TextBox = e.Item.FindControl("TXT_MB_TEL")
                        mbMEMBER.setAttribute("MB_TEL", Trim(TXT_MB_TEL.Text))

                        'e-mail
                        Dim TXT_MB_EMAIL As System.Web.UI.WebControls.TextBox = e.Item.FindControl("TXT_MB_EMAIL")
                        mbMEMBER.setAttribute("MB_EMAIL", Trim(TXT_MB_EMAIL.Text))

                        '通訊地址
                        Dim DDL_MB_CITY As DropDownList = e.Item.FindControl("DDL_MB_CITY")
                        If Utility.isValidateData(DDL_MB_CITY.SelectedValue) Then
                            mbMEMBER.setAttribute("MB_CITY", DDL_MB_CITY.SelectedItem.Text)
                        Else
                            mbMEMBER.setAttribute("MB_CITY", Nothing)
                        End If

                        Dim DDL_MB_VLG As DropDownList = e.Item.FindControl("DDL_MB_VLG")
                        If Utility.isValidateData(DDL_MB_VLG.SelectedValue) Then
                            mbMEMBER.setAttribute("MB_VLG", DDL_MB_VLG.SelectedItem.Text)
                        Else
                            mbMEMBER.setAttribute("MB_VLG", Nothing)
                        End If

                        Dim TXT_MB_ADDR As System.Web.UI.WebControls.TextBox = e.Item.FindControl("TXT_MB_ADDR")
                        mbMEMBER.setAttribute("MB_ADDR", Trim(TXT_MB_ADDR.Text))

                        mbMEMBER.save()
                    End If
                Catch ex As Exception
                    Throw
                Finally
                    UIShareFun.releaseConnection(dbManager)
                End Try

                Me.ShowMode(e.Item, True, False)

                Me.Bind_RP_RESULT(CDec(Me.DDL_Page.SelectedValue))
            ElseIf UCase(e.CommandName) = "CANCEL" Then
                Me.ShowMode(e.Item, True, False)
            ElseIf UCase(e.CommandName) = "DELETE" Then
                Dim dbManager As DatabaseManager = UIShareFun.getDataBaseManager
                Try
                    Dim mbMBIMP As New MB_MBIMP(dbManager)
                    If mbMBIMP.LoadByPK(e.CommandArgument) Then
                        mbMBIMP.remove()
                    End If

                    Dim mbMEMBER As New MB_MEMBER(dbManager)
                    If mbMEMBER.loadByPK(e.CommandArgument) Then
                        mbMEMBER.remove()
                    End If
                Catch ex As Exception
                    Throw
                Finally
                    UIShareFun.releaseConnection(dbManager)
                End Try

                Me.Bind_RP_RESULT(CDec(Me.DDL_Page.SelectedValue))
            End If
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    Sub ShowMode(ByVal objItem As RepeaterItem, ByVal isShow As Boolean, ByVal isEDIT As Boolean)
        Try
            '姓名
            Dim LTL_MB_NAME As Literal = objItem.FindControl("LTL_MB_NAME")
            LTL_MB_NAME.Visible = isShow
            Dim TXT_MB_NAME As System.Web.UI.WebControls.TextBox = objItem.FindControl("TXT_MB_NAME")
            TXT_MB_NAME.Visible = isEDIT

            '出生年月日
            Dim LTL_MB_BIRTH As Literal = objItem.FindControl("LTL_MB_BIRTH")
            LTL_MB_BIRTH.Visible = isShow
            Dim TXT_MB_BIRTH As System.Web.UI.WebControls.TextBox = objItem.FindControl("TXT_MB_BIRTH")
            TXT_MB_BIRTH.Visible = isEDIT

            '身分證字號
            Dim LTL_MB_ID As Literal = objItem.FindControl("LTL_MB_ID")
            LTL_MB_ID.Visible = isShow
            Dim TXT_MB_ID As System.Web.UI.WebControls.TextBox = objItem.FindControl("TXT_MB_ID")
            TXT_MB_ID.Visible = isEDIT

            '手機
            Dim LTL_MB_MOBIL As Literal = objItem.FindControl("LTL_MB_MOBIL")
            LTL_MB_MOBIL.Visible = isShow
            Dim TXT_MB_MOBIL As System.Web.UI.WebControls.TextBox = objItem.FindControl("TXT_MB_MOBIL")
            TXT_MB_MOBIL.Visible = isEDIT

            '電話
            Dim LTL_MB_TEL As Literal = objItem.FindControl("LTL_MB_TEL")
            LTL_MB_TEL.Visible = isShow
            Dim TXT_MB_TEL As System.Web.UI.WebControls.TextBox = objItem.FindControl("TXT_MB_TEL")
            TXT_MB_TEL.Visible = isEDIT

            'e-mail
            Dim LTL_MB_EMAIL As Literal = objItem.FindControl("LTL_MB_EMAIL")
            LTL_MB_EMAIL.Visible = isShow
            Dim TXT_MB_EMAIL As System.Web.UI.WebControls.TextBox = objItem.FindControl("TXT_MB_EMAIL")
            TXT_MB_EMAIL.Visible = isEDIT

            '通訊地址
            Dim LTL_ARRD As Literal = objItem.FindControl("LTL_ARRD")
            LTL_ARRD.Visible = isShow
            Dim DDL_MB_CITY As DropDownList = objItem.FindControl("DDL_MB_CITY")
            DDL_MB_CITY.Visible = isEDIT
            Dim DDL_MB_VLG As DropDownList = objItem.FindControl("DDL_MB_VLG")
            DDL_MB_VLG.Visible = isEDIT
            Dim TXT_MB_ADDR As System.Web.UI.WebControls.TextBox = objItem.FindControl("TXT_MB_ADDR")
            TXT_MB_ADDR.Visible = isEDIT

            '確定 取消
            Dim IMG_EDIT As ImageButton = objItem.FindControl("IMG_EDIT")
            IMG_EDIT.Visible = isShow
            Dim btYES As Button = objItem.FindControl("btYES")
            btYES.Visible = isEDIT
            Dim btCancel As Button = objItem.FindControl("btCancel")
            btCancel.Visible = isEDIT
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Sub DDL_MB_CITY_OnSelectedIndexChanged(ByVal Sender As Object, ByVal e As System.EventArgs)
        Dim dbManager As DatabaseManager = UIShareFun.getDataBaseManager
        Try
            Dim objRepeaterItem As RepeaterItem = Sender.namingContainer
            Dim DDL_MB_CITY As DropDownList = objRepeaterItem.FindControl("DDL_MB_CITY")

            Dim DDL_MB_VLG As DropDownList = objRepeaterItem.FindControl("DDL_MB_VLG")
            DDL_MB_VLG.Items.Clear()
            Dim apROADSECList As New AP_ROADSECList(dbManager)
            apROADSECList.loadByCityNOT_D(DDL_MB_CITY.SelectedItem.Text)
            DDL_MB_VLG.DataSource = apROADSECList.getCurrentDataSet.Tables(0)
            DDL_MB_VLG.DataBind()

            DDL_MB_VLG.Items.Insert(0, New ListItem("請選擇", ""))
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        Finally
            UIShareFun.releaseConnection(dbManager)
        End Try
    End Sub

    Function getMB_MEMSEQ(ByVal dbManager As com.Azion.NET.VB.DatabaseManager) As Decimal
        Try
            Dim sProcName As String = String.Empty
            sProcName = LCase(com.Azion.NET.VB.Properties.getSchemaName) & ".MB_GET_MAXID"
            Dim inParaAL As New ArrayList
            Dim outParaAL As New ArrayList
            inParaAL.Add("01")
            inParaAL.Add("1")

            outParaAL.Add(7)

            Dim HT_Return As Hashtable = com.Azion.NET.VB.DBObject.getProcedureData(dbManager, sProcName, inParaAL, outParaAL)
            Return HT_Return.Item("@IMAXID")
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function getMB_MEMSEQ(ByVal iMB_MEMSEQ As Object, ByVal sMB_AREA As Object) As String
        Try
            If IsNumeric(iMB_MEMSEQ) AndAlso Utility.isValidateData(sMB_AREA) Then
                Return sMB_AREA & "-" & com.Azion.EloanUtility.StrUtility.FillZero(CDec(iMB_MEMSEQ), 7)
            End If
            Return String.Empty
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Function getROCDate(ByVal MB_BIRTH As Object) As String
        Try
            Return Utility.ToYYYMMDD(MB_BIRTH)
        Catch ex As Exception
            Throw
        End Try
    End Function

#Region "換頁"
    Sub initCHGPage(ByVal iPageNUM As Integer)
        Dim dbManager As DatabaseManager = UIShareFun.getDataBaseManager
        Try
            Me.PLH_Page.Visible = True

            Dim mbMEMBERList As New MB_MEMBERList(dbManager)
            Dim DT_MB_MEMBER As DataTable = mbMEMBERList.loadExistsIMP()

            Dim iRowCount As Integer = DT_MB_MEMBER.Rows.Count

            '若筆數少於頁數，不顯示換頁
            If iRowCount <= Me.m_PageSize Then
                Me.PLH_Page.Visible = False
            End If

            '目前在第?頁
            Me.lblCurrentPage.Text = iPageNUM.ToString

            '每頁?筆
            Me.lblPageSize.Text = Me.m_PageSize.ToString

            '共?筆
            Me.lblTotleSize.Text = iRowCount.ToString

            Me.DDL_Page.Items.Clear()

            Dim iTopNum As Integer = iRowCount \ Me.m_PageSize
            If (iRowCount Mod Me.m_PageSize) > 0 Then
                iTopNum += 1
            End If
            Me.lblTotalPage.Text = iTopNum.ToString

            For i As Integer = 1 To iTopNum
                Dim objListItem As New ListItem(i.ToString, i.ToString)
                Me.DDL_Page.Items.Add(objListItem)
            Next

            Me.DDL_Page.SelectedIndex = -1
            If Not IsNothing(Me.DDL_Page.Items.FindByValue(iPageNUM.ToString)) Then
                Me.DDL_Page.Items.FindByValue(iPageNUM.ToString).Selected = True
            End If
        Catch ex As Exception
            Throw
        Finally
            dbManager.releaseConnection()
        End Try
    End Sub

    '下一頁
    Sub btnLinkNext_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnLinkNext.Click, btnImgNext.Click
        Try
            Dim iRowCount As Integer = CDec(Me.lblTotleSize.Text)
            Dim iTopNum As Integer = iRowCount \ Me.m_PageSize
            If (iRowCount Mod Me.m_PageSize) > 0 Then
                iTopNum += 1
            End If

            Dim iPageNUM As Integer = CDec(Me.lblCurrentPage.Text) + 1
            If iPageNUM >= iTopNum Then
                iPageNUM = iTopNum
            End If

            Me.Bind_RP_RESULT(iPageNUM)
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    '上一頁
    Sub btnLinkPrev_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnLinkPrev.Click, btnImgPrev.Click
        Try
            Dim iPageNUM As Integer = CDec(Me.lblCurrentPage.Text) - 1
            If iPageNUM <= 0 Then
                iPageNUM = 1
            End If

            Me.Bind_RP_RESULT(iPageNUM)
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    '第一頁
    Sub btnLinkFirst_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnLinkFirst.Click, btnImgFirst.Click
        Try
            Me.Bind_RP_RESULT(1)
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    '最後一頁
    Sub btnLinkLast_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnLinkLast.Click, btnImgLast.Click
        Try
            Dim iRowCount As Integer = CDec(Me.lblTotleSize.Text)
            Dim iTopNum As Integer = iRowCount \ Me.m_PageSize
            If (iRowCount Mod Me.m_PageSize) > 0 Then
                iTopNum += 1
            End If

            Me.Bind_RP_RESULT(iTopNum)
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    '下拉換頁
    Sub DDL_Page_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DDL_Page.SelectedIndexChanged
        Try
            Me.Bind_RP_RESULT(CDec(Me.DDL_Page.SelectedValue))
        Catch ex As Exception
            UIShareFun.showErrMsg(Me, ex)
        End Try
    End Sub

    Sub ClearPageValue()
        Try
            Me.lblCurrentPage.Text = "1"
            Me.lblTotalPage.Text = String.Empty
            Me.lblPageSize.Text = String.Empty
            Me.lblTotleSize.Text = String.Empty
            Me.DDL_Page.Items.Clear()

            Me.PLH_Page.Visible = False

            Me.RP_RESULT.DataSource = Nothing

            Me.RP_RESULT.DataBind()
        Catch ex As Exception
            Throw
        End Try
    End Sub
#End Region

End Class