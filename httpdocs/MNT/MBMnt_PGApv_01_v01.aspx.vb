Imports com.Azion.NET.VB
Imports MBSC.MB_OP
Imports MBSC.UICtl

Public Class MBMnt_PGApv_01_v01
    Inherits System.Web.UI.Page

    Dim m_iUPCODE_76 As Integer = 76
    Dim m_iUPCODE_78 As Integer = 78
    Dim m_DBManager As DatabaseManager = Nothing
    Dim m_DT_78 As DataTable = Nothing

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.m_DBManager = UIShareFun.getDataBaseManager

            '會員系統
            Me.m_iUPCODE_78 = com.Azion.EloanUtility.CodeList.getAppSettings("UPCODE78")

            '維護人員
            Me.m_iUPCODE_76 = com.Azion.EloanUtility.CodeList.getAppSettings("UPCODE76")

            If Not Page.IsPostBack Then
                Me.Bind_MDUSER(Me.LB_L_MDUSER)
            End If
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

    Private Sub MBMnt_PGApv_01_v01_Unload(sender As Object, e As EventArgs) Handles Me.Unload
        UIShareFun.releaseConnection(Me.m_DBManager)
    End Sub

    Sub Bind_MDUSER(ByVal LB As ListBox)
        Dim DT_76 As DataTable = Nothing
        Try
            DT_76 = Me.getDT_AP_CODE(Me.m_iUPCODE_76)
            LB.DataTextField = "TEXT"
            LB.DataValueField = "TEXT"
            LB.DataSource = DT_76
            LB.DataBind()
        Catch ex As Exception
            Throw
        Finally
            If Not IsNothing(DT_76) Then
                DT_76.Dispose()
            End If
        End Try
    End Sub

    Function getDT_AP_CODE(ByVal iUPCODE As Integer) As DataTable
        Try
            Dim apCODEList As New AP_CODEList(Me.m_DBManager)
            apCODEList.loadByUpCode(iUPCODE)
            Return apCODEList.getCurrentDataSet.Tables(0)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Private Sub btnCHOOSE_Click(sender As Object, e As EventArgs) Handles btnCHOOSE.Click
        Try
            Dim AL_C As New ArrayList
            For i As Integer = 0 To Me.LB_L_MDUSER.Items.Count - 1
                Dim objItem As ListItem = Me.LB_L_MDUSER.Items(i)
                If objItem.Selected Then
                    Me.LB_C_MDUSER.Items.Add(objItem)
                    AL_C.Add(objItem)
                End If
            Next

            For i As Integer = 0 To AL_C.Count - 1
                Me.LB_L_MDUSER.Items.Remove(AL_C.Item(i))
            Next

            Me.LB_L_MDUSER.SelectedIndex = -1
            Me.LB_C_MDUSER.SelectedIndex = -1
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

    Private Sub btnOUT_Click(sender As Object, e As EventArgs) Handles btnOUT.Click
        Try
            Dim AL_C As New ArrayList
            For i As Integer = 0 To Me.LB_C_MDUSER.Items.Count - 1
                Dim objItem As ListItem = Me.LB_C_MDUSER.Items(i)
                If objItem.Selected Then
                    Me.LB_L_MDUSER.Items.Add(objItem)
                    AL_C.Add(objItem)
                End If
            Next

            For i As Integer = 0 To AL_C.Count - 1
                Me.LB_C_MDUSER.Items.Remove(AL_C.Item(i))
            Next

            Me.LB_L_MDUSER.SelectedIndex = -1
            Me.LB_C_MDUSER.SelectedIndex = -1
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

    Private Sub btnQRY_Click(sender As Object, e As EventArgs) Handles btnQRY.Click
        Dim DT_ACCT As New DataTable
        Try
            DT_ACCT.Columns.Add("ACCT", Type.GetType("System.String"))

            For i As Integer = 0 To Me.LB_C_MDUSER.Items.Count - 1
                Dim objItem As ListItem = Nothing
                objItem = Me.LB_C_MDUSER.Items(i)

                Dim sqlRow As DataRow = DT_ACCT.NewRow
                sqlRow("ACCT") = objItem.Value
                DT_ACCT.Rows.Add(sqlRow)
            Next

            If DT_ACCT.Rows.Count = 0 Then
                com.Azion.EloanUtility.UIUtility.showJSMsg(Me, "請選擇維護人員")
                com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "請選擇維護人員")
                Me.RP_PROG.DataSource = Nothing
                Me.RP_PROG.DataBind()
                Me.PLH_PROG.Visible = False
                Return
            Else
                Me.PLH_LIST.Visible = False
                Me.PLH_PROG.Visible = True
            End If

            Me.RP_PROG.DataSource = DT_ACCT
            Me.RP_PROG.DataBind()
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        Finally
            If Not IsNothing(DT_ACCT) Then
                DT_ACCT.Dispose()
            End If
        End Try
    End Sub

    Private Sub RP_PROG_ItemCommand(source As Object, e As Web.UI.WebControls.RepeaterCommandEventArgs) Handles RP_PROG.ItemCommand
        Try
            If e.CommandName = "ALL" Then
                Dim CB_PROG As CheckBoxList = e.Item.FindControl("CB_PROG")
                For Each objItem As ListItem In CB_PROG.Items
                    objItem.Selected = True
                Next
            ElseIf e.CommandName = "CANCELALL" Then
                Dim CB_PROG As CheckBoxList = e.Item.FindControl("CB_PROG")
                For Each objItem As ListItem In CB_PROG.Items
                    objItem.Selected = False
                Next
            End If
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

    Private Sub RP_PROG_ItemDataBound(sender As Object, e As RepeaterItemEventArgs) Handles RP_PROG.ItemDataBound
        Try
            If e.Item.ItemType = ListItemType.Item OrElse e.Item.ItemType = ListItemType.AlternatingItem Then
                Dim CB_PROG As CheckBoxList = e.Item.FindControl("CB_PROG")

                If IsNothing(Me.m_DT_78) Then
                    Me.m_DT_78 = Me.getDT_AP_CODE(Me.m_iUPCODE_78)
                End If

                CB_PROG.DataTextField = "TEXT"
                CB_PROG.DataValueField = "VALUE"
                CB_PROG.DataSource = Me.m_DT_78
                CB_PROG.DataBind()

                Dim DT_MB_FNCAUTH As DataTable = Nothing
                Try
                    Dim LTL_ACCT As Literal = e.Item.FindControl("LTL_ACCT")
                    Dim mbFNCAUTHList As New MB_FNCAUTHList(Me.m_DBManager)
                    mbFNCAUTHList.LoadByACCT_CODE(LTL_ACCT.Text, Me.m_iUPCODE_78)
                    CB_PROG.SelectedIndex = -1
                    For Each ROW As DataRow In mbFNCAUTHList.getCurrentDataSet.Tables(0).Rows
                        If Not IsNothing(CB_PROG.Items.FindByValue(ROW("MB_FUCCODE").ToString)) Then
                            CB_PROG.Items.FindByValue(ROW("MB_FUCCODE").ToString).Selected = True
                        End If
                    Next
                Catch ex As Exception
                    Throw
                Finally
                    If Not IsNothing(DT_MB_FNCAUTH) Then
                        DT_MB_FNCAUTH.Dispose()
                    End If
                End Try
            End If
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Private Sub btnSAVE_Click(sender As Object, e As EventArgs) Handles btnSAVE.Click
        Try
            Try
                Me.m_DBManager.beginTran()

                Dim mbFNCAUTHList As New MB_FNCAUTHList(Me.m_DBManager)
                Dim mbFNCAUTH As New MB_FNCAUTH(Me.m_DBManager)
                For i As Integer = 0 To Me.RP_PROG.Items.Count - 1
                    Dim objItem As RepeaterItem = Me.RP_PROG.Items(i)

                    '維護人員帳號
                    Dim LTL_ACCT As Literal = objItem.FindControl("LTL_ACCT")
                    mbFNCAUTHList.clear()
                    mbFNCAUTHList.DelByACCT(LTL_ACCT.Text)

                    '功能清單
                    Dim CB_PROG As CheckBoxList = objItem.FindControl("CB_PROG")
                    For Each L_ITEM As ListItem In CB_PROG.Items
                        If L_ITEM.Selected Then
                            mbFNCAUTH.clear()
                            mbFNCAUTH.LoadByPK(LTL_ACCT.Text, Me.m_iUPCODE_78, L_ITEM.Value)
                            'MB_ACCT	varchar(100)	utf8_general_ci	NO	PRI			select,insert,update,references	帳號
                            mbFNCAUTH.setAttribute("MB_ACCT", LTL_ACCT.Text)
                            'MB_UPCODE	decimal(10,0)		NO	PRI	0		select,insert,update,references	父類別 REF AP_CODE.UPCODE
                            mbFNCAUTH.setAttribute("MB_UPCODE", Me.m_iUPCODE_78)
                            'MB_FUCCODE	varchar(20)	utf8_general_ci	NO	PRI			select,insert,update,references	功能代號Ref AP_CODE UPCODE=78->VALUE
                            mbFNCAUTH.setAttribute("MB_FUCCODE", L_ITEM.Value)
                            'CHGUID	varchar(500)	utf8_general_ci	YES				select,insert,update,references	修改員編
                            mbFNCAUTH.setAttribute("CHGUID", com.Azion.EloanUtility.UIUtility.getLoginUserID)
                            'CHGDATE	date		YES				select,insert,update,references	修改日期
                            mbFNCAUTH.setAttribute("CHGDATE", Now)
                            mbFNCAUTH.save()
                        End If
                    Next
                Next

                Me.m_DBManager.commit()
            Catch ex As Exception
                Me.m_DBManager.Rollback()
                Throw
            End Try

            Me.ClearUI()
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, "儲存成功")
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

    Private Sub btnBack_Click(sender As Object, e As EventArgs) Handles btnBack.Click
        Try
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub

    Sub ClearUI()
        Try
            Me.RP_PROG.DataSource = Nothing
            Me.RP_PROG.DataBind()

            Me.PLH_PROG.Visible = False

            Me.Bind_MDUSER(Me.LB_L_MDUSER)
            Me.LB_C_MDUSER.Items.Clear()
            Me.PLH_LIST.Visible = True
        Catch ex As Exception
            Throw
        End Try
    End Sub

    Private Sub btnCHOOSE_ALL_Click(sender As Object, e As EventArgs) Handles btnCHOOSE_ALL.Click
        Try
            Me.LB_L_MDUSER.Items.Clear()

            Me.Bind_MDUSER(Me.LB_C_MDUSER)
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub


    Private Sub btnOUT_ALL_Click(sender As Object, e As EventArgs) Handles btnOUT_ALL.Click
        Try
            Me.LB_C_MDUSER.Items.Clear()

            Me.Bind_MDUSER(Me.LB_L_MDUSER)
        Catch ex As Exception
            com.Azion.EloanUtility.UIUtility.showErrMsg(Me, ex)
        End Try
    End Sub
End Class