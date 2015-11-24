Imports com.Azion.NET.VB
Imports MBSC.MB_OP

Public Class ZipCode
    Inherits System.Web.UI.UserControl

#Region "Public Functions"
    ''' <summary>
    ''' 首次使用初始化
    ''' </summary>
    ''' <param name="dbmanager">db連接</param>
    ''' <param name="sCity">縣市Code</param>
    ''' <param name="sTown">鄉鎮Code</param>
    ''' <param name="sZIP">郵政區號</param>
    ''' <param name="sROAD">其他</param>
    ''' <param name="sSW">自有/租用</param>
    ''' <remarks></remarks>
    Public Sub FirstLoad(ByVal dbmanager As DatabaseManager, ByVal sCity As String, ByVal sTown As String, ByVal sZIP As String, ByVal sROAD As String, ByVal sSW As String, Optional ByVal sFlag As String = "", Optional sMaxLength As String = "")
        Try
            ' 綁定下拉選單首選值為“請選擇”
            Me.BindDDLCity(dbmanager)
            Me.ddlCity.SelectedIndex = -1

            ' 編輯狀態下，給控件賦值
            If Me.isValidateData(sCity) Then
                If Not IsNothing(ddlCity.Items.FindByValue(sCity)) Then
                    ddlCity.Items.FindByValue(sCity).Selected = True
                End If
            End If

            ' 綁定下拉選單首選值為“請選擇”
            Me.BindDDLTown(dbmanager, sCity)
            Me.ddlTown.SelectedIndex = -1

            ' 編輯狀態下，給控件賦值
            If Not IsNothing(ddlTown.Items.FindByValue(sTown)) Then
                ddlTown.Items.FindByValue(sTown).Selected = True
            End If

            ' 給郵編區號賦值
            Me.ltlPostCode.Text = sZIP
            Me.txtPostCode.Text = sZIP
            txtPostCode.Visible = False

            ' 給其他部份賦值
            Me.txtRoad.Text = sROAD

            '' 給自有租用賦值
            'Select Case sSW
            '    Case "0"
            '        rdoSw0.Checked = True
            '    Case "1"
            '        rdoSw1.Checked = True
            'End Select

            If Not String.IsNullOrEmpty(sFlag) Then
                plhPostCode.Visible = False
                'rdoSw0.Visible = False
                'rdoSw1.Visible = False
            Else
                plhPostCode.Visible = True
                'rdoSw0.Visible = True
                'rdoSw1.Visible = True
            End If

            ' 設置MaxLength
            If sMaxLength <> "" Then
                txtRoad.MaxLength = Convert.ToInt32(sMaxLength)
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ''' <summary>
    ''' 取得下拉縣市中文名(儲存資料時使用)
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getCityName() As String
        Try
            Return Me.ddlCity.SelectedItem.Text
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 取得下拉縣市代碼(儲存資料時使用)
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getCityValue() As String
        Try
            Return Me.ddlCity.SelectedValue
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 取得下拉鄉鎮中文名(儲存資料時使用)
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getAreaName() As String
        Try
            Return Me.ddlTown.SelectedItem.Text
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 取得下拉鄉鎮代碼(儲存資料時使用)
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getAreaValue() As String
        Try
            Return Me.ddlTown.SelectedValue
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 取得郵遞區號(儲存資料時使用)
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getZIP() As String
        Try
            Return IIf(Me.ltlPostCode.Text = "", txtPostCode.Text, ltlPostCode.Text)
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 其他部份
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function getRoad() As String
        Try
            Return Me.txtRoad.Text
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ' ''' <summary>
    ' ''' 取得自有還是租用
    ' ''' </summary>
    ' ''' <returns></returns>
    ' ''' <remarks></remarks>
    'Public Function getSW() As String
    '    Try
    '        If rdoSw0.Checked Then
    '            Return 0
    '        ElseIf rdoSw1.Checked Then
    '            Return 1
    '        Else
    '            Return ""
    '        End If
    '        'Return IIf(rdoSw0.Checked, "0", "1")
    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Function

    ''' <summary>
    ''' 初始化縣市
    ''' </summary>
    ''' <param name="dbmanager"></param>
    ''' <remarks></remarks>
    Sub BindDDLCity(ByVal dbmanager As DatabaseManager)
        Try
            ' 清空下拉選單 添加初始值"請選擇"
            Me.ddlCity.Items.Clear()
            Me.ddlCity.Items.Add(New ListItem("請選擇", ""))

            Dim ApRoadSecList As New AP_ROADSECList(dbmanager)
            Dim dtCity As DataTable

            ApRoadSecList.loadAllCity()
            dtCity = ApRoadSecList.getCurrentDataSet.Tables(0)

            ' 循環dtCity 將選項添加到下拉選單中
            For i As Integer = 0 To dtCity.Rows.Count - 1
                Dim objListItem As New ListItem(Me.CheckNull(dtCity.Rows(i).Item("CITY")), Me.CheckNull(dtCity.Rows(i).Item("CITY_CD")))
                Me.ddlCity.Items.Add(objListItem)
            Next
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ''' <summary>
    ''' 初始化鄉鎮市區
    ''' </summary>
    ''' <param name="dbManager"></param>
    ''' <param name="sCity"></param>
    ''' <remarks></remarks>
    Sub BindDDLTown(ByVal dbManager As DatabaseManager, ByVal sCity As String)
        Try
            ' 清空下拉選單 添加初始值"請選擇"
            Me.ddlTown.Items.Clear()
            Me.ddlTown.Items.Add(New ListItem("請選擇", ""))

            Dim ApRoadSecList As New AP_ROADSECList(dbManager)
            Dim dtArea As DataTable

            ApRoadSecList.loadDataByCityId(sCity)
            dtArea = ApRoadSecList.getCurrentDataSet.Tables(0)

            ' 循環dtArea 將選項添加到下拉選單中
            For i As Integer = 0 To dtArea.Rows.Count - 1
                Dim objListItem As New ListItem(Me.CheckNull(dtArea.Rows(i).Item("AREA")), Me.CheckNull(dtArea.Rows(i).Item("AREA_ID")))
                Me.ddlTown.Items.Add(objListItem)
            Next
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#Region "Event"
    ''' <summary>
    ''' 縣市自動帶出鄉鎮選單
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Sub ddlCity_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlCity.SelectedIndexChanged
        Dim dbManager As com.Azion.NET.VB.DatabaseManager = MBSC.UICtl.UIShareFun.getDataBaseManager
        Try
            ' 綁定鄉鎮市區
            Me.BindDDLTown(dbManager, Me.ddlCity.SelectedValue)
            ' 獲得焦點
            Me.setObjFocus(Me.ddlCity.ClientID, Me.Page)
        Catch ex As Exception
            Throw ex
        Finally
            dbManager.releaseConnection()
        End Try
    End Sub

    ''' <summary>
    ''' 鄉鎮市區自動帶出郵遞區號
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Sub ddlTown_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlTown.SelectedIndexChanged
        Dim dbManager As com.Azion.NET.VB.DatabaseManager = MBSC.UICtl.UIShareFun.getDataBaseManager
        Try
            ' 給郵政區號賦值
            Me.ltlPostCode.Text = Me.getZIP(dbManager, Me.ddlCity.SelectedItem.Text, Me.ddlTown.SelectedItem.Text)

            ' 如果有資料 顯示Label 如果沒有資料 開放textbox自行輸入
            If ltlPostCode.Text <> "" Then
                txtPostCode.Visible = False
            Else
                txtPostCode.Visible = True
            End If

            ' 獲得焦點
            Me.setObjFocus(Me.ddlCity.ClientID, Me.Page)
        Catch ex As Exception
            Throw ex
        Finally
            dbManager.releaseConnection()
        End Try
    End Sub

#End Region

#Region "Utility"

    ''' <summary>
    ''' 取得郵遞區號
    ''' </summary>
    ''' <param name="dbManager">db鏈接</param>
    ''' <param name="sCITY">縣市中文名</param>
    ''' <param name="sAREA">鄉鎮市區中文名</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function getZIP(ByVal dbManager As com.Azion.NET.VB.DatabaseManager, ByVal sCITY As String, ByVal sAREA As String) As String
        Try
            Dim apPOSTList As New AP_POSTLIST(dbManager)

            If apPOSTList.LoadDataByCityAndArea(sCITY, sAREA) Then

                Return apPOSTList.getCurrentDataSet.Tables(0).Rows(0)("ZIP")
            Else
                txtPostCode.Visible = True
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    ''' <summary>
    ''' 獲取焦點
    ''' </summary>
    ''' <param name="sClientid"></param>
    ''' <param name="page"></param>
    ''' <remarks></remarks>
    Sub setObjFocus(ByVal sClientid As String, ByVal page As System.Web.UI.Page)
        Dim script As String = ""

        script = "<script language='javascript'>"
        script &= "if (document.all('" & sClientid & "')!=undefined){"
        script &= "document.all('" & sClientid & "').scrollIntoView();"
        script &= "document.all('" & sClientid & "').focus();"
        script &= "}"
        script &= "</script>"

        page.RegisterStartupScript("setFocus", script)
    End Sub

    ''' <summary>
    ''' 是否為Null
    ''' </summary>
    ''' <param name="CheckValue"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CheckNull(ByVal CheckValue) As String
        If IsDBNull(CheckValue) Then
            Return ""
        Else
            Return CheckValue
        End If
    End Function

    ''' <summary>
    ''' 檢查欄位是否有效
    ''' </summary>
    ''' <param name="sInfo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function isValidateData(ByVal sInfo As String) As Boolean
        If (Not IsNothing(sInfo) AndAlso sInfo.Trim.Length > 0) Then
            Return True
        End If
        Return False
    End Function
#End Region
End Class