Option Explicit On
Option Strict On

Imports System.Web.UI.HtmlControls
Imports System.Web.UI.WebControls
Imports System.Web

Public NotInheritable Class UIUtility

    ''' <summary>
    ''' 是否為測試模式
    ''' </summary>
    ''' <returns>Boolean</returns>
    ''' <remarks>
    ''' [Titan] 	2011/07/20	Created
    ''' </remarks>
    Public Shared Function isTesting() As Boolean
        Dim sConfig As String = String.Empty
        Dim currContext As System.Web.HttpContext = System.Web.HttpContext.Current
        If Not IsNothing(currContext) Then
            sConfig = CStr(System.Web.HttpContext.Current.Application("EloanConf"))
            Return CBool(FileUtility.getAppSettings("Testing", sConfig))
        End If

        Return CBool(FileUtility.getAppSettings("Testing"))
    End Function

    ''' <summary>
    ''' 是否為debug模式
    ''' </summary>
    ''' <returns>Boolean</returns>
    ''' <remarks>
    ''' [Titan] 	2011/11/06	Created
    ''' </remarks>
    Public Shared Function isDebug() As Boolean
        Dim sConfig As String = String.Empty
        Dim currContext As System.Web.HttpContext = System.Web.HttpContext.Current
        If Not IsNothing(currContext) Then
            sConfig = CStr(System.Web.HttpContext.Current.Application("EloanConf"))
            Return CBool(FileUtility.getAppSettings("debug", sConfig))
        End If

        Return CBool(FileUtility.getAppSettings("debug"))
    End Function

    ''' <summary>
    ''' 判斷Session是否遺失，假使遺失由Request，賦予Session值，並從新設定TimeOut
    ''' </summary>
    ''' <remarks>
    ''' [Titan] 	2011/10/05	Created
    ''' </remarks>
    Public Shared Sub newSession()
        'Dim sSID As String = System.Web.HttpContext.Current.Request.QueryString("SID")
        'If sSID <> Nothing AndAlso System.Web.HttpContext.Current.Session.SessionID <> sSID Then
        '    System.Web.HttpContext.Current.Session.Timeout = 600
        '    For Each sReqSessionInfo As String In HttpUtility.UrlDecode(System.Web.HttpContext.Current.Request.QueryString("sessionInfo")).Split(CChar("|"))
        '        If sReqSessionInfo <> Nothing Then
        '            Dim sKey As String = sReqSessionInfo.Split(CChar("="))(0)
        '            Dim sValue As String = sReqSessionInfo.Split(CChar("="))(1)
        '            System.Web.HttpContext.Current.Session(sKey) = sValue
        '        End If
        '    Next
        'End If
    End Sub

    ''' <summary>
    ''' 由 session or Request 取得caseid
    ''' </summary>
    ''' <returns>String</returns>
    ''' <remarks>
    ''' [Titan] 	2011/07/20	Created
    ''' </remarks>
    Public Shared Function getCaseID() As String
        Dim currContext As System.Web.HttpContext = System.Web.HttpContext.Current
        Dim sCaseID As String = ""

        If IsNothing(currContext) OrElse IsNothing(currContext.Session) Then Return sCaseID

        If ValidateUtility.isValidateData(currContext.Request("caseid")) Then
            sCaseID = currContext.Request("caseid")
        ElseIf ValidateUtility.isValidateData(currContext.Request("caseid")) Then
            sCaseID = CStr(currContext.Request("caseid"))
        End If

        Return sCaseID
    End Function

    ''' <summary>
    ''' 由 Request 取得IDBAN(徵信對象統編)
    ''' </summary>
    ''' <returns>String</returns>
    ''' <remarks>
    ''' [Ted] 	2011/12/19	Created
    ''' </remarks>
    Public Shared Function getIDBAN() As String
        Dim currContext As System.Web.HttpContext = System.Web.HttpContext.Current
        Dim sIDBAN As String = String.Empty

        If IsNothing(currContext) OrElse IsNothing(currContext.Session) Then Return sIDBAN

        If ValidateUtility.isValidateData(currContext.Request("IDBAN")) Then
            sIDBAN = currContext.Request("IDBAN")
            'ElseIf ValidateUtility.isValidateData(currContext.Request("IDBAN")) Then
            '    sIDBAN = CStr(currContext.Request("IDBAN"))
        End If

        Return sIDBAN
    End Function

    ''' <summary>
    ''' 由 session or Request 取得userid
    ''' </summary>
    ''' <returns>String</returns>
    ''' <remarks>
    ''' [Titan] 	2011/07/20	Created
    ''' </remarks>
    Public Shared Function getLoginUserID() As String
        Dim currContext As System.Web.HttpContext = System.Web.HttpContext.Current
        Dim sUserId As String = String.Empty

        If IsNothing(currContext) OrElse IsNothing(currContext.Session) Then Return sUserId

        If ValidateUtility.isValidateData(currContext.Session("userid")) Then
            sUserId = CStr(currContext.Session("userid"))
        ElseIf ValidateUtility.isValidateData(currContext.Request("userid")) Then
            sUserId = currContext.Request("userid")
            currContext.Session("userid") = sUserId
        End If

        Return sUserId
    End Function

    ''' <summary>
    ''' 由 session or Request 取得安泰 userid(5碼)
    ''' </summary>
    ''' <returns>String</returns>
    ''' <remarks>
    ''' [Titan] 	2011/07/20	Created
    ''' </remarks>
    Public Shared Function getEnTieLoginUserID() As String
        Dim currContext As System.Web.HttpContext = System.Web.HttpContext.Current
        Dim sUserId As String = String.Empty

        If IsNothing(currContext) OrElse IsNothing(currContext.Session) Then Return sUserId

        If ValidateUtility.isValidateData(currContext.Session("userid")) Then
            sUserId = CStr(currContext.Session("userid"))
            sUserId = Right(sUserId, 5)
        ElseIf ValidateUtility.isValidateData(currContext.Request("userid")) Then
            sUserId = currContext.Request("userid")
            currContext.Session("userid") = sUserId
            sUserId = Right(sUserId, 5)
        End If

        Return sUserId
    End Function

    ''' <summary>
    ''' 取得JCIC SP_GETJCICDATA連接字串
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function getJCICConnectionString() As String
        Dim currContext As System.Web.HttpContext = System.Web.HttpContext.Current
        Dim sConnectionString As String = String.Empty

        If IsNothing(currContext) OrElse IsNothing(currContext.Session) Then Return sConnectionString

        If ValidateUtility.isValidateData(currContext.Application("JCICDSN")) Then
            sConnectionString = CStr(currContext.Application("JCICDSN"))
        End If

        Return sConnectionString
    End Function

    ''' <summary>
    ''' 取得UNI RISK風控資料庫連線字串
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function getUNIRISKConnectionString() As String
        Dim currContext As System.Web.HttpContext = System.Web.HttpContext.Current
        Dim sConnectionString As String = String.Empty

        If IsNothing(currContext) OrElse IsNothing(currContext.Session) Then Return sConnectionString

        If ValidateUtility.isValidateData(currContext.Application("UNIDSN")) Then
            sConnectionString = CStr(currContext.Application("UNIDSN"))
        End If

        Return sConnectionString
    End Function

    ''' <summary>
    ''' 取得TEJ資料庫連線字串
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function getTEJConnectionString() As String
        Dim currContext As System.Web.HttpContext = System.Web.HttpContext.Current
        Dim sConnectionString As String = String.Empty

        If IsNothing(currContext) OrElse IsNothing(currContext.Session) Then Return sConnectionString

        If ValidateUtility.isValidateData(currContext.Application("TEJDSN")) Then
            sConnectionString = CStr(currContext.Application("TEJDSN"))
        End If

        Return sConnectionString
    End Function

    ''' <summary>
    ''' 由 session or Request 取得經過NOTES驗證的Session id
    ''' </summary>
    ''' <returns>String</returns>
    ''' <remarks>
    ''' [Ted] 	2012/01/03	Created
    ''' </remarks>
    Public Shared Function getNOTESSessionID() As String
        Dim currContext As System.Web.HttpContext = System.Web.HttpContext.Current
        Dim sSessionID As String = String.Empty

        If IsNothing(currContext) OrElse IsNothing(currContext.Session) Then Return sSessionID

        If ValidateUtility.isValidateData(currContext.Session("sessionid")) Then
            sSessionID = CStr(currContext.Session("sessionid"))
        ElseIf ValidateUtility.isValidateData(currContext.Request("sessionid")) Then
            sSessionID = currContext.Request("sessionid")
            currContext.Session("sessionid") = sSessionID
        End If

        Return sSessionID
    End Function

    ''' <summary>
    ''' 由 Request 取得stepno
    ''' </summary>
    ''' <returns>String</returns>
    ''' <remarks>
    ''' [Titan] 	2011/07/20	Created
    ''' </remarks>
    Public Shared Function getStepno() As String
        Dim currContext As System.Web.HttpContext = System.Web.HttpContext.Current
        Dim sStepno As String = String.Empty

        If IsNothing(currContext) OrElse IsNothing(currContext.Request) Then Return sStepno

        If ValidateUtility.isValidateData(currContext.Request("stepno")) Then
            sStepno = currContext.Request("stepno")
        End If

        Return sStepno
    End Function

    ''' <summary>
    ''' 由 Request 取得Watch Stepno
    ''' 取得查詢的步驟
    ''' </summary>
    ''' <returns>String</returns>
    ''' <remarks>
    ''' [Titan] 	2011/12/16	Created
    ''' </remarks>
    Public Shared Function getWatchStepno() As String
        Dim currContext As System.Web.HttpContext = System.Web.HttpContext.Current
        Dim sWatchStepNo As String = String.Empty

        If IsNothing(currContext) OrElse IsNothing(currContext.Request) Then Return sWatchStepNo

        If ValidateUtility.isValidateData(currContext.Request("WatchStepNo")) Then
            sWatchStepNo = currContext.Request("WatchStepNo")
        End If

        Return sWatchStepNo
    End Function

    ''' <summary>
    ''' 由 Request 取得 版號
    ''' </summary>
    ''' <returns>String</returns>
    ''' <remarks>
    ''' [Titan] 	2011/12/16	Created
    ''' </remarks>
    Public Shared Function getVersion() As String
        Dim currContext As System.Web.HttpContext = System.Web.HttpContext.Current
        Dim sWatchStepNo As String = String.Empty

        If IsNothing(currContext) OrElse IsNothing(currContext.Request) Then Return sWatchStepNo

        If ValidateUtility.isValidateData(currContext.Request("Version")) Then
            sWatchStepNo = currContext.Request("Version")
        End If

        Return sWatchStepNo
    End Function

    ''' <summary>
    ''' 由 session or Request 取得BrID
    ''' </summary>
    ''' <returns>String</returns>
    ''' <remarks>
    ''' [Titan] 	2011/07/20	Created
    ''' </remarks>
    Public Shared Function getLoginBrID() As String
        Dim currContext As System.Web.HttpContext = System.Web.HttpContext.Current
        Dim sBrid As String = ""

        If IsNothing(currContext) OrElse IsNothing(currContext.Session) Then Return sBrid

        If ValidateUtility.isValidateData(currContext.Session("brid")) Then
            sBrid = CStr(currContext.Session("brid"))
        ElseIf ValidateUtility.isValidateData(currContext.Request("brid")) Then
            sBrid = currContext.Request("brid")
            currContext.Session("brid") = sBrid
        End If

        Return sBrid
    End Function

    ''' <summary>
    ''' 由 session or Request 取得RemoteFlag
    ''' </summary>
    ''' <returns>Boolean</returns>
    ''' <remarks>
    ''' [Titan] 	2011/07/20	Created
    ''' 尚未實作
    ''' </remarks>
    Public Shared Function isRemoteMode() As Boolean
        Dim currContext As System.Web.HttpContext = System.Web.HttpContext.Current
        Dim RemoteFlag As String = String.Empty

        If IsNothing(currContext) OrElse IsNothing(currContext.Session) Then Return False 

        Return False
    End Function

    ''' <summary>
    ''' 取得eLoan專案目錄
    ''' </summary>
    ''' <returns>String</returns>
    ''' <remarks>
    ''' [Titan] 	2011/07/20	Created
    ''' </remarks>
    Public Shared Function getRootPath(Optional ByVal hasSessionId As Boolean = True) As String

        'If hasSessionId Then Return HttpContext.Current.Request.ApplicationPath '& "/(S(" & System.Web.HttpContext.Current.Session.SessionID & "))"
        'Return HttpContext.Current.Request.ApplicationPath
        'HttpContext.Current.Request.PhysicalApplicationPath

        Return String.Empty

        'Dim sROOT As String = String.Empty
        'Dim sPath As String = String.Empty
        'sPath = HttpContext.Current.Request.ApplicationPath
        'If com.Azion.EloanUtility.ValidateUtility.isValidateData(sPath) AndAlso Right(sPath, 1) = "/" Then
        '    sPath = Left(sPath, sPath.Length - 1)
        'End If
        'sROOT = "//" & HttpContext.Current.Request.Url.Authority & sPath
        'Return sROOT
    End Function

    ''' <summary>
    ''' 取得eLoan專案目錄
    ''' </summary>
    ''' <returns>String</returns>
    ''' <remarks>
    ''' [Titan] 	2011/07/20	Created
    ''' </remarks>
    Public Shared Function getRootPathClient(Optional ByVal hasSessionId As Boolean = True) As String

        'If hasSessionId Then Return HttpContext.Current.Request.ApplicationPath '& "/(S(" & System.Web.HttpContext.Current.Session.SessionID & "))"
        'Return HttpContext.Current.Request.ApplicationPath
        'HttpContext.Current.Request.PhysicalApplicationPath

        'Return ""

        Dim sROOT As String = String.Empty
        Dim sPath As String = String.Empty
        sPath = HttpContext.Current.Request.ApplicationPath
        If com.Azion.EloanUtility.ValidateUtility.isValidateData(sPath) AndAlso Right(sPath, 1) = "/" Then
            sPath = Left(sPath, sPath.Length - 1)
        End If
        sROOT = "//" & HttpContext.Current.Request.Url.Authority & sPath
        Return sROOT
    End Function

    ''' <summary>
    ''' 取得企金目錄
    ''' </summary>
    ''' <returns>String</returns>
    ''' <remarks>
    ''' [Titan] 	2011/07/20	Created
    ''' </remarks>
    Public Shared Function getENPath() As String
        Return getRootPath() & "/EN/"
    End Function

    ''' <summary>
    ''' 取得企金目錄
    ''' </summary>
    ''' <returns>String</returns>
    ''' <remarks>
    ''' [Titan] 	2011/07/20	Created
    ''' </remarks>
    Public Shared Function getGAPath() As String
        Return getRootPath() & "/GA/"
    End Function

    ''' <summary>
    ''' 取得WebService目錄
    ''' </summary>
    ''' <returns>String</returns>
    ''' <remarks>
    ''' [Titan] 	2011/07/20	Created
    ''' </remarks>
    Public Shared Function getWSPath() As String
        Return getRootPath() & "/WebService/"
    End Function

    ''' <summary>
    ''' 取得UI Ctl目錄
    ''' </summary>
    ''' <returns>String</returns>
    ''' <remarks>
    ''' [Titan] 	2011/12/23	Created
    ''' </remarks>
    Public Shared Function getUICtlPath() As String
        Return getRootPath() & "/UICtl/"
    End Function
    ''' <summary>
    ''' 取得img目錄
    ''' </summary>
    ''' <returns>String</returns>
    ''' <remarks>
    ''' [Titan] 	2011/07/20	Created
    ''' </remarks>
    Public Shared Function getImgPath() As String
        Return getRootPath() & "/img/"
    End Function

    ''' <summary>
    ''' 取得img目錄
    ''' </summary>
    ''' <returns>String</returns>
    ''' <remarks>
    ''' [Titan] 	2011/07/20	Created
    ''' </remarks>
    Public Shared Function geteLandPDFPath() As String
        Dim sConfig As String = String.Empty
        Dim currContext As System.Web.HttpContext = System.Web.HttpContext.Current
        If Not IsNothing(currContext) Then
            sConfig = CStr(System.Web.HttpContext.Current.Application("EloanConf"))
            Return FileUtility.getAppSettings("eLandPDFPath", sConfig)
        End If

        Return FileUtility.getAppSettings("eLandPDFPath")
    End Function

    ''' <summary>
    ''' Response.Redirect
    ''' </summary>
    ''' <param name="sURL">String</param>
    ''' <remarks>
    ''' [Titan] 	2011/08/24	Created
    ''' </remarks>
    Public Shared Sub Redirect(ByVal sURL As String)
        Try
            If System.Web.HttpContext.Current.Response.IsClientConnected Then
                'com.Azion.EloanUtility.UIUtility.genSessionInfo(sURL)
                System.Web.HttpContext.Current.Response.Redirect(sURL, False)
                System.Web.HttpContext.Current.ApplicationInstance.CompleteRequest()
            Else
                System.Web.HttpContext.Current.Response.End()
            End If
            'com.Azion.EloanUtility.UIUtility.Transfer(sURL)
        Catch ex As System.Threading.AbandonedMutexException
        Catch ex As ArgumentNullException
            Throw ex
        Catch ex As ArgumentException
            Throw ex
        Catch ex As HttpException
            Throw ex
        Catch ex As ApplicationException
            Throw ex
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    ''' <summary>
    ''' Server.Transfer
    ''' </summary>
    ''' <param name="sURL">String</param>
    ''' <remarks>
    ''' [Titan] 	2011/08/24	Created
    ''' </remarks>
    Public Shared Sub Transfer(ByVal sURL As String)
        Try
            System.Web.HttpContext.Current.Server.Transfer(sURL, False)
            'System.Web.HttpContext.Current.ApplicationInstance.CompleteRequest()
        Catch ex As System.Threading.AbandonedMutexException
        Catch ex As System.Threading.ThreadAbortException
        Catch ex As Exception
            Throw ex
        End Try
    End Sub


    Public Shared Function getXMLPath() As String
        Return System.Web.HttpContext.Current.Request.PhysicalApplicationPath & "/EN/XML/"
    End Function

    Public Shared Sub clearMsg(ByVal page As System.Web.UI.Page)
        Dim lblErrorMsg As Label = CType(page.FindControl("lblErrorMsg"), Label)
        Dim palErrorMsg As Panel = CType(page.FindControl("ErrorMsg"), Panel)
        If Not IsNothing(lblErrorMsg) Then
            lblErrorMsg.Text = ""
            lblErrorMsg.Visible = False
        End If
        If Not IsNothing(palErrorMsg) Then
            palErrorMsg.Visible = False
        End If
    End Sub

#Region "UI"
    ''' <summary>
    ''' disable all Control
    ''' </summary>
    ''' <param name="control">控制項</param>
    ''' <param name="sDisplayCtl">子控制項</param>
    ''' <remarks>
    ''' [Titan] 	2011/07/20	Created 
    ''' </remarks>
    Public Shared Sub setControlRead(ByVal control As Web.UI.Control, Optional ByVal sDisplayCtl As Web.UI.Control = Nothing)
        For Each ctl As Web.UI.Control In control.Controls
            setCtlRead(ctl)
            If ctl.HasControls AndAlso Not ctl Is sDisplayCtl Then
                setControlRead(ctl, sDisplayCtl)
            End If
        Next
        If control.Controls.Count = 0 Then
            setCtlRead(control)
        End If
    End Sub

    Private Shared Sub setCtlRead(ByVal ctl As Web.UI.Control)
        If TypeOf (ctl) Is TextBox Then
            Dim txt As TextBox = CType(ctl, TextBox)
            'txt.ReadOnly = True
            '發文字號Textbox調整大小
            txt.Attributes.Add("ReadOnly", "ReadOnly")
            txt.Columns = txt.Text.Length
            'If txt.CssClass.Equals("bordernum") OrElse txt.CssClass.Equals("nobordernum") Then
            '    txt.CssClass = "nobordernum1"
            'Else
            '    txt.CssClass = "bgnobordertxt"
            'End If
            'If txt.TextMode = TextBoxMode.MultiLine Then
            '    txt.Attributes("Style") = "overflow: visible; border:none;"
            'End If '

            Dim sPreCssClass As String = ""
            sPreCssClass = txt.CssClass

            txt.Style.Item("overflow") = "visible"
            txt.Style.Item("border") = "none"

            If ValidateUtility.isValidateData(sPreCssClass) AndAlso InStr(UCase(sPreCssClass), "NUM") > 0 Then
                If Not ValidateUtility.isValidateData(txt.Style.Item("text-align")) Then
                    txt.Style.Item("text-align") = "right"
                End If

                txt.Style.Item("background") = "transparent"
            Else
                txt.CssClass = "td2"
            End If
        ElseIf TypeOf (ctl) Is HtmlTextArea Then
            Dim txt As HtmlTextArea = CType(ctl, HtmlTextArea)
            txt.Disabled = True
            txt.Style.Item("overflow") = "visible"
            txt.Style.Item("border") = "none"
        ElseIf TypeOf (ctl) Is RadioButton Then
            Dim rad As RadioButton = CType(ctl, RadioButton)
            'rad.Font.Bold = True
            rad.AutoPostBack = False
            If rad.Checked Then
            Else
                rad.GroupName = rad.ID
                rad.Attributes.Add("onclick", "return false;")
            End If
        ElseIf TypeOf (ctl) Is RadioButtonList Then
            Dim rad As RadioButtonList = CType(ctl, RadioButtonList)
            rad.AutoPostBack = False
            If rad.SelectedIndex = -1 Then
                rad.Attributes.Add("onclick", "return false;")
            Else
                ' rad.Attributes.Add("onclick", "document.all('" & rad.ClientID & "_" & rad.SelectedIndex.ToString & "').checked=true;return false;")
                rad.Attributes.Add("onclick", String.Format("document.all('{0}_{1}').checked=true;return false;", rad.ClientID, rad.SelectedIndex))
            End If


        ElseIf TypeOf (ctl) Is HtmlInputRadioButton Then
            Dim rad As HtmlInputRadioButton = CType(ctl, HtmlInputRadioButton)
            If rad.Checked Then
            Else
                rad.Name = rad.ID
                rad.Attributes.Add("onclick", "return false;")
            End If
        ElseIf TypeOf (ctl) Is Button Then
            Dim btn As Button = CType(ctl, Button)
            If btn.ID.IndexOf("btnQuery") = -1 Then
                btn.Visible = False
            End If
        ElseIf TypeOf (ctl) Is HtmlButton Then
            Dim btn As HtmlButton = CType(ctl, HtmlButton)
            If btn.ID.IndexOf("btnQuery") = -1 Then
                btn.Visible = False
            End If
        ElseIf TypeOf (ctl) Is HtmlInputButton Then
            Dim btn As HtmlInputButton = CType(ctl, HtmlInputButton)
            If btn.ID.IndexOf("btnQuery") = -1 Then
                btn.Visible = False
            End If
        ElseIf TypeOf (ctl) Is ImageButton Then
            Dim btn As ImageButton = CType(ctl, ImageButton)
            If btn.ImageUrl.IndexOf("imgView") = -1 Then 'OrElse btn.ImageUrl.IndexOf("search") = -1 
                btn.Visible = False
            End If
        ElseIf TypeOf (ctl) Is DropDownList Then
            'Dim ddl As DropDownList = CType(ctl, DropDownList)
            'If Not IsNothing(ddl.Parent) Then
            '    If TypeOf (ddl.Parent) Is HtmlTableCell Then
            '        Dim cell As HtmlTableCell = CType(ddl.Parent, HtmlTableCell)
            '        If ddl.Items.Count > 0 Then
            '            cell.InnerText = ddl.SelectedItem.Text
            '        End If
            '    End If
            'End If
            'ElseIf TypeOf (ctl) Is DropDownList Then
            '    Dim ddl As DropDownList = CType(ctl, DropDownList)
            '    If Not IsNothing(ddl.Parent) Then
            '        If TypeOf (ddl.Parent) Is HtmlTableCell Then
            '            Dim cell As HtmlTableCell = CType(ddl.Parent, HtmlTableCell)
            '            If ddl.Items.Count > 0 Then
            '                cell.InnerText = ddl.SelectedItem.Text
            '            End If
            '            ddl.Visible = False
            '        ElseIf TypeOf (ddl.Parent) Is Panel Then
            '            For i As Integer = 0 To CType(ddl.Parent, Panel).Controls.Count - 1
            '                If ddl.Items.Count > 0 Then
            '                    If ddl.Parent.Controls(i).ID = ddl.ID Then
            '                        If TypeOf ddl.Parent.Controls(i - 1) Is LiteralControl Then
            '                            CType(ddl.Parent.Controls(i - 1), LiteralControl).Text &= ddl.SelectedItem.Text
            '                            ddl.Visible = False
            '                            Exit For
            '                        End If
            '                    End If
            '                End If
            '            Next
            '        Else
            '            Dim sValue As String = ddl.SelectedValue
            '            Dim iIndex As Integer = 0
            '            While ddl.Items.Count > 1
            '                If ddl.Items(0).Value = sValue Then
            '                    iIndex = 1
            '                End If
            '                ddl.Items.Remove(ddl.Items(iIndex))
            '            End While
            '        End If
            '    End If
        ElseIf TypeOf (ctl) Is LinkButton Then
            Dim linkButton As LinkButton = CType(ctl, LinkButton)
            If linkButton.ClientID.ToUpper.IndexOf("PAGETAB") = -1 AndAlso linkButton.ClientID.ToUpper.IndexOf("LBSERAIL") = -1 AndAlso _
                linkButton.ClientID.ToUpper.IndexOf("LBGOMAN") = -1 AndAlso linkButton.ClientID.ToUpper.IndexOf("LNKBTNQUERY") = -1 Then
                linkButton.Visible = False
            End If
        ElseIf TypeOf (ctl) Is CheckBox Then
            Dim checkBox As CheckBox = CType(ctl, CheckBox)
            checkBox.AutoPostBack = False
            ' checkBox.Attributes.Add("onclick", "document.all('" & checkBox.ClientID & "').checked=true;return false;")
            checkBox.Attributes.Add("onclick", String.Format("document.all('{0}').checked=true;return false;", checkBox.ClientID))
            'checkBox.Enabled = False

        ElseIf TypeOf (ctl) Is CheckBoxList Then
            Dim checkBox As CheckBoxList = CType(ctl, CheckBoxList)
            checkBox.AutoPostBack = False
            If checkBox.SelectedIndex = -1 Then
                checkBox.Attributes.Add("onclick", "return false;")
            Else
                ' rad.Attributes.Add("onclick", "document.all('" & rad.ClientID & "_" & rad.SelectedIndex.ToString & "').checked=true;return false;")
                checkBox.Attributes.Add("onclick", String.Format("document.all('{0}_{1}').checked=true;return false;", checkBox.ClientID, checkBox.SelectedIndex))
            End If

        ElseIf TypeOf (ctl) Is System.Web.UI.HtmlControls.HtmlTableCell Then
            Dim td As System.Web.UI.HtmlControls.HtmlTableCell = CType(ctl, System.Web.UI.HtmlControls.HtmlTableCell)
            If ValidateUtility.isValidateData(td) AndAlso ValidateUtility.isValidateData(td.ID) AndAlso td.ID.ToUpper.IndexOf("NEXTPAGE") <> -1 Then
                td.Visible = False
            End If
        End If
    End Sub
#End Region

#Region " 顯示資訊(JS) "
    Public Shared Sub showJSMsg(ByVal page As System.Web.UI.Page, ByVal sMsg As String)
        Dim sJS As String = "<script Language='JAVAScript'>"
        sJS &= "alert('" & sMsg & "');"
        sJS &= "</script>"
        page.Response.Write(sJS)
    End Sub

    Public Shared Sub showJSMsg(ByVal sMsg As String)
        Dim sJS As String = "<script Language='JAVAScript'>"
        sJS &= "alert('" & sMsg & "');"
        sJS &= "</script>"
        HttpContext.Current.Response.Write(sJS)
    End Sub
#End Region

#Region " 顯示錯誤資訊 "

#Region "錯誤資訊的開關 "

#Region "錯誤資訊的開關 "

    Public Shared Sub closeErrMsg()
        Dim page As UI.Page = CType(HttpContext.Current.Handler, UI.Page)
        Try
            Dim lblErrorMsg As Label = CType(page.FindControl("lblErrorMsg"), System.Web.UI.WebControls.Label)
            Dim ErrorMsg As Panel = CType(page.FindControl("ErrorMsg"), System.Web.UI.WebControls.Panel)

            lblErrorMsg.Text = ""
            ErrorMsg.Visible = False
        Catch ex As Exception

        End Try
    End Sub


#End Region
    Public Shared Sub showErrMsg(ByRef page As System.Web.UI.Page, ByVal obj As Object)
        Try

            'If UIUtility.isTabPostBack(page) Then
            '    Dim pageTab As PageTab = CType(page.FindControl("PageTab"), PageTab)
            '    If TypeOf obj Is Exception Then
            '        pageTab.setException(CType(obj, Exception))
            '    Else
            '        pageTab.setException(New Exception(CType(obj, String)))
            '    End If
            'End If

            Dim lblErrorMsg As Label = CType(page.FindControl("lblErrorMsg"), System.Web.UI.WebControls.Label)
            Dim palErrorMsg As Panel = CType(page.FindControl("ErrorMsg"), System.Web.UI.WebControls.Panel)

            Dim sErrorMsg As String = String.Empty

            If TypeOf obj Is Exception Then
                Dim exception As Exception = CType(obj, Exception)
                Dim sb As New System.Text.StringBuilder

                If com.Azion.EloanUtility.UIUtility.isDebug Then
                    Do
                        If ValidateUtility.isValidateData(exception.Message) Then
                            sb.Append(exception.Message.ToString & "。<br>")
                        End If

                        If ValidateUtility.isValidateData(exception.StackTrace) Then
                            sb.Append(exception.StackTrace.ToString & "。<br>")
                        End If

                        exception = exception.InnerException
                    Loop Until (exception Is Nothing)
                Else
                    sb.Append(exception.Message.ToString & "。<br>")
                End If
                sErrorMsg = sb.ToString
            Else
                sErrorMsg = CType(obj, String)
            End If

            If Not IsNothing(palErrorMsg) Then
                palErrorMsg.Visible = True
                lblErrorMsg.Visible = True
                lblErrorMsg.Text = sErrorMsg
            Else
                sErrorMsg = "<span style='color:red'>" & sErrorMsg & "</span>"

                HttpContext.Current.Response.Write(sErrorMsg)
            End If
        Catch ex As Exception
            Dim s As String = CType(obj, Exception).Message.ToString & "<br>" & CType(obj, Exception).StackTrace
        End Try
    End Sub

#End Region

#Region "JavaScript 顯示錯誤訊息&回到待辦清單"
    'Shared Sub JSChglocation(ByVal page As System.Web.UI.Page, ByVal sPath As String)
    '    Dim js As String
    '    js = "<script Language='JAVAScript'>" & vbCrLf
    '    js += "window.location='" & sPath & "'"
    '    js += "</" & "script>" & vbCrLf
    '    page.Response.Write(js)
    'End Sub

    Shared Sub goMainPage(ByRef page As System.Web.UI.Page, ByVal sMsg As String)
        Dim js As String
        js = "<script Language='JAVAScript'>" & vbCrLf
        If com.Azion.EloanUtility.ValidateUtility.isValidateData(sMsg) Then
            js += "alert('" & sMsg & "');"
        End If
        js += "window.location.replace('" & UIUtility.getRootPath() & "/ss_mainpage.aspx?ftype=1')"
        js += "</" & "script>" & vbCrLf
        page.Response.Write(js)
        'page.RegisterStartupScript()
         
        'com.Azion.EloanUtility.UIUtility.Redirect(UIUtility.getRootPath() & "/ss_mainpage.aspx?ftype=1")
    End Sub

#Region "JavaScript 顯示錯誤訊息&回到待辦清單"

    Shared Sub alert(ByVal sMesg As String)
        Dim sResult As String = String.Empty
        Dim sbJS As New System.Text.StringBuilder
        Try
            '不管傳入字串有無單引號，都將它加上單引號
            If ValidateUtility.isValidateData(sMesg) Then
                sResult = "'" & sMesg & "'"
                sResult = sResult.Replace("''", "'")
            End If
            sbJS.Append("<script language='javascript'>").Append(vbNewLine)
            sbJS.Append("alert(").Append(sResult).Append(");").Append(vbNewLine)
            sbJS.Append("</script>").Append(vbNewLine)
            HttpContext.Current.Response.Write(sbJS.ToString)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region
#End Region

#Region "UI唯讀模式 mDISPLAYMODE"
    Shared Function isReadOnly() As Boolean
        Dim httpContext As System.Web.HttpContext = System.Web.HttpContext.Current
        If Not IsNothing(httpContext) Then

            If Not IsNothing(httpContext.Request.QueryString("WatchStepNo")) Then
                Return True
            End If

            If Not IsNothing(httpContext.Request.QueryString("MDISPLAYMODE")) Then
                Dim str() As String = httpContext.Request.QueryString("MDISPLAYMODE").Split(CType(",", Char))
                Dim s As String = str(str.Length - 1)
                '取得最後一個MDISPLAYMODE的參數
                If ValidateUtility.isValidateData(s) Then
                    If UCase(s) = "TRUE" Then
                        Return True
                    End If
                End If
            End If
        End If

        Return False
    End Function
#End Region

#Region "開關視窗"

    Public Shared Sub closeWindow(ByVal sMsg As String)
        Dim sJscript As String
        sJscript = "<Script Language='JAVAScript'>" & vbCrLf
        sJscript &= "alert('" & sMsg & "');"
        sJscript &= "window.self.close();" & vbCrLf
        sJscript &= "</Script>" & vbCrLf
        HttpContext.Current.Response.Write(sJscript)
    End Sub

    '關閉視窗
    Shared Sub closeWindow(ByRef page As System.Web.UI.Page)
        Dim sJS As String = "<script language='javascript'>window.self.close();</script>"
        page.Response.Write(sJS)
    End Sub

    Shared Sub closeWindow(ByRef page As System.Web.UI.Page, ByVal sMsg As String)
        Dim sJS As String
        sJS = "<script language='javascript'>  alert('" & sMsg & "'); window.self.close();</script>"
        page.Response.Write(sJS)
    End Sub

    '回傳值後關閉視窗 (限 dialogWidth)
    Shared Sub returnValueAndcloseWindow(ByVal page As System.Web.UI.Page, ByVal sReturnValue As String)
        Dim sJS As String = "<script Language='JAVAScript'>"
        sJS &= "var temp=new Array();"
        sJS &= "temp[0] =" & sReturnValue & " ;"
        sJS &= "parent.window.returnValue=temp;"
        sJS &= "window.self.close();"
        sJS &= "</script>"
        page.Response.Write(sJS)
    End Sub

    '跳出視窗
    Shared Sub showModalDialog(ByVal sUrl As String, Optional ByVal _Height_Width As String = "", Optional ByVal _Property As String = "")
        'HttpContext.Current.Response.Write(showModalDialogJS(sUrl, _Height_Width, _Property))
        showOpen(sUrl, _Height_Width, _Property)
    End Sub

    Private Shared Sub genSessionInfo(ByRef sUrl As String)
        'Dim sessionKeys As System.Collections.Specialized.NameObjectCollectionBase.KeysCollection = System.Web.HttpContext.Current.Session.Keys

        'Dim sSessionInfo As String = String.Empty
        'For Each sessionKey As String In sessionKeys
        '    If TypeOf System.Web.HttpContext.Current.Session.Item(sessionKey) Is String Then
        '        sSessionInfo = sSessionInfo & sessionKey & "=" & System.Web.HttpContext.Current.Session.Item(sessionKey).ToString & "|"
        '    End If
        'Next
        'sSessionInfo = sSessionInfo
        'sSessionInfo = "&sessionInfo=" & HttpUtility.UrlEncode(sSessionInfo) & "&SID=" & System.Web.HttpContext.Current.Session.SessionID

        'If sUrl.EndsWith(".aspx") Then
        '    sUrl = sUrl & "?"
        'End If
        'sUrl = sUrl & sSessionInfo
    End Sub

    Public Shared Function showModalDialogAddAtt(ByVal sUrl As String, Optional ByVal _Height_Width As String = "", Optional ByVal _Property As String = "", Optional ByVal _InEvent As Boolean = False) As String
        Return com.Azion.EloanUtility.UIUtility.showModalDialogJS(sUrl).Replace("<script lang='JavaScript'>", "").Replace("</script>", "")
    End Function

    '產生Open Window的JavaScript
    Private Shared Function showModalDialogJS(ByVal sUrl As String, Optional ByVal _Height_Width As String = "", Optional ByVal _Property As String = "", Optional ByVal _InEvent As Boolean = False) As String
        Dim sbJS As New System.Text.StringBuilder
        Dim sSize() As String = Nothing
        Dim sHeight As String
        Dim sWidth As String
        Dim sProperty As String = "'dialogWidth:{{W}}px; dialogHeight:{{H}}px; center:yes;scroll:1;status:0;help:0;resizable:0'"
        genSessionInfo(sUrl)
        If ValidateUtility.isValidateData(_Height_Width) Then sSize = _Height_Width.ToLower.Split(CChar("x"))

        If _InEvent Then
            sbJS.Append("JavaScript:")
        Else
            sbJS.Append("<script lang='JavaScript'>").Append(vbNewLine)
            'sbJS.Append(" var args = new Object; ").Append(vbNewLine)
            'sbJS.Append(" args.window = window; ").Append(vbNewLine)
        End If

        sbJS.Append("window.showModalDialog('" & sUrl & "',self,{{P}});").Append(vbNewLine)

        If Not _InEvent Then
            sbJS.Append("</script>").Append(vbNewLine)
        End If

        If Not IsNothing(sSize) AndAlso sSize.Length = 2 AndAlso IsNumeric(sSize(0)) AndAlso IsNumeric(sSize(1)) Then
            sHeight = sSize(0)
            sWidth = sSize(1)
        Else '沒設定長寬，就放到最大
            sHeight = "' + window.screen.height + '"
            sWidth = "' + window.screen.width + '"
        End If

        If ValidateUtility.isValidateData(_Property) Then
            '不管_Property有沒有加單引號、雙引號，一律刪掉重加
            sProperty = "'" & _Property.Replace("'", "").Replace("""", "") & "'"
            sbJS.Replace("{{P}}", sProperty)
        Else
            sbJS.Replace("{{P}}", sProperty.Replace("{{W}}", sWidth).Replace("{{H}}", sHeight))
        End If
        Return sbJS.ToString
    End Function

    '跳出視窗
    Public Shared Sub showOpen(ByVal sUrl As String, Optional ByVal _Height_Width As String = "", Optional ByVal _Property As String = "")
        HttpContext.Current.Response.Write(openJS(sUrl, _Height_Width, _Property))
    End Sub

    '產生Open Window的JavaScript
    Private Shared Function openJS(ByVal sUrl As String, Optional ByVal _Height_Width As String = "", Optional ByVal _Property As String = "", Optional ByVal _InEvent As Boolean = False) As String
        Dim sbJS As New System.Text.StringBuilder
        Dim sSize() As String = Nothing
        Dim sHeight As String
        Dim sWidth As String
        Dim sProperty As String = "'height={{H}}, width={{W}}, top=0, left=0, toolbar=yes, menubar=yes, scrollbars=yes, resizable=no,location=n o, status=no'"

        If ValidateUtility.isValidateData(_Height_Width) Then sSize = _Height_Width.ToLower.Split(CChar("x"))

        If _InEvent Then
            sbJS.Append("JavaScript:")
        Else
            sbJS.Append("<script lang='JavaScript'>").Append(vbNewLine)
        End If

        sbJS.Append("window.open('" & sUrl & "',null,{{P}});").Append(vbNewLine)

        If Not _InEvent Then
            sbJS.Append("</script>").Append(vbNewLine)
        End If

        If Not IsNothing(sSize) AndAlso sSize.Length = 2 AndAlso IsNumeric(sSize(0)) AndAlso IsNumeric(sSize(1)) Then
            sHeight = sSize(0)
            sWidth = sSize(1)
        Else '沒設定長寬，就放到最大
            sHeight = "' + window.screen.height + '"
            sWidth = "' + window.screen.width + '"
        End If

        If ValidateUtility.isValidateData(_Property) Then
            '不管_Property有沒有加單引號、雙引號，一律刪掉重加
            ' sProperty = "'" & _Property.Trim(CChar("'")).Replace("""", "") & "'"
            sProperty = String.Format("'{0}'", _Property.Trim(CChar("'")).Replace("""", ""))
            sbJS.Replace("{{P}}", sProperty)
        Else
            sbJS.Replace("{{P}}", sProperty.Replace("{{W}}", sWidth).Replace("{{H}}", sHeight))
        End If
        Return sbJS.ToString
    End Function
#End Region

#Region "是否為頁簽PostBack"

    Public Shared Function isTabPostBack(ByRef page As System.Web.UI.Page) As Boolean

        If Not IsNothing(page.Request) Then
            Dim sPostBackObj As String

            If IsNothing(page.Request.Params("__EVENTTARGET")) Then
                sPostBackObj = "Nothing"
            Else
                sPostBackObj = page.Request.Params("__EVENTTARGET")
            End If
            If sPostBackObj.ToUpper.IndexOf("PAGETAB") = -1 Then
                Return False
            End If
        End If

        If page.IsPostBack Then
            Return True
        Else
            Return False
        End If
    End Function
#End Region

#End Region

End Class
