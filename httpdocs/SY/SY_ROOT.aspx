<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SY_ROOT.aspx.vb" Inherits="MBSC.SY_ROOT" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head runat="server">
    <title>系統管理員登錄</title>
    <!--SY CSS-->
    <!-- #include virtual="~/inc/SYCSS.inc" -->
    <!--共用的SY javascript-->
    <!-- #include virtual="~/inc/SYJS.inc" -->
    <script type="text/javascript" language="javascript">
        // 驗證密碼是否為空
        function validate() {
            var txtPassword1 = document.getElementById("txtPWD1");
            var txtPassword2 = document.getElementById("txtPWD2");
            var errorControl = null;
            var errorMsg = "";

            // 如果沒有輸入密碼
            if (txtPassword1.value == "") {
                if (errorControl == null) {
                    errorControl = txtPassword1;
                    txtPassword1.focus();
                }
                errorMsg += "請輸入第一組密碼！\r\n"
            }

            if (txtPassword2.value == "" ) {
                if (errorControl == null) {
                    errorControl = txtPassword2;
                    txtPassword2.focus();
                }
                errorMsg += "請輸入第二組密碼！\r\n"
            }

            // 如果存在錯誤控件,就提示錯誤信息
            if (errorControl != null) {
                errorControl.focus();

                alert(errorMsg);

                return false;
            }

            return true;
        }


        // 点击复选框时触发事件
        function postBackByObject() {
            var o = window.event.srcElement;
            if (o.tagName == "INPUT" && o.type == "checkbox") {
                __doPostBack("", "");
            }
        }

    </script>
    <style type="text/css">
        .newinput
        {
            margin-bottom: 0px;
        }
    </style>
</head>
<body bgcolor="#ffffff" text="#000000" leftmargin="0" topmargin="0" marginwidth="0"
    marginheight="0">
    <form id="form1" runat="server">
    <div runat="server" id="divloginPassWord">
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr background="<%=com.Azion.EloanUtility.UIUtility.getImgPath()%>top_t01.gif">
                <td background="<%=com.Azion.EloanUtility.UIUtility.getImgPath()%>top_t01.gif">
                    <img src="<%=com.Azion.EloanUtility.UIUtility.getImgPath()%>top_tital.gif" width="209"
                        height="45">
                </td>
            </tr>
        </table>
        <center>
            <span id="userip" runat="server" />
            <br />
            <br />
            <br />
            <br />
            <br />
            <br />
            <div style="text-align: center; font-family: 標楷體; font-size: 17; font-weight: bold;
                color: #800080;">
                請輸入登入資訊</div>
            <br />
            <div style="text-align: center">
                <table name="logintab" width="30%" cellpadding="0" cellspacing="0" border="1">
                    <tr>
                        <td class='th1' width="10%" nowrap>
                            &nbsp;第一組密碼&nbsp;
                        </td>
                        <td class="td1">
                            <asp:TextBox ID="txtPWD1" runat="server" TabIndex="10" TextMode="Password" MaxLength="12"></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <td class='th1' width="10%" nowrap>
                            &nbsp;第二組密碼&nbsp;
                        </td>
                        <td class="td1">
                            <asp:TextBox ID="txtPWD2" runat="server" TabIndex="10" TextMode="Password" MaxLength="12"></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <td class='th1' width="10%" nowrap>
                            &nbsp;系統訊息
                        </td>
                        <td class="td1">
                            <asp:Label runat="server" ID="lblSystemMsg" ForeColor="#CC3300"></asp:Label>
                        </td>
                    </tr>
                </table>
                <br />
                <div style="text-align: center">
                    <asp:Button runat="server" name="btnChangePwd" ID="btnChangePwd" TabIndex="15" Text="變更密碼"
                        CssClass="bt" Width="80px" />
                    &nbsp;
                    <asp:Button runat="server" name="btnLogin" ID="btnLogin" TabIndex="15" Text="登入系統"
                        CssClass="bt" OnClientClick="return validate();" Width="80px" />
                    &nbsp;&nbsp;<asp:Button runat="server" name="btnCancel" ID="btnCancel" TabIndex="15"
                        Text="取消" CssClass="bt" Width="80px" />
                </div>
            </div>
        </center>
    </div>
    <div runat="server" id="divEditTwoPassWord">
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr background="<%=com.Azion.EloanUtility.UIUtility.getImgPath()%>top_t01.gif">
                <td background="<%=com.Azion.EloanUtility.UIUtility.getImgPath()%>top_t01.gif">
                    <img src="<%=com.Azion.EloanUtility.UIUtility.getImgPath()%>top_tital.gif" width="209"
                        height="45">
                </td>
            </tr>
        </table>
        <center>
            <span id="Span1" runat="server" />
            <br />
            <br />
            <br />
            <br />
            <br />
            <br />
            <div style="text-align: center; font-family: 標楷體; font-size: 17; font-weight: bold;
                color: #800080;">
                系統管理員分派</div>
            <br />
            <div style="text-align: center">
                <br />
                    <table id="tbData" class="mtr" cellspacing="1" cellpadding="1" runat="server">
                        <tr id="trSubSys" runat="server">
                            <td align="left" valign="top">
                                <asp:TreeView ID="treeViewAdministrators" runat="server" TabIndex="15" ShowCheckBoxes="All">
                                </asp:TreeView>
                            </td>
                        </tr>
                    </table>
                <br />
                <div style="text-align: center">
                    <asp:Button runat="server" ID="btnEidtTwoOK" Text="儲存" TabIndex="40" CssClass="bt" 
                        Width="80px" /> &nbsp;&nbsp;
                    <asp:Button runat="server" name="btnCancel2" ID="Cancel2" TabIndex="15" Text="取消"
                        CssClass="bt" Width="80px" />
                </div>
            </div>
        </center>
    </div>
    <div runat="server" id="divEditThridPassWord">
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr background="<%=com.Azion.EloanUtility.UIUtility.getImgPath()%>top_t01.gif">
                <td background="<%=com.Azion.EloanUtility.UIUtility.getImgPath()%>top_t01.gif">
                    <img src="<%=com.Azion.EloanUtility.UIUtility.getImgPath()%>top_tital.gif" width="209"
                        height="45">
                </td>
            </tr>
        </table>
        <center>
            <span id="Span2" runat="server" />
            <br />
            <br />
            <br />
            <br />
            <br />
            <br />
            <div style="text-align: center; font-family: 標楷體; font-size: 17; font-weight: bold;
                color: #800080;">
                密碼變更</div>
            <br />
            <br />
            <div style="text-align: center; font-family: 標楷體; font-size: 15; font-weight: bold;">
                <asp:RadioButton ID="RbChangePwd1" runat="server" Text="變更第一組密碼" GroupName="GroupChangePwd" />
                &nbsp;&nbsp;
                <asp:RadioButton ID="RbChangePwd2" runat="server" Text="變更第二組密碼" GroupName="GroupChangePwd" />
            </div>
            <br />
            <div style="text-align: center">
                <table name="logintab" width="30%" cellpadding="0" cellspacing="0" border="1">
                    <tr>
                        <td class='th1' width="10%" nowrap>
                            舊密碼
                        </td>
                        <td class="td1">
                            <asp:TextBox ID="txtEditThridOPWD" runat="server" TabIndex="50" TextMode="Password"
                                MaxLength="12"></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <td class='th1' width="10%" nowrap>
                            新密碼
                        </td>
                        <td class="td1">
                            <asp:TextBox ID="txtEditThridNewPWD" runat="server" TabIndex="55" TextMode="Password"
                                MaxLength="12"></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <td class='th1' width="10%" nowrap>
                            確認新密碼
                        </td>
                        <td class="td1">
                            <asp:TextBox ID="txtEditThridCheckPWD" runat="server" TabIndex="60" TextMode="Password"
                                MaxLength="12"></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <td class='th1' width="10%" nowrap>
                            系統訊息
                        </td>
                        <td class="td1">
                            <asp:Label runat="server" ID="lblEditThirdSystemMsg" ForeColor="#CC3300"></asp:Label>
                        </td>
                    </tr>
                </table>
                <br />
                <div style="text-align: center">
                    
                    <asp:Button runat="server" ID="btnEditThrid" Text="儲存" TabIndex="70" CssClass="bt"
                        Width="80px" /> &nbsp;&nbsp;
                    <asp:Button runat="server" name="btnCancel3" ID="btnCance3" TabIndex="15" Text="取消"
                        CssClass="bt" Width="80px" />
                </div>
            </div>
        </center>
    </div>
    <asp:HiddenField ID="hidScript" runat="server"></asp:HiddenField>
    </form>
</body>
</html>
