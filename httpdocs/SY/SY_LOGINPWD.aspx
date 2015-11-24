<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SY_LOGINPWD.aspx.vb" Inherits="MBSC.SY_LOGINPWD" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head runat="server">
    <title>外部使用者登入畫面</title>
    <!--SY CSS-->
    <!-- #include virtual="~/inc/SYCSS.inc" -->
    <!--共用的SY javascript-->
    <!-- #include virtual="~/inc/SYJS.inc" -->
    <script type="text/javascript" language="javascript">
        // 驗證使用者編號，密碼是否為空
        function validate() {
            var txtStaffid = document.getElementById("txtStaffId");
            var txtPassword = document.getElementById("txtPWD");
            var errorControl = null;
            var errorMsg = "";

            // 如果沒有輸入使用者員工編號
            if (txtStaffid.value == "") {
                if (errorControl == null) {
                    errorControl = txtStaffid;
                    txtStaffid.focus();
                }
                errorMsg += "請輸入使用者員工編號！\r\n"
            }

            // 如果沒有輸入密碼
            if (txtPassword.value == "") {
                if (errorControl == null) {
                    errorControl = txtPassword;
                    txtPassword.focus();
                }
                errorMsg += "請輸入密碼！\r\n"
            }

            // 如果存在錯誤控件,就提示錯誤信息
            if (errorControl != null) {
                errorControl.focus();

                alert(errorMsg);

                return false;
            }

            return true;
        }

        // 驗證Mail是否輸入，格式是否正確
        function validateMail() 
        {
            return true;
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
             <div style =" text-align :center; font-family: 標楷體; font-size: 17; font-weight: bold;color: #800080;">
    請輸入登入資訊</div>
            <font color="red" size="4"><br />
            </font>
            <p style="text-align: center">
                <table name="logintab" width="30%" cellpadding="0" cellspacing="0" border="1">
                    <tr>
                        <td class='th1' width="10%" nowrap>
                            使用者員工編號
                        </td>
                        <td class="td1">
                            <asp:textbox id="txtStaffId" runat="server" maxlength='5' cssclass="newinput" tabindex="5"></asp:textbox>
                        </td>
                    </tr>
                    <tr>
                        <td class='th1' width="10%" nowrap>
                            密碼
                        </td>
                        <td class="td1">
                            <asp:textbox id="txtPWD" runat="server" tabindex="10" textmode="Password" 
                                MaxLength="12"></asp:textbox>
                        </td>
                    </tr>
                    <tr>
                        <td class='th1' width="10%" nowrap>
                            系統訊息
                        </td>
                        <td class="td1">
                            <asp:label runat="server" id="lblSystemMsg" forecolor="#CC3300"></asp:label>
                        </td>
                    </tr>
                </table>
                <br />
                <div style="text-align: center">
                    <asp:button runat="server" name="btnLogin" id="btnLogin" tabindex="15" text="登入系統" cssclass="bt"
                        onclientclick="return validate();" />
                </div>
            </p>
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
             <div style =" text-align :center; font-family: 標楷體; font-size: 17; font-weight: bold;color: #800080;">
    密碼修改</div>
            </font>
            <p style="text-align: center">
                <table name="logintab" width="30%" cellpadding="0" cellspacing="0" border="1">
                    <tr>
                        <td class='th1' width="10%" align="left" nowrap>
                            使用者員工編號
                        </td>
                        <td class="td1">
                            <asp:textbox id="txtEditTwoStaffId" runat="server" maxlength='5' cssclass="newinput"
                                tabindex="20"></asp:textbox>
                        </td>
                    </tr>
                    <tr>
                        <td class='th1' width="10%" nowrap>
                            舊密碼
                        </td>
                        <td class="td1">
                            <asp:textbox id="txtEditTwoOPWD" runat="server" tabindex="25" 
                                textmode="Password" MaxLength="12"></asp:textbox>
                        </td>
                    </tr>
                    <tr>
                        <td class='th1' width="10%" nowrap>
                            新密碼
                        </td>
                        <td class="td1">
                            <asp:textbox id="txtEditTwoNewPWD" runat="server" tabindex="30" 
                                textmode="Password" MaxLength="12"></asp:textbox>
                        </td>
                    </tr>
                    <tr>
                        <td class='th1' width="10%" nowrap>
                            確認新密碼
                        </td>
                        <td class="td1">
                            <asp:textbox id="txtEditTwoCheckPWD" runat="server" tabindex="35" 
                                textmode="Password" MaxLength="12"></asp:textbox>
                        </td>
                    </tr>
                    <tr>
                        <td class='th1' width="10%" nowrap>
                            系統訊息
                        </td>
                        <td class="td1">
                            <asp:label runat="server" id="lblEditTwoSystemMsg" forecolor="#CC3300"></asp:label>
                        </td>
                    </tr>
                </table>
                <br />
                <div style="text-align: center">
                    <asp:button runat="server" id="btnEidtTwoOK" text="儲存" tabindex="40" cssclass="bt" />
                </div>
            </p>
        </center>
    </div>
    <div runat="server" id="divEditThridPassWord">
        <center>
            <span id="Span2" runat="server" />
            <br />
            <br />
            <br />
            <br />
            <br />
            <br />
                  <div style =" text-align :center; font-family: 標楷體; font-size: 17; font-weight: bold;color: #800080;">
    外部鑑價單位變更</div>
            </font>
            <p style="text-align: center">
                <table name="logintab" width="30%" cellpadding="0" cellspacing="0" border="1">
                    <tr>
                        <td class='th1' width="10%" nowrap>
                            使用者員工編號
                        </td>
                        <td class="td1">
                            <asp:textbox id="txtEditThridStaffId" runat="server" maxlength='5' cssclass="newinput"
                                tabindex="45" AutoPostBack = "true" ></asp:textbox>
                        </td>
                    </tr>
                    <tr>
                        <td class='th1' width="10%" nowrap>
                            舊密碼
                        </td>
                        <td class="td1">
                            <asp:textbox id="txtEditThridOPWD" runat="server" tabindex="50" 
                                textmode="Password" MaxLength="12"></asp:textbox>
                        </td>
                    </tr>
                    <tr>
                        <td class='th1' width="10%" nowrap>
                            新密碼
                        </td>
                        <td class="td1">
                            <asp:textbox id="txtEditThridNewPWD" runat="server" tabindex="55" 
                                textmode="Password" MaxLength="12"></asp:textbox>
                        </td>
                    </tr>
                    <tr>
                        <td class='th1' width="10%" nowrap>
                            確認新密碼
                        </td>
                        <td class="td1">
                            <asp:textbox id="txtEditThridCheckPWD" runat="server" tabindex="60" 
                                textmode="Password" MaxLength="12"></asp:textbox>
                        </td>
                    </tr>
                    <tr>
                        <td class='th1' width="10%" nowrap>
                            Mail
                        </td>
                        <td class="td1">
                            <asp:textbox id="txtMail" runat="server" tabindex="65"></asp:textbox>
                        </td>
                    </tr>
                </table>
                <br />
                <div style="text-align: center">
                    <asp:button runat="server" id="btnEditThrid" text="儲存" tabindex="70" cssclass="bt" onclientclick="return validateMail();" />
                </div>
            </p>
        </center>
    </div>
    <asp:hiddenfield id="hidScript" runat="server"></asp:hiddenfield>
    </form>
</body>
</html>
