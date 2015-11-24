<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SY_USERINFO.aspx.vb" Inherits="MBSC.SY_USERINFO" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>個人資訊</title>
    <base target="_self" />
    <!--SY CSS-->
    <!-- #include virtual="~/inc/SYCSS.inc" -->
    <!--共用的SY javascript-->
    <!-- #include virtual="~/inc/SYJS.inc" -->
    <script type="text/javascript">

        // E-Mail格式驗證
        function CheckValidate() {
            var $email = $("#txtEmail");
            var emailRule = new RegExp(/^\w+([-+.]\w+)*@\w+([-.]\\w+)*\.\w+([-.]\w+)*$/);

            if (!emailRule.test(jQuery.trim($email.val()))) {

                alert("郵件格式輸入有誤！");
                $email.val("");
                $email[0].focus();

                return false;
            }

            return true;
        }
    </script>
</head>
<body>
    <form id="form1" runat="server">
    <table class="mtr" cellspacing="1" cellpadding="2" width="100%">
        <tr>
            <td class="mtr" align="center">
                項目名稱
            </td>
            <td class="mtr" align="center">
                詳細內容
            </td>
        </tr>
        <tr>
            <td class="td2c">
                全名
            </td>
            <td align="center">
                <asp:Label ID="lblUserName" runat="server"></asp:Label>
            </td>
        </tr>
        <tr>
            <td class="td2c">
                登入名稱
            </td>
            <td align="center">
                <asp:Label ID="lblStaffId" runat="server"></asp:Label>
            </td>
        </tr>
        <tr>
            <td class="td2c">
                職稱
            </td>
            <td align="center">
                <asp:Label ID="lblJobTitle" runat="server"></asp:Label>
            </td>
        </tr>
        <tr>
            <td class="td2c">
                所屬部門
            </td>
            <td align="center">
                <asp:Label ID="lblBrcName" runat="server"></asp:Label>
            </td>
        </tr>
        <tr>
            <td class="td2c">
                電話號碼一
            </td>
            <td align="center">
                <asp:TextBox ID="txtOfficeTel1" MaxLength="20" TabIndex="5" runat="server"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <td class="td2c">
                E-Mail
            </td>
            <td align="center">
                <asp:TextBox ID="txtEmail" MaxLength="100" TabIndex="10" runat="server"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <td colspan="2" align="center">
                <asp:Button ID="btnSaveChange" runat="server" TabIndex="15" Text="儲存" OnClientClick="return CheckValidate();"
                    CssClass="bt" />
                <asp:Button ID="btnCloseWin" runat="server" TabIndex="20" Text="關閉" CssClass="bt" />
            </td>
        </tr>
    </table>
    </form>
</body>
</html>
