<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SY_USERINFO.aspx.vb" Inherits="MBSC.SY_USERINFO" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>�ӤH��T</title>
    <base target="_self" />
    <!--SY CSS-->
    <!-- #include virtual="~/inc/SYCSS.inc" -->
    <!--�@�Ϊ�SY javascript-->
    <!-- #include virtual="~/inc/SYJS.inc" -->
    <script type="text/javascript">

        // E-Mail�榡����
        function CheckValidate() {
            var $email = $("#txtEmail");
            var emailRule = new RegExp(/^\w+([-+.]\w+)*@\w+([-.]\\w+)*\.\w+([-.]\w+)*$/);

            if (!emailRule.test(jQuery.trim($email.val()))) {

                alert("�l��榡��J���~�I");
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
                ���ئW��
            </td>
            <td class="mtr" align="center">
                �ԲӤ��e
            </td>
        </tr>
        <tr>
            <td class="td2c">
                ���W
            </td>
            <td align="center">
                <asp:Label ID="lblUserName" runat="server"></asp:Label>
            </td>
        </tr>
        <tr>
            <td class="td2c">
                �n�J�W��
            </td>
            <td align="center">
                <asp:Label ID="lblStaffId" runat="server"></asp:Label>
            </td>
        </tr>
        <tr>
            <td class="td2c">
                ¾��
            </td>
            <td align="center">
                <asp:Label ID="lblJobTitle" runat="server"></asp:Label>
            </td>
        </tr>
        <tr>
            <td class="td2c">
                ���ݳ���
            </td>
            <td align="center">
                <asp:Label ID="lblBrcName" runat="server"></asp:Label>
            </td>
        </tr>
        <tr>
            <td class="td2c">
                �q�ܸ��X�@
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
                <asp:Button ID="btnSaveChange" runat="server" TabIndex="15" Text="�x�s" OnClientClick="return CheckValidate();"
                    CssClass="bt" />
                <asp:Button ID="btnCloseWin" runat="server" TabIndex="20" Text="����" CssClass="bt" />
            </td>
        </tr>
    </table>
    </form>
</body>
</html>
