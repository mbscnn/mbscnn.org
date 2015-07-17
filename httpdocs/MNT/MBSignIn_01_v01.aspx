<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="MBSignIn_01_v01.aspx.vb" Inherits="MBSC.MBSignIn_01_v01" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title>會員登入</title>
    <!-- #include virtual="~/inc/MBSCCSS.inc" -->
    <!-- #include virtual="~/inc/MBSCJS.inc" -->
    <script language="javascript" type="text/javascript">
                        $(document).ready(function () {
                            $('#txt_UserId').focus(function () {
                                if ($("#txt_UserId").val() == "帳號(e-Mail)") {
                                    $("#txt_UserId").css("color", "black");
                                    $("#txt_UserId").val("");
                                }
                            });
                            $('#txt_UserId').blur(function () {
                                if ($('#txt_UserId').val() == '') {
                                    $("#txt_UserId").css("color", "Gray");
                                    $("#txt_UserId").val("帳號(e-Mail)");
                                }
                            })

                            $('#password-clear').show();
                            $('#txt_Password').hide();

                            $('#password-clear').focus(function () {
                                $('#password-clear').hide();
                                $('#txt_Password').show();
                                $('#txt_Password').focus();
                            });
                            $('#txt_Password').blur(function () {
                                if ($('#txt_Password').val() == '') {
                                    $('#password-clear').show();
                                    $('#txt_Password').hide();
                                }
                            });
                        });
                    </script>
</head>
<body>
    <form id="form1" runat="server">
        <!-- #include virtual="~/inc/PageTab.inc" -->
        <!--錯誤訊息區-->
        <!-- #include virtual="~/inc/MBSCErrorMsg.inc" -->
        <table width="100%" cellspacing="0" cellpadding="0" style="width:1235px;background:transparent;margin-left: auto; margin-right: auto;" align="center">
            <tr>
                <td style="vertical-align:top;padding:0;text-align:left;background:transparent;width:185px" >
                    <!-- #include virtual="~/inc/vTab.inc" -->
                </td>
                <td style="vertical-align:top;padding:0;text-align:center;background:transparent;width:1050px" >
                    <!-- #include virtual="~/inc/Signin.inc" -->
                    <div style="clear: both;"></div>
                    <div>
                        <asp:TextBox ID="txt_UserId" runat="server" Width="175px" CssClass="mbsctxt" Text="帳號(e-Mail)" Style="color: Gray;height:30px;width:200px" />
                        <BR />
                        <input id="password-clear" type="text" value="密碼" autocomplete="off" size="20" class="mbsctxt" style="color: Gray;height:30px;width:200px" />
                        <asp:TextBox ID="txt_Password" runat="server" CssClass="mbsctxt" TextMode="Password" Columns="20" style="height:30px;width:200px" />
                        <BR />
                        <BR />
                        <asp:Button ID="btLogin" runat="server" Text="登入" CssClass="mbscbt" Style="height:30px;width:200px;font-size:12pt" />
                        <span></span>
                        <BR />
                        <BR />
                        <asp:Button ID="btSign" runat="server" Text="成為會員" CssClass="mbscbt" Style="height:30px;width:200px;font-size:12pt" />
                        <span></span>
                        <BR />
                        <BR />
                        <asp:Button ID="btnReMail" runat="server" Text="重發認證信" CssClass="mbscbt" Style="height:30px;width:200px;font-size:12pt" />
                    </div>
                    <div style="clear: both;"></div>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
