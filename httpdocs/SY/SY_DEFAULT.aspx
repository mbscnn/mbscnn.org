<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SY_DEFAULT.aspx.vb" Inherits="MBSC.SY_DEFAULT" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>EnTie Eloan �{�ɵn�J�e��</title>
    <meta http-equiv="Content-Type" content="text/html; charset=big5">
    <link href="<%=com.Azion.EloanUtility.UIUtility.getRootPath()%>/css/Style.css" rel="stylesheet"
        type="text/css">
    <link href="<%=com.Azion.EloanUtility.UIUtility.getRootPath()%>/css/default.css"
        type="text/css" rel="STYLESHEET">
    <link href="<%=com.Azion.EloanUtility.UIUtility.getRootPath()%>/css/std.css" type="text/css"
        rel="StyleSheet">
    <script type="text/javascript" language="JavaScript">
    </script>
</head>
<body bgcolor="#ffffff" text="#000000" leftmargin="0" topmargin="0" marginwidth="0"
    marginheight="0">
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
        <font color="red" size="4">�п�J�n�J��T<br />
        </font>
        <p>
            <form id="f1" name="f1" runat="server">
            &nbsp;
            <table name="logintab" width="30%" cellpadding="0" cellspacing="0" border="1">
                <tr>
                    <td class='th1' width="10%" nowrap>
                        �ϥΪ̭��u�s��
                    </td>
                    <td class="td1">
                        <asp:TextBox ID="txtStaffId" runat="server" MaxLength='5' onkeyup="changefacus()"
                            CssClass="newinput"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td class='th1' width="10%" nowrap>
                        �t�ΰT��
                    </td>
                    <td class="td1">
                        <font color="red">&nbsp;<span id="outError" runat="server" /></font>
                    </td>
                </tr>
            </table>
            <br />
            <div style="text-align: center">
                <asp:Button runat="server" name="btnLogin" ID="btnLogin" Text="�n�J�t��" CssClass="bt" />
            </div>
            </form>
        </p>
    </center>
    <script type="text/javascript" language="JavaScript">

        // �Y��J�����׬�6,Button��o�J�I
        function changefacus() {
            if (document.getElementById("txtStaffId").value.length == 8) {
                document.getElementById("btnLogin").focus();
            }
        }

        document.getElementById("txtStaffId").focus();
    </script>
    <p>
        <font face="�s�ө���"></font>
    </p>
</body>
</html>
