<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="MBQry_NEWs_01_v01.aspx.vb" Inherits="MBSC.MBQry_NEWs_01_v01" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title>內容明細</title>
    <!-- #include virtual="~/inc/MBSCCSS.inc" -->
    <!-- #include virtual="~/inc/MBSCJS.inc" --> 
</head>
<body>
    <form id="form1" runat="server">
        <!--顯示上方外框線-->
        <!-- #include virtual="~/inc/MBSCTableStart.inc" -->
        <!--錯誤訊息區-->
        <!-- #include virtual="~/inc/MBSCErrorMsg.inc" -->
        <table style="border-collapse:collapse;border-style:solid;empty-cells:show;border-width:1px;border-color:black;background:transparent;" cellspacing="0" cellpadding="10px" align="center" width="1050px" >
            <tr>
                <td style="text-align: left;" >
                    <asp:Literal ID="LTL_CNTHTML" runat="server" />
                </td>
            </tr>
        </table>
        <!--顯示下方外框線-->
        <!--#include virtual="~/inc/MBSCTableEnd.inc"--> 
    </form>
</body>
</html>
