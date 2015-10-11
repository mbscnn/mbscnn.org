<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="MBQry_Member_01_v01.aspx.vb" Inherits="MBSC.MBQry_Member_01_v01" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head runat="server">
    <title></title>
    <!-- #include virtual="~/inc/MBSCCSS.inc" -->
    <!-- #include virtual="~/inc/MBSCJS.inc" -->
    <script lang="JavaScript" src="<%=com.Azion.EloanUtility.UIUtility.getRootPathClient() + "/js/jquery-1.7.1.min.js"%>"></script>
    <base target="_self" />
</head>
<body>
    <form id="form1" runat="server">
    <!-- #include virtual="~/inc/PageTab.inc" -->
<table width="100%" cellspacing="0" cellpadding="0" style="width:1235px;background:transparent;margin-left: auto; margin-right: auto;" align="center">
    <tr>
        <td style="vertical-align:top;padding:0;text-align:left;background:transparent;width:1050px" >
            <!-- #include virtual="~/inc/Signin.inc" -->
            <div align="center">
                <h2 style="color: #800080; font-family: 標楷體">
                    MBSC課程報名人員查詢
                </h2>
            </div>
            <!--顯示上方外框線-->
            <!-- #include virtual="~/inc/MBSCTableStart.inc" -->

            <!--錯誤訊息區-->
            <!-- #include virtual="~/inc/MBSCErrorMsg.inc" -->
            <table id="tbl_QRY" runat="server" class="CRTable_Top" width="100%" cellspacing="0">
                <tr>
                    <td width="15%" class="th1c_b">
                      課程編號
                    </td>
                    <td width="85%" class="td2_b">
                        <asp:TextBox ID="txt_MB_SEQ" runat="server" CssClass="bordertxt" width="60%" AutoPostBack="true" />
                    </td>
                </tr>
                <tr>
                    <td width="15%" class="th1c_b">
                      梯次
                    </td>
                    <td width="85%" class="td2_b">
                        <asp:DropDownList ID="MB_BATCH" runat="server" />
                    </td>
                </tr>
                <tr>
                    <td class="th1c_b" align="center" colspan="2">
                        <asp:Button ID="btn_Confirm" runat="server" Text="查詢" CssClass="bt" />
                    </td>
                </tr>
            </table>
            <div id="divResult" runat="server" style="display:none">
            <table id="tblTitle" runat="server" width="100%" border="1" cellpadding="0" cellspacing="0" BorderColor="#77B6E3" >
		        <tr><td align="center" colspan="3"><H3 style="COLOR: #800080; FONT-FAMILY: 標楷體"><asp:Label ID="lbl_CLASSNAME" Runat="server" /><asp:Label ID="lbl_BATCH" Runat="server" />報名人員表</H3></td></tr>
	        </table>
            <asp:DataGrid ID="dg_Member" runat="server" CssClass="CRTABLE" Width="100%" ItemStyle-BorderWidth="1px" HeaderStyle-BorderWidth="1px"
                CellPadding="0" BorderColor="#77B6E3" AutoGenerateColumns="False" ShowHeader="True" HeaderStyle-ForeColor="#000000" GridLines="Both" >
                    <Columns>
                        <asp:BoundColumn HeaderText="報名人員" HeaderStyle-CssClass="th1c_b" HeaderStyle-Width="20%"  HeaderStyle-Wrap="False"  ItemStyle-Wrap="True" ItemStyle-HorizontalAlign="Center"></asp:BoundColumn>
                        <asp:BoundColumn HeaderText="簽到" HeaderStyle-CssClass="th1c_b" HeaderStyle-Width="20%"  HeaderStyle-Wrap="False"  ItemStyle-Wrap="True" ItemStyle-HorizontalAlign="Center"></asp:BoundColumn>
                        <asp:BoundColumn HeaderText="課程備註" HeaderStyle-CssClass="th1c_b" HeaderStyle-Width="40%"  HeaderStyle-Wrap="False"  ItemStyle-Wrap="True" ItemStyle-HorizontalAlign="Center"></asp:BoundColumn>  
                        <asp:BoundColumn HeaderText="正取備取" HeaderStyle-CssClass="th1c_b" HeaderStyle-Width="20%"  HeaderStyle-Wrap="False"  ItemStyle-Wrap="True" ItemStyle-HorizontalAlign="Center"></asp:BoundColumn>
                    </Columns>
                </asp:DataGrid>
                </div>
                <div align="center" class="th1c_b">
                    <asp:Button ID="btn_QSEQ" runat="server" Text="另存EXCEL檔" CssClass="bt" Visible="false" />
                </div>
            <!--顯示下方外框線-->
            <!--#include virtual="~/inc/MBSCTableEnd.inc"--> 
        </td>
    </tr>
</table>
    </form>
</body>
</html>
