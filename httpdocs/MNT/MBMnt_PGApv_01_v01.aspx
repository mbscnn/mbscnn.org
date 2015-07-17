<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="MBMnt_PGApv_01_v01.aspx.vb" Inherits="MBSC.MBMnt_PGApv_01_v01" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title>功能權限維護</title>
    <!-- #include virtual="~/inc/MBSCCSS.inc" -->
    <!-- #include virtual="~/inc/MBSCJS.inc" --> 
</head>
<body topmargin="0">
    <form id="form1" runat="server">
    <!-- #include virtual="~/inc/PageTab.inc" -->
<table width="100%" cellspacing="0" cellpadding="0" style="width:1235px;background:transparent;margin-left: auto; margin-right: auto;" align="center">
    <tr>
        <td style="vertical-align:top;padding:0;text-align:left;background:transparent;width:185px" >
            <!-- #include virtual="~/inc/vTab.inc" -->
        </td>
        <td style="vertical-align:top;padding:0;text-align:left;background:transparent;width:1050px" >
            <!-- #include virtual="~/inc/Signin.inc" -->
            <div align="center">
                <h2 style="color: #800080; font-family: 標楷體">
                    功能權限維護
                </h2>
            </div>
            <!--顯示上方外框線-->
            <!-- #include virtual="~/inc/MBSCTableStart.inc" -->
            <!--錯誤訊息區-->
            <!-- #include virtual="~/inc/MBSCErrorMsg.inc" -->
            <asp:PlaceHolder ID="PLH_LIST" runat="server" >
            <table class="CRTable_Top" width="100%" cellspacing="0">
                <tr>
                    <td width="100%" colspan="3" class="th1c_b">
                        選取維護人員
                    </td>
                </tr>
                <tr>
                    <td width="45%" class="th1c_b">
                        待維護人員
                    </td>
                    <td width="10%" class="th1c_b">
                        動作
                    </td>
                    <td width="45%" class="th1c_b">
                        欲維護人員
                    </td>
                </tr>
                <tr>
                    <td width="45%" class="td2c_b">
                        <asp:ListBox ID="LB_L_MDUSER" runat="server" Rows="10" Width="100%" SelectionMode="Multiple"  />
                    </td>
                    <td width="10%" class="td2c_b">
                        <asp:Button ID="btnCHOOSE" runat="server" CssClass="bt" Text="選入" />
                        <BR />
                        <asp:Button ID="btnOUT" runat="server" CssClass="bt" Text="選出" />
                        <BR />
                        <asp:Button ID="btnCHOOSE_ALL" runat="server" CssClass="bt" Text="全部選入" />
                        <BR />
                        <asp:Button ID="btnOUT_ALL" runat="server" CssClass="bt" Text="全部選出" />
                    </td>
                    <td width="45%" class="td2c_b">
                        <asp:ListBox ID="LB_C_MDUSER" runat="server" Rows="10" Width="100%" SelectionMode="Multiple" />
                    </td>
                </tr>
            </table>
            <div align="center">
                <asp:Button ID="btnQRY" runat="server" CssClass="bt" Text="確定" />
            </div>
            </asp:PlaceHolder>
            <asp:PlaceHolder ID="PLH_PROG" runat="server" Visible="false">
                <table class="CRTable_Top" width="100%" cellspacing="0">
                    <asp:Repeater ID="RP_PROG" runat="server" >
                        <ItemTemplate>
                            <tr>
                                <td width="100%" class="th1_b" >
                                    維護人員帳號：<asp:Literal ID="LTL_ACCT" runat="server" Text='<%#Container.DataItem("ACCT")%>' />
                                </td>
                            </tr>
                            <tr>
                                <td width="100%" class="th1_b">
                                    功能清單
                                    <asp:Button ID="btnALL" runat="server" CssClass="bt" Text="全選" CommandName="ALL" />
                                    <asp:Button ID="btnCANCELALL" runat="server" CssClass="bt" Text="全部取消" CommandName="CANCELALL" />
                                </td>
                            </tr>
                            <tr>
                                <td width="100%" class="td2_b" >
                                    <asp:CheckBoxList ID="CB_PROG" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow" />
                                </td>
                            </tr>
                        </ItemTemplate>
                    </asp:Repeater>
                </table>
                <div align="center">
                    <asp:Button ID="btnSAVE" runat="server" CssClass="bt" Text="儲存" />
                    &nbsp;
                    <asp:Button ID="btnBack" runat="server" CssClass="bt" Text="回選取維護人員清單" />
                </div>
            </asp:PlaceHolder>

            <!--顯示下方外框線-->
            <!--#include virtual="~/inc/MBSCTableEnd.inc"--> 
        </td>
    </tr>
</table>
    </form>
</body>
</html>
