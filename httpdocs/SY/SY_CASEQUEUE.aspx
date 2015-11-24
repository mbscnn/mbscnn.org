<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SY_CASEQUEUE.aspx.vb"
    Inherits="MBSC.SY_CASEQUEUE" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" >
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>佇列取件</title>
    <base target="_self" />
    <!--SY CSS-->
    <!-- #include virtual="~/inc/SYCSS.inc" -->
</head>
<body>
    <form id="form1" runat="server">
    <table width="100%" border="1px" cellspacing="0" cellpadding="0" class="mtr">
        <tr>
            <td>
                <asp:DataGrid ID="dgCase" runat="server" CssClass="CRTABLE" Width="100%" BorderWidth="1px"
                    CellPadding="0" BorderColor="#77B6E3" AutoGenerateColumns="False" HeaderStyle-ForeColor="#000000"
                    AllowSorting="True">
                    <AlternatingItemStyle CssClass="DataGrid_AlternatingItemStyle"></AlternatingItemStyle>
                    <HeaderStyle ForeColor="Black"></HeaderStyle>
                    <Columns>
                        <asp:TemplateColumn HeaderText="">
                            <HeaderStyle CssClass="th1c_b" Font-Size="9pt" Wrap="False" Height="22px" Width="9%" />
                            <ItemStyle Wrap="True" CssClass="td2c_b"></ItemStyle>
                            <ItemTemplate>
                                <asp:CheckBox ID="chkFlag" runat="server" />
                            </ItemTemplate>
                        </asp:TemplateColumn>
                        <asp:TemplateColumn HeaderText="案件編號">
                            <HeaderStyle CssClass="th1c_b" Font-Size="9pt" Wrap="False" Height="22px" Width="13%" />
                            <ItemStyle Wrap="True" CssClass="td2c_b"></ItemStyle>
                            <ItemTemplate>
                                <asp:Label ID="lblCaseId" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"CASEID")%>'>
                                </asp:Label>
                            </ItemTemplate>
                        </asp:TemplateColumn>
                        <asp:TemplateColumn HeaderText="核准編號">
                            <HeaderStyle CssClass="th1c_b" Font-Size="9pt" Wrap="False" Height="22px" Width="13%" />
                            <ItemStyle Wrap="True" CssClass="td2_b"></ItemStyle>
                            <ItemTemplate>
                                <asp:Label ID="lblApvCaseId" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"APV_CAS_ID")%>'>
                                </asp:Label>
                            </ItemTemplate>
                        </asp:TemplateColumn>
                        <asp:TemplateColumn HeaderText="業務類別">
                            <HeaderStyle CssClass="th1c_b" Font-Size="9pt" Wrap="False" Height="22px" Width="13%" />
                            <ItemStyle Wrap="True" CssClass="td2c_b"></ItemStyle>
                            <ItemTemplate>
                                <asp:Label ID="lblFlowCName" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"FLOW_CNAME")%>'>
                                </asp:Label>
                            </ItemTemplate>
                        </asp:TemplateColumn>
                        <asp:TemplateColumn HeaderText="授信戶名稱">
                            <HeaderStyle CssClass="th1c_b" Font-Size="9pt" Wrap="False" Height="22px" Width="13%" />
                            <ItemStyle Wrap="True" CssClass="td2c_b"></ItemStyle>
                            <ItemTemplate>
                                <asp:Label ID="lblCplAplId" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"CPL_APL_ID")%>'>
                                </asp:Label>
                                <asp:Label ID="lblCplAplNam" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"CPL_APL_NAM")%>'>
                                </asp:Label>
                            </ItemTemplate>
                        </asp:TemplateColumn>
                        <asp:TemplateColumn HeaderText="流程步驟">
                            <HeaderStyle CssClass="th1c_b" Font-Size="9pt" Wrap="False" Height="22px" Width="13%" />
                            <ItemStyle Wrap="True" CssClass="td2c_b"></ItemStyle>
                            <ItemTemplate>
                                <asp:Label ID="lblStepName" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"STEP_NAME")%>'>
                                </asp:Label>
                            </ItemTemplate>
                        </asp:TemplateColumn>
                        <asp:TemplateColumn HeaderText="申請時間">
                            <HeaderStyle CssClass="th1c_b" Font-Size="9pt" Wrap="False" Height="22px" Width="13%" />
                            <ItemStyle Wrap="True" CssClass="td2c_b"></ItemStyle>
                            <ItemTemplate>
                                <asp:Label ID="lblStartTime" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"STARTTIME")%>'>
                                </asp:Label>
                            </ItemTemplate>
                        </asp:TemplateColumn>
                        <asp:TemplateColumn HeaderText="申請單位">
                            <HeaderStyle CssClass="th1c_b" Font-Size="9pt" Wrap="False" Height="22px" Width="13%" />
                            <ItemStyle Wrap="True" CssClass="td2c_b"></ItemStyle>
                            <ItemTemplate>
                                <asp:Label ID="lblBrcName" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"BRCNAME")%>'>
                                </asp:Label>
                            </ItemTemplate>
                        </asp:TemplateColumn>
                    </Columns>
                </asp:DataGrid>
            </td>
        </tr>
    </table>
    <!--顯示下方外框線-->
    <!--#include virtual="~/inc/eLoantableEnd.inc"-->
    <table class="mtr" cellspacing="0" cellpadding="0" width="100%" border="0">
        <tr>
            <td align="left">
                <asp:Button ID="btnGetCase" runat="server" Text="取得案件" CssClass="bt" />
                <asp:Button ID="btnCloseWin" runat="server" Text="關閉" CssClass="bt" />
            </td>
        </tr>
    </table>
    </form>
</body>
</html>
