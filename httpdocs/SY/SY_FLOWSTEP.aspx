<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SY_FLOWSTEP.aspx.vb" Inherits="MBSC.SY_FLOWSTEP" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>流程步驟</title>
    <base target="_self" />
    <!--SY CSS-->
    <!-- #include virtual="~/inc/SYCSS.inc" -->
</head>
<body>
    <form id="form1" runat="server">
    <!--顯示上方外框線-->
    <!-- #include virtual="~/inc/eLoanTableStart.inc" -->
    <table class="mtr" cellspacing="1" cellpadding="2" width="100%">
        <tr>
            <td class="mtr" align="center">
                流程案件狀態
            </td>
        </tr>
        <tr>
            <td>
                <table class="mtr" cellspacing="0" cellpadding="0" width="100%" border="0">
                    <tr>
                        <td>
                            <asp:DataGrid ID="dgFlowCaseStatus" runat="server" CssClass="crtable" Width="100%"
                                BorderWidth="1px" CellPadding="0" BorderColor="#77B6E3" AutoGenerateColumns="False"
                                HeaderStyle-ForeColor="#000000" AllowSorting="True">
                                <Columns>
                                    <asp:TemplateColumn HeaderText="步驟名稱">
                                        <ItemTemplate>
                                            <asp:Label ID="lblFlowName" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"STEPNAME")%>'></asp:Label>
                                        </ItemTemplate>
                                        <ItemStyle Width="20%" CssClass="td2c_b"></ItemStyle>
                                        <HeaderStyle CssClass="td2c_b" />
                                    </asp:TemplateColumn>
                                    <asp:TemplateColumn HeaderText="用戶端">
                                        <ItemTemplate>
                                            <asp:HiddenField id="hid_CLIENT" runat="server" Value='<%#DataBinder.Eval(Container.DataItem,"CLIENT")%>' /> 
                                            <asp:HiddenField id="hid_SENDER" runat="server" Value='<%#DataBinder.Eval(Container.DataItem,"SENDER")%>' /> 
                                            <asp:Label ID="lblUserName" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"USERNAME")%>'></asp:Label>
                                        </ItemTemplate>
                                        <ItemStyle Width="15%" CssClass="td2c_b"></ItemStyle>
                                        <HeaderStyle CssClass="td2c_b" />
                                    </asp:TemplateColumn>
                                     <asp:TemplateColumn HeaderText="代理人">
                                        <ItemTemplate>
                                            <asp:Label ID="lblUserName_AGENT" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"USERNAME_AGENT")%>'></asp:Label>
                                        </ItemTemplate>
                                        <ItemStyle Width="15%" CssClass="td2c_b"></ItemStyle>
                                        <HeaderStyle CssClass="td2c_b" />
                                    </asp:TemplateColumn>
                                    <asp:TemplateColumn HeaderText="聯絡電話">
                                        <ItemTemplate>
                                            <asp:Label ID="lblOfficeTel" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"OFFICETEL")%>'></asp:Label>
                                        </ItemTemplate>
                                        <ItemStyle Width="20%" CssClass="td2c_b"></ItemStyle>
                                        <HeaderStyle CssClass="td2c_b" />
                                    </asp:TemplateColumn>
                                    <asp:TemplateColumn HeaderText="收件時間">
                                        <ItemTemplate>
                                            <asp:Label ID="lblStartTime" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"STARTTIME")%>'></asp:Label>
                                        </ItemTemplate>
                                        <ItemStyle Width="20%" CssClass="td2c_b"></ItemStyle>
                                        <HeaderStyle CssClass="td2c_b" />
                                    </asp:TemplateColumn>
                                    <asp:TemplateColumn HeaderText="案件狀態">
                                        <ItemTemplate>
                                            <asp:Label ID="lblStatus" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"STATUS")%>'></asp:Label>
                                        </ItemTemplate>
                                        <ItemStyle Width="10%" Wrap="False" CssClass="td2c_b"></ItemStyle>
                                        <HeaderStyle CssClass="td2c_b" />
                                    </asp:TemplateColumn>
                                </Columns>
                            </asp:DataGrid>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    <!--顯示下方外框線-->
    <!--#include virtual="~/inc/eLoanTableEnd.inc"-->
    </form>
</body>
</html>
