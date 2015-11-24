<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SY_UPDSYSUSER.aspx.vb"
    Inherits="MBSC.SY_UPDSYSUSER" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>應用系統人員分派</title>
    <base target="_self" />
    <!--SY CSS-->
    <!-- #include virtual="~/inc/SYCSS.inc" -->
    <!--共用的SY javascript-->
    <!-- #include virtual="~/inc/SYJS.inc" -->
</head>
<body>
    <form id="form1" runat="server">
    <div class="divhelp">
        <img alt="help" src="<%= com.Azion.EloanUtility.UIUtility.getImgPath()+"help.png"%>"
            onclick="javascript:window.open('./help/SY_UPDSYSUSER.docx')" />
    </div>
    <br />
    <div class="TaskTtitle">
        應用系統人員分派</div>
    <br />
    <!--顯示上方外框線-->
    <!-- #include virtual="~/inc/eLoanTableStart.inc" -->
    <div id="divEdit" runat="server">
        <table id="Table1" class="mtr" border="0" cellspacing="1" cellpadding="2" width="100%"
            runat="server">
            <tr>
                <td class="th1" style="width: 25%">
                    維護角色
                </td>
                <td style="width: 85%; text-align: right">
                    <div style="float: left">
                        <asp:RadioButtonList ID="rdolistRole" runat="server" RepeatDirection="Horizontal"
                            TabIndex="5" RepeatLayout="Flow" AutoPostBack ="true" >
                        </asp:RadioButtonList>
                    </div>
                    <div style="float: right">
                        <asp:Button ID="btnSend" runat="server" Text="確認送出" CssClass="bt" TabIndex="10" />
                        <asp:Button ID="btnCancel" runat="server" Text="全部取消" CssClass="bt" TabIndex="15" />
                    </div>
                </td>
            </tr>
            <tr>
                <td class="th1" style="width: 25%">
                    單位
                </td>
                <td style="width: 85%">
                    <asp:DropDownList ID="ddlDepart" runat="server" Width="200px" TabIndex="20" AutoPostBack="True">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="th1" style="width: 25%">
                    使用者
                </td>
                <td style="width: 85%">
                    <asp:DropDownList ID="ddlUser" runat="server" Width="200px" TabIndex="25">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td colspan="2" style="text-align: center">
                    <asp:Button ID="btnAdd" runat="server" Text="新增" CssClass="bt" TabIndex="30" />
                </td>
            </tr>
            <tr id="trDetail" runat="server">
                <td colspan="2">
                    <asp:DataGrid ID="dgAddDetail" runat="server" Width="100%" BorderWidth="1px" CellPadding="0"
                        BorderColor="#77B6E3" AutoGenerateColumns="False" HeaderStyle-ForeColor="#000000"
                        AllowSorting="True">
                        <Columns>
                            <asp:TemplateColumn HeaderText="員工編號">
                                <HeaderStyle CssClass="th1c_b" HorizontalAlign="Center" Font-Size="9pt" Wrap="False"
                                    Height="22px" Width="15%" />
                                <ItemStyle Wrap="True" CssClass="td2_b"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Label ID="lblSTAFFID" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"STAFFID")%>'>
                                    </asp:Label>
                                    <asp:HiddenField ID="hidROLEID" Value='<%#DataBinder.Eval(Container.DataItem,"ROLEID")%>'
                                        runat="server"></asp:HiddenField>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="姓名">
                                <HeaderStyle CssClass="th1c_b" HorizontalAlign="Center" Font-Size="9pt" Wrap="False"
                                    Height="22px" Width="20%" />
                                <ItemStyle Wrap="True" CssClass="td2_b"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Label ID="lblUSERNAME" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"USERNAME")%>'>
                                    </asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="單位">
                                <HeaderStyle CssClass="th1c_b" HorizontalAlign="Center" Font-Size="9pt" Wrap="False"
                                    Height="22px" Width="30%" />
                                <ItemStyle Wrap="True" CssClass="td2_b"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Label ID="lblBRANAME" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"BRANAME")%>'>
                                    </asp:Label>
                                    <asp:HiddenField ID="hidBRA_DEPNO" runat="server" Value='<%#DataBinder.Eval(Container.DataItem,"BRA_DEPNO")%>'>
                                    </asp:HiddenField>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="角色">
                                <HeaderStyle CssClass="th1c_b" HorizontalAlign="Center" Font-Size="9pt" Wrap="False"
                                    Height="22px" Width="25%" />
                                <ItemStyle Wrap="True" CssClass="td2_b"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Label ID="lblROLENAME" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"ROLENAME")%>'>
                                    </asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="刪除">
                                <HeaderStyle CssClass="th1c_b" HorizontalAlign="Center" Font-Size="9pt" Wrap="False"
                                    Height="22px" Width="10%" />
                                <ItemStyle Wrap="True" CssClass="td2_b"></ItemStyle>
                                <ItemTemplate>
                                    <asp:ImageButton onmousedown="if(window.confirm('確定要刪除資料嗎?')==true){hidOutAction.value='D';this.click();}"
                                        ID="ImgDel" CommandArgument='<%#Container.DataItem("ROLEID") %>' ImageUrl="~/img/imgDelete.gif"
                                        runat="server"></asp:ImageButton>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                        </Columns>
                    </asp:DataGrid>
                </td>
            </tr>
        </table>
    </div>
    <div id="divChecked" runat="server">
         <table width="100%" runat="server">
            <tr align="right">
                <td colspan="2">
                    案件編號:<asp:Label ID="lblCaseId" runat="server"></asp:Label>
                </td>
            </tr>
            <tr align="right">
                <td colspan="2">
                    申請單位:
                    <asp:Label ID="lblBranch" runat="server"></asp:Label>
                </td>
            </tr>
        </table>
        角色分派
        <table id="tblChecked" class="mtr" border="0" cellspacing="1" cellpadding="2" width="100%"
            runat="server">
            <tr>
                <td style="width: 40%; text-align: center; height: 300px">
                    <span style="float: left">修改前</span><br />
                    <asp:ListBox ID="lstPreEdit" runat="server" Width="260px" Height="300px" Enabled="false">
                    </asp:ListBox>
                </td>
                <td style="width: 20%; text-align: center">
                    <asp:Label ID="lblMsg" runat="server"></asp:Label>
                </td>
                <td style="width: 40%; text-align: center; height: 300px">
                    <span style="float: left">修改後</span><br />
                    <asp:ListBox ID="lstAfterEdit" runat="server" Width="260px" Height="300px" Enabled="false">
                    </asp:ListBox>
                    <br />
                    <span style="text-align: center">紅字為本次新增</span>
                </td>
            </tr>
            <tr>
                <td colspan="3">
                </td>
            </tr>
            <tr id="trUserGrant" runat="server">
                <td colspan="3">
                    <asp:DataGrid ID="dgRoleDetail" runat="server" Width="100%" BorderWidth="1px" CellPadding="0"
                        BorderColor="#77B6E3" AutoGenerateColumns="False" HeaderStyle-ForeColor="#000000"
                        AllowSorting="True">
                        <Columns>
                            <asp:TemplateColumn HeaderText="角色名稱">
                                <HeaderStyle CssClass="th1c_b" HorizontalAlign="Center" Font-Size="9pt" Wrap="False"
                                    Height="22px" Width="60%" />
                                <ItemStyle Wrap="True" CssClass="td2c_b" HorizontalAlign="Center"></ItemStyle>
                                <ItemTemplate>
                                    <asp:LinkButton ID="lnkBtnRoleName" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"ROLENAME")%>'></asp:LinkButton>
                                    <asp:HiddenField ID="hidRoleID" runat="server" Value='<%#DataBinder.Eval(Container.DataItem,"ROLEID")%>' />
                                </ItemTemplate>
                            </asp:TemplateColumn>
                        </Columns>
                    </asp:DataGrid>
                </td>
            </tr>
        </table>
        <div id="divBtnChecked" runat="server" style="text-align: center">
            <asp:Button ID="btnAgree" runat="server" Text="同意" CssClass="bt" />
            <asp:Button ID="btnDiffer" runat="server" Text="不同意" CssClass="bt" />
            <asp:Button ID="btnUpdate" runat="server" Text="修正補充" CssClass="bt" />
        </div>
    </div>
    <asp:HiddenField ID="hidOutAction" runat="server" />
    <!--顯示下方外框線-->
    <!--#include virtual="~/inc/eLoanTableEnd.inc"-->
    </form>
</body>
</html>
