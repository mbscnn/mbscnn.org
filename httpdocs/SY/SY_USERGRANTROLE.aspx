<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SY_USERGRANTROLE.aspx.vb"
    Inherits="MBSC.SY_USERGRANTROLE" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" >
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>一般人員分派</title>
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
            onclick="javascript:window.open('./help/SY_USERGRANTROLE.docx')" />
    </div>
    <br />
    <div class="TaskTtitle">
        人員分派</div>
    <br />
    <!--顯示上方外框線-->
    <!-- #include virtual="~/inc/eLoanTableStart.inc" -->
    <asp:HiddenField runat="server" ID="hidOutAction"></asp:HiddenField>
    <div runat="server" id="divSend" style=" float :right ">
        <asp:Button ID="btnSendFlow" runat="server" Text="確認送出" CssClass="bt" />
        &nbsp;
        <asp:Button ID="btnAllCancel" runat="server" Text="全部取消" CssClass="bt" />
    </div>
    <div id="divQuery" runat="server">
        <table id="tblQuery" class="mtr" border="0" cellspacing="1" cellpadding="2" width="100%"
            runat="server">
            <tr>
                <td class="th1" style="width: 15%">
                    設定方式
                </td>
                <td style="width: 85%">
                    <asp:RadioButtonList ID="rdolistSetStyle" runat="server" RepeatDirection="Horizontal"
                        TabIndex="5" RepeatLayout="Flow" AutoPostBack="true">
                        <asp:ListItem Value="2">依角色</asp:ListItem>
                        <asp:ListItem Value="1">依人員</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <td class="th1" style="width: 15%">
                    單位
                </td>
                <td style="width: 85%">
                    <asp:DropDownList ID="ddlDepart" runat="server" TabIndex="10" AutoPostBack="true">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr runat="server" id="trRole"  style="width: 100%">
                <td class="th1" style="width: 15%">
                    角色名稱
                </td>
                <td style="width: 85%">
                    <asp:DropDownList ID="ddlRole" runat="server" TabIndex="15" AutoPostBack="true">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr runat="server" id="trUser">
                <td class="th1" style="width: 15%">
                    人員名稱
                </td>
                <td style="width: 85%">
                    <asp:DropDownList ID="ddlUser" runat="server" TabIndex="20" AutoPostBack="true">
                    </asp:DropDownList>
                </td>
            </tr>
        </table>
    </div>
    <div id="divEdit" runat="server">
        <table id="tbEdit" class="mtr" border="0" cellspacing="1" cellpadding="2" width="100%"
            runat="server">
            <tr>
                <td>
                    <table id="tblOperate" class="mtr" border="0" cellspacing="1" cellpadding="2" width="100%"
                        runat="server">
                        <tr>
                            <td style="width: 40%; text-align: center; height: 300px">
                                <asp:ListBox ID="lstBoxLeft" runat="server" SelectionMode="Multiple" Width="250px"
                                    Height="300px"></asp:ListBox>
                            </td>
                            <td style="width: 20%; text-align: center">
                                <asp:Button ID="btnIn" runat="server" Text="選入" CssClass="bt" />
                                <br />
                                <br />
                                <br />
                                <asp:Button ID="btnOut" runat="server" Text="選出" CssClass="bt" />
                            </td>
                            <td style="width: 40%; text-align: center; height: 300px">
                                <asp:ListBox ID="lstBoxRight" runat="server" Width="200px" Height="300px" SelectionMode="Multiple">
                                </asp:ListBox>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td>
                    <div style="text-align: center">
                        <asp:Button ID="btnSave" runat="server" Text="儲存" CssClass="bt" />
                        &nbsp;&nbsp;&nbsp
                        <asp:Button ID="btnCancel" runat="server" Text="取消" CssClass="bt" />
                    </div>
                </td>
            </tr>
            <tr runat="server" id="trByRoleDetail">
                <td>
                    <asp:DataGrid ID="dgByRoleDetail" runat="server" Width="100%" BorderWidth="1px" CellPadding="0"
                        BorderColor="#77B6E3" AutoGenerateColumns="False" HeaderStyle-ForeColor="#000000"
                        AllowSorting="True">
                        <Columns>
                            <asp:TemplateColumn HeaderText="單位">
                                <HeaderStyle CssClass="th1c_b" HorizontalAlign="Center" Font-Size="9pt" Wrap="False"
                                    Height="22px" Width="15%" />
                                <ItemStyle Wrap="True" CssClass="td2_b"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Label ID="lblDepart" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"DEPART")%>'>
                                    </asp:Label>
                                    <asp:HiddenField ID="hidBraDepno" runat="server" Value='<%#DataBinder.Eval(Container.DataItem,"BRA_DEPNO")%>' />
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="角色名稱">
                                <HeaderStyle CssClass="th1c_b" HorizontalAlign="Center" Font-Size="9pt" Wrap="False"
                                    Height="22px" Width="25%" />
                                <ItemStyle Wrap="True" CssClass="td2_b"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Label ID="lblRoleName" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"ROLENAME")%>'>
                                    </asp:Label>
                                    <asp:HiddenField ID="hilRoleID" runat="server" Value='<%#DataBinder.Eval(Container.DataItem,"ROLEID")%>' />
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="人員名稱">
                                <HeaderStyle CssClass="th1c_b" HorizontalAlign="Center" Font-Size="9pt" Wrap="False"
                                    Height="22px" Width="40%" />
                                <ItemStyle Wrap="True" CssClass="td2_b"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Label ID="lblUserName" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"USERNAME")%>'>
                                    </asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="修改">
                                <HeaderStyle CssClass="th1c_b" Font-Size="9pt" Wrap="False" Height="22px" Width="10%"
                                    HorizontalAlign="Center" />
                                <ItemStyle Wrap="True" CssClass="td2c_b" HorizontalAlign="Center"></ItemStyle>
                                <ItemTemplate>
                                    <asp:ImageButton ID="ImgEdit" CommandArgument='<%#Container.DataItem("ROLEID") %>'
                                        ImageUrl="~/img/imgEdit.gif" runat="server" TabIndex="105" CommandName="Edit">
                                    </asp:ImageButton>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="刪除">
                                <HeaderStyle CssClass="th1c_b" HorizontalAlign="Center" Font-Size="9pt" Wrap="False"
                                    Height="22px" Width="10%" />
                                <ItemStyle Wrap="True" CssClass="td2c_b" HorizontalAlign="Center"></ItemStyle>
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
            <tr runat="server" id="trByPeopleDetail">
                <td>
                    <asp:DataGrid ID="dgByPeopleDetail" runat="server" Width="100%" BorderWidth="1px"
                        CellPadding="0" BorderColor="#77B6E3" AutoGenerateColumns="False" HeaderStyle-ForeColor="#000000"
                        AllowSorting="True">
                        <Columns>
                            <asp:TemplateColumn HeaderText="單位">
                                <HeaderStyle CssClass="th1c_b" HorizontalAlign="Center" Font-Size="9pt" Wrap="False"
                                    Height="22px" Width="15%" />
                                <ItemStyle Wrap="True" CssClass="td2_b"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Label ID="lblDepart" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"DEPART")%>'>
                                    </asp:Label>
                                    <asp:HiddenField ID="hidBraDepno" runat="server" Value='<%#DataBinder.Eval(Container.DataItem,"BRA_DEPNO")%>' />
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="人員名稱">
                                <HeaderStyle CssClass="th1c_b" HorizontalAlign="Center" Font-Size="9pt" Wrap="False"
                                    Height="22px" Width="25%" />
                                <ItemStyle Wrap="True" CssClass="td2_b"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Label ID="lblUserName" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"USERNAME")%>'>
                                    </asp:Label>
                                    <asp:HiddenField ID="hidUserID" runat="server" Value='<%#DataBinder.Eval(Container.DataItem,"STAFFID")%>' />
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="角色名稱">
                                <HeaderStyle CssClass="th1c_b" HorizontalAlign="Center" Font-Size="9pt" Wrap="False"
                                    Height="22px" Width="40%" />
                                <ItemStyle Wrap="True" CssClass="td2_b"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Label ID="lblRoleName" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"ROLENAME")%>'>
                                    </asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="修改">
                                <HeaderStyle CssClass="th1c_b" Font-Size="9pt" Wrap="False" Height="22px" Width="10%"
                                    HorizontalAlign="Center" />
                                <ItemStyle Wrap="True" CssClass="td2c_b" HorizontalAlign="Center"></ItemStyle>
                                <ItemTemplate>
                                    <asp:ImageButton ID="ImgEdit" CommandArgument='<%#Container.DataItem("STAFFID") %>'
                                        ImageUrl="~/img/imgEdit.gif" runat="server" TabIndex="105" CommandName="Edit">
                                    </asp:ImageButton>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="刪除">
                                <HeaderStyle CssClass="th1c_b" HorizontalAlign="Center" Font-Size="9pt" Wrap="False"
                                    Height="22px" Width="10%" />
                                <ItemStyle Wrap="True" CssClass="td2c_b" HorizontalAlign="Center"></ItemStyle>
                                <ItemTemplate>
                                    <asp:ImageButton onmousedown="if(window.confirm('確定要刪除資料嗎?')==true){hidOutAction.value='D';this.click();}"
                                        ID="ImgDel" CommandArgument='<%#Container.DataItem("STAFFID") %>' ImageUrl="~/img/imgDelete.gif"
                                        runat="server"></asp:ImageButton>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                        </Columns>
                    </asp:DataGrid>
                </td>
            </tr>
            <tr runat="server" id="trAuthor">
                <td>
                    <div id="divReviewer" runat="server">
                    維護主管二 <asp:DropDownList ID="ddlReviewer" runat="server"></asp:DropDownList>
                    </div>
                </td>
            </tr>
        </table>
    </div>
    <div id="divChecked" runat="server">
        <table id="Table1" width="100%" runat="server">
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
        <asp:Label ID="lblGrantStyle" runat="server"></asp:Label>
        <table id="tblChecked" class="mtr" border="0" cellspacing="1" cellpadding="2" width="100%"
            runat="server">
            <tr>
                <td style="width: 40%; text-align: center; height: 300px">
                    <span style="float: left">修改前</span><br />
                    <asp:ListBox ID="lstPreEdit" runat="server" Width="200px" Height="300px" Enabled="false">
                    </asp:ListBox>
                    <br />
                    <span style="text-align: center">紅字為本次刪除</span>
                </td>
                <td style="width: 20%; text-align: center">
                    <asp:Label ID="lblMsg" runat="server"></asp:Label>
                </td>
                <td style="width: 40%; text-align: center; height: 300px">
                    <span style="float: left">修改後</span><br />
                    <asp:ListBox ID="lstAfterEdit" runat="server" Width="200px" Height="300px" Enabled="false">
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
                    <asp:DataGrid ID="dgUserGrant" runat="server" Width="100%" BorderWidth="1px" CellPadding="0"
                        BorderColor="#77B6E3" AutoGenerateColumns="False" HeaderStyle-ForeColor="#000000"
                        AllowSorting="True">
                        <Columns>
                            <asp:TemplateColumn HeaderText="單位">
                                <HeaderStyle CssClass="th1c_b" HorizontalAlign="Center" Font-Size="9pt" Wrap="False"
                                    Height="22px" Width="40%" />
                                <ItemStyle Wrap="True" CssClass="td2_b"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Label ID="lblDepart" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"BRANAME")%>'>
                                    </asp:Label>
                                    <asp:HiddenField ID="hidBraDepno" runat="server" Value='<%#DataBinder.Eval(Container.DataItem,"DEPNO")%>' />
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="人員名稱">
                                <HeaderStyle CssClass="th1c_b" HorizontalAlign="Center" Font-Size="9pt" Wrap="False"
                                    Height="22px" Width="60%" />
                                <ItemStyle Wrap="True" CssClass="td2c_b" HorizontalAlign="Center"></ItemStyle>
                                <ItemTemplate>
                                    <asp:LinkButton ID="lnkBtnUserName" runat="server">
                                        <asp:Label ID="UserName" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"NAME")%>'>
                                        </asp:Label>
                                    </asp:LinkButton>
                                    <asp:HiddenField ID="hidBtnUserID" runat="server" Value='<%#DataBinder.Eval(Container.DataItem,"ID")%>' />
                                </ItemTemplate>
                            </asp:TemplateColumn>
                        </Columns>
                    </asp:DataGrid>
                </td>
            </tr>
            <tr id="trRoleGrant" runat="server">
                <td colspan="3">
                    <asp:DataGrid ID="dgRoleGrant" runat="server" Width="100%" BorderWidth="1px" CellPadding="0"
                        BorderColor="#77B6E3" AutoGenerateColumns="False" HeaderStyle-ForeColor="#000000"
                        AllowSorting="True">
                        <Columns>
                            <asp:TemplateColumn HeaderText="單位">
                                <HeaderStyle CssClass="th1c_b" HorizontalAlign="Center" Font-Size="9pt" Wrap="False"
                                    Height="22px" Width="40%" />
                                <ItemStyle Wrap="True" CssClass="td2_b"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Label ID="lblDepart" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"BRANAME")%>'>
                                    </asp:Label>
                                    <asp:HiddenField ID="hidBraDepno" runat="server" Value='<%#DataBinder.Eval(Container.DataItem,"DEPNO")%>' />
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="角色名稱">
                                <HeaderStyle CssClass="th1c_b" HorizontalAlign="Center" Font-Size="9pt" Wrap="False"
                                    Height="22px" Width="60%" />
                                <ItemStyle Wrap="True" CssClass="td2c_b" HorizontalAlign="Center"></ItemStyle>
                                <ItemTemplate>
                                    <asp:LinkButton ID="lnkBtnRoleName" runat="server">
                                        <asp:Label ID="lblRoleName" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"NAME")%>'>
                                        </asp:Label>
                                    </asp:LinkButton>
                                    <asp:HiddenField ID="hidRoleID" runat="server" Value='<%#DataBinder.Eval(Container.DataItem,"ID")%>' />
                                </ItemTemplate>
                            </asp:TemplateColumn>
                        </Columns>
                    </asp:DataGrid>
                </td>
            </tr>
        </table>
        <div id="divBtnChecked" runat="server">
            <asp:Button ID="btnAgree" runat="server" Text="同意" CssClass="bt" />
            <asp:Button ID="btnDiffer" runat="server" Text="不同意" CssClass="bt" />
            <asp:Button ID="btnReviseFlow" runat="server" Text="修正補充" CssClass="bt" />
        </div>
    </div>
    <div>
        <asp:HiddenField ID="hidOldXml" runat="server" />
    </div>
    <!--顯示下方外框線-->
    <!--#include virtual="~/inc/eLoanTableEnd.inc"-->
    </form>
</body>
</html>
