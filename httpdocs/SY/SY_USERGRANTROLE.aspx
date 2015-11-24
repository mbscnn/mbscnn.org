<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SY_USERGRANTROLE.aspx.vb"
    Inherits="MBSC.SY_USERGRANTROLE" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" >
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>�@��H������</title>
    <base target="_self" />
    <!--SY CSS-->
    <!-- #include virtual="~/inc/SYCSS.inc" -->
    <!--�@�Ϊ�SY javascript-->
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
        �H������</div>
    <br />
    <!--��ܤW��~�ؽu-->
    <!-- #include virtual="~/inc/eLoanTableStart.inc" -->
    <asp:HiddenField runat="server" ID="hidOutAction"></asp:HiddenField>
    <div runat="server" id="divSend" style=" float :right ">
        <asp:Button ID="btnSendFlow" runat="server" Text="�T�{�e�X" CssClass="bt" />
        &nbsp;
        <asp:Button ID="btnAllCancel" runat="server" Text="��������" CssClass="bt" />
    </div>
    <div id="divQuery" runat="server">
        <table id="tblQuery" class="mtr" border="0" cellspacing="1" cellpadding="2" width="100%"
            runat="server">
            <tr>
                <td class="th1" style="width: 15%">
                    �]�w�覡
                </td>
                <td style="width: 85%">
                    <asp:RadioButtonList ID="rdolistSetStyle" runat="server" RepeatDirection="Horizontal"
                        TabIndex="5" RepeatLayout="Flow" AutoPostBack="true">
                        <asp:ListItem Value="2">�̨���</asp:ListItem>
                        <asp:ListItem Value="1">�̤H��</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <td class="th1" style="width: 15%">
                    ���
                </td>
                <td style="width: 85%">
                    <asp:DropDownList ID="ddlDepart" runat="server" TabIndex="10" AutoPostBack="true">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr runat="server" id="trRole"  style="width: 100%">
                <td class="th1" style="width: 15%">
                    ����W��
                </td>
                <td style="width: 85%">
                    <asp:DropDownList ID="ddlRole" runat="server" TabIndex="15" AutoPostBack="true">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr runat="server" id="trUser">
                <td class="th1" style="width: 15%">
                    �H���W��
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
                                <asp:Button ID="btnIn" runat="server" Text="��J" CssClass="bt" />
                                <br />
                                <br />
                                <br />
                                <asp:Button ID="btnOut" runat="server" Text="��X" CssClass="bt" />
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
                        <asp:Button ID="btnSave" runat="server" Text="�x�s" CssClass="bt" />
                        &nbsp;&nbsp;&nbsp
                        <asp:Button ID="btnCancel" runat="server" Text="����" CssClass="bt" />
                    </div>
                </td>
            </tr>
            <tr runat="server" id="trByRoleDetail">
                <td>
                    <asp:DataGrid ID="dgByRoleDetail" runat="server" Width="100%" BorderWidth="1px" CellPadding="0"
                        BorderColor="#77B6E3" AutoGenerateColumns="False" HeaderStyle-ForeColor="#000000"
                        AllowSorting="True">
                        <Columns>
                            <asp:TemplateColumn HeaderText="���">
                                <HeaderStyle CssClass="th1c_b" HorizontalAlign="Center" Font-Size="9pt" Wrap="False"
                                    Height="22px" Width="15%" />
                                <ItemStyle Wrap="True" CssClass="td2_b"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Label ID="lblDepart" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"DEPART")%>'>
                                    </asp:Label>
                                    <asp:HiddenField ID="hidBraDepno" runat="server" Value='<%#DataBinder.Eval(Container.DataItem,"BRA_DEPNO")%>' />
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="����W��">
                                <HeaderStyle CssClass="th1c_b" HorizontalAlign="Center" Font-Size="9pt" Wrap="False"
                                    Height="22px" Width="25%" />
                                <ItemStyle Wrap="True" CssClass="td2_b"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Label ID="lblRoleName" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"ROLENAME")%>'>
                                    </asp:Label>
                                    <asp:HiddenField ID="hilRoleID" runat="server" Value='<%#DataBinder.Eval(Container.DataItem,"ROLEID")%>' />
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="�H���W��">
                                <HeaderStyle CssClass="th1c_b" HorizontalAlign="Center" Font-Size="9pt" Wrap="False"
                                    Height="22px" Width="40%" />
                                <ItemStyle Wrap="True" CssClass="td2_b"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Label ID="lblUserName" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"USERNAME")%>'>
                                    </asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="�ק�">
                                <HeaderStyle CssClass="th1c_b" Font-Size="9pt" Wrap="False" Height="22px" Width="10%"
                                    HorizontalAlign="Center" />
                                <ItemStyle Wrap="True" CssClass="td2c_b" HorizontalAlign="Center"></ItemStyle>
                                <ItemTemplate>
                                    <asp:ImageButton ID="ImgEdit" CommandArgument='<%#Container.DataItem("ROLEID") %>'
                                        ImageUrl="~/img/imgEdit.gif" runat="server" TabIndex="105" CommandName="Edit">
                                    </asp:ImageButton>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="�R��">
                                <HeaderStyle CssClass="th1c_b" HorizontalAlign="Center" Font-Size="9pt" Wrap="False"
                                    Height="22px" Width="10%" />
                                <ItemStyle Wrap="True" CssClass="td2c_b" HorizontalAlign="Center"></ItemStyle>
                                <ItemTemplate>
                                    <asp:ImageButton onmousedown="if(window.confirm('�T�w�n�R����ƶ�?')==true){hidOutAction.value='D';this.click();}"
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
                            <asp:TemplateColumn HeaderText="���">
                                <HeaderStyle CssClass="th1c_b" HorizontalAlign="Center" Font-Size="9pt" Wrap="False"
                                    Height="22px" Width="15%" />
                                <ItemStyle Wrap="True" CssClass="td2_b"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Label ID="lblDepart" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"DEPART")%>'>
                                    </asp:Label>
                                    <asp:HiddenField ID="hidBraDepno" runat="server" Value='<%#DataBinder.Eval(Container.DataItem,"BRA_DEPNO")%>' />
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="�H���W��">
                                <HeaderStyle CssClass="th1c_b" HorizontalAlign="Center" Font-Size="9pt" Wrap="False"
                                    Height="22px" Width="25%" />
                                <ItemStyle Wrap="True" CssClass="td2_b"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Label ID="lblUserName" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"USERNAME")%>'>
                                    </asp:Label>
                                    <asp:HiddenField ID="hidUserID" runat="server" Value='<%#DataBinder.Eval(Container.DataItem,"STAFFID")%>' />
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="����W��">
                                <HeaderStyle CssClass="th1c_b" HorizontalAlign="Center" Font-Size="9pt" Wrap="False"
                                    Height="22px" Width="40%" />
                                <ItemStyle Wrap="True" CssClass="td2_b"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Label ID="lblRoleName" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"ROLENAME")%>'>
                                    </asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="�ק�">
                                <HeaderStyle CssClass="th1c_b" Font-Size="9pt" Wrap="False" Height="22px" Width="10%"
                                    HorizontalAlign="Center" />
                                <ItemStyle Wrap="True" CssClass="td2c_b" HorizontalAlign="Center"></ItemStyle>
                                <ItemTemplate>
                                    <asp:ImageButton ID="ImgEdit" CommandArgument='<%#Container.DataItem("STAFFID") %>'
                                        ImageUrl="~/img/imgEdit.gif" runat="server" TabIndex="105" CommandName="Edit">
                                    </asp:ImageButton>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="�R��">
                                <HeaderStyle CssClass="th1c_b" HorizontalAlign="Center" Font-Size="9pt" Wrap="False"
                                    Height="22px" Width="10%" />
                                <ItemStyle Wrap="True" CssClass="td2c_b" HorizontalAlign="Center"></ItemStyle>
                                <ItemTemplate>
                                    <asp:ImageButton onmousedown="if(window.confirm('�T�w�n�R����ƶ�?')==true){hidOutAction.value='D';this.click();}"
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
                    ���@�D�ޤG <asp:DropDownList ID="ddlReviewer" runat="server"></asp:DropDownList>
                    </div>
                </td>
            </tr>
        </table>
    </div>
    <div id="divChecked" runat="server">
        <table id="Table1" width="100%" runat="server">
            <tr align="right">
                <td colspan="2">
                    �ץ�s��:<asp:Label ID="lblCaseId" runat="server"></asp:Label>
                </td>
            </tr>
            <tr align="right">
                <td colspan="2">
                    �ӽг��:
                    <asp:Label ID="lblBranch" runat="server"></asp:Label>
                </td>
            </tr>
        </table>
        <asp:Label ID="lblGrantStyle" runat="server"></asp:Label>
        <table id="tblChecked" class="mtr" border="0" cellspacing="1" cellpadding="2" width="100%"
            runat="server">
            <tr>
                <td style="width: 40%; text-align: center; height: 300px">
                    <span style="float: left">�ק�e</span><br />
                    <asp:ListBox ID="lstPreEdit" runat="server" Width="200px" Height="300px" Enabled="false">
                    </asp:ListBox>
                    <br />
                    <span style="text-align: center">���r�������R��</span>
                </td>
                <td style="width: 20%; text-align: center">
                    <asp:Label ID="lblMsg" runat="server"></asp:Label>
                </td>
                <td style="width: 40%; text-align: center; height: 300px">
                    <span style="float: left">�ק��</span><br />
                    <asp:ListBox ID="lstAfterEdit" runat="server" Width="200px" Height="300px" Enabled="false">
                    </asp:ListBox>
                    <br />
                    <span style="text-align: center">���r�������s�W</span>
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
                            <asp:TemplateColumn HeaderText="���">
                                <HeaderStyle CssClass="th1c_b" HorizontalAlign="Center" Font-Size="9pt" Wrap="False"
                                    Height="22px" Width="40%" />
                                <ItemStyle Wrap="True" CssClass="td2_b"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Label ID="lblDepart" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"BRANAME")%>'>
                                    </asp:Label>
                                    <asp:HiddenField ID="hidBraDepno" runat="server" Value='<%#DataBinder.Eval(Container.DataItem,"DEPNO")%>' />
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="�H���W��">
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
                            <asp:TemplateColumn HeaderText="���">
                                <HeaderStyle CssClass="th1c_b" HorizontalAlign="Center" Font-Size="9pt" Wrap="False"
                                    Height="22px" Width="40%" />
                                <ItemStyle Wrap="True" CssClass="td2_b"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Label ID="lblDepart" runat="server" Text='<%#DataBinder.Eval(Container.DataItem,"BRANAME")%>'>
                                    </asp:Label>
                                    <asp:HiddenField ID="hidBraDepno" runat="server" Value='<%#DataBinder.Eval(Container.DataItem,"DEPNO")%>' />
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="����W��">
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
            <asp:Button ID="btnAgree" runat="server" Text="�P�N" CssClass="bt" />
            <asp:Button ID="btnDiffer" runat="server" Text="���P�N" CssClass="bt" />
            <asp:Button ID="btnReviseFlow" runat="server" Text="�ץ��ɥR" CssClass="bt" />
        </div>
    </div>
    <div>
        <asp:HiddenField ID="hidOldXml" runat="server" />
    </div>
    <!--��ܤU��~�ؽu-->
    <!--#include virtual="~/inc/eLoanTableEnd.inc"-->
    </form>
</body>
</html>
