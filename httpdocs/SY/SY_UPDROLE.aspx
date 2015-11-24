<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SY_UPDROLE.aspx.vb" Inherits="MBSC.SY_UPDROLE" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>����إ�</title>
    <base target="_self" />
    <!--�@�Ϊ�SY CSS-->
    <!-- #include virtual="~/inc/SYCSS.inc" -->
    <script language="javascript" type="text/javascript">
        // ??�`?��?�D?�ƥ�
        function postBackByObject()
        {
            var o = window.event.srcElement;
            if (o.tagName == "INPUT" && o.type == "checkbox")
            {
                __doPostBack("", "");
            }
        }
    </script>
</head>
<body>
    <form id="form1" runat="server">
    <div class="divhelp">
        <img alt="help" src="<%= com.Azion.EloanUtility.UIUtility.getImgPath()+"help.png"%>"
            onclick="javascript:window.open('./help/SY_UPDROLE.docx')" />
    </div>
    <br />
    <div id="divBranch" runat="server" visible="false">
        <table width="100%" runat="server">
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
    </div>
    <div class="TaskTtitle">
        ����إ�
    </div>
    <br />
    <!--��ܤW��~�ؽu-->
    <!-- #include virtual="~/inc/eLoanTableStart.inc" -->
    <div id="divNormal" runat="server">
        <table width="100%" runat="server" id="tdNormal">
            <tr>
                <td align="left" width="30%" valign="top">
                    <asp:TreeView ID="treeView" runat="server" TabIndex="5">
                        <RootNodeStyle ImageUrl="~/img/folder.png" />
                        <ParentNodeStyle ImageUrl="~/img/group_add.png" />
                        <LeafNodeStyle ImageUrl="~/img/group.png" />
                    </asp:TreeView>
                </td>
                <td align="right" valign="top" width="70%">
                    <asp:Button ID="btnConfirmSend" runat="server" Text="�T�{�e�X" CssClass="bt" TabIndex="10" />&nbsp;&nbsp;
                    <asp:Button ID="btnCancelAll" runat="server" Text="��������" CssClass="bt" TabIndex="15" />
                    <table width="100%" class="mtr" cellspacing="1" cellpadding="2" runat="server">
                        <tr>
                            <td style="width: 15%">
                                �W�h����
                            </td>
                            <td style="width: 85%">
                                <asp:Label ID="lblParentName" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr runat="server">
                            <td>
                                �s��
                            </td>
                            <td>
                                <asp:Label ID="lblID" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr runat="server">
                            <td>
                                �W��
                            </td>
                            <td>
                                <asp:Label ID="lblName" runat="server"></asp:Label>
                                <asp:TextBox ID="txtName" runat="server" MaxLength="100" Visible="false" TabIndex="20"
                                    AutoPostBack="True"></asp:TextBox>
                            </td>
                        </tr>
                        <tr id = "trRoleMgr" runat="server" visible="false">
                            <td>�W�h�޲z����</td>
                            <td>
                                <asp:TextBox ID="tbRoleMgr" runat="server" AutoPostBack="True" MaxLength="5" Width="80px"></asp:TextBox> &nbsp;<asp:Label ID="lbRoleMgr" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr id="trSubSys" runat="server" visible="false">
                            <td>
                                ���ݤl�t��
                            </td>
                            <td>
                                <asp:TreeView ID="treeViewSys" runat="server" ShowCheckBoxes="All" ShowExpandCollapse="False"
                                    EnableTheming="True" TabIndex="25">
                                    <RootNodeStyle ImageUrl="~/img/system_run.png" />
                                    <ParentNodeStyle ImageUrl="~/img/package_system.png" />
                                    <LeafNodeStyle ImageUrl="~/img/package_system.png" />
                                </asp:TreeView>
                            </td>
                        </tr>
                        <tr id="trDisable" runat="server" visible="false">
                            <td>
                                �ҥΪ��A
                            </td>
                            <td>
                                <asp:RadioButtonList ID="rdoDisable" runat="server" RepeatDirection="Horizontal"
                                    TabIndex="30" AutoPostBack="True">
                                    <asp:ListItem Selected="True" Value="0">�ҥ�</asp:ListItem>
                                    <asp:ListItem Value="1">����</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="2" align="right">
                                <asp:Button ID="btnAdd" runat="server" Text="�s�W" CssClass="bt" TabIndex="35" />&nbsp;&nbsp;
                                <asp:Button ID="btnSave" runat="server" Text="�x�s" CssClass="bt" TabIndex="40" Visible="false" />&nbsp;&nbsp;
                                <asp:Button ID="btnDelete" runat="server" Text="�R��" CssClass="bt" TabIndex="45" Visible="false" />&nbsp;&nbsp;
                                <asp:Button ID="btnCancel" runat="server" Text="����" CssClass="bt" TabIndex="50" Visible="false" />
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </div>
    <div id="divCheck" runat="server" visible="false">
        <table width="100%" runat="server" id="tdCheck">
            <tr>
                <td align="left" width="30%" valign="top">
                    <asp:TreeView ID="treeViewCheck" runat="server" TabIndex="5">
                        <RootNodeStyle ImageUrl="~/img/folder.png" />
                        <ParentNodeStyle ImageUrl="~/img/group_add.png" />
                        <LeafNodeStyle ImageUrl="~/img/group.png" />
                    </asp:TreeView>
                </td>
                <td align="right" valign="top">
                    <table id="Table1" width="100%" class="mtr" cellspacing="1" cellpadding="2" runat="server">
                        <tr>
                            <td colspan="2">
                                �ק�e
                            </td>
                            <td colspan="2">
                                �ק�Z
                            </td>
                        </tr>
                        <tr>
                            <td>
                                �W�h����
                            </td>
                            <td>
                                <asp:Label ID="lblParentNameBefore" runat="server"></asp:Label>
                            </td>
                            <td>
                                �W�h����
                            </td>
                            <td>
                                <asp:Label ID="lblParentNameAfter" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr id="Tr1" runat="server">
                            <td>
                                �s��
                            </td>
                            <td>
                                <asp:Label ID="lblIDBefore" runat="server"></asp:Label>
                            </td>
                            <td>
                                �s��
                            </td>
                            <td>
                                <asp:Label ID="lblIDAfter" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr id="Tr2" runat="server">
                            <td>
                                �W��
                            </td>
                            <td>
                                <asp:Label ID="lblNameBefore" runat="server"></asp:Label>
                            </td>
                            <td>
                                �W��
                            </td>
                            <td>
                                <asp:Label ID="lblNameAfter" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr id="tr3" runat="server">
                            <td>
                                ���ݤl�t��
                            </td>
                            <td>
                                <asp:TreeView ID="treeViewSysBefore" runat="server" ShowCheckBoxes="All" ShowExpandCollapse="False"
                                    EnableTheming="True" TabIndex="26" Enabled="false">
                                    <RootNodeStyle ImageUrl="~/img/system_run.png" />
                                    <ParentNodeStyle ImageUrl="~/img/package_system.png" />
                                    <LeafNodeStyle ImageUrl="~/img/package_system.png" />
                                </asp:TreeView>
                            </td>
                            <td>
                                ���ݤl�t��
                            </td>
                            <td>
                                <asp:TreeView ID="treeViewSysAfter" runat="server" ShowCheckBoxes="All" ShowExpandCollapse="False"
                                    EnableTheming="True" TabIndex="25" Enabled="false">
                                    <RootNodeStyle ImageUrl="~/img/system_run.png" />
                                    <ParentNodeStyle ImageUrl="~/img/package_system.png" />
                                    <LeafNodeStyle ImageUrl="~/img/package_system.png" />
                                </asp:TreeView>
                            </td>
                        </tr>
                        <tr id="tr4" runat="server">
                            <td>
                                �ҥΪ��A
                            </td>
                            <td>
                                <asp:RadioButtonList ID="rdoDisableBefore" runat="server" RepeatDirection="Horizontal"
                                    TabIndex="30" Enabled="false">
                                    <asp:ListItem Selected="True" Value="0">�ҥ�</asp:ListItem>
                                    <asp:ListItem Value="1">����</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                            <td>
                                �ҥΪ��A
                            </td>
                            <td>
                                <asp:RadioButtonList ID="rdoDisableAfter" runat="server" RepeatDirection="Horizontal"
                                    TabIndex="30" Enabled="false">
                                    <asp:ListItem Selected="True" Value="0">�ҥ�</asp:ListItem>
                                    <asp:ListItem Value="1">����</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <table width="100%">
            <tr>
                <td align="center">
                    <asp:Button ID="btnArgee" runat="server" Text="�P�N" TabIndex="40" CssClass="bt" />
                    <asp:Button ID="btnNotArgee" runat="server" Text="���P�N" TabIndex="35" CssClass="bt" />
                    <asp:Button ID="btnReviseFlow" runat="server" Text="�ץ��ɥR" TabIndex="45" CssClass="bt" />
                </td>
            </tr>
        </table>
    </div>
    <asp:HiddenField ID="hidParentId" runat="server" />
    <asp:HiddenField ID="hidParentName" runat="server" />
    <asp:HiddenField ID="hidId" runat="server" />
    <asp:HiddenField ID="hidName" runat="server" />
    <asp:HiddenField ID="hidFlag" runat="server" />
    <asp:HiddenField ID="hidNodeParent" runat="server" />
    <asp:HiddenField ID="hidOldData" runat="server" />
    <asp:HiddenField ID="hidDelData" runat="server" />
    <asp:HiddenField ID="hidRedFlag" runat="server" />
    <asp:HiddenField ID="hidParId" runat="server" />
    <asp:HiddenField ID="hidValue" runat="server" />
    <!--��ܤU��~�ؽu-->
    <!--#include virtual="~/inc/eLoanTableEnd.inc"-->
    </form>
</body>
</html>

