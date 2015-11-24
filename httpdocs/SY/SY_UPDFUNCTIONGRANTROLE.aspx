<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SY_UPDFUNCTIONGRANTROLE.aspx.vb"
    Inherits="MBSC.SY_UPDFUNCTIONGRANTROLE" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>�������</title>
    <base target="_self" />
    <!--SY CSS-->
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
            onclick="javascript:window.open('./help/SY_UPDFUNCTIONGRANTROLE.docx')" />
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
        �������
    </div>
    <br />
    <!--��ܤW��~�ؽu-->
    <!-- #include virtual="~/inc/eLoanTableStart.inc" -->
    <div id="divNormal" runat="server">
        <table width="100%">
            <tr>
                <td align="right" valign="top">
                    <asp:Button ID="btnConfirmSend" runat="server" Text="�T�{�e�X" CssClass="bt" Visible="true"
                        TabIndex="5" />&nbsp;&nbsp;
                    <asp:Button ID="btnCancelAll" runat="server" Text="��������" CssClass="bt" TabIndex="15" />
                    <table id="tbData" width="100%" class="mtr" cellspacing="1" cellpadding="2" runat="server">
                        <tr>
                            <td align="center">
                                ���ⶵ��
                            </td>
                            <td align="center">
                                �������
                            </td>
                        </tr>
                        <tr id="trSubSys" runat="server">
                            <td align="left" width="50%" valign="top">
                                <asp:TreeView ID="treeViewRole" runat="server" TabIndex="10">
                                </asp:TreeView>
                            </td>
                            <td valign="top">
                                <asp:TreeView ID="treeViewFun" runat="server" TabIndex="15" ShowCheckBoxes="All">
                                </asp:TreeView>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td colspan="2" align="center">
                    <asp:Button ID="btnSave" runat="server" Text="�x�s" CssClass="bt" TabIndex="20" />&nbsp;&nbsp;
                    <asp:Button ID="btnCancel" runat="server" Text="����" CssClass="bt" TabIndex="25" />
                </td>
            </tr>
        </table>
    </div>
    <div id="divCheck" runat="server" visible="false">
        <table width="100%">
            <tr>
                <td align="right" valign="top">
                    <table id="Table1" width="100%" class="mtr" cellspacing="1" cellpadding="2" runat="server">
                        <tr>
                            <td align="center">
                                ���ⶵ��
                            </td>
                            <td align="center">
                                �������
                            </td>
                        </tr>
                        <tr id="tr1" runat="server">
                            <td align="left" width="50%" valign="top">
                                <asp:TreeView ID="treeViewRoleCheck" runat="server" TabIndex="30">
                                </asp:TreeView>
                            </td>
                            <td>
                                <asp:TreeView ID="treeViewFunCheck" runat="server" TabIndex="35" ShowCheckBoxes="All">
                                </asp:TreeView>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td colspan="2" align="center">
                    <asp:Button ID="btnArgee" runat="server" Text="�P�N" CssClass="bt" TabIndex="40" />&nbsp;&nbsp;
                    <asp:Button ID="btnNotArgee" runat="server" Text="���P�N" CssClass="bt" TabIndex="45" />&nbsp;&nbsp;
                    <asp:Button ID="btnReviseFlow" runat="server" Text="�ץ��ɥR" CssClass="bt" TabIndex="50" />&nbsp;&nbsp;
                </td>
            </tr>
        </table>
    </div>
    <asp:HiddenField ID="hidRoleId" runat="server" />
    <asp:HiddenField ID="hidRoleName" runat="server" />
    <asp:HiddenField ID="hidSysId" runat="server" />
    <asp:HiddenField ID="hidSubSysId" runat="server" />
    <asp:HiddenField ID="hidNodeParent" runat="server" />
    <asp:HiddenField ID="hidValuePath" runat="server" />
    <asp:HiddenField ID="hidDelData" runat="server" />
    <asp:HiddenField ID="hidOldRoleList" runat="server" />
    <asp:HiddenField ID="hidNodeParentF" runat="server" />
    <!--��ܤU��~�ؽu-->
    <!--#include virtual="~/inc/eLoanTableEnd.inc"-->
    </form>
</body>
</html>
