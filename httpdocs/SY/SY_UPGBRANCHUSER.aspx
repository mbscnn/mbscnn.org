<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SY_UPGBRANCHUSER.aspx.vb"
    Inherits="MBSC.SY_UPGBRANCHUSER" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>��´�H������</title>
    <base target="_self" />
    <!--�@�Ϊ�SY CSS-->
    <!-- #include virtual="~/inc/SYCSS.inc" -->
    <script type="text/javascript">
        // �O���R�����ѡA���ܫH��
        function SaveChangeToXML() {
            var isChange = true;

            if (document.getElementById("hidRoleFlag").value == "1") {
                isChange = window.confirm("���ʤH���w���t����A�O�_�R���H");
            }

            if (document.getElementById("hidRoleFlag").value == "1") {
                isChange = window.confirm("���ʤH���s�b�N�z���Y�A�O�_�R���H");
            }

            return isChange;
        }

        // �ե��x�s��e�H���ܰʸ�Ƥ�k
        function CallStaffSave() {
            document.getElementById("btnSaveStaff").click();
        }
    </script>
</head>
<body>
    <form id="form1" runat="server">
    <div class="divhelp">
        <img alt="help" src="<%= com.Azion.EloanUtility.UIUtility.getImgPath()+"help.png"%>"
            onclick="javascript:window.open('./help/SY_UPGBRANCHUSER.docx')" />
    </div>
    <table runat = "server" id ="trHeader" width="100%">
        <tr id="tr1" runat="server">
            <td colspan="2" align="right">
                �ץ�s���G<asp:Label ID="lblCaseId" runat="server"></asp:Label>
                <br />
                �ӽг��G<asp:Label ID="lblBranch" runat="server"></asp:Label>
            </td>
        </tr>
    </table>
    <br />
    <div class="TaskTtitle">
        ��´�H������
    </div>
    <br />
    <!--��ܤW��~�ؽu-->
    <!-- #include virtual="~/inc/eLoanTableStart.inc" -->
    <table width="100%">
        <tr id="trNormalCondition" runat="server">
            <td colspan="2">
                <div style="float: left;">
                    <asp:RadioButtonList ID="rdolNormalCondition" runat="server" AutoPostBack="true"
                        RepeatDirection="Horizontal">
                        <asp:ListItem Value="all">����</asp:ListItem>
                        <asp:ListItem Value="branch" Selected="True">�̳��</asp:ListItem>
                    </asp:RadioButtonList>
                </div>
                <div style="float: left;">
                    <asp:DropDownList ID="ddlBranch" AutoPostBack="true" runat="server" TabIndex="15">
                    </asp:DropDownList>
                </div>
                <div style="float: right;">
                    <asp:Button ID="btnSendFlow" runat="server" Text="�T�{�e�X" CssClass="bt" TabIndex="20" />
                    <asp:Button ID="btnCancelAll" runat="server" Text="��������" CssClass="bt" TabIndex="25" />
                </div>
            </td>
        </tr>
        <tr id="trNormalContent" runat="server">
            <td>
                <table id="tblNormalContent" width="100%" class="mtr" cellspacing="1" cellpadding="2"
                    runat="server">
                    <tr>
                        <td align="left" width="30%" valign="top">
                            <asp:TreeView ID="tvBranch" runat="server" TabIndex="30">
                            </asp:TreeView>
                        </td>
                        <td align="right" valign="top">
                            <table id="Table2" width="100%" class="mtr" cellspacing="1" cellpadding="2" runat="server">
                                <tr>
                                    <td align="center">
                                        <asp:ListBox ID="lstLeft" TabIndex="35" Width="98%" Height="300px" runat="server"
                                            SelectionMode="Multiple"></asp:ListBox>
                                    </td>
                                    <td align="center" style="width: 10%;">
                                        <asp:Button ID="btnCheckIn" runat="server" Text="��J" CssClass="bt" TabIndex="40" />
                                        <br />
                                        <br />
                                        <asp:Button ID="btnCheckOut" runat="server" Text="��X" CssClass="bt" TabIndex="45" />
                                    </td>
                                    <td align="center">
                                        <asp:ListBox ID="lstRight" TabIndex="50" Width="98%" Height="300px" runat="server"
                                            SelectionMode="Multiple"></asp:ListBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td colspan="3" align="right">
                                        <asp:Button ID="btnSave" runat="server" Text="�x�s" CssClass="bt" TabIndex="55" />&nbsp;&nbsp;
                                        <asp:Button ID="btnSaveStaff" runat="server" Width="0px" Height="0px" />
                                        <asp:Button ID="btnCancel" runat="server" Text="����" CssClass="bt" TabIndex="60" />
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr id="trCheckContent" runat="server">
            <td>
                <table id="tblCheckContent" width="100%" class="mtr" cellspacing="1" cellpadding="2"
                    runat="server">
                    <tr>
                        <td align="left" width="30%" valign="top">
                            <asp:TreeView ID="tvBranchCheck" runat="server" TabIndex="65"  ViewStateMode="Enabled">
                            </asp:TreeView>
                        </td>
                        <td align="right" valign="top">
                            <table id="Table4" width="100%" class="mtr" cellspacing="1" cellpadding="2" runat="server">
                                <tr>
                                    <td align="center">
                                        <span>�ק�e</span>
                                        <br />
                                        <asp:ListBox ID="lstEditBefore" TabIndex="70" Width="98%" Height="300px" runat="server">
                                        </asp:ListBox>
                                        <span>���r�������R��</span>
                                    </td>
                                    <td align="center">
                                        <span>�ק��</span>
                                        <br />
                                        <asp:ListBox ID="lstEditAfter" TabIndex="75" Width="98%" Height="300px" runat="server">
                                        </asp:ListBox>
                                        <span>���r�������s�W</span>
                                    </td>
                                </tr>
                                <tr>
                                    <td colspan="3" align="center">
                                        <asp:Button ID="btnYes" runat="server" Text="�P�N" CssClass="bt" TabIndex="80" />&nbsp;&nbsp;
                                        <asp:Button ID="btnNo" runat="server" Text="���P�N" CssClass="bt" TabIndex="85" />&nbsp;&nbsp;
                                        <asp:Button ID="btnReviseFlow" runat="server" Text="�ץ��ɥR" CssClass="bt" TabIndex="90" />
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    <asp:HiddenField ID="hidDeleteFlag" runat="server" />
    <asp:HiddenField ID="hidRoleFlag" runat="server" />
    <asp:HiddenField ID="hidAgentFlag" runat="server" />
    <asp:HiddenField ID="hidSelectedNode" runat="server" />
    <!--��ܤU��~�ؽu-->
    <!--#include virtual="~/inc/eLoanTableEnd.inc"-->
    </form>
</body>
</html>
