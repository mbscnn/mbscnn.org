<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SY_UPDBRANCH.aspx.vb"
    Inherits="MBSC.SY_UPDBRANCH" %>

<%@ Register Src="../Module/UIControl/Addr/ZipCode.ascx" TagName="ZipCode" TagPrefix="uc1" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>��´���@�@�~</title>
    <base target="_self" />
    <!--�@�Ϊ�SY CSS-->
    <!-- #include virtual="~/inc/SYCSS.inc" -->
    <script type="text/javascript">
        // ��J����
        function CheckValidate() {
            var branchName = document.getElementById("txtBranchName");
            var braTel = document.getElementById("txtBraTel");
            var brTel = document.getElementById("txtBrTel");
            var brid = document.getElementById("txtBrid");
            var mgrBrids = document.getElementById("txtMgrBrid");
            var braTelTest = new RegExp(/^\d{0,4}$/);
            var brTelTest = new RegExp(/^\d{0,10}$/);
            var errorMsg = "";
            var errorControl = null;

            // ?��W?
            if (branchName != null) {
                //  ���W�٥�������
                if (TrimText(branchName.value) == "") {
                    errorControl = branchName;

                    errorMsg = "�п�J���W�١I\n";
                }
            }

            // ?��N?
            if (TrimText(brid.value) == "") {
                if (errorControl == null) {
                    errorControl = brid;
                }

                errorMsg += "�п�J���N�X�I\n";
            }

            // �޲z?��
            if (TrimText(mgrBrids.value) != "") {
                var mgrBrid = mgrBrids.value.replace("�A", ",").split(",");

                for (var i = 0; i < mgrBrid.length; i++) {
                    if (!isNumber(mgrBrid[i])) {
                        if (errorControl == null) {
                            errorControl = mgrBrids;
                        }

                        errorMsg += "�޲z���u���J�Ʀr�]�h�ӮɥΡ��A���j?�^�I\n";
                        
                        break;
                    }
                }
            }

            // �q�ܰϸ���J����
            if (!braTelTest.test(TrimText(braTel.value))) {
                if (errorControl == null) {
                    errorControl = braTel;
                }

                errorMsg += "�q�ܰϳ̦h�u���J4�X��ơI\n";
            }

            // �q������ 
            if (!brTelTest.test(TrimText(brTel.value))) {
                if (errorControl == null) {
                    errorControl = brTel;
                }

                errorMsg += "�q�ܰϳ̦h�u���J10�X��ơI\n";
            }

            if (errorControl != null) {
                alert(errorMsg);
                errorControl.focus();
                errorControl.value = "";
                return false;
            }

            return true;
        }

        // �h�Ů�
        function TrimText(text) {
            return (text || "").replace(RegExp(/^\s+|\s+$/g), "");
        }

        // �ˬd�O�_���Ʀr
        function isNumber(strNum) {
            var num = new Number(stripCommas(strNum));
            if (isNaN(num)) {
                return false;
            }
            return true;
        }

        // ?���A����??��
        function stripCommas(numString) {
            var re = /,/g;
            if (typeof numString == "number") {
                return numString;
            }
            return numString.replace(re, "");
        }
    </script>
</head>
<body>
    <form id="form1" runat="server">
    <div class="divhelp">
        <img alt="help" src="<%= com.Azion.EloanUtility.UIUtility.getImgPath()+"help.png"%>"
            onclick="javascript:window.open('./help/SY_UPDBRANCH.docx')" />
    </div>
    <br />
    <div class="TaskTtitle">
        ��´���@�@�~
    </div>
    <br />
    <!--��ܤW��~�ؽu-->
    <!-- #include virtual="~/inc/eLoanTableStart.inc" -->
    <table width="100%">
        <tr>
            <td colspan="2">
                <asp:RadioButton ID="rdoAll" AutoPostBack="true" runat="server" GroupName="branch"
                    TabIndex="5" />
                ����
                <asp:RadioButton ID="rdoBranch" AutoPostBack="true" Checked="true" runat="server"
                    GroupName="branch" TabIndex="10" />
                �̳��
                <asp:DropDownList ID="ddlBranch" AutoPostBack="true" runat="server" TabIndex="15">
                </asp:DropDownList>
            </td>
        </tr>
        <tr>
            <td>
                <table id="tblContent" width="100%" class="mtr" cellspacing="1" cellpadding="2" runat="server">
                    <tr>
                        <td align="left" width="30%" valign="top">
                            <asp:TreeView ID="tvBranch" runat="server" ExpandDepth ="2">
                                <RootNodeStyle ImageUrl="~/img/company1.png" />
                                <ParentNodeStyle ImageUrl="~/img/bank2.png" />
                                <LeafNodeStyle ImageUrl="~/img/company3.png" />
                            </asp:TreeView>
                        </td>
                        <td align="right" valign="top">
                            <table id="tblDetail" width="100%" class="mtr" cellspacing="1" cellpadding="2" runat="server">
                                <tr>
                                    <td style="width: 30%">
                                        �W�h���
                                    </td>
                                    <td>
                                        <asp:Label ID="lblParentName" runat="server"></asp:Label>
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        ���W��
                                    </td>
                                    <td>
                                        <asp:Label ID="lblBranchName" runat="server"></asp:Label>
                                        <asp:TextBox ID="txtBranchName" runat="server" MaxLength="30" Visible="false" 
                                            TabIndex="20" Width="214px"></asp:TextBox>
                                    </td>
                                </tr>
                                <tr id="trBrid" runat="server">
                                    <td>
                                        ���N�X
                                    </td>
                                    <td>
                                        <asp:TextBox ID="txtBrid" runat="server" MaxLength = "3" Width ="50px" TabIndex="25"></asp:TextBox>
                                        <asp:HiddenField ID="hidBraDepNo" runat="server" />
                                        <asp:HiddenField ID="hidBrid" runat="server" />
                                    </td>
                                </tr>
                                <tr id="trMgrBrid" runat="server">
                                    <td>
                                        �޲z���
                                    </td>
                                    <td>
                                        <asp:TextBox ID="txtMgrBrid" runat="server" Width ="50px" TabIndex="30"></asp:TextBox>
                                    </td>
                                </tr>
                                <tr id="trAddr" runat="server">
                                    <td>
                                        �a�}
                                    </td>
                                    <td>
                                        <uc1:zipcode id="brAddr" runat="server"/>
                                    </td>
                                </tr>
                                <tr id="trTel" runat="server">
                                    <td>
                                        �q��
                                    </td>
                                    <td>
                                        <asp:TextBox ID="txtBraTel" runat="server" MaxLength="4" Width="10%" TabIndex="40"></asp:TextBox>
                                        ��
                                        <asp:TextBox ID="txtBrTel" runat="server" MaxLength="10" TabIndex="45"></asp:TextBox>
                                    </td>
                                </tr>
                                <tr id="trProperty" runat="server">
                                    <td>
                                        �ݩ�
                                    </td>
                                    <td>
                                        <asp:RadioButtonList ID="rdolProperty" runat="server" RepeatDirection="Horizontal" TabIndex="50">
                                        </asp:RadioButtonList>
                                    </td>
                                </tr>
                                <tr id="trStatus" runat="server">
                                    <td>
                                        �ҥΪ��A
                                    </td>
                                    <td>
                                        <asp:RadioButtonList ID="rdoDisable" runat="server" RepeatDirection="Horizontal"
                                            TabIndex="55">
                                            <asp:ListItem Selected="True" Value="0">�ҥ�</asp:ListItem>
                                            <asp:ListItem Value="1">����</asp:ListItem>
                                        </asp:RadioButtonList>
                                    </td>
                                </tr>
                                <tr>
                                    <td colspan="2" align="right">
                                        <asp:Button ID="btnAdd" runat="server" Text="�s�W" CssClass="bt" TabIndex="60" />&nbsp;&nbsp;
                                        <asp:Button ID="btnDelete" runat="server" Text="�R��" CssClass="bt" OnClientClick="return confirm('�T�{�R���ӳ��H');"
                                            TabIndex="65" Visible="false" />&nbsp;&nbsp;
                                        <asp:Button ID="btnSave" runat="server" Text="�x�s" CssClass="bt" TabIndex="70" Visible="false"
                                            OnClientClick="return CheckValidate();" />&nbsp;&nbsp;
                                        <asp:Button ID="btnCancel" runat="server" Text="����" CssClass="bt" TabIndex="75" />
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    <asp:HiddenField ID="hidAddNodeValuePath" runat="server" />
    <asp:HiddenField ID="hidOperateFlag" runat="server" />
    <asp:HiddenField ID="hidDdlBraDepNo" runat="server" />
    <asp:HiddenField ID = "hidScript" runat = "server" />
    <!--��ܤU��~�ؽu-->
    <!--#include virtual="~/inc/eLoanTableEnd.inc"-->
    </form>
</body>
</html>
