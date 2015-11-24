<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SY_UPGBRANCHUSER.aspx.vb"
    Inherits="MBSC.SY_UPGBRANCHUSER" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>組織人員分派</title>
    <base target="_self" />
    <!--共用的SY CSS-->
    <!-- #include virtual="~/inc/SYCSS.inc" -->
    <script type="text/javascript">
        // 記錄刪除標識，提示信息
        function SaveChangeToXML() {
            var isChange = true;

            if (document.getElementById("hidRoleFlag").value == "1") {
                isChange = window.confirm("異動人員已分配角色，是否刪除？");
            }

            if (document.getElementById("hidRoleFlag").value == "1") {
                isChange = window.confirm("異動人員存在代理關係，是否刪除？");
            }

            return isChange;
        }

        // 調用儲存當前人員變動資料方法
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
                案件編號：<asp:Label ID="lblCaseId" runat="server"></asp:Label>
                <br />
                申請單位：<asp:Label ID="lblBranch" runat="server"></asp:Label>
            </td>
        </tr>
    </table>
    <br />
    <div class="TaskTtitle">
        組織人員分派
    </div>
    <br />
    <!--顯示上方外框線-->
    <!-- #include virtual="~/inc/eLoanTableStart.inc" -->
    <table width="100%">
        <tr id="trNormalCondition" runat="server">
            <td colspan="2">
                <div style="float: left;">
                    <asp:RadioButtonList ID="rdolNormalCondition" runat="server" AutoPostBack="true"
                        RepeatDirection="Horizontal">
                        <asp:ListItem Value="all">全部</asp:ListItem>
                        <asp:ListItem Value="branch" Selected="True">依單位</asp:ListItem>
                    </asp:RadioButtonList>
                </div>
                <div style="float: left;">
                    <asp:DropDownList ID="ddlBranch" AutoPostBack="true" runat="server" TabIndex="15">
                    </asp:DropDownList>
                </div>
                <div style="float: right;">
                    <asp:Button ID="btnSendFlow" runat="server" Text="確認送出" CssClass="bt" TabIndex="20" />
                    <asp:Button ID="btnCancelAll" runat="server" Text="全部取消" CssClass="bt" TabIndex="25" />
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
                                        <asp:Button ID="btnCheckIn" runat="server" Text="選入" CssClass="bt" TabIndex="40" />
                                        <br />
                                        <br />
                                        <asp:Button ID="btnCheckOut" runat="server" Text="選出" CssClass="bt" TabIndex="45" />
                                    </td>
                                    <td align="center">
                                        <asp:ListBox ID="lstRight" TabIndex="50" Width="98%" Height="300px" runat="server"
                                            SelectionMode="Multiple"></asp:ListBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td colspan="3" align="right">
                                        <asp:Button ID="btnSave" runat="server" Text="儲存" CssClass="bt" TabIndex="55" />&nbsp;&nbsp;
                                        <asp:Button ID="btnSaveStaff" runat="server" Width="0px" Height="0px" />
                                        <asp:Button ID="btnCancel" runat="server" Text="取消" CssClass="bt" TabIndex="60" />
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
                                        <span>修改前</span>
                                        <br />
                                        <asp:ListBox ID="lstEditBefore" TabIndex="70" Width="98%" Height="300px" runat="server">
                                        </asp:ListBox>
                                        <span>紅字為本次刪除</span>
                                    </td>
                                    <td align="center">
                                        <span>修改後</span>
                                        <br />
                                        <asp:ListBox ID="lstEditAfter" TabIndex="75" Width="98%" Height="300px" runat="server">
                                        </asp:ListBox>
                                        <span>紅字為本次新增</span>
                                    </td>
                                </tr>
                                <tr>
                                    <td colspan="3" align="center">
                                        <asp:Button ID="btnYes" runat="server" Text="同意" CssClass="bt" TabIndex="80" />&nbsp;&nbsp;
                                        <asp:Button ID="btnNo" runat="server" Text="不同意" CssClass="bt" TabIndex="85" />&nbsp;&nbsp;
                                        <asp:Button ID="btnReviseFlow" runat="server" Text="修正補充" CssClass="bt" TabIndex="90" />
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
    <!--顯示下方外框線-->
    <!--#include virtual="~/inc/eLoanTableEnd.inc"-->
    </form>
</body>
</html>
